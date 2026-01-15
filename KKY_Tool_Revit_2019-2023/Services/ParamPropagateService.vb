Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Text
Imports WinForms = System.Windows.Forms
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

Namespace Services

    '==================== 공용: 로드 옵션 ====================
    Friend Class FamilyLoadOptionsAllOverwrite
        Implements IFamilyLoadOptions

        Public Function OnFamilyFound(familyInUse As Boolean,
                                      ByRef overwriteParameterValues As Boolean) As Boolean _
            Implements IFamilyLoadOptions.OnFamilyFound
            overwriteParameterValues = True
            Return True
        End Function

        Public Function OnSharedFamilyFound(sharedFamily As Family,
                                            familyInUse As Boolean,
                                            ByRef source As FamilySource,
                                            ByRef overwriteParameterValues As Boolean) As Boolean _
            Implements IFamilyLoadOptions.OnSharedFamilyFound
            source = FamilySource.Family
            overwriteParameterValues = True
            Return True
        End Function
    End Class

    '==================== 공용: 실패 전처리/트랜잭션 유틸 ====================
    Friend Class SwallowWarningsOnly
        Implements IFailuresPreprocessor

        Public Function PreprocessFailures(failAcc As FailuresAccessor) As FailureProcessingResult _
            Implements IFailuresPreprocessor.PreprocessFailures

            For Each f In failAcc.GetFailureMessages()
                If f.GetSeverity() = FailureSeverity.Warning Then
                    failAcc.DeleteWarning(f)
                End If
            Next

            Return FailureProcessingResult.Continue
        End Function
    End Class

    Friend Module TxnUtil

        Public Sub WithTxn(doc As Document, name As String, action As Action)
            Using t As New Transaction(doc, name)
                t.Start()

                Dim opt = t.GetFailureHandlingOptions()
                opt.SetFailuresPreprocessor(New SwallowWarningsOnly())
                opt.SetClearAfterRollback(True)
                t.SetFailureHandlingOptions(opt)

                action.Invoke()

                t.Commit()
            End Using
        End Sub

        ' LoadFamily는 트랜잭션 밖에서 호출
        Public Sub SafeLoadFamily(sourceFamDoc As Document,
                                  targetDoc As Document,
                                  Optional label As String = "LoadFamily")

            If targetDoc Is Nothing Then Throw New ArgumentNullException(NameOf(targetDoc))

            If targetDoc.IsModifiable Then
                Throw New InvalidOperationException("Target document has an open transaction. Close it before LoadFamily.")
            End If

            Try
                If sourceFamDoc IsNot Nothing AndAlso sourceFamDoc.IsModified Then
                    If Not String.IsNullOrEmpty(sourceFamDoc.PathName) Then
                        sourceFamDoc.Save()
                    End If
                End If
            Catch
                ' 저장 실패는 무시 (로드 자체는 시도)
            End Try

            Dim opts As New FamilyLoadOptionsAllOverwrite()
            sourceFamDoc.LoadFamily(targetDoc, opts)

            Try : targetDoc.Regenerate() : Catch : End Try
        End Sub

    End Module

    '==================== 공유 패밀리 판정 ====================
    Friend Module SharedFamilyHelper

        Private ReadOnly _cache As New Dictionary(Of Integer, Boolean)()

        ''' <summary>
        ''' 해당 패밀리의 "공유 패밀리(Shared)" 체크 여부를 ownerDocCanEdit 기준으로 판정.
        ''' </summary>
        Public Function IsFamilyShared(ownerDocCanEdit As Document, fam As Family) As Boolean
            If fam Is Nothing Then Return False

            Dim key As Integer = fam.Id.IntegerValue
            Dim cached As Boolean
            If _cache.TryGetValue(key, cached) Then Return cached

            Dim fdoc As Document = Nothing
            Try
                fdoc = ownerDocCanEdit.EditFamily(fam)
                Dim p As Parameter = fdoc.OwnerFamily.Parameter(BuiltInParameter.FAMILY_SHARED)
                Dim isSharedFlag As Boolean = (p IsNot Nothing AndAlso p.AsInteger() = 1)
                _cache(key) = isSharedFlag
                Return isSharedFlag
            Catch
                _cache(key) = False
                Return False
            Finally
                If fdoc IsNot Nothing Then
                    Try : fdoc.Close(False) : Catch : End Try
                End If
            End Try
        End Function

    End Module

    '==================== 파라미터 타입 문자열 헬퍼 ====================
    Friend Module SharedParamTypeHelper

        Friend Function GetParamTypeString(ed As ExternalDefinition) As String
            If ed Is Nothing Then Return String.Empty

#If REVIT2019 Or REVIT2021 Then
            Return ed.ParameterType.ToString()
#ElseIf REVIT2023 Or REVIT2025 Then
            Try
                Dim dataType As ForgeTypeId = ed.GetDataType()
                If dataType Is Nothing Then Return String.Empty

                Try
                    Return LabelUtils.GetLabelForSpec(dataType)
                Catch
                End Try

                Try
                    Return dataType.TypeId
                Catch
                End Try

                Return dataType.ToString()
            Catch
                Return String.Empty
            End Try
#End If
        End Function

    End Module

    Friend Module ProgressMessageHelper
        Friend Function MakePhaseError(phase As String, famName As String, famId As ElementId, message As String) As String
            Dim idVal As Integer = 0
            Try
                If famId IsNot Nothing Then idVal = famId.IntegerValue
            Catch
                idVal = 0
            End Try
            Return $"[{phase}] Family='{famName}' (Id:{idVal}) - {message}"
        End Function

        Friend Function MakeInvalidObjectMessage(phase As String,
                                                 famName As String,
                                                 famId As ElementId,
                                                 operation As String,
                                                 ex As Autodesk.Revit.Exceptions.InvalidObjectException) As String
            Dim action As String = If(String.IsNullOrWhiteSpace(operation), "Unknown", operation)
            Dim baseMsg As String = If(ex Is Nothing, String.Empty, ex.Message)
            Return MakePhaseError(phase, famName, famId, $"{action} -> {baseMsg}")
        End Function

        Friend Function TryGetFreshFamily(doc As Document,
                                           famName As String,
                                           ByRef famId As ElementId,
                                           nameToFamilyId As Dictionary(Of String, ElementId)) As Family
            Dim fam As Family = Nothing

            If doc Is Nothing OrElse String.IsNullOrWhiteSpace(famName) Then Return Nothing

            If famId IsNot Nothing Then
                Try
                    fam = TryCast(doc.GetElement(famId), Family)
                Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                    fam = Nothing
                Catch
                    fam = Nothing
                End Try
            End If

            If fam Is Nothing Then
                Dim foundId As ElementId = FindFamilyIdByName(doc, famName)
                famId = foundId
                If foundId IsNot Nothing Then
                    Try
                        fam = TryCast(doc.GetElement(foundId), Family)
                    Catch
                        fam = Nothing
                    End Try

                    If fam IsNot Nothing AndAlso nameToFamilyId IsNot Nothing Then
                        nameToFamilyId(famName) = foundId
                    End If
                End If
            End If

            Return fam
        End Function

        Friend Function FindFamilyIdByName(doc As Document, famName As String) As ElementId
            If doc Is Nothing OrElse String.IsNullOrWhiteSpace(famName) Then Return Nothing
            Return New FilteredElementCollector(doc).
                   OfClass(GetType(Family)).
                   Cast(Of Family)().
                   FirstOrDefault(Function(f) String.Equals(f.Name, famName, StringComparison.OrdinalIgnoreCase))?.
                   Id
        End Function
    End Module

    '==================== 파라미터 선택 폼 ====================
    Friend Class FormSharedParamPicker
        Inherits WinForms.Form

        Private ReadOnly _app As Application
        Private _defFile As DefinitionFile

        Private txtSearch As WinForms.TextBox
        Private lstGroups As WinForms.ListBox
        Private lvParams As WinForms.ListView
        Private btnOK As WinForms.Button
        Private btnCancel As WinForms.Button
        Private chkExcludeDummy As WinForms.CheckBox
        Private cboTargetGroup As WinForms.ComboBox
        Private rbInstance As WinForms.RadioButton
        Private rbType As WinForms.RadioButton

        Public ReadOnly Property SelectedExternalDefinitions As New List(Of ExternalDefinition)
        Public ReadOnly Property ExcludeDummy As Boolean
            Get
                Return chkExcludeDummy IsNot Nothing AndAlso chkExcludeDummy.Checked
            End Get
        End Property
        Public ReadOnly Property SelectedGroupPG As BuiltInParameterGroup
            Get
                Dim item = TryCast(cboTargetGroup.SelectedItem, GroupItem)
                If item Is Nothing Then Return BuiltInParameterGroup.PG_IDENTITY_DATA
                Return item.PG
            End Get
        End Property
        Public ReadOnly Property SelectedIsInstance As Boolean
            Get
                If rbInstance Is Nothing Then Return True
                Return rbInstance.Checked
            End Get
        End Property

        Private Class GroupItem
            Public ReadOnly PG As BuiltInParameterGroup
            Public ReadOnly Name As String
            Public Sub New(pg As BuiltInParameterGroup)
                Me.PG = pg : Me.Name = pg.ToString()
            End Sub
            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Class

        Public Sub New(app As Application)
            _app = app
            Me.Text = "Select Shared Parameter(s) – Multi-select"
            Me.StartPosition = WinForms.FormStartPosition.CenterScreen
            Me.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
            Me.MaximizeBox = False : Me.MinimizeBox = False
            Me.Width = 920 : Me.Height = 700
            BuildUI()
            LoadDefinitions()
            LoadGroupCombo()
        End Sub

        Private Sub BuildUI()
            Dim lblSearch As New WinForms.Label() With {.Text = "Search:", .Left = 12, .Top = 14, .AutoSize = True}
            txtSearch = New WinForms.TextBox() With {.Left = 70, .Top = 10, .Width = 820}
            AddHandler txtSearch.TextChanged, AddressOf OnSearchChanged

            Dim lblGroups As New WinForms.Label() With {
                .Text = "Shared Parameter Groups (Multi-select / (All Groups) available)",
                .Left = 12, .Top = 44, .AutoSize = True
            }
            lstGroups = New WinForms.ListBox() With {
                .Left = 12,
                .Top = 64,
                .Width = 320,
                .Height = 520,
                .SelectionMode = WinForms.SelectionMode.MultiExtended
            }
            AddHandler lstGroups.SelectedIndexChanged, AddressOf OnGroupChanged

            Dim lblParams As New WinForms.Label() With {
                .Text = "Parameters — Name / Type / Visible / Group",
                .Left = 344,
                .Top = 44,
                .AutoSize = True
            }
            lvParams = New WinForms.ListView() With {
                .Left = 344,
                .Top = 64,
                .Width = 556,
                .Height = 520,
                .View = WinForms.View.Details,
                .FullRowSelect = True,
                .MultiSelect = True
            }
            lvParams.Columns.Add("Name", 240)
            lvParams.Columns.Add("Type", 100)
            lvParams.Columns.Add("Visible", 70)
            lvParams.Columns.Add("Group", 130)
            AddHandler lvParams.DoubleClick, AddressOf OnParamDoubleClick

            chkExcludeDummy = New WinForms.CheckBox() With {
                .Text = "하위 패밀리 이름에 'Dummy' 포함 시 제외 (기본)",
                .Left = 12,
                .Top = 592,
                .Width = 480,
                .Checked = True
            }

            Dim lblTarget As New WinForms.Label() With {
                .Text = "추가할 파라미터 그룹:",
                .Left = 344,
                .Top = 592,
                .AutoSize = True
            }
            cboTargetGroup = New WinForms.ComboBox() With {
                .Left = 480,
                .Top = 588,
                .Width = 220,
                .DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            }

            rbInstance = New WinForms.RadioButton() With {
                .Text = "인스턴스",
                .Left = 720,
                .Top = 588,
                .AutoSize = True,
                .Checked = True
            }
            rbType = New WinForms.RadioButton() With {
                .Text = "타입",
                .Left = 800,
                .Top = 588,
                .AutoSize = True
            }

            btnOK = New WinForms.Button() With {.Text = "OK", .Left = 716, .Top = 620, .Width = 85}
            AddHandler btnOK.Click, AddressOf OnOK
            btnCancel = New WinForms.Button() With {.Text = "Cancel", .Left = 815, .Top = 620, .Width = 85}
            AddHandler btnCancel.Click, AddressOf OnCancel

            Controls.AddRange(New WinForms.Control() {
                lblSearch, txtSearch,
                lblGroups, lstGroups,
                lvParams,
                chkExcludeDummy,
                lblTarget, cboTargetGroup,
                rbInstance, rbType,
                btnOK, btnCancel
            })
        End Sub

        Private Sub LoadDefinitions()
            SelectedExternalDefinitions.Clear()
            lvParams.Items.Clear()
            lstGroups.Items.Clear()

            If String.IsNullOrEmpty(_app.SharedParametersFilename) OrElse Not File.Exists(_app.SharedParametersFilename) Then
                Throw New InvalidOperationException("Revit 옵션에서 Shared Parameters 파일을 지정해 주세요.")
            End If

            _defFile = _app.OpenSharedParameterFile()
            If _defFile Is Nothing Then Throw New InvalidOperationException("Shared Parameters 파일을 열 수 없습니다.")

            lstGroups.Items.Add("(All Groups)")
            For Each g As DefinitionGroup In _defFile.Groups
                lstGroups.Items.Add(g.Name)
            Next

            If lstGroups.Items.Count > 0 Then
                lstGroups.SelectedIndices.Clear()
                lstGroups.SelectedIndex = 0
            End If

            PopulateParams()
        End Sub

        Private Sub LoadGroupCombo()
            cboTargetGroup.Items.Clear()
            Dim preferred As BuiltInParameterGroup() = {
                BuiltInParameterGroup.PG_TEXT,
                BuiltInParameterGroup.PG_IDENTITY_DATA,
                BuiltInParameterGroup.PG_DATA,
                BuiltInParameterGroup.PG_CONSTRAINTS
            }
            Dim added As New HashSet(Of BuiltInParameterGroup)()
            For Each pg In preferred
                cboTargetGroup.Items.Add(New GroupItem(pg))
                added.Add(pg)
            Next

            For Each pg As BuiltInParameterGroup In [Enum].GetValues(GetType(BuiltInParameterGroup))
                If Not added.Contains(pg) Then
                    cboTargetGroup.Items.Add(New GroupItem(pg))
                End If
            Next

            Dim defIndex As Integer = -1
            For i = 0 To cboTargetGroup.Items.Count - 1
                Dim gi = TryCast(cboTargetGroup.Items(i), GroupItem)
                If gi IsNot Nothing AndAlso gi.PG = BuiltInParameterGroup.PG_TEXT Then
                    defIndex = i
                    Exit For
                End If
            Next

            If defIndex < 0 Then defIndex = 0
            cboTargetGroup.SelectedIndex = defIndex
        End Sub

        Private Sub OnGroupChanged(sender As Object, e As EventArgs)
            PopulateParams()
        End Sub

        Private Sub OnSearchChanged(sender As Object, e As EventArgs)
            PopulateParams()
        End Sub

        Private Sub OnParamDoubleClick(sender As Object, e As EventArgs)
            OnOK(sender, e)
        End Sub

        Private Sub OnOK(sender As Object, e As EventArgs)
            SelectedExternalDefinitions.Clear()
            If lvParams.SelectedItems.Count = 0 Then
                WinForms.MessageBox.Show("하나 이상의 파라미터를 선택하세요.", "KKY", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information)
                Return
            End If

            For Each it As WinForms.ListViewItem In lvParams.SelectedItems
                Dim ed = TryCast(it.Tag, ExternalDefinition)
                If ed IsNot Nothing Then
                    SelectedExternalDefinitions.Add(ed)
                End If
            Next

            Me.DialogResult = WinForms.DialogResult.OK
        End Sub

        Private Sub OnCancel(sender As Object, e As EventArgs)
            Me.DialogResult = WinForms.DialogResult.Cancel
        End Sub

        Private Sub PopulateParams()
            lvParams.Items.Clear()
            If _defFile Is Nothing Then Return

            Dim selectedGroupNames As New List(Of String)()
            For Each idx As Integer In lstGroups.SelectedIndices
                selectedGroupNames.Add(lstGroups.Items(idx).ToString())
            Next

            Dim search As String = If(txtSearch.Text, String.Empty).Trim()

            Dim useAll As Boolean =
                (selectedGroupNames.Count = 0) OrElse
                (selectedGroupNames.Count = 1 AndAlso String.Equals(selectedGroupNames(0), "(All Groups)", StringComparison.OrdinalIgnoreCase))

            Dim groupsToShow As IEnumerable(Of DefinitionGroup)
            If useAll Then
                groupsToShow = _defFile.Groups
            Else
                Dim picked = New HashSet(Of String)(
                    selectedGroupNames.Where(Function(n) Not String.Equals(n, "(All Groups)", StringComparison.OrdinalIgnoreCase)),
                    StringComparer.OrdinalIgnoreCase)
                groupsToShow = _defFile.Groups.Cast(Of DefinitionGroup)().Where(Function(g) picked.Contains(g.Name))
            End If

            For Each g In groupsToShow
                For Each d As Definition In g.Definitions
                    Dim ed = TryCast(d, ExternalDefinition)
                    If ed Is Nothing Then Continue For

                    If Not String.IsNullOrEmpty(search) Then
                        Dim okName As Boolean = d.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                        Dim okGroup As Boolean = g.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0
                        If Not (okName OrElse okGroup) Then Continue For
                    End If

                    Dim lvi As New WinForms.ListViewItem(d.Name)
                    lvi.SubItems.Add(SharedParamTypeHelper.GetParamTypeString(ed))
                    lvi.SubItems.Add(If(ed.Visible, "Yes", "No"))
                    lvi.SubItems.Add(g.Name)
                    lvi.Tag = ed
                    lvParams.Items.Add(lvi)
                Next
            Next

            If lvParams.Items.Count > 0 Then
                lvParams.Items(0).Selected = True
            End If
        End Sub
    End Class

    '==================== 결과 폼 (CSV 내보내기 포함) ====================
    Friend Class FormPropagateReport
        Inherits WinForms.Form

        Private Class RowInfo
            Public Property Kind As String
            Public Property Family As String
            Public Property Detail As String
        End Class

        Private ReadOnly _report As String
        Private ReadOnly _rows As List(Of RowInfo)

        Private txtReport As WinForms.TextBox
        Private lv As WinForms.ListView
        Private btnClose As WinForms.Button
        Private btnCsv As WinForms.Button

        Public Sub New(report As String,
                       scanFails As IEnumerable(Of String),
                       skips As IEnumerable(Of String),
                       parentFails As IEnumerable(Of String),
                       childFails As IEnumerable(Of String))
            _report = report
            _rows = New List(Of RowInfo)()
            BuildRows(scanFails, "ScanFail")
            BuildRows(skips, "Skip")
            BuildRows(parentFails, "Error")
            BuildRows(childFails, "ChildError")
            BuildUI()
        End Sub

        Private Sub BuildRows(items As IEnumerable(Of String), kind As String)
            If items Is Nothing Then Return
            For Each s In items
                If String.IsNullOrWhiteSpace(s) Then Continue For
                Dim fam As String = s
                Dim detail As String = String.Empty
                Dim parts = s.Split(New Char() {":"c}, 2)
                If parts.Length = 2 Then
                    fam = parts(0).Trim()
                    detail = parts(1).Trim()
                End If
                _rows.Add(New RowInfo With {.Kind = kind, .Family = fam, .Detail = detail})
            Next
        End Sub

        Private Sub BuildUI()
            Me.Text = "KKY Param Propagator - Report"
            Me.Width = 900
            Me.Height = 720
            Me.StartPosition = WinForms.FormStartPosition.CenterScreen

            txtReport = New WinForms.TextBox() With {
                .Left = 10,
                .Top = 10,
                .Width = 860,
                .Height = 260,
                .Multiline = True,
                .ScrollBars = WinForms.ScrollBars.Vertical,
                .ReadOnly = True,
                .Font = New Drawing.Font("Consolas", 9.0F),
                .Text = _report
            }

            lv = New WinForms.ListView() With {
                .Left = 10,
                .Top = 280,
                .Width = 860,
                .Height = 340,
                .View = WinForms.View.Details,
                .FullRowSelect = True
            }
            lv.Columns.Add("Type", 100)
            lv.Columns.Add("Family", 400)
            lv.Columns.Add("Detail", 340)

            For Each r In _rows
                Dim item As New WinForms.ListViewItem(r.Kind)
                item.SubItems.Add(r.Family)
                item.SubItems.Add(r.Detail)
                lv.Items.Add(item)
            Next

            btnClose = New WinForms.Button() With {.Text = "닫기", .Width = 100, .Left = 770, .Top = 630}
            AddHandler btnClose.Click, Sub() Me.Close()

            btnCsv = New WinForms.Button() With {.Text = "CSV 내보내기", .Width = 120, .Left = 640, .Top = 630}
            AddHandler btnCsv.Click, AddressOf OnExportCsv

            Controls.AddRange(New WinForms.Control() {txtReport, lv, btnCsv, btnClose})
        End Sub

        Private Sub OnExportCsv(sender As Object, e As EventArgs)
            Using dlg As New WinForms.SaveFileDialog()
                dlg.Filter = "CSV 파일 (*.csv)|*.csv|모든 파일 (*.*)|*.*"
                dlg.FileName = "KKY_ParamPropagator_Report.csv"
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                Try
                    Using sw As New StreamWriter(dlg.FileName, False, Encoding.UTF8)
                        sw.WriteLine("Type,FamilyName,Detail")
                        For Each r In _rows
                            Dim line = $"{EscapeCsv(r.Kind)},{EscapeCsv(r.Family)},{EscapeCsv(r.Detail)}"
                            sw.WriteLine(line)
                        Next
                    End Using
                    WinForms.MessageBox.Show("CSV 내보내기가 완료되었습니다.", "KKY Param Propagator")
                Catch ex As Exception
                    WinForms.MessageBox.Show($"CSV 내보내기 실패: {ex.Message}", "KKY Param Propagator")
                End Try
            End Using
        End Sub

        Private Function EscapeCsv(s As String) As String
            If s Is Nothing Then Return String.Empty

            Dim quote As String = """"
            Dim needsQuote As Boolean =
                (s.Contains(",") OrElse s.Contains(quote) OrElse s.Contains(vbCr) OrElse s.Contains(vbLf))

            Dim val As String = s.Replace(quote, quote & quote)

            If needsQuote Then
                Return quote & val & quote
            End If

            Return val
        End Function
    End Class

    '==================== 메인 서비스 ====================
    Public Class ParamPropagateService

        Public Enum RunStatus
            Succeeded
            Cancelled
            Failed
        End Enum

        Public Class SharedParamDefinitionDto
            Public Property GroupName As String
            Public Property Name As String
            Public Property ParamType As String
            Public Property Visible As Boolean
        End Class

        Public Class ParameterGroupOption
            Public Property Id As Integer
            Public Property Name As String
        End Class

        Public Class SharedParamListResponse
            Public Property Ok As Boolean
            Public Property Message As String
            Public Property Definitions As List(Of SharedParamDefinitionDto)
            Public Property TargetGroups As List(Of ParameterGroupOption)
        End Class

        Public Class SharedParamRunRequest
            Public Property ParamNames As List(Of String)
            Public Property TargetGroup As Integer
            Public Property IsInstance As Boolean
            Public Property ExcludeDummy As Boolean

            Public Shared Function FromPayload(payload As Object) As SharedParamRunRequest
                Dim req As New SharedParamRunRequest With {
                    .ParamNames = New List(Of String)(),
                    .TargetGroup = CInt(BuiltInParameterGroup.PG_TEXT),
                    .IsInstance = True,
                    .ExcludeDummy = True
                }

                Try
                    If payload Is Nothing Then Return req

                    Dim namesObj = ReadProp(payload, "paramNames")
                    If TypeOf namesObj Is IEnumerable Then
                        For Each n In CType(namesObj, IEnumerable)
                            If n IsNot Nothing Then req.ParamNames.Add(n.ToString())
                        Next
                    End If

                    Dim gObj = ReadProp(payload, "group")
                    If gObj IsNot Nothing Then req.TargetGroup = Convert.ToInt32(gObj)

                    Dim instObj = ReadProp(payload, "isInstance")
                    If instObj IsNot Nothing Then req.IsInstance = Convert.ToBoolean(instObj)

                    Dim dummyObj = ReadProp(payload, "excludeDummy")
                    If dummyObj IsNot Nothing Then req.ExcludeDummy = Convert.ToBoolean(dummyObj)
                Catch
                End Try

                Return req
            End Function

            Private Shared Function ReadProp(payload As Object, name As String) As Object
                If payload Is Nothing OrElse String.IsNullOrEmpty(name) Then Return Nothing

                Try
                    Dim t = payload.GetType()
                    Dim pi = t.GetProperty(name)
                    If pi IsNot Nothing Then Return pi.GetValue(payload)

                    Dim fi = t.GetField(name)
                    If fi IsNot Nothing Then Return fi.GetValue(payload)

                    Dim dict = TryCast(payload, IDictionary)
                    If dict IsNot Nothing AndAlso dict.Contains(name) Then Return dict(name)
                Catch
                End Try

                Return Nothing
            End Function
        End Class

        Public Class SharedParamDetailRow
            Public Property Kind As String
            Public Property Family As String
            Public Property Detail As String
        End Class

        Public Class SharedParamRunResult
            Public Property Status As RunStatus
            Public Property Message As String
            Public Property Report As String
            Public Property Details As List(Of SharedParamDetailRow)
        End Class

        '==================== 목록 제공 ====================
        Public Shared Function GetSharedParameterDefinitions(app As UIApplication) As SharedParamListResponse
            Dim res As New SharedParamListResponse With {
                .Ok = False,
                .Message = Nothing,
                .Definitions = New List(Of SharedParamDefinitionDto)(),
                .TargetGroups = BuildGroupOptions()
            }

            Try
                If app Is Nothing OrElse app.Application Is Nothing Then
                    res.Message = "UIApplication 이 없습니다."
                    Return res
                End If

                Dim sharedPath As String = app.Application.SharedParametersFilename
                If String.IsNullOrEmpty(sharedPath) OrElse Not File.Exists(sharedPath) Then
                    res.Message = "공유 파라미터 파일 먼저 지정해 주세요."
                    Return res
                End If

                Dim defFile = app.Application.OpenSharedParameterFile()
                If defFile Is Nothing Then
                    res.Message = "공유 파라미터 파일을 열 수 없습니다."
                    Return res
                End If

                For Each g As DefinitionGroup In defFile.Groups
                    For Each d As Definition In g.Definitions
                        Dim ed = TryCast(d, ExternalDefinition)
                        If ed Is Nothing Then Continue For
                        res.Definitions.Add(New SharedParamDefinitionDto With {
                            .GroupName = g.Name,
                            .Name = ed.Name,
                            .ParamType = SharedParamTypeHelper.GetParamTypeString(ed),
                            .Visible = ed.Visible
                        })
                    Next
                Next

                res.Ok = True
            Catch ex As Exception
                res.Ok = False
                res.Message = ex.Message
            End Try

            Return res
        End Function

        '==================== 실행 엔트리 ====================
        Private Class ProgressDispatcher
            Private ReadOnly _cb As Action(Of String, Double, Integer, Integer, String, String)
            Private ReadOnly _minMs As Integer
            Private _lastPhase As String = String.Empty
            Private _lastSent As DateTime = DateTime.MinValue

            Public Sub New(cb As Action(Of String, Double, Integer, Integer, String, String), Optional minIntervalMs As Integer = 180)
                _cb = cb
                _minMs = Math.Max(50, minIntervalMs)
            End Sub

            Public Sub Report(phase As String,
                              phaseProgress As Double,
                              current As Integer,
                              total As Integer,
                              message As String,
                              target As String,
                              Optional force As Boolean = False)
                If _cb Is Nothing Then Return
                Dim now As DateTime = DateTime.UtcNow
                Dim elapsed As Double = (now - _lastSent).TotalMilliseconds
                Dim phaseChanged As Boolean = Not String.Equals(_lastPhase, phase, StringComparison.OrdinalIgnoreCase)
                If Not force AndAlso Not phaseChanged AndAlso elapsed < _minMs Then Return
                _lastPhase = phase
                _lastSent = now
                Dim safeProg As Double = Math.Max(0.0R, Math.Min(1.0R, phaseProgress))
                _cb.Invoke(phase, safeProg, current, total, message, target)
            End Sub
        End Class

        Public Shared Function Run(app As UIApplication,
                                   request As SharedParamRunRequest,
                                   Optional progress As Action(Of String, Double, Integer, Integer, String, String) = Nothing) As SharedParamRunResult
            Dim result As New SharedParamRunResult With {
                .Status = RunStatus.Failed,
                .Details = New List(Of SharedParamDetailRow)()
            }

            If app Is Nothing Then
                result.Message = "UIApplication 이 없습니다."
                Return result
            End If

            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                result.Message = "활성 문서가 없습니다."
                Return result
            End If

            Dim doc As Document = uiDoc.Document
            If doc.IsFamilyDocument Then
                result.Message = "프로젝트 문서에서 실행하세요."
                Return result
            End If

            Dim reporter As New ProgressDispatcher(progress)

            Dim sharedPath As String = app.Application.SharedParametersFilename
            If String.IsNullOrEmpty(sharedPath) OrElse Not File.Exists(sharedPath) Then
                result.Message = "공유 파라미터 파일 먼저 지정해 주세요."
                Return result
            End If

            If request Is Nothing OrElse request.ParamNames Is Nothing OrElse request.ParamNames.Count = 0 Then
                result.Message = "선택된 공유 파라미터가 없습니다."
                result.Status = RunStatus.Cancelled
                Return result
            End If

            Dim chosenPG As BuiltInParameterGroup = BuiltInParameterGroup.PG_TEXT
            Try
                chosenPG = CType(request.TargetGroup, BuiltInParameterGroup)
            Catch
                chosenPG = BuiltInParameterGroup.PG_TEXT
            End Try

            Dim extDefs As List(Of ExternalDefinition) = ResolveDefinitions(app.Application, request.ParamNames)
            If extDefs Is Nothing OrElse extDefs.Count = 0 Then
                result.Message = "선택한 공유 파라미터를 Shared Parameters 파일에서 찾을 수 없습니다."
                Return result
            End If

            Dim status As RunStatus =
                ExecuteCore(doc, extDefs, request.ParamNames, request.ExcludeDummy, chosenPG, request.IsInstance, result, reporter)

            result.Status = status
            If String.IsNullOrEmpty(result.Message) Then
                result.Message = If(status = RunStatus.Succeeded,
                                    "공유 파라미터 연동을 완료했습니다.",
                                    "공유 파라미터 연동에 실패했습니다.")
            End If

            Return result
        End Function

        Public Shared Function RunOnDocument(app As UIApplication,
                                             doc As Document,
                                             request As SharedParamRunRequest,
                                             Optional progress As Action(Of String, Double, Integer, Integer, String, String) = Nothing) As SharedParamRunResult
            Dim result As New SharedParamRunResult With {
                .Status = RunStatus.Failed,
                .Details = New List(Of SharedParamDetailRow)()
            }

            If app Is Nothing Then
                result.Message = "UIApplication 이 없습니다."
                Return result
            End If
            If doc Is Nothing Then
                result.Message = "문서가 없습니다."
                Return result
            End If
            If doc.IsFamilyDocument Then
                result.Message = "프로젝트 문서에서 실행하세요."
                Return result
            End If

            Dim reporter As New ProgressDispatcher(progress)

            Dim sharedPath As String = app.Application.SharedParametersFilename
            If String.IsNullOrEmpty(sharedPath) OrElse Not File.Exists(sharedPath) Then
                result.Message = "공유 파라미터 파일 먼저 지정해 주세요."
                Return result
            End If

            If request Is Nothing OrElse request.ParamNames Is Nothing OrElse request.ParamNames.Count = 0 Then
                result.Message = "선택된 공유 파라미터가 없습니다."
                result.Status = RunStatus.Cancelled
                Return result
            End If

            Dim chosenPG As BuiltInParameterGroup = BuiltInParameterGroup.PG_TEXT
            Try
                chosenPG = CType(request.TargetGroup, BuiltInParameterGroup)
            Catch
                chosenPG = BuiltInParameterGroup.PG_TEXT
            End Try

            Dim extDefs As List(Of ExternalDefinition) = ResolveDefinitions(app.Application, request.ParamNames)
            If extDefs Is Nothing OrElse extDefs.Count = 0 Then
                result.Message = "선택한 공유 파라미터를 Shared Parameters 파일에서 찾을 수 없습니다."
                Return result
            End If

            Dim status As RunStatus =
                ExecuteCore(doc, extDefs, request.ParamNames, request.ExcludeDummy, chosenPG, request.IsInstance, result, reporter)

            result.Status = status
            If String.IsNullOrEmpty(result.Message) Then
                result.Message = If(status = RunStatus.Succeeded,
                                    "공유 파라미터 연동을 완료했습니다.",
                                    "공유 파라미터 연동에 실패했습니다.")
            End If

            Return result
        End Function

        '==================== 결과를 엑셀로 ====================
        Public Shared Function ExportResultToExcel(result As SharedParamRunResult, Optional doAutoFit As Boolean = False) As String
            If result Is Nothing OrElse result.Details Is Nothing OrElse result.Details.Count = 0 Then Return String.Empty

            Dim defaultName As String = $"ParamProp_{Date.Now:yyMMdd_HHmmss}.xlsx"

            Using sfd As New WinForms.SaveFileDialog()
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                sfd.FileName = defaultName
                If sfd.ShowDialog() <> WinForms.DialogResult.OK Then Return String.Empty

                Dim dt As New DataTable("SharedParamPropagate")
                dt.Columns.Add("Type")
                dt.Columns.Add("Family")
                dt.Columns.Add("Detail")

                For Each r In result.Details
                    Dim row = dt.NewRow()
                    row("Type") = r.Kind
                    row("Family") = r.Family
                    row("Detail") = r.Detail
                    dt.Rows.Add(row)
                Next

                Infrastructure.ExcelCore.SaveXlsx(sfd.FileName, "Results", dt, doAutoFit, "paramprop:progress")
                Return sfd.FileName
            End Using
        End Function

        '==================== 핵심 실행 (이름 기반 계층 + 역순 처리) ====================
        Private Shared Function ExecuteCore(doc As Document,
                                            extDefs As List(Of ExternalDefinition),
                                            paramNames As List(Of String),
                                            excludeDummy As Boolean,
                                            chosenPG As BuiltInParameterGroup,
                                            chosenIsInstance As Boolean,
                                            result As SharedParamRunResult,
                                            reporter As ProgressDispatcher) As RunStatus

            ' 1. 편집 가능한 모든 패밀리 수집 (프로젝트 문서 기준)
            Dim allEditableIds As List(Of ElementId) =
    New FilteredElementCollector(doc).
        OfClass(GetType(Family)).
        Cast(Of Family)().
        Where(Function(x) x IsNot Nothing AndAlso x.IsEditable).
        Select(Function(x) x.Id).
        ToList()

            Dim totalEditableCount As Integer = allEditableIds.Count
            Dim scanIndex As Integer = 0
            reporter.Report("COLLECT", 0.0R, 0, totalEditableCount, "패밀리 스캔 준비", String.Empty, True)

            ' 이름 → ElementId 매핑 (패밀리 객체 캐싱 금지)
            Dim nameToFamilyId As New Dictionary(Of String, ElementId)(StringComparer.OrdinalIgnoreCase)

            ' 부모이름 → 자식이름 그래프 (공유 체크된 하위만)
            Dim parentToChildren As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            Dim allTargetNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim childNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Dim dummyExcludedCount As Integer = 0
            Dim scanFails As New List(Of String)()

            '----- 1차 스캔: 그래프 구성 -----
            For Each famId As ElementId In allEditableIds
                Dim f As Family = Nothing
                Try
                    f = TryCast(doc.GetElement(famId), Family)
                Catch
                End Try
                If f Is Nothing OrElse f.FamilyCategory Is Nothing Then Continue For
                If IsAnnotationFamily(f) Then Continue For

                Dim famName As String = f.Name
                If Not nameToFamilyId.ContainsKey(famName) Then
                    nameToFamilyId.Add(famName, f.Id)
                End If

                Dim hostDoc As Document = Nothing
                Try
                    hostDoc = doc.EditFamily(f)

                    Dim hostName As String = f.Name

                    Dim insts = New FilteredElementCollector(hostDoc).
                                OfClass(GetType(FamilyInstance)).
                                Cast(Of FamilyInstance)()

                    For Each fi As FamilyInstance In insts
                        Dim sym As FamilySymbol = TryCast(hostDoc.GetElement(fi.Symbol.Id), FamilySymbol)
                        If sym Is Nothing Then Continue For

                        Dim childFam As Family = sym.Family
                        If childFam Is Nothing OrElse Not childFam.IsEditable Then Continue For
                        If IsAnnotationFamily(childFam) Then Continue For

                        If excludeDummy AndAlso
                           childFam.Name.IndexOf("Dummy", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            dummyExcludedCount += 1
                            Continue For
                        End If

                        Dim isShared As Boolean = False
                        Try
                            isShared = SharedFamilyHelper.IsFamilyShared(hostDoc, childFam)
                        Catch
                            isShared = False
                        End Try

                        ' 공유 체크 안 된 하위 패밀리는 완전히 제외
                        If Not isShared Then Continue For

                        Dim childName As String = childFam.Name

                        Dim setChildren As HashSet(Of String) = Nothing
                        If Not parentToChildren.TryGetValue(hostName, setChildren) Then
                            setChildren = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                            parentToChildren(hostName) = setChildren
                        End If
                        setChildren.Add(childName)

                        allTargetNames.Add(hostName)
                        allTargetNames.Add(childName)
                        childNames.Add(childName)
                    Next

                Catch ex As Exception
                    If Not IsNoTxnNoise(ex.Message) Then
                        scanFails.Add(famName)
                    End If
                Finally
                    If hostDoc IsNot Nothing Then
                        Try : hostDoc.Close(False) : Catch : End Try
                    End If
                End Try

                scanIndex += 1
                reporter.Report("COLLECT",
                                If(totalEditableCount = 0, 1.0R, CDbl(scanIndex) / CDbl(Math.Max(1, totalEditableCount))),
                                scanIndex,
                                totalEditableCount,
                                "패밀리 스캔 중",
                                famName)
            Next

            ' 복합 패밀리(상위) 목록 (리포트용)
            Dim compositeFamilyNames As New List(Of String)(parentToChildren.Keys)

            '----- 2. 그래프를 하위 → 상위 순으로 정렬 (DFS, 이름 기준) -----
            Dim order As New List(Of String)()
            Dim mark As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase) ' 0:미방문,1:방문중,2:완료

            Dim dfs As Action(Of String) =
                Sub(name As String)
                    Dim st As Integer = 0
                    If mark.TryGetValue(name, st) Then
                        If st = 2 Then Return
                        If st = 1 Then
                            ' 순환 구조가 있다면 더 들어가지 않고 끊음
                            Return
                        End If
                    End If

                    mark(name) = 1
                    Dim childs As HashSet(Of String) = Nothing
                    If parentToChildren.TryGetValue(name, childs) Then
                        For Each cn In childs
                            dfs(cn)
                        Next
                    End If
                    mark(name) = 2
                    order.Add(name)   ' 자식 먼저 처리 후 자기 자신 ⇒ 하위 → 상위
                End Sub

            For Each name In allTargetNames
                dfs(name)
            Next

            Dim totalToProcess As Integer = order.Count
            reporter.Report("ANALYZE", 1.0R, totalToProcess, totalToProcess, "그래프 분석 완료", String.Empty, True)

            '----- 3. 계층 역순으로 파라미터 추가 & 연동 -----
            Dim fatalEx As Exception = Nothing
            Dim addedChild As Integer = 0
            Dim addedHost As Integer = 0
            Dim linkCnt As Integer = 0
            Dim verifyOk As Integer = 0
            Dim verifyFail As Integer = 0
            Dim skipTotal As Integer = 0
            Dim parentFails As New List(Of String)()
            Dim childFails As New List(Of String)()
            Dim skips As New List(Of String)()
            Dim compositeSuccessCount As Integer = 0
            Dim applyIndex As Integer = 0
            Dim lastErrorMessage As String = String.Empty
            reporter.Report("APPLY", 0.0R, 0, totalToProcess, "파라미터 적용 준비", String.Empty, True)

            Using tgAll As New TransactionGroup(doc, "KKY Shared Param Propagate")
                tgAll.Start()
                Try
                    ' order: 하위 → 상위. leaf(A-6) 먼저 처리.
                    For Each famName In order
                        Dim famId As ElementId = Nothing
                        If Not nameToFamilyId.TryGetValue(famName, famId) Then
                            famId = FindFamilyIdByName(doc, famName)
                            If famId IsNot Nothing Then nameToFamilyId(famName) = famId
                        End If

                        Dim isParent As Boolean = parentToChildren.ContainsKey(famName)
                        Dim isChild As Boolean = childNames.Contains(famName)

                        Try
                            ProcessFamilyBottomUp(doc,
                                                  famId,
                                                  famName,
                                                  extDefs,
                                                  paramNames,
                                                  parentToChildren,
                                                  excludeDummy,
                                                  chosenIsInstance,
                                                  chosenPG,
                                                  isParent,
                                                  isChild,
                                                  addedHost,
                                                  addedChild,
                                                  linkCnt,
                                                  verifyOk,
                                                  verifyFail,
                                                  skipTotal,
                                                  parentFails,
                                                  childFails,
                                                  skips,
                                                  compositeSuccessCount,
                                                  nameToFamilyId)
                        Catch ex As Exception
                            lastErrorMessage = MakePhaseError("APPLY", famName, famId, ex.Message)
                            Throw
                        End Try

                        applyIndex += 1
                        reporter.Report("APPLY",
                                        If(totalToProcess = 0, 1.0R, CDbl(applyIndex) / CDbl(Math.Max(1, totalToProcess))),
                                        applyIndex,
                                        totalToProcess,
                                        "파라미터 적용 중",
                                        famName)
                    Next

                    tgAll.Assimilate()
                Catch ex As Exception
                    fatalEx = ex
                    Try : tgAll.RollBack() : Catch : End Try
                End Try
            End Using

            reporter.Report("SAVE", 1.0R, applyIndex, totalToProcess, "결과 저장/정리 중", String.Empty, True)

            If fatalEx IsNot Nothing Then
                result.Message = fatalEx.Message
                Return RunStatus.Failed
            End If

            '----- 4. 리포트 구성 -----
            Dim header As New StringBuilder()
            header.AppendLine($"패밀리 스캔: {totalEditableCount}개 / 복합 패밀리: {compositeFamilyNames.Count}개 / 성공: {compositeSuccessCount}개")
            header.AppendLine($"하위 패밀리 파라미터 추가: {addedChild}")
            header.AppendLine($"복합 패밀리 파라미터 추가/교정: {addedHost}")
            header.AppendLine($"파라미터 연동 성공 카운트: {linkCnt}")
            header.AppendLine($"연동 검증 OK: {verifyOk} / 미연동: {verifyFail}")
            header.AppendLine($"스킵(연동 불가) 건수: {skipTotal}")
            header.AppendLine($"선택 모드: {(If(chosenIsInstance, "인스턴스", "타입"))}")
            header.AppendLine($"Dummy 제외된 하위 패밀리 수: {dummyExcludedCount}")

            Dim skipLines = skips.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            Dim failLines = parentFails.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            Dim scanLines = scanFails.Distinct(StringComparer.OrdinalIgnoreCase).ToList()

            Dim hasDetail As Boolean = (scanLines.Count + skipLines.Count + failLines.Count + childFails.Count > 0)
            If hasDetail Then
                header.AppendLine().AppendLine("상세:")
                If scanLines.Count > 0 Then
                    header.AppendLine("스캔 실패(패밀리 편집 불가) 목록:")
                    header.AppendLine(String.Join(vbCrLf, scanLines.Take(80)))
                    If scanLines.Count > 80 Then header.AppendLine("...")
                End If
                If skipLines.Count > 0 Then
                    header.AppendLine("스킵(연동 불가) 목록:")
                    header.AppendLine(String.Join(vbCrLf, skipLines.Take(80)))
                    If skipLines.Count > 80 Then header.AppendLine("...")
                End If
                If failLines.Count + childFails.Count > 0 Then
                    header.AppendLine("오류 목록:")
                    Dim failMerged As New List(Of String)()
                    failMerged.AddRange(failLines)
                    failMerged.AddRange(childFails)
                    header.AppendLine(String.Join(vbCrLf, failMerged.Take(80)))
                    If failMerged.Count > 80 Then header.AppendLine("...")
                End If
            End If

            result.Report = header.ToString()
            result.Details = BuildDetails(scanLines, skipLines, failLines, childFails)
            reporter.Report("DONE", 1.0R, applyIndex, totalToProcess, "완료", String.Empty, True)
            If String.IsNullOrWhiteSpace(result.Message) AndAlso Not String.IsNullOrWhiteSpace(lastErrorMessage) Then
                result.Message = lastErrorMessage
            End If
            Return RunStatus.Succeeded
        End Function

        ' 패밀리 1개 단위 처리 (계층 역순에서 호출)
        Private Shared Sub ProcessFamilyBottomUp(projDoc As Document,
                                                 famId As ElementId,
                                                 famName As String,
                                                 extDefs As IEnumerable(Of ExternalDefinition),
                                                 paramNames As List(Of String),
                                                 parentToChildren As Dictionary(Of String, HashSet(Of String)),
                                                 excludeDummy As Boolean,
                                                 chosenIsInstance As Boolean,
                                                 chosenPG As BuiltInParameterGroup,
                                                 isParent As Boolean,
                                                 isChild As Boolean,
                                                 ByRef addedHost As Integer,
                                                 ByRef addedChild As Integer,
                                                 ByRef linkCnt As Integer,
                                                 ByRef verifyOk As Integer,
                                                 ByRef verifyFail As Integer,
                                                 ByRef skipTotal As Integer,
                                                 parentFails As List(Of String),
                                                 childFails As List(Of String),
                                                 skips As List(Of String),
                                                 ByRef compositeSuccessCount As Integer,
                                                 nameToFamilyId As Dictionary(Of String, ElementId))

            If projDoc Is Nothing OrElse String.IsNullOrWhiteSpace(famName) Then Return

            Dim fam As Family = TryGetFreshFamily(projDoc, famName, famId, nameToFamilyId)

            If fam Is Nothing OrElse IsAnnotationFamily(fam) Then
                Dim missingMsg = MakePhaseError("APPLY", famName, famId, "패밀리 인스턴스를 찾을 수 없습니다.")
                If isParent Then
                    parentFails.Add(missingMsg)
                End If
                If isChild Then
                    childFails.Add(missingMsg)
                End If
                Return
            End If

            Dim safeFamId As ElementId = If(fam IsNot Nothing, fam.Id, famId)
            Dim hasChildren As Boolean = parentToChildren.ContainsKey(famName)

            Dim famDoc As Document = Nothing
            Dim ok As Boolean = True
            Dim localAdded As Integer = 0
            Dim localSkipAssoc As Integer = 0

            Try
                Try
                    famDoc = projDoc.EditFamily(fam)
                Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                    ok = False
                    Dim msg = MakeInvalidObjectMessage("APPLY", famName, safeFamId, "EditFamily", ex)
                    If isParent Then
                        parentFails.Add(msg)
                    ElseIf isChild Then
                        childFails.Add(msg)
                    End If
                    Return
                End Try

                If famDoc Is Nothing Then
                    parentFails.Add(MakePhaseError("APPLY", famName, safeFamId, "패밀리 편집에 실패했습니다."))
                    Return
                End If
                Dim fm As FamilyManager = famDoc.FamilyManager

                ' 1) 이 패밀리에 공유 파라미터 추가/교정
                For Each ed In extDefs
                    Dim res = EnsureSharedParamInFamily(famDoc, ed, chosenIsInstance, chosenPG)
                    If res.Added Then localAdded += 1
                    If Not String.IsNullOrEmpty(res.ErrorMessage) Then
                        ok = False
                        If isParent Then
                            parentFails.Add(famName & ": " & res.ErrorMessage)
                        ElseIf isChild Then
                            childFails.Add(famName & ": " & res.ErrorMessage)
                        End If
                    End If
                    Try : famDoc.Regenerate() : Catch : End Try
                Next

                If localAdded > 0 Then
                    If isChild Then addedChild += localAdded
                    If isParent Then addedHost += localAdded
                End If

                ' 2) 자식이 있는 패밀리라면 하위 인스턴스와 연동
                If hasChildren Then
                    Dim hostParams As New Dictionary(Of String, FamilyParameter)(StringComparer.OrdinalIgnoreCase)
                    For Each p As FamilyParameter In fm.Parameters
                        If p.Definition IsNot Nothing Then
                            Dim nm = p.Definition.Name
                            If paramNames.Contains(nm, StringComparer.OrdinalIgnoreCase) Then
                                hostParams(nm) = p
                            End If
                        End If
                    Next

                    Dim childrenOfHost As HashSet(Of String) = Nothing
                    parentToChildren.TryGetValue(famName, childrenOfHost)

                    Using t As New Transaction(famDoc, "KKY: Associate nested shared params")
                        t.Start()

                        Dim insts = New FilteredElementCollector(famDoc).
                                    OfClass(GetType(FamilyInstance)).
                                    Cast(Of FamilyInstance)().
                                    ToList()

                        For Each fi In insts
                            Dim childF As Family = Nothing
                            Try
                                Dim sym As FamilySymbol = TryCast(famDoc.GetElement(fi.Symbol.Id), FamilySymbol)
                                If sym Is Nothing Then Continue For
                                childF = sym.Family
                            Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                                ok = False
                                Dim msg = MakeInvalidObjectMessage("APPLY", famName, safeFamId, "Associate:ResolveChildSymbol", ex)
                                If isParent Then
                                    parentFails.Add(msg)
                                ElseIf isChild Then
                                    childFails.Add(msg)
                                End If
                                Continue For
                            Catch ex As Exception
                                ok = False
                                parentFails.Add(MakePhaseError("APPLY", famName, safeFamId, $"[Associate] 오류 - {ex.Message}"))
                                Continue For
                            End Try

                            If childF Is Nothing Then Continue For
                            If IsAnnotationFamily(childF) Then Continue For
                            If ShouldSkipDummy(excludeDummy, childF) Then Continue For

                            ' 그래프상 이 호스트의 자식으로 등록된 패밀리만 대상
                            If childrenOfHost IsNot Nothing AndAlso childrenOfHost.Count > 0 AndAlso
                               Not childrenOfHost.Contains(childF.Name) Then
                                Continue For
                            End If

                            For Each name In paramNames
                                Try
                                    Dim hostParam As FamilyParameter = Nothing
                                    If Not hostParams.TryGetValue(name, hostParam) OrElse hostParam Is Nothing Then
                                        Continue For
                                    End If

                                    Dim p As Parameter = TryGetElementParameterByName(fi, name)
                                    If p Is Nothing Then
                                        ' 인스턴스에 파라미터가 없어서 스킵: 건수만 카운트
                                        localSkipAssoc += 1
                                        Continue For
                                    End If
                                    If p.IsReadOnly Then
                                        ' 읽기 전용이라 스킵: 건수만 카운트
                                        localSkipAssoc += 1
                                        Continue For
                                    End If
                                    If Not famDoc.FamilyManager.CanElementParameterBeAssociated(p) Then
                                        ' 연동 불가라 스킵: 건수만 카운트
                                        localSkipAssoc += 1
                                        Continue For
                                    End If

                                    Try
                                        famDoc.FamilyManager.AssociateElementParameterToFamilyParameter(p, hostParam)
                                        linkCnt += 1
                                    Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                                        ok = False
                                        Dim msg = MakeInvalidObjectMessage("APPLY", famName, safeFamId, $"Associate:{name}", ex)
                                        If isParent Then
                                            parentFails.Add(msg)
                                        ElseIf isChild Then
                                            childFails.Add(msg)
                                        End If
                                    Catch ex As Exception
                                        ok = False
                                        parentFails.Add(MakePhaseError("APPLY", famName, safeFamId, $"[Associate] 실패 - {name}: {ex.Message}"))
                                    End Try
                                Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                                    ok = False
                                    Dim msg = MakeInvalidObjectMessage("APPLY", famName, safeFamId, $"Associate:{name}", ex)
                                    If isParent Then
                                        parentFails.Add(msg)
                                    ElseIf isChild Then
                                        childFails.Add(msg)
                                    End If
                                    Continue For
                                Catch ex As Exception
                                    ok = False
                                    parentFails.Add(MakePhaseError("APPLY", famName, safeFamId, $"[Associate] 오류 - {ex.Message}"))
                                    Continue For
                                End Try
                            Next
                        Next

                        t.Commit()
                    End Using

                    Try : famDoc.Regenerate() : Catch : End Try

                    ' 3) 연동 검증 (그래프상의 자식만 대상으로)
                    Dim v = VerifyAssociations(famDoc, hostParams, paramNames, excludeDummy, childrenOfHost)
                    verifyOk += v.Ok : verifyFail += v.Fail

                    Dim hasAnySuccess As Boolean = (v.Ok > 0)
                    Dim localSkipTotal As Integer = localSkipAssoc + v.Skip

                    If Not hasAnySuccess Then
                        skipTotal += localSkipTotal
                    End If

                    ' 패밀리 안에서 한 번이라도 연동이 성공했다면
                    ' 상위 "Error / Skip 목록"에는 올리지 않는다 (부분 실패는 숫자 카운트로만 유지)
                    If v.FailItems IsNot Nothing AndAlso v.FailItems.Count > 0 AndAlso Not hasAnySuccess Then
                        ok = False
                        parentFails.Add(famName)
                    End If
                    If v.SkipItems IsNot Nothing AndAlso v.SkipItems.Count > 0 AndAlso Not hasAnySuccess Then
                        skips.Add(famName)
                    End If
                End If

                ' 4) 프로젝트에 로드
                TxnUtil.SafeLoadFamily(famDoc, projDoc, $"KKY Load '{famName}'")
                If ok AndAlso hasChildren Then
                    compositeSuccessCount += 1
                End If

            Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                ok = False
                Dim failMsg As String = MakeInvalidObjectMessage("APPLY", famName, safeFamId, "ProcessFamilyBottomUp", ex)
                If isParent Then
                    parentFails.Add(failMsg)
                ElseIf isChild Then
                    childFails.Add(failMsg)
                End If
            Catch ex As Exception
                ok = False
                If Not IsNoTxnNoise(ex.Message) Then
                    Dim failMsg As String = MakePhaseError("APPLY", famName, safeFamId, ex.Message)
                    If isParent Then
                        parentFails.Add(failMsg)
                    ElseIf isChild Then
                        childFails.Add(failMsg)
                    End If
                End If
            Finally
                If famDoc IsNot Nothing Then
                    Try : famDoc.Close(False) : Catch : End Try
                End If
            End Try
        End Sub

        '==================== 공통 헬퍼들 ====================
        Private Structure EnsureResult
            Public Added As Boolean
            Public FinalOk As Boolean
            Public HadMismatch As Boolean
            Public ErrorMessage As String
        End Structure

        Private Shared Function EnsureSharedParamInFamily(famDoc As Document,
                                                         extDef As ExternalDefinition,
                                                         isInstance As Boolean,
                                                         groupPG As BuiltInParameterGroup) As EnsureResult
            Dim fm As FamilyManager = famDoc.FamilyManager
            Dim result As New EnsureResult With {
                .Added = False,
                .FinalOk = False,
                .HadMismatch = False,
                .ErrorMessage = Nothing
            }

            Dim anyByName As FamilyParameter =
                fm.Parameters.Cast(Of FamilyParameter)().
                FirstOrDefault(Function(p) p.Definition IsNot Nothing AndAlso
                                           String.Equals(p.Definition.Name, extDef.Name, StringComparison.OrdinalIgnoreCase))

            If anyByName Is Nothing Then
                Try
                    TxnUtil.WithTxn(famDoc, $"TEMP non-shared add: {extDef.Name}",
                        Sub()
#If REVIT2019 Or REVIT2021 Then
                            fm.AddParameter(extDef, groupPG, isInstance) ' ← shared 생성
#ElseIf REVIT2023 Or REVIT2025 Then
                            fm.AddParameter(extDef, groupPG, isInstance)
#End If
                        End Sub)
                    result.Added = True
                    Try : famDoc.Regenerate() : Catch : End Try
                    anyByName =
                        fm.Parameters.Cast(Of FamilyParameter)().
                        FirstOrDefault(Function(x) x.Definition IsNot Nothing AndAlso
                                                   String.Equals(x.Definition.Name, extDef.Name, StringComparison.OrdinalIgnoreCase))
                Catch ex As Exception
                    result.ErrorMessage = $"임시 비공유 추가 실패: {ex.Message}"
                    Return result
                End Try
            End If

            If anyByName Is Nothing Then
                result.ErrorMessage = "동일 이름 파라미터를 찾을 수 없음(임시 생성 실패)."


                Return result
            End If

            Dim replaceFailed As Boolean = False
            Try
                TxnUtil.WithTxn(famDoc, $"Replace to shared: {extDef.Name}",
                    Sub()
                        fm.ReplaceParameter(anyByName, extDef, groupPG, isInstance)
                    End Sub)
            Catch
                replaceFailed = True
            End Try

            Dim corrected As FamilyParameter =
                fm.Parameters.Cast(Of FamilyParameter)().
                FirstOrDefault(Function(x) x.Definition IsNot Nothing AndAlso
                                           String.Equals(x.Definition.Name, extDef.Name, StringComparison.OrdinalIgnoreCase))
#If REVIT2025 Then
            Dim okNow As Boolean =
                (corrected IsNot Nothing AndAlso corrected.IsShared AndAlso
                 BuiltInParameterGroupCompat.IsInGroup(corrected.Definition, groupPG))
#Else
            Dim okNow As Boolean =
                (corrected IsNot Nothing AndAlso corrected.IsShared AndAlso
                 corrected.Definition.ParameterGroup = groupPG)
#End If

            If okNow AndAlso Not replaceFailed Then
                result.FinalOk = True
                Return result
            End If

            result.HadMismatch = True
            Try
                TxnUtil.WithTxn(famDoc, $"HardFix Remove/Add: {extDef.Name}",
                    Sub()
                        Try
                            fm.RemoveParameter(anyByName)
                        Catch remEx As Exception
                            Throw New InvalidOperationException(
                                $"Remove 실패(레이블/치수/공식 참조 중일 수 있음): {remEx.Message}")
                        End Try
#If REVIT2019 Or REVIT2021 Then
                        fm.AddParameter(extDef.Name, groupPG, extDef.ParameterType, isInstance)
#ElseIf REVIT2023 Or REVIT2025 Then
                        fm.AddParameter(extDef, groupPG, isInstance)
#End If
                    End Sub)
                Try : famDoc.Regenerate() : Catch : End Try
            Catch ex As Exception
                result.ErrorMessage = $"HardFix 실패: {ex.Message}"
                Return result
            End Try

            corrected =
                fm.Parameters.Cast(Of FamilyParameter)().
                FirstOrDefault(Function(x) x.Definition IsNot Nothing AndAlso
                                           x.IsShared AndAlso
                                           String.Equals(x.Definition.Name, extDef.Name, StringComparison.OrdinalIgnoreCase))
#If REVIT2025 Then
            result.FinalOk = (corrected IsNot Nothing AndAlso
                              BuiltInParameterGroupCompat.IsInGroup(corrected.Definition, groupPG))
#Else
            result.FinalOk = (corrected IsNot Nothing AndAlso
                              corrected.Definition.ParameterGroup = groupPG)
#End If
            Return result
        End Function

        Private Structure VerifyResult
            Public Ok As Integer
            Public Fail As Integer
            Public FailItems As List(Of String)
            Public Skip As Integer
            Public SkipItems As List(Of String)
        End Structure

        Private Shared Function VerifyAssociations(hostDoc As Document,
                                                   hostParams As Dictionary(Of String, FamilyParameter),
                                                   paramNames As List(Of String),
                                                   excludeDummy As Boolean,
                                                   allowedChildNames As HashSet(Of String)) As VerifyResult

            Dim fm = hostDoc.FamilyManager
            Dim okCnt As Integer = 0
            Dim failCnt As Integer = 0
            Dim skipCnt As Integer = 0
            Dim fails As New List(Of String)()
            Dim skipItems As New List(Of String)()

            Dim insts = New FilteredElementCollector(hostDoc).
                        OfClass(GetType(FamilyInstance)).
                        Cast(Of FamilyInstance)().
                        ToList()

            For Each fi In insts
                Dim childF As Family = Nothing
                Try
                    Dim sym As FamilySymbol = TryCast(hostDoc.GetElement(fi.Symbol.Id), FamilySymbol)
                    If sym Is Nothing Then Continue For
                    childF = sym.Family
                Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                    failCnt += 1
                    fails.Add($"{fi.Id.IntegerValue}:(verify-exception: {ex.Message})")
                    Continue For
                Catch ex As Exception
                    failCnt += 1
                    fails.Add($"{fi.Id.IntegerValue}:(verify-exception: {ex.Message})")
                    Continue For
                End Try

                If childF Is Nothing Then Continue For
                If IsAnnotationFamily(childF) Then Continue For
                If ShouldSkipDummy(excludeDummy, childF) Then Continue For

                ' 그래프 상 이 호스트의 자식으로 등록된 패밀리만 검증
                If allowedChildNames IsNot Nothing AndAlso allowedChildNames.Count > 0 AndAlso
                   Not allowedChildNames.Contains(childF.Name) Then
                    Continue For
                End If

                For Each name In paramNames
                    Try
                        Dim hostParam As FamilyParameter = Nothing
                        If Not hostParams.TryGetValue(name, hostParam) OrElse hostParam Is Nothing Then
                            failCnt += 1
                            fails.Add($"{fi.Id.IntegerValue}:{name} (host-missing)")
                            Continue For
                        End If

                        Dim p As Parameter = Nothing
                        Try
                            p = TryGetElementParameterByName(fi, name)
                        Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                            failCnt += 1
                            fails.Add($"{fi.Id.IntegerValue}:{name} (verify-exception: {ex.Message})")
                            Continue For
                        Catch ex As Exception
                            failCnt += 1
                            fails.Add($"{fi.Id.IntegerValue}:{name} (verify-exception: {ex.Message})")
                            Continue For
                        End Try

                        If p Is Nothing Then
                            skipCnt += 1
                            skipItems.Add($"{fi.Id.IntegerValue}:{name} (child-inst-missing)")
                            Continue For
                        End If

                        Dim associated As FamilyParameter = Nothing
                        Try
                            associated = fm.GetAssociatedFamilyParameter(p)
                        Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                            failCnt += 1
                            fails.Add($"{fi.Id.IntegerValue}:{name} (verify-exception: {ex.Message})")
                            Continue For
                        Catch ex As Exception
                            failCnt += 1
                            fails.Add($"{fi.Id.IntegerValue}:{name} (verify-exception: {ex.Message})")
                            Continue For
                        End Try

                        If associated IsNot Nothing AndAlso associated.Id = hostParam.Id Then
                            okCnt += 1
                        Else
                            failCnt += 1
                            fails.Add($"{fi.Id.IntegerValue}:{name} (not-associated)")
                        End If

                    Catch ex As Autodesk.Revit.Exceptions.InvalidObjectException
                        failCnt += 1
                        fails.Add($"{fi.Id.IntegerValue}:{name} (verify-exception: {ex.Message})")
                        Continue For
                    Catch ex As Exception
                        failCnt += 1
                        fails.Add($"{fi.Id.IntegerValue}:{name} (verify-exception: {ex.Message})")
                        Continue For
                    End Try
                Next
            Next

            Dim r As New VerifyResult
            r.Ok = okCnt : r.Fail = failCnt : r.FailItems = fails
            r.Skip = skipCnt : r.SkipItems = skipItems
            Return r
        End Function

        Private Shared Function TryGetElementParameterByName(elem As Element,
                                                             paramName As String) As Parameter
            Dim ps As IList(Of Parameter) = elem.GetParameters(paramName)
            If ps IsNot Nothing AndAlso ps.Count > 0 Then Return ps(0)

            For Each p As Parameter In elem.Parameters
                Dim dn As Definition = p.Definition
                If dn IsNot Nothing AndAlso
                   String.Equals(dn.Name, paramName, StringComparison.OrdinalIgnoreCase) Then
                    Return p
                End If
            Next

            Return Nothing
        End Function

        Private Shared Function IsAnnotationFamily(fam As Family) As Boolean
            If fam Is Nothing Then Return False
            Dim cat = fam.FamilyCategory
            If cat Is Nothing Then Return False
            Return cat.CategoryType = CategoryType.Annotation
        End Function

        Private Shared Function ShouldSkipDummy(excludeDummy As Boolean,
                                                fam As Family) As Boolean
            If Not excludeDummy OrElse fam Is Nothing Then Return False
            Return fam.Name.IndexOf("Dummy", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Shared Function IsNoTxnNoise(msg As String) As Boolean
            If String.IsNullOrEmpty(msg) Then Return False
            Dim key As String = "Modification of the document is forbidden"
            Dim key2 As String = "no open transaction"
            Dim key3 As String = "Creation was undone"
            Return (msg.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0) OrElse
                   (msg.IndexOf(key2, StringComparison.OrdinalIgnoreCase) >= 0) OrElse
                   (msg.IndexOf(key3, StringComparison.OrdinalIgnoreCase) >= 0)
        End Function

        Private Shared Function ResolveDefinitions(app As Application,
                                                   names As IEnumerable(Of String)) As List(Of ExternalDefinition)
            Dim list As New List(Of ExternalDefinition)()
            If app Is Nothing OrElse names Is Nothing Then Return list

            Dim setNames As New HashSet(Of String)(
                names.Where(Function(n) Not String.IsNullOrWhiteSpace(n)),
                StringComparer.OrdinalIgnoreCase)
            If setNames.Count = 0 Then Return list

            Dim defFile = app.OpenSharedParameterFile()
            If defFile Is Nothing Then Return list

            For Each g As DefinitionGroup In defFile.Groups
                For Each d As Definition In g.Definitions
                    Dim ed = TryCast(d, ExternalDefinition)
                    If ed Is Nothing Then Continue For
                    If setNames.Contains(ed.Name) Then
                        list.Add(ed)
                    End If
                Next
            Next

            Return list
        End Function

        Private Shared Function BuildGroupOptions() As List(Of ParameterGroupOption)
            Dim items As New List(Of ParameterGroupOption)()
            Dim preferred As BuiltInParameterGroup() = {
                BuiltInParameterGroup.PG_TEXT,
                BuiltInParameterGroup.PG_IDENTITY_DATA,
                BuiltInParameterGroup.PG_DATA,
                BuiltInParameterGroup.PG_CONSTRAINTS
            }

            Dim added As New HashSet(Of Integer)()
            For Each pg In preferred
                items.Add(New ParameterGroupOption With {
                    .Id = CInt(pg),
                    .Name = pg.ToString()
                })
                added.Add(CInt(pg))
            Next

            For Each pg As BuiltInParameterGroup In [Enum].GetValues(GetType(BuiltInParameterGroup))
                If Not added.Contains(CInt(pg)) Then
                    items.Add(New ParameterGroupOption With {
                        .Id = CInt(pg),
                        .Name = pg.ToString()
                    })
                End If
            Next

            Return items
        End Function

        Private Shared Function BuildDetails(scanFails As IEnumerable(Of String),
                                             skips As IEnumerable(Of String),
                                             parentFails As IEnumerable(Of String),
                                             childFails As IEnumerable(Of String)) As List(Of SharedParamDetailRow)
            Dim rows As New List(Of SharedParamDetailRow)()
            AddRows(rows, scanFails, "ScanFail")
            AddRows(rows, skips, "Skip")
            AddRows(rows, parentFails, "Error")
            AddRows(rows, childFails, "ChildError")
            Return rows
        End Function

        Private Shared Sub AddRows(rows As List(Of SharedParamDetailRow),
                                   items As IEnumerable(Of String),
                                   kind As String)
            If rows Is Nothing OrElse items Is Nothing Then Return
            For Each s In items
                If String.IsNullOrWhiteSpace(s) Then Continue For
                Dim fam As String = s
                Dim detail As String = String.Empty
                Dim parts = s.Split(New Char() {":"c}, 2)
                If parts.Length = 2 Then
                    fam = parts(0).Trim()
                    detail = parts(1).Trim()
                End If
                rows.Add(New SharedParamDetailRow With {
                    .Kind = kind,
                    .Family = fam,
                    .Detail = detail
                })
            Next
        End Sub

    End Class

End Namespace
