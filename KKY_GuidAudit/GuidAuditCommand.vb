Option Explicit On
Option Strict On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Reflection

Imports WinForms = System.Windows.Forms

Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

<Transaction(TransactionMode.ReadOnly)>
Public Class CmdGuidAudit
    Implements IExternalCommand

    Public Function Execute(commandData As ExternalCommandData,
                            ByRef message As String,
                            elements As ElementSet) As Result Implements IExternalCommand.Execute
        Try
            Dim uiApp = commandData.Application
            Dim doc = uiApp.ActiveUIDocument?.Document

            If doc Is Nothing Then
                TaskDialog.Show("GUID Audit", "활성 문서가 없습니다. RVT를 열고 다시 실행해 주세요.")
                Return Result.Cancelled
            End If

            Using f As New GuidAuditForm(uiApp, doc)
                Dim owner As WinForms.IWin32Window = RevitWin32Window.FromRevit(uiApp)
                f.ShowDialog(owner)
            End Using

            Return Result.Succeeded

        Catch ex As Exception
            message = ex.ToString()
            Return Result.Failed
        End Try
    End Function
End Class

Friend NotInheritable Class FileItem
    Public ReadOnly Property Path As String
    Public ReadOnly Property Name As String

    Public Sub New(fullPath As String)
        Path = If(fullPath, "")
        Name = GetName(fullPath)
    End Sub

    Private Shared Function GetName(p As String) As String
        If String.IsNullOrWhiteSpace(p) Then Return "(Active/Unsaved)"
        Try
            Return IO.Path.GetFileName(p)
        Catch
            Return p
        End Try
    End Function

    Public Overrides Function ToString() As String
        Return Name
    End Function
End Class

Friend NotInheritable Class FamilyNodeTag
    Public ReadOnly Property RvtPath As String
    Public ReadOnly Property FamilyName As String

    Public Sub New(rvtPath As String, famName As String)
        Me.RvtPath = If(rvtPath, "")
        Me.FamilyName = If(famName, "")
    End Sub
End Class

Friend Class GuidAuditForm
    Inherits WinForms.Form

    Private ReadOnly _uiApp As UIApplication
    Private ReadOnly _activeDoc As Document

    Private ReadOnly rbProject As WinForms.RadioButton
    Private ReadOnly rbFamilies As WinForms.RadioButton

    Private ReadOnly lstFiles As WinForms.ListBox
    Private ReadOnly btnAddFiles As WinForms.Button
    Private ReadOnly btnRemove As WinForms.Button
    Private ReadOnly btnClear As WinForms.Button

    Private ReadOnly btnRun As WinForms.Button
    Private ReadOnly btnExport As WinForms.Button
    Private ReadOnly btnClose As WinForms.Button
    Private ReadOnly lblStatus As WinForms.Label

    Private ReadOnly tabs As WinForms.TabControl
    Private ReadOnly tabSummary As WinForms.TabPage
    Private ReadOnly tabFamily As WinForms.TabPage

    Private ReadOnly gridSummary As WinForms.DataGridView

    Private ReadOnly split As WinForms.SplitContainer
    Private ReadOnly tv As WinForms.TreeView
    Private ReadOnly gridDetail As WinForms.DataGridView

    Private _mode As AuditMode = AuditMode.ProjectParams
    Private _summaryMaster As DataTable = Nothing
    Private _detailMaster As DataTable = Nothing
    Private _detailView As DataView = Nothing

    Private Enum AuditMode
        ProjectParams = 1
        FamilySharedParams = 2
    End Enum

    Public Sub New(uiApp As UIApplication, doc As Document)
        _uiApp = uiApp
        _activeDoc = doc

        Text = "GUID Audit (Multi RVT)"
        StartPosition = WinForms.FormStartPosition.CenterScreen
        Width = 1280
        Height = 840

        rbProject = New WinForms.RadioButton() With {
            .Text = "1) 프로젝트 파라미터(공유/프로젝트 구분 + 공유면 GUID 비교)",
            .AutoSize = True,
            .Checked = True,
            .Left = 12,
            .Top = 12
        }

        rbFamilies = New WinForms.RadioButton() With {
            .Text = "2) 로드된 패밀리의 공유 파라미터 vs 공유파라미터 파일 GUID 비교",
            .AutoSize = True,
            .Checked = False,
            .Left = 12,
            .Top = 36
        }

        Dim lblFiles As New WinForms.Label() With {
            .Text = "대상 RVT 목록 (비우면 현재 열려있는 활성 문서만 검토)",
            .AutoSize = True,
            .Left = 12,
            .Top = 66
        }

        lstFiles = New WinForms.ListBox() With {
            .Left = 12,
            .Top = 88,
            .Width = 980,
            .Height = 110
        }

        btnAddFiles = New WinForms.Button() With {
            .Text = "RVT 추가...",
            .Left = 1004,
            .Top = 88,
            .Width = 240,
            .Height = 30
        }
        AddHandler btnAddFiles.Click, AddressOf OnAddFiles

        btnRemove = New WinForms.Button() With {
            .Text = "선택 제거",
            .Left = 1004,
            .Top = 122,
            .Width = 240,
            .Height = 30
        }
        AddHandler btnRemove.Click, AddressOf OnRemoveFiles

        btnClear = New WinForms.Button() With {
            .Text = "목록 지우기",
            .Left = 1004,
            .Top = 156,
            .Width = 240,
            .Height = 30
        }
        AddHandler btnClear.Click, AddressOf OnClearFiles

        btnRun = New WinForms.Button() With {
            .Text = "검토 실행",
            .Width = 120,
            .Height = 30,
            .Left = 12,
            .Top = 210
        }
        AddHandler btnRun.Click, AddressOf OnRun

        btnExport = New WinForms.Button() With {
            .Text = "엑셀 저장...",
            .Width = 120,
            .Height = 30,
            .Left = 140,
            .Top = 210,
            .Enabled = False
        }
        AddHandler btnExport.Click, AddressOf OnExport

        btnClose = New WinForms.Button() With {
            .Text = "닫기",
            .Width = 120,
            .Height = 30,
            .Left = 268,
            .Top = 210
        }
        AddHandler btnClose.Click, Sub() Close()

        lblStatus = New WinForms.Label() With {
            .Text = "대기 중",
            .AutoSize = False,
            .Left = 410,
            .Top = 216,
            .Width = 834,
            .Height = 20
        }

        tabs = New WinForms.TabControl() With {
            .Left = 12,
            .Top = 250,
            .Width = 1232,
            .Height = 540
        }

        tabSummary = New WinForms.TabPage("결과(요약)")
        tabFamily = New WinForms.TabPage("패밀리/파라미터")

        gridSummary = New WinForms.DataGridView() With {
            .Dock = WinForms.DockStyle.Fill,
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.DisplayedCells,
            .SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = True
        }
        tabSummary.Controls.Add(gridSummary)

        split = New WinForms.SplitContainer() With {
            .Dock = WinForms.DockStyle.Fill,
            .Orientation = WinForms.Orientation.Vertical,
            .SplitterDistance = 380
        }

        tv = New WinForms.TreeView() With {
            .Dock = WinForms.DockStyle.Fill,
            .HideSelection = False
        }
        AddHandler tv.AfterSelect, AddressOf OnTreeSelect

        gridDetail = New WinForms.DataGridView() With {
            .Dock = WinForms.DockStyle.Fill,
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.DisplayedCells,
            .SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = True
        }

        split.Panel1.Controls.Add(tv)
        split.Panel2.Controls.Add(gridDetail)
        tabFamily.Controls.Add(split)

        tabs.TabPages.Add(tabSummary)
        tabs.TabPages.Add(tabFamily)

        Controls.Add(rbProject)
        Controls.Add(rbFamilies)

        Controls.Add(lblFiles)
        Controls.Add(lstFiles)
        Controls.Add(btnAddFiles)
        Controls.Add(btnRemove)
        Controls.Add(btnClear)

        Controls.Add(btnRun)
        Controls.Add(btnExport)
        Controls.Add(btnClose)
        Controls.Add(lblStatus)

        Controls.Add(tabs)

        tabFamily.Enabled = False
    End Sub

    Private Sub OnAddFiles(sender As Object, e As EventArgs)
        Using ofd As New WinForms.OpenFileDialog()
            ofd.Filter = "Revit Project (*.rvt)|*.rvt"
            ofd.Multiselect = True
            ofd.Title = "검토할 RVT 파일 선택"
            If ofd.ShowDialog(Me) <> WinForms.DialogResult.OK Then Return

            Dim existing As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each it As Object In lstFiles.Items
                Dim fi = TryCast(it, FileItem)
                If fi IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(fi.Path) Then existing.Add(fi.Path)
            Next

            For Each p In ofd.FileNames
                If String.IsNullOrWhiteSpace(p) Then Continue For
                If existing.Contains(p) Then Continue For
                lstFiles.Items.Add(New FileItem(p))
                existing.Add(p)
            Next
        End Using
    End Sub

    Private Sub OnRemoveFiles(sender As Object, e As EventArgs)
        Dim sel = lstFiles.SelectedItems.Cast(Of Object)().ToList()
        For Each it In sel
            lstFiles.Items.Remove(it)
        Next
    End Sub

    Private Sub OnClearFiles(sender As Object, e As EventArgs)
        lstFiles.Items.Clear()
    End Sub

    Private Sub OnRun(sender As Object, e As EventArgs)
        btnRun.Enabled = False
        btnExport.Enabled = False
        tabFamily.Enabled = False
        tv.Nodes.Clear()
        gridDetail.DataSource = Nothing

        Try
            _mode = If(rbFamilies.Checked, AuditMode.FamilySharedParams, AuditMode.ProjectParams)

            lblStatus.Text = "공유 파라미터 파일 읽는 중..."
            WinForms.Application.DoEvents()

            Dim fileMap = SharedParamReader.ReadSharedParamNameGuidMap(_uiApp.Application)
            If fileMap Is Nothing Then
                Throw New InvalidOperationException("공유 파라미터 파일이 설정되어 있지 않거나 읽을 수 없습니다. (Revit 옵션에서 Shared Parameter 파일 경로 확인)")
            End If

            Dim targets As New List(Of FileItem)()
            For Each it As Object In lstFiles.Items
                Dim fi = TryCast(it, FileItem)
                If fi IsNot Nothing Then targets.Add(fi)
            Next
            If targets.Count = 0 Then
                Dim ap As String = ""
                Try : ap = _activeDoc.PathName : Catch : ap = "" : End Try
                targets.Add(New FileItem(ap))
            End If

            Dim summaryMaster As DataTable = Nothing
            Dim detailMaster As DataTable = Nothing

            Dim total As Integer = targets.Count

            For i As Integer = 0 To total - 1
                Dim target = targets(i)
                Dim path As String = If(target.Path, "")
                Dim openedByMe As Boolean = False
                Dim doc As Document = Nothing

                Try
                    lblStatus.Text = $"문서 여는 중... ({i + 1}/{total}) {target.Name}"
                    WinForms.Application.DoEvents()

                    doc = ResolveOrOpenDocument(_uiApp, _activeDoc, path, openedByMe)

                    If doc Is Nothing Then
                        Dim fail = Auditors.MakeFailureSummaryTable(_mode)
                        Auditors.AddOpenFailRow(fail, target.Name, path, If(_mode = AuditMode.ProjectParams, "Project", "Family"), "OPEN_FAIL", "문서 열기 실패")
                        summaryMaster = MergeTable(summaryMaster, fail)
                        Continue For
                    End If

                    Dim rvtName As String = GetRvtName(doc, path)

                    If _mode = AuditMode.ProjectParams Then
                        lblStatus.Text = $"프로젝트 파라미터 수집/비교 중... ({i + 1}/{total}) {rvtName}"
                        WinForms.Application.DoEvents()

                        Dim dt = Auditors.RunProjectParameterAudit(doc, fileMap, rvtName, path)
                        summaryMaster = MergeTable(summaryMaster, dt)

                    Else
                        lblStatus.Text = $"패밀리/파라미터 수집/비교 중... ({i + 1}/{total}) {rvtName}"
                        WinForms.Application.DoEvents()

                        Dim pack = Auditors.RunFamilyAudit(doc, fileMap, lblStatus, rvtName, path)
                        summaryMaster = MergeTable(summaryMaster, pack.Summary)
                        detailMaster = MergeTable(detailMaster, pack.Detail)
                    End If

                Catch ex As Exception
                    Dim fail = Auditors.MakeFailureSummaryTable(_mode)
                    Auditors.AddOpenFailRow(fail, target.Name, path, If(_mode = AuditMode.ProjectParams, "Project", "Family"), "ERROR", ex.Message)
                    summaryMaster = MergeTable(summaryMaster, fail)

                Finally
                    If openedByMe AndAlso doc IsNot Nothing Then
                        Try
                            doc.Close(False)
                        Catch
                        End Try
                    End If
                End Try
            Next

            _summaryMaster = If(summaryMaster, New DataTable())
            gridSummary.DataSource = _summaryMaster

            If _mode = AuditMode.FamilySharedParams Then
                _detailMaster = If(detailMaster, New DataTable())
                _detailView = _detailMaster.DefaultView
                gridDetail.DataSource = _detailView

                BuildTreeFromDetail(_detailMaster)
                tabFamily.Enabled = True
                tabs.SelectedTab = tabSummary
            Else
                tabFamily.Enabled = False
            End If

            btnExport.Enabled = (_summaryMaster IsNot Nothing AndAlso _summaryMaster.Rows.Count > 0)
            lblStatus.Text = $"완료: {_summaryMaster.Rows.Count:#,0} rows (문서 {targets.Count:#,0}개)"

        Catch ex As Exception
            lblStatus.Text = "오류 발생"
            TaskDialog.Show("GUID Audit", ex.Message)
        Finally
            btnRun.Enabled = True
        End Try
    End Sub

    Private Sub OnExport(sender As Object, e As EventArgs)
        Dim dtToSave As DataTable = Nothing
        Dim sheetName As String = "Result"

        If tabs.SelectedTab Is tabFamily AndAlso _mode = AuditMode.FamilySharedParams Then
            dtToSave = _detailMaster
            sheetName = "FamilyParamDetail"
        Else
            dtToSave = _summaryMaster
            sheetName = If(_mode = AuditMode.ProjectParams, "ProjectParams", "FamilySharedParams")
        End If

        If dtToSave Is Nothing OrElse dtToSave.Rows.Count = 0 Then
            TaskDialog.Show("GUID Audit", "저장할 결과가 없습니다.")
            Return
        End If

        Using sfd As New WinForms.SaveFileDialog()
            sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
            sfd.Title = "GUID Audit 결과 저장"
            sfd.FileName = $"{sheetName}_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"

            If sfd.ShowDialog(Me) <> WinForms.DialogResult.OK Then Return

            Dim ok As Boolean = False
            Dim err As String = ""

            ExcelExporter.SaveDataTableAsXlsx(dtToSave, sheetName, sfd.FileName, ok, err)

            If ok Then
                TaskDialog.Show("GUID Audit", "저장 완료 (Excel AutoFit 적용됨)" & Environment.NewLine & sfd.FileName)
            Else
                TaskDialog.Show("GUID Audit", "저장 완료 (AutoFit 미적용)" & Environment.NewLine &
                                "원인: " & err & Environment.NewLine &
                                sfd.FileName)
            End If
        End Using
    End Sub

    Private Sub OnTreeSelect(sender As Object, e As WinForms.TreeViewEventArgs)
        If _detailView Is Nothing OrElse _detailMaster Is Nothing Then Return

        Dim tag = TryCast(e.Node.Tag, FamilyNodeTag)
        If tag Is Nothing Then Return

        Dim pathEsc = EscapeForRowFilter(tag.RvtPath)

        If String.IsNullOrWhiteSpace(tag.FamilyName) Then
            _detailView.RowFilter = $"RvtPath = '{pathEsc}'"
        Else
            Dim famEsc = EscapeForRowFilter(tag.FamilyName)
            _detailView.RowFilter = $"RvtPath = '{pathEsc}' AND FamilyName = '{famEsc}'"
        End If
    End Sub

    Private Sub BuildTreeFromDetail(detail As DataTable)
        tv.BeginUpdate()
        tv.Nodes.Clear()

        If detail Is Nothing OrElse detail.Rows.Count = 0 Then
            tv.EndUpdate()
            Return
        End If

        Dim map As New Dictionary(Of String, Tuple(Of String, HashSet(Of String)))(StringComparer.OrdinalIgnoreCase)

        For Each r As DataRow In detail.Rows
            Dim p As String = Convert.ToString(r("RvtPath"))
            Dim n As String = Convert.ToString(r("RvtName"))
            Dim f As String = Convert.ToString(r("FamilyName"))

            If String.IsNullOrWhiteSpace(p) Then p = "(Active/Unsaved)"
            If String.IsNullOrWhiteSpace(n) Then n = "(Doc)"
            If String.IsNullOrWhiteSpace(f) Then Continue For

            Dim tup As Tuple(Of String, HashSet(Of String)) = Nothing
            If Not map.TryGetValue(p, tup) Then
                tup = Tuple.Create(n, New HashSet(Of String)(StringComparer.OrdinalIgnoreCase))
                map(p) = tup
            End If
            tup.Item2.Add(f)
        Next

        For Each kv In map.OrderBy(Function(x) x.Value.Item1, StringComparer.OrdinalIgnoreCase)
            Dim rvtPath = kv.Key
            Dim rvtName = kv.Value.Item1
            Dim famSet = kv.Value.Item2

            Dim fileNode As New WinForms.TreeNode(rvtName) With {
                .Tag = New FamilyNodeTag(rvtPath, "")
            }

            For Each fam In famSet.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                fileNode.Nodes.Add(New WinForms.TreeNode(fam) With {
                    .Tag = New FamilyNodeTag(rvtPath, fam)
                })
            Next

            tv.Nodes.Add(fileNode)
        Next

        tv.ExpandAll()
        tv.EndUpdate()
    End Sub

    Private Function EscapeForRowFilter(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("'", "''")
    End Function

    Private Function GetRvtName(doc As Document, path As String) As String
        If Not String.IsNullOrWhiteSpace(path) Then
            Try
                Return IO.Path.GetFileName(path)
            Catch
            End Try
        End If

        Try
            Return doc.Title
        Catch
            Return "(Doc)"
        End Try
    End Function

    Private Function MergeTable(master As DataTable, part As DataTable) As DataTable
        If part Is Nothing Then Return master
        If master Is Nothing Then master = part.Clone()
        For Each r As DataRow In part.Rows
            master.ImportRow(r)
        Next
        Return master
    End Function

    '=========================================================
    ' Central(Workshared) => Detach + CloseAllWorksets
    '=========================================================
    Private Function ResolveOrOpenDocument(uiApp As UIApplication, activeDoc As Document, path As String, ByRef openedByMe As Boolean) As Document
        openedByMe = False

        Dim activePath As String = ""
        Try : activePath = activeDoc.PathName : Catch : activePath = "" : End Try

        If String.IsNullOrWhiteSpace(path) OrElse
           (Not String.IsNullOrWhiteSpace(activePath) AndAlso String.Equals(activePath, path, StringComparison.OrdinalIgnoreCase)) Then
            Return activeDoc
        End If

        Try
            For Each d As Document In uiApp.Application.Documents
                Dim dp As String = ""
                Try : dp = d.PathName : Catch : dp = "" : End Try
                If Not String.IsNullOrWhiteSpace(dp) AndAlso String.Equals(dp, path, StringComparison.OrdinalIgnoreCase) Then
                    Return d
                End If
            Next
        Catch
        End Try

        If Not File.Exists(path) Then Return Nothing

        Dim app = uiApp.Application

        Dim mp As ModelPath = Nothing
        Try
            mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        Catch
            mp = Nothing
        End Try
        If mp Is Nothing Then Return Nothing

        Dim opt As New OpenOptions()

        Dim applyDetachCloseAll As Boolean = False
        Try
            Dim bfi = BasicFileInfo.Extract(path)
            If bfi IsNot Nothing AndAlso bfi.IsWorkshared Then
                Dim isCentral As Boolean = True
                Try
                    Dim pIsCentral = bfi.GetType().GetProperty("IsCentral", BindingFlags.Public Or BindingFlags.Instance)
                    If pIsCentral IsNot Nothing Then
                        isCentral = Convert.ToBoolean(pIsCentral.GetValue(bfi, Nothing))
                    End If
                Catch
                    isCentral = True
                End Try

                applyDetachCloseAll = isCentral
            End If
        Catch
            applyDetachCloseAll = False
        End Try

        If applyDetachCloseAll Then
            Try
                opt.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
                Dim wc As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                opt.SetOpenWorksetsConfiguration(wc)
            Catch
            End Try
        End If

        Try
            Dim d = app.OpenDocumentFile(mp, opt)
            openedByMe = True
            Return d
        Catch
            Try
                Dim opt2 As New OpenOptions()
                Dim d2 = app.OpenDocumentFile(mp, opt2)
                openedByMe = True
                Return d2
            Catch
                openedByMe = False
                Return Nothing
            End Try
        End Try
    End Function

End Class

Friend Module SharedParamReader

    Public Function ReadSharedParamNameGuidMap(app As Autodesk.Revit.ApplicationServices.Application) As Dictionary(Of String, List(Of Guid))
        Dim defFile As DefinitionFile = Nothing
        Try
            defFile = app.OpenSharedParameterFile()
        Catch
            defFile = Nothing
        End Try

        If defFile Is Nothing Then Return Nothing

        Dim map As New Dictionary(Of String, List(Of Guid))(StringComparer.OrdinalIgnoreCase)

        For Each grp As DefinitionGroup In defFile.Groups
            For Each d As Definition In grp.Definitions
                Dim g As Guid = Guid.Empty
                If Not TryGetDefinitionGuid(d, g) Then Continue For

                Dim name = d.Name
                If Not map.ContainsKey(name) Then map(name) = New List(Of Guid)()
                map(name).Add(g)
            Next
        Next

        Return map
    End Function

    Private Function TryGetDefinitionGuid(d As Definition, ByRef g As Guid) As Boolean
        g = Guid.Empty
        If d Is Nothing Then Return False

        Dim t = d.GetType()
        Dim p = t.GetProperty("GUID", BindingFlags.Public Or BindingFlags.Instance)
        If p Is Nothing Then Return False

        Dim v = p.GetValue(d, Nothing)
        If v Is Nothing Then Return False

        If TypeOf v Is Guid Then
            g = DirectCast(v, Guid)
            Return g <> Guid.Empty
        End If

        Return False
    End Function

End Module

Friend NotInheritable Class FamilyAuditPack
    Public Property Summary As DataTable
    Public Property Detail As DataTable
End Class

Friend Module Auditors

    Public Function MakeFailureSummaryTable(mode As Object) As DataTable
        Dim m As Integer = Convert.ToInt32(mode)

        If m = 1 Then
            Dim dt As New DataTable("ProjectParams")
            dt.Columns.Add("RvtName", GetType(String))
            dt.Columns.Add("Scope", GetType(String))
            dt.Columns.Add("ParamName", GetType(String))
            dt.Columns.Add("ParamKind", GetType(String))
            dt.Columns.Add("ProjectGuid", GetType(String))
            dt.Columns.Add("FileGuid", GetType(String))
            dt.Columns.Add("Result", GetType(String))
            dt.Columns.Add("Notes", GetType(String))
            Return dt
        Else
            Dim dt As New DataTable("FamilySharedParams")
            dt.Columns.Add("RvtName", GetType(String))
            dt.Columns.Add("Scope", GetType(String))
            dt.Columns.Add("FamilyName", GetType(String))
            dt.Columns.Add("FamilyCategory", GetType(String))
            dt.Columns.Add("ParamName", GetType(String))
            dt.Columns.Add("FamilyGuid", GetType(String))
            dt.Columns.Add("FileGuid", GetType(String))
            dt.Columns.Add("Result", GetType(String))
            dt.Columns.Add("Notes", GetType(String))
            Return dt
        End If
    End Function

    Public Sub AddOpenFailRow(dt As DataTable, rvtName As String, rvtPath As String, scope As String, result As String, notes As String)
        Dim r = dt.NewRow()
        If dt.Columns.Contains("RvtName") Then r("RvtName") = If(rvtName, "")
        If dt.Columns.Contains("Scope") Then r("Scope") = scope
        If dt.Columns.Contains("FamilyName") Then r("FamilyName") = ""
        If dt.Columns.Contains("FamilyCategory") Then r("FamilyCategory") = ""
        If dt.Columns.Contains("ParamName") Then r("ParamName") = ""
        If dt.Columns.Contains("ParamKind") Then r("ParamKind") = ""
        If dt.Columns.Contains("ProjectGuid") Then r("ProjectGuid") = ""
        If dt.Columns.Contains("FamilyGuid") Then r("FamilyGuid") = ""
        If dt.Columns.Contains("FileGuid") Then r("FileGuid") = ""
        If dt.Columns.Contains("Result") Then r("Result") = result
        If dt.Columns.Contains("Notes") Then r("Notes") = notes
        dt.Rows.Add(r)
    End Sub

    Public Function RunProjectParameterAudit(doc As Document,
                                             fileMap As Dictionary(Of String, List(Of Guid)),
                                             rvtName As String,
                                             rvtPath As String) As DataTable

        Dim dt As New DataTable("ProjectParams")
        dt.Columns.Add("RvtName", GetType(String))
        dt.Columns.Add("Scope", GetType(String))
        dt.Columns.Add("ParamName", GetType(String))
        dt.Columns.Add("ParamKind", GetType(String))
        dt.Columns.Add("ProjectGuid", GetType(String))
        dt.Columns.Add("FileGuid", GetType(String))
        dt.Columns.Add("Result", GetType(String))
        dt.Columns.Add("Notes", GetType(String))

        Dim pes = New FilteredElementCollector(doc).
            OfClass(GetType(ParameterElement)).
            Cast(Of ParameterElement)().
            ToList()

        For Each pe As ParameterElement In pes
            Dim name As String = SafeParamElementName(pe)
            Dim kind As String = "Project"
            Dim projGuid As String = ""
            Dim fileGuid As String = ""
            Dim result As String = ""
            Dim notes As String = ""

            If TypeOf pe Is SharedParameterElement Then
                kind = "Shared"
                Dim spe = DirectCast(pe, SharedParameterElement)
                Dim gProj = spe.GuidValue
                projGuid = gProj.ToString()

                Dim fileGuids As List(Of Guid) = Nothing
                If fileMap.TryGetValue(name, fileGuids) Then
                    fileGuid = String.Join("; ", fileGuids.Select(Function(x) x.ToString()).Distinct().ToArray())
                    If fileGuids.Count > 1 Then notes = "Shared parameter file에 동일 이름 GUID가 여러 개 존재"

                    If fileGuids.Any(Function(x) x = gProj) Then
                        result = If(fileGuids.Count > 1, "OK(MULTI_IN_FILE)", "OK")
                    Else
                        result = "MISMATCH"
                    End If
                Else
                    result = "NOT_FOUND_IN_FILE"
                End If
            Else
                result = "PROJECT_PARAM"
            End If

            Dim r = dt.NewRow()
            r("RvtName") = If(rvtName, "")
            r("Scope") = "Project"
            r("ParamName") = name
            r("ParamKind") = kind
            r("ProjectGuid") = projGuid
            r("FileGuid") = fileGuid
            r("Result") = result
            r("Notes") = notes
            dt.Rows.Add(r)
        Next

        Return dt
    End Function

    Public Function RunFamilyAudit(doc As Document,
                                   fileMap As Dictionary(Of String, List(Of Guid)),
                                   statusLabel As WinForms.Label,
                                   rvtName As String,
                                   rvtPath As String) As FamilyAuditPack

        Dim pack As New FamilyAuditPack()

        Dim dtSum As New DataTable("FamilySharedParams")
        dtSum.Columns.Add("RvtName", GetType(String))
        dtSum.Columns.Add("Scope", GetType(String))
        dtSum.Columns.Add("FamilyName", GetType(String))
        dtSum.Columns.Add("FamilyCategory", GetType(String))
        dtSum.Columns.Add("ParamName", GetType(String))
        dtSum.Columns.Add("FamilyGuid", GetType(String))
        dtSum.Columns.Add("FileGuid", GetType(String))
        dtSum.Columns.Add("Result", GetType(String))
        dtSum.Columns.Add("Notes", GetType(String))

        Dim dtDet As New DataTable("FamilyParamDetail")
        dtDet.Columns.Add("RvtName", GetType(String))
        dtDet.Columns.Add("RvtPath", GetType(String)) ' UI 필터/트리용. Excel에서는 exporter에서 제외.
        dtDet.Columns.Add("FamilyName", GetType(String))
        dtDet.Columns.Add("FamilyCategory", GetType(String))
        dtDet.Columns.Add("ParamName", GetType(String))
        dtDet.Columns.Add("IsShared", GetType(String))
        dtDet.Columns.Add("ParamGroup", GetType(String))
        dtDet.Columns.Add("ParamType", GetType(String))
        dtDet.Columns.Add("IsInstance", GetType(String))
        dtDet.Columns.Add("FamilyGuid", GetType(String))
        dtDet.Columns.Add("FileGuid", GetType(String))
        dtDet.Columns.Add("Result", GetType(String))
        dtDet.Columns.Add("Notes", GetType(String))

        Dim fams = New FilteredElementCollector(doc).
            OfClass(GetType(Family)).
            Cast(Of Family)().
            OrderBy(Function(x) x.Name, StringComparer.OrdinalIgnoreCase).
            ToList()

        Dim total As Integer = fams.Count
        Dim idx As Integer = 0

        For Each fam As Family In fams
            idx += 1

            Dim famName = fam.Name
            Dim famCat = ""
            Try
                If fam.FamilyCategory IsNot Nothing Then famCat = fam.FamilyCategory.Name
            Catch
                famCat = ""
            End Try

            If statusLabel IsNot Nothing Then
                statusLabel.Text = $"[{rvtName}] 패밀리 처리... ({idx:#,0}/{total:#,0}) {famName}"
                WinForms.Application.DoEvents()
            End If

            Try
                If fam.IsInPlace Then
                    AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, "", "N/A", "", "", "", "", "", "SKIP_INPLACE", "In-place family")
                    Continue For
                End If
            Catch
            End Try

            Dim famDoc As Document = Nothing
            Try
                famDoc = doc.EditFamily(fam)
                If famDoc Is Nothing OrElse Not famDoc.IsFamilyDocument Then
                    AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, "", "N/A", "", "", "", "", "", "OPEN_FAIL", "EditFamily 실패")
                    Continue For
                End If

                Dim fm As FamilyManager = famDoc.FamilyManager
                If fm Is Nothing Then
                    AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, "", "N/A", "", "", "", "", "", "OPEN_FAIL", "FamilyManager 없음")
                    Continue For
                End If

                For Each fp As FamilyParameter In fm.Parameters
                    If fp Is Nothing Then Continue For

                    Dim pName As String = ""
                    Try : pName = fp.Definition.Name : Catch : pName = "" : End Try

                    Dim isSharedBool As Boolean = False
                    Try : isSharedBool = fp.IsShared : Catch : isSharedBool = False : End Try

                    Dim paramGroup As String = ""
                    Try : paramGroup = fp.Definition.ParameterGroup.ToString() : Catch : paramGroup = "" : End Try

                    Dim paramType As String = ""
                    Try : paramType = fp.Definition.ParameterType.ToString() : Catch : paramType = "" : End Try

                    Dim isInst As String = ""
                    Try : isInst = If(fp.IsInstance, "Y", "N") : Catch : isInst = "" : End Try

                    Dim famGuid As String = ""
                    Dim fileGuid As String = ""
                    Dim res As String = ""
                    Dim notes As String = ""

                    If isSharedBool Then
                        Dim gFam As Guid = Guid.Empty
                        If TryGetFamilyParameterGuid(fp, gFam) Then
                            famGuid = gFam.ToString()

                            Dim fileGuids As List(Of Guid) = Nothing
                            If fileMap.TryGetValue(pName, fileGuids) Then
                                fileGuid = String.Join("; ", fileGuids.Select(Function(x) x.ToString()).Distinct().ToArray())
                                If fileGuids.Count > 1 Then notes = "Shared parameter file에 동일 이름 GUID 여러 개"

                                If fileGuids.Any(Function(x) x = gFam) Then
                                    res = If(fileGuids.Count > 1, "OK(MULTI_IN_FILE)", "OK")
                                Else
                                    res = "MISMATCH"
                                End If
                            Else
                                res = "NOT_FOUND_IN_FILE"
                            End If
                        Else
                            res = "GUID_FAIL"
                            notes = "FamilyParameter GUID 추출 실패"
                        End If

                        Dim rs = dtSum.NewRow()
                        rs("RvtName") = If(rvtName, "")
                        rs("Scope") = "Family"
                        rs("FamilyName") = famName
                        rs("FamilyCategory") = famCat
                        rs("ParamName") = pName
                        rs("FamilyGuid") = famGuid
                        rs("FileGuid") = fileGuid
                        rs("Result") = res
                        rs("Notes") = notes
                        dtSum.Rows.Add(rs)
                    Else
                        res = "NON_SHARED"
                    End If

                    AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, pName,
                                 If(isSharedBool, "Y", "N"),
                                 paramGroup, paramType, isInst,
                                 famGuid, fileGuid, res, notes)
                Next

            Catch ex As Exception
                AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, "", "N/A", "", "", "", "", "", "OPEN_FAIL", ex.Message)

            Finally
                If famDoc IsNot Nothing Then
                    Try
                        famDoc.Close(False)
                    Catch
                    End Try
                End If
            End Try
        Next

        pack.Summary = dtSum
        pack.Detail = dtDet
        Return pack
    End Function

    Private Sub AddDetailRow(dt As DataTable,
                             rvtName As String,
                             rvtPath As String,
                             famName As String,
                             famCat As String,
                             pName As String,
                             isShared As String,
                             pGroup As String,
                             pType As String,
                             isInst As String,
                             famGuid As String,
                             fileGuid As String,
                             res As String,
                             notes As String)
        Dim r = dt.NewRow()
        r("RvtName") = If(rvtName, "")
        r("RvtPath") = If(rvtPath, "")
        r("FamilyName") = If(famName, "")
        r("FamilyCategory") = If(famCat, "")
        r("ParamName") = If(pName, "")
        r("IsShared") = If(isShared, "")
        r("ParamGroup") = If(pGroup, "")
        r("ParamType") = If(pType, "")
        r("IsInstance") = If(isInst, "")
        r("FamilyGuid") = If(famGuid, "")
        r("FileGuid") = If(fileGuid, "")
        r("Result") = If(res, "")
        r("Notes") = If(notes, "")
        dt.Rows.Add(r)
    End Sub

    Private Function SafeParamElementName(pe As ParameterElement) As String
        Try
            Return pe.Name
        Catch
            Return ""
        End Try
    End Function

    Private Function TryGetFamilyParameterGuid(fp As FamilyParameter, ByRef g As Guid) As Boolean
        g = Guid.Empty
        If fp Is Nothing Then Return False

        Dim t = fp.GetType()
        Dim p = t.GetProperty("GUID", BindingFlags.Public Or BindingFlags.Instance)
        If p Is Nothing Then Return False

        Dim v = p.GetValue(fp, Nothing)
        If v Is Nothing Then Return False

        If TypeOf v Is Guid Then
            g = DirectCast(v, Guid)
            Return g <> Guid.Empty
        End If

        Return False
    End Function

End Module

'=========================================================
' Excel Export (NPOI fast) + Excel COM AutoFit (perfect)
'  - RvtPath 컬럼은 Excel에서 제외
'  - Excel COM 호출은 Reflection으로 (Option Strict On OK)
'=========================================================
Friend Module ExcelExporter

    Private Const ENABLE_EXCEL_COM_AUTOFIT As Boolean = True
    Private Const EXCEL_AUTOFIT_MAX_COL_WIDTH_CHARS As Integer = 90

    ' Excel Calculation 상수
    Private Const XlCalculationManual As Integer = -4135
    Private Const XlCalculationAutomatic As Integer = -4105

    Public Sub SaveDataTableAsXlsx(dt As DataTable, sheetName As String, filePath As String, ByRef autofitOk As Boolean, ByRef autofitError As String)
        autofitOk = False
        autofitError = ""

        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))
        If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentException("filePath")

        Dim skipCols As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "RvtPath"
        }

        Dim exportColIdx As New List(Of Integer)()
        Dim exportColNames As New List(Of String)()

        For i As Integer = 0 To dt.Columns.Count - 1
            Dim cn = dt.Columns(i).ColumnName
            If skipCols.Contains(cn) Then Continue For
            exportColIdx.Add(i)
            exportColNames.Add(cn)
        Next

        Dim wb As IWorkbook = New XSSFWorkbook()
        Dim sh As ISheet = wb.CreateSheet(MakeSafeSheetName(sheetName))

        Dim header As IRow = sh.CreateRow(0)
        For c As Integer = 0 To exportColNames.Count - 1
            header.CreateCell(c).SetCellValue(exportColNames(c))
        Next

        For r As Integer = 0 To dt.Rows.Count - 1
            Dim row As IRow = sh.CreateRow(r + 1)
            Dim dr As DataRow = dt.Rows(r)

            For c As Integer = 0 To exportColIdx.Count - 1
                Dim srcIdx As Integer = exportColIdx(c)
                Dim v = dr(srcIdx)
                Dim s As String = If(v Is Nothing OrElse v Is DBNull.Value, "", Convert.ToString(v))
                row.CreateCell(c).SetCellValue(s)
            Next
        Next

        Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
            wb.Write(fs)
        End Using

        If ENABLE_EXCEL_COM_AUTOFIT Then
            Dim ok As Boolean = False
            Dim err As String = ""
            ok = TryExcelComAutofit(filePath, EXCEL_AUTOFIT_MAX_COL_WIDTH_CHARS, err)
            autofitOk = ok
            autofitError = If(err, "")
        Else
            autofitOk = False
            autofitError = "AutoFit 옵션이 비활성화되어 있습니다."
        End If
    End Sub

    '---------------------------
    ' Excel COM AutoFit (완벽)
    '---------------------------
    Private Function TryExcelComAutofit(filePath As String, maxWidthChars As Integer, ByRef err As String) As Boolean
        err = ""
        Dim xl As Object = Nothing
        Dim workbooks As Object = Nothing
        Dim wb As Object = Nothing
        Dim sheets As Object = Nothing

        Dim anySheetDone As Boolean = False
        Dim prevCalc As Object = Nothing

        Try
            Dim excelType As Type = Type.GetTypeFromProgID("Excel.Application")
            If excelType Is Nothing Then
                err = "Excel COM을 찾을 수 없습니다(Excel 미설치 또는 ProgID 등록 안됨)."
                Return False
            End If

            xl = Activator.CreateInstance(excelType)
            If xl Is Nothing Then
                err = "Excel.Application 인스턴스 생성 실패."
                Return False
            End If

            ' 성능 옵션
            ComSet(xl, "Visible", False)
            ComSet(xl, "DisplayAlerts", False)
            ComSet(xl, "ScreenUpdating", False)
            ComSet(xl, "EnableEvents", False)

            ' 이전 Calculation 저장 후 Manual로
            prevCalc = ComGet(xl, "Calculation")
            ComSet(xl, "Calculation", XlCalculationManual)

            workbooks = ComGet(xl, "Workbooks")
            wb = ComCall(workbooks, "Open", filePath)

            sheets = ComGet(wb, "Worksheets")
            Dim sheetCount As Integer = Convert.ToInt32(ComGet(sheets, "Count"))

            For i As Integer = 1 To sheetCount
                Dim ws As Object = Nothing
                Dim used As Object = Nothing
                Dim entireCol As Object = Nothing
                Dim cols As Object = Nothing

                Try
                    ws = ComGet(sheets, "Item", i) ' ✅ Item은 프로퍼티(GetProperty)
                    ComCall(ws, "Activate")         ' ✅ UsedRange/AutoFit 안정성

                    used = ComGet(ws, "UsedRange")
                    entireCol = ComGet(used, "EntireColumn")
                    ComCall(entireCol, "AutoFit")

                    cols = ComGet(used, "Columns")
                    Dim colCount As Integer = Convert.ToInt32(ComGet(cols, "Count"))

                    For c As Integer = 1 To colCount
                        Dim colObj As Object = Nothing
                        Try
                            colObj = ComGet(cols, "Item", c)
                            Dim w As Double = 0
                            Try
                                w = Convert.ToDouble(ComGet(colObj, "ColumnWidth"))
                            Catch
                                w = 0
                            End Try
                            If w > maxWidthChars Then
                                ComSet(colObj, "ColumnWidth", CDbl(maxWidthChars))
                            End If
                        Finally
                            ReleaseCom(colObj)
                        End Try
                    Next

                    anySheetDone = True

                Catch exSheet As Exception
                    ' 시트 하나 실패해도 나머지 진행
                    err = "AutoFit 중 오류: " & exSheet.Message
                Finally
                    ReleaseCom(cols)
                    ReleaseCom(entireCol)
                    ReleaseCom(used)
                    ReleaseCom(ws)
                End Try
            Next

            ' Calculation 복원
            Try
                If prevCalc IsNot Nothing Then
                    ComSet(xl, "Calculation", prevCalc)
                Else
                    ComSet(xl, "Calculation", XlCalculationAutomatic)
                End If
            Catch
            End Try

            ComCall(wb, "Save")
            ComCall(wb, "Close", False)

            Return anySheetDone

        Catch ex As Exception
            err = ex.Message
            Return False

        Finally
            Try
                ReleaseCom(sheets)
                ReleaseCom(wb)
                ReleaseCom(workbooks)

                If xl IsNot Nothing Then
                    Try : ComCall(xl, "Quit") : Catch : End Try
                    ReleaseCom(xl)
                End If
            Catch
            End Try

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    '---------------------------
    ' COM Reflection Helper (✅ Public/Instance 포함)
    '---------------------------
    Private Const BF As BindingFlags = BindingFlags.Public Or BindingFlags.Instance

    Private Function ComGet(obj As Object, propName As String, ParamArray args As Object()) As Object
        If obj Is Nothing Then Return Nothing
        Return obj.GetType().InvokeMember(propName, BF Or BindingFlags.GetProperty, Nothing, obj, args)
    End Function

    Private Sub ComSet(obj As Object, propName As String, value As Object)
        If obj Is Nothing Then Exit Sub
        obj.GetType().InvokeMember(propName, BF Or BindingFlags.SetProperty, Nothing, obj, New Object() {value})
    End Sub

    Private Function ComCall(obj As Object, methodName As String, ParamArray args As Object()) As Object
        If obj Is Nothing Then Return Nothing
        Return obj.GetType().InvokeMember(methodName, BF Or BindingFlags.InvokeMethod, Nothing, obj, args)
    End Function

    Private Sub ReleaseCom(o As Object)
        Try
            If o IsNot Nothing AndAlso System.Runtime.InteropServices.Marshal.IsComObject(o) Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o)
            End If
        Catch
        End Try
    End Sub

    Private Function MakeSafeSheetName(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "Sheet1"
        Dim invalid = New Char() {":"c, "\"c, "/"c, "?"c, "*"c, "["c, "]"c}
        Dim s = name
        For Each ch In invalid
            s = s.Replace(ch, "_"c)
        Next
        If s.Length > 31 Then s = s.Substring(0, 31)
        Return s
    End Function

End Module

Friend NotInheritable Class RevitWin32Window
    Implements WinForms.IWin32Window

    Private ReadOnly _hwnd As IntPtr

    Private Sub New(hwnd As IntPtr)
        _hwnd = hwnd
    End Sub

    Public ReadOnly Property Handle As IntPtr Implements WinForms.IWin32Window.Handle
        Get
            Return _hwnd
        End Get
    End Property

    Public Shared Function FromRevit(uiApp As UIApplication) As WinForms.IWin32Window
        Dim hwnd As IntPtr = IntPtr.Zero
        Try
            hwnd = uiApp.MainWindowHandle
        Catch
            hwnd = IntPtr.Zero
        End Try
        Return New RevitWin32Window(hwnd)
    End Function
End Class
