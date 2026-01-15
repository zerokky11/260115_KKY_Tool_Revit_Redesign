Option Strict On
Option Explicit On

Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Namespace KKY_Tool_Revit

    <Transaction(TransactionMode.Manual)>
    Public Class Cmd_NestedFamilyParamAudit
        Implements IExternalCommand

        Public Function Execute(commandData As ExternalCommandData,
                                ByRef message As String,
                                elements As ElementSet) As Result Implements IExternalCommand.Execute
            Try
                Dim uiapp As UIApplication = commandData.Application
                Using f As New AuditForm(uiapp)
                    f.ShowDialog(New RevitWin32Window(uiapp.MainWindowHandle))
                End Using
                Return Result.Succeeded
            Catch ex As Exception
                message = ex.ToString()
                Return Result.Failed
            End Try
        End Function
    End Class

    Friend Class RevitWin32Window
        Implements IWin32Window

        Private ReadOnly _handle As IntPtr
        Public Sub New(handle As IntPtr)
            _handle = handle
        End Sub

        Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
            Get
                Return _handle
            End Get
        End Property
    End Class

    Friend Enum AuditIssueType
        OK
        MissingAssociation
        GuidMismatch
        HostParamNotShared
        ParamNotFound
        [Error]
    End Enum

    Friend Enum FoundScope
        InstanceParam
        TypeParam
    End Enum

    Friend Class SharedParamItem
        Public Property Name As String = ""
        Public Property Guid As Guid
        Public Property GroupName As String = ""
        Public Property DataTypeToken As String = ""

        Public Overrides Function ToString() As String
            Dim g8 As String = ""
            Try
                g8 = Guid.ToString("D")
                If g8.Length >= 8 Then g8 = g8.Substring(0, 8)
            Catch
                g8 = ""
            End Try

            If String.IsNullOrWhiteSpace(GroupName) Then
                Return $"{Name}  [{g8}]"
            End If
            Return $"{Name}  ({GroupName})  [{g8}]"
        End Function
    End Class

    Friend Class FoundParam
        Public Property P As Parameter
        Public Property Scope As FoundScope
    End Class

    Friend Class AuditRow
        Public Property ProjectPath As String = ""

        Public Property HostFamilyName As String = ""
        Public Property HostFamilyCategory As String = ""

        Public Property NestedFamilyName As String = ""
        Public Property NestedTypeName As String = ""
        Public Property NestedCategory As String = ""

        Public Property TargetParamName As String = ""
        Public Property ExpectedGuid As String = ""

        Public Property FoundScope As String = ""
        Public Property NestedParamGuid As String = ""
        Public Property NestedParamDataType As String = ""

        Public Property AssocHostParamName As String = ""
        Public Property HostParamGuid As String = ""
        Public Property HostParamIsShared As String = ""

        Public Property Issue As String = ""
        Public Property Notes As String = ""
    End Class

    Friend Class AuditForm
        Inherits System.Windows.Forms.Form

        Private ReadOnly _uiapp As UIApplication

        ' --- Shared Param source: Revit configured SharedParameters file ---
        Private ReadOnly btnReloadShared As New Button()
        Private ReadOnly lblSharedSource As New System.Windows.Forms.Label()
        Private ReadOnly txtParamSearch As New System.Windows.Forms.TextBox()
        Private ReadOnly btnAddParam As New Button()
        Private ReadOnly btnRemoveParam As New Button()
        Private ReadOnly btnClearSelectedParams As New Button()
        Private ReadOnly lstAvailableParams As New ListBox()
        Private ReadOnly lstSelectedParams As New ListBox()
        Private ReadOnly lblParamCounts As New System.Windows.Forms.Label()

        ' --- RVT list ---
        Private ReadOnly lstFiles As New ListBox()
        Private ReadOnly btnAddRvt As New Button()
        Private ReadOnly btnRemoveRvt As New Button()
        Private ReadOnly btnClearRvt As New Button()

        ' --- Actions ---
        Private ReadOnly btnScan As New Button()
        Private ReadOnly btnExport As New Button()
        Private ReadOnly btnClose As New Button()

        Private ReadOnly dgv As New DataGridView()
        Private ReadOnly lblStatus As New System.Windows.Forms.Label()
        Private ReadOnly pbar As New ProgressBar()

        Private _allSharedParams As List(Of SharedParamItem) = New List(Of SharedParamItem)()
        Private _rows As List(Of AuditRow) = New List(Of AuditRow)()

        Public Sub New(uiapp As UIApplication)
            _uiapp = uiapp
            BuildUi()

            ' ✅ 사용자 입력 없이: 현재 Revit에 설정된 Shared Parameters 파일에서 로드
            ReloadSharedParamsFromRevit()

            RefreshAvailableParams()
            UpdateButtonStates()
        End Sub

        Private Sub BuildUi()
            Text = "Association Check - Nested Family Parameter Audit"
            Width = 1450
            Height = 820
            StartPosition = FormStartPosition.CenterScreen

            Dim root As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .RowCount = 2,
                .ColumnCount = 1
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 70))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            ' ---------- Top bar ----------
            Dim top As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .RowCount = 1,
                .ColumnCount = 6
            }
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120)) ' Scan
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120)) ' Export
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120)) ' Close
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))  ' Status
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 240)) ' Progress
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 10))

            btnScan.Text = "스캔 실행"
            btnExport.Text = "CSV 저장..."
            btnClose.Text = "닫기"
            lblStatus.Text = "준비됨"
            lblStatus.AutoSize = True

            AddHandler btnScan.Click, AddressOf OnScan
            AddHandler btnExport.Click, AddressOf OnExport
            AddHandler btnClose.Click, Sub() Close()

            pbar.Style = ProgressBarStyle.Blocks
            pbar.Minimum = 0
            pbar.Maximum = 100
            pbar.Value = 0
            pbar.Dock = DockStyle.Fill

            top.Controls.Add(btnScan, 0, 0)
            top.Controls.Add(btnExport, 1, 0)
            top.Controls.Add(btnClose, 2, 0)
            top.Controls.Add(lblStatus, 3, 0)
            top.Controls.Add(pbar, 4, 0)

            ' ---------- Body ----------
            Dim body As New SplitContainer() With {
                .Dock = DockStyle.Fill,
                .Orientation = Orientation.Vertical,
                .SplitterDistance = 520
            }

            Dim leftRoot As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .RowCount = 2,
                .ColumnCount = 1
            }
            leftRoot.RowStyles.Add(New RowStyle(SizeType.Percent, 58))
            leftRoot.RowStyles.Add(New RowStyle(SizeType.Percent, 42))

            leftRoot.Controls.Add(BuildParamGroup(), 0, 0)
            leftRoot.Controls.Add(BuildRvtGroup(), 0, 1)

            body.Panel1.Controls.Add(leftRoot)

            Dim right As New GroupBox() With {.Text = "결과", .Dock = DockStyle.Fill}
            dgv.Dock = DockStyle.Fill
            dgv.ReadOnly = True
            dgv.AllowUserToAddRows = False
            dgv.AllowUserToDeleteRows = False
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.MultiSelect = True
            right.Controls.Add(dgv)

            body.Panel2.Controls.Add(right)

            root.Controls.Add(top, 0, 0)
            root.Controls.Add(body, 0, 1)

            Controls.Add(root)

            btnExport.Enabled = False
        End Sub

        Private Function BuildParamGroup() As System.Windows.Forms.Control
            Dim gb As New GroupBox() With {.Text = "1) 검토할 파라미터 (Revit에 설정된 Shared Parameters 파일에서 검색 후 등록)", .Dock = DockStyle.Fill}

            Dim lay As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .RowCount = 4,
                .ColumnCount = 1
            }
            lay.RowStyles.Add(New RowStyle(SizeType.Absolute, 36)) ' source row
            lay.RowStyles.Add(New RowStyle(SizeType.Absolute, 36)) ' search row
            lay.RowStyles.Add(New RowStyle(SizeType.Percent, 100)) ' lists
            lay.RowStyles.Add(New RowStyle(SizeType.Absolute, 28)) ' counts

            ' source row
            Dim srcRow As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .RowCount = 1, .ColumnCount = 2}
            srcRow.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            srcRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            btnReloadShared.Text = "SharedParam 새로고침"
            AddHandler btnReloadShared.Click, Sub()
                                                  ReloadSharedParamsFromRevit()
                                                  RefreshAvailableParams()
                                                  UpdateButtonStates()
                                              End Sub

            lblSharedSource.Text = "(Revit Shared Parameters 파일 미확인)"
            lblSharedSource.AutoEllipsis = True
            lblSharedSource.Dock = DockStyle.Fill
            lblSharedSource.TextAlign = Drawing.ContentAlignment.MiddleLeft

            srcRow.Controls.Add(btnReloadShared, 0, 0)
            srcRow.Controls.Add(lblSharedSource, 1, 0)

            ' search row
            Dim searchRow As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .RowCount = 1, .ColumnCount = 2}
            searchRow.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 80))
            searchRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            Dim lblSearch As New System.Windows.Forms.Label() With {.Text = "검색:", .AutoSize = True, .Padding = New Padding(0, 8, 0, 0)}
            txtParamSearch.Dock = DockStyle.Fill
            AddHandler txtParamSearch.TextChanged, Sub() RefreshAvailableParams()

            searchRow.Controls.Add(lblSearch, 0, 0)
            searchRow.Controls.Add(txtParamSearch, 1, 0)

            ' lists row
            Dim lists As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .RowCount = 1, .ColumnCount = 3}
            lists.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 46))
            lists.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 84))
            lists.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 54))

            Dim gbAvail As New GroupBox() With {.Text = "검색 결과", .Dock = DockStyle.Fill}
            lstAvailableParams.Dock = DockStyle.Fill
            lstAvailableParams.SelectionMode = SelectionMode.MultiExtended
            gbAvail.Controls.Add(lstAvailableParams)

            Dim mid As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown}
            btnAddParam.Text = "추가 >>"
            btnRemoveParam.Text = "<< 제거"
            btnClearSelectedParams.Text = "선택목록 비우기"
            btnAddParam.Width = 76
            btnRemoveParam.Width = 76
            btnClearSelectedParams.Width = 76

            AddHandler btnAddParam.Click, AddressOf OnAddSelectedParam
            AddHandler btnRemoveParam.Click, AddressOf OnRemoveSelectedParam
            AddHandler btnClearSelectedParams.Click, AddressOf OnClearSelectedParams

            mid.Controls.Add(btnAddParam)
            mid.Controls.Add(btnRemoveParam)
            mid.Controls.Add(btnClearSelectedParams)

            Dim gbSel As New GroupBox() With {.Text = "검토 대상(등록됨)", .Dock = DockStyle.Fill}
            lstSelectedParams.Dock = DockStyle.Fill
            lstSelectedParams.SelectionMode = SelectionMode.MultiExtended
            gbSel.Controls.Add(lstSelectedParams)

            lists.Controls.Add(gbAvail, 0, 0)
            lists.Controls.Add(mid, 1, 0)
            lists.Controls.Add(gbSel, 2, 0)

            lblParamCounts.Text = "선택 0개 / 파일 0개"
            lblParamCounts.AutoSize = True
            lblParamCounts.Padding = New Padding(4, 4, 0, 0)

            lay.Controls.Add(srcRow, 0, 0)
            lay.Controls.Add(searchRow, 0, 1)
            lay.Controls.Add(lists, 0, 2)
            lay.Controls.Add(lblParamCounts, 0, 3)

            gb.Controls.Add(lay)
            Return gb
        End Function

        Private Function BuildRvtGroup() As System.Windows.Forms.Control
            Dim gb As New GroupBox() With {.Text = "2) RVT 파일 목록", .Dock = DockStyle.Fill}

            Dim lay As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .RowCount = 2,
                .ColumnCount = 1
            }
            lay.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))
            lay.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            Dim btnRow As New FlowLayoutPanel() With {.Dock = DockStyle.Fill}
            btnAddRvt.Text = "RVT 추가..."
            btnRemoveRvt.Text = "선택 제거"
            btnClearRvt.Text = "목록 지우기"
            AddHandler btnAddRvt.Click, AddressOf OnAddFiles
            AddHandler btnRemoveRvt.Click, AddressOf OnRemoveSelected
            AddHandler btnClearRvt.Click, AddressOf OnClearFiles

            btnRow.Controls.Add(btnAddRvt)
            btnRow.Controls.Add(btnRemoveRvt)
            btnRow.Controls.Add(btnClearRvt)

            lstFiles.Dock = DockStyle.Fill
            lay.Controls.Add(btnRow, 0, 0)
            lay.Controls.Add(lstFiles, 0, 1)

            gb.Controls.Add(lay)
            Return gb
        End Function

        Private Sub UpdateButtonStates()
            Dim hasParams As Boolean = (lstSelectedParams.Items.Count > 0)
            Dim hasRvts As Boolean = (lstFiles.Items.Count > 0)

            btnScan.Enabled = hasParams AndAlso hasRvts
            btnExport.Enabled = (_rows IsNot Nothing AndAlso _rows.Count > 0)

            lblParamCounts.Text = $"선택 {lstSelectedParams.Items.Count}개 / 파일 {lstFiles.Items.Count}개"
        End Sub

        Private Sub SetBusy(isBusy As Boolean, status As String)
            btnReloadShared.Enabled = Not isBusy
            txtParamSearch.Enabled = Not isBusy
            btnAddParam.Enabled = Not isBusy
            btnRemoveParam.Enabled = Not isBusy
            btnClearSelectedParams.Enabled = Not isBusy
            lstAvailableParams.Enabled = Not isBusy
            lstSelectedParams.Enabled = Not isBusy

            btnAddRvt.Enabled = Not isBusy
            btnRemoveRvt.Enabled = Not isBusy
            btnClearRvt.Enabled = Not isBusy
            lstFiles.Enabled = Not isBusy

            btnScan.Enabled = Not isBusy
            btnClose.Enabled = Not isBusy

            lblStatus.Text = status

            If isBusy Then
                pbar.Style = ProgressBarStyle.Marquee
                pbar.MarqueeAnimationSpeed = 30
            Else
                pbar.Style = ProgressBarStyle.Blocks
                pbar.MarqueeAnimationSpeed = 0
                pbar.Value = 0
            End If

            System.Windows.Forms.Application.DoEvents()
            UpdateButtonStates()
        End Sub

        ' =========================
        ' ✅ Shared params: from Revit configured file
        ' =========================
        Private Sub ReloadSharedParamsFromRevit()
            _allSharedParams = New List(Of SharedParamItem)()

            Dim spPath As String = ""
            Try
                spPath = SafeStr(_uiapp.Application.SharedParametersFilename)
            Catch
                spPath = ""
            End Try

            Dim defFile As DefinitionFile = Nothing
            Try
                defFile = _uiapp.Application.OpenSharedParameterFile()
            Catch
                defFile = Nothing
            End Try

            If defFile Is Nothing Then
                lblSharedSource.Text = If(String.IsNullOrWhiteSpace(spPath),
                                         "Revit에 Shared Parameters 파일이 설정되지 않았습니다. (Revit 옵션에서 설정 필요)",
                                         "Shared Parameters 파일을 열 수 없습니다: " & spPath)
                Return
            End If

            Dim list As New List(Of SharedParamItem)()

            Try
                For Each grp As DefinitionGroup In defFile.Groups
                    If grp Is Nothing Then Continue For

                    For Each defn As Definition In grp.Definitions
                        Dim ext As ExternalDefinition = TryCast(defn, ExternalDefinition)
                        If ext Is Nothing Then Continue For

                        list.Add(New SharedParamItem With {
                            .Name = SafeStr(ext.Name),
                            .Guid = ext.GUID,
                            .GroupName = SafeStr(grp.Name),
                            .DataTypeToken = SafeDefTypeToken(ext)
                        })
                    Next
                Next
            Catch ex As Exception
                lblSharedSource.Text = "Shared Parameters 로딩 중 오류: " & ex.Message
                _allSharedParams = New List(Of SharedParamItem)()
                Return
            End Try

            _allSharedParams = list.
                Where(Function(x) x IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(x.Name)).
                OrderBy(Function(x) If(x.GroupName, ""), StringComparer.OrdinalIgnoreCase).
                ThenBy(Function(x) If(x.Name, ""), StringComparer.OrdinalIgnoreCase).
                ToList()

            Dim fileName As String = ""
            Try
                fileName = If(String.IsNullOrWhiteSpace(spPath), "(unknown)", spPath)
            Catch
                fileName = "(unknown)"
            End Try

            lblSharedSource.Text = $"SharedParam Source: {fileName}  /  { _allSharedParams.Count }개"
        End Sub

        Private Sub RefreshAvailableParams()
            lstAvailableParams.BeginUpdate()
            Try
                lstAvailableParams.Items.Clear()

                Dim q As String = If(txtParamSearch.Text, "").Trim()
                Dim qHas As Boolean = Not String.IsNullOrWhiteSpace(q)

                Dim selectedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each it As Object In lstSelectedParams.Items
                    Dim sp As SharedParamItem = TryCast(it, SharedParamItem)
                    If sp IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(sp.Name) Then
                        selectedNames.Add(sp.Name)
                    End If
                Next

                For Each sp As SharedParamItem In _allSharedParams
                    If sp Is Nothing Then Continue For
                    If selectedNames.Contains(sp.Name) Then Continue For

                    If qHas Then
                        Dim hit As Boolean =
                            (sp.Name IsNot Nothing AndAlso sp.Name.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0) OrElse
                            (sp.GroupName IsNot Nothing AndAlso sp.GroupName.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0) OrElse
                            (sp.Guid.ToString("D").IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0)
                        If Not hit Then Continue For
                    End If

                    lstAvailableParams.Items.Add(sp)
                Next
            Finally
                lstAvailableParams.EndUpdate()
            End Try

            UpdateButtonStates()
        End Sub

        Private Sub OnAddSelectedParam(sender As Object, e As EventArgs)
            If lstAvailableParams.SelectedItems.Count = 0 Then Return

            Dim existingNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each it As Object In lstSelectedParams.Items
                Dim sp0 As SharedParamItem = TryCast(it, SharedParamItem)
                If sp0 IsNot Nothing Then existingNames.Add(sp0.Name)
            Next

            For Each it As Object In lstAvailableParams.SelectedItems
                Dim sp As SharedParamItem = TryCast(it, SharedParamItem)
                If sp Is Nothing Then Continue For
                If existingNames.Contains(sp.Name) Then Continue For
                lstSelectedParams.Items.Add(sp)
                existingNames.Add(sp.Name)
            Next

            RefreshAvailableParams()
            UpdateButtonStates()
        End Sub

        Private Sub OnRemoveSelectedParam(sender As Object, e As EventArgs)
            If lstSelectedParams.SelectedItems.Count = 0 Then Return

            Dim toRemove As New List(Of Object)()
            For Each it As Object In lstSelectedParams.SelectedItems
                toRemove.Add(it)
            Next
            For Each it As Object In toRemove
                lstSelectedParams.Items.Remove(it)
            Next

            RefreshAvailableParams()
            UpdateButtonStates()
        End Sub

        Private Sub OnClearSelectedParams(sender As Object, e As EventArgs)
            lstSelectedParams.Items.Clear()
            RefreshAvailableParams()
            UpdateButtonStates()
        End Sub

        ' =========================
        ' RVT list
        ' =========================
        Private Sub OnAddFiles(sender As Object, e As EventArgs)
            Using ofd As New OpenFileDialog()
                ofd.Filter = "Revit Project (*.rvt)|*.rvt"
                ofd.Multiselect = True
                ofd.Title = "스캔할 RVT 파일 선택"
                If ofd.ShowDialog(Me) <> DialogResult.OK Then Return

                Dim existing As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each it As Object In lstFiles.Items
                    existing.Add(CStr(it))
                Next

                For Each p As String In ofd.FileNames
                    If File.Exists(p) AndAlso Not existing.Contains(p) Then
                        lstFiles.Items.Add(p)
                        existing.Add(p)
                    End If
                Next
            End Using

            UpdateButtonStates()
        End Sub

        Private Sub OnRemoveSelected(sender As Object, e As EventArgs)
            Dim selected As New List(Of Object)()
            For Each it As Object In lstFiles.SelectedItems
                selected.Add(it)
            Next
            For Each it As Object In selected
                lstFiles.Items.Remove(it)
            Next
            UpdateButtonStates()
        End Sub

        Private Sub OnClearFiles(sender As Object, e As EventArgs)
            lstFiles.Items.Clear()
            UpdateButtonStates()
        End Sub

        ' =========================
        ' Scan / Export
        ' =========================
        Private Sub OnScan(sender As Object, e As EventArgs)
            If lstSelectedParams.Items.Count = 0 Then
                MessageBox.Show(Me, "검토할 파라미터를 먼저 등록하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            If lstFiles.Items.Count = 0 Then
                MessageBox.Show(Me, "RVT 파일을 먼저 추가하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selectedParams As New List(Of SharedParamItem)()
            For Each it As Object In lstSelectedParams.Items
                Dim sp As SharedParamItem = TryCast(it, SharedParamItem)
                If sp IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(sp.Name) Then
                    selectedParams.Add(sp)
                End If
            Next
            If selectedParams.Count = 0 Then
                MessageBox.Show(Me, "검토할 파라미터가 비어있습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim paths As New List(Of String)()
            For Each it As Object In lstFiles.Items
                paths.Add(CStr(it))
            Next

            SetBusy(True, "스캔 중... (복합 패밀리만 대상으로 연동 검토)")

            Dim results As List(Of AuditRow) = Nothing
            Try
                results = AuditFiles(_uiapp, paths, selectedParams)
            Catch ex As Exception
                SetBusy(False, "오류")
                MessageBox.Show(Me, ex.ToString(), "스캔 오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            _rows = results

            dgv.DataSource = Nothing
            dgv.DataSource = New BindingList(Of AuditRow)(_rows)

            btnExport.Enabled = (_rows IsNot Nothing AndAlso _rows.Count > 0)

            Dim cntOk As Integer = _rows.Where(Function(r) r.Issue = AuditIssueType.OK.ToString()).Count()
            Dim cntMiss As Integer = _rows.Where(Function(r) r.Issue = AuditIssueType.MissingAssociation.ToString()).Count()
            Dim cntGuid As Integer = _rows.Where(Function(r) r.Issue = AuditIssueType.GuidMismatch.ToString()).Count()
            Dim cntHostNotShared As Integer = _rows.Where(Function(r) r.Issue = AuditIssueType.HostParamNotShared.ToString()).Count()
            Dim cntNotFound As Integer = _rows.Where(Function(r) r.Issue = AuditIssueType.ParamNotFound.ToString()).Count()
            Dim cntErr As Integer = _rows.Where(Function(r) r.Issue = AuditIssueType.[Error].ToString()).Count()

            SetBusy(False, $"완료: RVT {paths.Count}개 / Rows={_rows.Count} (OK {cntOk}, Missing {cntMiss}, Guid {cntGuid}, HostNotShared {cntHostNotShared}, NotFound {cntNotFound}, Err {cntErr})")
        End Sub

        Private Sub OnExport(sender As Object, e As EventArgs)
            If _rows Is Nothing OrElse _rows.Count = 0 Then
                MessageBox.Show(Me, "저장할 결과가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Using sfd As New SaveFileDialog()
                sfd.Filter = "CSV (*.csv)|*.csv"
                sfd.Title = "결과 CSV 저장"
                sfd.FileName = "AssociationCheck_Result.csv"
                If sfd.ShowDialog(Me) <> DialogResult.OK Then Return

                Try
                    WriteCsv(sfd.FileName, _rows)
                    MessageBox.Show(Me, "저장 완료: " & sfd.FileName, "완료", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show(Me, ex.ToString(), "저장 오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Sub

        ' -------------------------
        ' Core Audit (Complex family only)
        ' -------------------------
        Private Shared Function AuditFiles(uiapp As UIApplication,
                                           paths As List(Of String),
                                           selectedParams As List(Of SharedParamItem)) As List(Of AuditRow)

            Dim rows As New List(Of AuditRow)()

            Dim expectedByName As New Dictionary(Of String, SharedParamItem)(StringComparer.OrdinalIgnoreCase)
            For Each sp As SharedParamItem In selectedParams
                If sp Is Nothing OrElse String.IsNullOrWhiteSpace(sp.Name) Then Continue For
                If Not expectedByName.ContainsKey(sp.Name) Then expectedByName.Add(sp.Name, sp)
            Next

            For Each path As String In paths
                Dim doc As Document = Nothing
                Try
                    doc = OpenProjectDocument(uiapp.Application, path)

                    Dim fams As IEnumerable(Of Family) =
                        New FilteredElementCollector(doc).
                            OfClass(GetType(Family)).
                            Cast(Of Family)()

                    For Each fam As Family In fams
                        If fam Is Nothing Then Continue For
                        If Not fam.IsEditable Then Continue For
                        If fam.IsInPlace Then Continue For

                        AuditFamilyAsHost(doc, fam, path, expectedByName, rows)
                    Next

                Catch ex As Exception
                    rows.Add(New AuditRow With {
                        .ProjectPath = path,
                        .Issue = AuditIssueType.[Error].ToString(),
                        .Notes = "Project open/scan error: " & ex.Message
                    })
                Finally
                    If doc IsNot Nothing Then
                        Try
                            doc.Close(False)
                        Catch
                        End Try
                    End If
                End Try
            Next

            Return rows
        End Function

        Private Shared Sub AuditFamilyAsHost(hostDoc As Document,
                                             hostFamily As Family,
                                             projectPath As String,
                                             expectedByName As Dictionary(Of String, SharedParamItem),
                                             rows As List(Of AuditRow))

            Dim famDoc As Document = Nothing
            Try
                famDoc = hostDoc.EditFamily(hostFamily)
                If famDoc Is Nothing OrElse Not famDoc.IsFamilyDocument Then Return

                Dim nestedInstances As List(Of FamilyInstance) =
                    New FilteredElementCollector(famDoc).
                        OfClass(GetType(FamilyInstance)).
                        WhereElementIsNotElementType().
                        Cast(Of FamilyInstance)().
                        Where(Function(x) x IsNot Nothing AndAlso x.Symbol IsNot Nothing AndAlso x.Symbol.Family IsNot Nothing).
                        ToList()

                If nestedInstances.Count = 0 Then Return

                ' ✅ 같은 Symbol 반복검사 방지
                Dim repInstances As List(Of FamilyInstance) =
                    nestedInstances.
                        GroupBy(Function(fi) fi.Symbol.Id.IntegerValue).
                        Select(Function(g) g.First()).
                        ToList()

                Dim hostCat As String = ""
                Try
                    If hostFamily.FamilyCategory IsNot Nothing Then hostCat = hostFamily.FamilyCategory.Name
                Catch
                End Try

                For Each fi As FamilyInstance In repInstances
                    Dim nestedFam As Family = fi.Symbol.Family

                    Dim nestedCat As String = ""
                    Try
                        If fi.Category IsNot Nothing Then nestedCat = fi.Category.Name
                    Catch
                    End Try

                    Dim map As Dictionary(Of String, List(Of FoundParam)) = CollectParamMap(fi)

                    For Each kv As KeyValuePair(Of String, SharedParamItem) In expectedByName
                        Dim targetName As String = kv.Key
                        Dim expected As SharedParamItem = kv.Value

                        Dim found As List(Of FoundParam) = Nothing
                        If Not map.TryGetValue(targetName, found) OrElse found Is Nothing OrElse found.Count = 0 Then
                            rows.Add(New AuditRow With {
                                .ProjectPath = projectPath,
                                .HostFamilyName = hostFamily.Name,
                                .HostFamilyCategory = hostCat,
                                .NestedFamilyName = nestedFam.Name,
                                .NestedTypeName = SafeStr(fi.Symbol.Name),
                                .NestedCategory = nestedCat,
                                .TargetParamName = targetName,
                                .ExpectedGuid = expected.Guid.ToString("D"),
                                .Issue = AuditIssueType.ParamNotFound.ToString(),
                                .Notes = "네스티드(하위) 패밀리 인스턴스/타입에서 해당 이름의 파라미터를 찾지 못함"
                            })
                            Continue For
                        End If

                        For Each fp As FoundParam In found
                            Dim p As Parameter = fp.P
                            If p Is Nothing OrElse p.Definition Is Nothing Then Continue For

                            ' ✅ 여기부터 핵심 변경: ExternalDefinition 캐스팅이 아니라 Parameter 자체에서 GUID/Shared 판별
                            Dim nestedGuid As Guid
                            Dim nestedGuidOk As Boolean = TryGetParameterGuid(p, nestedGuid)
                            Dim nestedGuidStr As String = If(nestedGuidOk, nestedGuid.ToString("D"), "")

                            Dim nestedIsShared As Boolean = False
                            Dim nestedIsSharedKnown As Boolean = TryGetParameterIsShared(p, nestedIsShared)

                            Dim assoc As FamilyParameter = Nothing
                            Try
                                assoc = famDoc.FamilyManager.GetAssociatedFamilyParameter(p)
                            Catch ex As Exception
                                rows.Add(New AuditRow With {
                                    .ProjectPath = projectPath,
                                    .HostFamilyName = hostFamily.Name,
                                    .HostFamilyCategory = hostCat,
                                    .NestedFamilyName = nestedFam.Name,
                                    .NestedTypeName = SafeStr(fi.Symbol.Name),
                                    .NestedCategory = nestedCat,
                                    .TargetParamName = targetName,
                                    .ExpectedGuid = expected.Guid.ToString("D"),
                                    .FoundScope = fp.Scope.ToString(),
                                    .NestedParamGuid = nestedGuidStr,
                                    .NestedParamDataType = SafeDefTypeToken(p.Definition),
                                    .Issue = AuditIssueType.[Error].ToString(),
                                    .Notes = "GetAssociatedFamilyParameter 실패: " & ex.Message
                                })
                                Continue For
                            End Try

                            Dim issue As AuditIssueType = AuditIssueType.OK
                            Dim notes As String = ""

                            If assoc Is Nothing Then
                                issue = AuditIssueType.MissingAssociation
                                notes = "호스트 패밀리 파라미터로 연동(Associate)되지 않음"
                            Else
                                ' 1) 네스티드 GUID 체크(Shared 파라미터면 GUID가 나와야 정상)
                                If nestedGuidOk Then
                                    If nestedGuid <> expected.Guid Then
                                        issue = AuditIssueType.GuidMismatch
                                        notes = $"네스티드 파라미터 GUID 불일치 (Expected {expected.Guid:D}, Nested {nestedGuid:D})"
                                    End If
                                Else
                                    ' ✅ 여기서 이제 “정말 Shared가 아닌지 / 확인이 안되는지”를 구분해서 메시지
                                    If nestedIsSharedKnown Then
                                        If nestedIsShared Then
                                            notes = "네스티드 파라미터 IsShared=True 이지만 GUID 추출 실패(특이 케이스)"
                                        Else
                                            notes = "네스티드 파라미터 IsShared=False (Shared 아님, 이름만 일치)"
                                        End If
                                    Else
                                        notes = "네스티드 파라미터 Shared 여부 확인 실패(이름만 일치)"
                                    End If
                                End If

                                ' 2) 호스트 연결 파라미터가 Shared가 아니면 표시
                                If assoc.IsShared = False Then
                                    If issue = AuditIssueType.OK Then issue = AuditIssueType.HostParamNotShared
                                    If notes <> "" Then notes &= " / "
                                    notes &= "연결된 호스트 FamilyParameter가 Shared가 아님"
                                End If

                                ' 3) 호스트 GUID도 비교 가능하면 비교
                                Dim hostGuid As Guid
                                If TryGetDefinitionGuid(assoc.Definition, hostGuid) Then
                                    If hostGuid <> expected.Guid Then
                                        If issue = AuditIssueType.OK Then issue = AuditIssueType.GuidMismatch
                                        If notes <> "" Then notes &= " / "
                                        notes &= $"호스트 파라미터 GUID 불일치 (Expected {expected.Guid:D}, Host {hostGuid:D})"
                                    End If
                                End If
                            End If

                            Dim row As New AuditRow With {
                                .ProjectPath = projectPath,
                                .HostFamilyName = hostFamily.Name,
                                .HostFamilyCategory = hostCat,
                                .NestedFamilyName = nestedFam.Name,
                                .NestedTypeName = SafeStr(fi.Symbol.Name),
                                .NestedCategory = nestedCat,
                                .TargetParamName = targetName,
                                .ExpectedGuid = expected.Guid.ToString("D"),
                                .FoundScope = fp.Scope.ToString(),
                                .NestedParamGuid = nestedGuidStr,
                                .NestedParamDataType = SafeDefTypeToken(p.Definition),
                                .Issue = issue.ToString(),
                                .Notes = notes
                            }

                            If assoc IsNot Nothing Then
                                row.AssocHostParamName = SafeStr(assoc.Definition.Name)
                                row.HostParamIsShared = assoc.IsShared.ToString()

                                Dim hostGuid2 As Guid
                                If TryGetDefinitionGuid(assoc.Definition, hostGuid2) Then
                                    row.HostParamGuid = hostGuid2.ToString("D")
                                End If
                            End If

                            rows.Add(row)
                        Next
                    Next
                Next

            Finally
                If famDoc IsNot Nothing Then
                    Try
                        famDoc.Close(False)
                    Catch
                    End Try
                End If
            End Try
        End Sub

        Private Shared Function CollectParamMap(fi As FamilyInstance) As Dictionary(Of String, List(Of FoundParam))
            Dim map As New Dictionary(Of String, List(Of FoundParam))(StringComparer.OrdinalIgnoreCase)

            ' Instance params
            Try
                For Each p As Parameter In fi.Parameters
                    If p Is Nothing OrElse p.Definition Is Nothing Then Continue For
                    Dim name As String = p.Definition.Name
                    If String.IsNullOrWhiteSpace(name) Then Continue For
                    If Not map.ContainsKey(name) Then map(name) = New List(Of FoundParam)()
                    map(name).Add(New FoundParam With {.P = p, .Scope = FoundScope.InstanceParam})
                Next
            Catch
            End Try

            ' Type params
            Try
                If fi.Symbol IsNot Nothing Then
                    For Each p As Parameter In fi.Symbol.Parameters
                        If p Is Nothing OrElse p.Definition Is Nothing Then Continue For
                        Dim name As String = p.Definition.Name
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        If Not map.ContainsKey(name) Then map(name) = New List(Of FoundParam)()
                        map(name).Add(New FoundParam With {.P = p, .Scope = FoundScope.TypeParam})
                    Next
                End If
            Catch
            End Try

            Return map
        End Function

        ' ✅ 기존 방식(Definition -> ExternalDefinition) 유지(호스트 FamilyParameter쪽에서 필요)
        Private Shared Function TryGetDefinitionGuid(defn As Definition, ByRef guid As Guid) As Boolean
            guid = Guid.Empty
            Try
                Dim ext As ExternalDefinition = TryCast(defn, ExternalDefinition)
                If ext IsNot Nothing Then
                    guid = ext.GUID
                    If guid <> Guid.Empty Then Return True
                End If
            Catch
            End Try
            Return False
        End Function

        ' ✅ 핵심: Parameter 자체에서 Shared/GUID 추출 (Reflection로 안전하게)
        Private Shared Function TryGetParameterIsShared(p As Parameter, ByRef isShared As Boolean) As Boolean
            isShared = False
            If p Is Nothing Then Return False
            Try
                Dim t As Type = p.GetType()
                Dim prop As Reflection.PropertyInfo = t.GetProperty("IsShared")
                If prop Is Nothing Then Return False
                Dim v As Object = prop.GetValue(p, Nothing)
                If TypeOf v Is Boolean Then
                    isShared = CBool(v)
                    Return True
                End If
            Catch
            End Try
            Return False
        End Function

        Private Shared Function TryGetParameterGuid(p As Parameter, ByRef guid As Guid) As Boolean
            guid = Guid.Empty
            If p Is Nothing Then Return False

            ' 1) Parameter.IsShared/Parameter.GUID 우선
            Try
                Dim isShared As Boolean = False
                Dim isSharedKnown As Boolean = TryGetParameterIsShared(p, isShared)
                If isSharedKnown AndAlso isShared Then
                    Dim t As Type = p.GetType()

                    Dim propGuid As Reflection.PropertyInfo = t.GetProperty("GUID")
                    If propGuid Is Nothing Then
                        propGuid = t.GetProperty("Guid") ' 혹시 몰라서 폴백
                    End If

                    If propGuid IsNot Nothing Then
                        Dim v As Object = propGuid.GetValue(p, Nothing)
                        If TypeOf v Is Guid Then
                            guid = CType(v, Guid)
                            If guid <> Guid.Empty Then Return True
                        End If
                    End If
                End If
            Catch
            End Try

            ' 2) 폴백: Definition -> ExternalDefinition
            Return TryGetDefinitionGuid(p.Definition, guid)
        End Function

        Private Shared Function SafeDefTypeToken(defn As Definition) As String
            If defn Is Nothing Then Return ""
            Try
#If REVIT2023 = 1 Then
                Dim dt As ForgeTypeId = defn.GetDataType()
                If dt IsNot Nothing Then Return SafeStr(dt.TypeId)
                Return ""
#Else
                Return SafeStr(defn.ParameterType.ToString())
#End If
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function OpenProjectDocument(app As Autodesk.Revit.ApplicationServices.Application, path As String) As Document
            If String.IsNullOrWhiteSpace(path) Then Throw New ArgumentException("path is empty.")
            If Not File.Exists(path) Then Throw New FileNotFoundException("RVT not found.", path)

            Dim mp As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)

            Dim opts As New OpenOptions()
            opts.Audit = False

            Try
                opts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
            Catch
            End Try

            Try
                Dim ws As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                opts.SetOpenWorksetsConfiguration(ws)
            Catch
            End Try

            Try
                Return app.OpenDocumentFile(mp, opts)
            Catch
                Dim opts2 As New OpenOptions()
                opts2.Audit = False
                Try
                    Dim ws2 As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                    opts2.SetOpenWorksetsConfiguration(ws2)
                Catch
                End Try
                Return app.OpenDocumentFile(mp, opts2)
            End Try
        End Function

        Private Shared Function SafeStr(s As String) As String
            Return If(s, "")
        End Function

        Private Shared Sub WriteCsv(path As String, rows As List(Of AuditRow))
            Dim headers As String() = {
                "ProjectPath",
                "HostFamilyName", "HostFamilyCategory",
                "NestedFamilyName", "NestedTypeName", "NestedCategory",
                "TargetParamName", "ExpectedGuid",
                "FoundScope", "NestedParamGuid", "NestedParamDataType",
                "AssocHostParamName", "HostParamGuid", "HostParamIsShared",
                "Issue", "Notes"
            }

            Using fs As New FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read)
                Using sw As New StreamWriter(fs, New UTF8Encoding(True))
                    sw.WriteLine(String.Join(",", headers.Select(Function(h) CsvEscape(h))))

                    For Each r As AuditRow In rows
                        Dim cols As New List(Of String) From {
                            r.ProjectPath,
                            r.HostFamilyName, r.HostFamilyCategory,
                            r.NestedFamilyName, r.NestedTypeName, r.NestedCategory,
                            r.TargetParamName, r.ExpectedGuid,
                            r.FoundScope, r.NestedParamGuid, r.NestedParamDataType,
                            r.AssocHostParamName, r.HostParamGuid, r.HostParamIsShared,
                            r.Issue, r.Notes
                        }
                        sw.WriteLine(String.Join(",", cols.Select(Function(c) CsvEscape(c))))
                    Next
                End Using
            End Using
        End Sub

        Private Shared Function CsvEscape(s As String) As String
            If s Is Nothing Then s = ""
            Dim needsQuotes As Boolean = s.Contains(","c) OrElse s.Contains(""""c) OrElse s.Contains(vbCr) OrElse s.Contains(vbLf)
            s = s.Replace("""", """""")
            If needsQuotes Then
                Return """" & s & """"
            End If
            Return s
        End Function

    End Class

End Namespace
