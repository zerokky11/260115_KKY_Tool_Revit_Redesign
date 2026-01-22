Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Drawing

Imports ClosedXML.Excel

' Revit 네임스페이스는 "Imports"로 직접 끌어오지 않고 별칭(alias)로만 사용 (모호함 방지)
Imports RUI = Autodesk.Revit.UI
Imports RDB = Autodesk.Revit.DB

Namespace ColumnEqCodeBatchTools

    ' ============================================================
    ' 단일 진입점: UI 런처
    ' ============================================================
    <Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)>
    <Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)>
    Public Class Cmd_MainUI
        Implements RUI.IExternalCommand

        Public Function Execute(commandData As RUI.ExternalCommandData,
                                ByRef message As String,
                                elements As RDB.ElementSet) As RUI.Result Implements RUI.IExternalCommand.Execute
            Try
                DependencyResolver.Ensure()

                Dim uiapp As RUI.UIApplication = commandData.Application
                Using f As New MainForm(uiapp)
                    f.StartPosition = FormStartPosition.CenterScreen
                    f.ShowDialog()
                End Using

                Return RUI.Result.Succeeded
            Catch ex As Exception
                RUI.TaskDialog.Show("Column EqCode Batch Tools", "실행 실패: " & ex.Message)
                Return RUI.Result.Failed
            End Try
        End Function
    End Class

    ' ============================================================
    ' Dependency loader (ClosedXML 등) - Addin 폴더에서 DLL 로드
    ' ============================================================
    Friend Module DependencyResolver
        Private _installed As Boolean = False

        Friend Sub Ensure()
            If _installed Then Return
            _installed = True
            AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf ResolveFromAddinFolder
        End Sub

        Private Function ResolveFromAddinFolder(sender As Object, args As ResolveEventArgs) As Assembly
            Try
                Dim asmName As New AssemblyName(args.Name)
                Dim baseDir As String = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim candidate As String = System.IO.Path.Combine(baseDir, asmName.Name & ".dll")
                If File.Exists(candidate) Then
                    Return Assembly.LoadFrom(candidate)
                End If
            Catch
            End Try
            Return Nothing
        End Function
    End Module

    ' ============================================================
    ' Main UI (WinForms / Modal)
    ' ============================================================
    Friend Class MainForm
        Inherits System.Windows.Forms.Form

        Private ReadOnly _uiapp As RUI.UIApplication
        Private ReadOnly _doc As RDB.Document

        Private tab As TabControl

        ' Tab1 controls
        Private dgvLinks As DataGridView
        Private btnRefreshLinks As Button
        Private btnExport As Button
        Private txtExportPath As TextBox
        Private btnBrowseExport As Button

        ' Tab2 controls
        Private txtRawdataPath As TextBox
        Private btnBrowseRawdata As Button
        Private txtMappingOutPath As TextBox
        Private btnBrowseMappingOut As Button
        Private nudTol As NumericUpDown
        Private nudExact As NumericUpDown
        Private btnBuildMapping As Button

        ' Tab3 controls
        Private txtMappingInPath As TextBox
        Private btnBrowseMappingIn As Button
        Private clbSPaths As CheckedListBox
        Private btnCheckAllS As Button
        Private btnUncheckAllS As Button
        Private chkSaveAsNew As CheckBox
        Private txtSuffix As TextBox
        Private btnApply As Button

        ' Log
        Private txtLog As TextBox

        Public Sub New(uiapp As RUI.UIApplication)
            _uiapp = uiapp
            _doc = uiapp.ActiveUIDocument.Document

            Me.Text = "Column EqCode Batch Tools (Revit 2019)"
            Me.Width = 980
            Me.Height = 720
            Me.MinimizeBox = True
            Me.MaximizeBox = True

            BuildUI()
            LoadLinksToGrid()
        End Sub

        Private Sub BuildUI()
            tab = New TabControl()
            tab.Dock = DockStyle.Top
            tab.Height = 520

            Dim tp1 As New TabPage("1) Export RAWDATA")
            Dim tp2 As New TabPage("2) Build Mapping")
            Dim tp3 As New TabPage("3) Apply Mapping")

            BuildTab1(tp1)
            BuildTab2(tp2)
            BuildTab3(tp3)

            tab.TabPages.Add(tp1)
            tab.TabPages.Add(tp2)
            tab.TabPages.Add(tp3)

            txtLog = New TextBox()
            txtLog.Dock = DockStyle.Fill
            txtLog.Multiline = True
            txtLog.ScrollBars = ScrollBars.Vertical
            txtLog.ReadOnly = True
            txtLog.Font = New Font("Consolas", 9.0F)

            Dim lblLog As New Label()
            lblLog.Dock = DockStyle.Top
            lblLog.Height = 20
            lblLog.Text = "Log"
            lblLog.Font = New Font(lblLog.Font, FontStyle.Bold)

            Dim pnlLog As New System.Windows.Forms.Panel()
            pnlLog.Dock = DockStyle.Fill
            pnlLog.Padding = New Padding(10, 0, 10, 10)
            pnlLog.Controls.Add(txtLog)
            pnlLog.Controls.Add(lblLog)

            Me.Controls.Add(pnlLog)
            Me.Controls.Add(tab)
        End Sub

        Private Sub BuildTab1(tp As TabPage)
            tp.Padding = New Padding(10)

            Dim lblDesc As New Label()
            lblDesc.AutoSize = False
            lblDesc.Dock = DockStyle.Top
            lblDesc.Height = 70
            lblDesc.Text =
                "설명" & vbCrLf &
                "- 현재 프로젝트에 '로드된 링크' 목록을 보여줍니다." & vbCrLf &
                "- 각 링크를 SC(Structural Columns 추출) / GM(Generic Models 추출)로 체크해서 구분합니다." & vbCrLf &
                "- Export 버튼을 누르면 RAWDATA 엑셀을 생성합니다."

            dgvLinks = New DataGridView()
            dgvLinks.Dock = DockStyle.Top
            dgvLinks.Height = 340
            dgvLinks.AllowUserToAddRows = False
            dgvLinks.AllowUserToDeleteRows = False
            dgvLinks.RowHeadersVisible = False
            dgvLinks.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgvLinks.MultiSelect = False
            dgvLinks.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            Dim colSC As New DataGridViewCheckBoxColumn()
            colSC.HeaderText = "SC"
            colSC.Width = 40
            colSC.FillWeight = 7

            Dim colGM As New DataGridViewCheckBoxColumn()
            colGM.HeaderText = "GM"
            colGM.Width = 40
            colGM.FillWeight = 7

            Dim colLoaded As New DataGridViewTextBoxColumn()
            colLoaded.HeaderText = "Loaded"
            colLoaded.FillWeight = 8

            Dim colName As New DataGridViewTextBoxColumn()
            colName.HeaderText = "Link Name"
            colName.FillWeight = 20

            Dim colInstId As New DataGridViewTextBoxColumn()
            colInstId.HeaderText = "InstanceId"
            colInstId.FillWeight = 10

            Dim colPath As New DataGridViewTextBoxColumn()
            colPath.HeaderText = "Path"
            colPath.FillWeight = 48

            dgvLinks.Columns.AddRange(New DataGridViewColumn() {colSC, colGM, colLoaded, colName, colInstId, colPath})

            AddHandler dgvLinks.CellValueChanged, AddressOf LinksGrid_CellValueChanged
            AddHandler dgvLinks.CurrentCellDirtyStateChanged,
                Sub(sender, e)
                    If dgvLinks.IsCurrentCellDirty Then dgvLinks.CommitEdit(DataGridViewDataErrorContexts.Commit)
                End Sub

            btnRefreshLinks = New Button()
            btnRefreshLinks.Text = "Refresh Link List"
            btnRefreshLinks.Width = 160
            AddHandler btnRefreshLinks.Click, Sub() LoadLinksToGrid()

            txtExportPath = New TextBox()
            txtExportPath.Width = 600

            btnBrowseExport = New Button()
            btnBrowseExport.Text = "Browse..."
            btnBrowseExport.Width = 90
            AddHandler btnBrowseExport.Click, AddressOf BrowseExport_Click

            btnExport = New Button()
            btnExport.Text = "Export RAWDATA"
            btnExport.Width = 160
            btnExport.Height = 30
            AddHandler btnExport.Click, AddressOf Export_Click

            Dim pnlBottom As New FlowLayoutPanel()
            pnlBottom.Dock = DockStyle.Top
            pnlBottom.Height = 80
            pnlBottom.FlowDirection = FlowDirection.LeftToRight
            pnlBottom.WrapContents = True

            pnlBottom.Controls.Add(btnRefreshLinks)

            Dim lblOut As New Label()
            lblOut.Text = "Output:"
            lblOut.AutoSize = True
            lblOut.Margin = New Padding(20, 8, 0, 0)
            pnlBottom.Controls.Add(lblOut)

            pnlBottom.Controls.Add(txtExportPath)
            pnlBottom.Controls.Add(btnBrowseExport)
            pnlBottom.Controls.Add(btnExport)

            tp.Controls.Add(pnlBottom)
            tp.Controls.Add(dgvLinks)
            tp.Controls.Add(lblDesc)
        End Sub

        Private Sub BuildTab2(tp As TabPage)
            tp.Padding = New Padding(10)

            Dim lblDesc As New Label()
            lblDesc.AutoSize = False
            lblDesc.Dock = DockStyle.Top
            lblDesc.Height = 85
            lblDesc.Text =
                "설명" & vbCrLf &
                "- 1단계 RAWDATA 엑셀을 입력으로 받아, 파일명 규칙으로 A/S 쌍을 묶고 XY로 1:1 매핑합니다." & vbCrLf &
                "- 결과는 MAPPING 엑셀로 출력됩니다. (EXACT / MATCH / UNMATCH)" & vbCrLf &
                "- Tolerance(mm): 허용 거리, Exact(mm): 완전일치 판정 거리"

            Dim pnl As New TableLayoutPanel()
            pnl.Dock = DockStyle.Top
            pnl.Height = 220
            pnl.ColumnCount = 4
            pnl.RowCount = 5
            pnl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
            pnl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            pnl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 90))
            pnl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            pnl.RowStyles.Add(New RowStyle(SizeType.Absolute, 35))
            pnl.RowStyles.Add(New RowStyle(SizeType.Absolute, 35))
            pnl.RowStyles.Add(New RowStyle(SizeType.Absolute, 35))
            pnl.RowStyles.Add(New RowStyle(SizeType.Absolute, 35))
            pnl.RowStyles.Add(New RowStyle(SizeType.Absolute, 50))

            txtRawdataPath = New TextBox()
            btnBrowseRawdata = New Button() With {.Text = "Browse..."}
            AddHandler btnBrowseRawdata.Click, AddressOf BrowseRawdata_Click

            txtMappingOutPath = New TextBox()
            btnBrowseMappingOut = New Button() With {.Text = "Browse..."}
            AddHandler btnBrowseMappingOut.Click, AddressOf BrowseMappingOut_Click

            nudTol = New NumericUpDown()
            nudTol.Minimum = 1
            nudTol.Maximum = 10000
            nudTol.Value = 50
            nudTol.DecimalPlaces = 0

            nudExact = New NumericUpDown()
            nudExact.Minimum = 0
            nudExact.Maximum = 1000
            nudExact.Value = 1
            nudExact.DecimalPlaces = 0

            btnBuildMapping = New Button()
            btnBuildMapping.Text = "Build Mapping"
            btnBuildMapping.Width = 160
            btnBuildMapping.Height = 30
            AddHandler btnBuildMapping.Click, AddressOf BuildMapping_Click

            pnl.Controls.Add(New Label() With {.Text = "RAWDATA:", .TextAlign = ContentAlignment.MiddleRight}, 0, 0)
            pnl.Controls.Add(txtRawdataPath, 1, 0)
            pnl.Controls.Add(btnBrowseRawdata, 2, 0)

            pnl.Controls.Add(New Label() With {.Text = "Output:", .TextAlign = ContentAlignment.MiddleRight}, 0, 1)
            pnl.Controls.Add(txtMappingOutPath, 1, 1)
            pnl.Controls.Add(btnBrowseMappingOut, 2, 1)

            pnl.Controls.Add(New Label() With {.Text = "Tolerance(mm):", .TextAlign = ContentAlignment.MiddleRight}, 0, 2)
            pnl.Controls.Add(nudTol, 1, 2)

            pnl.Controls.Add(New Label() With {.Text = "Exact(mm):", .TextAlign = ContentAlignment.MiddleRight}, 0, 3)
            pnl.Controls.Add(nudExact, 1, 3)

            pnl.Controls.Add(btnBuildMapping, 3, 4)

            tp.Controls.Add(pnl)
            tp.Controls.Add(lblDesc)
        End Sub

        Private Sub BuildTab3(tp As TabPage)
            tp.Padding = New Padding(10)

            Dim lblDesc As New Label()
            lblDesc.AutoSize = False
            lblDesc.Dock = DockStyle.Top
            lblDesc.Height = 95
            lblDesc.Text =
                "설명" & vbCrLf &
                "- 검토/정리한 MAPPING 엑셀을 입력으로 받아, ColumnNumber 값을 구조기둥(S 파일)의 S5_EQCODE에 입력합니다." & vbCrLf &
                "- 아래 리스트는 MAPPING에서 발견된 S_FullPath 목록입니다. 적용할 S 파일만 체크하세요." & vbCrLf &
                "- SaveAs New 권장(원본 보호)"

            Dim pnlTop As New TableLayoutPanel()
            pnlTop.Dock = DockStyle.Top
            pnlTop.Height = 110
            pnlTop.ColumnCount = 4
            pnlTop.RowCount = 3
            pnlTop.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
            pnlTop.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            pnlTop.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 90))
            pnlTop.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            pnlTop.RowStyles.Add(New RowStyle(SizeType.Absolute, 35))
            pnlTop.RowStyles.Add(New RowStyle(SizeType.Absolute, 35))
            pnlTop.RowStyles.Add(New RowStyle(SizeType.Absolute, 35))

            txtMappingInPath = New TextBox()
            btnBrowseMappingIn = New Button() With {.Text = "Browse..."}
            AddHandler btnBrowseMappingIn.Click, AddressOf BrowseMappingIn_Click

            chkSaveAsNew = New CheckBox() With {.Text = "SaveAs New", .Checked = True, .AutoSize = True}
            txtSuffix = New TextBox() With {.Text = "_EQCODE", .Width = 90}

            btnApply = New Button()
            btnApply.Text = "Apply Mapping"
            btnApply.Width = 160
            btnApply.Height = 30
            AddHandler btnApply.Click, AddressOf Apply_Click

            pnlTop.Controls.Add(New Label() With {.Text = "MAPPING:", .TextAlign = ContentAlignment.MiddleRight}, 0, 0)
            pnlTop.Controls.Add(txtMappingInPath, 1, 0)
            pnlTop.Controls.Add(btnBrowseMappingIn, 2, 0)

            pnlTop.Controls.Add(chkSaveAsNew, 1, 1)

            Dim pnlSuffix As New FlowLayoutPanel() With {.FlowDirection = FlowDirection.LeftToRight, .Dock = DockStyle.Fill}
            pnlSuffix.Controls.Add(New Label() With {.Text = "Suffix:", .AutoSize = True, .Margin = New Padding(0, 7, 5, 0)})
            pnlSuffix.Controls.Add(txtSuffix)
            pnlTop.Controls.Add(pnlSuffix, 2, 1)

            pnlTop.Controls.Add(btnApply, 3, 2)

            clbSPaths = New CheckedListBox()
            clbSPaths.Dock = DockStyle.Fill

            btnCheckAllS = New Button() With {.Text = "Check All", .Width = 100}
            btnUncheckAllS = New Button() With {.Text = "Uncheck All", .Width = 100}
            AddHandler btnCheckAllS.Click,
                Sub()
                    For i As Integer = 0 To clbSPaths.Items.Count - 1
                        clbSPaths.SetItemChecked(i, True)
                    Next
                End Sub
            AddHandler btnUncheckAllS.Click,
                Sub()
                    For i As Integer = 0 To clbSPaths.Items.Count - 1
                        clbSPaths.SetItemChecked(i, False)
                    Next
                End Sub

            Dim pnlButtons As New FlowLayoutPanel()
            pnlButtons.Dock = DockStyle.Top
            pnlButtons.Height = 40
            pnlButtons.Controls.Add(btnCheckAllS)
            pnlButtons.Controls.Add(btnUncheckAllS)

            Dim pnlList As New System.Windows.Forms.Panel()
            pnlList.Dock = DockStyle.Fill
            pnlList.Controls.Add(clbSPaths)
            pnlList.Controls.Add(pnlButtons)

            tp.Controls.Add(pnlList)
            tp.Controls.Add(pnlTop)
            tp.Controls.Add(lblDesc)
        End Sub

        ' =========================
        ' Tab1: Links
        ' =========================
        Private Sub LoadLinksToGrid()
            dgvLinks.Rows.Clear()

            Dim links As List(Of RDB.RevitLinkInstance) = BatchCore.GetLoadedLinkInstances(_doc)

            For Each li As RDB.RevitLinkInstance In links
                Dim linkDoc As RDB.Document = li.GetLinkDocument()
                Dim loaded As Boolean = (linkDoc IsNot Nothing)

                Dim fileName As String = ""
                Dim linkPath As String = ""

                If loaded Then
                    linkPath = If(linkDoc.PathName, "")
                    If Not String.IsNullOrWhiteSpace(linkPath) Then
                        fileName = System.IO.Path.GetFileName(linkPath)
                    Else
                        fileName = linkDoc.Title & ".rvt"
                    End If
                Else
                    fileName = "(UNLOADED)"
                    linkPath = ""
                End If

                Dim idx As Integer = dgvLinks.Rows.Add(False, False, If(loaded, "Yes", "No"), fileName, li.Id.IntegerValue.ToString(), linkPath)
                dgvLinks.Rows(idx).Tag = li

                If Not loaded Then
                    dgvLinks.Rows(idx).DefaultCellStyle.ForeColor = System.Drawing.Color.Gray
                    dgvLinks.Rows(idx).ReadOnly = True
                End If
            Next

            If String.IsNullOrWhiteSpace(txtExportPath.Text) Then
                txtExportPath.Text = System.IO.Path.Combine(BatchCore.GetDesktop(), $"RAWDATA_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx")
            End If

            Log($"[Link List] Loaded links: {links.Count}")
        End Sub

        ' SC/GM 상호 배타 처리
        Private Sub LinksGrid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
            If e.ColumnIndex <> 0 AndAlso e.ColumnIndex <> 1 Then Return

            Dim row = dgvLinks.Rows(e.RowIndex)
            If row Is Nothing OrElse row.ReadOnly Then Return

            Dim scVal As Boolean = ToBool(row.Cells(0).Value)
            Dim gmVal As Boolean = ToBool(row.Cells(1).Value)

            If e.ColumnIndex = 0 AndAlso scVal Then
                row.Cells(1).Value = False
            ElseIf e.ColumnIndex = 1 AndAlso gmVal Then
                row.Cells(0).Value = False
            End If
        End Sub

        Private Shared Function ToBool(v As Object) As Boolean
            If v Is Nothing Then Return False
            Try
                Return Convert.ToBoolean(v)
            Catch
                Return False
            End Try
        End Function

        Private Sub BrowseExport_Click(sender As Object, e As EventArgs)
            Using sfd As New SaveFileDialog()
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                sfd.InitialDirectory = BatchCore.GetDesktop()
                sfd.FileName = $"RAWDATA_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                If sfd.ShowDialog() = DialogResult.OK Then
                    txtExportPath.Text = sfd.FileName
                End If
            End Using
        End Sub

        Private Sub Export_Click(sender As Object, e As EventArgs)
            Dim scLinks As New List(Of RDB.RevitLinkInstance)()
            Dim gmLinks As New List(Of RDB.RevitLinkInstance)()

            For Each r As DataGridViewRow In dgvLinks.Rows
                If r Is Nothing OrElse r.Tag Is Nothing Then Continue For
                If r.ReadOnly Then Continue For

                Dim li As RDB.RevitLinkInstance = TryCast(r.Tag, RDB.RevitLinkInstance)
                If li Is Nothing Then Continue For

                Dim isSC As Boolean = ToBool(r.Cells(0).Value)
                Dim isGM As Boolean = ToBool(r.Cells(1).Value)

                If isSC Then scLinks.Add(li)
                If isGM Then gmLinks.Add(li)
            Next

            If scLinks.Count = 0 OrElse gmLinks.Count = 0 Then
                MessageBox.Show("SC와 GM 링크를 각각 최소 1개 이상 체크하세요.", "Export RAWDATA", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim outPath As String = txtExportPath.Text.Trim()
            If String.IsNullOrWhiteSpace(outPath) Then
                MessageBox.Show("Output 경로를 지정하세요.", "Export RAWDATA", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                Cursor.Current = Cursors.WaitCursor
                Log("[Export] Start")
                Dim summary As String = BatchCore.ExportRawDataToExcel(_uiapp, scLinks, gmLinks, outPath)
                Log(summary)
                MessageBox.Show("완료" & vbCrLf & outPath, "Export RAWDATA", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                Log("[Export] FAIL: " & ex.Message)
                MessageBox.Show(ex.Message, "Export RAWDATA - Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Cursor.Current = Cursors.Default
            End Try
        End Sub

        ' =========================
        ' Tab2: Build Mapping
        ' =========================
        Private Sub BrowseRawdata_Click(sender As Object, e As EventArgs)
            Using ofd As New OpenFileDialog()
                ofd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                ofd.InitialDirectory = BatchCore.GetDesktop()
                If ofd.ShowDialog() = DialogResult.OK Then
                    txtRawdataPath.Text = ofd.FileName
                    If String.IsNullOrWhiteSpace(txtMappingOutPath.Text) Then
                        txtMappingOutPath.Text = System.IO.Path.Combine(BatchCore.GetDesktop(), $"MAPPING_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx")
                    End If
                End If
            End Using
        End Sub

        Private Sub BrowseMappingOut_Click(sender As Object, e As EventArgs)
            Using sfd As New SaveFileDialog()
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                sfd.InitialDirectory = BatchCore.GetDesktop()
                sfd.FileName = $"MAPPING_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                If sfd.ShowDialog() = DialogResult.OK Then
                    txtMappingOutPath.Text = sfd.FileName
                End If
            End Using
        End Sub

        Private Sub BuildMapping_Click(sender As Object, e As EventArgs)
            Dim inPath As String = txtRawdataPath.Text.Trim()
            Dim outPath As String = txtMappingOutPath.Text.Trim()

            If String.IsNullOrWhiteSpace(inPath) OrElse Not File.Exists(inPath) Then
                MessageBox.Show("RAWDATA 파일을 선택하세요.", "Build Mapping", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If String.IsNullOrWhiteSpace(outPath) Then
                MessageBox.Show("Output 경로를 지정하세요.", "Build Mapping", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim tolMm As Double = CDbl(nudTol.Value)
            Dim exactMm As Double = CDbl(nudExact.Value)

            Try
                Cursor.Current = Cursors.WaitCursor
                Log("[BuildMapping] Start")
                Dim summary As String = BatchCore.BuildMappingToExcel(inPath, outPath, tolMm, exactMm)
                Log(summary)
                MessageBox.Show("완료" & vbCrLf & outPath, "Build Mapping", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                Log("[BuildMapping] FAIL: " & ex.Message)
                MessageBox.Show(ex.Message, "Build Mapping - Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Cursor.Current = Cursors.Default
            End Try
        End Sub

        ' =========================
        ' Tab3: Apply Mapping
        ' =========================
        Private Sub BrowseMappingIn_Click(sender As Object, e As EventArgs)
            Using ofd As New OpenFileDialog()
                ofd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                ofd.InitialDirectory = BatchCore.GetDesktop()
                If ofd.ShowDialog() = DialogResult.OK Then
                    txtMappingInPath.Text = ofd.FileName
                    ReloadSPathListFromMapping(ofd.FileName)
                End If
            End Using
        End Sub

        Private Sub ReloadSPathListFromMapping(mappingPath As String)
            clbSPaths.Items.Clear()
            Try
                Dim paths As List(Of String) = BatchCore.ReadDistinctSPathsFromMapping(mappingPath)
                For Each p In paths
                    clbSPaths.Items.Add(p, True)
                Next
                Log($"[Apply] Loaded S_FullPath list: {paths.Count}")
            Catch ex As Exception
                Log("[Apply] FAIL to read S paths: " & ex.Message)
            End Try
        End Sub

        Private Sub Apply_Click(sender As Object, e As EventArgs)
            Dim mappingPath As String = txtMappingInPath.Text.Trim()
            If String.IsNullOrWhiteSpace(mappingPath) OrElse Not File.Exists(mappingPath) Then
                MessageBox.Show("MAPPING 엑셀 파일을 선택하세요.", "Apply Mapping", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim selectedS As New List(Of String)()
            For i As Integer = 0 To clbSPaths.Items.Count - 1
                If clbSPaths.GetItemChecked(i) Then
                    selectedS.Add(clbSPaths.Items(i).ToString())
                End If
            Next

            Dim saveAsNew As Boolean = chkSaveAsNew.Checked
            Dim suffix As String = txtSuffix.Text.Trim()
            If saveAsNew AndAlso String.IsNullOrWhiteSpace(suffix) Then
                MessageBox.Show("SaveAs New 사용 시 suffix를 입력하세요. (예: _EQCODE)", "Apply Mapping", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                Cursor.Current = Cursors.WaitCursor
                Log("[Apply] Start")
                Dim summary As String = BatchCore.ApplyMappingFromExcel(_uiapp, mappingPath, selectedS, saveAsNew, suffix)
                Log(summary)
                MessageBox.Show("Apply 완료 (바탕화면 APPLY_LOG 엑셀 확인)", "Apply Mapping", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                Log("[Apply] FAIL: " & ex.Message)
                MessageBox.Show(ex.Message, "Apply Mapping - Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Cursor.Current = Cursors.Default
            End Try
        End Sub

        Private Sub Log(msg As String)
            If String.IsNullOrWhiteSpace(msg) Then Return
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}")
        End Sub

    End Class

    ' ============================================================
    ' Core logic
    ' ============================================================
    Friend Module BatchCore

        Private ReadOnly RAW_HEADERS As String() = New String() {
            "Id", "ElementName", "CategoryName",
            "LocationX", "LocationY", "LocationZ",
            "Grid",
            "SB_NAME", "SB_FIELD", "SB_FL", "SB_BLDG",
            "Stud Number", "Column Number",
            "EXCLUSION", "CLASS",
            "SAMOO-Other_1", "SAMOO-Other_2", "SAMOO-Other_3",
            "SECC-Other_3", "SECC-Other_4"
        }

        Private Const TARGET_PARAM_EQCODE As String = "S5_EQCODE"
        Private Const ALSO_SET_TARGET_COLNO As Boolean = False
        Private Const TARGET_PARAM_COLNO As String = "Column Number"

        Private Class LinkInfo
            Public Kind As String
            Public FileName As String
            Public FullPath As String
            Public LinkInstanceId As Integer
            Public SheetName As String
        End Class

        Private Class RawRow
            Public Kind As String
            Public FileName As String
            Public FullPath As String
            Public LinkInstanceId As Integer

            Public Id As Integer
            Public UniqueId As String

            Public ElementName As String
            Public CategoryName As String

            Public Xmm As Double
            Public Ymm As Double
            Public Zmm As Double

            Public Grid As String
            Public StudNumber As String
            Public ColumnNumber As String
        End Class

        Private Class MapRow
            Public PairKey As String

            Public A_File As String
            Public A_Path As String
            Public A_Id As Integer
            Public A_Uid As String
            Public A_Xmm As Double
            Public A_Ymm As Double
            Public A_ColumnNumber As String

            Public S_File As String
            Public S_Path As String
            Public S_Id As Integer
            Public S_Uid As String
            Public S_Xmm As Double
            Public S_Ymm As Double

            Public DistMm As Double
            Public Status As String
            Public Note As String
        End Class

        Private Structure CellKey
            Public ReadOnly X As Integer
            Public ReadOnly Y As Integer
            Public Sub New(ix As Integer, iy As Integer)
                X = ix : Y = iy
            End Sub
            Public Overrides Function GetHashCode() As Integer
                Return (X * 397) Xor Y
            End Function
            Public Overrides Function Equals(obj As Object) As Boolean
                If Not TypeOf obj Is CellKey Then Return False
                Dim o As CellKey = CType(obj, CellKey)
                Return o.X = X AndAlso o.Y = Y
            End Function
        End Structure

        Friend Function GetDesktop() As String
            Return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        End Function

        ' 로드된 링크만 반환
        Friend Function GetLoadedLinkInstances(doc As RDB.Document) As List(Of RDB.RevitLinkInstance)
            Dim result As New List(Of RDB.RevitLinkInstance)()

            Dim col As New RDB.FilteredElementCollector(doc)
            col.OfClass(GetType(RDB.RevitLinkInstance))

            For Each e As RDB.Element In col
                Dim li As RDB.RevitLinkInstance = TryCast(e, RDB.RevitLinkInstance)
                If li Is Nothing Then Continue For
                If li.GetLinkDocument() IsNot Nothing Then
                    result.Add(li)
                End If
            Next

            Return result
        End Function

        ' ==========================================================
        ' (1) Export RAWDATA
        ' ==========================================================
        Friend Function ExportRawDataToExcel(uiapp As RUI.UIApplication,
                                            scLinks As IList(Of RDB.RevitLinkInstance),
                                            gmLinks As IList(Of RDB.RevitLinkInstance),
                                            outPath As String) As String

            If scLinks Is Nothing OrElse scLinks.Count = 0 Then Throw New InvalidOperationException("SC 링크가 비어있습니다.")
            If gmLinks Is Nothing OrElse gmLinks.Count = 0 Then Throw New InvalidOperationException("GM 링크가 비어있습니다.")
            If String.IsNullOrWhiteSpace(outPath) Then Throw New InvalidOperationException("Output 경로가 비어있습니다.")

            Dim wb As New XLWorkbook()

            Dim wsLinks As IXLWorksheet = wb.Worksheets.Add("LINKS")
            wsLinks.Cell(1, 1).Value = "Kind"
            wsLinks.Cell(1, 2).Value = "FileName"
            wsLinks.Cell(1, 3).Value = "FullPath"
            wsLinks.Cell(1, 4).Value = "LinkInstanceId"
            wsLinks.Cell(1, 5).Value = "SheetName"
            wsLinks.Row(1).Style.Font.Bold = True

            Dim wsUid As IXLWorksheet = wb.Worksheets.Add("UIDMAP")
            wsUid.Cell(1, 1).Value = "Kind"
            wsUid.Cell(1, 2).Value = "FileName"
            wsUid.Cell(1, 3).Value = "LinkInstanceId"
            wsUid.Cell(1, 4).Value = "Id"
            wsUid.Cell(1, 5).Value = "UniqueId"
            wsUid.Row(1).Style.Font.Bold = True

            Dim usedSheetNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim linkRow As Integer = 2
            Dim uidRow As Integer = 2

            For Each li As RDB.RevitLinkInstance In scLinks
                Dim info As LinkInfo = BuildLinkInfo(li, "SC")
                If info Is Nothing Then Continue For

                info.SheetName = MakeUniqueSheetName(wb, usedSheetNames, "SC_" & SafeSheetBase(info.FileName) & "_" & info.LinkInstanceId.ToString())
                WriteLinksRow(wsLinks, linkRow, info) : linkRow += 1

                Dim ws As IXLWorksheet = wb.Worksheets.Add(info.SheetName)
                WriteRawHeader(ws)

                Dim wrote As Integer = WriteRawRowsFromLink(li, RDB.BuiltInCategory.OST_StructuralColumns, ws, info, wsUid, uidRow)
                uidRow += wrote
            Next

            For Each li As RDB.RevitLinkInstance In gmLinks
                Dim info As LinkInfo = BuildLinkInfo(li, "GM")
                If info Is Nothing Then Continue For

                info.SheetName = MakeUniqueSheetName(wb, usedSheetNames, "GM_" & SafeSheetBase(info.FileName) & "_" & info.LinkInstanceId.ToString())
                WriteLinksRow(wsLinks, linkRow, info) : linkRow += 1

                Dim ws As IXLWorksheet = wb.Worksheets.Add(info.SheetName)
                WriteRawHeader(ws)

                Dim wrote As Integer = WriteRawRowsFromLink(li, RDB.BuiltInCategory.OST_GenericModel, ws, info, wsUid, uidRow)
                uidRow += wrote
            Next

            wsLinks.Columns().AdjustToContents()
            wsUid.Columns().AdjustToContents()

            wb.SaveAs(outPath)

            Return $"Export 완료: {outPath} (Links={linkRow - 2}, UIDRows={uidRow - 2})"
        End Function

        Private Function BuildLinkInfo(li As RDB.RevitLinkInstance, kind As String) As LinkInfo
            Dim linkDoc As RDB.Document = li.GetLinkDocument()
            If linkDoc Is Nothing Then Return Nothing

            Dim fullPath As String = If(linkDoc.PathName, "")
            Dim fileName As String

            If Not String.IsNullOrWhiteSpace(fullPath) Then
                fileName = System.IO.Path.GetFileName(fullPath)
            Else
                fileName = linkDoc.Title & ".rvt"
            End If

            Return New LinkInfo With {
                .Kind = kind,
                .FileName = fileName,
                .FullPath = fullPath,
                .LinkInstanceId = li.Id.IntegerValue,
                .SheetName = ""
            }
        End Function

        Private Sub WriteLinksRow(ws As IXLWorksheet, row As Integer, info As LinkInfo)
            ws.Cell(row, 1).Value = info.Kind
            ws.Cell(row, 2).Value = info.FileName
            ws.Cell(row, 3).Value = info.FullPath
            ws.Cell(row, 4).Value = info.LinkInstanceId
            ws.Cell(row, 5).Value = info.SheetName
        End Sub

        Private Sub WriteRawHeader(ws As IXLWorksheet)
            For i As Integer = 0 To RAW_HEADERS.Length - 1
                ws.Cell(1, i + 1).Value = RAW_HEADERS(i)
            Next
            ws.Row(1).Style.Font.Bold = True
        End Sub

        Private Function WriteRawRowsFromLink(li As RDB.RevitLinkInstance,
                                             bic As RDB.BuiltInCategory,
                                             ws As IXLWorksheet,
                                             info As LinkInfo,
                                             wsUid As IXLWorksheet,
                                             uidRowStart As Integer) As Integer

            Dim linkDoc As RDB.Document = li.GetLinkDocument()
            If linkDoc Is Nothing Then Return 0

            Dim t As RDB.Transform = li.GetTransform()

            Dim col As New RDB.FilteredElementCollector(linkDoc)
            col.OfCategory(bic)
            col.WhereElementIsNotElementType()

            Dim row As Integer = 2
            Dim uidRow As Integer = uidRowStart

            For Each e As RDB.Element In col
                Dim ptLink As RDB.XYZ = GetInsertPoint(e)
                If ptLink Is Nothing Then Continue For

                Dim ptHost As RDB.XYZ = t.OfPoint(ptLink)

                Dim xmm As Double = RDB.UnitUtils.ConvertFromInternalUnits(ptHost.X, RDB.DisplayUnitType.DUT_MILLIMETERS)
                Dim ymm As Double = RDB.UnitUtils.ConvertFromInternalUnits(ptHost.Y, RDB.DisplayUnitType.DUT_MILLIMETERS)
                Dim zmm As Double = RDB.UnitUtils.ConvertFromInternalUnits(ptHost.Z, RDB.DisplayUnitType.DUT_MILLIMETERS)

                Dim elementName As String = GetElementTypeDisplayName(linkDoc, e)
                Dim categoryName As String = If(e.Category Is Nothing, "", e.Category.Name)

                Dim values As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                values("Id") = e.Id.IntegerValue
                values("ElementName") = elementName
                values("CategoryName") = categoryName
                values("LocationX") = Math.Round(xmm, 3)
                values("LocationY") = Math.Round(ymm, 3)
                values("LocationZ") = Math.Round(zmm, 3)

                For Each h As String In RAW_HEADERS
                    If values.ContainsKey(h) Then Continue For
                    values(h) = GetParamString(linkDoc, e, h)
                Next

                For i As Integer = 0 To RAW_HEADERS.Length - 1
                    ws.Cell(row, i + 1).Value = values(RAW_HEADERS(i))
                Next

                wsUid.Cell(uidRow, 1).Value = info.Kind
                wsUid.Cell(uidRow, 2).Value = info.FileName
                wsUid.Cell(uidRow, 3).Value = info.LinkInstanceId
                wsUid.Cell(uidRow, 4).Value = e.Id.IntegerValue
                wsUid.Cell(uidRow, 5).Value = e.UniqueId

                row += 1
                uidRow += 1
            Next

            ws.Columns().AdjustToContents()
            Return uidRow - uidRowStart
        End Function

        ' ==========================================================
        ' (2) Build Mapping
        ' ==========================================================
        Friend Function BuildMappingToExcel(rawdataPath As String,
                                            outPath As String,
                                            tolMm As Double,
                                            exactMm As Double) As String

            If String.IsNullOrWhiteSpace(rawdataPath) OrElse Not File.Exists(rawdataPath) Then
                Throw New FileNotFoundException("RAWDATA 파일을 찾을 수 없습니다.", rawdataPath)
            End If
            If String.IsNullOrWhiteSpace(outPath) Then Throw New InvalidOperationException("Output 경로가 비어있습니다.")
            If tolMm <= 0 Then Throw New InvalidOperationException("Tolerance(mm)는 0보다 커야 합니다.")
            If exactMm < 0 Then Throw New InvalidOperationException("Exact(mm)는 0 이상이어야 합니다.")

            Dim links As List(Of LinkInfo)
            Dim uidMap As Dictionary(Of String, String)
            Dim raws As List(Of RawRow)

            Using wb As New XLWorkbook(rawdataPath)
                links = ReadLinksSheet(wb)
                uidMap = ReadUidMap(wb)
                raws = ReadAllRawSheets(wb, links, uidMap)
            End Using

            Dim scAll As List(Of RawRow) = raws.Where(Function(r) r.Kind = "SC").ToList()
            Dim gmAll As List(Of RawRow) = raws.Where(Function(r) r.Kind = "GM").ToList()
            Dim gm As List(Of RawRow) = gmAll.Where(Function(r) Not String.IsNullOrWhiteSpace(r.ColumnNumber)).ToList()
            Dim sc As List(Of RawRow) = scAll

            Dim gmByKey = gm.GroupBy(Function(r) PairKeyOf(r.FileName)).ToDictionary(Function(g) g.Key, Function(g) g.ToList())
            Dim scByKey = sc.GroupBy(Function(r) PairKeyOf(r.FileName)).ToDictionary(Function(g) g.Key, Function(g) g.ToList())

            Dim keys As List(Of String) = gmByKey.Keys.Intersect(scByKey.Keys).ToList()
            If keys.Count = 0 Then Throw New InvalidOperationException("A/S 쌍을 찾지 못했습니다. 파일명 규칙(예: 1_A_a / 1_S_a)을 확인하세요.")

            Dim allMaps As New List(Of MapRow)()
            Dim allUnmatchedS As New List(Of RawRow)()

            For Each k As String In keys
                Dim gmList As List(Of RawRow) = gmByKey(k)
                Dim scList As List(Of RawRow) = scByKey(k)

                Dim aFile As String = gmList.Select(Function(x) x.FileName).Distinct().First()
                Dim aPath As String = gmList.Select(Function(x) x.FullPath).Distinct().FirstOrDefault()

                Dim sFile As String = scList.Select(Function(x) x.FileName).Distinct().First()
                Dim sPath As String = scList.Select(Function(x) x.FullPath).Distinct().FirstOrDefault()

                Dim pairMaps As List(Of MapRow) = MapByXY_OneToOne(k, aFile, aPath, gmList, sFile, sPath, scList, tolMm, exactMm)
                allMaps.AddRange(pairMaps)

                Dim matchedSIds As New HashSet(Of Integer)(
                    pairMaps.Where(Function(m) m.S_Id <> 0 AndAlso (m.Status = "EXACT" OrElse m.Status = "MATCH")).Select(Function(m) m.S_Id)
                )
                allUnmatchedS.AddRange(scList.Where(Function(s) Not matchedSIds.Contains(s.Id)))
            Next

            WriteMappingWorkbook(outPath, allMaps, tolMm, allUnmatchedS)

            Dim exactCnt As Integer = allMaps.Where(Function(m) m.Status = "EXACT").Count()
            Dim matchCnt As Integer = allMaps.Where(Function(m) m.Status = "MATCH").Count()
            Dim unCnt As Integer = allMaps.Where(Function(m) m.Status = "UNMATCH").Count()

            Return $"BuildMapping 완료: {outPath} (EXACT={exactCnt}, MATCH={matchCnt}, UNMATCH={unCnt})"
        End Function

        Private Function MapByXY_OneToOne(pairKey As String,
                                         aFile As String, aPath As String, gmList As List(Of RawRow),
                                         sFile As String, sPath As String, scList As List(Of RawRow),
                                         tolMm As Double, exactMm As Double) As List(Of MapRow)

            Dim cell As Double = tolMm

            Dim idx As New Dictionary(Of CellKey, List(Of Integer))()
            For i As Integer = 0 To scList.Count - 1
                Dim s As RawRow = scList(i)
                Dim ck As CellKey = CellKeyOf(s.Xmm, s.Ymm, cell)
                Dim lst As List(Of Integer) = Nothing
                If Not idx.TryGetValue(ck, lst) Then
                    lst = New List(Of Integer)()
                    idx.Add(ck, lst)
                End If
                lst.Add(i)
            Next

            Dim edges As New List(Of Tuple(Of Integer, Integer, Double))()

            For ai As Integer = 0 To gmList.Count - 1
                Dim a As RawRow = gmList(ai)
                Dim baseKey As CellKey = CellKeyOf(a.Xmm, a.Ymm, cell)

                For dx As Integer = -1 To 1
                    For dy As Integer = -1 To 1
                        Dim k As New CellKey(baseKey.X + dx, baseKey.Y + dy)
                        Dim lst As List(Of Integer) = Nothing
                        If Not idx.TryGetValue(k, lst) Then Continue For

                        For Each si As Integer In lst
                            Dim s As RawRow = scList(si)
                            Dim d As Double = Dist2Dmm(a.Xmm, a.Ymm, s.Xmm, s.Ymm)
                            If d <= tolMm Then
                                edges.Add(Tuple.Create(ai, si, d))
                            End If
                        Next
                    Next
                Next
            Next

            edges.Sort(Function(x, y) x.Item3.CompareTo(y.Item3))

            Dim aAssigned(gmList.Count - 1) As Integer
            Dim sAssigned(scList.Count - 1) As Integer
            For i As Integer = 0 To aAssigned.Length - 1 : aAssigned(i) = -1 : Next
            For i As Integer = 0 To sAssigned.Length - 1 : sAssigned(i) = -1 : Next

            For Each ed In edges
                Dim ai As Integer = ed.Item1
                Dim si As Integer = ed.Item2
                If aAssigned(ai) <> -1 Then Continue For
                If sAssigned(si) <> -1 Then Continue For
                aAssigned(ai) = si
                sAssigned(si) = ai
            Next

            Dim result As New List(Of MapRow)()

            For ai As Integer = 0 To gmList.Count - 1
                Dim a As RawRow = gmList(ai)

                Dim mr As New MapRow()
                mr.PairKey = pairKey

                mr.A_File = aFile
                mr.A_Path = aPath
                mr.A_Id = a.Id
                mr.A_Uid = a.UniqueId
                mr.A_Xmm = a.Xmm
                mr.A_Ymm = a.Ymm
                mr.A_ColumnNumber = If(a.ColumnNumber, "").Trim()

                mr.S_File = sFile
                mr.S_Path = sPath

                Dim si As Integer = aAssigned(ai)
                If si = -1 Then
                    mr.S_Id = 0
                    mr.S_Uid = ""
                    mr.S_Xmm = 0
                    mr.S_Ymm = 0
                    mr.DistMm = 0
                    mr.Status = "UNMATCH"
                    mr.Note = "NO_TARGET_WITHIN_TOL"
                Else
                    Dim s As RawRow = scList(si)
                    Dim d As Double = Dist2Dmm(a.Xmm, a.Ymm, s.Xmm, s.Ymm)

                    mr.S_Id = s.Id
                    mr.S_Uid = s.UniqueId
                    mr.S_Xmm = s.Xmm
                    mr.S_Ymm = s.Ymm
                    mr.DistMm = Math.Round(d, 3)
                    mr.Status = If(d <= exactMm, "EXACT", "MATCH")
                    mr.Note = ""
                End If

                result.Add(mr)
            Next

            Return result
        End Function

        Private Sub WriteMappingWorkbook(path As String, maps As List(Of MapRow), tolMm As Double, unmatchedS As List(Of RawRow))
            Using wb As New XLWorkbook()

                Dim ws As IXLWorksheet = wb.Worksheets.Add("MAPPING")

                Dim headers As String() = New String() {
                    "PairKey",
                    "A_File", "A_FullPath", "A_Id", "A_UniqueId", "A_X(mm)", "A_Y(mm)", "A_ColumnNumber",
                    "S_File", "S_FullPath", "S_Id", "S_UniqueId", "S_X(mm)", "S_Y(mm)",
                    "DistXY(mm)", "Status", "Note"
                }

                For i As Integer = 0 To headers.Length - 1
                    ws.Cell(1, i + 1).Value = headers(i)
                Next
                ws.Row(1).Style.Font.Bold = True

                Dim r As Integer = 2
                For Each m As MapRow In maps
                    ws.Cell(r, 1).Value = m.PairKey

                    ws.Cell(r, 2).Value = m.A_File
                    ws.Cell(r, 3).Value = m.A_Path
                    ws.Cell(r, 4).Value = m.A_Id
                    ws.Cell(r, 5).Value = m.A_Uid
                    ws.Cell(r, 6).Value = m.A_Xmm
                    ws.Cell(r, 7).Value = m.A_Ymm
                    ws.Cell(r, 8).Value = m.A_ColumnNumber

                    ws.Cell(r, 9).Value = m.S_File
                    ws.Cell(r, 10).Value = m.S_Path
                    ws.Cell(r, 11).Value = m.S_Id
                    ws.Cell(r, 12).Value = m.S_Uid
                    ws.Cell(r, 13).Value = m.S_Xmm
                    ws.Cell(r, 14).Value = m.S_Ymm

                    ws.Cell(r, 15).Value = m.DistMm
                    ws.Cell(r, 16).Value = m.Status
                    ws.Cell(r, 17).Value = m.Note
                    r += 1
                Next

                ws.Columns().AdjustToContents()

                Dim ws2 As IXLWorksheet = wb.Worksheets.Add("SUMMARY")
                ws2.Cell(1, 1).Value = "MapTol(mm)" : ws2.Cell(1, 2).Value = tolMm
                ws2.Cell(2, 1).Value = "EXACT" : ws2.Cell(2, 2).Value = maps.Where(Function(m) m.Status = "EXACT").Count()
                ws2.Cell(3, 1).Value = "MATCH" : ws2.Cell(3, 2).Value = maps.Where(Function(m) m.Status = "MATCH").Count()
                ws2.Cell(4, 1).Value = "UNMATCH" : ws2.Cell(4, 2).Value = maps.Where(Function(m) m.Status = "UNMATCH").Count()
                ws2.Columns().AdjustToContents()

                Dim ws3 As IXLWorksheet = wb.Worksheets.Add("UNMATCHED_S")
                ws3.Cell(1, 1).Value = "S_File"
                ws3.Cell(1, 2).Value = "S_FullPath"
                ws3.Cell(1, 3).Value = "S_Id"
                ws3.Cell(1, 4).Value = "S_UniqueId"
                ws3.Cell(1, 5).Value = "S_X(mm)"
                ws3.Cell(1, 6).Value = "S_Y(mm)"
                ws3.Row(1).Style.Font.Bold = True

                Dim rr As Integer = 2
                For Each s As RawRow In unmatchedS
                    ws3.Cell(rr, 1).Value = s.FileName
                    ws3.Cell(rr, 2).Value = s.FullPath
                    ws3.Cell(rr, 3).Value = s.Id
                    ws3.Cell(rr, 4).Value = s.UniqueId
                    ws3.Cell(rr, 5).Value = s.Xmm
                    ws3.Cell(rr, 6).Value = s.Ymm
                    rr += 1
                Next
                ws3.Columns().AdjustToContents()

                wb.SaveAs(path)
            End Using
        End Sub

        ' ==========================================================
        ' (3) Apply Mapping
        ' ==========================================================
        Friend Function ApplyMappingFromExcel(uiapp As RUI.UIApplication,
                                             mappingPath As String,
                                             selectedSPaths As IList(Of String),
                                             saveAsNew As Boolean,
                                             suffix As String) As String

            If String.IsNullOrWhiteSpace(mappingPath) OrElse Not File.Exists(mappingPath) Then
                Throw New FileNotFoundException("MAPPING 파일을 찾을 수 없습니다.", mappingPath)
            End If

            Dim maps As List(Of MapRow) = ReadMappingWorkbook(mappingPath)

            Dim applyRows = maps.Where(Function(m) (m.Status = "EXACT" OrElse m.Status = "MATCH") AndAlso
                                                   Not String.IsNullOrWhiteSpace(m.A_ColumnNumber) AndAlso
                                                   Not String.IsNullOrWhiteSpace(m.S_Path)).ToList()
            If applyRows.Count = 0 Then
                Return "적용할 행이 없습니다. (Status=EXACT/MATCH, A_ColumnNumber, S_FullPath 확인)"
            End If

            If selectedSPaths IsNot Nothing AndAlso selectedSPaths.Count > 0 Then
                Dim setPicked As New HashSet(Of String)(selectedSPaths, StringComparer.OrdinalIgnoreCase)
                applyRows = applyRows.Where(Function(m) setPicked.Contains(m.S_Path)).ToList()
            End If

            If applyRows.Count = 0 Then
                Return "선택된 S_FullPath에 해당하는 적용 행이 없습니다."
            End If

            Dim app As Autodesk.Revit.ApplicationServices.Application = uiapp.Application

            Dim total As Integer = 0
            Dim ok As Integer = 0
            Dim fail As Integer = 0
            Dim log As New List(Of Tuple(Of String, Integer, String, String))()

            Dim byDoc = applyRows.GroupBy(Function(m) m.S_Path).ToList()

            For Each g In byDoc
                Dim sPath As String = g.Key
                If String.IsNullOrWhiteSpace(sPath) OrElse Not File.Exists(sPath) Then
                    For Each m In g
                        total += 1 : fail += 1
                        log.Add(Tuple.Create(sPath, m.S_Id, m.A_ColumnNumber, "FAIL:FILE_NOT_FOUND"))
                    Next
                    Continue For
                End If

                Dim targetDoc As RDB.Document = Nothing
                Dim openedByMe As Boolean = False

                Try
                    targetDoc = FindOpenDocumentByPath(app, sPath)
                    If targetDoc Is Nothing Then
                        targetDoc = app.OpenDocumentFile(sPath)
                        openedByMe = True
                    End If

                    Using tr As New RDB.Transaction(targetDoc, "Apply S5_EQCODE From Mapping")
                        tr.Start()

                        For Each m As MapRow In g
                            total += 1

                            Dim elem As RDB.Element = Nothing
                            If Not String.IsNullOrWhiteSpace(m.S_Uid) Then
                                elem = targetDoc.GetElement(m.S_Uid)
                            End If
                            If elem Is Nothing AndAlso m.S_Id <> 0 Then
                                elem = targetDoc.GetElement(New RDB.ElementId(m.S_Id))
                            End If

                            If elem Is Nothing Then
                                fail += 1
                                log.Add(Tuple.Create(sPath, m.S_Id, m.A_ColumnNumber, "FAIL:ELEMENT_NOT_FOUND"))
                                Continue For
                            End If

                            Dim p As RDB.Parameter = elem.LookupParameter(TARGET_PARAM_EQCODE)
                            If p Is Nothing OrElse p.IsReadOnly Then
                                fail += 1
                                log.Add(Tuple.Create(sPath, m.S_Id, m.A_ColumnNumber, "FAIL:PARAM_NOT_FOUND_OR_READONLY"))
                                Continue For
                            End If

                            Dim v As String = m.A_ColumnNumber.Trim()
                            Dim okSet As Boolean = False
                            Try
                                okSet = p.Set(v)
                            Catch
                                okSet = False
                            End Try

                            If ALSO_SET_TARGET_COLNO Then
                                Dim p2 As RDB.Parameter = elem.LookupParameter(TARGET_PARAM_COLNO)
                                If p2 IsNot Nothing AndAlso Not p2.IsReadOnly Then
                                    Try : p2.Set(v) : Catch : End Try
                                End If
                            End If

                            If okSet Then
                                ok += 1
                                log.Add(Tuple.Create(sPath, m.S_Id, v, "OK"))
                            Else
                                fail += 1
                                log.Add(Tuple.Create(sPath, m.S_Id, v, "FAIL:SET_FAILED"))
                            End If
                        Next

                        tr.Commit()
                    End Using

                    If saveAsNew Then
                        Dim dir As String = System.IO.Path.GetDirectoryName(sPath)
                        Dim name As String = System.IO.Path.GetFileNameWithoutExtension(sPath)
                        Dim outRvt As String = System.IO.Path.Combine(dir, $"{name}{suffix}_{DateTime.Now:yyyyMMdd_HHmmss}.rvt")

                        Dim sao As New RDB.SaveAsOptions()
                        sao.OverwriteExistingFile = True
                        targetDoc.SaveAs(outRvt, sao)
                    Else
                        targetDoc.Save()
                    End If

                Catch ex As Exception
                    For Each m In g
                        total += 1 : fail += 1
                        log.Add(Tuple.Create(sPath, m.S_Id, m.A_ColumnNumber, "FAIL:DOC_ERROR:" & ex.Message))
                    Next

                Finally
                    If targetDoc IsNot Nothing AndAlso openedByMe Then
                        Try : targetDoc.Close(False) : Catch : End Try
                    End If
                End Try
            Next

            Dim logPath As String = System.IO.Path.Combine(GetDesktop(), $"APPLY_LOG_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx")
            Try
                WriteApplyLog(logPath, log)
            Catch
                logPath = ""
            End Try

            Dim sb As New StringBuilder()
            sb.AppendLine($"Apply 완료: Total={total}, OK={ok}, FAIL={fail}")
            If Not String.IsNullOrWhiteSpace(logPath) Then sb.AppendLine($"Log: {logPath}")
            Return sb.ToString()
        End Function

        Private Sub WriteApplyLog(path As String, log As List(Of Tuple(Of String, Integer, String, String)))
            Using wb As New XLWorkbook()
                Dim ws = wb.Worksheets.Add("LOG")
                ws.Cell(1, 1).Value = "S_FullPath"
                ws.Cell(1, 2).Value = "S_Id"
                ws.Cell(1, 3).Value = "Value(A_ColumnNumber)"
                ws.Cell(1, 4).Value = "Result"
                ws.Row(1).Style.Font.Bold = True

                Dim r As Integer = 2
                For Each t In log
                    ws.Cell(r, 1).Value = t.Item1
                    ws.Cell(r, 2).Value = t.Item2
                    ws.Cell(r, 3).Value = t.Item3
                    ws.Cell(r, 4).Value = t.Item4
                    r += 1
                Next

                ws.Columns().AdjustToContents()
                wb.SaveAs(path)
            End Using
        End Sub

        Friend Function ReadDistinctSPathsFromMapping(mappingPath As String) As List(Of String)
            Dim setP As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Using wb As New XLWorkbook(mappingPath)
                Dim ws = wb.Worksheet("MAPPING")
                Dim used = ws.RangeUsed()
                If used Is Nothing Then Return New List(Of String)()

                Dim header As IXLRangeRow = used.FirstRowUsed()
                Dim idx = BuildHeaderIndex(header)
                If Not idx.ContainsKey("S_FullPath") Then Return New List(Of String)()

                Dim lastRow As Integer = used.LastRowUsed().RowNumber()
                For r As Integer = header.RowNumber() + 1 To lastRow
                    Dim p As String = ws.Cell(r, idx("S_FullPath")).GetString().Trim()
                    If String.IsNullOrWhiteSpace(p) Then Continue For
                    setP.Add(p)
                Next
            End Using
            Return setP.OrderBy(Function(x) x).ToList()
        End Function

        ' ========================= Excel Read =========================
        Private Function ReadLinksSheet(wb As XLWorkbook) As List(Of LinkInfo)
            Dim ws = wb.Worksheet("LINKS")
            Dim used = ws.RangeUsed()
            Dim list As New List(Of LinkInfo)()
            If used Is Nothing Then Return list

            Dim header As IXLRangeRow = used.FirstRowUsed()
            Dim idx = BuildHeaderIndex(header)
            Dim lastRow As Integer = used.LastRowUsed().RowNumber()

            For r As Integer = header.RowNumber() + 1 To lastRow
                Dim kind As String = GetCellStr(ws, r, idx, "Kind")
                Dim sheetName As String = GetCellStr(ws, r, idx, "SheetName")
                If String.IsNullOrWhiteSpace(kind) OrElse String.IsNullOrWhiteSpace(sheetName) Then Continue For

                list.Add(New LinkInfo With {
                    .Kind = kind.Trim(),
                    .FileName = GetCellStr(ws, r, idx, "FileName").Trim(),
                    .FullPath = GetCellStr(ws, r, idx, "FullPath"),
                    .LinkInstanceId = GetCellInt(ws, r, idx, "LinkInstanceId"),
                    .SheetName = sheetName.Trim()
                })
            Next

            Return list
        End Function

        Private Function ReadUidMap(wb As XLWorkbook) As Dictionary(Of String, String)
            Dim ws = wb.Worksheet("UIDMAP")
            Dim used = ws.RangeUsed()
            Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If used Is Nothing Then Return dict

            Dim header As IXLRangeRow = used.FirstRowUsed()
            Dim idx = BuildHeaderIndex(header)
            Dim lastRow As Integer = used.LastRowUsed().RowNumber()

            For r As Integer = header.RowNumber() + 1 To lastRow
                Dim fileName As String = GetCellStr(ws, r, idx, "FileName")
                Dim linkInstId As Integer = GetCellInt(ws, r, idx, "LinkInstanceId")
                Dim id As Integer = GetCellInt(ws, r, idx, "Id")
                Dim uid As String = GetCellStr(ws, r, idx, "UniqueId")

                If String.IsNullOrWhiteSpace(fileName) OrElse id = 0 OrElse String.IsNullOrWhiteSpace(uid) Then Continue For
                Dim key As String = $"{fileName}|{linkInstId}|{id}"
                If Not dict.ContainsKey(key) Then dict.Add(key, uid)
            Next

            Return dict
        End Function

        Private Function ReadAllRawSheets(wb As XLWorkbook, links As List(Of LinkInfo), uidMap As Dictionary(Of String, String)) As List(Of RawRow)
            Dim rows As New List(Of RawRow)()

            For Each li As LinkInfo In links
                Dim ws As IXLWorksheet
                Try
                    ws = wb.Worksheet(li.SheetName)
                Catch
                    Continue For
                End Try

                Dim used = ws.RangeUsed()
                If used Is Nothing Then Continue For

                Dim header As IXLRangeRow = used.FirstRowUsed()
                Dim idx = BuildHeaderIndex(header)
                Dim lastRow As Integer = used.LastRowUsed().RowNumber()

                For r As Integer = header.RowNumber() + 1 To lastRow
                    Dim id As Integer = GetCellInt(ws, r, idx, "Id")
                    If id = 0 Then Continue For

                    Dim rr As New RawRow()
                    rr.Kind = li.Kind
                    rr.FileName = li.FileName
                    rr.FullPath = li.FullPath
                    rr.LinkInstanceId = li.LinkInstanceId

                    rr.Id = id
                    rr.ElementName = GetCellStr(ws, r, idx, "ElementName")
                    rr.CategoryName = GetCellStr(ws, r, idx, "CategoryName")
                    rr.Xmm = GetCellDbl(ws, r, idx, "LocationX")
                    rr.Ymm = GetCellDbl(ws, r, idx, "LocationY")
                    rr.Zmm = GetCellDbl(ws, r, idx, "LocationZ")

                    rr.Grid = GetCellStr(ws, r, idx, "Grid")
                    rr.StudNumber = GetCellStr(ws, r, idx, "Stud Number")
                    rr.ColumnNumber = GetCellStr(ws, r, idx, "Column Number")

                    Dim key As String = $"{li.FileName}|{li.LinkInstanceId}|{id}"
                    rr.UniqueId = If(uidMap.ContainsKey(key), uidMap(key), "")

                    rows.Add(rr)
                Next
            Next

            Return rows
        End Function

        Private Function ReadMappingWorkbook(path As String) As List(Of MapRow)
            Dim rows As New List(Of MapRow)()

            Using wb As New XLWorkbook(path)
                Dim ws = wb.Worksheet("MAPPING")
                Dim used = ws.RangeUsed()
                If used Is Nothing Then Return rows

                Dim header As IXLRangeRow = used.FirstRowUsed()
                Dim colIndex = BuildHeaderIndex(header)
                Dim lastRow As Integer = used.LastRowUsed().RowNumber()

                For r As Integer = header.RowNumber() + 1 To lastRow
                    Dim sPath As String = GetCellStr(ws, r, colIndex, "S_FullPath")
                    Dim aCol As String = GetCellStr(ws, r, colIndex, "A_ColumnNumber")
                    If String.IsNullOrWhiteSpace(sPath) AndAlso String.IsNullOrWhiteSpace(aCol) Then Continue For

                    Dim mr As New MapRow()
                    mr.PairKey = GetCellStr(ws, r, colIndex, "PairKey")

                    mr.A_File = GetCellStr(ws, r, colIndex, "A_File")
                    mr.A_Path = GetCellStr(ws, r, colIndex, "A_FullPath")
                    mr.A_Id = GetCellInt(ws, r, colIndex, "A_Id")
                    mr.A_Uid = GetCellStr(ws, r, colIndex, "A_UniqueId")
                    mr.A_Xmm = GetCellDbl(ws, r, colIndex, "A_X(mm)")
                    mr.A_Ymm = GetCellDbl(ws, r, colIndex, "A_Y(mm)")
                    mr.A_ColumnNumber = aCol

                    mr.S_File = GetCellStr(ws, r, colIndex, "S_File")
                    mr.S_Path = sPath
                    mr.S_Id = GetCellInt(ws, r, colIndex, "S_Id")
                    mr.S_Uid = GetCellStr(ws, r, colIndex, "S_UniqueId")
                    mr.S_Xmm = GetCellDbl(ws, r, colIndex, "S_X(mm)")
                    mr.S_Ymm = GetCellDbl(ws, r, colIndex, "S_Y(mm)")

                    mr.DistMm = GetCellDbl(ws, r, colIndex, "DistXY(mm)")
                    mr.Status = GetCellStr(ws, r, colIndex, "Status")
                    mr.Note = GetCellStr(ws, r, colIndex, "Note")

                    rows.Add(mr)
                Next
            End Using

            Return rows
        End Function

        ' ========================= Helpers =========================
        Private Function PairKeyOf(fileName As String) As String
            Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(fileName)
            Dim parts As String() = baseName.Split("_"c)

            For i As Integer = 0 To parts.Length - 1
                If parts(i).Equals("A", StringComparison.OrdinalIgnoreCase) OrElse
                   parts(i).Equals("S", StringComparison.OrdinalIgnoreCase) Then
                    parts(i) = "<X>"
                End If
            Next

            Return String.Join("_", parts)
        End Function

        Private Function GetInsertPoint(e As RDB.Element) As RDB.XYZ
            If e Is Nothing Then Return Nothing
            Dim loc As RDB.Location = e.Location
            If loc Is Nothing Then Return Nothing

            Dim lp As RDB.LocationPoint = TryCast(loc, RDB.LocationPoint)
            If lp IsNot Nothing Then Return lp.Point

            Dim lc As RDB.LocationCurve = TryCast(loc, RDB.LocationCurve)
            If lc IsNot Nothing AndAlso lc.Curve IsNot Nothing Then
                Return lc.Curve.Evaluate(0.5, True)
            End If

            Return Nothing
        End Function

        Private Function Dist2Dmm(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
            Dim dx As Double = x1 - x2
            Dim dy As Double = y1 - y2
            Return Math.Sqrt(dx * dx + dy * dy)
        End Function

        Private Function CellKeyOf(xmm As Double, ymm As Double, cellMm As Double) As CellKey
            Dim ix As Integer = CInt(Math.Floor(xmm / cellMm))
            Dim iy As Integer = CInt(Math.Floor(ymm / cellMm))
            Return New CellKey(ix, iy)
        End Function

        Private Function GetElementTypeDisplayName(doc As RDB.Document, e As RDB.Element) As String
            If e Is Nothing Then Return ""
            Try
                Dim tid As RDB.ElementId = e.GetTypeId()
                If tid IsNot Nothing AndAlso tid.IntegerValue > 0 Then
                    Dim t As RDB.Element = doc.GetElement(tid)
                    Dim fs As RDB.FamilySymbol = TryCast(t, RDB.FamilySymbol)
                    If fs IsNot Nothing Then
                        Dim famName As String = ""
                        Try : famName = fs.Family.Name : Catch : famName = "" : End Try
                        Return famName & ":" & fs.Name
                    End If
                    If t IsNot Nothing Then Return t.Name
                End If
            Catch
            End Try
            Try : Return e.Name : Catch : End Try
            Return ""
        End Function

        Private Function GetParamString(doc As RDB.Document, e As RDB.Element, paramName As String) As String
            If e Is Nothing OrElse String.IsNullOrWhiteSpace(paramName) Then Return ""

            Dim p As RDB.Parameter = e.LookupParameter(paramName)

            If p Is Nothing Then
                Try
                    Dim t As RDB.Element = doc.GetElement(e.GetTypeId())
                    If t IsNot Nothing Then p = t.LookupParameter(paramName)
                Catch
                End Try
            End If

            If p Is Nothing Then Return ""

            Try
                If p.StorageType = RDB.StorageType.String Then Return If(p.AsString(), "")
                Dim s As String = p.AsValueString()
                If s IsNot Nothing Then Return s
            Catch
            End Try

            Try
                Select Case p.StorageType
                    Case RDB.StorageType.Integer
                        Return p.AsInteger().ToString()
                    Case RDB.StorageType.Double
                        Return p.AsDouble().ToString()
                    Case RDB.StorageType.ElementId
                        Return p.AsElementId().IntegerValue.ToString()
                    Case Else
                        Return ""
                End Select
            Catch
                Return ""
            End Try
        End Function

        Private Function BuildHeaderIndex(headerRow As IXLRangeRow) As Dictionary(Of String, Integer)
            Dim dict As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For Each cell In headerRow.CellsUsed()
                Dim name As String = (If(cell.GetString(), "")).Trim()
                If String.IsNullOrWhiteSpace(name) Then Continue For
                If Not dict.ContainsKey(name) Then dict.Add(name, cell.Address.ColumnNumber)
            Next
            Return dict
        End Function

        Private Function GetCellStr(ws As IXLWorksheet, r As Integer, idx As Dictionary(Of String, Integer), name As String) As String
            If Not idx.ContainsKey(name) Then Return ""
            Return ws.Cell(r, idx(name)).GetString().Trim()
        End Function

        Private Function GetCellInt(ws As IXLWorksheet, r As Integer, idx As Dictionary(Of String, Integer), name As String) As Integer
            If Not idx.ContainsKey(name) Then Return 0
            Dim c = ws.Cell(r, idx(name))
            If c.IsEmpty() Then Return 0
            Dim s As String = c.GetString().Trim()
            Dim v As Integer
            If Integer.TryParse(s, v) Then Return v
            Dim d As Double
            If Double.TryParse(s, d) Then Return CInt(Math.Truncate(d))
            Return 0
        End Function

        Private Function GetCellDbl(ws As IXLWorksheet, r As Integer, idx As Dictionary(Of String, Integer), name As String) As Double
            If Not idx.ContainsKey(name) Then Return 0
            Dim c = ws.Cell(r, idx(name))
            If c.IsEmpty() Then Return 0
            Dim s As String = c.GetString().Trim()
            Dim d As Double
            If Double.TryParse(s, d) Then Return d
            Return 0
        End Function

        Private Function SafeSheetBase(fileName As String) As String
            Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(fileName)
            Dim s As String = Regex.Replace(baseName, "[:\\/?*\[\]]", "_")
            If s.Length > 25 Then s = s.Substring(0, 25)
            Return s
        End Function

        Private Function MakeUniqueSheetName(wb As XLWorkbook, used As HashSet(Of String), proposed As String) As String
            Dim name As String = proposed
            If name.Length > 31 Then name = name.Substring(0, 31)

            Dim i As Integer = 1
            While used.Contains(name) OrElse wb.Worksheets.Any(Function(x) x.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                Dim suffix As String = "_" & i.ToString()
                Dim baseName As String = proposed
                If baseName.Length + suffix.Length > 31 Then
                    baseName = baseName.Substring(0, 31 - suffix.Length)
                End If
                name = baseName & suffix
                i += 1
            End While

            used.Add(name)
            Return name
        End Function

        Private Function FindOpenDocumentByPath(app As Autodesk.Revit.ApplicationServices.Application, fullPath As String) As RDB.Document
            For Each d As RDB.Document In app.Documents
                Try
                    If String.Equals(d.PathName, fullPath, StringComparison.OrdinalIgnoreCase) Then
                        Return d
                    End If
                Catch
                End Try
            Next
            Return Nothing
        End Function

    End Module

End Namespace
