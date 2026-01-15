Option Explicit On
Option Strict On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions

Imports WinForms = System.Windows.Forms

Imports RvtAttr = Autodesk.Revit.Attributes
Imports RvtApp = Autodesk.Revit.ApplicationServices
Imports RvtDB = Autodesk.Revit.DB
Imports RvtPlumb = Autodesk.Revit.DB.Plumbing
Imports RvtUI = Autodesk.Revit.UI

Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

'=========================================================
' DTO / Options
'=========================================================
Public Class BatchOptions
    Public Enum CheckMode
        Segment = 0
        RoutingPreference = 1
    End Enum

    Public Enum RoutingCategory
        Pipe = 0
        Duct = 1
        Conduit = 2
        CableTray = 3
    End Enum

    Public Enum UnitType
        Millimeter = 0
        Inch = 1
    End Enum

    Public Property RvtFiles As List(Of String)
    Public Property Mode As CheckMode
    Public Property RoutingCat As Nullable(Of RoutingCategory)
    Public Property Unit As UnitType
    Public Property OutputPath As String
    Public Property ClassMaterialCheck As Boolean

    Public Sub New()
        RvtFiles = New List(Of String)()
        Mode = CheckMode.Segment
        RoutingCat = Nothing
        Unit = UnitType.Millimeter
        OutputPath = ""
        ClassMaterialCheck = False
    End Sub
End Class

Public Class SegmentSizeEntry
    Public Property FileName As String
    Public Property SegmentName As String
    Public Property Nd As Double
    Public Property Id As Double
    Public Property Od As Double
End Class

Public Class RoutingRuleEntry
    Public Property FileName As String
    Public Property CategoryName As String
    Public Property TypeName As String
    Public Property GroupType As RvtDB.RoutingPreferenceRuleGroupType
    Public Property RuleIndex As Integer
    Public Property PartFamily As String
    Public Property PartType As String
    Public Property MinSize As Double
    Public Property MaxSize As Double
End Class

Public Class ClassMaterialResult
    Public Property FileName As String
    Public Property PipeTypeName As String
    Public Property SubjectKind As String
    Public Property FittingPart As String
    Public Property SubjectName As String
    Public Property SubjectClass As String
    Public Property SegmentKey As String
    Public Property SegmentClass As String
    Public Property Status As String
End Class

Public Class ErrorEntry
    Public Property FilePath As String
    Public Property Stage As String
    Public Property Message As String
End Class

'=========================================================
' UI - Options Form (WinForms)
'=========================================================
Public Class BatchOptionsForm
    Inherits WinForms.Form

    Private ReadOnly btnAddFolder As WinForms.Button
    Private ReadOnly lstFolderFiles As WinForms.ListView
    Private ReadOnly btnAddChecked As WinForms.Button
    Private ReadOnly lstSelectedFiles As WinForms.ListBox
    Private ReadOnly btnRemoveSelected As WinForms.Button

    Private ReadOnly grpMode As WinForms.GroupBox
    Private ReadOnly rbSegment As WinForms.RadioButton
    Private ReadOnly rbRouting As WinForms.RadioButton

    Private ReadOnly grpUnit As WinForms.GroupBox
    Private ReadOnly rbMm As WinForms.RadioButton
    Private ReadOnly rbInch As WinForms.RadioButton

    Private ReadOnly grpRoutingCategory As WinForms.GroupBox
    Private ReadOnly cmbRoutingCategory As WinForms.ComboBox

    Private ReadOnly lblOutput As WinForms.Label
    Private ReadOnly txtOutputPath As WinForms.TextBox
    Private ReadOnly btnBrowseOutput As WinForms.Button
    Private ReadOnly chkClassMaterial As WinForms.CheckBox

    Private ReadOnly btnRun As WinForms.Button
    Private ReadOnly btnCancel As WinForms.Button

    Public Property ResultOptions As BatchOptions

    Public Sub New()
        Me.Text = "Batch Segment / Routing Preference Check"
        Me.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = WinForms.FormStartPosition.CenterScreen
        Me.Width = 920
        Me.Height = 660

        ResultOptions = Nothing

        btnAddFolder = New WinForms.Button() With {.Text = "경로 추가...", .Left = 10, .Top = 10, .Width = 120}
        AddHandler btnAddFolder.Click, AddressOf OnAddFolder

        lstFolderFiles = New WinForms.ListView() With {
            .Left = 10, .Top = 40, .Width = 560, .Height = 260,
            .CheckBoxes = True, .View = WinForms.View.Details, .FullRowSelect = True
        }
        lstFolderFiles.Columns.Add("FileName", 200)
        lstFolderFiles.Columns.Add("FullPath", 340)

        btnAddChecked = New WinForms.Button() With {.Text = "체크된 파일 추가 →", .Left = 580, .Top = 120, .Width = 150}
        AddHandler btnAddChecked.Click, AddressOf OnAddChecked

        lstSelectedFiles = New WinForms.ListBox() With {.Left = 740, .Top = 40, .Width = 160, .Height = 260}
        btnRemoveSelected = New WinForms.Button() With {.Text = "선택 제거", .Left = 740, .Top = 310, .Width = 160}
        AddHandler btnRemoveSelected.Click, AddressOf OnRemoveSelected

        grpMode = New WinForms.GroupBox() With {.Text = "검토 모드", .Left = 10, .Top = 315, .Width = 260, .Height = 95}
        rbSegment = New WinForms.RadioButton() With {.Text = "Segment 검토", .Left = 10, .Top = 28, .Width = 220, .Checked = True}
        rbRouting = New WinForms.RadioButton() With {.Text = "Routing Preference 검토", .Left = 10, .Top = 55, .Width = 220}
        AddHandler rbSegment.CheckedChanged, AddressOf OnModeChanged
        AddHandler rbRouting.CheckedChanged, AddressOf OnModeChanged
        grpMode.Controls.Add(rbSegment)
        grpMode.Controls.Add(rbRouting)

        grpUnit = New WinForms.GroupBox() With {.Text = "단위", .Left = 280, .Top = 315, .Width = 200, .Height = 95}
        rbMm = New WinForms.RadioButton() With {.Text = "mm", .Left = 10, .Top = 28, .Width = 100, .Checked = True}
        rbInch = New WinForms.RadioButton() With {.Text = "inch", .Left = 10, .Top = 55, .Width = 100}
        grpUnit.Controls.Add(rbMm)
        grpUnit.Controls.Add(rbInch)

        grpRoutingCategory = New WinForms.GroupBox() With {.Text = "Routing Category (Routing 모드에서만)", .Left = 490, .Top = 315, .Width = 410, .Height = 95}
        cmbRoutingCategory = New WinForms.ComboBox() With {.Left = 10, .Top = 35, .Width = 260, .DropDownStyle = WinForms.ComboBoxStyle.DropDownList}
        cmbRoutingCategory.Items.AddRange(New Object() {"Pipe", "Duct", "Conduit", "CableTray"})
        grpRoutingCategory.Controls.Add(cmbRoutingCategory)
        grpRoutingCategory.Enabled = False

        chkClassMaterial = New WinForms.CheckBox() With {.Text = "Class 재질 매칭 검토", .Left = 490, .Top = 415, .Width = 220, .Checked = False, .Enabled = False}

        lblOutput = New WinForms.Label() With {.Text = "출력 Excel 경로:", .Left = 10, .Top = 430, .Width = 180}
        txtOutputPath = New WinForms.TextBox() With {.Left = 10, .Top = 455, .Width = 720}
        btnBrowseOutput = New WinForms.Button() With {.Text = "찾아보기...", .Left = 740, .Top = 452, .Width = 160}
        AddHandler btnBrowseOutput.Click, AddressOf OnBrowseOutput

        btnRun = New WinForms.Button() With {.Text = "시작", .Left = 580, .Top = 520, .Width = 150}
        btnCancel = New WinForms.Button() With {.Text = "취소", .Left = 750, .Top = 520, .Width = 150}
        AddHandler btnRun.Click, AddressOf OnRun
        AddHandler btnCancel.Click, Sub(sender, e) Me.DialogResult = WinForms.DialogResult.Cancel

        Me.Controls.Add(btnAddFolder)
        Me.Controls.Add(lstFolderFiles)
        Me.Controls.Add(btnAddChecked)
        Me.Controls.Add(lstSelectedFiles)
        Me.Controls.Add(btnRemoveSelected)
        Me.Controls.Add(grpMode)
        Me.Controls.Add(grpUnit)
        Me.Controls.Add(grpRoutingCategory)
        Me.Controls.Add(chkClassMaterial)
        Me.Controls.Add(lblOutput)
        Me.Controls.Add(txtOutputPath)
        Me.Controls.Add(btnBrowseOutput)
        Me.Controls.Add(btnRun)
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub OnModeChanged(sender As Object, e As EventArgs)
        grpRoutingCategory.Enabled = rbRouting.Checked
        chkClassMaterial.Enabled = rbRouting.Checked
        If Not rbRouting.Checked Then
            chkClassMaterial.Checked = False
        End If
    End Sub

    Private Sub OnAddFolder(sender As Object, e As EventArgs)
        Using dlg As New WinForms.FolderBrowserDialog()
            dlg.Description = "RVT 파일이 있는 폴더를 선택하세요 (하위 폴더 포함 검색)."
            If dlg.ShowDialog(Me) <> WinForms.DialogResult.OK Then Return

            Dim folder As String = dlg.SelectedPath
            Dim files As String() = Array.Empty(Of String)()
            Try
                files = Directory.GetFiles(folder, "*.rvt", SearchOption.AllDirectories)
            Catch ex As Exception
                WinForms.MessageBox.Show(Me, "폴더 검색 실패: " & ex.Message)
                Return
            End Try

            For Each f As String In files
                Dim exists As Boolean = lstFolderFiles.Items.Cast(Of WinForms.ListViewItem)().
                    Any(Function(it) String.Equals(it.SubItems(1).Text, f, StringComparison.OrdinalIgnoreCase))
                If exists Then Continue For

                Dim item As New WinForms.ListViewItem(Path.GetFileName(f))
                item.SubItems.Add(f)
                item.Checked = True
                lstFolderFiles.Items.Add(item)
            Next
        End Using
    End Sub

    Private Sub OnAddChecked(sender As Object, e As EventArgs)
        For Each item As WinForms.ListViewItem In lstFolderFiles.Items
            If Not item.Checked Then Continue For

            Dim fullPath As String = item.SubItems(1).Text
            Dim exists As Boolean = lstSelectedFiles.Items.Cast(Of String)().
                Any(Function(s) String.Equals(s, fullPath, StringComparison.OrdinalIgnoreCase))
            If Not exists Then lstSelectedFiles.Items.Add(fullPath)
        Next
    End Sub

    Private Sub OnRemoveSelected(sender As Object, e As EventArgs)
        Dim removeList As New List(Of Object)()
        For Each sel As Object In lstSelectedFiles.SelectedItems
            removeList.Add(sel)
        Next
        For Each x As Object In removeList
            lstSelectedFiles.Items.Remove(x)
        Next
    End Sub

    Private Sub OnBrowseOutput(sender As Object, e As EventArgs)
        Using dlg As New WinForms.SaveFileDialog()
            dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
            dlg.FileName = If(rbSegment.Checked, "SegmentCheck.xlsx", "RoutingCheck.xlsx")
            If dlg.ShowDialog(Me) = WinForms.DialogResult.OK Then
                txtOutputPath.Text = dlg.FileName
            End If
        End Using
    End Sub

    Private Sub OnRun(sender As Object, e As EventArgs)
        If lstSelectedFiles.Items.Count = 0 Then
            WinForms.MessageBox.Show(Me, "선택된 RVT 파일이 없습니다.", "경고", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)
            Return
        End If

        Dim outPath As String = txtOutputPath.Text.Trim()
        If String.IsNullOrWhiteSpace(outPath) OrElse Not outPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) Then
            WinForms.MessageBox.Show(Me, "출력 파일 경로(.xlsx)를 지정해 주세요.", "경고", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)
            Return
        End If

        Dim opt As New BatchOptions()
        opt.RvtFiles = lstSelectedFiles.Items.Cast(Of String)().Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        opt.Mode = If(rbSegment.Checked, BatchOptions.CheckMode.Segment, BatchOptions.CheckMode.RoutingPreference)
        opt.Unit = If(rbMm.Checked, BatchOptions.UnitType.Millimeter, BatchOptions.UnitType.Inch)
        opt.OutputPath = outPath
        opt.ClassMaterialCheck = chkClassMaterial.Checked

        If opt.Mode = BatchOptions.CheckMode.RoutingPreference Then
            If cmbRoutingCategory.SelectedIndex < 0 Then
                WinForms.MessageBox.Show(Me, "Routing 모드에서는 카테고리를 선택해야 합니다.", "경고", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)
                Return
            End If

            Select Case cmbRoutingCategory.SelectedItem.ToString()
                Case "Pipe" : opt.RoutingCat = BatchOptions.RoutingCategory.Pipe
                Case "Duct" : opt.RoutingCat = BatchOptions.RoutingCategory.Duct
                Case "Conduit" : opt.RoutingCat = BatchOptions.RoutingCategory.Conduit
                Case "CableTray" : opt.RoutingCat = BatchOptions.RoutingCategory.CableTray
                Case Else : opt.RoutingCat = BatchOptions.RoutingCategory.Pipe
            End Select
        Else
            opt.RoutingCat = Nothing
        End If

        ResultOptions = opt
        Me.DialogResult = WinForms.DialogResult.OK
        Me.Close()
    End Sub
End Class

'=========================================================
' UI - Progress Form
'=========================================================
Public Class ProgressForm
    Inherits WinForms.Form

    Private ReadOnly lblInfo As WinForms.Label

    Public Sub New()
        Me.Text = "Processing..."
        Me.Width = 420
        Me.Height = 140
        Me.FormBorderStyle = WinForms.FormBorderStyle.FixedToolWindow
        Me.StartPosition = WinForms.FormStartPosition.CenterScreen

        lblInfo = New WinForms.Label() With {.Left = 10, .Top = 10, .Width = 380, .Height = 80}
        Me.Controls.Add(lblInfo)
    End Sub

    Public Sub UpdateProgress(currentIndex As Integer, total As Integer, filePath As String)
        lblInfo.Text = $"Processing {currentIndex} / {total}{Environment.NewLine}{Path.GetFileName(filePath)}"
        lblInfo.Refresh()
        WinForms.Application.DoEvents()
    End Sub
End Class

'=========================================================
' Command
'=========================================================
<RvtAttr.Transaction(RvtAttr.TransactionMode.Manual)>
Public Class CmdBatchSegmentAndRouting
    Implements RvtUI.IExternalCommand

    Public Function Execute(commandData As RvtUI.ExternalCommandData,
                            ByRef message As String,
                            elements As RvtDB.ElementSet) As RvtUI.Result Implements RvtUI.IExternalCommand.Execute

        Try
            WinForms.Application.EnableVisualStyles()

            Dim uiApp As RvtUI.UIApplication = commandData.Application
            Dim app As RvtApp.Application = uiApp.Application

            Dim opt As BatchOptions = Nothing
            Using f As New BatchOptionsForm()
                If f.ShowDialog() <> WinForms.DialogResult.OK Then
                    Return RvtUI.Result.Cancelled
                End If
                opt = f.ResultOptions
            End Using

            If opt Is Nothing OrElse opt.RvtFiles Is Nothing OrElse opt.RvtFiles.Count = 0 Then
                Return RvtUI.Result.Cancelled
            End If

            Dim unitFactor As Double = If(opt.Unit = BatchOptions.UnitType.Millimeter, 304.8R, 12.0R)
            Dim unitLabel As String = If(opt.Unit = BatchOptions.UnitType.Millimeter, "mm", "in")

            Dim segAll As New List(Of SegmentSizeEntry)()
            Dim routingAll As New List(Of RoutingRuleEntry)()
            Dim errors As New List(Of ErrorEntry)()
            Dim classMatches As List(Of ClassMaterialResult) = Nothing

            Dim prog As New ProgressForm()
            prog.Show()

            Dim openedByMe As New List(Of RvtDB.Document)()

            Try
                Dim total As Integer = opt.RvtFiles.Count
                For i As Integer = 0 To total - 1
                    Dim filePath As String = opt.RvtFiles(i)
                    prog.UpdateProgress(i + 1, total, filePath)

                    If Not File.Exists(filePath) Then
                        errors.Add(New ErrorEntry With {.FilePath = filePath, .Stage = "FileMissing", .Message = "파일이 존재하지 않습니다."})
                        Continue For
                    End If

                    Dim shortName As String = System.IO.Path.GetFileName(filePath)
                    Dim doc As RvtDB.Document = FindOpenDocumentByPath(app, filePath)
                    Dim mustClose As Boolean = False

                    Try
                        If doc Is Nothing Then
                            doc = OpenDetachedWithClosedWorksets(app, filePath)
                            mustClose = True
                            openedByMe.Add(doc)
                        End If

                        If opt.Mode = BatchOptions.CheckMode.Segment Then
                            segAll.AddRange(CollectPipeSegments(doc, shortName, unitFactor))
                        Else
                            If Not opt.RoutingCat.HasValue Then
                                errors.Add(New ErrorEntry With {.FilePath = filePath, .Stage = "Input", .Message = "RoutingCat이 선택되지 않았습니다."})
                            Else
                                routingAll.AddRange(CollectRoutingPrefs(doc, shortName, opt.RoutingCat.Value, unitFactor))
                            End If
                        End If

                    Catch ex As Exception
                        errors.Add(New ErrorEntry With {
                            .FilePath = filePath,
                            .Stage = If(opt.Mode = BatchOptions.CheckMode.Segment, "CollectSegment", "CollectRouting"),
                            .Message = ex.Message
                        })
                    Finally
                        If mustClose AndAlso doc IsNot Nothing Then
                            Try
                                doc.Close(False)
                            Catch
                                ' ignore
                            End Try
                        End If
                    End Try
                Next
            Finally
                prog.Close()
            End Try

            ' 출력 폴더 보장
            Try
                Dim outDir As String = Path.GetDirectoryName(opt.OutputPath)
                If Not String.IsNullOrWhiteSpace(outDir) AndAlso Not Directory.Exists(outDir) Then
                    Directory.CreateDirectory(outDir)
                End If
            Catch
            End Try

            If opt.Mode = BatchOptions.CheckMode.Segment Then
                ExportSegmentResultToExcel(opt.OutputPath, segAll, errors, unitLabel, opt.RvtFiles)
            Else
                If opt.ClassMaterialCheck Then
                    classMatches = BuildClassMaterialResults(routingAll)
                End If
                ExportRoutingResultToExcel(opt.OutputPath, routingAll, errors, unitLabel, opt.RvtFiles, classMatches)
            End If

            Dim failFiles As Integer = errors.Select(Function(x) x.FilePath).Distinct(StringComparer.OrdinalIgnoreCase).Count()
            Dim successFiles As Integer = Math.Max(0, opt.RvtFiles.Count - failFiles)

            RvtUI.TaskDialog.Show("Batch Segment / Routing Check",
                                  $"완료되었습니다.{Environment.NewLine}" &
                                  $"총 파일: {opt.RvtFiles.Count}{Environment.NewLine}" &
                                  $"성공: {successFiles}{Environment.NewLine}" &
                                  $"실패(오류시트): {failFiles}{Environment.NewLine}" &
                                  $"결과: {opt.OutputPath}")

            Return RvtUI.Result.Succeeded

        Catch exAll As Exception
            message = exAll.ToString()
            Return RvtUI.Result.Failed
        End Try
    End Function

    '=========================================================
    ' Open: Detached + CloseAllWorksets (Revit 2019 안전버전)
    '=========================================================
    Private Function OpenDetachedWithClosedWorksets(app As RvtApp.Application, filePath As String) As RvtDB.Document
        Dim mp As RvtDB.ModelPath = RvtDB.ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath)

        Dim openOpts As New RvtDB.OpenOptions()

        Dim wsCfg As New RvtDB.WorksetConfiguration(RvtDB.WorksetConfigurationOption.CloseAllWorksets)
        openOpts.SetOpenWorksetsConfiguration(wsCfg)

        ' Revit 2019: WorksharingUtils.GetWorksharingModelStatus 같은 걸 쓰지 말고 BasicFileInfo로 판단
        Try
            Dim bfi As RvtDB.BasicFileInfo = RvtDB.BasicFileInfo.Extract(filePath)
            If bfi IsNot Nothing AndAlso bfi.IsWorkshared AndAlso bfi.IsCentral Then
                openOpts.DetachFromCentralOption = RvtDB.DetachFromCentralOption.DetachAndDiscardWorksets
            End If
        Catch
            ' 판단 실패해도 그냥 오픈 시도
        End Try

        Return app.OpenDocumentFile(mp, openOpts)
    End Function

    Private Function FindOpenDocumentByPath(app As RvtApp.Application, filePath As String) As RvtDB.Document
        For Each d As RvtDB.Document In app.Documents
            If String.Equals(d.PathName, filePath, StringComparison.OrdinalIgnoreCase) Then
                Return d
            End If
        Next
        Return Nothing
    End Function

    '=========================================================
    ' Segment Collect
    '=========================================================
    Private Function CollectPipeSegments(doc As RvtDB.Document, fileLabel As String, unitFactor As Double) As List(Of SegmentSizeEntry)
        Dim result As New List(Of SegmentSizeEntry)()

        Dim col As New RvtDB.FilteredElementCollector(doc)
        col.OfClass(GetType(RvtPlumb.PipeSegment))

        For Each e As RvtDB.Element In col
            Dim seg As RvtPlumb.PipeSegment = TryCast(e, RvtPlumb.PipeSegment)
            If seg Is Nothing Then Continue For

            Dim segName As String = seg.Name

            For Each s As RvtDB.MEPSize In seg.GetSizes()
                result.Add(New SegmentSizeEntry With {
                    .FileName = fileLabel,
                    .SegmentName = segName,
                    .Nd = s.NominalDiameter * unitFactor,
                    .Id = s.InnerDiameter * unitFactor,
                    .Od = s.OuterDiameter * unitFactor
                })
            Next
        Next

        Return result
    End Function

    '=========================================================
    ' Routing Collect (타입들은 문자열로 찾아서 접근: ConduitType/CableTrayType Friend 문제 회피)
    '=========================================================
    Private Function CollectRoutingPrefs(doc As RvtDB.Document,
                                         fileLabel As String,
                                         cat As BatchOptions.RoutingCategory,
                                         unitFactor As Double) As List(Of RoutingRuleEntry)

        Dim result As New List(Of RoutingRuleEntry)()
        Dim asm As Reflection.Assembly = GetType(RvtDB.Element).Assembly

        Dim typeFullName As String
        Dim categoryName As String

        Select Case cat
            Case BatchOptions.RoutingCategory.Pipe
                typeFullName = "Autodesk.Revit.DB.Plumbing.PipeType"
                categoryName = "Pipe"
            Case BatchOptions.RoutingCategory.Duct
                typeFullName = "Autodesk.Revit.DB.Mechanical.DuctType"
                categoryName = "Duct"
            Case BatchOptions.RoutingCategory.Conduit
                typeFullName = "Autodesk.Revit.DB.Electrical.ConduitType"
                categoryName = "Conduit"
            Case BatchOptions.RoutingCategory.CableTray
                typeFullName = "Autodesk.Revit.DB.Electrical.CableTrayType"
                categoryName = "CableTray"
            Case Else
                typeFullName = "Autodesk.Revit.DB.Plumbing.PipeType"
                categoryName = "Pipe"
        End Select

        Dim targetType As Type = asm.GetType(typeFullName, throwOnError:=False, ignoreCase:=False)
        If targetType Is Nothing Then
            Throw New InvalidOperationException("RevitAPI에서 타입을 찾을 수 없습니다: " & typeFullName)
        End If

        Dim groups As RvtDB.RoutingPreferenceRuleGroupType() = {
            RvtDB.RoutingPreferenceRuleGroupType.Segments,
            RvtDB.RoutingPreferenceRuleGroupType.Elbows,
            RvtDB.RoutingPreferenceRuleGroupType.Junctions,
            RvtDB.RoutingPreferenceRuleGroupType.Crosses,
            RvtDB.RoutingPreferenceRuleGroupType.Transitions
        }

        Dim typesCol As New RvtDB.FilteredElementCollector(doc)
        typesCol.OfClass(targetType)

        For Each el As RvtDB.Element In typesCol
            Dim et As RvtDB.ElementType = TryCast(el, RvtDB.ElementType)
            If et Is Nothing Then Continue For

            Dim rpmProp As Reflection.PropertyInfo = et.GetType().GetProperty("RoutingPreferenceManager")
            If rpmProp Is Nothing Then Continue For

            Dim rpmObj As Object = rpmProp.GetValue(et, Nothing)
            Dim rpm As RvtDB.RoutingPreferenceManager = TryCast(rpmObj, RvtDB.RoutingPreferenceManager)
            If rpm Is Nothing Then Continue For

            Dim typeName As String = et.Name

            For Each g As RvtDB.RoutingPreferenceRuleGroupType In groups
                Dim n As Integer = 0
                Try
                    n = rpm.GetNumberOfRules(g)
                Catch
                    n = 0
                End Try

                For i As Integer = 0 To n - 1
                    Dim rule As RvtDB.RoutingPreferenceRule = Nothing
                    Try
                        rule = rpm.GetRule(g, i)
                    Catch
                        rule = Nothing
                    End Try
                    If rule Is Nothing Then Continue For

                    Dim partFamily As String = ""
                    Dim partType As String = ""
                    Try
                        Dim pid As RvtDB.ElementId = rule.MEPPartId
                        If pid IsNot Nothing AndAlso pid.IntegerValue <> -1 Then
                            Dim part As RvtDB.Element = doc.GetElement(pid)
                            If part IsNot Nothing Then
                                If TypeOf part Is RvtDB.FamilySymbol Then
                                    Dim sym As RvtDB.FamilySymbol = DirectCast(part, RvtDB.FamilySymbol)
                                    partType = sym.Name
                                    If sym.Family IsNot Nothing Then partFamily = sym.Family.Name
                                ElseIf TypeOf part Is RvtDB.ElementType Then
                                    Dim pet As RvtDB.ElementType = DirectCast(part, RvtDB.ElementType)
                                    partType = pet.Name
                                    Dim fnProp As Reflection.PropertyInfo = pet.GetType().GetProperty("FamilyName")
                                    If fnProp IsNot Nothing Then
                                        Dim v As Object = fnProp.GetValue(pet, Nothing)
                                        If v IsNot Nothing Then partFamily = v.ToString()
                                    End If
                                Else
                                    partType = part.Name
                                End If
                            End If
                        End If
                    Catch
                    End Try

                    Dim minFeet As Double = 0.0R
                    Dim maxFeet As Double = 0.0R

                    Try
                        Dim critCount As Integer = rule.NumberOfCriteria
                        For ci As Integer = 0 To critCount - 1
                            Dim crit As Object = rule.GetCriterion(ci)
                            If crit Is Nothing Then Continue For

                            Dim ct As Type = crit.GetType()
                            Dim pMin As Reflection.PropertyInfo = ct.GetProperty("MinimumSize")
                            Dim pMax As Reflection.PropertyInfo = ct.GetProperty("MaximumSize")

                            If pMin IsNot Nothing AndAlso pMax IsNot Nothing Then
                                Dim vMin As Object = pMin.GetValue(crit, Nothing)
                                Dim vMax As Object = pMax.GetValue(crit, Nothing)
                                If vMin IsNot Nothing Then minFeet = CDbl(vMin)
                                If vMax IsNot Nothing Then maxFeet = CDbl(vMax)
                                Exit For
                            End If
                        Next
                    Catch
                        ' 사이즈 기준 없는 룰도 있음
                    End Try

                    result.Add(New RoutingRuleEntry With {
                        .FileName = fileLabel,
                        .CategoryName = categoryName,
                        .TypeName = typeName,
                        .GroupType = g,
                        .RuleIndex = i,
                        .PartFamily = partFamily,
                        .PartType = partType,
                        .MinSize = minFeet * unitFactor,
                        .MaxSize = maxFeet * unitFactor
                    })
                Next
            Next
        Next

        Return result
    End Function

    '=========================================================
    ' Excel Helpers
    '=========================================================
    Private Shared Function Fmt3(v As Double) As String
        Return v.ToString("0.###", CultureInfo.InvariantCulture)
    End Function

    Private Shared Sub SetCell(row As IRow, col As Integer, value As String)
        Dim cell As ICell = row.GetCell(col)
        If cell Is Nothing Then cell = row.CreateCell(col)
        cell.SetCellValue(value)
    End Sub

    Private Shared Sub SetCell(row As IRow, col As Integer, value As Double)
        Dim cell As ICell = row.GetCell(col)
        If cell Is Nothing Then cell = row.CreateCell(col)
        cell.SetCellValue(value)
    End Sub

    '=========================================================
    ' Class Material Helpers
    '=========================================================
    Private Shared Function NormalizeClassToken(raw As String) As String
        If String.IsNullOrWhiteSpace(raw) Then
            Return String.Empty
        End If
        Dim trimmed As String = raw.Trim()
        Dim normalized As String = Regex.Replace(trimmed, "\s+", " ")
        Return normalized.Trim().ToLowerInvariant()
    End Function

    Private Shared Function ExtractSegmentClass(seg As String) As String
        If String.IsNullOrWhiteSpace(seg) Then
            Return String.Empty
        End If
        Dim token As String = seg.Trim()
        Dim dashIdx As Integer = token.IndexOf(" - ", StringComparison.Ordinal)
        If dashIdx >= 0 Then
            token = token.Substring(0, dashIdx)
        Else
            Dim simpleDash As Integer = token.IndexOf("-"c)
            If simpleDash > 0 Then
                token = token.Substring(0, simpleDash)
            End If
        End If
        token = token.Trim()

        If token.EndsWith("Ref.", StringComparison.OrdinalIgnoreCase) Then
            token = token.Substring(0, token.Length - 4)
        ElseIf token.EndsWith("REF", StringComparison.OrdinalIgnoreCase) Then
            token = token.Substring(0, token.Length - 3)
        ElseIf token.EndsWith(".ref", StringComparison.OrdinalIgnoreCase) Then
            token = token.Substring(0, token.Length - 4)
        End If

        token = token.Trim()
        Return token
    End Function

    Private Shared Function ExtractSubjectClass(src As String) As String
        If String.IsNullOrWhiteSpace(src) Then
            Return String.Empty
        End If

        Dim trimmed As String = src.Trim()
        Dim parts As String() = trimmed.Split(","c)
        Dim candidate As String = String.Empty
        If parts.Length >= 2 Then
            candidate = parts(1).Trim()
        End If

        If String.IsNullOrWhiteSpace(candidate) Then
            Dim m As Match = Regex.Match(trimmed, "\b[^,]*\([^()]+\)")
            If m.Success Then
                candidate = m.Value.Trim()
            End If
        End If

        Return candidate
    End Function

    Private Shared Function BuildClassMaterialResults(entries As List(Of RoutingRuleEntry)) As List(Of ClassMaterialResult)
        Dim res As New List(Of ClassMaterialResult)()
        If entries Is Nothing OrElse entries.Count = 0 Then
            Return res
        End If

        Dim groups = entries.GroupBy(Function(x) x.FileName & "|" & x.TypeName, StringComparer.OrdinalIgnoreCase)
        For Each grp In groups
            Dim keyParts As String() = grp.Key.Split("|"c)
            Dim fileName As String = keyParts(0)
            Dim typeName As String = keyParts(1)

            Dim segmentNames As List(Of String) = grp.
                Where(Function(x) x.GroupType = RvtDB.RoutingPreferenceRuleGroupType.Segments).
                Select(Function(x) If(String.IsNullOrWhiteSpace(x.PartType), x.PartFamily, x.PartType)).
                Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()

            Dim subjects As New List(Of Tuple(Of String, String, String))()
            subjects.Add(Tuple.Create("PipeType", "", typeName))

            For Each fitting In grp.Where(Function(x) x.GroupType <> RvtDB.RoutingPreferenceRuleGroupType.Segments)
                Dim subjectName As String = fitting.PartType
                If String.IsNullOrWhiteSpace(subjectName) Then
                    subjectName = fitting.PartFamily
                End If
                If String.IsNullOrWhiteSpace(subjectName) Then
                    subjectName = fitting.GroupType.ToString()
                End If
                subjects.Add(Tuple.Create("Fitting", fitting.GroupType.ToString(), subjectName))
            Next

            If segmentNames.Count = 0 Then
                For Each subj In subjects
                    res.Add(New ClassMaterialResult With {
                        .FileName = fileName,
                        .PipeTypeName = typeName,
                        .SubjectKind = subj.Item1,
                        .FittingPart = subj.Item2,
                        .SubjectName = subj.Item3,
                        .SubjectClass = ExtractSubjectClass(subj.Item3),
                        .SegmentKey = "",
                        .SegmentClass = "",
                        .Status = "NO_SEGMENT"
                    })
                Next
                Continue For
            End If

            For Each subj In subjects
                Dim subjClass As String = ExtractSubjectClass(subj.Item3)
                Dim subjNorm As String = NormalizeClassToken(subjClass)
                For Each segName In segmentNames
                    Dim segClass As String = ExtractSegmentClass(segName)
                    Dim segNorm As String = NormalizeClassToken(segClass)

                    Dim status As String
                    If String.IsNullOrWhiteSpace(segClass) OrElse String.IsNullOrWhiteSpace(subjClass) Then
                        status = "UNPARSEABLE"
                    ElseIf String.Equals(segNorm, subjNorm, StringComparison.Ordinal) Then
                        status = "PASS"
                    Else
                        status = "MISMATCH"
                    End If

                    res.Add(New ClassMaterialResult With {
                        .FileName = fileName,
                        .PipeTypeName = typeName,
                        .SubjectKind = subj.Item1,
                        .FittingPart = subj.Item2,
                        .SubjectName = subj.Item3,
                        .SubjectClass = subjClass,
                        .SegmentKey = segName,
                        .SegmentClass = segClass,
                        .Status = status
                    })
                Next
            Next
        Next

        Return res
    End Function

    '=========================================================
    ' Export - Segment
    '=========================================================
    Private Sub ExportSegmentResultToExcel(outputXlsxPath As String,
                                          entries As List(Of SegmentSizeEntry),
                                          errors As List(Of ErrorEntry),
                                          unitLabel As String,
                                          allFiles As List(Of String))

        Dim wb As IWorkbook = New XSSFWorkbook()

        ' Raw
        Dim shRaw As ISheet = wb.CreateSheet("SegmentRaw")
        Dim hr As IRow = shRaw.CreateRow(0)
        SetCell(hr, 0, "SegmentName")
        SetCell(hr, 1, "FileName")
        SetCell(hr, 2, "ND_" & unitLabel)
        SetCell(hr, 3, "ID_" & unitLabel)
        SetCell(hr, 4, "OD_" & unitLabel)

        Dim r As Integer = 1
        For Each e As SegmentSizeEntry In entries
            Dim row As IRow = shRaw.CreateRow(r)
            SetCell(row, 0, e.SegmentName)
            SetCell(row, 1, e.FileName)
            SetCell(row, 2, e.Nd)
            SetCell(row, 3, e.Id)
            SetCell(row, 4, e.Od)
            r += 1
        Next

        ' Compare
        Dim shCmp As ISheet = wb.CreateSheet("SegmentCompare")
        Dim hc As IRow = shCmp.CreateRow(0)
        SetCell(hc, 0, "SegmentName")
        SetCell(hc, 1, "FileName")
        SetCell(hc, 2, "Signature")
        SetCell(hc, 3, "IsDifferent")

        Dim fileShorts As List(Of String) =
            allFiles.Select(Function(p) System.IO.Path.GetFileName(p)).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        r = 1
        For Each segGrp In entries.GroupBy(Function(x) x.SegmentName, StringComparer.OrdinalIgnoreCase)
            Dim segName As String = segGrp.Key

            Dim sigByFile As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each fileGrp In segGrp.GroupBy(Function(x) x.FileName, StringComparer.OrdinalIgnoreCase)
                Dim ordered = fileGrp.OrderBy(Function(x) x.Nd).ThenBy(Function(x) x.Id).ThenBy(Function(x) x.Od).ToList()
                Dim parts As New List(Of String)()
                For Each it In ordered
                    parts.Add($"ND={Fmt3(it.Nd)}{unitLabel},ID={Fmt3(it.Id)}{unitLabel},OD={Fmt3(it.Od)}{unitLabel}")
                Next
                sigByFile(fileGrp.Key) = String.Join(";", parts)
            Next

            Dim missing As Boolean = sigByFile.Count <> fileShorts.Count
            Dim distinctSig As Integer = sigByFile.Values.Distinct(StringComparer.Ordinal).Count()
            Dim isDiff As String = If(missing OrElse distinctSig > 1, "TRUE", "FALSE")

            For Each fn As String In fileShorts
                Dim sig As String = ""
                If sigByFile.ContainsKey(fn) Then
                    sig = sigByFile(fn)
                Else
                    sig = "<MISSING>"
                End If

                Dim row As IRow = shCmp.CreateRow(r)
                SetCell(row, 0, segName)
                SetCell(row, 1, fn)
                SetCell(row, 2, sig)
                SetCell(row, 3, isDiff)
                r += 1
            Next
        Next

        ' Error
        Dim shErr As ISheet = wb.CreateSheet("Error")
        Dim he As IRow = shErr.CreateRow(0)
        SetCell(he, 0, "FilePath")
        SetCell(he, 1, "Stage")
        SetCell(he, 2, "ErrorMessage")

        r = 1
        For Each ee As ErrorEntry In errors
            Dim row As IRow = shErr.CreateRow(r)
            SetCell(row, 0, ee.FilePath)
            SetCell(row, 1, ee.Stage)
            SetCell(row, 2, ee.Message)
            r += 1
        Next

        ' Save
        Using fs As New FileStream(outputXlsxPath, FileMode.Create, FileAccess.Write, FileShare.None)
            wb.Write(fs)
        End Using
    End Sub

    '=========================================================
    ' Export - Routing
    '=========================================================
    Private Sub ExportRoutingResultToExcel(outputXlsxPath As String,
                                          entries As List(Of RoutingRuleEntry),
                                          errors As List(Of ErrorEntry),
                                          unitLabel As String,
                                          allFiles As List(Of String),
                                          classMatches As List(Of ClassMaterialResult))

        Dim wb As IWorkbook = New XSSFWorkbook()

        ' Raw
        Dim shRaw As ISheet = wb.CreateSheet("RoutingRaw")
        Dim hr As IRow = shRaw.CreateRow(0)
        SetCell(hr, 0, "Category")
        SetCell(hr, 1, "TypeName")
        SetCell(hr, 2, "FileName")
        SetCell(hr, 3, "GroupType")
        SetCell(hr, 4, "RuleIndex")
        SetCell(hr, 5, "PartFamily")
        SetCell(hr, 6, "PartType")
        SetCell(hr, 7, "MinSize_" & unitLabel)
        SetCell(hr, 8, "MaxSize_" & unitLabel)

        Dim r As Integer = 1
        For Each e As RoutingRuleEntry In entries
            Dim row As IRow = shRaw.CreateRow(r)
            SetCell(row, 0, e.CategoryName)
            SetCell(row, 1, e.TypeName)
            SetCell(row, 2, e.FileName)
            SetCell(row, 3, e.GroupType.ToString())
            SetCell(row, 4, CDbl(e.RuleIndex))
            SetCell(row, 5, e.PartFamily)
            SetCell(row, 6, e.PartType)
            SetCell(row, 7, e.MinSize)
            SetCell(row, 8, e.MaxSize)
            r += 1
        Next

        ' Compare
        Dim shCmp As ISheet = wb.CreateSheet("RoutingCompare")
        Dim hc As IRow = shCmp.CreateRow(0)
        SetCell(hc, 0, "Category")
        SetCell(hc, 1, "TypeName")
        SetCell(hc, 2, "FileName")
        SetCell(hc, 3, "Signature")
        SetCell(hc, 4, "IsDifferent")

        Dim fileShorts As List(Of String) =
            allFiles.Select(Function(p) System.IO.Path.GetFileName(p)).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        r = 1
        For Each keyGrp In entries.GroupBy(Function(x) x.CategoryName & "|" & x.TypeName, StringComparer.OrdinalIgnoreCase)
            Dim partsKey = keyGrp.Key.Split("|"c)
            Dim catName As String = partsKey(0)
            Dim typeName As String = partsKey(1)

            Dim sigByFile As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each fileGrp In keyGrp.GroupBy(Function(x) x.FileName, StringComparer.OrdinalIgnoreCase)
                Dim ordered = fileGrp.
                    OrderBy(Function(x) x.GroupType.ToString()).
                    ThenBy(Function(x) x.RuleIndex).
                    ToList()

                Dim sigParts As New List(Of String)()
                For Each it In ordered
                    sigParts.Add($"[{it.GroupType}]#{it.RuleIndex}:{it.PartFamily}|{it.PartType}|{Fmt3(it.MinSize)}{unitLabel}-{Fmt3(it.MaxSize)}{unitLabel}")
                Next

                sigByFile(fileGrp.Key) = String.Join(";", sigParts)
            Next

            Dim missing As Boolean = sigByFile.Count <> fileShorts.Count
            Dim distinctSig As Integer = sigByFile.Values.Distinct(StringComparer.Ordinal).Count()
            Dim isDiff As String = If(missing OrElse distinctSig > 1, "TRUE", "FALSE")

            For Each fn As String In fileShorts
                Dim sig As String = If(sigByFile.ContainsKey(fn), sigByFile(fn), "<MISSING>")

                Dim row As IRow = shCmp.CreateRow(r)
                SetCell(row, 0, catName)
                SetCell(row, 1, typeName)
                SetCell(row, 2, fn)
                SetCell(row, 3, sig)
                SetCell(row, 4, isDiff)
                r += 1
            Next
        Next

        ' Error
        Dim shErr As ISheet = wb.CreateSheet("Error")
        Dim he As IRow = shErr.CreateRow(0)
        SetCell(he, 0, "FilePath")
        SetCell(he, 1, "Stage")
        SetCell(he, 2, "ErrorMessage")

        r = 1
        For Each ee As ErrorEntry In errors
            Dim row As IRow = shErr.CreateRow(r)
            SetCell(row, 0, ee.FilePath)
            SetCell(row, 1, ee.Stage)
            SetCell(row, 2, ee.Message)
            r += 1
        Next

        If classMatches IsNot Nothing Then
            Dim shClass As ISheet = wb.CreateSheet("Routing_ClassMatch")
            Dim hc2 As IRow = shClass.CreateRow(0)
            SetCell(hc2, 0, "File")
            SetCell(hc2, 1, "PipeTypeName")
            SetCell(hc2, 2, "SubjectKind")
            SetCell(hc2, 3, "FittingPart")
            SetCell(hc2, 4, "SubjectName")
            SetCell(hc2, 5, "SubjectClass")
            SetCell(hc2, 6, "SegmentKey")
            SetCell(hc2, 7, "SegmentClass")
            SetCell(hc2, 8, "Status")

            r = 1
            If classMatches.Count > 0 Then
                For Each cm As ClassMaterialResult In classMatches
                    Dim row As IRow = shClass.CreateRow(r)
                    SetCell(row, 0, cm.FileName)
                    SetCell(row, 1, cm.PipeTypeName)
                    SetCell(row, 2, cm.SubjectKind)
                    SetCell(row, 3, cm.FittingPart)
                    SetCell(row, 4, cm.SubjectName)
                    SetCell(row, 5, cm.SubjectClass)
                    SetCell(row, 6, cm.SegmentKey)
                    SetCell(row, 7, cm.SegmentClass)
                    SetCell(row, 8, cm.Status)
                    r += 1
                Next
            End If
        End If

        Using fs As New FileStream(outputXlsxPath, FileMode.Create, FileAccess.Write, FileShare.None)
            wb.Write(fs)
        End Using
    End Sub

End Class
