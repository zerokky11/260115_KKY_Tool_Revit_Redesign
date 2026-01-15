Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Drawing
Imports WinForms = System.Windows.Forms

Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

' NPOI 타입은 전부 Object + 리플렉션으로만 사용 (모호성 방지)
' 아래 경고는 의도적인 late binding이라 전역으로 끈다.
#Disable Warning BC42016 ' Object에서 다른 형식으로의 암시적 변환
#Disable Warning BC42017 ' 런타임에 바인딩 확인 (late binding)
#Disable Warning BC42019 ' 연산자에 Object 피연산자 사용

<Transaction(TransactionMode.ReadOnly)>
Public Class CmdCheckParameterCombinations
    Implements IExternalCommand

    '==================== 메인 진입점 ====================
    Public Function Execute(commandData As ExternalCommandData,
                            ByRef message As String,
                            elements As ElementSet) As Result Implements IExternalCommand.Execute

        Dim uiApp As UIApplication = commandData.Application
        Dim uiDoc As UIDocument = uiApp.ActiveUIDocument

        If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
            message = "열린 프로젝트 문서가 없습니다."
            Return Result.Failed
        End If

        Dim doc As Document = uiDoc.Document

        ' 설정 파일은 사용자 AppData 아래에 저장해서, 다음 실행 때도 유지
        Dim appData As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim configDir As String = Path.Combine(appData, "ParamConstructionClassification")
        Dim configPath As String = Path.Combine(configDir, "settings.txt")

        Try
            Using frm As New ParamCheckForm(doc, configPath)
                Dim r = frm.ShowDialog()
                If r = WinForms.DialogResult.OK OrElse r = WinForms.DialogResult.Cancel Then
                    Return Result.Succeeded
                Else
                    Return Result.Succeeded
                End If
            End Using
        Catch ex As Exception
            message = ex.Message
            Return Result.Failed
        End Try

    End Function

    '==================== 도메인 모델 ====================

    Public Class Criteria
        Public Property ParamA As String
        Public Property ParamB As String
        Public Property ParamC As String
        Public Property AllowedCombos As HashSet(Of String)
        Public Property TargetMap As Dictionary(Of String, String)
        Public Property SheetName As String
    End Class

    Public Class AnalysisResult
        Public Property ResultTable As DataTable

        Public Property TotalChecked As Integer
        Public Property OkCount As Integer
        Public Property NeedFixCount As Integer
        Public Property MissingParamCount As Integer
        Public Property NotInCriteriaCount As Integer
        Public Property EmptyValueCount As Integer
    End Class

    '==================== NPOI Workbook 생성(리플렉션) ====================

    Friend Shared Function CreateWorkbookFromStream(ext As String, fs As Stream) As Object
        Dim t As Type

        If ext = ".xlsx" Then
            ' XSSFWorkbook (NPOI.OOXML)
            t = Type.GetType("NPOI.XSSF.UserModel.XSSFWorkbook, NPOI.OOXML", True)
            Return Activator.CreateInstance(t, fs)
        ElseIf ext = ".xls" Then
            ' HSSFWorkbook (NPOI.Core)
            t = Type.GetType("NPOI.HSSF.UserModel.HSSFWorkbook, NPOI.Core", True)
            Return Activator.CreateInstance(t, fs)
        Else
            Throw New InvalidOperationException("지원되지 않는 엑셀 형식입니다. (*.xlsx 또는 *.xls만 지원)")
        End If
    End Function

    Friend Shared Function CreateEmptyWorkbook(ext As String) As Object
        Dim t As Type

        If ext = ".xlsx" Then
            t = Type.GetType("NPOI.XSSF.UserModel.XSSFWorkbook, NPOI.OOXML", True)
            Return Activator.CreateInstance(t)
        ElseIf ext = ".xls" Then
            t = Type.GetType("NPOI.HSSF.UserModel.HSSFWorkbook, NPOI.Core", True)
            Return Activator.CreateInstance(t)
        Else
            Throw New InvalidOperationException("지원되지 않는 엑셀 형식입니다. (*.xlsx 또는 *.xls만 지원)")
        End If
    End Function

    ' ParamCheckForm에서도 써야 하니까 Friend
    Friend Shared Sub TryCloseWorkbook(wb As Object)
        If wb Is Nothing Then Return
        Try
            wb.Close()
        Catch
            ' 무시 (버전에 따라 Close 없을 수 있음)
        End Try
    End Sub

    '==================== 엑셀 기준 로딩 ====================

    ''' <summary>
    ''' 엑셀에서 기준 조합 + 검토대상 읽기
    ''' 시트: 이름이 "CUSTOM1" 인 시트를 우선 시도, 없으면 첫 번째 시트 사용
    ''' 헤더: SB_PHASE, SB_EXCLUSION, SB_FUTURE, 검토대상
    ''' </summary>
    Friend Shared Function LoadCriteria(excelPath As String) As Criteria
        If String.IsNullOrEmpty(excelPath) OrElse Not File.Exists(excelPath) Then
            Throw New FileNotFoundException("엑셀 파일을 찾을 수 없습니다.", excelPath)
        End If

        Dim ext As String = Path.GetExtension(excelPath).ToLowerInvariant()

        Using fs As New FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim workbook As Object = CreateWorkbookFromStream(ext, fs)

            Dim sheet As Object = workbook.GetSheet("CUSTOM1")
            If sheet Is Nothing Then
                sheet = workbook.GetSheetAt(0)
            End If

            If sheet Is Nothing Then
                Throw New InvalidOperationException("엑셀에 시트가 없습니다.")
            End If

            Dim headerRow As Object = sheet.GetRow(sheet.FirstRowNum)
            If headerRow Is Nothing Then
                Throw New InvalidOperationException("엑셀에 헤더 행이 없습니다.")
            End If

            ' 헤더 위치 찾기
            Dim idxPhase As Integer = -1
            Dim idxExclusion As Integer = -1
            Dim idxFuture As Integer = -1
            Dim idxTarget As Integer = -1

            For ci As Integer = headerRow.FirstCellNum To headerRow.LastCellNum - 1
                Dim name As String = GetCellString(headerRow.GetCell(ci)).Trim()
                Select Case name
                    Case "SB_PHASE"
                        idxPhase = ci
                    Case "SB_EXCLUSION"
                        idxExclusion = ci
                    Case "SB_FUTURE"
                        idxFuture = ci
                    Case "검토대상"
                        idxTarget = ci
                End Select
            Next

            If idxPhase = -1 OrElse idxExclusion = -1 OrElse idxFuture = -1 Then
                Throw New InvalidOperationException("헤더 행에 SB_PHASE, SB_EXCLUSION, SB_FUTURE 열을 찾을 수 없습니다.")
            End If

            If idxTarget = -1 Then
                Throw New InvalidOperationException("헤더 행에 '검토대상' 열을 찾을 수 없습니다.")
            End If

            Dim combos As New HashSet(Of String)(StringComparer.Ordinal)
            Dim targetMap As New Dictionary(Of String, String)(StringComparer.Ordinal)

            For i As Integer = sheet.FirstRowNum + 1 To sheet.LastRowNum
                Dim row As Object = sheet.GetRow(i)
                If row Is Nothing Then Continue For

                Dim aVal As String = GetCellString(row.GetCell(idxPhase)).Trim()
                Dim bVal As String = GetCellString(row.GetCell(idxExclusion)).Trim()
                Dim cVal As String = GetCellString(row.GetCell(idxFuture)).Trim()
                Dim tVal As String = GetCellString(row.GetCell(idxTarget)).Trim()

                ' 세 값 모두 비어 있으면 무시 (빈 행)
                If String.IsNullOrEmpty(aVal) AndAlso
                   String.IsNullOrEmpty(bVal) AndAlso
                   String.IsNullOrEmpty(cVal) Then
                    Continue For
                End If

                Dim key As String = MakeKey(aVal, bVal, cVal)
                combos.Add(key)

                If Not targetMap.ContainsKey(key) Then
                    targetMap.Add(key, tVal)
                End If
            Next

            TryCloseWorkbook(workbook)

            If combos.Count = 0 Then
                Throw New InvalidOperationException("엑셀에서 허용 조합을 찾지 못했습니다. SB_PHASE/SB_EXCLUSION/SB_FUTURE 값을 확인해 주세요.")
            End If

            Dim c As New Criteria()
            c.ParamA = "SB_PHASE"
            c.ParamB = "SB_EXCLUSION"
            c.ParamC = "SB_FUTURE"
            c.AllowedCombos = combos
            c.TargetMap = targetMap
            c.SheetName = sheet.SheetName

            Return c
        End Using
    End Function

    ' NPOI CellType/ICell 의존 안 쓰고 문자열만 뽑기
    Friend Shared Function GetCellString(cell As Object) As String
        If cell Is Nothing Then Return String.Empty
        Return cell.ToString()
    End Function

    '==================== 도큐먼트 검사 로직 ====================

    Friend Shared Function AnalyzeDocument(doc As Document,
                                           criteria As Criteria,
                                           bqcOnly As Boolean) As AnalysisResult

        Dim result As New AnalysisResult()
        Dim table As New DataTable("ParamCombinationResult")

        ' ===== 결과 테이블 스키마 (엑셀 헤더와 1:1, 순서 고정) =====
        Dim headers As String() = {
            "항목",
            "ID",
            "Name",
            "결과",
            "내용",
            "검토내용",
            "비고(답변)",
            "Category",
            "Family",
            "SB_PHASE",
            "SB_EXCLUSION",
            "SB_FUTURE",
            "SB_CONDATE",
            "SB_REV",
            "SB_REVDATE",
            "S5_MATERIAL",
            "S5_UTILITY",
            "SB_FIELD",
            "SB_FL",
            "S5_EQCODE",
            "SB_SC",
            "CUSTOM_1",
            "CUSTOM_2",
            "CUSTOM_3",
            "NOTE-1",
            "NOTE-2",
            "EXCLUSION",
            "Comments",
            "UPDATE_TIME_DESIGN",
            "USER_NAME_DESIGN",
            "UPDATE_TIME_CONSTRUCTION",
            "USER_NAME_CONSTRUCTION",
            "UPDATE_TIME-설계정보",
            "USER_NAME-설계정보",
            "UPDATE_TIME-시공정보",
            "USER_NAME-시공정보"
        }

        For Each h In headers
            table.Columns.Add(h, GetType(String))
        Next

        result.ResultTable = table

        Dim collector As New FilteredElementCollector(doc)
        collector.WhereElementIsNotElementType()

        Dim index As Integer = 0

        For Each e As Element In collector
            If e Is Nothing Then Continue For
            If e.Category Is Nothing Then Continue For

            ' 0) 링크 / CAD 타입은 무조건 제외
            If TypeOf e Is ImportInstance Then Continue For
            If TypeOf e Is RevitLinkInstance Then Continue For

            ' 1) 모델 카테고리만
            If e.Category.CategoryType <> CategoryType.Model Then Continue For

            ' 2) 모델 카테고리 중에서도 특정 카테고리는 제외
            If IsSkippableModelCategory(e.Category, e) Then Continue For

            ' 여기까지 통과한 애들만 "총 요소"에 카운트
            result.TotalChecked += 1

            Dim valA As String = Nothing
            Dim valB As String = Nothing
            Dim valC As String = Nothing

            Dim hasA As Boolean = TryGetParamString(e, criteria.ParamA, valA)
            Dim hasB As Boolean = TryGetParamString(e, criteria.ParamB, valB)
            Dim hasC As Boolean = TryGetParamString(e, criteria.ParamC, valC)

            Dim key As String = MakeKey(If(valA, String.Empty),
                                        If(valB, String.Empty),
                                        If(valC, String.Empty))

            Dim targetText As String = Nothing
            If criteria.TargetMap IsNot Nothing AndAlso criteria.TargetMap.TryGetValue(key, targetText) Then
                ' OK
            Else
                targetText = Nothing
            End If

            Dim missingParam As Boolean = (Not hasA) OrElse (Not hasB) OrElse (Not hasC)
            Dim allEmpty As Boolean =
                String.IsNullOrEmpty(valA) AndAlso
                String.IsNullOrEmpty(valB) AndAlso
                String.IsNullOrEmpty(valC)

            Dim needFix As Boolean = False

            If missingParam Then
                needFix = True
                result.MissingParamCount += 1

            ElseIf allEmpty Then
                needFix = True
                result.EmptyValueCount += 1

            Else
                If criteria.AllowedCombos IsNot Nothing AndAlso criteria.AllowedCombos.Contains(key) Then
                    needFix = False
                    result.OkCount += 1
                Else
                    needFix = True
                    result.NotInCriteriaCount += 1
                End If
            End If

            If needFix Then
                ' BQC 필터 적용
                If bqcOnly Then
                    ' 엑셀 기준에서 검토대상 문자열이 BQC 포함일 때만 결과에 포함
                    If String.IsNullOrEmpty(targetText) OrElse
                       targetText.IndexOf("BQC", StringComparison.OrdinalIgnoreCase) < 0 Then
                        Continue For
                    End If
                End If

                result.NeedFixCount += 1
                index += 1

                Dim row As DataRow = table.NewRow()

                ' ===== 공통 메타 =====
                row("항목") = index.ToString()
                row("ID") = e.Id.IntegerValue.ToString()
                row("Name") = GetTypeName(e)

                ' 내용/검토내용/비고(답변)
                Dim noteText As String =
                    String.Format("[{0}] 파라미터 기준(오/탈자) 오류: SB_PHASE,SB_EXCLUSION,SB_FUTURE",
                                  criteria.SheetName)

                Dim answerText As String = ""   ' 비고(답변) 초기값은 공란
                Dim reviewText As String = BuildReviewMessage(noteText, answerText)

                row("결과") = "오류"
                row("내용") = noteText
                row("검토내용") = reviewText
                row("비고(답변)") = answerText

                row("Category") = SafeCategoryName(e)
                row("Family") = GetFamilyName(e)

                ' ===== 파라미터 값 채우기 =====
                Dim paramNames As String() = {
                    "SB_PHASE",
                    "SB_EXCLUSION",
                    "SB_FUTURE",
                    "SB_CONDATE",
                    "SB_REV",
                    "SB_REVDATE",
                    "S5_MATERIAL",
                    "S5_UTILITY",
                    "SB_FIELD",
                    "SB_FL",
                    "S5_EQCODE",
                    "SB_SC",
                    "CUSTOM_1",
                    "CUSTOM_2",
                    "CUSTOM_3",
                    "NOTE-1",
                    "NOTE-2",
                    "EXCLUSION",
                    "Comments",
                    "UPDATE_TIME_DESIGN",
                    "USER_NAME_DESIGN",
                    "UPDATE_TIME_CONSTRUCTION",
                    "USER_NAME_CONSTRUCTION",
                    "UPDATE_TIME-설계정보",
                    "USER_NAME-설계정보",
                    "UPDATE_TIME-시공정보",
                    "USER_NAME-시공정보"
                }

                For Each pName In paramNames
                    Dim pVal As String = Nothing
                    If TryGetParamString(e, pName, pVal) Then
                        row(pName) = pVal
                    Else
                        row(pName) = String.Empty
                    End If
                Next

                table.Rows.Add(row)
            End If
        Next

        Return result
    End Function

    ' ====== 모델 카테고리 중에서도 제외해야 할 것들 ======
    Friend Shared Function IsSkippableModelCategory(cat As Category, elem As Element) As Boolean
        If cat Is Nothing Then Return True

        ' 1) BuiltInCategory 기반 필터 – 명백히 실물과 무관한 것들만 제외
        Try
            Dim bic As BuiltInCategory = CType(cat.Id.IntegerValue, BuiltInCategory)
            Select Case bic
                Case BuiltInCategory.OST_Lines               ' 모델 선
                    Return True
                Case BuiltInCategory.OST_RvtLinks            ' 링크 파일 카테고리
                    Return True
                Case BuiltInCategory.OST_ProjectBasePoint    ' 프로젝트 기준점
                    Return True
                Case BuiltInCategory.OST_Levels              ' 레벨
                    Return True
                Case BuiltInCategory.OST_Grids               ' 그리드
                    Return True
                Case BuiltInCategory.OST_Site                ' 사이트(지형 등) – 이 도구에서는 제외
                    Return True
                Case BuiltInCategory.OST_Cameras             ' 카메라
                    Return True
            End Select
        Catch
            ' 캐스팅 실패 시 아래 이름 기반 필터로만 처리
        End Try

        ' 2) 이름 기반 필터 (환경/시스템/도면/해석 계열 등)
        Dim name As String = cat.Name
        If String.IsNullOrEmpty(name) = False Then
            Dim n As String = name.Trim()

            ' 실물이 아닌 카테고리 이름들
            Dim skipNames As String() = {
                "Project Information",
                "Sun Path",
                "Primary Contours",
                "Legend Components",
                "HVAC Zones",
                "Material Assets",
                "Materials",
                "Survey Point",
                "Pipe Segments",
                "Piping Systems",
                "Duct Systems",
                "Center line",
                "<Sketch>"
            }

            For Each s In skipNames
                If String.Equals(n, s, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next

            ' Category 이름에 .dwg / .dxf 등이 들어가면 CAD라고 보고 제외
            Dim lower As String = n.ToLowerInvariant()
            If lower.Contains(".dwg") OrElse lower.Contains(".dxf") Then
                Return True
            End If
        End If

        ' 3) 요소 타입 기반 안전장치 (Import / Link)
        If elem IsNot Nothing Then
            If TypeOf elem Is ImportInstance Then Return True
            If TypeOf elem Is RevitLinkInstance Then Return True
        End If

        ' 나머지는 전부 "실물 모델링 요소"로 간주 (구조/건축/설비/전기 등)
        Return False
    End Function

    Friend Shared Function TryGetParamString(elem As Element, paramName As String, ByRef value As String) As Boolean
        value = Nothing

        If elem Is Nothing OrElse String.IsNullOrEmpty(paramName) Then
            Return False
        End If

        Dim p As Parameter = elem.LookupParameter(paramName)
        If p Is Nothing Then Return False

        If p.StorageType = StorageType.String Then
            value = p.AsString()
        ElseIf p.StorageType = StorageType.Double Then
            value = p.AsDouble().ToString()
        ElseIf p.StorageType = StorageType.Integer Then
            value = p.AsInteger().ToString()
        ElseIf p.StorageType = StorageType.ElementId Then
            Dim id As ElementId = p.AsElementId()
            If id IsNot Nothing AndAlso id.IntegerValue <> -1 Then
                value = id.IntegerValue.ToString()
            Else
                value = String.Empty
            End If
        Else
            value = p.AsValueString()
        End If

        If value Is Nothing Then value = String.Empty
        Return True
    End Function

    Friend Shared Function SafeCategoryName(e As Element) As String
        If e Is Nothing OrElse e.Category Is Nothing Then Return String.Empty
        Return e.Category.Name
    End Function

    Friend Shared Function GetFamilyName(e As Element) As String
        If e Is Nothing Then Return String.Empty

        Dim famName As String = String.Empty

        Dim inst As FamilyInstance = TryCast(e, FamilyInstance)
        If inst IsNot Nothing AndAlso inst.Symbol IsNot Nothing AndAlso inst.Symbol.Family IsNot Nothing Then
            famName = inst.Symbol.Family.Name
        Else
            Dim et As ElementType = TryCast(e.Document.GetElement(e.GetTypeId()), ElementType)
            Dim fs As FamilySymbol = TryCast(et, FamilySymbol)
            If fs IsNot Nothing AndAlso fs.Family IsNot Nothing Then
                famName = fs.Family.Name
            End If
        End If

        Return famName
    End Function

    Friend Shared Function GetTypeName(e As Element) As String
        If e Is Nothing Then Return String.Empty

        Dim inst As FamilyInstance = TryCast(e, FamilyInstance)
        If inst IsNot Nothing AndAlso inst.Symbol IsNot Nothing Then
            Return inst.Symbol.Name
        End If

        Dim et As ElementType = TryCast(e.Document.GetElement(e.GetTypeId()), ElementType)
        If et IsNot Nothing Then
            Return et.Name
        End If

        Return String.Empty
    End Function

    Friend Shared Function MakeKey(a As String, b As String, c As String) As String
        If a Is Nothing Then a = String.Empty
        If b Is Nothing Then b = String.Empty
        If c Is Nothing Then c = String.Empty

        Return String.Concat(a.Trim(), "||", b.Trim(), "||", c.Trim())
    End Function

    ' ====== 엑셀 LET 수식이 하는 일을 VB로 그대로 구현 ======
    Friend Shared Function BuildReviewMessage(note As String, answer As String) As String
        ' LET_비고구분 = answer
        ' LET_내용     = note

        If answer = "REV, REVDATE" Then
            Return ""
        End If

        If answer = "검토제외대상" Then
            Return "검토제외대상 : 해당 납품대상 객체 아님"
        End If

        Dim prefix As String = If(answer = "공사구분", "납품대상 객체이면 ", "")

        If note Is Nothing Then note = String.Empty

        Dim msg As String

        If note.Contains("[CUSTOM1]") Then
            ' SB_PHASE, SB_EXCLUSION, SB_FUTURE 기준
            msg = "SB_PHASE, SB_EXCLUSION, SB_FUTURE 기준 및 대소문자 확인필요"
        ElseIf note.Contains("[CUSTOM2]") Then
            msg = "SB_FL 대소문자 및 오탈자 확인 필요"
        ElseIf note.Contains("[S5_UTILITY]") Then
            msg = "S5_UTILITY 대소문자 및 오탈자 확인필요"
        ElseIf note.Contains("[S5_MATERIAL]") Then
            msg = "S5_MATERIAL 대소문자 및 오탈자 확인필요"
        Else
            msg = "CUSTOM조건 확인필요"
        End If

        Return prefix & msg
    End Function

End Class

'==================== UI 폼 ====================

Friend Class ParamCheckForm
    Inherits WinForms.Form

    Private ReadOnly _doc As Document
    Private ReadOnly _settingsPath As String

    Private _excelPath As String
    Private _bqcOnly As Boolean
    Private _criteria As CmdCheckParameterCombinations.Criteria
    Private _analysis As CmdCheckParameterCombinations.AnalysisResult

    Private _txtExcel As WinForms.TextBox
    Private _btnBrowse As WinForms.Button
    Private _chkBqc As WinForms.CheckBox
    Private _btnRun As WinForms.Button
    Private _btnExport As WinForms.Button
    Private _lblStatus As WinForms.Label
    Private _grid As WinForms.DataGridView

    Public Sub New(doc As Document, settingsPath As String)
        _doc = doc
        _settingsPath = settingsPath

        InitializeComponent()
        LoadSettings()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "파라미터 조합 검토"
        Me.StartPosition = WinForms.FormStartPosition.CenterScreen
        Me.Size = New Size(1200, 700)

        _txtExcel = New WinForms.TextBox() With {
            .Left = 12,
            .Top = 12,
            .Width = 800,
            .ReadOnly = True
        }

        _btnBrowse = New WinForms.Button() With {
            .Left = 820,
            .Top = 10,
            .Width = 120,
            .Text = "기준 엑셀..."
        }

        _chkBqc = New WinForms.CheckBox() With {
            .Left = 12,
            .Top = 40,
            .Width = 280,
            .Text = "검토대상에 'BQC' 포함 오류만 보기"
        }

        _btnRun = New WinForms.Button() With {
            .Left = 310,
            .Top = 38,
            .Width = 100,
            .Text = "검토 실행"
        }

        _btnExport = New WinForms.Button() With {
            .Left = 420,
            .Top = 38,
            .Width = 120,
            .Text = "엑셀로 내보내기",
            .Enabled = False
        }

        _lblStatus = New WinForms.Label() With {
            .Left = 12,
            .Top = 70,
            .Width = 1150,
            .AutoSize = True
        }

        _grid = New WinForms.DataGridView() With {
            .Left = 12,
            .Top = 95,
            .Width = 1150,
            .Height = 520,
            .Anchor = WinForms.AnchorStyles.Top Or WinForms.AnchorStyles.Bottom Or WinForms.AnchorStyles.Left Or WinForms.AnchorStyles.Right,
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        }

        AddHandler _btnBrowse.Click, AddressOf OnBrowseClick
        AddHandler _btnRun.Click, AddressOf OnRunClick
        AddHandler _btnExport.Click, AddressOf OnExportClick

        Me.Controls.AddRange(New WinForms.Control() {
            _txtExcel,
            _btnBrowse,
            _chkBqc,
            _btnRun,
            _btnExport,
            _lblStatus,
            _grid
        })
    End Sub

    '=========== 설정 저장/로드 ===========

    Private Sub LoadSettings()
        Try
            Dim dir As String = Path.GetDirectoryName(_settingsPath)
            If String.IsNullOrEmpty(dir) OrElse Not Directory.Exists(dir) Then
                Return
            End If
            If Not File.Exists(_settingsPath) Then
                Return
            End If

            Dim lines As String() = File.ReadAllLines(_settingsPath, Encoding.UTF8)
            For Each line In lines
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim parts = line.Split(New Char() {"="c}, 2)
                If parts.Length <> 2 Then Continue For

                Dim key = parts(0).Trim()
                Dim value = parts(1).Trim()

                Select Case key
                    Case "ExcelPath"
                        _excelPath = value
                    Case "BqcOnly"
                        Dim b As Boolean = False
                        Boolean.TryParse(value, b)
                        _bqcOnly = b
                End Select
            Next

            _txtExcel.Text = _excelPath
            _chkBqc.Checked = _bqcOnly

        Catch
            ' 무시
        End Try
    End Sub

    Private Sub SaveSettings()
        Try
            Dim dir As String = Path.GetDirectoryName(_settingsPath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim pathValue As String = If(_excelPath, String.Empty)
            Dim bqcValue As String = _bqcOnly.ToString()

            Dim lines As String() = {
                "ExcelPath=" & pathValue,
                "BqcOnly=" & bqcValue
            }

            File.WriteAllLines(_settingsPath, lines, Encoding.UTF8)
        Catch
            ' 무시
        End Try
    End Sub

    '=========== 버튼 핸들러 ===========

    Private Sub OnBrowseClick(sender As Object, e As EventArgs)
        Using dlg As New WinForms.OpenFileDialog()
            dlg.Title = "기준 조합 엑셀 파일 선택"
            dlg.Filter = "Excel 파일 (*.xlsx;*.xls)|*.xlsx;*.xls|모든 파일 (*.*)|*.*"
            dlg.Multiselect = False

            If Not String.IsNullOrEmpty(_excelPath) AndAlso File.Exists(_excelPath) Then
                dlg.InitialDirectory = Path.GetDirectoryName(_excelPath)
                dlg.FileName = Path.GetFileName(_excelPath)
            End If

            If dlg.ShowDialog(Me) = WinForms.DialogResult.OK Then
                _excelPath = dlg.FileName
                _txtExcel.Text = _excelPath
            End If
        End Using
    End Sub

    Private Sub OnRunClick(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(_excelPath) OrElse Not File.Exists(_excelPath) Then
            WinForms.MessageBox.Show(Me,
                                     "기준 엑셀 파일을 먼저 지정해 주세요.",
                                     "알림",
                                     WinForms.MessageBoxButtons.OK,
                                     WinForms.MessageBoxIcon.Information)
            Return
        End If

        Try
            Me.Cursor = WinForms.Cursors.WaitCursor
            _btnRun.Enabled = False

            _criteria = CmdCheckParameterCombinations.LoadCriteria(_excelPath)
            _bqcOnly = _chkBqc.Checked

            _analysis = CmdCheckParameterCombinations.AnalyzeDocument(_doc, _criteria, _bqcOnly)

            _grid.DataSource = _analysis.ResultTable
            _btnExport.Enabled = (_analysis.ResultTable IsNot Nothing AndAlso _analysis.ResultTable.Rows.Count > 0)

            UpdateStatusLabel()
            SaveSettings()

        Catch ex As Exception
            WinForms.MessageBox.Show(Me,
                                     "검토 중 오류가 발생했습니다." & Environment.NewLine & ex.Message,
                                     "오류",
                                     WinForms.MessageBoxButtons.OK,
                                     WinForms.MessageBoxIcon.Error)
        Finally
            _btnRun.Enabled = True
            Me.Cursor = WinForms.Cursors.Default
        End Try
    End Sub

    Private Sub OnExportClick(sender As Object, e As EventArgs)
        If _analysis Is Nothing OrElse _analysis.ResultTable Is Nothing OrElse _analysis.ResultTable.Rows.Count = 0 Then
            WinForms.MessageBox.Show(Me,
                                     "내보낼 검토 결과가 없습니다.",
                                     "알림",
                                     WinForms.MessageBoxButtons.OK,
                                     WinForms.MessageBoxIcon.Information)
            Return
        End If

        Using dlg As New WinForms.SaveFileDialog()
            dlg.Title = "검토 결과 엑셀로 내보내기"
            dlg.Filter = "Excel 파일 (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls"
            dlg.FileName = "ParamCombinationResult.xlsx"

            If dlg.ShowDialog(Me) <> WinForms.DialogResult.OK Then
                Return
            End If

            Try
                Me.Cursor = WinForms.Cursors.WaitCursor

                Dim dt As DataTable = _analysis.ResultTable
                Dim ext As String = Path.GetExtension(dlg.FileName).ToLowerInvariant()

                Dim wb As Object = CmdCheckParameterCombinations.CreateEmptyWorkbook(ext)
                Dim sheet As Object = wb.CreateSheet("Result")

                ' 헤더
                Dim headerRow As Object = sheet.CreateRow(0)
                For c As Integer = 0 To dt.Columns.Count - 1
                    headerRow.CreateCell(c).SetCellValue(dt.Columns(c).ColumnName)
                Next

                ' 데이터
                For r As Integer = 0 To dt.Rows.Count - 1
                    Dim dataRow As Object = sheet.CreateRow(r + 1)
                    For c As Integer = 0 To dt.Columns.Count - 1
                        Dim v As Object = dt.Rows(r)(c)
                        Dim s As String = If(v IsNot Nothing, v.ToString(), String.Empty)
                        dataRow.CreateCell(c).SetCellValue(s)
                    Next
                Next

                ' 저장
                Using fs As New FileStream(dlg.FileName, FileMode.Create, FileAccess.Write)
                    wb.Write(fs)
                End Using

                CmdCheckParameterCombinations.TryCloseWorkbook(wb)

                WinForms.MessageBox.Show(Me,
                                         "엑셀 파일로 저장되었습니다." & Environment.NewLine & dlg.FileName,
                                         "완료",
                                         WinForms.MessageBoxButtons.OK,
                                         WinForms.MessageBoxIcon.Information)

            Catch ex As Exception
                WinForms.MessageBox.Show(Me,
                                         "엑셀 내보내기 중 오류가 발생했습니다." & Environment.NewLine & ex.Message,
                                         "오류",
                                         WinForms.MessageBoxButtons.OK,
                                         WinForms.MessageBoxIcon.Error)
            Finally
                Me.Cursor = WinForms.Cursors.Default
            End Try
        End Using
    End Sub

    Private Sub UpdateStatusLabel()
        If _analysis Is Nothing Then
            _lblStatus.Text = ""
            Return
        End If

        Dim sb As New StringBuilder()
        sb.AppendFormat("총 요소: {0}, OK: {1}, 수정필요: {2} (Missing:{3}, NotInCriteria:{4}, Empty:{5})",
                        _analysis.TotalChecked,
                        _analysis.OkCount,
                        _analysis.NeedFixCount,
                        _analysis.MissingParamCount,
                        _analysis.NotInCriteriaCount,
                        _analysis.EmptyValueCount)
        sb.Append("  |  기준 시트: ")
        If _criteria IsNot Nothing Then
            sb.Append(_criteria.SheetName)
        Else
            sb.Append("-")
        End If
        sb.Append("  |  필터: ")
        sb.Append(If(_bqcOnly, "검토대상 BQC 포함만", "전체 오류"))

        _lblStatus.Text = sb.ToString()
    End Sub

End Class

#Enable Warning BC42016
#Enable Warning BC42017
#Enable Warning BC42019
