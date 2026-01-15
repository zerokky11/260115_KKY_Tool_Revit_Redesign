Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports WinForms = System.Windows.Forms

Imports Microsoft.VisualBasic.FileIO

Imports Autodesk.Navisworks.Api
Imports Autodesk.Navisworks.Api.Clash
Imports Autodesk.Navisworks.Api.Plugins

<Plugin("KKY_ClashExcelMapper", "KKY1",
        DisplayName:="Clash CSV Mapper",
        ToolTip:="CSV 매핑(테스트/클래시/ID 기준)으로 Clash Group & Status 자동 적용")>
<AddInPlugin(AddInLocation.AddIn)>
Public Class ClashExcelMapperPlugin
    Inherits AddInPlugin

    Public Overrides Function Execute(ParamArray parameters() As String) As Integer
        Try
            Dim doc As Document = Application.ActiveDocument
            If doc Is Nothing Then
                WinForms.MessageBox.Show("열려 있는 Navisworks 문서가 없습니다.", "Clash CSV Mapper")
                Return 0
            End If

            Dim documentClash As DocumentClash = doc.GetClash()
            If documentClash Is Nothing OrElse documentClash.TestsData Is Nothing Then
                WinForms.MessageBox.Show("이 문서에는 Clash 테스트가 없습니다.", "Clash CSV Mapper")
                Return 0
            End If

            Dim dlg As New WinForms.OpenFileDialog() With {
                .Title = "Clash 매핑 CSV 파일 선택",
                .Filter = "CSV / 텍스트 파일|*.csv;*.txt|모든 파일|*.*"
            }

            If dlg.ShowDialog() <> WinForms.DialogResult.OK Then
                Return 0 ' 취소
            End If

            Dim csvPath As String = dlg.FileName
            If Not File.Exists(csvPath) Then
                WinForms.MessageBox.Show("선택한 파일을 찾을 수 없습니다." & Environment.NewLine & csvPath,
                                         "Clash CSV Mapper")
                Return 0
            End If

            Dim rows As List(Of ClashMappingRow) = ClashCsvReader.LoadFromCsv(csvPath)
            If rows.Count = 0 Then
                WinForms.MessageBox.Show("CSV에서 유효한 매핑 데이터를 찾지 못했습니다.", "Clash CSV Mapper")
                Return 0
            End If

            Dim testsData As DocumentClashTests = documentClash.TestsData

            ' 1) 테스트 컨텍스트(기존 그룹들) + 2) Clash 키 → 위치 매핑 사전 구성
            Dim testsByName As New Dictionary(Of String, TestContext)(StringComparer.OrdinalIgnoreCase)
            Dim resultsByKey As New Dictionary(Of String, ClashResultLocation)(StringComparer.OrdinalIgnoreCase)

            For Each saved As SavedItem In testsData.Tests.Value
                Dim test As ClashTest = TryCast(saved, ClashTest)
                If test Is Nothing Then Continue For

                Dim testKey As String = NormalizeName(test.DisplayName)

                Dim ctx As New TestContext(test)
                testsByName(testKey) = ctx

                ' 기존 그룹 수집
                For Each child As SavedItem In test.Children
                    Dim grp As ClashResultGroup = TryCast(child, ClashResultGroup)
                    If grp IsNot Nothing Then
                        Dim gName As String = grp.DisplayName
                        If Not String.IsNullOrWhiteSpace(gName) Then
                            Dim gKey As String = NormalizeName(gName)
                            If Not ctx.Groups.ContainsKey(gKey) Then
                                ctx.Groups.Add(gKey, grp)
                            End If
                        End If
                    End If
                Next

                ' Clash 결과 위치 수집 (테스트 + 그룹 재귀)
                CollectClashResults(test, test, testsData, resultsByKey)
            Next

            If resultsByKey.Count = 0 Then
                WinForms.MessageBox.Show("Clash 결과가 없습니다.", "Clash CSV Mapper")
                Return 0
            End If

            Dim matched As Integer = 0
            Dim unmatched As Integer = 0
            Dim statusChanged As Integer = 0
            Dim groupChanged As Integer = 0

            For Each row As ClashMappingRow In rows
                Dim testKey As String = NormalizeName(row.TestName)

                Dim ctx As TestContext = Nothing
                If Not testsByName.TryGetValue(testKey, ctx) Then
                    unmatched += 1
                    Continue For
                End If

                Dim clashKey As String = BuildKey(row.TestName, row.ClashName, row.Item1Id, row.Item2Id)

                Dim loc As ClashResultLocation = Nothing
                If Not resultsByKey.TryGetValue(clashKey, loc) Then
                    unmatched += 1
                    Continue For
                End If

                Dim targetParent As GroupItem = ctx.Test

                ' 그룹명이 있으면 그룹 생성/획득
                If Not String.IsNullOrWhiteSpace(row.GroupName) Then
                    Dim grp As ClashResultGroup = GetOrCreateGroup(ctx, testsData, row.GroupName)
                    targetParent = grp
                End If

                Dim clash As ClashResult = MoveResultIfNeeded(testsData, loc, targetParent)
                If Not Object.ReferenceEquals(loc.Parent, targetParent) Then
                    groupChanged += 1
                End If

                Dim statusEnum As ClashResultStatus
                If TryParseStatus(row.StatusText, statusEnum) Then
                    testsData.TestsEditResultStatus(clash, statusEnum)
                    statusChanged += 1
                End If

                matched += 1
            Next

            Dim msg As New StringBuilder()
            msg.AppendLine("CSV 매핑 적용이 완료되었습니다.")
            msg.AppendLine()
            msg.AppendLine("총 매핑 행 수         : " & rows.Count.ToString())
            msg.AppendLine("매칭된 Clash 수       : " & matched.ToString())
            msg.AppendLine("매칭 실패(무시) 행 수 : " & unmatched.ToString())
            msg.AppendLine("그룹 이동/생성 수     : " & groupChanged.ToString())
            msg.AppendLine("Status 변경 수        : " & statusChanged.ToString())

            WinForms.MessageBox.Show(msg.ToString(), "Clash CSV Mapper")

        Catch ex As Exception
            WinForms.MessageBox.Show("Clash CSV Mapper 실행 중 오류가 발생했습니다." &
                                     Environment.NewLine & ex.Message,
                                     "Clash CSV Mapper")
        End Try

        Return 0
    End Function

    '=========================
    ' Clash 결과 수집 (재귀)
    '=========================
    Private Shared Sub CollectClashResults(parent As GroupItem,
                                           test As ClashTest,
                                           testsData As DocumentClashTests,
                                           resultsByKey As Dictionary(Of String, ClashResultLocation))

        Dim children As SavedItemCollection = parent.Children

        For i As Integer = 0 To children.Count - 1
            Dim child As SavedItem = children(i)

            Dim clash As ClashResult = TryCast(child, ClashResult)
            If clash IsNot Nothing Then
                Dim id1 As String = GetRevitElementIdString(clash.Item1)
                Dim id2 As String = GetRevitElementIdString(clash.Item2)

                If String.IsNullOrEmpty(id1) OrElse String.IsNullOrEmpty(id2) Then
                    Continue For
                End If

                Dim key As String = BuildKey(test.DisplayName, clash.DisplayName, id1, id2)

                If Not resultsByKey.ContainsKey(key) Then
                    Dim loc As New ClashResultLocation(test, parent, i, clash)
                    resultsByKey.Add(key, loc)
                End If
            Else
                Dim grp As ClashResultGroup = TryCast(child, ClashResultGroup)
                If grp IsNot Nothing Then
                    CollectClashResults(grp, test, testsData, resultsByKey)
                End If
            End If
        Next
    End Sub

    '=========================
    ' Revit Element Id (Element/Id) → 문자열
    '=========================
    Private Shared Function GetRevitElementIdString(item As ModelItem) As String
        If item Is Nothing Then Return String.Empty

        Dim cat As PropertyCategory =
            item.PropertyCategories.FindCategoryByDisplayName("Element")

        If cat Is Nothing Then Return String.Empty

        Dim prop As DataProperty =
            cat.Properties.FindPropertyByDisplayName("Id")

        If prop Is Nothing OrElse prop.Value Is Nothing Then
            Return String.Empty
        End If

        Try
            If prop.Value.IsInteger Then
                Return prop.Value.ToInt32().ToString()
            End If

            Dim s As String = prop.Value.ToDisplayString()
            If String.IsNullOrWhiteSpace(s) Then Return String.Empty

            Dim n As Long
            If Long.TryParse(s.Trim(), n) Then
                Return n.ToString()
            End If

            Return s.Trim()
        Catch
            Return String.Empty
        End Try
    End Function

    '=========================
    ' 키 생성 (Test + Clash + 두 ID 정렬)
    '=========================
    Private Shared Function BuildKey(testName As String,
                                     clashName As String,
                                     item1Id As String,
                                     item2Id As String) As String
        Dim t As String = NormalizeName(testName)
        Dim c As String = NormalizeName(clashName)
        Dim a As String = item1Id.Trim()
        Dim b As String = item2Id.Trim()

        ' ID 두 개는 순서 상관없이 정렬
        If String.CompareOrdinal(a, b) > 0 Then
            Dim tmp As String = a
            a = b
            b = tmp
        End If

        Return String.Format("{0}||{1}||{2}||{3}", t, c, a, b)
    End Function

    Private Shared Function NormalizeName(name As String) As String
        If name Is Nothing Then Return String.Empty
        Return name.Trim().ToUpperInvariant()
    End Function

    '=========================
    ' 그룹 가져오거나 새로 생성
    '=========================
    Private Shared Function GetOrCreateGroup(ctx As TestContext,
                                             testsData As DocumentClashTests,
                                             groupName As String) As ClashResultGroup

        Dim key As String = NormalizeName(groupName)

        Dim existing As ClashResultGroup = Nothing
        If ctx.Groups.TryGetValue(key, existing) Then
            Return existing
        End If

        Dim temp As New ClashResultGroup()
        temp.DisplayName = groupName

        testsData.TestsAddCopy(ctx.Test, temp)

        Dim children As SavedItemCollection = ctx.Test.Children
        Dim newGrp As ClashResultGroup = Nothing

        If children.Count > 0 Then
            newGrp = TryCast(children(children.Count - 1), ClashResultGroup)
        End If

        If newGrp Is Nothing Then
            Throw New InvalidOperationException(
                String.Format("테스트 '{0}' 에 그룹 '{1}' 를 생성하지 못했습니다.",
                              ctx.Test.DisplayName, groupName))
        End If

        ctx.Groups(key) = newGrp
        Return newGrp
    End Function

    '=========================
    ' 결과를 필요한 부모로 이동 (필요시만)
    '=========================
    Private Shared Function MoveResultIfNeeded(testsData As DocumentClashTests,
                                               loc As ClashResultLocation,
                                               newParent As GroupItem) As ClashResult
        Dim currentParent As GroupItem = loc.Parent
        Dim currentResult As ClashResult = loc.Result

        If Object.ReferenceEquals(currentParent, newParent) Then
            Return currentResult
        End If

        testsData.TestsAddCopy(newParent, currentResult)

        Dim newChildren As SavedItemCollection = newParent.Children
        Dim newResult As ClashResult =
            TryCast(newChildren(newChildren.Count - 1), ClashResult)

        testsData.TestsRemove(currentParent, currentResult)

        loc.Parent = newParent
        loc.Index = newChildren.Count - 1
        loc.Result = newResult

        Return newResult
    End Function

    '=========================
    ' Status 텍스트 → ClashResultStatus
    '=========================
    Private Shared Function TryParseStatus(text As String,
                                           ByRef status As ClashResultStatus) As Boolean
        If String.IsNullOrWhiteSpace(text) Then Return False

        Select Case text.Trim().ToLowerInvariant()
            Case "new"
                status = ClashResultStatus.New
            Case "active"
                status = ClashResultStatus.Active
            Case "reviewed"
                status = ClashResultStatus.Reviewed
            Case "approved"
                status = ClashResultStatus.Approved
            Case "resolved"
                status = ClashResultStatus.Resolved
            Case Else
                Return False
        End Select

        Return True
    End Function

End Class

'===============================
' 보조 클래스들
'===============================
Friend Class ClashMappingRow
    Public Property TestName As String
    Public Property ClashName As String
    Public Property Item1Id As String
    Public Property Item2Id As String
    Public Property GroupName As String
    Public Property StatusText As String
End Class

Friend Class ClashResultLocation
    Public Property Test As ClashTest
    Public Property Parent As GroupItem
    Public Property Index As Integer
    Public Property Result As ClashResult

    Public Sub New(test As ClashTest,
                   parent As GroupItem,
                   index As Integer,
                   result As ClashResult)
        Me.Test = test
        Me.Parent = parent
        Me.Index = index
        Me.Result = result
    End Sub
End Class

Friend Class TestContext
    Public ReadOnly Property Test As ClashTest
    Public ReadOnly Property Groups As Dictionary(Of String, ClashResultGroup)

    Public Sub New(test As ClashTest)
        Me.Test = test
        Me.Groups = New Dictionary(Of String, ClashResultGroup)(StringComparer.OrdinalIgnoreCase)
    End Sub
End Class

'===============================
' CSV Reader (엑셀을 CSV로 저장해서 사용)
'===============================
Friend NotInheritable Class ClashCsvReader
    Private Sub New()
    End Sub

    Public Shared Function LoadFromCsv(path As String) As List(Of ClashMappingRow)
        Dim list As New List(Of ClashMappingRow)()

        If Not File.Exists(path) Then
            Return list
        End If

        ' 기본 UTF-8 가정 (필요하면 Encoding.Default 로 변경)
        Using parser As New TextFieldParser(path, Encoding.UTF8)
            parser.TextFieldType = FieldType.Delimited
            ' 쉼표, 탭 둘 다 허용
            parser.SetDelimiters(",", vbTab)
            parser.HasFieldsEnclosedInQuotes = True

            If parser.EndOfData Then
                Return list
            End If

            ' 헤더 읽기
            Dim headerFields() As String = parser.ReadFields()
            If headerFields Is Nothing OrElse headerFields.Length = 0 Then
                Return list
            End If

            Dim colMap As Dictionary(Of String, Integer) = BuildHeaderMap(headerFields)

            While Not parser.EndOfData
                Dim fields() As String = Nothing
                Try
                    fields = parser.ReadFields()
                Catch ex As MalformedLineException
                    ' 깨진 라인은 무시
                    Continue While
                End Try

                If fields Is Nothing OrElse fields.Length = 0 Then
                    Continue While
                End If

                Dim testName As String = GetField(fields, colMap, "TEST NAME")
                Dim clashName As String = GetField(fields, colMap, "CLASH NAME")
                Dim item1Id As String = GetField(fields, colMap, "ITEM1 ID")
                Dim item2Id As String = GetField(fields, colMap, "ITEM2 ID")
                Dim groupName As String = GetField(fields, colMap, "CLASH GROUP")
                Dim statusText As String = GetField(fields, colMap, "STATUS")

                If String.IsNullOrWhiteSpace(testName) OrElse
                   String.IsNullOrWhiteSpace(clashName) OrElse
                   String.IsNullOrWhiteSpace(item1Id) OrElse
                   String.IsNullOrWhiteSpace(item2Id) Then

                    Continue While
                End If

                Dim item As New ClashMappingRow() With {
                    .TestName = testName,
                    .ClashName = clashName,
                    .Item1Id = item1Id,
                    .Item2Id = item2Id,
                    .GroupName = groupName,
                    .StatusText = statusText
                }

                list.Add(item)
            End While
        End Using

        Return list
    End Function

    Private Shared Function BuildHeaderMap(headerFields() As String) As Dictionary(Of String, Integer)
        Dim map As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

        For i As Integer = 0 To headerFields.Length - 1
            Dim text As String = headerFields(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For

            Dim key As String = text.Trim().ToUpperInvariant()
            If Not map.ContainsKey(key) Then
                map.Add(key, i)
            End If
        Next

        Return map
    End Function

    Private Shared Function GetField(fields() As String,
                                     colMap As Dictionary(Of String, Integer),
                                     logicalName As String) As String
        Dim idx As Integer
        If Not colMap.TryGetValue(logicalName.ToUpperInvariant(), idx) Then
            Return String.Empty
        End If

        If idx < 0 OrElse idx >= fields.Length Then
            Return String.Empty
        End If

        Dim value As String = fields(idx)
        If value Is Nothing Then Return String.Empty
        Return value.Trim()
    End Function

End Class
