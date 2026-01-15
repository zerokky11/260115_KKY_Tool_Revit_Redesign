Option Explicit On
Option Strict On

Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Services

Namespace Exports

    Public Module FamilyLinkAuditExport

        Public ReadOnly Schema As String() = {
            "FileName",
            "HostFamilyName",
            "HostFamilyCategory",
            "NestedFamilyName",
            "NestedTypeName",
            "NestedCategory",
            "TargetParamName",
            "ExpectedGuid",
            "FoundScope",
            "NestedParamGuid",
            "NestedParamDataType",
            "AssocHostParamName",
            "HostParamGuid",
            "HostParamIsShared",
            "Issue",
            "Notes"
        }

        Public Function Export(rows As IEnumerable(Of FamilyLinkAuditRow), Optional fastExport As Boolean = True, Optional autoFit As Boolean = False) As String
            If rows Is Nothing Then Return String.Empty
            Dim table As DataTable = BuildTable(rows)
            If Not ValidateSchema(table) Then
                Throw New InvalidOperationException("스키마 검증 실패: 컬럼 순서/헤더가 규격과 다릅니다.")
            End If

            Dim defaultName As String = $"FamilyLinkAudit_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            Using dlg As New SaveFileDialog()
                dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv"
                dlg.FileName = defaultName
                dlg.AddExtension = True
                dlg.DefaultExt = "xlsx"
                dlg.OverwritePrompt = True
                dlg.RestoreDirectory = True

                If dlg.ShowDialog() <> DialogResult.OK Then Return String.Empty

                Dim filePath As String = dlg.FileName
                Dim ext As String = Path.GetExtension(filePath).ToLowerInvariant()
                If String.IsNullOrWhiteSpace(ext) Then
                    ext = If(dlg.FilterIndex = 2, ".csv", ".xlsx")
                    filePath = filePath & ext
                End If

                If ext = ".csv" Then
                    SaveCsv(filePath, table)
                Else
                    Dim doAutoFit As Boolean = (Not fastExport) AndAlso autoFit
                    ExcelCore.SaveXlsx(filePath, "FamilyLinkAudit", table, doAutoFit)
                End If
                Return filePath
            End Using
        End Function

        Private Function BuildTable(rows As IEnumerable(Of FamilyLinkAuditRow)) As DataTable
            Dim dt As New DataTable("FamilyLinkAudit")
            For Each h In Schema
                dt.Columns.Add(h)
            Next

            For Each r In rows
                Dim dr = dt.NewRow()
                dr("FileName") = SafeStr(r.FileName)
                dr("HostFamilyName") = SafeStr(r.HostFamilyName)
                dr("HostFamilyCategory") = SafeStr(r.HostFamilyCategory)
                dr("NestedFamilyName") = SafeStr(r.NestedFamilyName)
                dr("NestedTypeName") = SafeStr(r.NestedTypeName)
                dr("NestedCategory") = SafeStr(r.NestedCategory)
                dr("TargetParamName") = SafeStr(r.TargetParamName)
                dr("ExpectedGuid") = SafeStr(r.ExpectedGuid)
                dr("FoundScope") = SafeStr(r.FoundScope)
                dr("NestedParamGuid") = SafeStr(r.NestedParamGuid)
                dr("NestedParamDataType") = SafeStr(r.NestedParamDataType)
                dr("AssocHostParamName") = SafeStr(r.AssocHostParamName)
                dr("HostParamGuid") = SafeStr(r.HostParamGuid)
                dr("HostParamIsShared") = SafeStr(r.HostParamIsShared)
                dr("Issue") = SafeStr(r.Issue)
                dr("Notes") = SafeStr(r.Notes)
                dt.Rows.Add(dr)
            Next

            Return dt
        End Function

        Private Function ValidateSchema(table As DataTable) As Boolean
            If table Is Nothing Then Return False
            If table.Columns.Count <> Schema.Length Then Return False
            For i As Integer = 0 To Schema.Length - 1
                If Not String.Equals(table.Columns(i).ColumnName, Schema(i), StringComparison.Ordinal) Then
                    Return False
                End If
            Next
            Return True
        End Function

        Private Sub SaveCsv(filePath As String, table As DataTable)
            Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.Read)
                Using sw As New StreamWriter(fs, New UTF8Encoding(True))
                    sw.WriteLine(String.Join(",", Schema.Select(Function(h) CsvEscape(h))))
                    For Each row As DataRow In table.Rows
                        Dim cols As New List(Of String)()
                        For Each h In Schema
                            cols.Add(CsvEscape(If(row(h), "").ToString()))
                        Next
                        sw.WriteLine(String.Join(",", cols))
                    Next
                End Using
            End Using
        End Sub

        Private Function CsvEscape(s As String) As String
            If s Is Nothing Then s = ""
            Dim quote As String = """"
            Dim needsQuotes As Boolean = s.Contains(","c) OrElse s.Contains(quote) OrElse s.Contains(vbCr) OrElse s.Contains(vbLf)
            s = s.Replace(quote, quote & quote)
            If needsQuotes Then
                Return quote & s & quote
            End If
            Return s
        End Function

        Private Function SafeStr(s As String) As String
            Return If(s, "")
        End Function

    End Module

End Namespace
