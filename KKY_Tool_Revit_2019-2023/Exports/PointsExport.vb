Imports System.Data
Imports KKY_Tool_Revit.Infrastructure

Namespace Exports

    Public Module PointsExport

        ' 저장 대화상자 사용
        Public Function SaveWithDialog(resultTable As DataTable) As String
            If resultTable Is Nothing Then Return String.Empty
            Return ExcelCore.PickAndSaveXlsx("Exported Points", resultTable, "ExportPoints.xlsx")
        End Function

        ' 경로 지정 저장
        Public Sub Save(outPath As String, resultTable As DataTable)
            If resultTable Is Nothing Then Exit Sub
            ' 일부 레거시에서는 SaveXlsx(filePath, table) 시그니처를 사용함 → ExcelCore 오버로드로 흡수
            ExcelCore.SaveXlsx(outPath, "Exported Points", resultTable)
        End Sub

    End Module
End Namespace
