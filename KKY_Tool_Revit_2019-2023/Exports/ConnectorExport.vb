Imports System.Data
Imports KKY_Tool_Revit.Infrastructure

Namespace Exports

    Public Module ConnectorExport

        ' 저장 대화 상자 사용 버전(선택시 호출)
        Public Function SaveWithDialog(resultTable As DataTable) As String
            If resultTable Is Nothing Then Return String.Empty
            Return ExcelCore.PickAndSaveXlsx("Connector Diagnostics", resultTable, "ConnectorDiagnostics.xlsx")
        End Function

        ' 경로 고정 저장 버전(브리지에서 경로를 이미 받은 경우)
        Public Sub Save(outPath As String, resultTable As DataTable)
            If resultTable Is Nothing Then Exit Sub
            ExcelCore.SaveXlsx(outPath, "Connector Diagnostics", resultTable)
        End Sub

    End Module
End Namespace
