Option Explicit On
Option Strict On

Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.UI.Hub
' ʿϸ ӽ̽  :
'Imports KKY_Tool_Revit.UI.Hub
' Ǵ Ʈ UiBridgeExternalEvent  Ʈ ӽ̽ Imports

<Transaction(TransactionMode.Manual)>
<Regeneration(RegenerationOption.Manual)>
Public Class DuplicateExport
    Implements IExternalCommand

    Public Function Execute(commandData As ExternalCommandData,
                            ByRef message As String,
                            elements As ElementSet) As Result Implements IExternalCommand.Execute
        Try
            Dim uiapp = commandData.Application

            ' 허브는 싱글톤으로 관리한다.
            HubHostWindow.ShowSingleton(uiapp)
            Return Result.Succeeded

        Catch ex As Exception
            message = ex.Message
            Return Result.Failed
        End Try
    End Function

End Class
