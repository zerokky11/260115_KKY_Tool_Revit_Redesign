Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

Namespace Services

    Public Class SharedParameterStatus
        Public Property Path As String = ""
        Public Property IsSet As Boolean
        Public Property ExistsOnDisk As Boolean
        Public Property CanOpen As Boolean
        Public Property Status As String = "warn"
        Public Property StatusLabel As String = "설정 필요"
        Public Property WarningMessage As String = ""
        Public Property ErrorMessage As String = ""
    End Class

    Public Class SharedParameterDefinitionItem
        Public Property Name As String = ""
        Public Property Guid As String = ""
        Public Property GroupName As String = ""
        Public Property DataTypeToken As String = ""
    End Class

    Public NotInheritable Class SharedParameterStatusService

        Private Sub New()
        End Sub

        Public Shared Function GetStatus(app As UIApplication) As SharedParameterStatus
            If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))

            Dim status As New SharedParameterStatus()
            Dim path As String = app.Application.SharedParametersFilename
            status.Path = If(path, String.Empty)
            status.IsSet = Not String.IsNullOrWhiteSpace(path)
            status.ExistsOnDisk = status.IsSet AndAlso File.Exists(path)

            If Not status.IsSet Then
                status.Status = "warn"
                status.StatusLabel = "설정 필요"
                status.WarningMessage = "Shared Parameter 파일 경로가 설정되지 않았습니다."
                Return status
            End If

            If Not status.ExistsOnDisk Then
                status.Status = "error"
                status.StatusLabel = "파일 없음"
                status.ErrorMessage = "Shared Parameter 파일을 찾을 수 없습니다."
                Return status
            End If

            Dim defFile As DefinitionFile = Nothing
            Try
                defFile = app.Application.OpenSharedParameterFile()
            Catch ex As Exception
                status.Status = "error"
                status.StatusLabel = "열기 실패"
                status.ErrorMessage = ex.Message
                Return status
            End Try

            status.CanOpen = defFile IsNot Nothing
            If Not status.CanOpen Then
                status.Status = "error"
                status.StatusLabel = "열기 실패"
                status.ErrorMessage = "Shared Parameter 파일을 열 수 없습니다."
                Return status
            End If

            status.Status = "ok"
            status.StatusLabel = "정상"
            Return status
        End Function

        Public Shared Function ListDefinitions(app As UIApplication) As List(Of SharedParameterDefinitionItem)
            If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))

            Dim items As New List(Of SharedParameterDefinitionItem)()
            Dim defFile = app.Application.OpenSharedParameterFile()
            If defFile Is Nothing Then Return items

            For Each grp As DefinitionGroup In defFile.Groups
                If grp Is Nothing Then Continue For
                For Each def As Definition In grp.Definitions
                    If def Is Nothing Then Continue For
                    Dim ext = TryCast(def, ExternalDefinition)
                    Dim guidValue As String = ""
                    Dim dataToken As String = ""
                    If ext IsNot Nothing Then
                        guidValue = ext.GUID.ToString("D")
                        Try
                            Dim dataType = ext.GetDataType()
                            If dataType IsNot Nothing Then dataToken = dataType.TypeId
                        Catch
                        End Try
                        If String.IsNullOrWhiteSpace(dataToken) Then
                            Try
                                dataToken = ext.ParameterType.ToString()
                            Catch
                            End Try
                        End If
                    End If
                    items.Add(New SharedParameterDefinitionItem With {
                        .Name = def.Name,
                        .Guid = guidValue,
                        .GroupName = grp.Name,
                        .DataTypeToken = dataToken
                    })
                Next
            Next

            Return items
        End Function

    End Class

End Namespace
