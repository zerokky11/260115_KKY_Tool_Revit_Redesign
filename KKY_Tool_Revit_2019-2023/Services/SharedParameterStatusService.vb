Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection
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

            status.CanOpen = (defFile IsNot Nothing)
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

            Dim defFile As DefinitionFile = Nothing
            Try
                defFile = app.Application.OpenSharedParameterFile()
            Catch
                Return items
            End Try
            If defFile Is Nothing Then Return items

            For Each grp As DefinitionGroup In defFile.Groups
                If grp Is Nothing Then Continue For

                For Each defn As Definition In grp.Definitions
                    If defn Is Nothing Then Continue For

                    Dim guidValue As String = ""
                    Dim ext = TryCast(defn, ExternalDefinition)
                    If ext IsNot Nothing Then
                        Try
                            guidValue = ext.GUID.ToString("D")
                        Catch
                            guidValue = ""
                        End Try
                    End If

                    Dim dataToken As String = TryGetDefinitionDataTypeToken(defn)

                    items.Add(New SharedParameterDefinitionItem With {
                        .Name = defn.Name,
                        .Guid = guidValue,
                        .GroupName = grp.Name,
                        .DataTypeToken = dataToken
                    })
                Next
            Next

            Return items
        End Function

        ' Revit 버전별 API 차이를 리플렉션으로 안전하게 흡수
        ' - Revit 2023+: Definition.GetDataType() -> ForgeTypeId.TypeId
        ' - Revit 2022- : Definition.ParameterType
        Private Shared Function TryGetDefinitionDataTypeToken(defn As Definition) As String
            If defn Is Nothing Then Return ""

            ' 1) Revit 2023+ : GetDataType()
            Try
                Dim m As MethodInfo = defn.GetType().GetMethod("GetDataType", Type.EmptyTypes)
                If m IsNot Nothing Then
                    Dim dtObj As Object = m.Invoke(defn, Nothing)
                    If dtObj IsNot Nothing Then
                        Dim pTypeId As PropertyInfo = dtObj.GetType().GetProperty("TypeId")
                        If pTypeId IsNot Nothing Then
                            Dim v As Object = pTypeId.GetValue(dtObj, Nothing)
                            Dim s As String = TryCast(v, String)
                            If Not String.IsNullOrWhiteSpace(s) Then Return s
                        End If
                    End If
                End If
            Catch
                ' ignore
            End Try

            ' 2) Revit 2022- : ParameterType
            Try
                Dim p As PropertyInfo = defn.GetType().GetProperty("ParameterType")
                If p IsNot Nothing Then
                    Dim v As Object = p.GetValue(defn, Nothing)
                    If v IsNot Nothing Then Return v.ToString()
                End If
            Catch
                ' ignore
            End Try

            Return ""
        End Function

    End Class

End Namespace
