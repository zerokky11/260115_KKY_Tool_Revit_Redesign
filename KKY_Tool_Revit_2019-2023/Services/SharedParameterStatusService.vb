Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports Autodesk.Revit.UI

Namespace Services
    Public NotInheritable Class SharedParameterStatusService
        Private Sub New()
        End Sub

        Public Class SharedParameterStatus
            Public Property Path As String = String.Empty
            Public Property IsSet As Boolean
            Public Property ExistsOnDisk As Boolean
            Public Property CanOpen As Boolean
            Public Property ErrorMessage As String = String.Empty
            Public Property Status As String = "unknown"
            Public Property StatusLabel As String = "알 수 없음"
            Public Property WarningMessage As String = String.Empty
        End Class

        Public Shared Function GetStatus(app As UIApplication) As SharedParameterStatus
            Dim result As New SharedParameterStatus()
            If app Is Nothing OrElse app.Application Is Nothing Then
                result.Status = "error"
                result.StatusLabel = "조회 실패"
                result.WarningMessage = "Revit Application이 준비되지 않았습니다."
                Return result
            End If

            Dim path As String = app.Application.SharedParametersFilename
            result.Path = If(path, String.Empty)
            result.IsSet = Not String.IsNullOrWhiteSpace(path)
            result.ExistsOnDisk = result.IsSet AndAlso File.Exists(path)

            If Not result.IsSet Then
                result.Status = "unset"
                result.StatusLabel = "미설정(등록 필요)"
                result.WarningMessage = "Shared Parameter 텍스트 파일이 설정되지 않았습니다."
                Return result
            End If

            If Not result.ExistsOnDisk Then
                result.Status = "missing"
                result.StatusLabel = "경로는 있으나 파일 없음"
                result.WarningMessage = "Shared Parameter 파일 경로가 존재하지 않습니다."
                Return result
            End If

            Try
                Dim defFile = app.Application.OpenSharedParameterFile()
                If defFile Is Nothing Then
                    result.CanOpen = False
                    result.Status = "open_failed"
                    result.StatusLabel = "파일 열기 실패"
                    result.WarningMessage = "Shared Parameter 파일을 열 수 없습니다."
                Else
                    result.CanOpen = True
                    result.Status = "ok"
                    result.StatusLabel = "정상(등록됨)"
                End If
            Catch ex As Exception
                result.CanOpen = False
                result.Status = "open_failed"
                result.StatusLabel = "파일 열기 실패"
                result.ErrorMessage = ex.Message
                result.WarningMessage = "Shared Parameter 파일을 열 수 없습니다."
            End Try

            Return result
        End Function
    End Class
End Namespace
