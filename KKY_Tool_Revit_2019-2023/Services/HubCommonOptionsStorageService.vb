Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Web.Script.Serialization

Namespace Services
    Public NotInheritable Class HubCommonOptionsStorageService
        Private Sub New()
        End Sub

        Public Class HubCommonOptions
            Public Property ExtraParamsText As String = String.Empty
            Public Property TargetFilterText As String = String.Empty
            Public Property ExcludeEndDummy As Boolean
        End Class

        Public Shared Function Load() As HubCommonOptions
            Dim result As New HubCommonOptions()
            Try
                Dim path = GetOptionsPath()
                If Not File.Exists(path) Then
                    Return result
                End If
                Dim json = File.ReadAllText(path)
                If String.IsNullOrWhiteSpace(json) Then
                    Return result
                End If
                Dim serializer As New JavaScriptSerializer()
                Dim loaded = serializer.Deserialize(Of HubCommonOptions)(json)
                If loaded IsNot Nothing Then
                    result = loaded
                End If
            Catch
                Return result
            End Try
            Return result
        End Function

        Public Shared Function Save(options As HubCommonOptions) As Boolean
            If options Is Nothing Then Return False
            Try
                Dim path = GetOptionsPath()
                Dim dir = IO.Path.GetDirectoryName(path)
                If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If
                Dim serializer As New JavaScriptSerializer()
                Dim json = serializer.Serialize(options)
                File.WriteAllText(path, json)
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function GetOptionsPath() As String
            Dim root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            If String.IsNullOrWhiteSpace(root) Then
                root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            End If
            Dim dir = Path.Combine(root, "KKY_Tool_Revit")
            Return Path.Combine(dir, "hub-common-options.json")
        End Function
    End Class
End Namespace
