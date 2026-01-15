Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Windows
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.Wpf
Imports System.Web.Script.Serialization
Imports System.Reflection
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB

Namespace UI.Hub
    Public Class HubHostWindow
        Inherits Window

        Private Const BaseTitle As String = "KKY Tool Hub"

        Private Shared _instance As HubHostWindow
        Private Shared ReadOnly _gate As New Object()

        Private ReadOnly _web As New WebView2()
        Private ReadOnly _serializer As New JavaScriptSerializer()

        Private _uiApp As UIApplication
        Private _currentDocName As String = String.Empty
        Private _currentDocPath As String = String.Empty
        Private _initStarted As Boolean = False
        Private _isClosing As Boolean = False

        Public ReadOnly Property Web As WebView2
            Get
                Return _web
            End Get
        End Property

        Public ReadOnly Property IsClosing As Boolean
            Get
                Return _isClosing
            End Get
        End Property

        Public Shared Sub ShowSingleton(uiApp As UIApplication)
            If uiApp Is Nothing Then Return

            SyncLock _gate
                If _instance IsNot Nothing AndAlso Not _instance.IsClosing Then
                    _instance.AttachTo(uiApp)
                    UiBridgeExternalEvent.Initialize(_instance)
                    If _instance.WindowState = WindowState.Minimized Then
                        _instance.WindowState = WindowState.Normal
                    End If
                    _instance.Activate()
                    _instance.Focus()
                    Return
                End If

                Dim wnd As New HubHostWindow(uiApp)
                UiBridgeExternalEvent.Initialize(wnd)
                _instance = wnd
                wnd.Show()
            End SyncLock
        End Sub

        Public Shared Sub NotifyActiveDocumentChanged(doc As Document)
            Dim inst = _instance
            If inst Is Nothing OrElse inst.IsClosing Then Return
            inst.UpdateActiveDocument(doc)
        End Sub

        Public Shared Sub NotifyDocumentListChanged()
            Dim inst = _instance
            If inst Is Nothing OrElse inst.IsClosing Then Return
            inst.BroadcastDocumentList()
        End Sub

        Public Sub New(uiApp As UIApplication)
            _uiApp = uiApp
            Title = BaseTitle
            Dim workArea = System.Windows.SystemParameters.WorkArea
            Dim desiredWidth As Double = 1400
            Dim desiredHeight As Double = 900
            Width = Math.Min(desiredWidth, workArea.Width * 0.93)
            Height = Math.Min(desiredHeight, workArea.Height * 0.93)
            MinWidth = Math.Min(1100, Width)
            MinHeight = Math.Min(720, Height)
            WindowStartupLocation = WindowStartupLocation.CenterScreen
            Content = _web

            AddHandler Loaded, AddressOf OnLoaded
            AddHandler Closing, AddressOf OnWindowClosing
            AddHandler Closed, AddressOf OnWindowClosed

            UpdateActiveDocument(GetActiveDocument())
        End Sub

        Public Sub AttachTo(uiApp As UIApplication)
            _uiApp = uiApp
            UpdateActiveDocument(GetActiveDocument())
            BroadcastDocumentList()
        End Sub

        Private Function GetActiveDocument() As Document
            Try
                If _uiApp Is Nothing Then Return Nothing
                Dim uidoc = _uiApp.ActiveUIDocument
                If uidoc Is Nothing Then Return Nothing
                Return uidoc.Document
            Catch
                Return Nothing
            End Try
        End Function

        Private Sub UpdateActiveDocument(doc As Document)
            Dim name As String = String.Empty
            Dim path As String = String.Empty

            If doc IsNot Nothing Then
                Try
                    name = doc.Title
                Catch
                End Try

                Try
                    path = doc.PathName
                Catch
                End Try
            End If

            If String.IsNullOrWhiteSpace(path) Then
                path = name
            End If

            _currentDocName = name
            _currentDocPath = path

            UpdateWindowTitle()
            SendActiveDocument()
        End Sub

        Private Sub UpdateWindowTitle()
            If String.IsNullOrWhiteSpace(_currentDocName) Then
                Title = BaseTitle
            Else
                Title = $"{BaseTitle} - {_currentDocName}"
            End If
        End Sub

        Private Function ResolveUiFolder() As String
            Try
                Dim asm = Assembly.GetExecutingAssembly()
                Dim baseDir = Path.GetDirectoryName(asm.Location)
                Dim ui = Path.Combine(baseDir, "Resources", "HubUI")
                If Directory.Exists(ui) Then Return Path.GetFullPath(ui)
            Catch
            End Try
            Return Nothing
        End Function

        Private Async Sub OnLoaded(sender As Object, e As RoutedEventArgs)
            If _initStarted Then Return
            _initStarted = True
            Try
                Dim userData = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "KKY_Tool_Revit", "WebView2UserData")
                Directory.CreateDirectory(userData)

                Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userData, Nothing)
                Await _web.EnsureCoreWebView2Async(env)
                Dim core = _web.CoreWebView2

                core.Settings.AreDefaultContextMenusEnabled = False
                core.Settings.IsStatusBarEnabled = False
#If DEBUG Then
                core.Settings.AreDevToolsEnabled = True
#Else
                core.Settings.AreDevToolsEnabled = False
#End If

                ' 가상 호스트 매핑
                Dim uiFolder = ResolveUiFolder()
                If String.IsNullOrEmpty(uiFolder) Then
                    Throw New DirectoryNotFoundException("Resources\HubUI 폴더를 찾을 수 없습니다.")
                End If
                core.SetVirtualHostNameToFolderMapping(
                    "hub.local", uiFolder, CoreWebView2HostResourceAccessKind.Allow)

                ' 메시지 브리지
                AddHandler core.WebMessageReceived, AddressOf OnWebMessage

                ' 허브 진입
                _web.Source = New Uri("https://hub.local/index.html")

                ' 초기 상태 알림
                SendToWeb("host:topmost", New With {.on = Me.Topmost})
                SendActiveDocument()
                BroadcastDocumentList()

            Catch ex As Exception
                Dim hr As Integer = Runtime.InteropServices.Marshal.GetHRForException(ex)
                MessageBox.Show($"WebView 초기화 실패 (0x{hr:X8}) : {ex.Message}",
                                "KKY Tool", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub

        Private Sub OnWindowClosing(sender As Object, e As System.ComponentModel.CancelEventArgs)
            _isClosing = True
        End Sub

        Private Sub OnWindowClosed(sender As Object, e As EventArgs)
            SyncLock _gate
                If _instance Is Me Then
                    _instance = Nothing
                End If
            End SyncLock
        End Sub

        Private Sub OnWebMessage(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
            Try
                Dim root As Dictionary(Of String, Object) =
                    _serializer.Deserialize(Of Dictionary(Of String, Object))(e.WebMessageAsJson)

                Dim name As String = Nothing
                If root IsNot Nothing Then
                    If root.ContainsKey("ev") AndAlso root("ev") IsNot Nothing Then
                        name = Convert.ToString(root("ev"))
                    ElseIf root.ContainsKey("name") AndAlso root("name") IsNot Nothing Then
                        name = Convert.ToString(root("name"))
                    End If
                End If
                If String.IsNullOrEmpty(name) Then Return

                Dim payload As Object = Nothing
                If root IsNot Nothing AndAlso root.ContainsKey("payload") Then
                    payload = root("payload")
                End If

                Select Case name
                    Case "ui:ping"
                        SendToWeb("host:pong", New With {.t = Date.Now.Ticks})

                    Case "ui:toggle-topmost"
                        Me.Topmost = Not Me.Topmost
                        SendToWeb("host:topmost", New With {.on = Me.Topmost})

                    Case "ui:query-topmost"
                        SendToWeb("host:topmost", New With {.on = Me.Topmost})

                    Case Else
                        ' 나머지는 ExternalEvent로 위임
                        UiBridgeExternalEvent.Raise(name, payload)
                End Select

            Catch ex As Exception
                SendToWeb("host:error", New With {ex.Message})
            End Try
        End Sub

        Private Sub BroadcastDocumentList()
            Dim docs As New List(Of Object)()
            Try
                If _uiApp IsNot Nothing AndAlso _uiApp.Application IsNot Nothing Then
                    For Each d As Document In _uiApp.Application.Documents
                        Try
                            Dim name = d.Title
                            Dim path = d.PathName
                            If String.IsNullOrWhiteSpace(path) Then
                                path = name
                            End If
                            docs.Add(New With {.name = name, path})
                        Catch
                        End Try
                    Next
                End If
            Catch
            End Try

            SendToWeb("host:doc-list", docs)
        End Sub

        Private Sub SendActiveDocument()
            SendToWeb("host:doc-changed", New With {.name = _currentDocName, .path = _currentDocPath})
        End Sub

        ' .NET → JS (양쪽 호환: ev & name 둘 다 포함해서 송신)
        Public Sub SendToWeb(ev As String, payload As Object)
            Dim core = _web.CoreWebView2
            If core Is Nothing Then Return
            Dim msg As New Dictionary(Of String, Object) From {
                {"ev", ev}, {"name", ev}, {"payload", payload}
            }
            Dim json = _serializer.Serialize(msg)
            core.PostWebMessageAsJson(json)
        End Sub
    End Class
End Namespace
