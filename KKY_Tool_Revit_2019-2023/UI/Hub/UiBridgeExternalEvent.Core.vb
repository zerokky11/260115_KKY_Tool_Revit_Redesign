Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports Autodesk.Revit.UI

' 네임스페이스는 프로젝트와 일치시켜 주세요.
Namespace UI.Hub

    ' Web(JS) ↔ Revit(VB) 브릿지의 중심. 외부이벤트 큐, 메시지 라우팅, 공통 로그/브로드캐스트 담당
    Partial Public Class UiBridgeExternalEvent

        ' -----------------------------
        ' 상태/싱글톤
        ' -----------------------------
        Friend Shared _host As HubHostWindow ' 다른 partial(Export 등)에서 사용
        Private Shared ReadOnly _self As New UiBridgeExternalEvent() ' 인스턴스 메서드 호출용
        Private Shared ReadOnly _gate As New Object()
        Private Shared ReadOnly _queue As New Queue(Of Action(Of UIApplication))()

        Private Shared _extEv As ExternalEvent
        Private Shared _handler As IExternalEventHandler

        ' -----------------------------
        ' 공개 진입점
        ' -----------------------------

        ''' <summary>
        ''' CmdOpenHub 에서 최초 1회 호출. Host 저장 + ExternalEvent 준비 + 초기 브로드캐스트.
        ''' </summary>
        Public Shared Sub Initialize(host As HubHostWindow)
            _host = host

            If _extEv Is Nothing Then
                _handler = New BridgeHandler(AddressOf ProcessQueue)
                _extEv = ExternalEvent.Create(_handler)
            End If

            ' 초기 상태 브로드캐스트(항상 위, 연결)
            BroadcastTopmost()
            SendToWeb("host:connected", New With {.ok = True})
        End Sub

        ''' <summary>
        ''' Host(또는 Web 메시지 핸들러)에서 호출: 이름/페이로드를 외부이벤트 큐에 넣고 Revit UI 스레드에서 처리.
        ''' </summary>
        Public Shared Sub Raise(name As String, payload As Object)
            Enqueue(Sub(app) Dispatch(app, name, payload))
        End Sub

        ' -----------------------------
        ' 큐/외부이벤트
        ' -----------------------------
        Private Shared Sub Enqueue(work As Action(Of UIApplication))
            SyncLock _gate
                _queue.Enqueue(work)
            End SyncLock
            If _extEv IsNot Nothing Then
                _extEv.Raise()
            End If
        End Sub

        Private Shared Sub ProcessQueue(app As UIApplication)
            Dim todo As Action(Of UIApplication) = Nothing
            Do
                SyncLock _gate
                    If _queue.Count > 0 Then
                        todo = _queue.Dequeue()
                    Else
                        todo = Nothing
                    End If
                End SyncLock
                If todo Is Nothing Then Exit Do
                Try
                    todo(app)
                Catch ex As Exception
                    SendToWeb("host:error", New With {.message = ex.Message})
                End Try
            Loop
        End Sub

        ' -----------------------------
        ' 라우팅
        ' -----------------------------
        Private Shared Sub Dispatch(app As UIApplication, name As String, payload As Object)
            ' 공통(웹 UI 쪽 요청)
            Select Case name
                Case "ui:query-topmost"
                    BroadcastTopmost()
                    Return

                Case "ui:set-topmost"
                    Dim turnOn As Boolean = False
                    Try
                        Dim raw = GetProp(payload, "on")
                        If raw IsNot Nothing Then turnOn = Convert.ToBoolean(raw)
                    Catch
                    End Try
                    Try
                        If _host IsNot Nothing Then _host.Topmost = turnOn
                    Catch
                    End Try
                    BroadcastTopmost()
                    Return

                Case "ui:toggle-topmost"
                    Try
                        If _host IsNot Nothing Then _host.Topmost = Not _host.Topmost
                    Catch
                    End Try
                    BroadcastTopmost()
                    Return
            End Select

            ' 기능 디스패치: 이벤트명 → 내부 핸들러명
            Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            ' Duplicate Inspector
            map.Add("dup:run", "HandleDupRun")
            map.Add("duplicate:export", "HandleDuplicateExport")
            map.Add("duplicate:delete", "HandleDuplicateDelete")
            map.Add("duplicate:restore", "HandleDuplicateRestore")
            map.Add("duplicate:select", "HandleDuplicateSelect")
            ' Connector Diagnostics
            map.Add("connector:run", "HandleConnectorRun")
            map.Add("connector:save-excel", "HandleConnectorSaveExcel")
            ' Export Points with Angle
            map.Add("export:browse-folder", "HandleExportBrowse")
            map.Add("export:add-rvt-files", "HandleExportAddRvtFiles")
            map.Add("export:preview", "HandleExportPreview")
            map.Add("export:save-excel", "HandleExportSaveExcel")
            ' Shared Parameter Propagator
            map.Add("paramprop:run", "HandleSharedParamRun")
            map.Add("sharedparam:run", "HandleSharedParamRun")
            map.Add("sharedparam:list", "HandleSharedParamList")
            map.Add("sharedparam:status", "HandleSharedParamStatus")
            map.Add("sharedparam:export-excel", "HandleSharedParamExport")
            ' 공통 Excel 동작
            map.Add("excel:open", "HandleExcelOpen")
            ' Segment ↔ PMS Check
            map.Add("segmentpms:rvt-pick-files", "HandleSegmentPmsRvtPickFiles")
            map.Add("segmentpms:rvt-pick-folder", "HandleSegmentPmsRvtPickFolder")
            map.Add("segmentpms:extract", "HandleSegmentPmsExtractStart")
            map.Add("segmentpms:load-extract", "HandleSegmentPmsLoadExtract")
            map.Add("segmentpms:save-extract", "HandleSegmentPmsSaveExtract")
            map.Add("segmentpms:register-pms", "HandleSegmentPmsRegisterPms")
            map.Add("segmentpms:pms-template", "HandleSegmentPmsExportTemplate")
            map.Add("segmentpms:prepare-mapping", "HandleSegmentPmsPrepareMapping")
            map.Add("segmentpms:run", "HandleSegmentPmsRun")
            map.Add("segmentpms:save-result", "HandleSegmentPmsSaveResult")
            ' GUID Audit
            map.Add("guid:add-files", "HandleGuidAddFiles")
            map.Add("guid:run", "HandleGuidRun")
            map.Add("guid:export", "HandleGuidExport")
            map.Add("guid:request-family-detail", "HandleGuidRequestFamilyDetail")
            ' Family Link Audit
            map.Add("familylink:init", "HandleFamilyLinkInit")
            map.Add("familylink:pick-rvts", "HandleFamilyLinkPickRvts")
            map.Add("familylink:run", "HandleFamilyLinkRun")
            map.Add("familylink:export", "HandleFamilyLinkExport")
            ' Multi RVT Hub
            map.Add("hub:pick-rvt", "HandleMultiPickRvt")
            map.Add("hub:multi-run", "HandleMultiRun")
            map.Add("hub:multi-export", "HandleMultiExport")
            map.Add("hub:multi-clear", "HandleMultiClear")
            map.Add("commonoptions:get", "HandleCommonOptionsGet")
            map.Add("commonoptions:save", "HandleCommonOptionsSave")

            Dim methodName As String = Nothing
            If Not map.TryGetValue(name, methodName) Then
                SendToWeb("host:warn", New With {.message = String.Format("알 수 없는 이벤트 '{0}'", name)})
                Return
            End If

            ' 동일 Partial 클래스 안의 Private 메서드를 리플렉션으로 찾아 호출
            Dim t As Type = GetType(UiBridgeExternalEvent)
            Dim flags As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public
            Dim m As MethodInfo = t.GetMethod(methodName, flags)

            If m Is Nothing Then
                ' 메서드가 아직 구현되지 않은 경우에도 앱이 죽지 않도록 안전하게 로그만 남김
                SendToWeb("host:warn", New With {.message = String.Format("핸들러 '{0}' 가 구현되어 있지 않습니다.", methodName)})
                Return
            End If

            ' 시그니처: (UIApplication, payload) or (UIApplication) or (payload) or ()
            Dim ps() As ParameterInfo = m.GetParameters()
            Dim args() As Object
            Select Case ps.Length
                Case 2
                    args = New Object() {app, payload}
                Case 1
                    If ps(0).ParameterType Is GetType(UIApplication) Then
                        args = New Object() {app}
                    Else
                        args = New Object() {payload}
                    End If
                Case Else
                    args = New Object() {}
            End Select

            Try
                m.Invoke(_self, args)
            Catch ex As TargetInvocationException
                Dim msg As String
                If ex.InnerException IsNot Nothing Then
                    msg = ex.InnerException.Message
                Else
                    msg = ex.Message
                End If
                SendToWeb("host:error", New With {.message = String.Format("핸들러 실행 오류({0}): {1}", methodName, msg)})
            Catch ex As Exception
                SendToWeb("host:error", New With {.message = String.Format("핸들러 실행 오류({0}): {1}", methodName, ex.Message)})
            End Try
        End Sub

        ' -----------------------------
        ' 공통 유틸/브로드캐스트
        ' -----------------------------

        Friend Shared Sub SendToWeb(channel As String, payload As Object)
            Try
                If _host IsNot Nothing Then
                    _host.SendToWeb(channel, payload)
                End If
            Catch
                ' Host 없는 초기 구간 등에서는 조용히 무시
            End Try
        End Sub

        Private Shared Sub BroadcastTopmost()
            Try
                Dim onTop As Boolean = False
                If _host IsNot Nothing Then onTop = _host.Topmost
                SendToWeb("host:topmost", New With {.on = onTop})
            Catch
            End Try
        End Sub

        Friend Shared Function ParseExcelMode(payload As Object) As Boolean
            Try
                Dim mode As String = Nothing
                If payload IsNot Nothing Then
                    Dim prop = GetProp(payload, "excelMode")
                    If prop IsNot Nothing Then mode = Convert.ToString(prop)
                End If
                If String.Equals(mode, "normal", StringComparison.OrdinalIgnoreCase) Then Return True
            Catch
            End Try
            Return False
        End Function

        ' payload 속성 안전 추출(익명/Dictionary 수용)
        Private Shared Function GetProp(obj As Object, prop As String) As Object
            If obj Is Nothing Then Return Nothing

            Dim d = TryCast(obj, IDictionary(Of String, Object))
            If d IsNot Nothing Then
                Dim v As Object = Nothing
                If d.TryGetValue(prop, v) Then Return v
                Return Nothing
            End If

            Dim t = obj.GetType()
            Dim p = t.GetProperty(prop, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase)
            If p Is Nothing Then Return Nothing
            Return p.GetValue(obj, Nothing)
        End Function

        ' (필요시) 외부에서 직접 로그를 남길 때 사용
        Public Shared Sub HostLog(kind As String, text As String)
            SendToWeb("host:log", New With {.kind = kind, .text = text})
        End Sub

        Friend Shared Sub LogAutoFitDecision(doAutoFit As Boolean, context As String)
            Dim tag As String = If(doAutoFit, "NORMAL_EXPORT: AutoFit applied", "FAST_EXPORT: AutoFit skipped")
            HostLog("debug", $"{tag} [{context}]")
        End Sub

        Private Sub HandleExcelOpen(payload As Object)
            Dim inputPath As String = TryCast(GetProp(payload, "path"), String)
            If String.IsNullOrWhiteSpace(inputPath) Then
                SendToWeb("host:warn", New With {.message = "엑셀 경로가 비어 있습니다."})
                Return
            End If

            Dim fullPath As String = inputPath
            Try
                fullPath = System.IO.Path.GetFullPath(inputPath)
            Catch
                ' ignore
            End Try

            Dim fileExists As Boolean = False
            Try
                fileExists = System.IO.File.Exists(fullPath)
            Catch
                ' ignore
            End Try

            If Not fileExists Then
                SendToWeb("host:warn", New With {.message = "엑셀 파일을 찾을 수 없습니다: " & fullPath, .path = fullPath})
                Return
            End If

            Try
                Dim info As New System.IO.FileInfo(fullPath)
                Dim lengthOk As Boolean = False
                For i As Integer = 0 To 9
                    info.Refresh()
                    If info.Exists AndAlso info.Length > 0 Then
                        lengthOk = True
                        Exit For
                    End If
                    System.Threading.Thread.Sleep(200)
                Next
                If Not lengthOk Then
                    SendToWeb("host:warn", New With {.message = "파일 크기가 0입니다. 열기를 시도합니다: " & fullPath, .path = fullPath})
                End If
            Catch ex As Exception
                SendToWeb("host:warn", New With {.message = "파일 상태 확인에 실패했습니다: " & ex.Message, .path = fullPath})
            End Try

            Dim opened As Boolean = False
            Dim firstError As Exception = Nothing

            Try
                Dim psi As New System.Diagnostics.ProcessStartInfo(fullPath)
                psi.UseShellExecute = True

                Dim directoryPath As String = String.Empty
                Try
                    directoryPath = System.IO.Path.GetDirectoryName(fullPath)
                Catch
                    directoryPath = String.Empty
                End Try
                If Not String.IsNullOrWhiteSpace(directoryPath) Then
                    psi.WorkingDirectory = directoryPath
                End If

                System.Diagnostics.Process.Start(psi)
                opened = True
                SendToWeb("host:info", New With {.message = "엑셀 열기를 시도했습니다: " & fullPath, .path = fullPath})
            Catch ex As Exception
                firstError = ex
            End Try

            If Not opened Then
                Try
                    Dim psi As New System.Diagnostics.ProcessStartInfo("explorer.exe", "/select,""" & fullPath & """")
                    psi.UseShellExecute = True
                    System.Diagnostics.Process.Start(psi)
                    opened = True

                    Dim warnMsg As String = "엑셀 열기에 실패하여 탐색기로 열었습니다: " & fullPath
                    If firstError IsNot Nothing Then
                        warnMsg &= " (" & firstError.Message & ")"
                    End If
                    SendToWeb("host:warn", New With {.message = warnMsg, .path = fullPath})
                Catch ex As Exception
                    Dim msg As String = "엑셀 열기 실패: " & ex.Message
                    If firstError IsNot Nothing Then
                        msg &= " / 최초 오류: " & firstError.Message
                    End If
                    SendToWeb("host:warn", New With {.message = msg, .path = fullPath})
                End Try
            End If
        End Sub

        Private Sub HandleSwitchDocument(app As UIApplication, payload As Object)
            ' 더 이상 문서를 OpenAndActivateDocument로 다시 열지 않는다.
            ' 단순 안내 로그만 남긴다.
            Dim name As String = TryCast(GetProp(payload, "name"), String)
            If String.IsNullOrWhiteSpace(name) Then
                name = TryCast(GetProp(payload, "path"), String)
            End If

            SendToWeb("host:info", New With {
                .message = "문서 전환은 Revit 창에서 직접 선택해 주세요.",
                .target = name
            })
        End Sub

    End Class

    ' 외부이벤트 핸들러(큐를 비우는 역할만 수행)
    Friend Class BridgeHandler
        Implements IExternalEventHandler

        Private ReadOnly _run As Action(Of UIApplication)
        Public Sub New(run As Action(Of UIApplication))
            _run = run
        End Sub

        Public Sub Execute(uiApp As UIApplication) Implements IExternalEventHandler.Execute
            If _run IsNot Nothing Then
                _run.Invoke(uiApp)
            End If
        End Sub

        ' Revit API는 Function GetName() As String 을 요구합니다.
        Public Function GetName() As String Implements IExternalEventHandler.GetName
            Return "KKY Hub Bridge"
        End Function
    End Class

End Namespace
