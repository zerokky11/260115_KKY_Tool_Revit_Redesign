Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports System.Windows.Forms
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Services
Imports KKY_Tool_Revit.Exports

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Class MultiCommonOptions
            Public Property ExtraParams As String = String.Empty
            Public Property TargetFilter As String = String.Empty
            Public Property ExcludeEndDummy As Boolean
        End Class

        ' === commonoptions:get ===
        Private Sub HandleCommonOptionsGet(app As UIApplication, payload As Object)
            Try
                Dim stored = HubCommonOptionsStorageService.Load()
                SendToWeb("commonoptions:loaded", New With {
                    .extraParamsText = stored.ExtraParamsText,
                    .targetFilterText = stored.TargetFilterText,
                    .excludeEndDummy = stored.ExcludeEndDummy
                })
            Catch ex As Exception
                SendToWeb("commonoptions:loaded", New With {
                    .extraParamsText = "",
                    .targetFilterText = "",
                    .excludeEndDummy = False,
                    .errorMessage = ex.Message
                })
            End Try
        End Sub

        ' === commonoptions:save ===
        Private Sub HandleCommonOptionsSave(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim extraText As String = Convert.ToString(GetProp(pd, "extraParamsText"))
            Dim filterText As String = Convert.ToString(GetProp(pd, "targetFilterText"))
            Dim excludeEndDummy As Boolean = SafeBoolObj(GetProp(pd, "excludeEndDummy"), False)
            Dim options As New HubCommonOptionsStorageService.HubCommonOptions() With {
                .ExtraParamsText = If(extraText, String.Empty),
                .TargetFilterText = If(filterText, String.Empty),
                .ExcludeEndDummy = excludeEndDummy
            }

            Dim ok = HubCommonOptionsStorageService.Save(options)
            SendToWeb("commonoptions:saved", New With {.ok = ok})
        End Sub

        Private Class MultiConnectorOptions
            Public Property Enabled As Boolean
            Public Property Tol As Double = 1.0R
            Public Property Unit As String = "inch"
            Public Property Param As String = "Comments"
        End Class

        Private Class MultiPmsOptions
            Public Property Enabled As Boolean
            Public Property NdRound As Integer = 3
            Public Property TolMm As Double = 0.01R
            Public Property ClassMatch As Boolean
        End Class

        Private Class MultiGuidOptions
            Public Property Enabled As Boolean
            Public Property IncludeFamily As Boolean
            Public Property IncludeAnnotation As Boolean
        End Class

        Private Class MultiFamilyLinkOptions
            Public Property Enabled As Boolean
            Public Property Targets As List(Of FamilyLinkTargetParam) = New List(Of FamilyLinkTargetParam)()
        End Class

        Private Class MultiPointsOptions
            Public Property Enabled As Boolean
            Public Property Unit As String = "ft"
        End Class

        Private Class MultiRunRequest
            Public Property Common As MultiCommonOptions = New MultiCommonOptions()
            Public Property Connector As MultiConnectorOptions = New MultiConnectorOptions()
            Public Property Pms As MultiPmsOptions = New MultiPmsOptions()
            Public Property Guid As MultiGuidOptions = New MultiGuidOptions()
            Public Property FamilyLink As MultiFamilyLinkOptions = New MultiFamilyLinkOptions()
            Public Property Points As MultiPointsOptions = New MultiPointsOptions()
            Public Property RvtPaths As List(Of String) = New List(Of String)()
        End Class

        Private Class MultiRunItem
            Public Property File As String = ""
            Public Property Status As String = ""
            Public Property Reason As String = ""
            Public Property Phase As String = ""
            Public Property ElapsedMs As Long
        End Class

        Private Shared ReadOnly _multiLock As New Object()
        Private Shared _multiQueue As Queue(Of String)
        Private Shared _multiTotal As Integer
        Private Shared _multiIndex As Integer
        Private Shared _multiActive As Boolean
        Private Shared _multiPending As Boolean
        Private Shared _multiBusy As Boolean
        Private Shared _multiRequest As MultiRunRequest
        Private Shared _multiApp As UIApplication
        Private Shared _multiIdlingBound As Boolean

        Private Shared _multiConnectorRows As List(Of Dictionary(Of String, Object))
        Private Shared _multiConnectorExtras As List(Of String)
        Private Shared _multiPmsClassRows As List(Of Dictionary(Of String, Object))
        Private Shared _multiPmsSizeRows As List(Of Dictionary(Of String, Object))
        Private Shared _multiPmsRoutingRows As List(Of Dictionary(Of String, Object))
        Private Shared _multiGuidProject As DataTable
        Private Shared _multiGuidFamilyDetail As DataTable
        Private Shared _multiGuidFamilyIndex As DataTable
        Private Shared _multiFamilyLinkRows As List(Of FamilyLinkAuditRow)
        Private Shared _multiPointRows As List(Of ExportPointsService.Row)
        Private Shared _multiRunItems As List(Of MultiRunItem)

        ' === hub:pick-rvt ===
        ' payload: none
        ' response: hub:rvt-picked { paths:[string] }
        Private Sub HandleMultiPickRvt()
            Using dlg As New OpenFileDialog()
                dlg.Filter = "Revit Project (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.Title = "RVT 파일 선택"
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> DialogResult.OK Then Return
                Dim files As String() = dlg.FileNames
                SendToWeb("hub:rvt-picked", New With {.paths = files})
            End Using
        End Sub

        ' === hub:multi-clear ===
        ' payload: { key?: string }
        Private Sub HandleMultiClear(payload As Object)
            Dim key As String = TryCast(GetProp(payload, "key"), String)
            If String.IsNullOrWhiteSpace(key) Then
                ResetMultiCaches()
                Return
            End If
            Select Case key.ToLowerInvariant()
                Case "connector"
                    _multiConnectorRows = Nothing
                    _multiConnectorExtras = Nothing
                Case "pms"
                    _multiPmsClassRows = Nothing
                    _multiPmsSizeRows = Nothing
                    _multiPmsRoutingRows = Nothing
                Case "guid"
                    _multiGuidProject = Nothing
                    _multiGuidFamilyDetail = Nothing
                    _multiGuidFamilyIndex = Nothing
                Case "familylink"
                    _multiFamilyLinkRows = Nothing
                Case "points"
                    _multiPointRows = Nothing
            End Select
        End Sub

        ' === hub:multi-run ===
        ' payload:
        '  { rvtPaths:[], commonOptions:{extraParams,targetFilter,excludeEndDummy},
        '    features:{connector,pms,guid,familylink,points} }
        ' response:
        '  hub:multi-progress {percent,message,detail}
        '  hub:multi-done { summary:{key:{rows}} }
        Private Sub HandleMultiRun(app As UIApplication, payload As Object)
            Dim req As MultiRunRequest = ParseMultiRequest(payload)
            If req Is Nothing OrElse req.RvtPaths.Count = 0 Then
                SendToWeb("hub:multi-error", New With {.message = "검토할 RVT 파일이 없습니다."})
                Return
            End If
            If Not AnyFeatureEnabled(req) Then
                SendToWeb("hub:multi-error", New With {.message = "선택된 기능이 없습니다."})
                Return
            End If

            SyncLock _multiLock
                _multiRequest = req
                _multiQueue = New Queue(Of String)(req.RvtPaths)
                _multiTotal = req.RvtPaths.Count
                _multiIndex = 0
                _multiActive = True
                _multiPending = True
                _multiBusy = False
                _multiApp = app
                _multiRunItems = New List(Of MultiRunItem)()
                ResetMultiCaches()
                If Not _multiIdlingBound Then
                    AddHandler app.Idling, AddressOf HandleMultiIdling
                    _multiIdlingBound = True
                End If
            End SyncLock

            ReportMultiProgress(0.0R, "배치 검토 시작", $"{req.RvtPaths.Count}개 파일 준비")
        End Sub

        Private Sub HandleMultiIdling(sender As Object, e As Autodesk.Revit.UI.Events.IdlingEventArgs)
            Dim shouldRun As Boolean = False
            SyncLock _multiLock
                shouldRun = _multiActive AndAlso _multiPending AndAlso Not _multiBusy
                If shouldRun Then
                    _multiPending = False
                    _multiBusy = True
                End If
            End SyncLock
            If shouldRun Then
                ProcessMultiNext(_multiApp)
            End If
        End Sub

        Private Sub ProcessMultiNext(app As UIApplication)
            Dim filePath As String = Nothing
            SyncLock _multiLock
                If _multiQueue IsNot Nothing AndAlso _multiQueue.Count > 0 Then
                    filePath = _multiQueue.Dequeue()
                    _multiIndex += 1
                End If
            End SyncLock

            If String.IsNullOrWhiteSpace(filePath) Then
                FinishMultiRun()
                Return
            End If

            Dim safeName As String = System.IO.Path.GetFileName(filePath)
            Dim basePct As Double = If(_multiTotal > 0, CDbl(_multiIndex - 1) / CDbl(_multiTotal), 0.0R)
            ReportMultiProgress(basePct * 100.0R, "파일 여는 중", safeName)

            Dim doc As Document = Nothing
            Dim phase As String = "OPEN"
            Dim started = Date.Now
            Try
                If Not System.IO.File.Exists(filePath) Then
                    ReportMultiProgress(basePct * 100.0R, "파일을 찾을 수 없습니다.", safeName)
                    AppendMultiConnectorError(safeName, "파일을 찾을 수 없습니다.")
                    AppendMultiRunItem(safeName, "skipped", "파일을 찾을 수 없습니다.", "OPEN", started)
                    GoTo NextItem
                End If

                Dim mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath)
                doc = app.Application.OpenDocumentFile(mp, BuildOpenOptions())
                ReportMultiProgress(basePct * 100.0R, "파일 열기 완료", safeName)
                phase = "RUN"

                RunMultiForDocument(app, doc, filePath, safeName, basePct)
                AppendMultiRunItem(safeName, "success", "", "DONE", started)
            Catch ex As Exception
                AppendMultiConnectorError(safeName, $"파일 처리 실패: {ex.Message}")
                ReportMultiProgress(basePct * 100.0R, "파일 처리 실패 (건너뜀)", safeName)
                AppendMultiRunItem(safeName, "failed", ex.Message, phase, started)
                SendToWeb("host:warn", New With {.message = $"파일 처리 실패: {safeName} - {ex.Message}"})
            Finally
                If doc IsNot Nothing Then
                    Try
                        doc.Close(False)
                    Catch
                    End Try
                End If
            End Try

NextItem:
            SyncLock _multiLock
                _multiBusy = False
                If _multiQueue IsNot Nothing AndAlso _multiQueue.Count > 0 Then
                    _multiPending = True
                Else
                    _multiActive = False
                End If
            End SyncLock

            If Not _multiActive Then
                FinishMultiRun()
            End If
        End Sub

        Private Sub RunMultiForDocument(app As UIApplication, doc As Document, path As String, safeName As String, basePct As Double)
            Dim steps As Integer = CountEnabledFeatures(_multiRequest)
            Dim stepIndex As Integer = 0

            If _multiRequest.Connector.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "커넥터 진단 실행 중", safeName)
                Dim extras = ParseExtraParams(_multiRequest.Common.ExtraParams)
                Dim rows = ConnectorDiagnosticsService.RunOnDocument(doc, _multiRequest.Connector.Tol, _multiRequest.Connector.Unit, _multiRequest.Connector.Param, extras, _multiRequest.Common.TargetFilter, _multiRequest.Common.ExcludeEndDummy, Sub(pct, msg)
                                                                                                                                                                                                      Dim overallPct = ((basePct + (pct / 100.0R) / Math.Max(_multiTotal, 1)) * 100.0R)
                                                                                                                                                                                                      ReportMultiProgress(overallPct, "커넥터 진단 실행 중", $"{safeName} · {msg}")
                                                                                                                                                                                                  End Sub)
                If rows IsNot Nothing AndAlso rows.Count > 0 Then
                    For Each row In rows
                        If row IsNot Nothing Then row("File") = safeName
                    Next
                    If _multiConnectorRows Is Nothing Then _multiConnectorRows = New List(Of Dictionary(Of String, Object))()
                    _multiConnectorRows.AddRange(rows)
                    _multiConnectorExtras = extras
                Else
                    If _multiConnectorRows Is Nothing Then _multiConnectorRows = New List(Of Dictionary(Of String, Object))()
                    _multiConnectorRows.Add(New Dictionary(Of String, Object) From {
                        {"File", safeName},
                        {"ConnectionType", "OK"},
                        {"ParamCompare", "OK"},
                        {"Status", "오류 없음"},
                        {"ErrorMessage", ""}
                    })
                    _multiConnectorExtras = extras
                End If
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "커넥터 진단 완료", safeName)
            End If

            If _multiRequest.Pms.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "PMS 검토 실행 중", safeName)
                If _pmsRows Is Nothing OrElse _pmsRows.Count = 0 Then
                    SendToWeb("host:warn", New With {.message = "PMS Excel이 등록되지 않았습니다."})
                Else
                    Dim opts As New SegmentPmsCheckService.ExtractOptions With {
                        .NdRound = _multiRequest.Pms.NdRound,
                        .ToleranceMm = _multiRequest.Pms.TolMm
                    }
                    Dim compareOpts As New SegmentPmsCheckService.CompareOptions With {
                        .NdRound = _multiRequest.Pms.NdRound,
                        .TolMm = _multiRequest.Pms.TolMm,
                        .ClassMatch = _multiRequest.Pms.ClassMatch
                    }
                    Dim ds = SegmentPmsCheckService.ExtractFromDocument(app, doc, path, opts, Nothing)
                    Dim groups = SegmentPmsCheckService.BuildGroups(ds)
                    Dim suggestions = SegmentPmsCheckService.SuggestGroupMappings(groups, _pmsRows)
                    Dim mappings = BuildMappingsFromSuggestions(suggestions)
                    Dim run = SegmentPmsCheckService.RunCompare(ds, _pmsRows, mappings, compareOpts)
                    AppendSegmentPmsRows(run, ds)
                End If
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "PMS 검토 완료", safeName)
            End If

            If _multiRequest.Guid.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "GUID 검토 실행 중", safeName)
                Dim res = GuidAuditService.Run(app, If(_multiRequest.Guid.IncludeFamily, 2, 1), New List(Of String) From {path}, Nothing, Nothing, _multiRequest.Guid.IncludeFamily, _multiRequest.Guid.IncludeAnnotation)
                MergeGuidResult(res)
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "GUID 검토 완료", safeName)
            End If

            If _multiRequest.FamilyLink.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "패밀리 연동 검토 실행 중", safeName)
                Dim rows = FamilyLinkAuditService.RunOnDocument(doc, path, _multiRequest.FamilyLink.Targets, Nothing)
                If rows IsNot Nothing Then
                    If _multiFamilyLinkRows Is Nothing Then _multiFamilyLinkRows = New List(Of FamilyLinkAuditRow)()
                    _multiFamilyLinkRows.AddRange(rows)
                End If
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "패밀리 연동 검토 완료", safeName)
            End If

            If _multiRequest.Points.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "Point 추출 실행 중", safeName)
                Dim rows = ExportPointsService.RunOnDocument(doc, safeName, Nothing)
                If rows IsNot Nothing Then
                    If _multiPointRows Is Nothing Then _multiPointRows = New List(Of ExportPointsService.Row)()
                    _multiPointRows.AddRange(rows)
                End If
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "Point 추출 완료", safeName)
            End If
        End Sub

        Private Sub FinishMultiRun()
            Dim summary As New Dictionary(Of String, Object) From {
                {"connector", New With {.rows = If(_multiConnectorRows, New List(Of Dictionary(Of String, Object))()).Count}},
                {"pms", New With {.rows = If(_multiPmsSizeRows, New List(Of Dictionary(Of String, Object))()).Count}},
                {"guid", New With {.rows = If(_multiGuidProject, New DataTable()).Rows.Count}},
                {"familylink", New With {.rows = If(_multiFamilyLinkRows, New List(Of FamilyLinkAuditRow)()).Count}},
                {"points", New With {.rows = If(_multiPointRows, New List(Of ExportPointsService.Row)()).Count}}
            }
            SendToWeb("hub:multi-done", New With {.summary = summary})
            SendToWeb("multi:review-summary", BuildMultiSummaryPayload())
        End Sub

        ' === hub:multi-export ===
        ' payload: { key, excelMode }
        Private Sub HandleMultiExport(payload As Object)
            Dim key As String = TryCast(GetProp(payload, "key"), String)
            Dim excelMode As String = TryCast(GetProp(payload, "excelMode"), String)
            Dim doAutoFit As Boolean = String.Equals(excelMode, "normal", StringComparison.OrdinalIgnoreCase)

            Try
                Select Case If(key, "").ToLowerInvariant()
                    Case "connector"
                        ExportConnector(doAutoFit)
                    Case "pms"
                        ExportSegmentPms(doAutoFit)
                    Case "guid"
                        ExportGuid(excelMode)
                    Case "familylink"
                        ExportFamilyLink(doAutoFit)
                    Case "points"
                        ExportPoints(doAutoFit)
                    Case Else
                        SendToWeb("hub:multi-exported", New With {.ok = False, .message = "알 수 없는 기능 키입니다."})
                        Return
                End Select
            Catch ex As Exception
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = ex.Message})
                Return
            End Try
        End Sub

        Private Sub ExportConnector(doAutoFit As Boolean)
            Dim rows = If(_multiConnectorRows, New List(Of Dictionary(Of String, Object))())
            Dim allFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing Then
                For Each path In _multiRequest.RvtPaths
                    Dim name = System.IO.Path.GetFileName(TryCast(path, String))
                    If Not String.IsNullOrWhiteSpace(name) Then allFiles.Add(name)
                Next
            End If
            If rows.Count = 0 AndAlso allFiles.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "커넥터 결과가 없습니다."})
                Return
            End If

            Dim existingFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each row In rows
                If row IsNot Nothing AndAlso row.ContainsKey("File") AndAlso row("File") IsNot Nothing Then
                    existingFiles.Add(row("File").ToString())
                End If
            Next
            For Each fileName In allFiles
                If Not existingFiles.Contains(fileName) Then
                    rows.Add(New Dictionary(Of String, Object) From {
                        {"File", fileName},
                        {"Status", "오류 없음"}
                    })
                End If
            Next

            Dim extras = If(_multiConnectorExtras, New List(Of String)())
            Dim headers As List(Of String) = BuildConnectorHeaders(extras)
            Dim table = BuildTableFromRows(headers, rows)
            If Not ValidateSchema(table, headers) Then Throw New InvalidOperationException("스키마 검증 실패: 커넥터")
            Dim saved = ExcelCore.PickAndSaveXlsx("Connector Diagnostics", table, $"Connector_{Date.Now:yyyyMMdd_HHmm}.xlsx", doAutoFit, "hub:multi-progress")
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportSegmentPms(doAutoFit As Boolean)
            Dim classRows = If(_multiPmsClassRows, New List(Of Dictionary(Of String, Object))())
            Dim sizeRows = If(_multiPmsSizeRows, New List(Of Dictionary(Of String, Object))())
            Dim routingRows = If(_multiPmsRoutingRows, New List(Of Dictionary(Of String, Object))())
            If classRows.Count = 0 AndAlso sizeRows.Count = 0 AndAlso routingRows.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "PMS 결과가 없습니다."})
                Return
            End If
            Dim sheetList As New List(Of KeyValuePair(Of String, DataTable))()
            Dim classHeaders = New List(Of String) From {"File", "PipeType", "Segment", "Class검토결과"}
            Dim sizeHeaders = New List(Of String) From {"FileName", "PipeType", "RevitSegment", "PMSCompared", "ND", "ID", "OD", "PMS ND", "PMS ID", "PMS OD", "Result"}
            Dim routingHeaders = New List(Of String) From {"File", "PipeType", "Part", "Type", "Class검토"}

            Dim classTable = BuildTableFromRows(classHeaders, classRows)
            Dim sizeTable = BuildTableFromRows(sizeHeaders, sizeRows)
            Dim routingTable = BuildTableFromRows(routingHeaders, routingRows)

            If classTable.Rows.Count > 0 AndAlso Not ValidateSchema(classTable, classHeaders) Then Throw New InvalidOperationException("스키마 검증 실패: PMS Class")
            If sizeTable.Rows.Count > 0 AndAlso Not ValidateSchema(sizeTable, sizeHeaders) Then Throw New InvalidOperationException("스키마 검증 실패: PMS Size")
            If routingTable.Rows.Count > 0 AndAlso Not ValidateSchema(routingTable, routingHeaders) Then Throw New InvalidOperationException("스키마 검증 실패: PMS Routing")

            If classTable.Rows.Count > 0 Then sheetList.Add(New KeyValuePair(Of String, DataTable)("Pipe Segment Class검토", classTable))
            If sizeTable.Rows.Count > 0 Then sheetList.Add(New KeyValuePair(Of String, DataTable)("PMS vs Segment Size검토", sizeTable))
            If routingTable.Rows.Count > 0 Then sheetList.Add(New KeyValuePair(Of String, DataTable)("Routing Class검토", routingTable))

            Dim saved = ExcelCore.PickAndSaveXlsxMulti(sheetList, $"SegmentPms_{Date.Now:yyyyMMdd_HHmm}.xlsx", doAutoFit, "hub:multi-progress")
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportGuid(excelMode As String)
            If _multiGuidProject Is Nothing Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "GUID 결과가 없습니다."})
                Return
            End If
            Dim sheets As New List(Of KeyValuePair(Of String, DataTable))()
            sheets.Add(New KeyValuePair(Of String, DataTable)("RVT 검토결과", GuidAuditService.PrepareExportTable(_multiGuidProject, 1)))
            If _multiGuidFamilyDetail IsNot Nothing Then
                sheets.Add(New KeyValuePair(Of String, DataTable)("Family 검토결과", GuidAuditService.PrepareExportTable(_multiGuidFamilyDetail, 2)))
            End If
            Dim saved = GuidAuditService.ExportMulti(sheets, excelMode, "hub:multi-progress")
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub



        Private Sub ExportFamilyLink(doAutoFit As Boolean)
            Dim rows = If(_multiFamilyLinkRows, New List(Of FamilyLinkAuditRow)())
            If rows.Count = 0 Then
                rows.Add(New FamilyLinkAuditRow With {
                    .FileName = "(Summary)",
                    .Issue = "OK",
                    .Notes = "검토 완료: 이상 없음(0건)"
                })
            End If
            Dim saved = FamilyLinkAuditExport.Export(rows, fastExport:=Not doAutoFit, autoFit:=doAutoFit)
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportPoints(doAutoFit As Boolean)
            If _multiPointRows Is Nothing OrElse _multiPointRows.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "Point 추출 결과가 없습니다."})
                Return
            End If
            Dim unit As String = "ft"
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Points IsNot Nothing Then
                unit = _multiRequest.Points.Unit
            End If
            Dim headers = BuildPointHeaders(unit)
            Dim rows As New List(Of Dictionary(Of String, Object))()
            For Each r In _multiPointRows
                rows.Add(New Dictionary(Of String, Object) From {
                    {"File", r.File},
                    {"ProjectPoint_E", ConvertPoint(r.ProjectE, unit)},
                    {"ProjectPoint_N", ConvertPoint(r.ProjectN, unit)},
                    {"ProjectPoint_Z", ConvertPoint(r.ProjectZ, unit)},
                    {"SurveyPoint_E", ConvertPoint(r.SurveyE, unit)},
                    {"SurveyPoint_N", ConvertPoint(r.SurveyN, unit)},
                    {"SurveyPoint_Z", ConvertPoint(r.SurveyZ, unit)},
                    {"TrueNorthAngle", Math.Round(r.TrueNorth, 3)}
                })
            Next
            Dim table = BuildPointTable(headers, rows)
            If Not ValidateSchema(table, headers) Then Throw New InvalidOperationException("스키마 검증 실패: Points")
            Dim saved = ExcelCore.PickAndSaveXlsx("Points", table, $"Points_{Date.Now:yyyyMMdd_HHmm}.xlsx", doAutoFit, "hub:multi-progress")
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Shared Sub ResetMultiCaches()
            _multiConnectorRows = Nothing
            _multiConnectorExtras = Nothing
            _multiPmsClassRows = Nothing
            _multiPmsSizeRows = Nothing
            _multiPmsRoutingRows = Nothing
            _multiGuidProject = Nothing
            _multiGuidFamilyDetail = Nothing
            _multiGuidFamilyIndex = Nothing
            _multiFamilyLinkRows = Nothing
            _multiPointRows = Nothing
        End Sub

        Private Function ParseMultiRequest(payload As Object) As MultiRunRequest
            Dim req As New MultiRunRequest()
            Dim pd = ToDict(payload)
            req.RvtPaths = ExtractStringList(pd, "rvtPaths")

            Dim commonObj As Object = Nothing
            If pd.TryGetValue("commonOptions", commonObj) Then
                Dim commonDict = ToDict(commonObj)
                req.Common.ExtraParams = SafeStr(GetDictValue(commonDict, "extraParams"))
                req.Common.TargetFilter = SafeStr(GetDictValue(commonDict, "targetFilter"))
                req.Common.ExcludeEndDummy = ToBool(GetDictValue(commonDict, "excludeEndDummy"))
            End If

            Dim featuresObj As Object = Nothing
            If pd.TryGetValue("features", featuresObj) Then
                Dim fd = ToDict(featuresObj)
                req.Connector = ParseConnector(fd)
                req.Pms = ParsePms(fd)
                req.Guid = ParseGuid(fd)
                req.FamilyLink = ParseFamilyLink(fd)
                req.Points = ParsePoints(fd)
            End If
            Return req
        End Function

        Private Function ParseConnector(fd As Dictionary(Of String, Object)) As MultiConnectorOptions
            Dim opt As New MultiConnectorOptions()
            Dim obj = GetDictValue(fd, "connector")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.Tol = ToDouble(GetDictValue(d, "tol"), 1.0R)
            opt.Unit = SafeStr(GetDictValue(d, "unit"))
            opt.Param = SafeStr(GetDictValue(d, "param"))
            If String.IsNullOrWhiteSpace(opt.Unit) Then opt.Unit = "inch"
            If String.IsNullOrWhiteSpace(opt.Param) Then opt.Param = "Comments"
            Return opt
        End Function

        Private Function ParsePms(fd As Dictionary(Of String, Object)) As MultiPmsOptions
            Dim opt As New MultiPmsOptions()
            Dim obj = GetDictValue(fd, "pms")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.NdRound = ToInt(GetDictValue(d, "ndRound"), 3)
            opt.TolMm = ToDouble(GetDictValue(d, "tolMm"), 0.01R)
            opt.ClassMatch = ToBool(GetDictValue(d, "classMatch"))
            Return opt
        End Function

        Private Function ParseGuid(fd As Dictionary(Of String, Object)) As MultiGuidOptions
            Dim opt As New MultiGuidOptions()
            Dim obj = GetDictValue(fd, "guid")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.IncludeFamily = ToBool(GetDictValue(d, "includeFamily"))
            opt.IncludeAnnotation = ToBool(GetDictValue(d, "includeAnnotation"))
            Return opt
        End Function

        Private Function ParseFamilyLink(fd As Dictionary(Of String, Object)) As MultiFamilyLinkOptions
            Dim opt As New MultiFamilyLinkOptions()
            Dim obj = GetDictValue(fd, "familylink")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            Dim rawTargets = GetDictValue(d, "targets")
            Dim targets As New List(Of FamilyLinkTargetParam)()
            Dim arr = TryCast(rawTargets, System.Collections.IEnumerable)
            If arr IsNot Nothing AndAlso Not TypeOf rawTargets Is String Then
                For Each o In arr
                    Dim td = ToDict(o)
                    Dim name = SafeStr(GetDictValue(td, "name"))
                    Dim guidStr = SafeStr(GetDictValue(td, "guid"))
                    Dim g As Guid
                    If Guid.TryParse(guidStr, g) Then
                        targets.Add(New FamilyLinkTargetParam With {.Name = name, .Guid = g})
                    End If
                Next
            End If
            opt.Targets = targets
            Return opt
        End Function

        Private Function ParsePoints(fd As Dictionary(Of String, Object)) As MultiPointsOptions
            Dim opt As New MultiPointsOptions()
            Dim obj = GetDictValue(fd, "points")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.Unit = SafeStr(GetDictValue(d, "unit"))
            If String.IsNullOrWhiteSpace(opt.Unit) Then opt.Unit = "ft"
            Return opt
        End Function

        Private Shared Function AnyFeatureEnabled(req As MultiRunRequest) As Boolean
            If req Is Nothing Then Return False
            Return req.Connector.Enabled OrElse req.Pms.Enabled OrElse req.Guid.Enabled OrElse req.FamilyLink.Enabled OrElse req.Points.Enabled
        End Function

        Private Shared Function CountEnabledFeatures(req As MultiRunRequest) As Integer
            If req Is Nothing Then Return 0
            Dim count As Integer = 0
            If req.Connector.Enabled Then count += 1
            If req.Pms.Enabled Then count += 1
            If req.Guid.Enabled Then count += 1
            If req.FamilyLink.Enabled Then count += 1
            If req.Points.Enabled Then count += 1
            Return Math.Max(count, 1)
        End Function

        Private Sub AppendMultiRunItem(fileName As String, status As String, reason As String, phase As String, started As DateTime)
            If _multiRunItems Is Nothing Then _multiRunItems = New List(Of MultiRunItem)()
            Dim elapsed = CLng((Date.Now - started).TotalMilliseconds)
            _multiRunItems.Add(New MultiRunItem With {
                .File = fileName,
                .Status = status,
                .Reason = reason,
                .Phase = phase,
                .ElapsedMs = elapsed
            })
        End Sub

        Private Function BuildMultiSummaryPayload() As Object
            Dim items As List(Of MultiRunItem) = If(_multiRunItems, New List(Of MultiRunItem)())

            Dim total As Integer = If(_multiTotal > 0, _multiTotal, items.Count)

            Dim success As Integer = items.Where(Function(x) String.Equals(x.Status, "success", StringComparison.OrdinalIgnoreCase)).Count()
            Dim skipped As Integer = items.Where(Function(x) String.Equals(x.Status, "skipped", StringComparison.OrdinalIgnoreCase)).Count()
            Dim failed As Integer = items.Where(Function(x) String.Equals(x.Status, "failed", StringComparison.OrdinalIgnoreCase)).Count()

            Return New With {
        .ok = True,
        .mode = "multiRvt",
        .featureId = "multi_rvt_batch",
        .title = "다중 RVT 검토",
        .finishedAt = Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
        .total = total,
        .success = success,
        .skipped = skipped,
        .failed = failed,
        .canceled = False,
        .items = items
    }
        End Function

        Private Sub ReportMultiProgress(percent As Double, message As String, detail As String)
            SendToWeb("hub:multi-progress", New With {
                .percent = Math.Max(0.0R, Math.Min(100.0R, percent)),
                .message = message,
                .detail = detail,
                .title = "다중 RVT 검토"
            })
        End Sub

        Private Function CalcStepPercent(basePct As Double, stepIndex As Integer, totalSteps As Integer) As Double
            Dim perFile As Double = If(_multiTotal > 0, 1.0R / CDbl(_multiTotal), 1.0R)
            Dim stepShare As Double = perFile / CDbl(Math.Max(totalSteps, 1))
            Dim stepPct As Double = (basePct + (stepShare * CDbl(stepIndex))) * 100.0R
            Return Math.Min(stepPct, 99.9R)
        End Function

        Private Shared Function BuildOpenOptions() As OpenOptions
            Dim opt As New OpenOptions()
            Try
                opt.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
            Catch
            End Try
            Try
                Dim ws = New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                opt.SetOpenWorksetsConfiguration(ws)
            Catch
            End Try
            Return opt
        End Function

        Private Shared Function BuildConnectorHeaders(extras As IList(Of String)) As List(Of String)
            Dim headers As New List(Of String) From {
                "File", "Id1", "Id2", "Category1", "Category2", "Family1", "Family2", "Distance (inch)", "ConnectionType", "ParamName", "Value1", "Value2", "ParamCompare", "Status", "ErrorMessage"
            }
            If extras IsNot Nothing Then
                For Each name In extras
                    headers.Add($"{name}(ID1)")
                    headers.Add($"{name}(ID2)")
                Next
            End If
            Return headers
        End Function

        Private Sub AppendMultiConnectorError(fileName As String, message As String)
            If _multiRequest Is Nothing OrElse _multiRequest.Connector Is Nothing OrElse Not _multiRequest.Connector.Enabled Then Return
            If _multiConnectorRows Is Nothing Then _multiConnectorRows = New List(Of Dictionary(Of String, Object))()
            _multiConnectorRows.Add(New Dictionary(Of String, Object) From {
                {"File", fileName},
                {"ConnectionType", "ERROR"},
                {"ParamName", _multiRequest.Connector.Param},
                {"ParamCompare", "N/A"},
                {"Status", "ERROR"},
                {"ErrorMessage", message}
            })
        End Sub

        Private Shared Function BuildTableFromRows(headers As IList(Of String), rows As IList(Of Dictionary(Of String, Object))) As DataTable
            Dim dt As New DataTable("Export")
            For Each h In headers
                dt.Columns.Add(h)
            Next
            If rows Is Nothing Then Return dt
            For Each r In rows
                Dim dr = dt.NewRow()
                For i As Integer = 0 To headers.Count - 1
                    Dim key = headers(i)
                    dr(i) = If(r IsNot Nothing AndAlso r.ContainsKey(key) AndAlso r(key) IsNot Nothing, r(key).ToString(), String.Empty)
                Next
                dt.Rows.Add(dr)
            Next
            Return dt
        End Function

        Private Shared Function ValidateSchema(table As DataTable, headers As IList(Of String)) As Boolean
            If table Is Nothing OrElse headers Is Nothing Then Return False
            If table.Columns.Count <> headers.Count Then Return False
            For i As Integer = 0 To headers.Count - 1
                If Not String.Equals(table.Columns(i).ColumnName, headers(i), StringComparison.Ordinal) Then Return False
            Next
            Return True
        End Function

        Private Sub AppendSegmentPmsRows(run As SegmentPmsCheckService.RunResult, ds As DataSet)
            If run Is Nothing Then Return
            Dim classRows = SegmentPmsCheckService.BuildClassCheckRows(run.MapTable)
            Dim sizeRows = SegmentPmsCheckService.BuildSizeCheckRows(run.CompareTable)
            Dim routingRows = SegmentPmsCheckService.BuildRoutingClassRows(ds)

            If _multiPmsClassRows Is Nothing Then _multiPmsClassRows = New List(Of Dictionary(Of String, Object))()
            If _multiPmsSizeRows Is Nothing Then _multiPmsSizeRows = New List(Of Dictionary(Of String, Object))()
            If _multiPmsRoutingRows Is Nothing Then _multiPmsRoutingRows = New List(Of Dictionary(Of String, Object))()
            _multiPmsClassRows.AddRange(classRows)
            _multiPmsSizeRows.AddRange(sizeRows)
            _multiPmsRoutingRows.AddRange(routingRows)
        End Sub

        Private Sub MergeGuidResult(res As GuidAuditService.RunResult)
            If res Is Nothing Then Return
            _multiGuidProject = MergeTable(_multiGuidProject, res.Project)
            If res.IncludeFamily Then
                _multiGuidFamilyDetail = MergeTable(_multiGuidFamilyDetail, res.FamilyDetail)
                _multiGuidFamilyIndex = MergeTable(_multiGuidFamilyIndex, res.FamilyIndex)
            End If
        End Sub

        Private Shared Function MergeTable(master As DataTable, part As DataTable) As DataTable
            If part Is Nothing Then Return master
            If master Is Nothing Then
                master = part.Clone()
            End If
            For Each r As DataRow In part.Rows
                master.ImportRow(r)
            Next
            Return master
        End Function

        Private Shared Function BuildMappingsFromSuggestions(suggestions As IList(Of SegmentPmsCheckService.SuggestedMapping)) As List(Of SegmentPmsCheckService.MappingSelection)
            Dim list As New List(Of SegmentPmsCheckService.MappingSelection)()
            If suggestions Is Nothing Then Return list
            For Each s In suggestions
                If s Is Nothing Then Continue For
                If String.IsNullOrWhiteSpace(s.PmsSegmentKey) Then Continue For
                Dim item As New SegmentPmsCheckService.MappingSelection With {
                    .File = s.File,
                    .PipeTypeName = s.PipeTypeName,
                    .RuleIndex = s.RuleIndex,
                    .SegmentId = s.SegmentId,
                    .SegmentKey = s.SegmentKey,
                    .SelectedClass = s.PmsClass,
                    .SelectedPmsSegment = s.PmsSegmentKey,
                    .MappingSource = "AutoSuggest"
                }
                list.Add(item)
            Next
            Return list
        End Function

        Private Shared Function BuildPointHeaders(unit As String) As List(Of String)
            Dim suffix As String = "(ft)"
            If String.Equals(unit, "m", StringComparison.OrdinalIgnoreCase) Then
                suffix = "(m)"
            ElseIf String.Equals(unit, "mm", StringComparison.OrdinalIgnoreCase) Then
                suffix = "(mm)"
            End If
            Return New List(Of String) From {
                "File",
                $"ProjectPoint_E{suffix}", $"ProjectPoint_N{suffix}", $"ProjectPoint_Z{suffix}",
                $"SurveyPoint_E{suffix}", $"SurveyPoint_N{suffix}", $"SurveyPoint_Z{suffix}",
                "TrueNorthAngle(deg)"
            }
        End Function

        Private Shared Function BuildPointTable(headers As IList(Of String), rows As IList(Of Dictionary(Of String, Object))) As DataTable
            Dim dt As New DataTable("Points")
            For Each h In headers
                dt.Columns.Add(h)
            Next
            For Each r In rows
                Dim dr = dt.NewRow()
                dr(0) = SafeStr(GetRowValue(r, "File"))
                dr(1) = SafeStr(GetRowValue(r, "ProjectPoint_E"))
                dr(2) = SafeStr(GetRowValue(r, "ProjectPoint_N"))
                dr(3) = SafeStr(GetRowValue(r, "ProjectPoint_Z"))
                dr(4) = SafeStr(GetRowValue(r, "SurveyPoint_E"))
                dr(5) = SafeStr(GetRowValue(r, "SurveyPoint_N"))
                dr(6) = SafeStr(GetRowValue(r, "SurveyPoint_Z"))
                dr(7) = SafeStr(GetRowValue(r, "TrueNorthAngle"))
                dt.Rows.Add(dr)
            Next
            Return dt
        End Function

        Private Shared Function ConvertPoint(valueFt As Double, unit As String) As Double
            If String.Equals(unit, "m", StringComparison.OrdinalIgnoreCase) Then
                Return Math.Round(valueFt * 0.3048R, 6)
            End If
            If String.Equals(unit, "mm", StringComparison.OrdinalIgnoreCase) Then
                Return Math.Round(valueFt * 304.8R, 3)
            End If
            Return Math.Round(valueFt, 6)
        End Function

        Private Shared Function ToDict(obj As Object) As Dictionary(Of String, Object)
            Dim dict = TryCast(obj, Dictionary(Of String, Object))
            If dict IsNot Nothing Then Return dict
            Dim res As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If obj Is Nothing Then Return res
            Dim t = obj.GetType()
            For Each p In t.GetProperties()
                Try
                    res(p.Name) = p.GetValue(obj, Nothing)
                Catch
                End Try
            Next
            Return res
        End Function

        Private Shared Function ExtractStringList(dict As Dictionary(Of String, Object), key As String) As List(Of String)
            Dim list As New List(Of String)()
            Dim raw = GetDictValue(dict, key)
            Dim arr = TryCast(raw, System.Collections.IEnumerable)
            If arr IsNot Nothing AndAlso Not TypeOf raw Is String Then
                For Each o In arr
                    Dim s = SafeStr(o)
                    If Not String.IsNullOrWhiteSpace(s) AndAlso Not list.Contains(s) Then
                        list.Add(s)
                    End If
                Next
            Else
                Dim s = SafeStr(raw)
                If Not String.IsNullOrWhiteSpace(s) Then list.Add(s)
            End If
            Return list
        End Function

        Private Shared Function ToBool(obj As Object, Optional defaultValue As Boolean = False) As Boolean
            If obj Is Nothing Then Return defaultValue
            Try
                Return Convert.ToBoolean(obj)
            Catch
                Return defaultValue
            End Try
        End Function

        Private Shared Function ToDouble(obj As Object, defaultValue As Double) As Double
            If obj Is Nothing Then Return defaultValue
            Try
                Return Convert.ToDouble(obj)
            Catch
                Return defaultValue
            End Try
        End Function

        Private Shared Function ToInt(obj As Object, defaultValue As Integer) As Integer
            If obj Is Nothing Then Return defaultValue
            Try
                Return Convert.ToInt32(obj)
            Catch
                Return defaultValue
            End Try
        End Function

        Private Shared Function GetRowValue(row As Dictionary(Of String, Object), key As String) As Object
            If row Is Nothing Then Return Nothing
            Dim val As Object = Nothing
            If row.TryGetValue(key, val) Then Return val
            Return Nothing
        End Function

    End Class

End Namespace
