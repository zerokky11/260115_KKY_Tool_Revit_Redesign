Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Services
Imports KKY_Tool_Revit.Infrastructure
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports NPOI.XSSF.UserModel

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private _extractData As DataSet
        Private _pmsRows As List(Of SegmentPmsCheckService.PmsRow)
        Private _pmsUnitPref As String = "mm"
        Private _lastExtractPath As String = String.Empty
        Private _lastRunResult As SegmentPmsCheckService.RunResult = Nothing
        Private _segmentPmsLastResult As SegmentPmsResultCache = Nothing

        Private Class SegmentPmsResultCache
            Public Property RunResult As SegmentPmsCheckService.RunResult
            Public Property TotalCount As Integer
        End Class

        Private Shared Function ParsePayloadDict(payload As Object) As Dictionary(Of String, Object)
            Dim dict = TryCast(payload, Dictionary(Of String, Object))
            If dict IsNot Nothing Then
                Return dict
            End If

            Dim result As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If payload Is Nothing Then
                Return result
            End If

            Dim t = payload.GetType()
            For Each p In t.GetProperties()
                Try
                    result(p.Name) = p.GetValue(payload, Nothing)
                Catch
                End Try
            Next
            Return result
        End Function

        Private Shared Function ParseStringList(payload As Dictionary(Of String, Object), key As String) As List(Of String)
            Dim list As New List(Of String)()
            If payload Is Nothing OrElse Not payload.ContainsKey(key) Then
                Return list
            End If

            Dim raw = payload(key)
            Dim enumerable = TryCast(raw, System.Collections.IEnumerable)
            If enumerable IsNot Nothing AndAlso Not (TypeOf raw Is String) Then
                For Each o As Object In enumerable
                    Dim s = If(o, String.Empty).ToString()
                    If Not String.IsNullOrWhiteSpace(s) AndAlso Not list.Contains(s) Then
                        list.Add(s)
                    End If
                Next
            Else
                Dim s = If(raw, String.Empty).ToString()
                If Not String.IsNullOrWhiteSpace(s) Then
                    list.Add(s)
                End If
            End If
            Return list
        End Function

        Private Shared Function ParseExtractOptions(payload As Dictionary(Of String, Object)) As SegmentPmsCheckService.ExtractOptions
            Dim opts As New SegmentPmsCheckService.ExtractOptions()
            If payload Is Nothing Then
                Return opts
            End If

            Dim optionsObj As Object = Nothing
            If payload.TryGetValue("options", optionsObj) Then
                Dim dict = ParsePayloadDict(optionsObj)
                If dict.ContainsKey("ndRound") Then
                    Dim v = dict("ndRound")
                    Dim iv As Integer
                    If Integer.TryParse(If(v, 3).ToString(), iv) Then
                        opts.NdRound = iv
                    End If
                End If
                If dict.ContainsKey("tolMm") Then
                    Dim v = dict("tolMm")
                    Dim dv As Double
                    If Double.TryParse(If(v, opts.ToleranceMm).ToString(), dv) Then
                        opts.ToleranceMm = dv
                    End If
                End If
            End If
            If payload.ContainsKey("ndRound") Then
                Dim v = payload("ndRound")
                Dim iv As Integer
                If Integer.TryParse(If(v, opts.NdRound).ToString(), iv) Then
                    opts.NdRound = iv
                End If
            End If
            If payload.ContainsKey("tolMm") Then
                Dim v = payload("tolMm")
                Dim dv As Double
                If Double.TryParse(If(v, opts.ToleranceMm).ToString(), dv) Then
                    opts.ToleranceMm = dv
                End If
            End If
            Return opts
        End Function

        Private Shared Function ParseCompareOptions(payload As Dictionary(Of String, Object)) As SegmentPmsCheckService.CompareOptions
            Dim opts As New SegmentPmsCheckService.CompareOptions()
            If payload Is Nothing Then
                Return opts
            End If

            If payload.ContainsKey("ndRound") Then
                Dim v = payload("ndRound")
                Dim iv As Integer
                If Integer.TryParse(If(v, 3).ToString(), iv) Then
                    opts.NdRound = iv
                End If
            End If
            If payload.ContainsKey("tolMm") Then
                Dim v = payload("tolMm")
                Dim dv As Double
                If Double.TryParse(If(v, 0.01R).ToString(), dv) Then
                    opts.TolMm = dv
                End If
            End If
            If payload.ContainsKey("classMatch") Then
                Try
                    opts.ClassMatch = Convert.ToBoolean(payload("classMatch"))
                Catch
                End Try
            End If
            Return opts
        End Function

        Private Shared Function ParseMappings(payload As Dictionary(Of String, Object)) As List(Of SegmentPmsCheckService.MappingSelection)
            Dim res As New List(Of SegmentPmsCheckService.MappingSelection)()
            If payload Is Nothing OrElse Not payload.ContainsKey("mappings") Then
                Return res
            End If

            Dim raw = payload("mappings")
            Dim arr = TryCast(raw, System.Collections.IEnumerable)
            If arr Is Nothing OrElse TypeOf raw Is String Then
                Return res
            End If

            For Each o In arr
                Dim d = ParsePayloadDict(o)
                Dim item As New SegmentPmsCheckService.MappingSelection()
                If d.ContainsKey("file") Then
                    item.File = If(d("file"), String.Empty).ToString()
                End If
                If d.ContainsKey("pipeType") Then
                    item.PipeTypeName = If(d("pipeType"), String.Empty).ToString()
                End If
                If d.ContainsKey("ruleIndex") Then
                    Dim iv As Integer
                    If Integer.TryParse(If(d("ruleIndex"), 0).ToString(), iv) Then
                        item.RuleIndex = iv
                    End If
                End If
                If d.ContainsKey("segmentId") Then
                    Dim iv As Integer
                    If Integer.TryParse(If(d("segmentId"), 0).ToString(), iv) Then
                        item.SegmentId = iv
                    End If
                End If
                If d.ContainsKey("segmentKey") Then
                    item.SegmentKey = If(d("segmentKey"), String.Empty).ToString()
                End If
                If d.ContainsKey("cls") Then
                    item.SelectedClass = If(d("cls"), String.Empty).ToString()
                End If
                If d.ContainsKey("segment") Then
                    item.SelectedPmsSegment = If(d("segment"), String.Empty).ToString()
                End If
                If d.ContainsKey("source") Then
                    item.MappingSource = If(d("source"), String.Empty).ToString()
                End If
                res.Add(item)
            Next
            Return res
        End Function

        Private Shared Function ParseGroupSelections(payload As Dictionary(Of String, Object)) As List(Of SegmentPmsCheckService.GroupSelection)
            Dim res As New List(Of SegmentPmsCheckService.GroupSelection)()
            If payload Is Nothing OrElse Not payload.ContainsKey("groups") Then
                Return res
            End If
            Dim raw = payload("groups")
            Dim arr = TryCast(raw, System.Collections.IEnumerable)
            If arr Is Nothing OrElse TypeOf raw Is String Then
                Return res
            End If
            For Each o In arr
                Dim d = ParsePayloadDict(o)
                Dim g As New SegmentPmsCheckService.GroupSelection()
                If d.ContainsKey("groupKey") Then
                    g.GroupKey = If(d("groupKey"), String.Empty).ToString()
                End If
                If d.ContainsKey("cls") Then
                    g.SelectedClass = If(d("cls"), String.Empty).ToString()
                End If
                If d.ContainsKey("segment") Then
                    g.SelectedPmsSegment = If(d("segment"), String.Empty).ToString()
                End If
                If d.ContainsKey("source") Then
                    g.SelectionSource = If(d("source"), String.Empty).ToString()
                End If
                res.Add(g)
            Next
            Return res
        End Function

        Private Shared Function SafeIntObj(o As Object, Optional def As Integer = 0) As Integer
            If o Is Nothing Then
                Return def
            End If
            Dim v As Integer
            If Integer.TryParse(o.ToString(), v) Then
                Return v
            End If
            Dim dv As Double
            If Double.TryParse(o.ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, dv) Then
                Return CInt(Math.Truncate(dv))
            End If
            Return def
        End Function

        Private Sub HandleSegmentPmsRvtPickFiles(app As UIApplication, payload As Object)
            Using dlg As New OpenFileDialog()
                dlg.Filter = "Revit Files (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.RestoreDirectory = True
                dlg.Title = "RVT 파일 선택"
                If dlg.ShowDialog() <> DialogResult.OK Then
                    Return
                End If
                Dim files As New List(Of String)()
                For Each f As String In dlg.FileNames
                    files.Add(f)
                Next
                SendToWeb("segmentpms:rvt-picked-files", New With {.paths = files})
            End Using
        End Sub

        Private Sub HandleSegmentPmsRvtPickFolder(app As UIApplication, payload As Object)
            Using dlg As New FolderBrowserDialog()
                dlg.Description = "RVT가 있는 폴더를 선택하세요."
                If dlg.ShowDialog() <> DialogResult.OK Then
                    Return
                End If

                Dim files As New List(Of String)()
                Try
                    For Each f As String In Directory.GetFiles(dlg.SelectedPath, "*.rvt", SearchOption.TopDirectoryOnly)
                        files.Add(f)
                    Next
                Catch
                End Try
                SendToWeb("segmentpms:rvt-picked-folder", New With {.paths = files})
            End Using
        End Sub

        Private Sub HandleSegmentPmsExtractStart(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim files = ParseStringList(pd, "files")
            Dim opts = ParseExtractOptions(pd)
            If files.Count = 0 Then
                SendToWeb("segmentpms:error", New With {.message = "추출할 RVT 파일을 선택하세요."})
                Return
            End If

            Using dlg As New SaveFileDialog()
                dlg.Filter = "Excel (*.xlsx)|*.xlsx"
                dlg.FileName = "SegmentPmsExtract.xlsx"
                dlg.AddExtension = True
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> DialogResult.OK Then
                    Return
                End If

                Try
                    _segmentPmsLastResult = Nothing
                    ReportProgress(files.Count, 0, "open", "추출 시작", String.Empty)
                    _extractData = SegmentPmsCheckService.ExtractToDataSet(app, files, opts, AddressOf ReportProgress)
                    _lastExtractPath = dlg.FileName
                    ReportProgress(files.Count, files.Count, "save", "엑셀 내보내기 중", dlg.FileName)
                    Dim doAutoFit As Boolean = ParseExcelMode(payload)
                    SegmentPmsCheckService.SaveDataSetToXlsx(_extractData, dlg.FileName, doAutoFit, "segmentpms:progress")
                    WaitForFileReady(dlg.FileName)
                    Dim summary = BuildExtractSummary(_extractData)
                    ReportProgress(files.Count, files.Count, "done", "추출 완료", dlg.FileName)
                    SendToWeb("segmentpms:extract-saved", New With {.path = dlg.FileName, .summary = summary})
                Catch ex As Exception
                    ReportProgress(files.Count, files.Count, "error", ex.Message, String.Empty)
                    SendToWeb("segmentpms:error", New With {.message = ex.Message})
                End Try
            End Using
        End Sub

        Private Sub HandleSegmentPmsSaveExtract(app As UIApplication, payload As Object)
            If _extractData Is Nothing Then
                SendToWeb("segmentpms:error", New With {.message = "저장할 추출 데이터가 없습니다."})
                Return
            End If
            Using dlg As New SaveFileDialog()
                dlg.Filter = "Excel (*.xlsx)|*.xlsx"
                dlg.FileName = If(String.IsNullOrWhiteSpace(_lastExtractPath), "SegmentPmsExtract.xlsx", Path.GetFileName(_lastExtractPath))
                dlg.AddExtension = True
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> DialogResult.OK Then
                    Return
                End If
                Try
                    Dim doAutoFit As Boolean = ParseExcelMode(payload)
                    SegmentPmsCheckService.SaveDataSetToXlsx(_extractData, dlg.FileName, doAutoFit, "segmentpms:progress")
                    _lastExtractPath = dlg.FileName
                    WaitForFileReady(dlg.FileName)
                    Dim summary = BuildExtractSummary(_extractData)
                    SendToWeb("segmentpms:extract-saved", New With {.path = dlg.FileName, .summary = summary})
                Catch ex As Exception
                    SendToWeb("segmentpms:error", New With {.message = ex.Message})
                End Try
            End Using
        End Sub

        Private Sub HandleSegmentPmsLoadExtract(app As UIApplication, payload As Object)
            Using dlg As New OpenFileDialog()
                dlg.Filter = "Excel (*.xlsx)|*.xlsx"
                dlg.RestoreDirectory = True
                dlg.Title = "추출 Excel 불러오기"
                If dlg.ShowDialog() <> DialogResult.OK Then
                    Return
                End If

                Try
                    _segmentPmsLastResult = Nothing
                    _extractData = SegmentPmsCheckService.LoadExtractFromXlsx(dlg.FileName)
                    _lastExtractPath = dlg.FileName
                    Dim summary = BuildExtractSummary(_extractData)
                    Dim groups = SegmentPmsCheckService.BuildGroups(_extractData)
                    Dim suggest = SegmentPmsCheckService.SuggestGroupMappings(groups, _pmsRows)
                    Dim pmsOpts = BuildPmsOptions()
                    Dim groupPayload = BuildGroupPayload(groups)
                    SendToWeb("segmentpms:extract-loaded", New With {
                        .summary = summary,
                        .groups = groupPayload,
                        .suggestions = suggest,
                        .pms = pmsOpts,
                        .path = dlg.FileName
                    })
                Catch ex As Exception
                    SendToWeb("segmentpms:error", New With {.message = ex.Message})
                End Try
            End Using
        End Sub

        Private Sub HandleSegmentPmsRegisterPms(app As UIApplication, payload As Object)
            Dim unitPref As String = "mm"
            Dim pd = ParsePayloadDict(payload)
            If pd.ContainsKey("unit") Then
                unitPref = If(pd("unit"), "mm").ToString()
            End If
            Using dlg As New OpenFileDialog()
                dlg.Filter = "Excel (*.xlsx)|*.xlsx"
                dlg.Title = "PMS Excel 선택"
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> DialogResult.OK Then
                    SendToWeb("segmentpms:error", New With {.message = "PMS 불러오기가 취소되었습니다."})
                    Return
                End If

                Try
                    Dim loaded = SegmentPmsCheckService.LoadPmsExcel(dlg.FileName, unitPref)
                    _pmsRows = loaded.Rows
                    _pmsUnitPref = unitPref
                    If loaded.Errors IsNot Nothing AndAlso loaded.Errors.Count > 0 Then
                        SendToWeb("segmentpms:error", New With {.message = String.Join(";", loaded.Errors)})
                        Return
                    End If
                    Dim pmsOpts = BuildPmsOptions()
                    Dim suggestList As List(Of SegmentPmsCheckService.SuggestedMapping) = Nothing
                    Dim groupPayload As List(Of Object) = Nothing
                    If _extractData IsNot Nothing Then
                        Dim groups = SegmentPmsCheckService.BuildGroups(_extractData)
                        suggestList = SegmentPmsCheckService.SuggestGroupMappings(groups, _pmsRows)
                        groupPayload = BuildGroupPayload(groups)
                    End If
                    SendToWeb("segmentpms:pms-registered", New With {.path = dlg.FileName, .options = pmsOpts, .suggestions = suggestList, .groups = groupPayload})
                Catch ex As Exception
                    SendToWeb("segmentpms:error", New With {.message = ex.Message})
                End Try
            End Using
        End Sub

        Private Sub HandleSegmentPmsPrepareMapping(app As UIApplication, payload As Object)
            If _extractData Is Nothing Then
                SendToWeb("segmentpms:error", New With {.message = "추출 데이터를 먼저 불러오세요."})
                Return
            End If
            Dim groups = SegmentPmsCheckService.BuildGroups(_extractData)
            Dim suggestions As List(Of SegmentPmsCheckService.SuggestedMapping) = Nothing
            If _pmsRows IsNot Nothing Then
                suggestions = SegmentPmsCheckService.SuggestGroupMappings(groups, _pmsRows)
            End If
            Dim pmsOpts = BuildPmsOptions()
            Dim groupPayload = BuildGroupPayload(groups)
            SendToWeb("segmentpms:mapping-ready", New With {.groups = groupPayload, .pms = pmsOpts, .suggestions = suggestions})
        End Sub

        Private Sub HandleSegmentPmsRun(app As UIApplication, payload As Object)
            If _extractData Is Nothing Then
                SendToWeb("segmentpms:error", New With {.message = "추출 데이터를 먼저 불러오세요."})
                Return
            End If
            If _pmsRows Is Nothing Then
                SendToWeb("segmentpms:error", New With {.message = "PMS Excel을 등록하세요."})
                Return
            End If

            Dim pd = ParsePayloadDict(payload)
            Dim maps = ParseMappings(pd)
            Dim groupSelections = ParseGroupSelections(pd)
            Dim opts = ParseCompareOptions(pd)
            If (maps Is Nothing OrElse maps.Count = 0) AndAlso groupSelections IsNot Nothing AndAlso groupSelections.Count > 0 Then
                Dim groups = SegmentPmsCheckService.BuildGroups(_extractData)
                maps = SegmentPmsCheckService.ExpandGroupSelections(groups, groupSelections)
            End If

            _segmentPmsLastResult = Nothing
            Try
                Dim run = SegmentPmsCheckService.RunCompare(_extractData, _pmsRows, maps, opts)
                _lastRunResult = run
                _segmentPmsLastResult = New SegmentPmsResultCache With {.RunResult = run, .TotalCount = If(run.CompareTable Is Nothing, 0, run.CompareTable.Rows.Count)}
                Dim compare = DataTableToObjects(run.CompareTable)
                Dim map = DataTableToObjects(run.MapTable)
                Dim revitRaw = DataTableToObjects(run.RevitSizeTable)
                Dim pmsRaw = DataTableToObjects(run.PmsSizeTable)
                Dim err = DataTableToObjects(run.ErrorTable)
                Dim summary = DataTableToObjects(run.SummaryTable)
                SendToWeb("segmentpms:result", New With {
                    .compare = compare,
                    .totalCount = _segmentPmsLastResult.TotalCount,
                    .map = map,
                    .revitRaw = revitRaw,
                    .pmsRaw = pmsRaw,
                    .summary = summary,
                    .errors = err
                })
            Catch ex As Exception
                SendToWeb("segmentpms:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleSegmentPmsSaveResult(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim doAutoFit As Boolean = ParseExcelMode(payload)
            Using dlg As New SaveFileDialog()
                dlg.Filter = "Excel (*.xlsx)|*.xlsx"
                dlg.FileName = "SegmentPmsResult.xlsx"
                dlg.AddExtension = True
                If dlg.ShowDialog() <> DialogResult.OK Then
                    Return
                End If

                Dim mapTable As DataTable = Nothing
                Dim compareTable As DataTable = Nothing
                If _segmentPmsLastResult IsNot Nothing AndAlso _segmentPmsLastResult.RunResult IsNot Nothing Then
                    mapTable = _segmentPmsLastResult.RunResult.MapTable
                    compareTable = _segmentPmsLastResult.RunResult.CompareTable
                End If

                Dim compareRows = If(compareTable, DictListToDataTable(CoerceRowsToDictList(GetDictValue(pd, "compare")), "SizeCompare"))
                Dim mapRows = If(mapTable, DictListToDataTable(CoerceRowsToDictList(GetDictValue(pd, "map")), "PipeTypeSegmentMap"))

                Dim classRows = SegmentPmsCheckService.BuildClassCheckRows(mapRows)
                Dim sizeRows = SegmentPmsCheckService.BuildSizeCheckRows(compareRows)
                Dim routingRows = SegmentPmsCheckService.BuildRoutingClassRows(_extractData)

                Dim totalRowsCount As Integer = classRows.Count + sizeRows.Count + routingRows.Count
                Dim written As Integer = 0
                Try
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset("segmentpms:progress")
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report("segmentpms:progress", "EXCEL_INIT", "엑셀 워크북 준비", 0, totalRowsCount, Nothing, True)
                    LogAutoFitDecision(doAutoFit, "UiBridgeExternalEvent.HandleSegmentPmsSaveResult")
                    Dim wb As IWorkbook = New XSSFWorkbook()

                    If classRows.Count = 0 AndAlso sizeRows.Count = 0 AndAlso routingRows.Count = 0 Then
                        SendToWeb("segmentpms:error", New With {.message = "먼저 검토를 실행하세요."})
                        Return
                    End If

                    AddSheet(wb, "Pipe Segment Class검토", classRows, New List(Of String) From {"File", "PipeType", "Segment", "Class검토결과"}, "segmentpms:progress", written, totalRowsCount, doAutoFit)
                    AddSheet(wb, "PMS vs Segment Size검토", sizeRows, New List(Of String) From {"FileName", "PipeType", "RevitSegment", "PMSCompared", "ND", "ID", "OD", "PMS ND", "PMS ID", "PMS OD", "Result"}, "segmentpms:progress", written, totalRowsCount, doAutoFit)
                    AddSheet(wb, "Routing Class검토", routingRows, New List(Of String) From {"File", "PipeType", "Part", "Type", "Class검토"}, "segmentpms:progress", written, totalRowsCount, doAutoFit)
                    Dim savePath As String = dlg.FileName
                    Try
                        savePath = System.IO.Path.GetFullPath(dlg.FileName)
                    Catch
                    End Try
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report("segmentpms:progress", "EXCEL_SAVE", "파일 저장 중", written, totalRowsCount, Nothing, True)
                    SaveWorkbookSafe(wb, savePath)
                    WaitForFileReady(savePath)
                    wb.Close()
                    Dim autoFitMessage As String = If(doAutoFit, "AutoFit 적용", "빠른 모드: AutoFit 생략")
                    If doAutoFit Then
                        Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report("segmentpms:progress", "AUTOFIT", autoFitMessage, written, totalRowsCount, Nothing, True)
                        Global.KKY_Tool_Revit.Infrastructure.ExcelCore.TryAutoFitWithExcel(savePath)
                    Else
                        Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report("segmentpms:progress", "AUTOFIT", autoFitMessage, written, totalRowsCount, Nothing, True)
                    End If
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report("segmentpms:progress", "DONE", "엑셀 내보내기 완료", written, totalRowsCount, 100.0R, True)
                    SendToWeb("segmentpms:saved", New With {.path = savePath})
                Catch ex As Exception
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report("segmentpms:progress", "ERROR", ex.Message, written, totalRowsCount, Nothing, True)
                    SendToWeb("segmentpms:error", New With {.message = ex.Message})
                End Try
            End Using
        End Sub

        Private Function BuildExtractSummary(ds As DataSet) As String
            If ds Is Nothing OrElse Not ds.Tables.Contains(SegmentPmsCheckService.TableRules) Then
                Return String.Empty
            End If
            Dim rules = ds.Tables(SegmentPmsCheckService.TableRules)
            Dim sizes As DataTable = Nothing
            If ds.Tables.Contains(SegmentPmsCheckService.TableSizes) Then
                sizes = ds.Tables(SegmentPmsCheckService.TableSizes)
            End If
            Dim fileSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim pipeSet As New HashSet(Of Tuple(Of String, String))(TupleComparer())
            For Each r As DataRow In rules.Rows
                fileSet.Add(NormalizePath(SafeStr(r("File"))))
                pipeSet.Add(Tuple.Create(Of String, String)(SafeStr(r("File")), SafeStr(r("PipeTypeName"))))
            Next
            Dim sizeCount As Integer = 0
            If sizes IsNot Nothing Then
                sizeCount = sizes.Rows.Count
            End If
            Return $"파일 {fileSet.Count}개, PipeType {pipeSet.Count}, Segment 후보 {rules.Rows.Count}, 사이즈 {sizeCount}"
        End Function

        Private Function BuildGroupPayload(groups As List(Of SegmentPmsCheckService.MappingGroup)) As List(Of Object)
            Dim list As New List(Of Object)()
            If groups Is Nothing Then
                Return list
            End If
            For Each g In groups
                Dim usages As New List(Of Object)()
                For Each u In g.Usages
                    usages.Add(New With {
                        .file = u.File,
                        .pipeType = u.PipeTypeName,
                        .ruleIndex = u.RuleIndex,
                        .segmentId = u.SegmentId,
                        .segmentKey = u.SegmentKey
                    })
                Next
                list.Add(New With {
                    .groupKey = g.GroupKey,
                    .displayKey = g.DisplayKey,
                    .normalizedKey = g.NormalizedKey,
                    .usageSummary = g.UsageSummary,
                    .fileCount = g.FileCount,
                    .pipeTypeCount = g.PipeTypeCount,
                    .usages = usages,
                    .suggestedClass = g.SuggestedClass,
                    .suggestedSegmentKey = g.SuggestedSegmentKey
                })
            Next
            Return list
        End Function

        Private Function BuildPmsOptions() As List(Of Object)
            Dim list As New List(Of Object)()
            If _pmsRows Is Nothing Then
                Return list
            End If
            For Each r In _pmsRows
                list.Add(New With {.label = $"{r.Class} | {r.SegmentKey}", .cls = r.Class, .segment = r.SegmentKey})
            Next
            Return list
        End Function

        Private Shared Function DataTableToObjects(t As DataTable, Optional maxRows As Integer = Integer.MaxValue) As List(Of Dictionary(Of String, Object))
            Dim list As New List(Of Dictionary(Of String, Object))()
            If t Is Nothing Then
                Return list
            End If
            Dim count As Integer = 0
            For Each r As DataRow In t.Rows
                count += 1
                If count > maxRows Then
                    Exit For
                End If
                Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                For Each c As DataColumn In t.Columns
                    d(c.ColumnName) = r(c)
                Next
                list.Add(d)
            Next
            Return list
        End Function

        Private Shared Function CoerceRowsToDictList(raw As Object) As List(Of Dictionary(Of String, Object))
            Dim res As New List(Of Dictionary(Of String, Object))()
            If raw Is Nothing Then
                Return res
            End If

            Dim en = TryCast(raw, System.Collections.IEnumerable)
            If en Is Nothing OrElse TypeOf raw Is String Then
                Return res
            End If

            For Each it As Object In en
                If it Is Nothing Then
                    Continue For
                End If

                Dim gen As IDictionary(Of String, Object) = TryCast(it, IDictionary(Of String, Object))
                If gen IsNot Nothing Then
                    res.Add(New Dictionary(Of String, Object)(gen, StringComparer.OrdinalIgnoreCase))
                    Continue For
                End If

                Dim idic As System.Collections.IDictionary = TryCast(it, System.Collections.IDictionary)
                If idic IsNot Nothing Then
                    Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                    For Each k As Object In idic.Keys
                        d(SegPmsSafeStr(k)) = idic(k)
                    Next
                    res.Add(d)
                    Continue For
                End If

                Try
                    Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                    Dim props = it.GetType().GetProperties()
                    For Each p In props
                        If p Is Nothing OrElse Not p.CanRead Then
                            Continue For
                        End If
                        d(p.Name) = p.GetValue(it, Nothing)
                    Next
                    If d.Count > 0 Then
                        res.Add(d)
                    End If
                Catch
                End Try
            Next

            Return res
        End Function

        Private Shared Sub AddSheet(wb As IWorkbook,
                                    name As String,
                                    rows As List(Of Dictionary(Of String, Object)),
                                    columns As IList(Of String),
                                    Optional progressChannel As String = Nothing,
                                    Optional ByRef written As Integer = 0,
                                    Optional totalRows As Integer = 0,
                                    Optional doAutoFit As Boolean = False)
            Dim sh = wb.CreateSheet(name)
            If columns Is Nothing OrElse columns.Count = 0 Then
                Return
            End If

            Dim headRow = sh.CreateRow(0)
            For ci As Integer = 0 To columns.Count - 1
                headRow.CreateCell(ci).SetCellValue(columns(ci))
            Next

            Dim numStyle = wb.CreateCellStyle()
            numStyle.DataFormat = wb.CreateDataFormat().GetFormat("0.###############")
            Dim dataRows = If(rows, New List(Of Dictionary(Of String, Object))())
            Dim rIndex As Integer = 1
            For Each item In dataRows
                Dim row = sh.CreateRow(rIndex)
                For ci As Integer = 0 To columns.Count - 1
                    Dim key = columns(ci)
                    Dim v As Object = Nothing
                    If item IsNot Nothing Then
                        item.TryGetValue(key, v)
                    End If
                    Dim cell = row.CreateCell(ci)
                    If v Is Nothing OrElse TypeOf v Is DBNull Then
                        cell.SetCellValue(String.Empty)
                    ElseIf TypeOf v Is Double OrElse TypeOf v Is Single OrElse TypeOf v Is Decimal Then
                        cell.SetCellValue(Convert.ToDouble(v))
                        cell.CellStyle = numStyle
                    ElseIf TypeOf v Is Integer OrElse TypeOf v Is Long OrElse TypeOf v Is Short Then
                        cell.SetCellValue(Convert.ToDouble(v))
                    Else
                        cell.SetCellValue(v.ToString())
                    End If
                Next
                rIndex += 1
                written += 1
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_WRITE", "엑셀 데이터 작성", written, totalRows)
            Next

            ExcelCore.ApplyStandardSheetStyle(wb, sh, headerRowIndex:=0, autoFilter:=True, freezeTopRow:=True, borderAll:=True, autoFit:=doAutoFit)
            ExcelCore.ApplyNumberFormatByHeader(wb, sh, 0, New String() {"ND", "ID", "OD", "PMS ND", "PMS ID", "PMS OD", "PMS_ND", "PMS_ID", "PMS_OD", "Diff_ID", "Diff_OD", "ND_mm", "ID_mm", "OD_mm"}, "0.####################")
            ExcelCore.ApplyResultFillByHeader(wb, sh, 0)
        End Sub

        Private Shared Function DictListToDataTable(rows As List(Of Dictionary(Of String, Object)), tableName As String) As DataTable
            Dim t As New DataTable(tableName)
            If rows Is Nothing OrElse rows.Count = 0 Then
                Return t
            End If

            Dim cols As New List(Of String)()
            Dim colSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each src In rows
                If src Is Nothing Then
                    Continue For
                End If
                For Each k In src.Keys
                    If colSet.Add(k) Then
                        cols.Add(k)
                    End If
                Next
            Next

            For Each c In cols
                t.Columns.Add(c, GetType(Object))
            Next

            For Each src In rows
                Dim r = t.NewRow()
                For Each c In cols
                    Dim v As Object = Nothing
                    If src IsNot Nothing Then
                        src.TryGetValue(c, v)
                    End If
                    r(c) = If(v, Nothing)
                Next
                t.Rows.Add(r)
            Next
            Return t
        End Function

        Private Shared Sub SaveWorkbookSafe(wb As IWorkbook, outPath As String)
            If wb Is Nothing OrElse String.IsNullOrWhiteSpace(outPath) Then
                Return
            End If

            Dim tmpPath As String = outPath & ".tmp"
            Dim dir As String = Path.GetDirectoryName(outPath)
            If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Try
                If File.Exists(tmpPath) Then
                    File.Delete(tmpPath)
                End If

                Using ms As New MemoryStream()
                    wb.Write(ms)
                    File.WriteAllBytes(tmpPath, ms.ToArray())
                End Using

                Try
                    If File.Exists(outPath) Then
                        Try
                            File.Replace(tmpPath, outPath, Nothing)
                        Catch
                            File.Delete(outPath)
                            File.Move(tmpPath, outPath)
                        End Try
                    Else
                        File.Move(tmpPath, outPath)
                    End If
                Finally
                    If File.Exists(tmpPath) Then
                        File.Delete(tmpPath)
                    End If
                End Try
            Catch
                If File.Exists(tmpPath) Then
                    Try
                        File.Delete(tmpPath)
                    Catch
                    End Try
                End If
                Throw
            End Try
        End Sub

        Private Shared Sub WaitForFileReady(path As String)
            If String.IsNullOrWhiteSpace(path) Then
                Return
            End If
            For i As Integer = 0 To 4
                Try
                    If File.Exists(path) Then
                        Dim info As New FileInfo(path)
                        If info.Length > 0 Then
                            Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)
                            End Using
                            Exit For
                        End If
                    End If
                Catch
                End Try
                Thread.Sleep(150)
            Next
        End Sub

        Private Shared Function NormalizePath(p As String) As String
            If String.IsNullOrWhiteSpace(p) Then
                Return String.Empty
            End If
            Try
                Return Path.GetFullPath(p)
            Catch
                Return p
            End Try
        End Function

        Private Shared Function GetDictValue(dict As Dictionary(Of String, Object), key As String) As Object
            If dict Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return Nothing
            End If
            Dim v As Object = Nothing
            If dict.TryGetValue(key, v) Then
                Return v
            End If
            Return Nothing
        End Function

        Private Shared Function SafeStr(o As Object) As String
            If o Is Nothing OrElse o Is DBNull.Value Then
                Return String.Empty
            End If
            Return Convert.ToString(o)
        End Function

        Private Shared Function SegPmsSafeStr(o As Object) As String
            Return SafeStr(o)
        End Function

        Private Shared Function TupleComparer() As IEqualityComparer(Of Tuple(Of String, String))
            Return New TupleComparerImpl()
        End Function

        Private Class TupleComparerImpl
            Implements IEqualityComparer(Of Tuple(Of String, String))

            Public Overloads Function Equals(x As Tuple(Of String, String), y As Tuple(Of String, String)) As Boolean Implements IEqualityComparer(Of Tuple(Of String, String)).Equals
                If x Is y Then
                    Return True
                End If
                If x Is Nothing OrElse y Is Nothing Then
                    Return False
                End If
                Return String.Equals(x.Item1, y.Item1, StringComparison.OrdinalIgnoreCase) AndAlso String.Equals(x.Item2, y.Item2, StringComparison.OrdinalIgnoreCase)
            End Function

            Public Overloads Function GetHashCode(obj As Tuple(Of String, String)) As Integer Implements IEqualityComparer(Of Tuple(Of String, String)).GetHashCode
                If obj Is Nothing Then
                    Return 0
                End If
                Return (If(obj.Item1, String.Empty).ToLowerInvariant().GetHashCode() Xor (If(obj.Item2, String.Empty).ToLowerInvariant().GetHashCode() << 3))
            End Function
        End Class

        Private Sub ReportProgress(total As Integer, index As Integer, stage As String, message As String, filePath As String)
            Dim stepFraction As Double
            Select Case (If(stage, String.Empty).ToLowerInvariant())
                Case "open"
                    stepFraction = 0.0R
                Case "extract"
                    stepFraction = 0.33R
                Case "route"
                    stepFraction = 0.66R
                Case "save"
                    stepFraction = 0.9R
                Case "finish", "done"
                    stepFraction = 1.0R
                Case "error"
                    stepFraction = 1.0R
                Case Else
                    stepFraction = 0.0R
            End Select
            Dim safeTotal As Integer = Math.Max(1, Math.Max(total, 1))
            Dim baseIdx As Integer = Math.Max(0, index - 1)
            Dim ratio As Double = (baseIdx + stepFraction) / safeTotal
            Dim pct As Integer = CInt(Math.Max(0, Math.Min(100, Math.Floor(ratio * 100))))
            Dim fileName As String = String.Empty
            If Not String.IsNullOrWhiteSpace(filePath) Then
                Try
                    fileName = Path.GetFileName(filePath)
                Catch
                    fileName = filePath
                End Try
            End If
            SendToWeb("segmentpms:progress", New With {
                .total = total,
                .index = Math.Max(0, index),
                .percent = pct,
                .file = fileName,
                .stage = stage,
                .message = message
            })
        End Sub

    End Class

End Namespace
