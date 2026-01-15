Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Diagnostics
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Services
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace UI.Hub
    Partial Public Class UiBridgeExternalEvent

        ' export:progress 페이로드
        ' { phase: COLLECT|EXTRACT|EXCEL|DONE|ERROR, message, current, total, phaseProgress, percent }
        Private Shared ReadOnly ExportProgressWeights As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) From {
            {"COLLECT", 0.1},
            {"EXTRACT", 0.7},
            {"EXCEL", 0.05},
            {"EXCEL_INIT", 0.02},
            {"EXCEL_WRITE", 0.11},
            {"EXCEL_SAVE", 0.02},
            {"AUTOFIT", 0.0}
        }
        Private Shared ReadOnly ExportProgressOrder As String() = {"COLLECT", "EXTRACT", "EXCEL", "EXCEL_INIT", "EXCEL_WRITE", "EXCEL_SAVE", "AUTOFIT"}
        Private Shared ExportProgressLastSent As DateTime = DateTime.MinValue
        Private Shared ExportProgressLastPct As Double = 0.0
        Private Shared ExportProgressLastRow As Integer = 0
        Private Shared ReadOnly ExportProgressGate As New Object()

        ' ========== Export: 폴더 선택 ==========
        Private Sub HandleExportBrowse()
            Using dlg As New System.Windows.Forms.FolderBrowserDialog()
                Dim r = dlg.ShowDialog()
                If r = System.Windows.Forms.DialogResult.OK Then
                    Dim files As String() = Directory.GetFiles(dlg.SelectedPath, "*.rvt", SearchOption.AllDirectories)
                    _host?.SendToWeb("export:files", New With {.files = files})
                End If
            End Using
        End Sub

        ' ========== Export: RVT 파일 추가(파일 대화상자) ==========
        Private Sub HandleExportAddRvtFiles()
            Using dlg As New System.Windows.Forms.OpenFileDialog()
                dlg.Filter = "Revit Project (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.Title = "Export Points 대상 RVT 선택"
                dlg.RestoreDirectory = True

                If dlg.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Return

                Dim files As New List(Of String)()
                Dim dedup As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each p In dlg.FileNames
                    If String.IsNullOrWhiteSpace(p) Then Continue For
                    If dedup.Add(p) Then files.Add(p)
                Next

                If files.Count > 0 Then
                    _host?.SendToWeb("export:rvt-files", New With {.files = files})
                End If
            End Using
        End Sub

        ' ========== Export: 미리보기 ==========
        Private Sub HandleExportPreview(app As UIApplication, payload As Dictionary(Of String, Object))
            ResetExportProgressState()
            Try
                Dim files = ExtractStringListLocal(payload, "files")
                ReportExportProgress("COLLECT", "파일 목록 준비 중", 0, If(files, New List(Of String)()).Count, 0.0, True)
                Dim rows = TryCallExportPointsService(app, files)
                If rows Is Nothing Then
                    ReportExportProgress("ERROR", "Export Points 서비스가 준비되지 않았습니다.", 0, 0, 0.0, True)
                    _host?.SendToWeb("revit:error", New With {.message = "Export Points 서비스가 준비되지 않았습니다."})
                    _host?.SendToWeb("export:previewed", New With {.rows = New List(Of Dictionary(Of String, Object))()})
                    Return
                End If
                rows = rows.Select(Function(r) AdaptExportRow(r)).ToList()
                Export_LastExportRows = rows
                _host?.SendToWeb("export:previewed", New With {.rows = rows})
            Catch ex As Exception
                ReportExportProgress("ERROR", ex.Message, 0, 0, 0.0, True)
                _host?.SendToWeb("revit:error", New With {.message = "미리보기 실패: " & ex.Message})
                _host?.SendToWeb("export:previewed", New With {.rows = New List(Of Dictionary(Of String, Object))()})
            End Try
        End Sub

        ' ========== Export: 엑셀 내보내기 ==========
        Private Sub HandleExportSaveExcel(payload As Dictionary(Of String, Object))
            ResetExportProgressState()
            Try
                Dim doAutoFit As Boolean = ParseExcelMode(payload)
                Dim unit As String = ExtractUnit(payload)
                Dim rows = TryGetRowsFromPayload(payload)
                If rows Is Nothing OrElse rows.Count = 0 Then rows = Export_LastExportRows
                If rows Is Nothing Then rows = New List(Of Dictionary(Of String, Object))()
                Dim total As Integer = rows.Count
                ReportExportProgress("EXCEL", "엑셀 내보내기 준비 중", 0, total, 0.0, True)

                Dim dt = BuildExportDataTableFromRows(rows, unit, True)
                Dim todayToken As String = Date.Now.ToString("yyMMdd")
                Dim defaultName As String = $"{todayToken}_좌표 추출 결과.xlsx"
                Dim savePath As String = SaveExcelWithDialog(dt, defaultName, doAutoFit)

                If Not String.IsNullOrEmpty(savePath) Then
                    ReportExportProgress("DONE", "엑셀 내보내기 완료", total, total, 1.0, True)
                    _host?.SendToWeb("export:saved", New With {.path = savePath})
                Else
                    ReportExportProgress("DONE", "엑셀 내보내기가 취소되었습니다.", total, total, 1.0, True)
                End If
            Catch ex As Exception
                ReportExportProgress("ERROR", ex.Message, 0, 0, 0.0, True)
                _host?.SendToWeb("revit:error", New With {.message = "엑셀 내보내기 실패: " & ex.Message})
            End Try
        End Sub

        ' -------- 서비스 호출/어댑터/테이블 --------
        Private Function TryCallExportPointsService(app As UIApplication, files As List(Of String)) As List(Of Dictionary(Of String, Object))
            Try
                Dim direct = ExportPointsService.Run(app, files, AddressOf HandleExportProgressFromService)
                Return AnyToRows(direct)
            Catch
            End Try

            Dim names = {"KKY_Tool_Revit.Services.ExportPointsService", "Services.ExportPointsService"}
            For Each n In names
                Dim t = FindType(n, "ExportPointsService")
                If t Is Nothing Then Continue For
                Dim m = t.GetMethod("Run", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance)
                If m Is Nothing Then Continue For
                Dim inst As Object = If(m.IsStatic, Nothing, Activator.CreateInstance(t))
                Dim args As Object()
                Dim ps = m.GetParameters()
                If ps IsNot Nothing AndAlso ps.Length >= 3 Then
                    Dim cb As [Delegate] = Nothing
                    Try
                        Dim cbMethod = GetType(UiBridgeExternalEvent).GetMethod("HandleExportProgressFromObject", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
                        cb = [Delegate].CreateDelegate(ps(2).ParameterType, Me, cbMethod)
                    Catch
                    End Try
                    If cb IsNot Nothing Then
                        args = New Object() {app, files, cb}
                    Else
                        args = New Object() {app, files}
                    End If
                Else
                    args = New Object() {app, files}
                End If
                Dim result = m.Invoke(inst, args)
                Return AnyToRows(result)
            Next
            Return Nothing
        End Function

        Private Function AdaptExportRow(r As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
            If r Is Nothing Then Return New Dictionary(Of String, Object)(StringComparer.Ordinal)
            ' 컬럼 명세를 고정 (home.js/export.js 스키마와 일치)
            Dim d As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            d("File") = If(r.ContainsKey("File"), r("File"), If(r.ContainsKey("file"), r("file"), ""))
            d("ProjectPoint_E(mm)") = FirstNonEmpty(r, {"ProjectPoint_E(mm)", "ProjectE", "ProjectPoint_E", "ProjectPoint_E_ft"})
            d("ProjectPoint_N(mm)") = FirstNonEmpty(r, {"ProjectPoint_N(mm)", "ProjectN", "ProjectPoint_N", "ProjectPoint_N_ft"})
            d("ProjectPoint_Z(mm)") = FirstNonEmpty(r, {"ProjectPoint_Z(mm)", "ProjectZ", "ProjectPoint_Z", "ProjectPoint_Z_ft"})
            d("SurveyPoint_E(mm)") = FirstNonEmpty(r, {"SurveyPoint_E(mm)", "SurveyE", "SurveyPoint_E", "SurveyPoint_E_ft"})
            d("SurveyPoint_N(mm)") = FirstNonEmpty(r, {"SurveyPoint_N(mm)", "SurveyN", "SurveyPoint_N", "SurveyPoint_N_ft"})
            d("SurveyPoint_Z(mm)") = FirstNonEmpty(r, {"SurveyPoint_Z(mm)", "SurveyZ", "SurveyPoint_Z", "SurveyPoint_Z_ft"})
            d("TrueNorthAngle(deg)") = FirstNonEmpty(r, {"TrueNorthAngle(deg)", "TrueNorth", "TrueNorthAngle", "TrueNorthAngle_deg"})
            Return d
        End Function

        Private Function BuildExportDataTableFromRows(rows As List(Of Dictionary(Of String, Object)), unit As String, applyConversion As Boolean) As DataTable
            Dim normalizedUnit As String = NormalizeUnit(unit)
            Dim suffix As String = "(ft)"
            If normalizedUnit = "m" Then
                suffix = "(m)"
            ElseIf normalizedUnit = "mm" Then
                suffix = "(mm)"
            End If

            Dim dt As New DataTable("Export")
            Dim headers = {
                "File",
                $"ProjectPoint_E{suffix}", $"ProjectPoint_N{suffix}", $"ProjectPoint_Z{suffix}",
                $"SurveyPoint_E{suffix}", $"SurveyPoint_N{suffix}", $"SurveyPoint_Z{suffix}",
                "TrueNorthAngle(deg)"
            }
            For Each h In headers : dt.Columns.Add(h) : Next
            Dim total As Integer = If(rows IsNot Nothing, rows.Count, 0)
            Dim idx As Integer = 0
            For Each r In rows
                Dim dr = dt.NewRow()
                dr(0) = SafeToString(r, "File")
                dr(1) = FormatCoordForUnit(r, {"ProjectPoint_E(mm)", "ProjectPoint_E(ft)", "ProjectPoint_E(m)", "ProjectE", "ProjectPoint_E"}, normalizedUnit, applyConversion)
                dr(2) = FormatCoordForUnit(r, {"ProjectPoint_N(mm)", "ProjectPoint_N(ft)", "ProjectPoint_N(m)", "ProjectN", "ProjectPoint_N"}, normalizedUnit, applyConversion)
                dr(3) = FormatCoordForUnit(r, {"ProjectPoint_Z(mm)", "ProjectPoint_Z(ft)", "ProjectPoint_Z(m)", "ProjectZ", "ProjectPoint_Z"}, normalizedUnit, applyConversion)
                dr(4) = FormatCoordForUnit(r, {"SurveyPoint_E(mm)", "SurveyPoint_E(ft)", "SurveyPoint_E(m)", "SurveyE", "SurveyPoint_E"}, normalizedUnit, applyConversion)
                dr(5) = FormatCoordForUnit(r, {"SurveyPoint_N(mm)", "SurveyPoint_N(ft)", "SurveyPoint_N(m)", "SurveyN", "SurveyPoint_N"}, normalizedUnit, applyConversion)
                dr(6) = FormatCoordForUnit(r, {"SurveyPoint_Z(mm)", "SurveyPoint_Z(ft)", "SurveyPoint_Z(m)", "SurveyZ", "SurveyPoint_Z"}, normalizedUnit, applyConversion)
                dr(7) = FormatAngleValue(r, "TrueNorthAngle(deg)")
                dt.Rows.Add(dr)
                idx += 1
                Dim progress As Double = If(total > 0, CDbl(idx) / CDbl(total), 1.0)
                ReportExportProgress("EXCEL", "엑셀 데이터 구성", idx, total, progress, False)
            Next
            If total = 0 Then
                ReportExportProgress("EXCEL", "엑셀 데이터 구성", 0, 0, 0.0, False)
            End If
            Return dt
        End Function

        Private Sub HandleExportProgressFromService(info As ExportPointsService.ProgressInfo)
            If info Is Nothing Then Return
            ReportExportProgress(info.Phase, info.Message, info.Current, info.Total, info.PhaseProgress, False)
        End Sub

        Private Sub HandleExportProgressFromObject(info As Object)
            If info Is Nothing Then Return
            Try
                Dim t = info.GetType()
                Dim phase As String = Convert.ToString(t.GetProperty("Phase")?.GetValue(info, Nothing))
                Dim message As String = Convert.ToString(t.GetProperty("Message")?.GetValue(info, Nothing))
                Dim current As Integer = ToIntSafe(t.GetProperty("Current")?.GetValue(info, Nothing))
                Dim total As Integer = ToIntSafe(t.GetProperty("Total")?.GetValue(info, Nothing))
                Dim phaseProgress As Double = ToDoubleSafe(t.GetProperty("PhaseProgress")?.GetValue(info, Nothing))
                ReportExportProgress(phase, message, current, total, phaseProgress, False)
            Catch
            End Try
        End Sub

        ' ==================================================================
        ' Export local helpers (self-contained; no cross-module dependency)
        ' ==================================================================

        ' 마지막 미리보기 결과(엑셀 내보내기 시 payload 없을 때 사용)
        Private Shared Export_LastExportRows As List(Of Dictionary(Of String, Object)) _
            = New List(Of Dictionary(Of String, Object))()

        ' payload에서 string 리스트 추출(e.g., files[])
        Private Shared Function ExtractStringListLocal(payload As Dictionary(Of String, Object), key As String) As List(Of String)
            Dim res As New List(Of String)()
            If payload Is Nothing OrElse Not payload.ContainsKey(key) OrElse payload(key) Is Nothing Then Return res
            Dim v = payload(key)
            Dim arr = TryCast(v, System.Collections.IEnumerable)
            If arr Is Nothing Then
                Dim s As String = TryCast(v, String)
                If Not String.IsNullOrEmpty(s) Then res.Add(s)
                Return res
            End If
            For Each o In arr
                If o Is Nothing Then Continue For
                Dim s = o.ToString()
                If Not String.IsNullOrWhiteSpace(s) Then res.Add(s)
            Next
            Return res
        End Function

        ' 다양한 반환값 → 표준 rows
        Private Shared Function AnyToRows(any As Object) As List(Of Dictionary(Of String, Object))
            Dim result As New List(Of Dictionary(Of String, Object))()
            If any Is Nothing Then Return result

            If TypeOf any Is List(Of Dictionary(Of String, Object)) Then
                Return DirectCast(any, List(Of Dictionary(Of String, Object)))
            End If

            Dim dt As DataTable = TryCast(any, DataTable)
            If dt IsNot Nothing Then
                Return DataTableToRows(dt)
            End If

            Dim ie = TryCast(any, System.Collections.IEnumerable)
            If ie IsNot Nothing Then
                For Each item In ie
                    Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                    Dim dict = TryCast(item, System.Collections.IDictionary)
                    If dict IsNot Nothing Then
                        For Each k In dict.Keys
                            d(k.ToString()) = dict(k)
                        Next
                    Else
                        Dim t = item.GetType()
                        For Each p In t.GetProperties()
                            d(p.Name) = p.GetValue(item, Nothing)
                        Next
                    End If
                    result.Add(d)
                Next
            End If
            Return result
        End Function

        ' 행 딕셔너리에서 컬럼값 안전 추출
        Private Shared Function SafeToString(row As Dictionary(Of String, Object), col As String) As String
            If row Is Nothing Then Return String.Empty
            Dim v As Object = Nothing
            If row.TryGetValue(col, v) AndAlso v IsNot Nothing Then
                Return Convert.ToString(v, Globalization.CultureInfo.InvariantCulture)
            End If
            Return String.Empty
        End Function

        Private Shared Function SafeToDouble(row As Dictionary(Of String, Object), cols As IEnumerable(Of String)) As Double?
            If row Is Nothing OrElse cols Is Nothing Then Return Nothing
            For Each k In cols
                Dim s = SafeToString(row, k)
                Dim d As Double
                If Double.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then
                    Return d
                End If
            Next
            Return Nothing
        End Function

        ' 여러 키 중 첫 번째로 값이 존재하는 항목을 문자열로 반환
        Private Shared Function FirstNonEmpty(row As Dictionary(Of String, Object), keys As IEnumerable(Of String)) As String
            If row Is Nothing OrElse keys Is Nothing Then Return String.Empty
            For Each k In keys
                Dim s = SafeToString(row, k)
                If Not String.IsNullOrEmpty(s) Then Return s
            Next
            Return String.Empty
        End Function

        Private Shared Function NormalizeUnit(unit As String) As String
            Dim u As String = If(unit, "").Trim().ToLowerInvariant()
            If u = "m" OrElse u = "meter" OrElse u = "meters" Then Return "m"
            If u = "mm" OrElse u = "millimeter" OrElse u = "millimeters" Then Return "mm"
            Return "ft"
        End Function

        Private Shared Function UnitFactor(unit As String) As Double
            If unit = "m" Then Return 0.3048
            If unit = "mm" Then Return 304.8
            Return 1.0
        End Function

        Private Shared Function FormatCoordForUnit(row As Dictionary(Of String, Object), keys As IEnumerable(Of String), unit As String, applyConversion As Boolean) As String
            Dim val As Double? = SafeToDouble(row, keys)
            If Not val.HasValue Then Return String.Empty
            Dim v As Double = val.Value
            If applyConversion Then
                v = v * UnitFactor(unit)
            End If
            Return v.ToString("0.####", Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function FormatAngleValue(row As Dictionary(Of String, Object), key As String) As String
            Dim ang As Double?
            Dim s = SafeToString(row, key)
            Dim d As Double
            If Double.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then
                ang = d
            End If
            If ang.HasValue Then
                Return ang.Value.ToString("0.###", Globalization.CultureInfo.InvariantCulture)
            End If
            Return s
        End Function

        Private Shared Function ExtractUnit(payload As Dictionary(Of String, Object)) As String
            If payload Is Nothing Then Return "ft"
            Dim v As Object = Nothing
            If payload.TryGetValue("unit", v) AndAlso v IsNot Nothing Then
                Return NormalizeUnit(Convert.ToString(v, Globalization.CultureInfo.InvariantCulture))
            End If
            Return "ft"
        End Function

        ' DataTable → rows
        Private Shared Function DataTableToRows(dt As DataTable) As List(Of Dictionary(Of String, Object))
            Dim list As New List(Of Dictionary(Of String, Object))()
            If dt Is Nothing Then Return list
            For Each r As DataRow In dt.Rows
                Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                For Each c As DataColumn In dt.Columns
                    d(c.ColumnName) = If(r.IsNull(c), Nothing, r(c))
                Next
                list.Add(d)
            Next
            Return list
        End Function

        ' 로딩된 어셈블리들에서 타입 찾기(정식명/간단명)
        Private Shared Function FindType(fullOrSimple As String, Optional simpleMatch As String = Nothing) As Type
            ' 직접 시도
            Dim t = Type.GetType(fullOrSimple, False)
            If t IsNot Nothing Then Return t
            ' 로드된 어셈블리 순회
            For Each asm In AppDomain.CurrentDomain.GetAssemblies()
                Try
                    t = asm.GetType(fullOrSimple, False)
                    If t IsNot Nothing Then Return t
                    If Not String.IsNullOrEmpty(simpleMatch) Then
                        For Each ti In asm.GetTypes()
                            If String.Equals(ti.Name, simpleMatch, StringComparison.OrdinalIgnoreCase) Then
                                Return ti
                            End If
                        Next
                    End If
                Catch
                End Try
            Next
            Return Nothing
        End Function

        ' payload에서 rows 추출(있으면) — 없으면 빈 리스트
        Private Shared Function TryGetRowsFromPayload(payload As Dictionary(Of String, Object)) As List(Of Dictionary(Of String, Object))
            If payload Is Nothing Then Return New List(Of Dictionary(Of String, Object))()
            If payload.ContainsKey("rows") AndAlso payload("rows") IsNot Nothing Then
                Return AnyToRows(payload("rows"))
            End If
            If payload.ContainsKey("data") AndAlso payload("data") IsNot Nothing Then
                Return AnyToRows(payload("data"))
            End If
            Dim ie = TryCast(payload, System.Collections.IEnumerable)
            If ie IsNot Nothing Then
                Return AnyToRows(ie)
            End If
            Return New List(Of Dictionary(Of String, Object))()
        End Function

        ' DataTable을 저장 대화상자로 엑셀로 저장하고 경로 반환(취소 시 "")
        Private Shared Function SaveExcelWithDialog(dt As DataTable, Optional defaultName As String = "export.xlsx", Optional doAutoFit As Boolean = False) As String
            If dt Is Nothing OrElse dt.Columns.Count = 0 Then Return String.Empty
            Dim dlg As New Microsoft.Win32.SaveFileDialog() With {
                .Filter = "Excel (*.xlsx)|*.xlsx",
                .FileName = defaultName
            }
            Dim ok = dlg.ShowDialog()
            If ok <> True Then Return String.Empty
            Dim path = dlg.FileName
            Dim totalRows As Integer = dt.Rows.Count
            Dim writtenRows As Integer = 0
                ReportExportProgress("EXCEL_INIT", "엑셀 워크북 준비", 0, totalRows, 0.0, True)
                LogAutoFitDecision(doAutoFit, "UiBridgeExternalEvent.SaveExcelWithDialog")
                Try
                    Dim wb As IWorkbook = New XSSFWorkbook()
                    Dim sh = wb.CreateSheet("Export")
                Dim xssf = TryCast(wb, XSSFWorkbook)
                Dim baseStyle As ICellStyle = If(xssf IsNot Nothing, CreateBorderedStyle(xssf), Nothing)
                Dim headerStyle As ICellStyle = If(xssf IsNot Nothing, CreateHeaderStyle(xssf, baseStyle), Nothing)

                ' 헤더
                Dim hr = sh.CreateRow(0)
                For c = 0 To dt.Columns.Count - 1
                    Dim cell = hr.CreateCell(c)
                    cell.SetCellValue(dt.Columns(c).ColumnName)
                    If headerStyle IsNot Nothing Then cell.CellStyle = headerStyle
                Next
                ' 데이터
                Dim rIndex = 1
                For Each dr As DataRow In dt.Rows
                    Dim rr = sh.CreateRow(rIndex) : rIndex += 1
                    For c = 0 To dt.Columns.Count - 1
                        Dim v = If(dr.IsNull(c), "", Convert.ToString(dr(c), Globalization.CultureInfo.InvariantCulture))
                        Dim cell = rr.CreateCell(c)
                        cell.SetCellValue(v)
                        If baseStyle IsNot Nothing Then cell.CellStyle = baseStyle
                    Next
                    writtenRows += 1
                    ReportExportProgress("EXCEL_WRITE", "엑셀 데이터 작성", writtenRows, totalRows, If(totalRows > 0, CDbl(writtenRows) / CDbl(totalRows), 1.0), False)
                Next
                ReportExportProgress("EXCEL_WRITE", "엑셀 데이터 작성", totalRows, totalRows, 1.0, True)
                ' 자동 너비
                If doAutoFit Then
                    For c = 0 To dt.Columns.Count - 1
                        sh.AutoSizeColumn(c)
                    Next
                End If
                ReportExportProgress("EXCEL_SAVE", "엑셀 파일 내보내기", totalRows, totalRows, 1.0, True)
                Using fs As New FileStream(path, FileMode.Create, FileAccess.Write)
                    wb.Write(fs)
                End Using
                Dim autoFitMessage As String = If(doAutoFit, "AutoFit 적용", "빠른 모드: AutoFit 생략")
                If doAutoFit Then
                    ReportExportProgress("AUTOFIT", autoFitMessage, totalRows, totalRows, 1.0, True)
                    Global.KKY_Tool_Revit.Infrastructure.ExcelCore.TryAutoFitWithExcel(path)
                Else
                    ReportExportProgress("AUTOFIT", autoFitMessage, totalRows, totalRows, 1.0, True)
                End If
                ReportExportProgress("DONE", "엑셀 내보내기 완료", totalRows, totalRows, 1.0, True)
                Return path
            Catch ex As Exception
                _host?.SendToWeb("host:error", New With {.message = "엑셀 내보내기 실패: " & ex.Message})
                ReportExportProgress("ERROR", ex.Message, writtenRows, totalRows, 0.0, True)
                Return String.Empty
            End Try
        End Function

        ' 진행률 헬퍼
        Private Shared Function ToIntSafe(obj As Object) As Integer
            If obj Is Nothing Then Return 0
            Try
                Return Convert.ToInt32(obj)
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function ToDoubleSafe(obj As Object) As Double
            If obj Is Nothing Then Return 0.0
            Try
                Return Convert.ToDouble(obj)
            Catch
                Return 0.0
            End Try
        End Function

        Private Shared Sub ResetExportProgressState()
            SyncLock ExportProgressGate
                ExportProgressLastSent = DateTime.MinValue
                ExportProgressLastPct = 0.0
                ExportProgressLastRow = 0
            End SyncLock
        End Sub

        Private Shared Sub ReportExportProgress(phase As String,
                                                message As String,
                                                current As Integer,
                                                total As Integer,
                                                phaseProgress As Double,
                                                Optional force As Boolean = False)
            Dim normalized As String = NormalizeExportPhase(phase)
            Dim pctToSend As Double = 0.0
            Dim shouldSend As Boolean = False
            Dim now As DateTime = DateTime.UtcNow
            SyncLock ExportProgressGate
                Dim computed As Double = ComputeExportPercent(normalized, current, total, phaseProgress, ExportProgressLastPct)
                Dim elapsed As Double = (now - ExportProgressLastSent).TotalMilliseconds
                Dim delta As Double = Math.Abs(computed - ExportProgressLastPct)
                Dim deltaRows As Integer = Math.Abs(current - ExportProgressLastRow)
                Dim important As Boolean = normalized = "DONE" OrElse normalized = "ERROR"
                If force OrElse important OrElse elapsed >= 200.0 OrElse delta >= 1.0 OrElse deltaRows >= 200 Then
                    ExportProgressLastSent = now
                    ExportProgressLastPct = Math.Max(ExportProgressLastPct, computed)
                    ExportProgressLastRow = current
                    pctToSend = ExportProgressLastPct
                    shouldSend = True
                End If
            End SyncLock
            If Not shouldSend Then Return

            _host?.SendToWeb("export:progress", New With {
                .phase = normalized,
                .message = message,
                .current = current,
                .total = total,
                .phaseProgress = Clamp01(phaseProgress),
                .percent = pctToSend
            })
        End Sub

        Private Shared Function ComputeExportPercent(phase As String,
                                                     current As Integer,
                                                     total As Integer,
                                                     phaseProgress As Double,
                                                     lastPct As Double) As Double
            If phase = "DONE" Then Return 100.0
            If phase = "ERROR" Then Return lastPct

            Dim completed As Double = 0.0
            Dim found As Boolean = False
            For Each key In ExportProgressOrder
                If String.Equals(key, phase, StringComparison.OrdinalIgnoreCase) Then
                    found = True
                    Exit For
                End If
                If ExportProgressWeights.ContainsKey(key) Then completed += ExportProgressWeights(key)
            Next
            Dim weight As Double
            If ExportProgressWeights.ContainsKey(phase) Then
                weight = ExportProgressWeights(phase)
            ElseIf Not found Then
                weight = 1.0
                completed = 0.0
            Else
                weight = 0.0
            End If
            Dim ratio As Double = 0.0
            If total > 0 Then ratio = Math.Max(0.0, Math.Min(1.0, CDbl(current) / CDbl(total)))
            ratio = Math.Max(ratio, Clamp01(phaseProgress))

            Dim pct As Double = (completed + weight * ratio) * 100.0
            If pct < lastPct Then Return lastPct
            If pct > 100.0 Then Return 100.0
            Return pct
        End Function

        Private Shared Function NormalizeExportPhase(phase As String) As String
            Dim p As String = If(phase, String.Empty).Trim().ToUpperInvariant()
            If String.IsNullOrEmpty(p) Then Return "EXTRACT"
            Return p
        End Function

        Private Shared Function Clamp01(v As Double) As Double
            If Double.IsNaN(v) OrElse Double.IsInfinity(v) Then Return 0.0
            If v < 0.0 Then Return 0.0
            If v > 1.0 Then Return 1.0
            Return v
        End Function

    End Class
End Namespace
