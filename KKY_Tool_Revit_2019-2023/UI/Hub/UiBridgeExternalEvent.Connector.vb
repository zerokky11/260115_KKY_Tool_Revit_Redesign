Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Diagnostics
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
Imports System.Windows.Forms ' WinForms 다이얼로그 사용
Imports KKY_Tool_Revit.Infrastructure

Namespace UI.Hub
    ' 커넥터 진단 (fix2 이벤트명/스키마 유지)
    Partial Public Class UiBridgeExternalEvent

        ' 최근 로드/실행 결과(엑셀 내보내기 시 기본 소스)
        Private lastConnRows As List(Of Dictionary(Of String, Object)) = Nothing

        ' 전체 커넥터 결과(엑셀 내보내기용) - Total/Detail 분리
        Private _connectorTotalRows As List(Of Dictionary(Of String, Object)) = Nothing
        Private _connectorDetailRows As List(Of Dictionary(Of String, Object)) = Nothing

        ' 추가 추출 파라미터
        Private _connectorExtraParams As List(Of String) = Nothing

        ' 추가 필터
        Private _connectorTargetFilter As String = String.Empty
        Private _connectorExcludeEndDummy As Boolean = False

        ' 디버그 로그를 웹(F12 콘솔)로 보내는 헬퍼
        Private Sub LogDebug(message As String)
            Try
                Dim ts As String = Date.Now.ToString("HH:mm:ss")
                SendToWeb("host:log", New With {
                    .message = $"[{ts}] {message}"
                })
            Catch
                ' 로깅 중 예외는 무시
            End Try
        End Sub

        ' 오류 로그용
        Private Sub LogError(message As String)
            Try
                Dim ts As String = Date.Now.ToString("HH:mm:ss")
                SendToWeb("host:error", New With {
                    .message = $"[{ts}] {message}"
                })
            Catch
            End Try
        End Sub

        Private Sub ReportConnectorProgress(pct As Double, text As String)
            SendToWeb("connector:progress", New With {.pct = pct, .text = text})
        End Sub

        Private Function SafePayloadSnapshot(payload As Object) As String
            If payload Is Nothing Then Return "(null)"
            Try
                Dim dict = TryCast(payload, IDictionary(Of String, Object))
                If dict IsNot Nothing Then
                    Dim parts As New List(Of String)()
                    For Each kv In dict
                        Dim v As Object = kv.Value
                        Dim text As String = If(v Is Nothing, "(null)", v.ToString())
                        parts.Add(kv.Key & "=" & text)
                    Next
                    Return "{" & String.Join(", ", parts) & "}"
                End If
                Return payload.ToString()
            Catch
                Return "(payload)"
            End Try
        End Function

#Region "핸들러 (Core에서 리플렉션으로 호출)"

        ' === connector:run ===
        Private Sub HandleConnectorRun(app As UIApplication, payload As Object)
            Try
                LogDebug("[connector] HandleConnectorRun 진입")
                LogDebug("[connector] payload 수신: " & SafePayloadSnapshot(payload))
                ReportConnectorProgress(0.1R, "커넥터 진단 시작...")

                Dim uidoc = app.ActiveUIDocument
                Dim doc = If(uidoc Is Nothing, Nothing, uidoc.Document)
                If doc Is Nothing Then
                    LogError("[connector] 활성 문서가 없습니다.")
                    SendToWeb("revit:error", New With {.message = "활성 문서가 없습니다."})
                    SendToWeb("connector:done", New With {.ok = False, .message = "활성 문서가 없습니다."})
                    Return
                End If

                _connectorTotalRows = Nothing
                _connectorDetailRows = Nothing

                ' === payload 파싱 ===
                Dim tol As Double = 1.0 ' 기본 1 inch
                Dim unit As String = "inch"
                Dim param As String = "Comments"
                Try
                    Dim vTol = GetProp(payload, "tol")
                    If vTol IsNot Nothing Then tol = Convert.ToDouble(vTol)
                Catch : End Try
                Try
                    Dim vUnit = TryCast(GetProp(payload, "unit"), String)
                    If Not String.IsNullOrEmpty(vUnit) Then unit = vUnit
                Catch : End Try
                Try
                    Dim vParam = TryCast(GetProp(payload, "param"), String)
                    If Not String.IsNullOrEmpty(vParam) Then param = vParam
                Catch : End Try
                _connectorExtraParams = ParseExtraParams(TryCast(GetProp(payload, "extraParams"), String))
                Try
                    Dim vFilter = TryCast(GetProp(payload, "targetFilter"), String)
                    _connectorTargetFilter = If(vFilter, String.Empty)
                Catch
                    _connectorTargetFilter = String.Empty
                End Try
                Try
                    Dim vExclude = GetProp(payload, "excludeEndDummy")
                    If vExclude IsNot Nothing Then
                        _connectorExcludeEndDummy = Convert.ToBoolean(vExclude)
                    Else
                        _connectorExcludeEndDummy = False
                    End If
                Catch
                    _connectorExcludeEndDummy = False
                End Try
                LogDebug($"[connector] 파라미터 파싱 완료 (tol={tol}, unit={unit}, param={param}, extra={String.Join(",", _connectorExtraParams)} )")

                ' === 단위 변환 → feet ===
                Dim tolFt As Double = 0.0
                Dim u = (If(unit, "inch")).Trim().ToLowerInvariant()
                If u = "mm" OrElse u = "millimeter" OrElse u = "millimeters" Then
                    tolFt = tol / 304.8R
                Else
                    ' inch 또는 기타 → inch 가정
                    tolFt = tol / 12.0R
                End If
                If tolFt < 0.0000001 Then tolFt = 0.0000001R
                LogDebug($"[connector] tolFt 계산 완료: {tolFt}")

                ' === 서비스 호출 ===
                LogDebug("[connector] 커넥터 수집/진단 실행 시작")
                Const PREVIEW_LIMIT As Integer = 150
                Dim rows As List(Of Dictionary(Of String, Object)) = Nothing
                Try
                    rows = Services.ConnectorDiagnosticsService.Run(app, tolFt, param, _connectorExtraParams, _connectorTargetFilter, _connectorExcludeEndDummy, AddressOf ReportConnectorProgress)
                Catch ex As Exception
                    ' 네임스페이스 변동 대비 리플렉션 재시도
                    Try
                        Dim t = Type.GetType("KKY_Tool_Revit.Services.ConnectorDiagnosticsService, KKY_Tool_Revit")
                        If t Is Nothing Then t = Type.GetType("ConnectorDiagnosticsService")
                        If t IsNot Nothing Then
                            Dim m = t.GetMethod("Run", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static)
                            If m IsNot Nothing Then
                                Dim args As Object()
                                Dim ps = m.GetParameters()
                                If ps.Length >= 7 Then
                                    args = New Object() {app, tolFt, param, _connectorExtraParams, _connectorTargetFilter, _connectorExcludeEndDummy, CType(AddressOf ReportConnectorProgress, Action(Of Double, String))}
                                ElseIf ps.Length >= 6 Then
                                    args = New Object() {app, tolFt, param, _connectorExtraParams, _connectorTargetFilter, _connectorExcludeEndDummy}
                                ElseIf ps.Length = 4 Then
                                    args = New Object() {app, tolFt, param, _connectorExtraParams}
                                Else
                                    args = New Object() {app, tolFt, param}
                                End If
                                rows = CType(m.Invoke(Nothing, args), List(Of Dictionary(Of String, Object)))
                            End If
                        End If
                    Catch
                    End Try
                End Try
                If rows Is Nothing Then rows = New List(Of Dictionary(Of String, Object))()

                Try
                    Dim svcLog = Services.ConnectorDiagnosticsService.LastDebug
                    If svcLog IsNot Nothing Then
                        For Each line In svcLog
                            LogDebug("[connector][svc] " & line)
                        Next
                    End If
                Catch
                End Try

                Dim totalRows = BuildTotalRows(rows)
                Dim filteredRows = totalRows.Where(Function(r) ShouldIncludeRow(r)).ToList()

                Dim mismatchAll = filteredRows.Where(Function(r) IsMismatchRow(r)).ToList()
                Dim nearAll = filteredRows.Where(Function(r) IsNearConnection(r)).ToList()

                Dim mismatchPreview As List(Of Dictionary(Of String, Object)) = mismatchAll.Take(PREVIEW_LIMIT).ToList()
                Dim nearPreview As List(Of Dictionary(Of String, Object)) = nearAll.Take(PREVIEW_LIMIT).ToList()
                Dim previewRows As List(Of Dictionary(Of String, Object)) = filteredRows.Take(PREVIEW_LIMIT).ToList()

                _connectorTotalRows = filteredRows
                _connectorDetailRows = rows
                lastConnRows = filteredRows

                Dim mismatchCount As Integer = mismatchAll.Count
                Dim okCount As Integer = Math.Max(filteredRows.Count - mismatchCount, 0)
                LogDebug($"[connector] 규칙/비교 로직 적용 완료: 정상 {okCount}개, 경고/오류 {mismatchCount}개")
                LogDebug($"[connector] 커넥터 수집 완료: 결과 행 {filteredRows.Count}개 (Mismatch={mismatchAll.Count}, Near={nearAll.Count})")

                LogDebug("[connector] 결과 전송 준비 완료, connector:done/connector:loaded emit 직전")
                Dim hasMore As Boolean = filteredRows.Count > PREVIEW_LIMIT
                SendToWeb("connector:loaded", New With {
                    .rows = previewRows,
                    .total = filteredRows.Count,
                    .previewCount = previewRows.Count,
                    .hasMore = hasMore,
                    .mismatch = New With {
                        .rows = mismatchPreview,
                        .total = mismatchAll.Count,
                        .previewCount = mismatchPreview.Count,
                        .hasMore = mismatchAll.Count > PREVIEW_LIMIT
                    },
                    .near = New With {
                        .rows = nearPreview,
                        .total = nearAll.Count,
                        .previewCount = nearPreview.Count,
                        .hasMore = nearAll.Count > PREVIEW_LIMIT
                    },
                    .extraParams = _connectorExtraParams
                })
                SendToWeb("connector:done", New With {
                    .rows = previewRows,
                    .total = filteredRows.Count,
                    .previewCount = previewRows.Count,
                    .hasMore = hasMore,
                    .mismatch = New With {
                        .rows = mismatchPreview,
                        .total = mismatchAll.Count,
                        .previewCount = mismatchPreview.Count,
                        .hasMore = mismatchAll.Count > PREVIEW_LIMIT
                    },
                    .near = New With {
                        .rows = nearPreview,
                        .total = nearAll.Count,
                        .previewCount = nearPreview.Count,
                        .hasMore = nearAll.Count > PREVIEW_LIMIT
                    },
                    .extraParams = _connectorExtraParams
                })
                LogDebug("[connector] 결과 전송 완료, connector:done emit")
                LogDebug("[connector] HandleConnectorRun 정상 종료")

            Catch ex As Exception
                LogError("[connector] 검사 중 예외 발생: " & ex.ToString())
                SendToWeb("connector:done", New With {.ok = False, .message = ex.Message})
                SendToWeb("revit:error", New With {.message = "실행 실패: " & ex.Message})
            Finally
                ReportConnectorProgress(0R, String.Empty)
            End Try
        End Sub

        ' === connector:save-excel ===
        Private Sub HandleConnectorSaveExcel(app As UIApplication, payload As Object)
            Try
                Dim rows As List(Of Dictionary(Of String, Object)) = _connectorTotalRows
                If rows Is Nothing OrElse rows.Count = 0 Then rows = TryGetRowsFromPayload(payload)
                If rows Is Nothing OrElse rows.Count = 0 Then rows = lastConnRows

                If rows Is Nothing OrElse rows.Count = 0 Then
                    SendToWeb("revit:error", New With {.message = "저장할 데이터가 없습니다."})
                    Return
                End If

                Dim filteredTotal = rows.Where(AddressOf ShouldExportToExcel).ToList()

                If filteredTotal Is Nothing OrElse filteredTotal.Count = 0 Then
                    System.Windows.Forms.MessageBox.Show("이슈 항목이 없습니다.", "검토 결과", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim mismatchCount As Integer = CountMismatches(filteredTotal)

                Dim doAutoFit As Boolean = ParseExcelMode(payload)
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset("connector:progress")
                Dim saved As String = SaveRowsToExcel(filteredTotal, mismatchCount, _connectorExtraParams, doAutoFit, "connector:progress")

                SendToWeb("connector:saved", New With {.path = saved})

            Catch ex As Exception
                SendToWeb("revit:error", New With {.message = "엑셀 내보내기 실패: " & ex.Message})
            End Try
        End Sub

#End Region

#Region "엑셀 입출력/유틸 (스키마 불변)"

        Private Function TryReadExcelAsDataTable() As DataTable
            Using ofd As New OpenFileDialog()
                ofd.Filter = "Excel Files|*.xlsx;*.xls"
                ofd.Multiselect = False
                If ofd.ShowDialog() <> DialogResult.OK Then Return Nothing

                Dim filePath = ofd.FileName
                Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Dim wb As IWorkbook
                    Dim ext As String = System.IO.Path.GetExtension(filePath)
                    If ext IsNot Nothing AndAlso ext.Equals(".xls", StringComparison.OrdinalIgnoreCase) Then
                        wb = New HSSFWorkbook(fs)
                    Else
                        wb = New XSSFWorkbook(fs)
                    End If

                    Dim sh = wb.GetSheetAt(0)
                    Dim dt As New DataTable()

                    ' 헤더
                    Dim hr = sh.GetRow(sh.FirstRowNum)
                    If hr Is Nothing Then Return Nothing
                    For c = 0 To hr.LastCellNum - 1
                        Dim name = If(hr.GetCell(c)?.ToString(), $"C{c + 1}")
                        dt.Columns.Add(name)
                    Next

                    ' 데이터
                    For r = sh.FirstRowNum + 1 To sh.LastRowNum
                        Dim sr = sh.GetRow(r)
                        If sr Is Nothing Then Continue For
                        Dim dr = dt.NewRow()
                        For c = 0 To dt.Columns.Count - 1
                            dr(c) = If(sr.GetCell(c)?.ToString(), "")
                        Next
                        dt.Rows.Add(dr)
                    Next

                    Return dt
                End Using
            End Using
        End Function

        Private Function DataTableRows(dt As DataTable) As List(Of Dictionary(Of String, Object))
            Dim list As New List(Of Dictionary(Of String, Object))()
            For Each r As DataRow In dt.Rows
                Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                For Each c As DataColumn In dt.Columns
                    d(c.ColumnName) = If(r(c), "")
                Next
                list.Add(d)
            Next
            Return list
        End Function

        Private Function TryGetRowsFromPayload(payload As Object) As List(Of Dictionary(Of String, Object))
            If payload Is Nothing Then Return Nothing
            Try
                Dim d = TryCast(payload, IDictionary(Of String, Object))
                If d IsNot Nothing AndAlso d.ContainsKey("rows") Then
                    Return TryCast(d("rows"), List(Of Dictionary(Of String, Object)))
                End If
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function ReadField(r As Dictionary(Of String, Object), key As String) As String
            If r Is Nothing Then Return String.Empty
            If r.ContainsKey(key) AndAlso r(key) IsNot Nothing Then
                Return r(key).ToString()
            End If
            Return String.Empty
        End Function

        Private Shared Function ReadFieldInsensitive(r As Dictionary(Of String, Object), key As String) As String
            If r Is Nothing Then Return String.Empty
            For Each kv In r
                If kv.Key Is Nothing Then Continue For
                If String.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) Then
                    If kv.Value Is Nothing Then Return String.Empty
                    Return kv.Value.ToString()
                End If
            Next
            Return String.Empty
        End Function

        Private Shared Function ParseExtraParams(raw As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(raw) Then Return result

            Dim parts = raw.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each part In parts
                Dim name = part.Trim()
                If String.IsNullOrEmpty(name) Then Continue For
                If seen.Add(name) Then result.Add(name)
            Next
            Return result
        End Function

        Private Shared Function IsNearConnection(r As Dictionary(Of String, Object)) As Boolean
            Dim conn As String = ReadField(r, "ConnectionType")
            If String.IsNullOrEmpty(conn) Then conn = ReadField(r, "Connection Type")
            If String.Equals(conn, "Near", StringComparison.OrdinalIgnoreCase) Then Return True
            If conn.IndexOf("Proximity", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            Dim status As String = ReadField(r, "Status")
            If String.Equals(status, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        Private Shared Function IsMismatchRow(r As Dictionary(Of String, Object)) As Boolean
            Dim status As String = ReadField(r, "Status")
            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        Private Shared Function IsMismatchStatus(status As String) As Boolean
            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        Private Shared Function IsIssueStatus(status As String) As Boolean
            If String.IsNullOrEmpty(status) Then Return False
            If String.Equals(status, "ERROR", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        Private Shared Function IsMatchOrOk(status As String) As Boolean
            If String.IsNullOrEmpty(status) Then Return False
            If String.Equals(status, "Match", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        Private Shared Function ShouldExportToExcel(row As Dictionary(Of String, Object)) As Boolean
            If row Is Nothing Then Return False
            Dim status As String = ReadField(row, "Status")
            Return IsIssueStatus(status)
        End Function

        Private Shared Function ShouldIncludeRow(r As Dictionary(Of String, Object)) As Boolean
            If IsMismatchRow(r) Then Return True
            If IsNearConnection(r) Then Return True
            Dim status As String = ReadField(r, "Status")
            Return IsIssueStatus(status)
        End Function

        Private Function CountMismatches(rows As List(Of Dictionary(Of String, Object))) As Integer
            If rows Is Nothing Then Return 0
            Dim cnt As Integer = 0
            For Each row In rows
                Dim status As String = Nothing
                If row IsNot Nothing AndAlso row.ContainsKey("Status") AndAlso row("Status") IsNot Nothing Then
                    status = row("Status").ToString()
                End If
                If IsIssueStatus(status) Then
                    cnt += 1
                End If
            Next
            Return cnt
        End Function

        Private Function SaveRowsToExcel(totalRows As List(Of Dictionary(Of String, Object)),
                                         Optional mismatchCount As Integer = -1,
                                         Optional extraParams As List(Of String) = Nothing,
                                         Optional doAutoFit As Boolean = False,
                                         Optional progressChannel As String = Nothing) As String
            Dim todayToken As String = Date.Now.ToString("yyMMdd")
            Dim count As Integer = If(mismatchCount < 0, CountMismatches(totalRows), mismatchCount)
            Dim defaultName As String = $"{todayToken}_커넥터기반 속성값 검토 결과_{count}개.xlsx"
            Dim totalCount As Integer = If(totalRows, New List(Of Dictionary(Of String, Object))()).Count
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset(progressChannel)
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_INIT", "엑셀 워크북 준비", 0, totalCount, Nothing, True)
            LogAutoFitDecision(doAutoFit, "UiBridgeExternalEvent.SaveRowsToExcel")
            Dim written As Integer = 0

            Try
                Using sfd As New SaveFileDialog()
                    sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                    sfd.FileName = defaultName
                    If sfd.ShowDialog() <> DialogResult.OK Then Throw New OperationCanceledException()

                    Dim savePath = sfd.FileName
                    Using fs As New FileStream(savePath, FileMode.Create, FileAccess.Write, FileShare.None)
                        Using wb As New XSSFWorkbook()
                            Dim extrasSource As List(Of String) = If(extraParams, New List(Of String)())
                            If (extrasSource Is Nothing OrElse extrasSource.Count = 0) AndAlso totalRows IsNot Nothing AndAlso totalRows.Count > 0 Then
                                extrasSource = InferExtrasFromRow(totalRows(0))
                            End If
                            Dim extrasHeaders = BuildExtraHeaders(extrasSource)

                            Dim headersTotal = BuildHeaders(extrasHeaders)
                            Dim baseStyle As ICellStyle = CreateBorderedStyle(wb)
                            Dim headerStyle As ICellStyle = CreateHeaderStyle(wb, baseStyle)
                            Dim mismatchStyle As ICellStyle = CreateFillStyle(wb, baseStyle, New Byte() {&HF9, &HD3, &HD7}) ' light red
                            Dim matchStyle As ICellStyle = CreateFillStyle(wb, baseStyle, New Byte() {&HD6, &HEF, &HD6})   ' light green
                            Dim nearStyle As ICellStyle = CreateFillStyle(wb, baseStyle, New Byte() {&HFA, &HF3, &HD1})    ' light yellow
                            Dim errorStyle As ICellStyle = CreateFillStyle(wb, baseStyle, New Byte() {&HFD, &HEA, &HCC})   ' light orange

                            Dim totalBase = totalRows.Select(Function(r) StripExtras(r, extrasHeaders)).ToList()
                            WriteSheet(wb, "Total", headersTotal, totalBase, headerStyle, baseStyle, matchStyle, mismatchStyle, nearStyle, errorStyle, progressChannel, written, totalCount, doAutoFit)

                            wb.Write(fs)
                        End Using
                    End Using
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_SAVE", "파일 저장 중", totalCount, totalCount, Nothing, True)
                    Dim autoFitMessage As String = If(doAutoFit, "AutoFit 적용", "빠른 모드: AutoFit 생략")
                    If doAutoFit Then
                        Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", autoFitMessage, totalCount, totalCount, Nothing, True)
                        Global.KKY_Tool_Revit.Infrastructure.ExcelCore.TryAutoFitWithExcel(savePath)
                    Else
                        Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", autoFitMessage, totalCount, totalCount, Nothing, True)
                    End If
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "DONE", "엑셀 내보내기 완료", totalCount, totalCount, 100.0R, True)
                    Return savePath
                End Using
            Catch ex As OperationCanceledException
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "DONE", "엑셀 내보내기가 취소되었습니다.", written, totalCount, 100.0R, True)
                Return String.Empty
            Catch ex As Exception
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "ERROR", ex.Message, written, totalCount, Nothing, True)
                Throw
            End Try
        End Function

        ' 테두리/헤더/색상 스타일 헬퍼 (같은 워크북 내 공유)
        Private Shared Function CreateBorderedStyle(wb As XSSFWorkbook) As ICellStyle
            Dim st As ICellStyle = wb.CreateCellStyle()
            st.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin
            st.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin
            st.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin
            st.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin
            Return st
        End Function

        Private Shared Function CreateHeaderStyle(wb As XSSFWorkbook, baseStyle As ICellStyle) As ICellStyle
            Dim st As XSSFCellStyle = CType(wb.CreateCellStyle(), XSSFCellStyle)
            st.CloneStyleFrom(baseStyle)
            st.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground
            ' use available ctor for current NPOI version
            st.SetFillForegroundColor(New XSSFColor(New Byte() {&H2A, &H3B, &H52}))

            Dim f As XSSFFont = CType(wb.CreateFont(), XSSFFont)
            f.IsBold = True
            f.Color = IndexedColors.White.Index
            st.SetFont(f)
            Return st
        End Function

        Private Shared Function CreateFillStyle(wb As XSSFWorkbook, baseStyle As ICellStyle, rgb As Byte()) As ICellStyle
            Dim st As XSSFCellStyle = CType(wb.CreateCellStyle(), XSSFCellStyle)
            st.CloneStyleFrom(baseStyle)
            st.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground
            ' same ctor form as header style to keep compatibility with the current NPOI version
            st.SetFillForegroundColor(New XSSFColor(rgb))
            Return st
        End Function

        Private Shared Function SafeCellString(row As Dictionary(Of String, Object), key As String) As String
            If row Is Nothing OrElse String.IsNullOrEmpty(key) OrElse Not row.ContainsKey(key) Then Return String.Empty
            Dim v = row(key)
            Return If(v Is Nothing, String.Empty, v.ToString())
        End Function

        Private Shared Function BuildBaseHeaders() As List(Of String)
            Return New List(Of String) From {
                "Id1", "Id2", "Category1", "Category2", "Family1", "Family2", "Distance (inch)", "ConnectionType", "ParamName", "Value1", "Value2", "ParamCompare", "Status", "ErrorMessage"
            }
        End Function

        Private Shared Function BuildExtraHeaders(extras As IList(Of String)) As List(Of String)
            Dim list As New List(Of String)()
            If extras Is Nothing Then Return list

            For Each name In extras
                list.Add($"{name}(ID1)")
                list.Add($"{name}(ID2)")
            Next

            Return list
        End Function

        Private Shared Function BuildHeaders(extras As IList(Of String)) As List(Of String)
            Dim headers = BuildBaseHeaders()
            If extras IsNot Nothing Then
                For Each name In extras
                    headers.Add(name)
                Next
            End If
            Return headers
        End Function

        Private Shared Function InferExtrasFromRow(row As Dictionary(Of String, Object)) As List(Of String)
            Dim extras As New List(Of String)()
            If row Is Nothing Then Return extras

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each key In row.Keys
                If key Is Nothing Then Continue For
                If key.EndsWith("(ID1)", StringComparison.OrdinalIgnoreCase) Then
                    Dim name = key.Substring(0, key.Length - "(ID1)".Length)
                    If seen.Add(name) Then extras.Add(name)
                End If
            Next
            Return extras
        End Function

        Private Shared Function BuildReviewRows(rows As List(Of Dictionary(Of String, Object)), extras As IList(Of String)) As List(Of Dictionary(Of String, Object))
            If rows Is Nothing Then Return New List(Of Dictionary(Of String, Object))()

            Dim dirRows As New Dictionary(Of String, List(Of Dictionary(Of String, Object)))(StringComparer.Ordinal)
            Dim pairMembers As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

            For Each row In rows
                If row Is Nothing Then Continue For
                Dim id1 = ToIntLocal(ReadField(row, "Id1"))
                Dim id2 = ToIntLocal(ReadField(row, "Id2"))
                If id1 = 0 AndAlso id2 = 0 Then Continue For

                Dim dirKey = $"{id1}->{id2}"
                Dim pairKey = If(id1 <= id2, $"{id1}_{id2}", $"{id2}_{id1}")

                If Not dirRows.ContainsKey(dirKey) Then dirRows(dirKey) = New List(Of Dictionary(Of String, Object))()
                dirRows(dirKey).Add(row)

                If Not pairMembers.ContainsKey(pairKey) Then pairMembers(pairKey) = New HashSet(Of String)(StringComparer.Ordinal)
                pairMembers(pairKey).Add(dirKey)
            Next

            Dim result As New List(Of Dictionary(Of String, Object))()

            For Each kv In pairMembers
                Dim ids = kv.Key.Split("_"c)
                If ids.Length <> 2 Then Continue For
                Dim a As Integer = ToIntLocal(ids(0))
                Dim b As Integer = ToIntLocal(ids(1))
                Dim keyAB = $"{a}->{b}"
                Dim keyBA = $"{b}->{a}"

                Dim bestAB = SelectBestRow(dirRows, keyAB)
                Dim bestBA = SelectBestRow(dirRows, keyBA)

                If bestAB Is Nothing AndAlso bestBA IsNot Nothing Then bestAB = SwapRow(bestBA)
                If bestBA Is Nothing AndAlso bestAB IsNot Nothing Then bestBA = SwapRow(bestAB)

                If bestAB IsNot Nothing Then result.Add(AppendExtrasForId1(bestAB, extras))
                If bestBA IsNot Nothing Then result.Add(AppendExtrasForId1(bestBA, extras))
            Next

            Return result
        End Function

        Private Shared Function StripExtras(row As Dictionary(Of String, Object), Optional extras As IList(Of String) = Nothing) As Dictionary(Of String, Object)
            Dim headers = BuildHeaders(extras)
            Dim d As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            For Each key In headers
                If row.ContainsKey(key) Then d(key) = row(key)
            Next
            Return d
        End Function

        Private Shared Function CloneRow(row As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
            Dim d As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            If row Is Nothing Then Return d
            For Each kv In row
                d(kv.Key) = kv.Value
            Next
            Return d
        End Function

        Private Shared Function RowPriority(row As Dictionary(Of String, Object)) As Integer
            Dim status = SafeCellString(row, "Status")
            Dim conn = SafeCellString(row, "ConnectionType")

            If String.Equals(status, "ERROR", StringComparison.OrdinalIgnoreCase) Then Return 5
            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then Return 4
            If String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then Return 4
            If conn.IndexOf("Proximity", StringComparison.OrdinalIgnoreCase) >= 0 OrElse String.Equals(conn, "Near", StringComparison.OrdinalIgnoreCase) Then Return 3
            If String.Equals(status, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) Then Return 3
            If String.Equals(status, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then Return 3
            If String.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) Then Return 2
            Return 1
        End Function

        Private Shared Function SelectBestRow(map As Dictionary(Of String, List(Of Dictionary(Of String, Object))), key As String) As Dictionary(Of String, Object)
            If Not map.ContainsKey(key) Then Return Nothing
            Dim best As Dictionary(Of String, Object) = Nothing

            For Each row In map(key)
                If row Is Nothing Then Continue For
                If best Is Nothing Then
                    best = row
                    Continue For
                End If

                Dim pr = RowPriority(row)
                Dim pb = RowPriority(best)
                If pr > pb Then
                    best = row
                ElseIf pr = pb Then
                    Dim dr = ToDoubleLocal(SafeCellString(row, "Distance (inch)"))
                    Dim db = ToDoubleLocal(SafeCellString(best, "Distance (inch)"))
                    If dr < db Then best = row
                End If
            Next

            Return best
        End Function

        Private Shared Function SwapRow(row As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
            If row Is Nothing Then Return Nothing
            Dim swapped As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            swapped("Id1") = SafeCellString(row, "Id2")
            swapped("Id2") = SafeCellString(row, "Id1")
            swapped("Category1") = SafeCellString(row, "Category2")
            swapped("Category2") = SafeCellString(row, "Category1")
            swapped("Family1") = SafeCellString(row, "Family2")
            swapped("Family2") = SafeCellString(row, "Family1")
            swapped("Distance (inch)") = SafeCellString(row, "Distance (inch)")
            swapped("ConnectionType") = SafeCellString(row, "ConnectionType")
            swapped("ParamName") = SafeCellString(row, "ParamName")
            swapped("Value1") = SafeCellString(row, "Value2")
            swapped("Value2") = SafeCellString(row, "Value1")
            swapped("ParamCompare") = SafeCellString(row, "ParamCompare")
            swapped("Status") = SafeCellString(row, "Status")
            swapped("ErrorMessage") = SafeCellString(row, "ErrorMessage")

            For Each kv In row
                If kv.Key Is Nothing Then Continue For
                If kv.Key.EndsWith("(ID1)", StringComparison.OrdinalIgnoreCase) Then
                    Dim name = kv.Key.Substring(0, kv.Key.Length - "(ID1)".Length)
                    swapped($"{name}(ID1)") = SafeCellString(row, $"{name}(ID2)")
                ElseIf kv.Key.EndsWith("(ID2)", StringComparison.OrdinalIgnoreCase) Then
                    Dim name = kv.Key.Substring(0, kv.Key.Length - "(ID2)".Length)
                    swapped($"{name}(ID2)") = SafeCellString(row, $"{name}(ID1)")
                End If
            Next

            Return swapped
        End Function

        Private Shared Function AppendExtrasForId1(row As Dictionary(Of String, Object), extras As IList(Of String)) As Dictionary(Of String, Object)
            Dim d = StripExtras(row, extras)
            If extras IsNot Nothing Then
                For Each name In extras
                    If row.ContainsKey(name) Then d(name) = SafeCellString(row, name)
                    Dim key1 = $"{name}(ID1)"
                    Dim key2 = $"{name}(ID2)"
                    d(key1) = SafeCellString(row, key1)
                    d(key2) = SafeCellString(row, key2)
                    If Not d.ContainsKey(name) Then d(name) = SafeCellString(row, key1)
                Next
            End If
            Return d
        End Function

        Private Shared Function BuildTotalRows(rows As List(Of Dictionary(Of String, Object))) As List(Of Dictionary(Of String, Object))
            If rows Is Nothing Then Return New List(Of Dictionary(Of String, Object))()
            Dim pairRows As New Dictionary(Of String, Dictionary(Of String, Object))(StringComparer.Ordinal)

            For Each row In rows
                If row Is Nothing Then Continue For
                Dim id1 = ToIntLocal(ReadField(row, "Id1"))
                Dim id2 = ToIntLocal(ReadField(row, "Id2"))
                Dim key = If(id1 <= id2, $"{id1}_{id2}", $"{id2}_{id1}")

                If Not pairRows.ContainsKey(key) Then
                    pairRows(key) = row
                Else
                    Dim cur = row
                    Dim best = pairRows(key)
                    Dim pr = RowPriority(cur)
                    Dim pb = RowPriority(best)
                    If pr > pb Then
                        pairRows(key) = cur
                    ElseIf pr = pb Then
                        Dim dr = ToDoubleLocal(SafeCellString(cur, "Distance (inch)"))
                        Dim db = ToDoubleLocal(SafeCellString(best, "Distance (inch)"))
                        If dr < db Then pairRows(key) = cur
                    End If
                End If
            Next

            Return pairRows.Values.Select(Function(r) CloneRow(r)).ToList()
        End Function

        Private Shared Sub WriteSheet(wb As XSSFWorkbook,
                                      sheetName As String,
                                      headers As List(Of String),
                                      rows As List(Of Dictionary(Of String, Object)),
                                      headerStyle As ICellStyle,
                                      baseStyle As ICellStyle,
                                      matchStyle As ICellStyle,
                                      mismatchStyle As ICellStyle,
                                      nearStyle As ICellStyle,
                                      errorStyle As ICellStyle,
                                      Optional progressChannel As String = Nothing,
                                      Optional ByRef written As Integer = 0,
                                      Optional totalRows As Integer = 0,
                                      Optional doAutoFit As Boolean = False)
            Dim sh = wb.CreateSheet(sheetName)

            Dim headerRow = sh.CreateRow(0)
            For i = 0 To headers.Count - 1
                Dim c = headerRow.CreateCell(i)
                c.SetCellValue(headers(i))
                c.CellStyle = headerStyle
            Next

            sh.CreateFreezePane(0, 1)

            If headers.Count > 0 Then
                Dim range As New NPOI.SS.Util.CellRangeAddress(0, 0, 0, headers.Count - 1)
                sh.SetAutoFilter(range)
            End If

            If rows IsNot Nothing Then
                Dim r As Integer = 1
                For Each row In rows
                    Dim sr = sh.CreateRow(r) : r += 1

                    Dim statusVal As String = SafeCellString(row, "Status")
                    Dim connVal As String = SafeCellString(row, "ConnectionType")
                    Dim styleToUse As ICellStyle = baseStyle

                    If String.Equals(statusVal, "ERROR", StringComparison.OrdinalIgnoreCase) Then
                        styleToUse = errorStyle
                    ElseIf IsMismatchStatus(statusVal) Then
                        styleToUse = mismatchStyle
                    ElseIf String.Equals(statusVal, "OK", StringComparison.OrdinalIgnoreCase) Then
                        styleToUse = matchStyle
                    ElseIf String.Equals(statusVal, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(statusVal, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then
                        styleToUse = nearStyle
                    ElseIf String.Equals(connVal.Trim(), "Near", StringComparison.OrdinalIgnoreCase) OrElse connVal.IndexOf("Proximity", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        styleToUse = nearStyle
                    End If

                    For c = 0 To headers.Count - 1
                        Dim key = headers(c)
                        Dim v As Object = Nothing
                        If row.ContainsKey(key) Then v = row(key)
                        Dim cell = sr.CreateCell(c)

                        Dim text As String = If(v Is Nothing, "", v.ToString())
                        If String.Equals(key, "Id2", StringComparison.OrdinalIgnoreCase) Then
                            Dim t = text.Trim()
                            If t = "" OrElse t = "0" Then
                                text = ""
                            ElseIf Not t.StartsWith(",", StringComparison.Ordinal) Then
                                text = "," & t
                            Else
                                text = t
                            End If
                        End If

                        cell.SetCellValue(text)
                        cell.CellStyle = styleToUse
                    Next
                    written += 1
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_WRITE", "엑셀 데이터 작성", written, totalRows)
                Next
            End If

            If doAutoFit Then
                ApplyFastColumnWidths(sh, headers, rows)
            End If
        End Sub

        Private Shared Sub ApplyFastColumnWidths(sh As ISheet, headers As List(Of String), rows As List(Of Dictionary(Of String, Object)))
            If sh Is Nothing OrElse headers Is Nothing Then Return

            Const MAX_SAMPLE As Integer = 2000
            Const MIN_CHARS As Integer = 6
            Const MAX_CHARS As Integer = 60
            Const MAX_WIDTH As Integer = 255 * 256

            Dim maxLens(headers.Count - 1) As Integer
            For i = 0 To headers.Count - 1
                maxLens(i) = If(headers(i) Is Nothing, 0, headers(i).Length)
            Next

            Dim sampleCount As Integer = 0
            If rows IsNot Nothing Then sampleCount = Math.Min(rows.Count, MAX_SAMPLE)

            For i = 0 To sampleCount - 1
                Dim row = rows(i)
                If row Is Nothing Then Continue For

                For c = 0 To headers.Count - 1
                    Dim key = headers(c)
                    Dim v As Object = Nothing
                    If row.ContainsKey(key) Then v = row(key)
                    Dim text As String = If(v Is Nothing, String.Empty, v.ToString())
                    If text.Length > maxLens(c) Then maxLens(c) = text.Length
                Next
            Next

            For c = 0 To headers.Count - 1
                Dim chars As Integer = maxLens(c) + 2
                If chars < MIN_CHARS Then chars = MIN_CHARS
                If chars > MAX_CHARS Then chars = MAX_CHARS

                Dim width As Integer = chars * 256
                If width > MAX_WIDTH Then width = MAX_WIDTH

                Try
                    sh.SetColumnWidth(c, width)
                Catch
                End Try
            Next
        End Sub

        Private Shared Function StatusRank(status As String) As Integer
            If String.IsNullOrEmpty(status) Then Return 0
            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then Return 3
            If String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then Return 3
            If String.Equals(status, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then Return 2
            If String.Equals(status, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) Then Return 2
            If String.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) Then Return 1
            Return 0
        End Function

        Private Shared Function ToDoubleLocal(val As String) As Double
            Try
                If String.IsNullOrWhiteSpace(val) Then Return Double.MaxValue
                Return Convert.ToDouble(val)
            Catch
                Return Double.MaxValue
            End Try
        End Function

        Private Shared Function ToIntLocal(val As String) As Integer
            Try
                If String.IsNullOrEmpty(val) Then Return 0
                Dim s = val.Trim()
                Return Convert.ToInt32(s)
            Catch
                Return 0
            End Try
        End Function

#End Region

    End Class
End Namespace
