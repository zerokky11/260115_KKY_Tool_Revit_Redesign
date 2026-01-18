Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.DB.Plumbing
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.UI.Hub
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports NPOI.XSSF.UserModel
Imports RvtDB = Autodesk.Revit.DB
Imports NpoiCellType = NPOI.SS.UserModel.CellType

Namespace Services

    Public Class SegmentPmsCheckService

        Public Class PmsRow
            Public Property [Class] As String = String.Empty
            Public Property SegmentKey As String = String.Empty
            Public Property NdMm As Double
            Public Property IdMm As Double
            Public Property OdMm As Double
        End Class

        Public Class MappingSelection
            Public Property [File] As String = String.Empty
            Public Property PipeTypeName As String = String.Empty
            Public Property RuleIndex As Integer
            Public Property SegmentId As Integer
            Public Property SegmentKey As String = String.Empty
            Public Property SelectedClass As String = String.Empty
            Public Property SelectedPmsSegment As String = String.Empty
            Public Property MappingSource As String = String.Empty
        End Class

        Public Class GroupSelection
            Public Property GroupKey As String = String.Empty
            Public Property SelectedClass As String = String.Empty
            Public Property SelectedPmsSegment As String = String.Empty
            Public Property SelectionSource As String = String.Empty
        End Class

        Public Class ExtractOptions
            Public Property NdRound As Integer = 3
            Public Property DetachFromCentral As Boolean = True
            Public Property OpenReadOnly As Boolean = True
            Public Property ToleranceMm As Double = 0.01R
        End Class

        Public Class CompareOptions
            Public Property NdRound As Integer = 3
            Public Property TolMm As Double = 0.01R
            Public Property ClassMatch As Boolean = False
        End Class

        Public Class SuggestedMapping
            Public Property [File] As String = String.Empty
            Public Property PipeTypeName As String = String.Empty
            Public Property RuleIndex As Integer
            Public Property SegmentId As Integer
            Public Property SegmentKey As String = String.Empty
            Public Property PmsClass As String = String.Empty
            Public Property PmsSegmentKey As String = String.Empty
            Public Property Score As Double
        End Class

        Public Class MappingGroup
            Public Property GroupKey As String = String.Empty
            Public Property DisplayKey As String = String.Empty
            Public Property NormalizedKey As String = String.Empty
            Public Property Usages As List(Of MappingUsage)
            Public Property SuggestedClass As String = String.Empty
            Public Property SuggestedSegmentKey As String = String.Empty
            Public Property FileCount As Integer
            Public Property PipeTypeCount As Integer
            Public Property UsageSummary As String = String.Empty
        End Class

        Public Class MappingUsage
            Public Property [File] As String = String.Empty
            Public Property PipeTypeName As String = String.Empty
            Public Property RuleIndex As Integer
            Public Property SegmentId As Integer
            Public Property SegmentKey As String = String.Empty
        End Class

        Public Class RunResult
            Public Property MapTable As DataTable
            Public Property RevitSizeTable As DataTable
            Public Property PmsSizeTable As DataTable
            Public Property CompareTable As DataTable
            Public Property ErrorTable As DataTable
            Public Property SummaryTable As DataTable
        End Class

        Public Class LoadPmsResult
            Public Property Table As DataTable
            Public Property Rows As List(Of PmsRow)
            Public Property Errors As List(Of String)
        End Class

        Public Const TableMeta As String = "Extract_Meta"
        Public Const TableFiles As String = "Extract_Files"
        Public Const TableRules As String = "Extract_Rules"
        Public Const TableSizes As String = "Extract_Sizes"
        Public Const TableRouting As String = "Extract_Routing"
        Private Const FeetToMm As Double = 304.8R
        Private Const ToolVersion As String = "SegmentPms 2.0"

        ' ---------------------------
        ' Extract stage
        ' ---------------------------
        Public Shared Function ExtractToDataSet(app As UIApplication,
                                                files As IEnumerable(Of String),
                                                options As ExtractOptions,
                                                Optional progress As Action(Of Integer, Integer, String, String, String) = Nothing) As DataSet
            Dim ds As New DataSet()
            Dim meta = BuildMetaTable()
            Dim fileTable = BuildFileTable()
            Dim rules = BuildRuleTable()
            Dim sizes = BuildSizeTable()
            Dim routing = BuildRoutingTable()
            ds.Tables.Add(meta)
            ds.Tables.Add(fileTable)
            ds.Tables.Add(rules)
            ds.Tables.Add(sizes)
            ds.Tables.Add(routing)

            Dim metaRow = meta.NewRow()
            metaRow("NdRound") = options.NdRound
            metaRow("Tolerance") = options.ToleranceMm
            metaRow("CreatedAt") = DateTime.Now.ToString("s", CultureInfo.InvariantCulture)
            metaRow("ToolVersion") = ToolVersion
            meta.Rows.Add(metaRow)

            Dim valid As New List(Of String)()
            If files IsNot Nothing Then
                For Each f As String In files
                    If Not String.IsNullOrWhiteSpace(f) AndAlso File.Exists(f) Then
                        Dim already As Boolean = False
                        For Each v As String In valid
                            If v.Equals(f, StringComparison.OrdinalIgnoreCase) Then
                                already = True
                                Exit For
                            End If
                        Next
                        If Not already Then
                            valid.Add(f)
                        End If
                    End If
                Next
            End If

            If valid.Count = 0 Then
                Return ds
            End If

            Dim totalCount As Integer = valid.Count
            If progress IsNot Nothing Then
                progress(totalCount, 0, "open", "추출 시작", String.Empty)
            End If

            Dim appObj = app.Application
            For i As Integer = 0 To valid.Count - 1
                Dim p As String = valid(i)
                Dim fileIndex As Integer = i + 1
                Dim doc As RvtDB.Document = Nothing
                Try
                    If progress IsNot Nothing Then
                        progress(totalCount, fileIndex, "open", "파일 여는 중", p)
                    End If
                    Dim opt = BuildOpenOptions(options, p)
                    Dim mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(p)
                    doc = appObj.OpenDocumentFile(mp, opt)

                    If progress IsNot Nothing Then
                        progress(totalCount, fileIndex, "extract", "Segment 후보 수집 중", p)
                    End If
                    Dim fileRow = fileTable.NewRow()
                    fileRow("File") = p
                    fileRow("FileName") = Path.GetFileName(p)
                    fileRow("ExtractedAt") = DateTime.Now.ToString("s", CultureInfo.InvariantCulture)
                    fileTable.Rows.Add(fileRow)

                    If progress IsNot Nothing Then
                        progress(totalCount, fileIndex, "route", "RoutingPreference 수집 중", p)
                    End If
                    Dim routingInfos = CollectRouting(doc, p)
                    For Each info In routingInfos
                        Dim row = routing.NewRow()
                        row("File") = info.File
                        row("PipeTypeName") = info.PipeTypeName
                        row("RuleGroup") = info.RuleGroup
                        row("RuleIndex") = info.RuleIndex
                        row("RuleType") = info.RuleType
                        row("PartId") = info.PartId
                        row("PartName") = info.PartName
                        row("TypeName") = info.TypeName
                        routing.Rows.Add(row)
                    Next

                    Dim pipeInfos = CollectPipeTypeSegmentCandidates(doc, p)
                    For Each pi In pipeInfos
                        For Each cand In pi.Candidates
                            Dim rr = rules.NewRow()
                            rr("File") = p
                            rr("PipeTypeName") = pi.PipeTypeName
                            rr("RuleIndex") = cand.RuleIndex
                            rr("SegmentId") = cand.SegmentId
                            rr("SegmentKey") = cand.SegmentKey
                            rules.Rows.Add(rr)
                        Next
                    Next

                    Dim segIds As New HashSet(Of Integer)()
                    For Each pi In pipeInfos
                        For Each c In pi.Candidates
                            segIds.Add(c.SegmentId)
                        Next
                    Next

                    Dim sizeRows = CollectSegmentSizes(doc, segIds, p, options.NdRound)
                    For Each s In sizeRows
                        Dim sr = sizes.NewRow()
                        sr("File") = s.File
                        sr("SegmentId") = s.SegmentId
                        sr("SegmentKey") = s.SegmentKey
                        sr("ND_mm") = s.NdMm
                        sr("ID_mm") = s.IdMm
                        sr("OD_mm") = s.OdMm
                        sizes.Rows.Add(sr)
                    Next

                    If progress IsNot Nothing Then
                        progress(totalCount, fileIndex, "done", "파일 처리 완료", p)
                    End If
                Catch
                    ' 개별 파일 오류는 누적하지 않고 건너뜀
                    If progress IsNot Nothing Then
                        progress(totalCount, fileIndex, "error", "파일 처리 중 오류", p)
                    End If
                Finally
                    If doc IsNot Nothing Then
                        Try
                            doc.Close(False)
                        Catch
                        End Try
                    End If
                End Try
            Next

            Return ds
        End Function

        Public Shared Function ExtractFromDocument(app As UIApplication,
                                                   doc As RvtDB.Document,
                                                   filePath As String,
                                                   options As ExtractOptions,
                                                   Optional progress As Action(Of Integer, Integer, String, String, String) = Nothing) As DataSet
            Dim ds As New DataSet()
            Dim meta = BuildMetaTable()
            Dim fileTable = BuildFileTable()
            Dim rules = BuildRuleTable()
            Dim sizes = BuildSizeTable()
            Dim routing = BuildRoutingTable()
            ds.Tables.Add(meta)
            ds.Tables.Add(fileTable)
            ds.Tables.Add(rules)
            ds.Tables.Add(sizes)
            ds.Tables.Add(routing)

            Dim metaRow = meta.NewRow()
            metaRow("NdRound") = options.NdRound
            metaRow("Tolerance") = options.ToleranceMm
            metaRow("CreatedAt") = DateTime.Now.ToString("s", CultureInfo.InvariantCulture)
            metaRow("ToolVersion") = ToolVersion
            meta.Rows.Add(metaRow)

            If doc Is Nothing Then
                Return ds
            End If

            Dim totalCount As Integer = 1
            Dim fileIndex As Integer = 1
            Dim safePath As String = If(filePath, String.Empty)
            Dim fileName As String = Path.GetFileName(safePath)

            If progress IsNot Nothing Then
                progress(totalCount, fileIndex, "open", "파일 여는 중", safePath)
            End If

            Dim fileRow = fileTable.NewRow()
            fileRow("File") = safePath
            fileRow("FileName") = fileName
            fileRow("ExtractedAt") = DateTime.Now.ToString("s", CultureInfo.InvariantCulture)
            fileTable.Rows.Add(fileRow)

            If progress IsNot Nothing Then
                progress(totalCount, fileIndex, "extract", "Segment 후보 수집 중", safePath)
            End If
            Dim pipeInfos = CollectPipeTypeSegmentCandidates(doc, safePath)
            For Each info In pipeInfos
                Dim row = rules.NewRow()
                row("File") = info.File
                row("FileName") = info.FileName
                row("PipeTypeName") = info.PipeTypeName
                row("RuleIndex") = info.RuleIndex
                row("RevitSegmentKey") = info.RevitSegmentKey
                row("SegmentId") = info.SegmentId
                row("SegmentClass") = info.SegmentClass
                row("SegmentName") = info.SegmentName
                row("SegmentPipeType") = info.SegmentPipeType
                row("SegmentIdParam") = info.SegmentIdParam
                rules.Rows.Add(row)

                For Each sz In info.Sizes
                    Dim sr = sizes.NewRow()
                    sr("File") = info.File
                    sr("FileName") = info.FileName
                    sr("PipeTypeName") = info.PipeTypeName
                    sr("RuleIndex") = info.RuleIndex
                    sr("SegmentId") = info.SegmentId
                    sr("SegmentKey") = info.RevitSegmentKey
                    sr("ND_mm") = sz.NdMm
                    sr("Revit_ID") = sz.IdMm
                    sr("Revit_OD") = sz.OdMm
                    sizes.Rows.Add(sr)
                Next
            Next

            If progress IsNot Nothing Then
                progress(totalCount, fileIndex, "route", "RoutingPreference 수집 중", safePath)
            End If
            Dim routingInfos = CollectRouting(doc, safePath)
            For Each info In routingInfos
                Dim row = routing.NewRow()
                row("File") = info.File
                row("FileName") = info.FileName
                row("PipeTypeName") = info.PipeTypeName
                row("PartName") = info.PartName
                row("TypeName") = info.TypeName
                routing.Rows.Add(row)
            Next

            If progress IsNot Nothing Then
                progress(totalCount, fileIndex, "done", "파일 처리 완료", safePath)
            End If

            Return ds
        End Function

        Public Shared Sub SaveDataSetToXlsx(ds As DataSet, path As String, Optional doAutoFit As Boolean = False, Optional progressChannel As String = Nothing)
            If ds Is Nothing Then Return
            Dim totalRows As Integer = 0
            For Each t As DataTable In ds.Tables
                If t IsNot Nothing Then
                    totalRows += t.Rows.Count
                End If
            Next
            Dim written As Integer = 0
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset(progressChannel)
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_INIT", "엑셀 워크북 준비", 0, totalRows, Nothing, True)
            UI.Hub.UiBridgeExternalEvent.LogAutoFitDecision(doAutoFit, "SegmentPmsCheckService.SaveDataSetToXlsx")
            Try
                Dim wb As IWorkbook = New XSSFWorkbook()
                If ds.Tables.Contains(TableMeta) Then WriteSheet(wb, TableMeta, ds.Tables(TableMeta), progressChannel, written, totalRows)
                If ds.Tables.Contains(TableFiles) Then WriteSheet(wb, TableFiles, ds.Tables(TableFiles), progressChannel, written, totalRows)
                If ds.Tables.Contains(TableRules) Then WriteSheet(wb, TableRules, ds.Tables(TableRules), progressChannel, written, totalRows)
                If ds.Tables.Contains(TableSizes) Then WriteSheet(wb, TableSizes, ds.Tables(TableSizes), progressChannel, written, totalRows)
                If ds.Tables.Contains(TableRouting) Then WriteSheet(wb, TableRouting, ds.Tables(TableRouting), progressChannel, written, totalRows)
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_SAVE", "파일 저장 중", written, totalRows, Nothing, True)
                SaveWorkbookSafely(wb, path)
                wb.Close()
                Dim autoFitMessage As String = If(doAutoFit, "AutoFit 적용", "빠른 모드: AutoFit 생략")
                If doAutoFit Then
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", autoFitMessage, written, totalRows, Nothing, True)
                    Global.KKY_Tool_Revit.Infrastructure.ExcelCore.TryAutoFitWithExcel(path)
                Else
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", autoFitMessage, written, totalRows, Nothing, True)
                End If
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "DONE", "엑셀 내보내기 완료", written, totalRows, 100.0R, True)
            Catch ex As Exception
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "ERROR", ex.Message, written, totalRows, Nothing, True)
                Throw
            End Try
        End Sub

        Public Shared Function LoadExtractFromXlsx(path As String) As DataSet
            Dim ds As New DataSet()
            If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then
                Return ds
            End If

            Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim wb As IWorkbook = New XSSFWorkbook(fs)
                For i As Integer = 0 To wb.NumberOfSheets - 1
                    Dim sh As ISheet = wb.GetSheetAt(i)
                    If sh Is Nothing Then
                        Continue For
                    End If
                    Dim t As New DataTable(sh.SheetName)
                    Dim head As IRow = sh.GetRow(0)
                    If head Is Nothing Then
                        Continue For
                    End If
                    For ci As Integer = 0 To head.LastCellNum - 1
                        t.Columns.Add(CellStr(head, ci), GetType(String))
                    Next
                    For r As Integer = 1 To sh.LastRowNum
                        Dim row As IRow = sh.GetRow(r)
                        If row Is Nothing Then
                            Continue For
                        End If
                        Dim dr = t.NewRow()
                        For ci As Integer = 0 To t.Columns.Count - 1
                            dr(ci) = CellStr(row, ci)
                        Next
                        t.Rows.Add(dr)
                    Next
                    ds.Tables.Add(t)
                Next
            End Using

            EnsureSchema(ds)
            Return ds
        End Function

        ' ---------------------------
        ' PMS
        ' ---------------------------
        Public Shared Function LoadPmsExcel(xlsxPath As String, preferredUnit As String) As LoadPmsResult
            Dim res As New LoadPmsResult With {.Table = BuildPmsTableSkeleton(), .Rows = New List(Of PmsRow)(), .Errors = New List(Of String)()}
            If String.IsNullOrWhiteSpace(xlsxPath) OrElse Not File.Exists(xlsxPath) Then
                res.Errors.Add("PMS 파일이 존재하지 않습니다.")
                Return res
            End If

            Dim wb As IWorkbook = Nothing
            Using fs As New FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                wb = New XSSFWorkbook(fs)
            End Using
            If wb.NumberOfSheets <= 0 Then
                res.Errors.Add("PMS 파일에 시트가 없습니다.")
                Return res
            End If

            Dim sh As ISheet = wb.GetSheetAt(0)
            If sh Is Nothing Then
                res.Errors.Add("PMS 시트를 읽을 수 없습니다.")
                Return res
            End If

            Dim headerMap = DetectHeader(sh)
            If Not headerMap.ContainsKey("class") Then res.Errors.Add("CLASS 헤더를 찾을 수 없습니다.")
            If Not headerMap.ContainsKey("segment") Then res.Errors.Add("Segment 헤더를 찾을 수 없습니다.")
            If Not headerMap.ContainsKey("nd") Then res.Errors.Add("ND 헤더를 찾을 수 없습니다.")
            If Not headerMap.ContainsKey("id") Then res.Errors.Add("ID 헤더를 찾을 수 없습니다.")
            If Not headerMap.ContainsKey("od") Then res.Errors.Add("OD 헤더를 찾을 수 없습니다.")
            If res.Errors.Count > 0 Then
                Return res
            End If

            Dim unitLabel As String = If(String.IsNullOrWhiteSpace(preferredUnit), "mm", preferredUnit).ToLowerInvariant()

            Dim lastRow As Integer = sh.LastRowNum
            For i As Integer = 1 To lastRow
                Dim row As IRow = sh.GetRow(i)
                If row Is Nothing Then
                    Continue For
                End If
                Dim cls As String = CellStr(row, headerMap("class"))
                Dim seg As String = CellStr(row, headerMap("segment"))
                Dim nd As Double = CellDbl(row, headerMap("nd"), 0)
                Dim id As Double = CellDbl(row, headerMap("id"), 0)
                Dim od As Double = CellDbl(row, headerMap("od"), 0)

                If String.IsNullOrWhiteSpace(seg) Then
                    Continue For
                End If

                Dim ndMm As Double = nd
                Dim idMm As Double = id
                Dim odMm As Double = od

                If unitLabel.IndexOf("in", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    ndMm = nd * 25.4R
                    idMm = id * 25.4R
                    odMm = od * 25.4R
                End If

                Dim dataRow = res.Table.NewRow()
                dataRow("CLASS") = cls
                dataRow("PMS_SegmentKey") = seg
                dataRow("ND_mm") = ndMm
                dataRow("ID_mm") = idMm
                dataRow("OD_mm") = odMm
                res.Table.Rows.Add(dataRow)

                res.Rows.Add(New PmsRow With {
                    .Class = cls,
                    .SegmentKey = seg,
                    .NdMm = ndMm,
                    .IdMm = idMm,
                    .OdMm = odMm
                })
            Next

            Return res
        End Function

        Public Shared Sub ExportPmsTemplateXlsx(outPath As String, preferredUnit As String)
            If String.IsNullOrWhiteSpace(outPath) Then
                Throw New ArgumentException("저장 경로가 비어 있습니다.", NameOf(outPath))
            End If

            Dim unitLabel As String = If(String.IsNullOrWhiteSpace(preferredUnit), "mm", preferredUnit).ToLowerInvariant()
            Dim useInch As Boolean = unitLabel.IndexOf("in", StringComparison.OrdinalIgnoreCase) >= 0
            Dim ndHeader As String = If(useInch, "ND_in", "ND_mm")
            Dim idHeader As String = If(useInch, "ID_in", "ID_mm")
            Dim odHeader As String = If(useInch, "OD_in", "OD_mm")
            Dim headers As New List(Of String) From {"CLASS", "Segment", ndHeader, idHeader, odHeader}

            Using wb As New XSSFWorkbook()
                Dim headerStyle As ICellStyle = wb.CreateCellStyle()
                Dim headerFont As IFont = wb.CreateFont()
                headerFont.IsBold = True
                headerStyle.SetFont(headerFont)

                Dim sheet = wb.CreateSheet("PMS_Template")
                Dim headerRow = sheet.CreateRow(0)
                For i As Integer = 0 To headers.Count - 1
                    Dim cell = headerRow.CreateCell(i)
                    cell.SetCellValue(headers(i))
                    cell.CellStyle = headerStyle
                Next

                Dim samples As New List(Of Object()) From {
                    New Object() {"HVAC", "SEG-100", If(useInch, 4.0R, 100.0R), If(useInch, 3.75R, 95.0R), If(useInch, 4.5R, 108.0R)},
                    New Object() {"HVAC", "SEG-150", If(useInch, 6.0R, 150.0R), If(useInch, 5.5R, 140.0R), If(useInch, 6.5R, 165.0R)},
                    New Object() {"Plumbing", "SEG-200", If(useInch, 8.0R, 200.0R), If(useInch, 7.5R, 190.0R), If(useInch, 8.5R, 214.0R)}
                }

                For r As Integer = 0 To samples.Count - 1
                    Dim row = sheet.CreateRow(r + 1)
                    Dim values = samples(r)
                    For c As Integer = 0 To values.Length - 1
                        Dim cell = row.CreateCell(c)
                        Dim val = values(c)
                        If TypeOf val Is Double OrElse TypeOf val Is Single OrElse TypeOf val Is Decimal Then
                            cell.SetCellValue(Convert.ToDouble(val))
                        Else
                            cell.SetCellValue(If(val Is Nothing, String.Empty, val.ToString()))
                        End If
                    Next
                Next

                Dim readme = wb.CreateSheet("README")
                Dim notes As New List(Of String) From {
                    "PMS 등록용 샘플 양식입니다.",
                    "필수 헤더: CLASS, Segment, " & ndHeader & ", " & idHeader & ", " & odHeader,
                    "단위: " & If(useInch, "inch", "mm"),
                    "Segment 컬럼은 PMS Segment Key 값입니다.",
                    "값이 비어있는 행은 자동으로 무시됩니다.",
                    "헤더명이 변경되면 PMS 등록에 실패할 수 있습니다."
                }
                For i As Integer = 0 To notes.Count - 1
                    Dim row = readme.CreateRow(i)
                    row.CreateCell(0).SetCellValue(notes(i))
                Next

                Using fs As New FileStream(outPath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fs)
                End Using
            End Using
        End Sub

        ' ---------------------------
        ' Suggestion / Compare
        ' ---------------------------
        Public Shared Function SuggestMappings(extractData As DataSet, pmsData As List(Of PmsRow)) As List(Of SuggestedMapping)
            Dim result As New List(Of SuggestedMapping)()
            If extractData Is Nothing OrElse pmsData Is Nothing Then
                Return result
            End If
            If Not extractData.Tables.Contains(TableRules) Then
                Return result
            End If
            Dim rules = extractData.Tables(TableRules)

            For Each r As DataRow In rules.Rows
                Dim filePath = SafeStr(r("File"))
                Dim pipeName = SafeStr(r("PipeTypeName"))
                Dim ruleIdx = SafeIntObj(r("RuleIndex"))
                Dim segId = SafeIntObj(r("SegmentId"))
                Dim segKey = SafeStr(r("SegmentKey"))
                Dim best = FindBestSuggestion(segKey, pmsData)

                result.Add(New SuggestedMapping With {
                    .File = filePath,
                    .PipeTypeName = pipeName,
                    .RuleIndex = ruleIdx,
                    .SegmentId = segId,
                    .SegmentKey = segKey,
                    .PmsClass = best.BestClass,
                    .PmsSegmentKey = best.BestSegment,
                    .Score = best.Score
                })
            Next

            Return result
        End Function

        Public Shared Function BuildGroups(extractData As DataSet) As List(Of MappingGroup)
            Dim groups As New Dictionary(Of String, MappingGroup)(StringComparer.OrdinalIgnoreCase)
            If extractData Is Nothing OrElse Not extractData.Tables.Contains(TableRules) Then
                Return New List(Of MappingGroup)()
            End If
            Dim rules = extractData.Tables(TableRules)
            For Each r As DataRow In rules.Rows
                Dim segKey = SafeStr(r("SegmentKey"))
                Dim norm = NormalizeSegmentGroupKey(segKey)
                Dim groupKey = If(String.IsNullOrWhiteSpace(norm), segKey, norm)
                If Not groups.ContainsKey(groupKey) Then
                    groups(groupKey) = New MappingGroup With {
                        .GroupKey = groupKey,
                        .DisplayKey = segKey,
                        .NormalizedKey = norm,
                        .Usages = New List(Of MappingUsage)(),
                        .FileCount = 0,
                        .PipeTypeCount = 0,
                        .UsageSummary = String.Empty
                    }
                End If
                Dim g = groups(groupKey)
                If String.IsNullOrWhiteSpace(g.DisplayKey) Then
                    g.DisplayKey = segKey
                End If
                g.Usages.Add(New MappingUsage With {
                    .File = SafeStr(r("File")),
                    .PipeTypeName = SafeStr(r("PipeTypeName")),
                    .RuleIndex = SafeIntObj(r("RuleIndex")),
                    .SegmentId = SafeIntObj(r("SegmentId")),
                    .SegmentKey = segKey
                })
            Next

            For Each kv In groups
                Dim g = kv.Value
                Dim fileSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim pipeSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each u In g.Usages
                    fileSet.Add(NormalizePath(u.File))
                    pipeSet.Add(u.File & "|" & u.PipeTypeName)
                Next
                g.FileCount = fileSet.Count
                g.PipeTypeCount = pipeSet.Count
                Dim sb As New StringBuilder()
                sb.Append("Used in: ")
                Dim fileInfo As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                For Each u In g.Usages
                    Dim key = NormalizePath(u.File)
                    If Not fileInfo.ContainsKey(key) Then
                        fileInfo(key) = 0
                    End If
                    fileInfo(key) += 1
                Next
                Dim first As Boolean = True
                For Each kvp In fileInfo
                    If Not first Then
                        sb.Append(", ")
                    End If
                    first = False
                    sb.Append(Path.GetFileName(kvp.Key))
                    sb.Append("("c)
                    sb.Append(kvp.Value.ToString(CultureInfo.InvariantCulture))
                    sb.Append(" PipeTypes)")
                Next
                g.UsageSummary = sb.ToString()
            Next

            Return New List(Of MappingGroup)(groups.Values)
        End Function

        Public Shared Function SuggestGroupMappings(groups As List(Of MappingGroup), pmsData As List(Of PmsRow)) As List(Of SuggestedMapping)
            Dim result As New List(Of SuggestedMapping)()
            If groups Is Nothing OrElse pmsData Is Nothing Then
                Return result
            End If
            For Each g In groups
                Dim best = FindBestSuggestion(g.DisplayKey, pmsData)
                If best.Score > 0 Then
                    g.SuggestedClass = best.BestClass
                    g.SuggestedSegmentKey = best.BestSegment
                    result.Add(New SuggestedMapping With {
                        .File = g.GroupKey,
                        .PipeTypeName = g.DisplayKey,
                        .RuleIndex = 0,
                        .SegmentId = 0,
                        .SegmentKey = g.GroupKey,
                        .PmsClass = best.BestClass,
                        .PmsSegmentKey = best.BestSegment,
                        .Score = best.Score
                    })
                End If
            Next
            Return result
        End Function

        Private Class SegmentTokenInfo
            Public Property Raw As String = String.Empty
            Public Property BaseCode As String = String.Empty
            Public Property VariantCode As String = String.Empty
            Public Property Tokens As HashSet(Of String)
            Public Property MaterialTokens As HashSet(Of String)
            Public Property Normalized As String = String.Empty
            Public Property IsGroupLike As Boolean
        End Class

        Private Class SuggestionResult
            Public Property BestClass As String = String.Empty
            Public Property BestSegment As String = String.Empty
            Public Property Score As Double
            Public Property DebugHint As String = String.Empty
        End Class

        Private Class CandidateScore
            Public Property BaseScore As Double
            Public Property VariantScore As Double
            Public Property TokenScore As Double
            Public Property Similarity As Double
            Public Property Score As Double
            Public Property LenDiff As Integer
            Public Property Detail As String = String.Empty
        End Class

        Private Shared Function FindBestSuggestion(segmentKey As String, pmsData As List(Of PmsRow)) As SuggestionResult
            Dim res As New SuggestionResult()
            If String.IsNullOrWhiteSpace(segmentKey) OrElse pmsData Is Nothing Then
                Return res
            End If

            Dim revitTokens = TokenizeRevitSegmentKey(segmentKey)
            If revitTokens.Tokens Is Nothing OrElse revitTokens.Tokens.Count = 0 Then
                Return res
            End If

            Dim bestScore As Double = Double.MinValue
            Dim bestLenDiff As Integer = Integer.MaxValue
            Dim bestDetail As String = String.Empty
            Dim cache As New Dictionary(Of String, SegmentTokenInfo)(StringComparer.OrdinalIgnoreCase)

            For Each p In pmsData
                Dim pTokens As SegmentTokenInfo = Nothing
                If Not cache.TryGetValue(p.SegmentKey, pTokens) Then
                    pTokens = TokenizePmsSegmentKey(p.SegmentKey)
                    cache(p.SegmentKey) = pTokens
                End If

                Dim info = ComputeCandidateScore(revitTokens, pTokens)
                If info.Score > bestScore OrElse (Math.Abs(info.Score - bestScore) < 0.0001R AndAlso info.LenDiff < bestLenDiff) Then
                    bestScore = info.Score
                    bestLenDiff = info.LenDiff
                    bestDetail = info.Detail
                    res.BestClass = p.Class
                    res.BestSegment = p.SegmentKey
                    res.Score = info.Score
                    res.DebugHint = bestDetail
                End If
            Next

            If res.Score <= 0 Then
                res.BestClass = String.Empty
                res.BestSegment = String.Empty
                res.DebugHint = String.Empty
            End If
            Return res
        End Function

        Private Shared Function TokenizeRevitSegmentKey(segmentKey As String) As SegmentTokenInfo
            Return TokenizeSegment(segmentKey, False)
        End Function

        Private Shared Function TokenizePmsSegmentKey(segmentKey As String) As SegmentTokenInfo
            Return TokenizeSegment(segmentKey, True)
        End Function

        Private Shared Function TokenizeSegment(segmentKey As String, isPms As Boolean) As SegmentTokenInfo
            Dim info As New SegmentTokenInfo With {
                .Raw = SafeStr(segmentKey),
                .Tokens = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                .MaterialTokens = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            }
            If String.IsNullOrWhiteSpace(info.Raw) Then
                Return info
            End If

            Dim upperRaw = info.Raw.ToUpperInvariant()
            Dim normalizedVariant = NormalizeVariantMarkers(upperRaw)
            Dim withoutRef = Regex.Replace(normalizedVariant, "\bREF\.?\b", " ", RegexOptions.IgnoreCase)
            Dim baseSource = withoutRef
            If isPms Then
                Dim pipeIdx As Integer = baseSource.LastIndexOf("|"c)
                If pipeIdx >= 0 AndAlso pipeIdx < baseSource.Length - 1 Then
                    baseSource = baseSource.Substring(pipeIdx + 1)
                End If
            End If

            info.BaseCode = ExtractBaseCode(baseSource)
            info.IsGroupLike = IsGroupLikeKey(baseSource)
            info.VariantCode = ExtractVariantToken(baseSource)
            info.MaterialTokens = ExtractMaterialTokens(upperRaw)
            info.Tokens = BuildTokens(withoutRef, info.MaterialTokens, info.BaseCode, info.VariantCode)
            info.Normalized = NormalizeForSimilarityTokens(info.Tokens)
            Return info
        End Function

        Private Shared Function NormalizeVariantMarkers(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then
                Return String.Empty
            End If
            Dim work = text
            work = Regex.Replace(work, "\(\s*A\.?P\.?\s*\)", " AP ", RegexOptions.IgnoreCase)
            work = Regex.Replace(work, "\(\s*M\.?P\.?\s*\)", " MP ", RegexOptions.IgnoreCase)
            work = Regex.Replace(work, "\(\s*E\.?P\.?\s*\)", " EP ", RegexOptions.IgnoreCase)
            work = Regex.Replace(work, "\bA\.?P\.?\b", " AP ", RegexOptions.IgnoreCase)
            work = Regex.Replace(work, "\bM\.?P\.?\b", " MP ", RegexOptions.IgnoreCase)
            work = Regex.Replace(work, "\bE\.?P\.?\b", " EP ", RegexOptions.IgnoreCase)
            Return work
        End Function

        Private Shared Function ExtractBaseCode(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then
                Return String.Empty
            End If
            Dim m = Regex.Match(text, "^\s*([A-Z]\d+[A-Z0-9]*)", RegexOptions.IgnoreCase)
            If m.Success AndAlso m.Groups.Count > 1 Then
                Return m.Groups(1).Value.ToUpperInvariant()
            End If
            Return String.Empty
        End Function

        Private Shared Function ExtractVariantToken(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then
                Return String.Empty
            End If
            Dim normalized = NormalizeVariantMarkers(text)
            Dim m = Regex.Match(normalized, "\b(AP|MP|EP)\b", RegexOptions.IgnoreCase)
            If m.Success AndAlso m.Groups.Count > 1 Then
                Return m.Groups(1).Value.ToUpperInvariant()
            End If
            Return String.Empty
        End Function

        Private Shared Function ExtractMaterialTokens(text As String) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(text) Then
                Return result
            End If
            For Each m As Match In Regex.Matches(text, "STS\s*\d+[A-Z]*", RegexOptions.IgnoreCase)
                Dim norm = NormalizeMaterialToken(m.Value)
                If Not String.IsNullOrWhiteSpace(norm) Then
                    result.Add(norm)
                End If
            Next
            Return result
        End Function

        Private Shared Function NormalizeMaterialToken(token As String) As String
            If String.IsNullOrWhiteSpace(token) Then
                Return String.Empty
            End If
            Dim normalized = Regex.Replace(token, "\s+", " ").Trim().ToUpperInvariant()
            Return normalized
        End Function

        Private Shared Function BuildTokens(text As String,
                                            materialTokens As HashSet(Of String),
                                            baseCode As String,
                                            variantCode As String) As HashSet(Of String)
            Dim tokens As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If Not String.IsNullOrWhiteSpace(text) Then
                Dim cleaned = Regex.Replace(text, "[\.,/\\\|\(\)\[\]:\-\+]", " ")
                cleaned = Regex.Replace(cleaned, "\s+", " ").Trim()
                For Each raw In cleaned.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim t = raw.Trim()
                    If t.Length > 0 Then
                        tokens.Add(t)
                    End If
                Next
            End If
            If Not String.IsNullOrWhiteSpace(baseCode) Then
                tokens.Add(baseCode)
            End If
            If Not String.IsNullOrWhiteSpace(variantCode) Then
                tokens.Add(variantCode)
            End If
            If materialTokens IsNot Nothing Then
                For Each m In materialTokens
                    tokens.Add(m)
                Next
            End If
            Return tokens
        End Function

        Private Shared Function NormalizeForSimilarityTokens(tokens As HashSet(Of String)) As String
            If tokens Is Nothing OrElse tokens.Count = 0 Then
                Return String.Empty
            End If
            Dim ordered = tokens.Where(Function(t) Not String.IsNullOrWhiteSpace(t)).Select(Function(t) t.ToUpperInvariant()).OrderBy(Function(t) t)
            Return String.Join(String.Empty, ordered)
        End Function

        Private Shared Function IsGroupLikeKey(text As String) As Boolean
            If String.IsNullOrWhiteSpace(text) Then
                Return False
            End If
            Return Regex.IsMatch(text, "[/\\+]", RegexOptions.IgnoreCase)
        End Function

        Private Shared Function ContainsBaseToken(pmsRaw As String, baseCode As String) As Boolean
            If String.IsNullOrWhiteSpace(pmsRaw) OrElse String.IsNullOrWhiteSpace(baseCode) Then
                Return False
            End If
            Dim compact = Regex.Replace(pmsRaw.ToUpperInvariant(), "[^A-Z0-9]", String.Empty)
            If compact.Contains(baseCode.ToUpperInvariant()) Then
                Return True
            End If
            For Each m As Match In Regex.Matches(pmsRaw, "[A-Z]\d+[A-Z0-9]*", RegexOptions.IgnoreCase)
                If m.Success AndAlso baseCode.Equals(m.Value, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Private Shared Function ComputeTokenScore(revitInfo As SegmentTokenInfo, pmsInfo As SegmentTokenInfo) As Double
            If revitInfo Is Nothing OrElse pmsInfo Is Nothing Then
                Return 0
            End If
            Dim commonGeneral As Integer = 0
            Dim commonMaterial As Integer = 0
            For Each t In revitInfo.Tokens
                If pmsInfo.Tokens.Contains(t) Then
                    Dim isMaterial = (revitInfo.MaterialTokens IsNot Nothing AndAlso revitInfo.MaterialTokens.Contains(t)) OrElse
                                     (pmsInfo.MaterialTokens IsNot Nothing AndAlso pmsInfo.MaterialTokens.Contains(t))
                    If isMaterial Then
                        commonMaterial += 1
                    Else
                        commonGeneral += 1
                    End If
                End If
            Next
            Return commonGeneral * 5 + commonMaterial * 2
        End Function

        Private Shared Function ComputeCandidateScore(revitInfo As SegmentTokenInfo, pmsInfo As SegmentTokenInfo) As CandidateScore
            Dim baseScore As Double = 0
            Dim variantScore As Double = 0
            Dim notes As New List(Of String)()
            Dim penalty As Double = 0

            If Not String.IsNullOrWhiteSpace(revitInfo.BaseCode) Then
                If Not String.IsNullOrWhiteSpace(pmsInfo.BaseCode) AndAlso revitInfo.BaseCode.Equals(pmsInfo.BaseCode, StringComparison.OrdinalIgnoreCase) Then
                    baseScore = If(pmsInfo.IsGroupLike, 40, 100)
                    notes.Add(String.Format("base:{0}", revitInfo.BaseCode))
                ElseIf ContainsBaseToken(pmsInfo.Raw, revitInfo.BaseCode) Then
                    baseScore = 40
                    notes.Add(String.Format("group:{0}", revitInfo.BaseCode))
                End If
            End If

            If Not String.IsNullOrWhiteSpace(revitInfo.VariantCode) Then
                If Not String.IsNullOrWhiteSpace(pmsInfo.VariantCode) AndAlso revitInfo.VariantCode.Equals(pmsInfo.VariantCode, StringComparison.OrdinalIgnoreCase) Then
                    variantScore = 80
                    notes.Add(String.Format("variant:{0}", revitInfo.VariantCode))
                ElseIf String.IsNullOrWhiteSpace(pmsInfo.VariantCode) Then
                    penalty -= 80
                Else
                    penalty -= 160
                End If
            Else
                If Not String.IsNullOrWhiteSpace(pmsInfo.VariantCode) Then
                    penalty -= 60
                End If
            End If

            Dim tokenScore = ComputeTokenScore(revitInfo, pmsInfo)
            Dim similarity = ComputeSimilarityBoost(revitInfo.Normalized, pmsInfo.Normalized)
            Dim orderSimilarity = ComputeOrderedSimilarityScore(revitInfo.Raw, pmsInfo.Raw)
            Dim total = baseScore + variantScore + tokenScore + similarity + orderSimilarity + penalty
            Dim lenDiff = Math.Abs(revitInfo.Normalized.Length - pmsInfo.Normalized.Length)
            Dim detail = String.Join(" + ", notes)

            Return New CandidateScore With {
                .BaseScore = baseScore,
                .VariantScore = variantScore,
                .TokenScore = tokenScore,
                .Similarity = similarity + orderSimilarity,
                .Score = total,
                .LenDiff = lenDiff,
                .Detail = detail
            }
        End Function

        Private Shared Function ComputeSimilarityBoost(a As String, b As String) As Double
            If String.IsNullOrWhiteSpace(a) OrElse String.IsNullOrWhiteSpace(b) Then
                Return 0
            End If
            Dim lcs = LongestCommonSubsequenceLength(a, b)
            Dim maxLen = Math.Max(a.Length, b.Length)
            If maxLen = 0 Then
                Return 0
            End If
            Dim ratio = CDbl(lcs) / CDbl(maxLen)
            Return Math.Max(0, Math.Min(20.0R, ratio * 20.0R))
        End Function

        Private Shared Function ComputeOrderedSimilarityScore(a As String, b As String) As Double
            If String.IsNullOrWhiteSpace(a) OrElse String.IsNullOrWhiteSpace(b) Then
                Return 0
            End If
            Dim cleanA = Regex.Replace(a.ToUpperInvariant(), "[^A-Z0-9]+", String.Empty)
            Dim cleanB = Regex.Replace(b.ToUpperInvariant(), "[^A-Z0-9]+", String.Empty)
            Dim lcs = LongestCommonSubsequenceLength(cleanA, cleanB)
            Dim maxLen = Math.Max(cleanA.Length, cleanB.Length)
            If maxLen = 0 Then
                Return 0
            End If
            Dim ratio = CDbl(lcs) / CDbl(maxLen)
            Return Math.Max(0, Math.Min(30.0R, ratio * 30.0R))
        End Function

        Private Shared Function LongestCommonSubsequenceLength(a As String, b As String) As Integer
            If String.IsNullOrEmpty(a) OrElse String.IsNullOrEmpty(b) Then
                Return 0
            End If
            Dim n As Integer = a.Length
            Dim m As Integer = b.Length
            Dim dp(n, m) As Integer
            For i As Integer = 1 To n
                For j As Integer = 1 To m
                    If a(i - 1) = b(j - 1) Then
                        dp(i, j) = dp(i - 1, j - 1) + 1
                    Else
                        dp(i, j) = Math.Max(dp(i - 1, j), dp(i, j - 1))
                    End If
                Next
            Next
            Return dp(n, m)
        End Function

        Public Shared Function ExpandGroupSelections(groups As List(Of MappingGroup), selections As List(Of GroupSelection)) As List(Of MappingSelection)
            Dim result As New List(Of MappingSelection)()
            If groups Is Nothing OrElse selections Is Nothing Then
                Return result
            End If
            Dim groupDict As New Dictionary(Of String, MappingGroup)(StringComparer.OrdinalIgnoreCase)
            For Each g In groups
                groupDict(g.GroupKey) = g
            Next
            For Each sel In selections
                Dim g As MappingGroup = Nothing
                If Not groupDict.TryGetValue(sel.GroupKey, g) Then
                    Continue For
                End If
                For Each u In g.Usages
                    result.Add(New MappingSelection With {
                        .File = u.File,
                        .PipeTypeName = u.PipeTypeName,
                        .RuleIndex = u.RuleIndex,
                        .SegmentId = u.SegmentId,
                        .SegmentKey = u.SegmentKey,
                        .SelectedClass = sel.SelectedClass,
                        .SelectedPmsSegment = sel.SelectedPmsSegment,
                        .MappingSource = If(String.IsNullOrWhiteSpace(sel.SelectionSource), "Manual", sel.SelectionSource)
                    })
                Next
            Next
            Return result
        End Function

        Public Shared Function RunCompare(extractData As DataSet,
                                          pmsData As List(Of PmsRow),
                                          mappings As List(Of MappingSelection),
                                          options As CompareOptions) As RunResult
            Dim res As New RunResult With {
                .MapTable = BuildMapTable(),
                .RevitSizeTable = BuildRevitSizeTable(),
                .PmsSizeTable = BuildPmsTableSkeleton(),
                .CompareTable = BuildCompareTable(),
                .ErrorTable = BuildErrorTable(),
                .SummaryTable = BuildSummaryTable()
            }

            If extractData Is Nothing OrElse Not extractData.Tables.Contains(TableRules) OrElse Not extractData.Tables.Contains(TableSizes) Then
                AddError(res.ErrorTable, "extract", "추출 데이터가 없습니다.")
                Return res
            End If

            If options Is Nothing Then
                options = New CompareOptions()
            End If

            Dim ndRound As Integer = options.NdRound
            Dim tol As Double = options.TolMm
            Dim doClassMatch As Boolean = options.ClassMatch

            Dim meta = extractData.Tables(TableMeta)
            If meta IsNot Nothing AndAlso meta.Rows.Count > 0 Then
                Dim ndVal = SafeDouble(meta.Rows(0)("NdRound"))
                If ndVal > 0 Then
                    ndRound = CInt(Math.Truncate(ndVal))
                End If
                Dim tolVal = SafeDouble(meta.Rows(0)("Tolerance"))
                If tolVal > 0 Then
                    tol = tolVal
                End If
            End If

            Dim pipeClassMap = BuildPipeTypeClassMap(extractData)
            Dim segmentClassMap = BuildSegmentClassMap(extractData)
            Dim routingClassMap = BuildRoutingClassMap(extractData)

            Dim sizeRows As New List(Of ExtractSizeRow)()
            Dim sizeTable As DataTable = extractData.Tables(TableSizes)
            For Each r As DataRow In sizeTable.Rows
                sizeRows.Add(New ExtractSizeRow With {
                    .File = SafeStr(r("File")),
                    .SegmentId = SafeIntObj(r("SegmentId")),
                    .SegmentKey = SafeStr(r("SegmentKey")),
                    .NdMm = SafeDouble(r("ND_mm")),
                    .IdMm = SafeDouble(r("ID_mm")),
                    .OdMm = SafeDouble(r("OD_mm"))
                })
            Next

            For Each s In sizeRows
                s.NdKey = Math.Round(s.NdMm, ndRound)
            Next

            Dim pmsRows As New List(Of PmsRow)()
            If pmsData IsNot Nothing Then
                For Each p In pmsData
                    pmsRows.Add(New PmsRow With {
                        .Class = p.Class,
                        .SegmentKey = p.SegmentKey,
                        .NdMm = p.NdMm,
                        .IdMm = p.IdMm,
                        .OdMm = p.OdMm
                    })
                Next
            End If

            Dim sizeByKey As New Dictionary(Of Tuple(Of String, String), List(Of ExtractSizeRow))(TupleComparer())
            For Each s In sizeRows
                Dim key = Tuple.Create(NormalizePath(s.File), s.SegmentKey)
                If Not sizeByKey.ContainsKey(key) Then
                    sizeByKey(key) = New List(Of ExtractSizeRow)()
                End If
                sizeByKey(key).Add(s)
            Next

            Dim pmsDict As New Dictionary(Of Tuple(Of String, String), List(Of PmsRow))(TupleComparer())
            For Each p In pmsRows
                Dim key = Tuple.Create(p.Class, p.SegmentKey)
                If Not pmsDict.ContainsKey(key) Then
                    pmsDict(key) = New List(Of PmsRow)()
                End If
                pmsDict(key).Add(p)
            Next

            For Each m In mappings
                Dim pipeKey = Tuple.Create(NormalizePath(m.File), m.PipeTypeName)
                Dim pipeClassRaw As String = GetDictValue(pipeClassMap, pipeKey)
                Dim segmentClassRaw As String = ExtractClassToken(If(GetDictValue(segmentClassMap, m.SegmentKey), m.SegmentKey))
                Dim routingSet As List(Of String) = Nothing
                Dim routingKey = Tuple.Create(NormalizePath(m.File), m.PipeTypeName)
                routingClassMap.TryGetValue(routingKey, routingSet)
                Dim classCheck = EvaluateClassMatch(doClassMatch, pipeClassRaw, segmentClassRaw, routingSet)
                If Not doClassMatch Then
                    pipeClassRaw = String.Empty
                    segmentClassRaw = String.Empty
                    routingSet = Nothing
                End If
                Dim routingSetStr As String = If(doClassMatch, String.Join("|", If(routingSet, New List(Of String)())), String.Empty)

                Dim mapRow = res.MapTable.NewRow()
                mapRow("File") = m.File
                mapRow("PipeTypeName") = m.PipeTypeName
                mapRow("SegmentRuleIndex") = m.RuleIndex
                mapRow("RevitSegmentKey") = m.SegmentKey
                mapRow("Selected_CLASS") = m.SelectedClass
                mapRow("Selected_PMS_SegmentKey") = m.SelectedPmsSegment
                mapRow("MappingSource") = If(String.IsNullOrWhiteSpace(m.MappingSource), "Manual", m.MappingSource)
                res.MapTable.Rows.Add(mapRow)

                Dim revKey = Tuple.Create(NormalizePath(m.File), m.SegmentKey)
                Dim revSizes As List(Of ExtractSizeRow) = Nothing
                sizeByKey.TryGetValue(revKey, revSizes)

                Dim pmsKey = Tuple.Create(m.SelectedClass, m.SelectedPmsSegment)
                Dim pmsSizes As List(Of PmsRow) = Nothing
                pmsDict.TryGetValue(pmsKey, pmsSizes)

                If String.IsNullOrWhiteSpace(m.SelectedPmsSegment) Then
                    AddMissingMappingRows(res.CompareTable, revSizes, m, ndRound, pipeClassRaw, segmentClassRaw, routingSetStr, classCheck.Status, classCheck.Note)
                    Continue For
                End If

                If revSizes Is Nothing OrElse revSizes.Count = 0 Then
                    AddCompareRow(res.CompareTable, m.File, m.PipeTypeName, m.RuleIndex, m.SegmentKey, m.SelectedClass, m.SelectedPmsSegment,
                                  0, 0, 0, 0, 0, 0, "MissingRevitRow", pipeClassRaw, segmentClassRaw, routingSetStr, classCheck.Status, classCheck.Note)
                    Continue For
                End If

                Dim revByNd As New Dictionary(Of Double, ExtractSizeRow)()
                For Each r In revSizes
                    Dim k = Math.Round(r.NdMm, ndRound)
                    If Not revByNd.ContainsKey(k) Then
                        revByNd(k) = r
                    End If
                Next

                Dim pmsByNd As New Dictionary(Of Double, PmsRow)()
                If pmsSizes IsNot Nothing Then
                    For Each p In pmsSizes
                        Dim k = Math.Round(p.NdMm, ndRound)
                        If Not pmsByNd.ContainsKey(k) Then
                            pmsByNd(k) = p
                        End If
                    Next
                End If

                Dim ndKeys As New HashSet(Of Double)()
                For Each k As Double In revByNd.Keys
                    ndKeys.Add(k)
                Next
                For Each k As Double In pmsByNd.Keys
                    ndKeys.Add(k)
                Next

                If ndKeys.Count = 0 Then
                    AddCompareRow(res.CompareTable, m.File, m.PipeTypeName, m.RuleIndex, m.SegmentKey, m.SelectedClass, m.SelectedPmsSegment,
                                  0, 0, 0, 0, 0, 0, If(pmsSizes Is Nothing OrElse pmsSizes.Count = 0, "MissingPmsRow", "MissingRevitRow"), pipeClassRaw, segmentClassRaw, routingSetStr, classCheck.Status, classCheck.Note)
                    Continue For
                End If

                For Each k In New List(Of Double)(ndKeys)
                    Dim r As ExtractSizeRow = Nothing
                    revByNd.TryGetValue(k, r)
                    Dim p As PmsRow = Nothing
                    pmsByNd.TryGetValue(k, p)

                    If r Is Nothing AndAlso p IsNot Nothing Then
                        AddCompareRow(res.CompareTable, m.File, m.PipeTypeName, m.RuleIndex, m.SegmentKey, m.SelectedClass, m.SelectedPmsSegment,
                                      0, 0, 0, p.NdMm, p.IdMm, p.OdMm, "MissingRevitRow", pipeClassRaw, segmentClassRaw, routingSetStr, classCheck.Status, classCheck.Note)
                        Continue For
                    End If

                    If r IsNot Nothing AndAlso p Is Nothing Then
                        AddCompareRow(res.CompareTable, m.File, m.PipeTypeName, m.RuleIndex, m.SegmentKey, m.SelectedClass, m.SelectedPmsSegment,
                                      r.NdMm, r.IdMm, r.OdMm, 0, 0, 0, "MissingPmsRow", pipeClassRaw, segmentClassRaw, routingSetStr, classCheck.Status, classCheck.Note)
                        Continue For
                    End If

                    Dim idDiff = Math.Abs(r.IdMm - p.IdMm)
                    Dim odDiff = Math.Abs(r.OdMm - p.OdMm)
                    Dim status As String = "OK"
                    If idDiff > tol AndAlso odDiff > tol Then
                        status = "Mismatch"
                    ElseIf idDiff > tol Then
                        status = "MismatchID"
                    ElseIf odDiff > tol Then
                        status = "MismatchOD"
                    End If

                    AddCompareRow(res.CompareTable, m.File, m.PipeTypeName, m.RuleIndex, m.SegmentKey, m.SelectedClass, m.SelectedPmsSegment,
                                  r.NdMm, r.IdMm, r.OdMm, p.NdMm, p.IdMm, p.OdMm, status, pipeClassRaw, segmentClassRaw, routingSetStr, classCheck.Status, classCheck.Note)
                Next
            Next

            For Each s In sizeRows
                Dim row = res.RevitSizeTable.NewRow()
                row("File") = s.File
                row("SegmentId") = s.SegmentId
                row("RevitSegmentKey") = s.SegmentKey
                row("ND_mm") = s.NdMm
                row("ID_mm") = s.IdMm
                row("OD_mm") = s.OdMm
                res.RevitSizeTable.Rows.Add(row)
            Next

            For Each p In pmsRows
                Dim row = res.PmsSizeTable.NewRow()
                row("CLASS") = p.Class
                row("PMS_SegmentKey") = p.SegmentKey
                row("ND_mm") = p.NdMm
                row("ID_mm") = p.IdMm
                row("OD_mm") = p.OdMm
                res.PmsSizeTable.Rows.Add(row)
            Next

            Dim summaryRow = res.SummaryTable.NewRow()
            Dim total As Integer = res.CompareTable.Rows.Count
            summaryRow("Total") = total
            summaryRow("OK") = CountStatus(res.CompareTable, "OK")
            summaryRow("Mismatch") = CountStatus(res.CompareTable, "Mismatch")
            summaryRow("MismatchID") = CountStatus(res.CompareTable, "MismatchID")
            summaryRow("MismatchOD") = CountStatus(res.CompareTable, "MismatchOD")
            summaryRow("MissingMapping") = CountStatus(res.CompareTable, "MissingMapping")
            summaryRow("MissingRevitRow") = CountStatus(res.CompareTable, "MissingRevitRow")
            summaryRow("MissingPmsRow") = CountStatus(res.CompareTable, "MissingPmsRow")
            res.SummaryTable.Rows.Add(summaryRow)

            Return res
        End Function

        ' ---------------------------
        ' Helpers
        ' ---------------------------
        Private Class ClassMatchResult
            Public Property Status As String = String.Empty
            Public Property Note As String = String.Empty
        End Class

        Private Shared Function BuildPipeTypeClassMap(extractData As DataSet) As Dictionary(Of Tuple(Of String, String), String)
            Dim map As New Dictionary(Of Tuple(Of String, String), String)(TupleComparer())
            If extractData Is Nothing OrElse Not extractData.Tables.Contains(TableRules) Then
                Return map
            End If
            Dim rules = extractData.Tables(TableRules)
            For Each r As DataRow In rules.Rows
                Dim key = Tuple.Create(NormalizePath(SafeStr(r("File"))), SafeStr(r("PipeTypeName")))
                If Not map.ContainsKey(key) Then
                    map(key) = ExtractClassToken(SafeStr(r("PipeTypeName")))
                End If
            Next
            Return map
        End Function

        Private Shared Function BuildSegmentClassMap(extractData As DataSet) As Dictionary(Of String, String)
            Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If extractData Is Nothing OrElse Not extractData.Tables.Contains(TableRules) Then
                Return map
            End If
            Dim rules = extractData.Tables(TableRules)
            For Each r As DataRow In rules.Rows
                Dim segKey = SafeStr(r("SegmentKey"))
                If Not map.ContainsKey(segKey) Then
                    map(segKey) = ExtractClassToken(segKey)
                End If
            Next
            Return map
        End Function

        Private Shared Function BuildRoutingClassMap(extractData As DataSet) As Dictionary(Of Tuple(Of String, String), List(Of String))
            Dim map As New Dictionary(Of Tuple(Of String, String), List(Of String))(TupleComparer())
            If extractData Is Nothing OrElse Not extractData.Tables.Contains(TableRouting) Then
                Return map
            End If
            Dim routing = extractData.Tables(TableRouting)
            For Each r As DataRow In routing.Rows
                Dim key = Tuple.Create(NormalizePath(SafeStr(r("File"))), SafeStr(r("PipeTypeName")))
                Dim rawPart = SafeStr(r("PartName"))
                Dim parts = rawPart.Split(New String() {"|"}, StringSplitOptions.None)
                For Each p In parts
                    Dim cls = ExtractClassToken(p)
                    Dim norm = NormalizeClassToken(cls)
                    If String.IsNullOrWhiteSpace(norm) Then
                        Continue For
                    End If
                    If Not map.ContainsKey(key) Then
                        map(key) = New List(Of String)()
                    End If
                    Dim existingNorms As New HashSet(Of String)(map(key).ConvertAll(Function(x) NormalizeClassToken(x)), StringComparer.OrdinalIgnoreCase)
                    If Not existingNorms.Contains(norm) Then
                        map(key).Add(cls)
                    End If
                Next
            Next
            Return map
        End Function

        Private Shared Function EvaluateClassMatch(doCheck As Boolean, pipeTypeClass As String, segmentClass As String, routingSet As List(Of String)) As ClassMatchResult
            Dim res As New ClassMatchResult()
            If Not doCheck Then
                Return res
            End If

            Dim pipeNorm = NormalizeClassToken(pipeTypeClass)
            Dim segNorm = NormalizeClassToken(segmentClass)
            Dim routingNorms As New List(Of String)()
            Dim routingRaw As New List(Of String)()
            If routingSet IsNot Nothing Then
                For Each r In routingSet
                    Dim norm = NormalizeClassToken(r)
                    If Not String.IsNullOrWhiteSpace(norm) Then
                        routingNorms.Add(norm)
                        routingRaw.Add(r)
                    End If
                Next
            End If

            If String.IsNullOrWhiteSpace(pipeNorm) AndAlso String.IsNullOrWhiteSpace(segNorm) AndAlso routingNorms.Count = 0 Then
                res.Status = "N/A"
                Return res
            End If

            Dim expected = If(Not String.IsNullOrWhiteSpace(segNorm), segNorm, pipeNorm)
            Dim noteParts As New List(Of String)()

            If Not String.IsNullOrWhiteSpace(pipeNorm) AndAlso Not String.IsNullOrWhiteSpace(segNorm) AndAlso Not pipeNorm.Equals(segNorm, StringComparison.OrdinalIgnoreCase) Then
                noteParts.Add(String.Format("PipeType:{0} vs Segment:{1}", pipeTypeClass, segmentClass))
            End If

            If routingNorms.Count > 0 AndAlso Not String.IsNullOrWhiteSpace(expected) Then
                Dim expectedRaw As String = If(String.IsNullOrWhiteSpace(segmentClass), pipeTypeClass, segmentClass)
                For i As Integer = 0 To routingNorms.Count - 1
                    If Not routingNorms(i).Equals(expected, StringComparison.OrdinalIgnoreCase) Then
                        Dim rawVal As String = routingRaw(i)
                        noteParts.Add(String.Format("Routing:{0} vs {1}", rawVal, expectedRaw))
                    End If
                Next
            End If

            If noteParts.Count > 0 Then
                res.Status = "Mismatch"
                res.Note = String.Join("; ", noteParts)
            Else
                res.Status = If(String.IsNullOrWhiteSpace(expected) AndAlso routingNorms.Count = 0, "N/A", "OK")
                res.Note = String.Empty
            End If

            Return res
        End Function

        Private Shared Function ExtractClassToken(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then
                Return String.Empty
            End If

            Dim firstPass As String = String.Empty
            Dim parts = text.Split(","c)
            For Each p In parts
                Dim trimmed = p.Trim()
                If trimmed.IndexOf("("c) >= 0 AndAlso trimmed.IndexOf(")"c) > trimmed.IndexOf("("c) Then
                    firstPass = trimmed
                    Exit For
                End If
            Next

            If String.IsNullOrWhiteSpace(firstPass) Then
                Dim m = Regex.Match(text, "^\s*([A-Za-z0-9]+\(.*?\))")
                If m.Success AndAlso m.Groups.Count > 1 Then
                    firstPass = m.Groups(1).Value.Trim()
                End If
            End If

            If String.IsNullOrWhiteSpace(firstPass) AndAlso parts.Length > 0 Then
                firstPass = parts(0).Trim()
            End If

            Return firstPass
        End Function

        Private Shared Function NormalizeClassToken(token As String) As String
            If String.IsNullOrWhiteSpace(token) Then
                Return String.Empty
            End If
            Dim t = token.Trim().ToUpperInvariant()
            t = t.Replace(" - REF.", String.Empty).Replace("- REF.", String.Empty).Replace("-REF.", String.Empty).Replace(" -REF.", String.Empty)
            While t.Contains("  ")
                t = t.Replace("  ", " ")
            End While
            Return t.Trim()
        End Function

        Private Shared Function GetDictValue(Of TKey, TValue)(dict As Dictionary(Of TKey, TValue), key As TKey) As TValue
            Dim val As TValue = Nothing
            If dict IsNot Nothing AndAlso dict.TryGetValue(key, val) Then
                Return val
            End If
            Return Nothing
        End Function

        Private Shared Function BuildOpenOptions(opts As ExtractOptions, filePath As String) As OpenOptions
            Dim opt As New OpenOptions()
            opt.Audit = False
            opt.AllowOpeningLocalByWrongUser = True
            opt.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach

            If opts IsNot Nothing Then
                opt.DetachFromCentralOption = If(opts.DetachFromCentral, DetachFromCentralOption.DetachAndPreserveWorksets, DetachFromCentralOption.DoNotDetach)
            End If

            ApplyWorksetConfiguration(opt, filePath)

            Return opt
        End Function

        Private Shared Sub ApplyWorksetConfiguration(opt As OpenOptions, filePath As String)
            If opt Is Nothing Then
                Return
            End If

            If String.IsNullOrWhiteSpace(filePath) Then
                Return
            End If

            Try
                Dim fileInfo = BasicFileInfo.Extract(filePath)
                If fileInfo IsNot Nothing AndAlso fileInfo.IsWorkshared Then
                    Dim wsConfig As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                    opt.SetOpenWorksetsConfiguration(wsConfig)
                End If
            Catch
                ' 워크셰어링 여부 확인 실패는 무시하고 기본 옵션으로 계속 진행
            End Try
        End Sub

        Private Shared Function BuildMetaTable() As DataTable
            Dim t As New DataTable(TableMeta)
            t.Columns.Add("NdRound", GetType(Integer))
            t.Columns.Add("CreatedAt", GetType(String))
            t.Columns.Add("Tolerance", GetType(Double))
            t.Columns.Add("ToolVersion", GetType(String))
            Return t
        End Function

        Private Shared Function BuildFileTable() As DataTable
            Dim t As New DataTable(TableFiles)
            t.Columns.Add("File", GetType(String))
            t.Columns.Add("FileName", GetType(String))
            t.Columns.Add("ExtractedAt", GetType(String))
            Return t
        End Function

        Private Shared Function BuildRuleTable() As DataTable
            Dim t As New DataTable(TableRules)
            t.Columns.Add("File", GetType(String))
            t.Columns.Add("PipeTypeName", GetType(String))
            t.Columns.Add("RuleIndex", GetType(Integer))
            t.Columns.Add("SegmentId", GetType(Integer))
            t.Columns.Add("SegmentKey", GetType(String))
            Return t
        End Function

        Private Shared Function BuildSizeTable() As DataTable
            Dim t As New DataTable(TableSizes)
            t.Columns.Add("File", GetType(String))
            t.Columns.Add("SegmentId", GetType(Integer))
            t.Columns.Add("SegmentKey", GetType(String))
            t.Columns.Add("ND_mm", GetType(Double))
            t.Columns.Add("ID_mm", GetType(Double))
            t.Columns.Add("OD_mm", GetType(Double))
            Return t
        End Function

        Private Shared Function BuildRoutingTable() As DataTable
            Dim t As New DataTable(TableRouting)
            t.Columns.Add("File", GetType(String))
            t.Columns.Add("PipeTypeName", GetType(String))
            t.Columns.Add("RuleGroup", GetType(String))
            t.Columns.Add("RuleIndex", GetType(Integer))
            t.Columns.Add("RuleType", GetType(String))
            t.Columns.Add("PartId", GetType(Integer))
            t.Columns.Add("PartName", GetType(String))
            t.Columns.Add("TypeName", GetType(String))
            Return t
        End Function

        Private Shared Function BuildMapTable() As DataTable
            Dim t As New DataTable("PipeTypeSegmentMap")
            t.Columns.Add("File", GetType(String))
            t.Columns.Add("PipeTypeName", GetType(String))
            t.Columns.Add("SegmentRuleIndex", GetType(Integer))
            t.Columns.Add("RevitSegmentKey", GetType(String))
            t.Columns.Add("Selected_CLASS", GetType(String))
            t.Columns.Add("Selected_PMS_SegmentKey", GetType(String))
            t.Columns.Add("MappingSource", GetType(String))
            Return t
        End Function

        Private Shared Function BuildRevitSizeTable() As DataTable
            Dim t As New DataTable("SegmentSizeRaw_Revit")
            t.Columns.Add("File", GetType(String))
            t.Columns.Add("SegmentId", GetType(Integer))
            t.Columns.Add("RevitSegmentKey", GetType(String))
            t.Columns.Add("ND_mm", GetType(Double))
            t.Columns.Add("ID_mm", GetType(Double))
            t.Columns.Add("OD_mm", GetType(Double))
            Return t
        End Function

        Private Shared Function BuildPmsTableSkeleton() As DataTable
            Dim t As New DataTable("SegmentSizeRaw_PMS")
            t.Columns.Add("CLASS", GetType(String))
            t.Columns.Add("PMS_SegmentKey", GetType(String))
            t.Columns.Add("ND_mm", GetType(Double))
            t.Columns.Add("ID_mm", GetType(Double))
            t.Columns.Add("OD_mm", GetType(Double))
            Return t
        End Function

        Private Shared Function BuildCompareTable() As DataTable
            Dim t As New DataTable("SizeCompare")
            t.Columns.Add("File", GetType(String))
            t.Columns.Add("PipeTypeName", GetType(String))
            t.Columns.Add("SegmentRuleIndex", GetType(Integer))
            t.Columns.Add("RevitSegmentKey", GetType(String))
            t.Columns.Add("CLASS", GetType(String))
            t.Columns.Add("PMS_SegmentKey", GetType(String))
            t.Columns.Add("ND_mm", GetType(Double))
            t.Columns.Add("PMS_ND", GetType(Double))
            t.Columns.Add("Revit_ID", GetType(Double))
            t.Columns.Add("Revit_OD", GetType(Double))
            t.Columns.Add("PMS_ID", GetType(Double))
            t.Columns.Add("PMS_OD", GetType(Double))
            t.Columns.Add("Diff_ID", GetType(Double))
            t.Columns.Add("Diff_OD", GetType(Double))
            t.Columns.Add("Status", GetType(String))
            t.Columns.Add("PipeTypeClass", GetType(String))
            t.Columns.Add("SegmentClass", GetType(String))
            t.Columns.Add("RoutingClassSet", GetType(String))
            t.Columns.Add("ClassMatchStatus", GetType(String))
            t.Columns.Add("ClassMatchNote", GetType(String))
            Return t
        End Function

        Private Shared Function BuildErrorTable() As DataTable
            Dim t As New DataTable("Error")
            t.Columns.Add("Stage", GetType(String))
            t.Columns.Add("Message", GetType(String))
            t.Columns.Add("ExceptionSummary", GetType(String))
            Return t
        End Function

        Private Shared Function BuildSummaryTable() As DataTable
            Dim t As New DataTable("Summary")
            t.Columns.Add("Total", GetType(Integer))
            t.Columns.Add("OK", GetType(Integer))
            t.Columns.Add("Mismatch", GetType(Integer))
            t.Columns.Add("MismatchID", GetType(Integer))
            t.Columns.Add("MismatchOD", GetType(Integer))
            t.Columns.Add("MissingMapping", GetType(Integer))
            t.Columns.Add("MissingRevitRow", GetType(Integer))
            t.Columns.Add("MissingPmsRow", GetType(Integer))
            Return t
        End Function

        Private Shared Sub AddCompareRow(table As DataTable,
                                         file As String,
                                         pipeType As String,
                                         ruleIdx As Integer,
                                         revSeg As String,
                                         cls As String,
                                         pmsSeg As String,
                                         revNd As Double,
                                         revId As Double,
                                         revOd As Double,
                                         pmsNd As Double,
                                         pmsId As Double,
                                         pmsOd As Double,
                                         status As String,
                                         Optional pipeTypeClass As String = "",
                                         Optional segmentClass As String = "",
                                         Optional routingClassSet As String = "",
                                         Optional classMatchStatus As String = "",
                                         Optional classMatchNote As String = "")
            Dim row = table.NewRow()
            row("File") = file
            row("PipeTypeName") = pipeType
            row("SegmentRuleIndex") = ruleIdx
            row("RevitSegmentKey") = revSeg
            row("CLASS") = cls
            row("PMS_SegmentKey") = pmsSeg
            row("ND_mm") = revNd
            row("PMS_ND") = pmsNd
            row("Revit_ID") = revId
            row("Revit_OD") = revOd
            row("PMS_ID") = pmsId
            row("PMS_OD") = pmsOd
            row("Diff_ID") = revId - pmsId
            row("Diff_OD") = revOd - pmsOd
            row("Status") = status
            row("PipeTypeClass") = pipeTypeClass
            row("SegmentClass") = segmentClass
            row("RoutingClassSet") = routingClassSet
            row("ClassMatchStatus") = classMatchStatus
            row("ClassMatchNote") = classMatchNote
            table.Rows.Add(row)
        End Sub

        Private Shared Sub AddMissingMappingRows(table As DataTable,
                                                 revSizes As List(Of ExtractSizeRow),
                                                 m As MappingSelection,
                                                 ndRound As Integer,
                                                 Optional pipeTypeClass As String = "",
                                                 Optional segmentClass As String = "",
                                                 Optional routingClassSet As String = "",
                                                 Optional classMatchStatus As String = "",
                                                 Optional classMatchNote As String = "")
            If revSizes IsNot Nothing AndAlso revSizes.Count > 0 Then
                For Each r In revSizes
                    AddCompareRow(table, m.File, m.PipeTypeName, m.RuleIndex, m.SegmentKey, m.SelectedClass, m.SelectedPmsSegment,
                                  r.NdMm, r.IdMm, r.OdMm, 0, 0, 0, "MissingMapping", pipeTypeClass, segmentClass, routingClassSet, classMatchStatus, classMatchNote)
                Next
            Else
                AddCompareRow(table, m.File, m.PipeTypeName, m.RuleIndex, m.SegmentKey, m.SelectedClass, m.SelectedPmsSegment,
                              0, 0, 0, 0, 0, 0, "MissingMapping", pipeTypeClass, segmentClass, routingClassSet, classMatchStatus, classMatchNote)
            End If
        End Sub

        Public Shared Function BuildClassCheckRows(mapTable As DataTable) As List(Of Dictionary(Of String, Object))
            Dim list As New List(Of Dictionary(Of String, Object))()
            If mapTable Is Nothing Then
                Return list
            End If

            Dim hasFile As Boolean = mapTable.Columns.Contains("File")
            Dim hasPipeType As Boolean = mapTable.Columns.Contains("PipeTypeName")
            Dim hasSegment As Boolean = mapTable.Columns.Contains("RevitSegmentKey")

            For Each r As DataRow In mapTable.Rows
                Dim fileName = SafeFileName(If(hasFile, SafeStr(r("File")), String.Empty))
                Dim pipeType = If(hasPipeType, SafeStr(r("PipeTypeName")), String.Empty)
                Dim segment = If(hasSegment, SafeStr(r("RevitSegmentKey")), String.Empty)
                Dim pipeCls = NormalizeClassToken(ExtractClassToken(pipeType))
                Dim segCls = NormalizeClassToken(ExtractClassToken(segment))

                Dim result As String
                If String.IsNullOrWhiteSpace(pipeCls) OrElse String.IsNullOrWhiteSpace(segCls) Then
                    result = "N/A"
                ElseIf pipeCls.Equals(segCls, StringComparison.OrdinalIgnoreCase) Then
                    result = "OK"
                Else
                    result = "Mismatch"
                End If

                Dim item As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                    {"File", fileName},
                    {"PipeType", pipeType},
                    {"Segment", segment},
                    {"Class검토결과", result}
                }
                list.Add(item)
            Next

            Return list.OrderBy(Function(x) SafeStr(x("File"))).
                        ThenBy(Function(x) SafeStr(x("PipeType"))).
                        ThenBy(Function(x) SafeStr(x("Segment"))).ToList()
        End Function

        Public Shared Function BuildSizeCheckRows(compareTable As DataTable) As List(Of Dictionary(Of String, Object))
            Dim list As New List(Of Dictionary(Of String, Object))()
            If compareTable Is Nothing Then
                Return list
            End If

            Dim hasFile As Boolean = compareTable.Columns.Contains("File")
            Dim hasPipeType As Boolean = compareTable.Columns.Contains("PipeTypeName")
            Dim hasNd As Boolean = compareTable.Columns.Contains("ND_mm")
            Dim hasRevId As Boolean = compareTable.Columns.Contains("Revit_ID")
            Dim hasRevOd As Boolean = compareTable.Columns.Contains("Revit_OD")
            Dim hasPmsNd As Boolean = compareTable.Columns.Contains("PMS_ND")
            Dim hasPmsId As Boolean = compareTable.Columns.Contains("PMS_ID")
            Dim hasPmsOd As Boolean = compareTable.Columns.Contains("PMS_OD")
            Dim hasStatus As Boolean = compareTable.Columns.Contains("Status")
            Dim hasRevSeg As Boolean = compareTable.Columns.Contains("RevitSegmentKey")
            Dim hasPmsSeg As Boolean = compareTable.Columns.Contains("PMS_SegmentKey")

            For Each r As DataRow In compareTable.Rows
                Dim item As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                    {"FileName", SafeFileName(If(hasFile, SafeStr(r("File")), String.Empty))},
                    {"PipeType", If(hasPipeType, SafeStr(r("PipeTypeName")), String.Empty)},
                    {"RevitSegment", If(hasRevSeg, SafeStr(r("RevitSegmentKey")), String.Empty)},
                    {"PMSCompared", If(hasPmsSeg, SafeStr(r("PMS_SegmentKey")), String.Empty)},
                    {"ND", If(hasNd AndAlso Not r.IsNull("ND_mm"), r("ND_mm"), Nothing)},
                    {"ID", If(hasRevId AndAlso Not r.IsNull("Revit_ID"), r("Revit_ID"), Nothing)},
                    {"OD", If(hasRevOd AndAlso Not r.IsNull("Revit_OD"), r("Revit_OD"), Nothing)},
                    {"PMS ND", If(hasPmsNd AndAlso Not r.IsNull("PMS_ND"), r("PMS_ND"), Nothing)},
                    {"PMS ID", If(hasPmsId AndAlso Not r.IsNull("PMS_ID"), r("PMS_ID"), Nothing)},
                    {"PMS OD", If(hasPmsOd AndAlso Not r.IsNull("PMS_OD"), r("PMS_OD"), Nothing)},
                    {"Result", MapSizeStatus(If(hasStatus, SafeStr(r("Status")), String.Empty))}
                }
                list.Add(item)
            Next

            Return list.OrderBy(Function(x) SafeStr(x("FileName"))).
                        ThenBy(Function(x) SafeStr(x("PipeType"))).
                        ThenBy(Function(x) SafeStr(x("ND"))).ToList()
        End Function

        Public Shared Function BuildRoutingClassRows(extractData As DataSet) As List(Of Dictionary(Of String, Object))
            Dim list As New List(Of Dictionary(Of String, Object))()
            If extractData Is Nothing OrElse Not extractData.Tables.Contains(TableRouting) Then
                Return list
            End If

            Dim routing = extractData.Tables(TableRouting)
            For Each r As DataRow In routing.Rows
                Dim fileName = SafeFileName(SafeStr(r("File")))
                Dim pipeType = SafeStr(r("PipeTypeName"))
                Dim partRaw = SafeStr(r("PartName"))
                Dim typeName As String = String.Empty
                If routing.Columns.Contains("TypeName") Then
                    typeName = SafeStr(r("TypeName"))
                End If
                Dim partLabel = ExtractRoutingPartLabel(partRaw)

                Dim pipeClass = NormalizeClassToken(ExtractClassToken(pipeType))
                Dim partClasses = ExtractRoutingClasses(partRaw)

                Dim status As String
                If String.IsNullOrWhiteSpace(pipeClass) OrElse partClasses.Count = 0 Then
                    status = "N/A"
                Else
                    Dim matched As Boolean = False
                    For Each cls In partClasses
                        If cls.Equals(pipeClass, StringComparison.OrdinalIgnoreCase) Then
                            matched = True
                            Exit For
                        End If
                    Next
                    status = If(matched, "OK", "Mismatch")
                End If

                Dim item As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                    {"File", fileName},
                    {"PipeType", pipeType},
                    {"Part", partLabel},
                    {"Type", typeName},
                    {"Class검토", status}
                }
                list.Add(item)
            Next

            Return list.OrderBy(Function(x) SafeStr(x("File"))).
                        ThenBy(Function(x) SafeStr(x("PipeType"))).
                        ThenBy(Function(x) SafeStr(x("Part"))).ToList()
        End Function

        Private Shared Function MapSizeStatus(status As String) As String
            Select Case status
                Case "OK"
                    Return "OK"
                Case "MismatchID"
                    Return "MismatchID"
                Case "MismatchOD"
                    Return "MismatchOD"
                Case "Mismatch"
                    Return "Mismatch"
                Case "MissingPmsRow", "MissingMapping"
                    Return "MissingPMS"
                Case "MissingRevitRow"
                    Return "MissingRevit"
                Case Else
                    Return "N/A"
            End Select
        End Function

        Private Shared Function NormalizeNumberText(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then
                Return String.Empty
            End If
            Dim val As Double
            If Double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, val) Then
                If Math.Abs(val) < Double.Epsilon Then
                    Return String.Empty
                End If
                Return val.ToString("0.###############", CultureInfo.InvariantCulture)
            End If
            Return text
        End Function

        Private Shared Function ExtractRoutingPartLabel(partName As String) As String
            If String.IsNullOrWhiteSpace(partName) Then
                Return String.Empty
            End If
            Dim first = partName.Split(New String() {"|"}, StringSplitOptions.None)(0).Trim()
            Dim idx = first.IndexOf(","c)
            If idx >= 0 Then
                Return first.Substring(0, idx).Trim()
            End If
            Return first
        End Function

        Private Shared Function ExtractRoutingClasses(partName As String) As List(Of String)
            Dim list As New List(Of String)()
            If String.IsNullOrWhiteSpace(partName) Then
                Return list
            End If
            Dim segments = partName.Split(New String() {"|"}, StringSplitOptions.None)
            Dim added As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each seg In segments
                Dim cls = NormalizeClassToken(ExtractClassToken(seg))
                If String.IsNullOrWhiteSpace(cls) Then
                    Continue For
                End If
                If Not added.Contains(cls) Then
                    added.Add(cls)
                    list.Add(cls)
                End If
            Next
            Return list
        End Function

        Private Shared Function SafeFileName(path As String) As String
            If String.IsNullOrWhiteSpace(path) Then
                Return String.Empty
            End If
            Try
                Return System.IO.Path.GetFileName(path)
            Catch
                Return path
            End Try
        End Function

        Private Shared Function DetectHeader(sh As ISheet) As Dictionary(Of String, Integer)
            Dim map As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim head As IRow = sh.GetRow(0)
            If head Is Nothing Then
                Return map
            End If
            For i As Integer = 0 To head.LastCellNum - 1
                Dim name = CellStr(head, i).Trim()
                Dim key = NormalizeHeader(name)
                If Not String.IsNullOrEmpty(key) AndAlso Not map.ContainsKey(key) Then
                    map(key) = i
                End If
            Next
            Return map
        End Function

        Private Shared Function NormalizeHeader(name As String) As String
            Dim n = (If(name, String.Empty)).Trim()
            If String.IsNullOrEmpty(n) Then
                Return String.Empty
            End If
            Select Case n.ToLowerInvariant()
                Case "class", "discipline", "trade"
                    Return "class"
                Case "segment", "segmentkey", "segmentname", "segment key", "segment_name", "pms_segment", "seg_pms", "seg", "segkey"
                    Return "segment"
                Case "nd", "nominaldiameter", "nominal", "nominal diameter", "nd_mm", "nd_in"
                    Return "nd"
                Case "id", "innerdiameter", "inner", "inner diameter", "id_mm", "id_in"
                    Return "id"
                Case "od", "outerdiameter", "outer", "outer diameter", "od_mm", "od_in"
                    Return "od"
                Case Else
                    Return String.Empty
            End Select
        End Function

        Private Shared Function CollectPipeTypeSegmentCandidates(doc As RvtDB.Document, filePath As String) As List(Of PreparePipeInfo)
            Dim result As New List(Of PreparePipeInfo)()
            If doc Is Nothing Then
                Return result
            End If

            Dim typesCol As New FilteredElementCollector(doc)
            typesCol.OfClass(GetType(PipeType))
            Dim byPipe As New Dictionary(Of String, List(Of PrepareRow))(StringComparer.OrdinalIgnoreCase)

            For Each el As RvtDB.Element In typesCol
                Dim pt As PipeType = TryCast(el, PipeType)
                If pt Is Nothing Then
                    Continue For
                End If
                Dim rpm = pt.RoutingPreferenceManager
                If rpm Is Nothing Then
                    Continue For
                End If

                Dim count = rpm.GetNumberOfRules(RoutingPreferenceRuleGroupType.Segments)
                For idx As Integer = 0 To count - 1
                    Dim rule = rpm.GetRule(RoutingPreferenceRuleGroupType.Segments, idx)
                    If rule Is Nothing Then
                        Continue For
                    End If
                    Dim segId = rule.MEPPartId
                    Dim segName = ToSegmentKey(doc, segId)
                    Dim key = pt.Name
                    If Not byPipe.ContainsKey(key) Then
                        byPipe(key) = New List(Of PrepareRow)()
                    End If
                    byPipe(key).Add(New PrepareRow With {
                        .File = filePath,
                        .PipeTypeName = pt.Name,
                        .RuleIndex = idx,
                        .SegmentId = segId.IntegerValue,
                        .SegmentName = segName,
                        .SegmentKey = segName
                    })
                Next
            Next

            For Each kv In byPipe
                Dim list = kv.Value
                list.Sort(Function(a, b) a.RuleIndex.CompareTo(b.RuleIndex))
                Dim defaultIdx As Integer = If(list.Count > 0, list(0).RuleIndex, 0)
                result.Add(New PreparePipeInfo With {
                    .File = filePath,
                    .PipeTypeName = kv.Key,
                    .Candidates = list,
                    .DefaultRuleIndex = defaultIdx
                })
            Next

            Return result
        End Function

        Private Shared Function CollectSegmentSizes(doc As RvtDB.Document, segIds As IEnumerable(Of Integer), filePath As String, ndRound As Integer) As List(Of ExtractSizeRow)
            Dim res As New List(Of ExtractSizeRow)()
            If doc Is Nothing Then
                Return res
            End If
            Dim cache As New Dictionary(Of Integer, PipeSegment)()

            For Each id In segIds
                Dim eid As New ElementId(id)
                Dim seg As PipeSegment = Nothing
                If Not cache.TryGetValue(id, seg) Then
                    seg = TryCast(doc.GetElement(eid), PipeSegment)
                    cache(id) = seg
                End If
                If seg Is Nothing Then
                    Continue For
                End If

                Dim sizes = seg.GetSizes()
                If sizes Is Nothing Then
                    Continue For
                End If

                For Each sz In sizes
                    Dim ndMm = sz.NominalDiameter * FeetToMm
                    Dim idMm = sz.InnerDiameter * FeetToMm
                    Dim odMm = sz.OuterDiameter * FeetToMm
                    res.Add(New ExtractSizeRow With {
                        .File = filePath,
                        .SegmentId = id,
                        .SegmentKey = ToSegmentKey(doc, eid),
                        .NdMm = ndMm,
                        .IdMm = idMm,
                        .OdMm = odMm,
                        .NdKey = Math.Round(ndMm, ndRound)
                    })
                Next
            Next
            Return res
        End Function

        Private Shared Function CollectRouting(doc As RvtDB.Document, filePath As String) As List(Of RoutingRow)
            Dim res As New List(Of RoutingRow)()
            If doc Is Nothing Then
                Return res
            End If

            Dim col As New FilteredElementCollector(doc)
            col.OfClass(GetType(PipeType))
            For Each el As RvtDB.Element In col
                Dim pt As PipeType = TryCast(el, PipeType)
                If pt Is Nothing Then
                    Continue For
                End If
                Dim rpm = pt.RoutingPreferenceManager
                If rpm Is Nothing Then
                    Continue For
                End If

                For Each obj As Object In [Enum].GetValues(GetType(RoutingPreferenceRuleGroupType))
                    Dim group = CType(obj, RoutingPreferenceRuleGroupType)
                    Dim count = rpm.GetNumberOfRules(group)
                    For i As Integer = 0 To count - 1
                        Dim rule = rpm.GetRule(group, i)
                        If rule Is Nothing Then
                            Continue For
                        End If
                        Dim partId = rule.MEPPartId
                        Dim partName = ToSegmentKey(doc, partId)
                        Dim typeName = ExtractRoutingTypeName(doc, partId)
                        res.Add(New RoutingRow With {
                            .File = filePath,
                            .PipeTypeName = pt.Name,
                            .RuleGroup = group.ToString(),
                            .RuleIndex = i,
                            .RuleType = rule.GetType().Name,
                            .PartId = partId.IntegerValue,
                            .PartName = partName,
                            .TypeName = typeName
                        })
                    Next
                Next
            Next

            Return res
        End Function

        Private Class RoutingRow
            Public Property [File] As String = String.Empty
            Public Property PipeTypeName As String = String.Empty
            Public Property RuleGroup As String = String.Empty
            Public Property RuleIndex As Integer
            Public Property RuleType As String = String.Empty
            Public Property PartId As Integer
            Public Property PartName As String = String.Empty
            Public Property TypeName As String = String.Empty
            Public ReadOnly Property FileName As String
                Get
                    If String.IsNullOrWhiteSpace([File]) Then
                        Return String.Empty
                    End If
                    Return Path.GetFileName([File])
                End Get
            End Property
        End Class

        Private Shared Function ExtractRoutingTypeName(doc As RvtDB.Document, partId As RvtDB.ElementId) As String
            If doc Is Nothing Then
                Return String.Empty
            End If
            Try
                Dim el = doc.GetElement(partId)
                If el Is Nothing Then
                    Return String.Empty
                End If
                Dim val As String = String.Empty
                Try
                    Dim typeParam As RvtDB.Parameter = el.LookupParameter("Type")
                    If typeParam Is Nothing Then
                        typeParam = el.Parameter(RvtDB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                    End If
                    If typeParam IsNot Nothing Then
                        val = typeParam.AsString()
                    End If
                Catch
                End Try
                If String.IsNullOrWhiteSpace(val) Then
                    Dim et As ElementType = TryCast(el, ElementType)
                    If et IsNot Nothing Then
                        val = et.Name
                    End If
                End If
                If String.IsNullOrWhiteSpace(val) Then
                    val = el.Name
                End If
                Return If(val, String.Empty)
            Catch
                Return String.Empty
            End Try
        End Function

        Private Shared Function ToSegmentKey(doc As RvtDB.Document, segId As RvtDB.ElementId) As String
            If doc Is Nothing Then
                Return String.Empty
            End If
            Try
                Dim el As RvtDB.Element = doc.GetElement(segId)
                If el Is Nothing Then
                    Return segId.IntegerValue.ToString(CultureInfo.InvariantCulture)
                End If
                Dim fam As String = String.Empty
                Dim typ As String = String.Empty
                Try
                    Dim revEl As RvtDB.Element = el
                    Dim famParam As RvtDB.Parameter = revEl.LookupParameter("Family")
                    If famParam Is Nothing Then
                        famParam = revEl.Parameter(RvtDB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                    End If
                    If famParam IsNot Nothing Then
                        fam = famParam.AsString()
                    End If
                Catch
                End Try
                Try
                    Dim revEl As RvtDB.Element = el
                    Dim typeParam As RvtDB.Parameter = revEl.LookupParameter("Type")
                    If typeParam Is Nothing Then
                        typeParam = revEl.Parameter(RvtDB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                    End If
                    If typeParam IsNot Nothing Then
                        typ = typeParam.AsString()
                    End If
                Catch
                End Try
                Dim name As String = el.Name
                Dim parts As New List(Of String)()
                If Not String.IsNullOrWhiteSpace(name) Then
                    parts.Add(name)
                End If
                If Not String.IsNullOrWhiteSpace(fam) Then
                    parts.Add(fam)
                End If
                If Not String.IsNullOrWhiteSpace(typ) Then
                    parts.Add(typ)
                End If
                Dim joined = String.Join(" | ", parts)
                If String.IsNullOrWhiteSpace(joined) Then
                    joined = segId.IntegerValue.ToString(CultureInfo.InvariantCulture)
                End If
                Return joined
            Catch
                Return segId.IntegerValue.ToString(CultureInfo.InvariantCulture)
            End Try
        End Function

        Private Shared Function NormalizeSegmentGroupKey(key As String) As String
            If String.IsNullOrWhiteSpace(key) Then
                Return String.Empty
            End If
            Dim trimmed = key.Trim()
            Dim withoutRef = Regex.Replace(trimmed, "\s*-\s*Ref\.?\s*$", String.Empty, RegexOptions.IgnoreCase)
            Dim compact = Regex.Replace(withoutRef, "\s+", " ").Trim()
            Return compact
        End Function

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

        Private Shared Function TupleComparer() As IEqualityComparer(Of Tuple(Of String, String))
            Return New TupleComparerImpl()
        End Function

        Private Shared Function CountStatus(t As DataTable, status As String) As Integer
            If t Is Nothing Then
                Return 0
            End If
            Dim cnt As Integer = 0
            For Each r As DataRow In t.Rows
                Dim s = SafeStr(r("Status"))
                If String.Equals(s, status, StringComparison.OrdinalIgnoreCase) Then
                    cnt += 1
                End If
            Next
            Return cnt
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

        Private Shared Sub AddError(t As DataTable, stage As String, msg As String)
            Dim r = t.NewRow()
            r("Stage") = stage
            r("Message") = msg
            r("ExceptionSummary") = String.Empty
            t.Rows.Add(r)
        End Sub

        Private Class ExtractSizeRow
            Public Property [File] As String = String.Empty
            Public Property SegmentId As Integer
            Public Property SegmentKey As String = String.Empty
            Public Property NdMm As Double
            Public Property IdMm As Double
            Public Property OdMm As Double
            Public Property NdKey As Double
        End Class

        Private Class PrepareRow
            Public Property [File] As String = String.Empty
            Public Property PipeTypeName As String = String.Empty
            Public Property RuleIndex As Integer
            Public Property SegmentId As Integer
            Public Property SegmentName As String = String.Empty
            Public Property SegmentKey As String = String.Empty
        End Class

        Private Class PreparePipeInfo
            Public Property [File] As String = String.Empty
            Public Property PipeTypeName As String = String.Empty
            Public Property Candidates As List(Of PrepareRow)
            Public Property DefaultRuleIndex As Integer
            Public Property SegmentClass As String = String.Empty
            Public Property SegmentPipeType As String = String.Empty
            Public Property SegmentIdParam As String = String.Empty
            Public Property Sizes As List(Of ExtractSizeRow) = New List(Of ExtractSizeRow)()

            Public ReadOnly Property FileName As String
                Get
                    If String.IsNullOrWhiteSpace([File]) Then
                        Return String.Empty
                    End If
                    Return Path.GetFileName([File])
                End Get
            End Property

            Public ReadOnly Property RuleIndex As Integer
                Get
                    Dim cand = GetDefaultCandidate()
                    If cand Is Nothing Then Return DefaultRuleIndex
                    Return cand.RuleIndex
                End Get
            End Property

            Public ReadOnly Property RevitSegmentKey As String
                Get
                    Dim cand = GetDefaultCandidate()
                    Return If(cand?.SegmentKey, String.Empty)
                End Get
            End Property

            Public ReadOnly Property SegmentId As Integer
                Get
                    Dim cand = GetDefaultCandidate()
                    If cand Is Nothing Then Return 0
                    Return cand.SegmentId
                End Get
            End Property

            Public ReadOnly Property SegmentName As String
                Get
                    Dim cand = GetDefaultCandidate()
                    Return If(cand?.SegmentName, String.Empty)
                End Get
            End Property

            Private Function GetDefaultCandidate() As PrepareRow
                If Candidates Is Nothing OrElse Candidates.Count = 0 Then
                    Return Nothing
                End If
                For Each cand In Candidates
                    If cand IsNot Nothing AndAlso cand.RuleIndex = DefaultRuleIndex Then
                        Return cand
                    End If
                Next
                Return Candidates(0)
            End Function
        End Class

        Private Shared Function CellStr(row As IRow, col As Integer) As String
            If row Is Nothing OrElse col < 0 Then
                Return String.Empty
            End If
            Dim cell As ICell = row.GetCell(col)
            If cell Is Nothing Then
                Return String.Empty
            End If
            Try
                Select Case cell.CellType
                    Case NpoiCellType.String
                        Return cell.StringCellValue
                    Case NpoiCellType.Boolean
                        Return cell.BooleanCellValue.ToString(CultureInfo.InvariantCulture)
                    Case NpoiCellType.Numeric
                        Return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture)
                    Case NpoiCellType.Formula
                        If cell.CachedFormulaResultType = NpoiCellType.Numeric Then
                            Return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture)
                        End If
                        If cell.CachedFormulaResultType = NpoiCellType.String Then
                            Return cell.StringCellValue
                        End If
                    Case Else
                        Return cell.ToString()
                End Select
            Catch
                Return String.Empty
            End Try
            Return String.Empty
        End Function

        Private Shared Function CellDbl(row As IRow, col As Integer, Optional def As Double = Double.NaN) As Double
            If row Is Nothing OrElse col < 0 Then
                Return def
            End If
            Dim cell As ICell = row.GetCell(col)
            If cell Is Nothing Then
                Return def
            End If
            Try
                Select Case cell.CellType
                    Case NpoiCellType.Numeric
                        Return cell.NumericCellValue
                    Case NpoiCellType.String
                        Dim txt = cell.StringCellValue
                        Dim v As Double
                        If Double.TryParse(txt, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then
                            Return v
                        End If
                    Case NpoiCellType.Formula
                        If cell.CachedFormulaResultType = NpoiCellType.Numeric Then
                            Return cell.NumericCellValue
                        End If
                        If cell.CachedFormulaResultType = NpoiCellType.String Then
                            Dim txt = cell.StringCellValue
                            Dim v As Double
                            If Double.TryParse(txt, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then
                                Return v
                            End If
                        End If
                End Select
            Catch
            End Try
            Return def
        End Function

        Private Shared Function SafeStr(o As Object) As String
            If o Is Nothing OrElse o Is DBNull.Value Then
                Return String.Empty
            End If
            Return o.ToString().Trim()
        End Function

        Private Shared Function SafeDouble(o As Object) As Double
            If o Is Nothing OrElse o Is DBNull.Value Then
                Return 0
            End If
            Dim d As Double
            If Double.TryParse(o.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, d) Then
                Return d
            End If
            Return 0
        End Function

        Private Shared Function SafeIntObj(o As Object) As Integer
            If o Is Nothing OrElse o Is DBNull.Value Then
                Return 0
            End If
            Dim v As Integer
            If Integer.TryParse(o.ToString(), v) Then
                Return v
            End If
            Dim dbl As Double
            If Double.TryParse(o.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, dbl) Then
                Return CInt(Math.Truncate(dbl))
            End If
            Return 0
        End Function

        Private Shared Sub EnsureSchema(ds As DataSet)
            If Not ds.Tables.Contains(TableMeta) Then
                ds.Tables.Add(BuildMetaTable())
            End If
            If Not ds.Tables.Contains(TableFiles) Then
                ds.Tables.Add(BuildFileTable())
            End If
            If Not ds.Tables.Contains(TableRules) Then
                ds.Tables.Add(BuildRuleTable())
            End If
            If Not ds.Tables.Contains(TableSizes) Then
                ds.Tables.Add(BuildSizeTable())
            End If
            If Not ds.Tables.Contains(TableRouting) Then
                ds.Tables.Add(BuildRoutingTable())
            End If
        End Sub

        Private Shared Sub SaveWorkbookSafely(wb As IWorkbook, outPath As String)
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

        Private Shared Sub WriteSheet(wb As IWorkbook,
                                      name As String,
                                      t As DataTable,
                                      Optional progressChannel As String = Nothing,
                                      Optional ByRef written As Integer = 0,
                                      Optional totalRows As Integer = 0)
            If t Is Nothing Then
                Return
            End If
            Dim sh = wb.CreateSheet(name)
            Dim head = sh.CreateRow(0)
            For ci As Integer = 0 To t.Columns.Count - 1
                head.CreateCell(ci).SetCellValue(t.Columns(ci).ColumnName)
            Next
            Dim numStyle = wb.CreateCellStyle()
            numStyle.DataFormat = wb.CreateDataFormat().GetFormat("0.###############")
            Dim r As Integer = 1
            For Each row As DataRow In t.Rows
                Dim rr = sh.CreateRow(r)
                For ci As Integer = 0 To t.Columns.Count - 1
                    Dim v As Object = row(ci)
                    Dim cell = rr.CreateCell(ci)
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
                r += 1
                written += 1
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_WRITE", "엑셀 데이터 작성", written, totalRows)
            Next
        End Sub

    End Class

End Namespace
