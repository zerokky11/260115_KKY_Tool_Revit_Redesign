Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Windows.Forms
Imports KKY_Tool_Revit.UI.Hub
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports NPOI.XSSF.UserModel
Imports NPOI.HSSF.Util

Namespace Infrastructure

    Public Module ExcelCore

        ' 저장 대화상자 + 저장
        Public Function PickAndSaveXlsx(sheetName As String, table As DataTable, Optional defaultFileName As String = Nothing, Optional doAutoFit As Boolean = False, Optional progressChannel As String = Nothing) As String
            If table Is Nothing Then Return String.Empty
            Dim fileName As String = If(String.IsNullOrWhiteSpace(defaultFileName), $"{sheetName}.xlsx", defaultFileName)
            Using sfd As New SaveFileDialog()
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                sfd.FileName = fileName
                sfd.AddExtension = True
                sfd.DefaultExt = "xlsx"
                sfd.OverwritePrompt = True
                sfd.RestoreDirectory = True
                If sfd.ShowDialog() = DialogResult.OK Then
                    SaveXlsx(sfd.FileName, sheetName, table, doAutoFit, progressChannel)
                    Return sfd.FileName
                End If
            End Using
            Return String.Empty
        End Function

        ' 저장 대화상자 + 다중 시트 저장
        Public Function PickAndSaveXlsxMulti(sheets As IList(Of KeyValuePair(Of String, DataTable)), Optional defaultFileName As String = Nothing, Optional doAutoFit As Boolean = False, Optional progressChannel As String = Nothing) As String
            If sheets Is Nothing OrElse sheets.Count = 0 Then Return String.Empty
            Dim fileName As String = If(String.IsNullOrWhiteSpace(defaultFileName), "Results.xlsx", defaultFileName)
            Using sfd As New SaveFileDialog()
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                sfd.FileName = fileName
                sfd.AddExtension = True
                sfd.DefaultExt = "xlsx"
                sfd.OverwritePrompt = True
                sfd.RestoreDirectory = True
                If sfd.ShowDialog() = DialogResult.OK Then
                    SaveXlsxMulti(sfd.FileName, sheets, doAutoFit, progressChannel)
                    Return sfd.FileName
                End If
            End Using
            Return String.Empty
        End Function

        ' 일반 테이블 저장
        Public Sub SaveXlsx(filePath As String, sheetName As String, table As DataTable, Optional doAutoFit As Boolean = False, Optional progressChannel As String = Nothing)
            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))
            Dim totalRows As Integer = table.Rows.Count
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset(progressChannel)
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_INIT", "엑셀 워크북 준비", 0, totalRows, Nothing, True)
            UI.Hub.UiBridgeExternalEvent.LogAutoFitDecision(doAutoFit, "ExcelCore.SaveXlsx")
            Dim wb As IWorkbook = New XSSFWorkbook()

            Dim headFont = wb.CreateFont() : headFont.IsBold = True
            Dim headStyle = wb.CreateCellStyle()
            headStyle.SetFont(headFont)
            headStyle.FillPattern = FillPattern.SolidForeground
            headStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index
            SetThinBorders(headStyle)

            Dim bodyStyle = wb.CreateCellStyle() : SetThinBorders(bodyStyle)

            Dim sh = wb.CreateSheet(SafeSheetName(If(sheetName, "Sheet1")))

            ' 헤더
            Dim r0 = sh.CreateRow(0)
            For ci = 0 To table.Columns.Count - 1
                Dim c = r0.CreateCell(ci)
                c.SetCellValue(table.Columns(ci).ColumnName)
                c.CellStyle = headStyle
            Next

            If table.Columns.Count > 0 Then
                Dim lastCol As Integer = table.Columns.Count - 1
                Dim range As New CellRangeAddress(0, 0, 0, lastCol)
                sh.SetAutoFilter(range)
            End If
            sh.CreateFreezePane(0, 1)

            ' 바디
            For ri = 0 To table.Rows.Count - 1
                Dim rr = sh.CreateRow(ri + 1)
                For ci = 0 To table.Columns.Count - 1
                    Dim cc = rr.CreateCell(ci)
                    Dim v = If(table.Rows(ri)(ci), "").ToString()
                    cc.SetCellValue(v)
                    cc.CellStyle = bodyStyle
                Next
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_WRITE", "엑셀 데이터 작성", ri + 1, totalRows)
            Next

            If doAutoFit Then
                AutoSizeAll(sh, table.Columns.Count)
            End If

            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_SAVE", "파일 저장 중", totalRows, totalRows, Nothing, True)
            SaveWorkbookToFile(wb, filePath)

            Dim autoFitMessage As String = If(doAutoFit, "열 너비 자동 조정 중…", "빠른 모드: 열 너비 자동 조정 생략")
            If doAutoFit Then
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", autoFitMessage, totalRows, totalRows, Nothing, True)
                Global.KKY_Tool_Revit.Infrastructure.ExcelCore.TryAutoFitWithExcel(filePath)
            Else
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", autoFitMessage, totalRows, totalRows, Nothing, True)
            End If

            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "DONE", "엑셀 내보내기 완료", totalRows, totalRows, 100.0R, True)
            wb.Close()
        End Sub

        ' 다중 시트 저장
        Public Sub SaveXlsxMulti(filePath As String, sheets As IList(Of KeyValuePair(Of String, DataTable)), Optional doAutoFit As Boolean = False, Optional progressChannel As String = Nothing)
            If sheets Is Nothing OrElse sheets.Count = 0 Then Throw New ArgumentNullException(NameOf(sheets))
            Dim totalRows As Integer = 0
            For Each kv In sheets
                If kv.Value Is Nothing Then Continue For
                totalRows += kv.Value.Rows.Count
            Next

            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset(progressChannel)
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_INIT", "엑셀 워크북 준비", 0, totalRows, Nothing, True)
            UI.Hub.UiBridgeExternalEvent.LogAutoFitDecision(doAutoFit, "ExcelCore.SaveXlsxMulti")

            Dim wb As IWorkbook = New XSSFWorkbook()
            Dim headFont = wb.CreateFont() : headFont.IsBold = True
            Dim headStyle = wb.CreateCellStyle()
            headStyle.SetFont(headFont)
            headStyle.FillPattern = FillPattern.SolidForeground
            headStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index
            SetThinBorders(headStyle)

            Dim bodyStyle = wb.CreateCellStyle() : SetThinBorders(bodyStyle)
            Dim current As Integer = 0

            For Each kv In sheets
                Dim table = kv.Value
                If table Is Nothing Then Continue For
                Dim sh = wb.CreateSheet(SafeSheetName(If(kv.Key, "Sheet1")))

                Dim r0 = sh.CreateRow(0)
                For ci = 0 To table.Columns.Count - 1
                    Dim c = r0.CreateCell(ci)
                    c.SetCellValue(table.Columns(ci).ColumnName)
                    c.CellStyle = headStyle
                Next

                If table.Columns.Count > 0 Then
                    Dim lastCol As Integer = table.Columns.Count - 1
                    Dim range As New CellRangeAddress(0, 0, 0, lastCol)
                    sh.SetAutoFilter(range)
                End If
                sh.CreateFreezePane(0, 1)

                For ri = 0 To table.Rows.Count - 1
                    Dim rr = sh.CreateRow(ri + 1)
                    For ci = 0 To table.Columns.Count - 1
                        Dim cc = rr.CreateCell(ci)
                        Dim v = If(table.Rows(ri)(ci), "").ToString()
                        cc.SetCellValue(v)
                        cc.CellStyle = bodyStyle
                    Next
                    current += 1
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_WRITE", "엑셀 데이터 작성", current, totalRows)
                Next

                If doAutoFit Then
                    AutoSizeAll(sh, table.Columns.Count)
                End If
            Next

            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_SAVE", "파일 저장 중", totalRows, totalRows, Nothing, True)
            SaveWorkbookToFile(wb, filePath)

            Dim autoFitMessage As String = If(doAutoFit, "열 너비 자동 조정 중…", "빠른 모드: 열 너비 자동 조정 생략")
            If doAutoFit Then
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", autoFitMessage, totalRows, totalRows, Nothing, True)
                Global.KKY_Tool_Revit.Infrastructure.ExcelCore.TryAutoFitWithExcel(filePath)
            Else
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", autoFitMessage, totalRows, totalRows, Nothing, True)
            End If

            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "DONE", "엑셀 내보내기 완료", totalRows, totalRows, 100.0R, True)
            wb.Close()
        End Sub

        ' 요약 테이블(그룹 밴딩 + 그룹 글꼴 강조)
        Public Sub SaveStyledSimple(outPath As String, sheetName As String, table As DataTable, groupColumnName As String, Optional doAutoFit As Boolean = False, Optional progressChannel As String = Nothing)
            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))
            Dim totalRows As Integer = table.Rows.Count
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset(progressChannel)
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_INIT", "엑셀 워크북 준비", 0, totalRows, Nothing, True)
            UI.Hub.UiBridgeExternalEvent.LogAutoFitDecision(doAutoFit, "ExcelCore.SaveStyledSimple")
            Dim wb As IWorkbook = New XSSFWorkbook()

            Dim headFont = wb.CreateFont() : headFont.IsBold = True
            Dim headStyle = wb.CreateCellStyle()
            headStyle.SetFont(headFont)
            headStyle.FillPattern = FillPattern.SolidForeground
            headStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index
            SetThinBorders(headStyle)

            Dim bodyA = wb.CreateCellStyle() : SetThinBorders(bodyA)
            bodyA.FillPattern = FillPattern.SolidForeground
            bodyA.FillForegroundColor = IndexedColors.PaleBlue.Index

            Dim bodyB = wb.CreateCellStyle() : SetThinBorders(bodyB)
            bodyB.FillPattern = FillPattern.SolidForeground
            bodyB.FillForegroundColor = IndexedColors.LightCornflowerBlue.Index

            Dim bodyPlain = wb.CreateCellStyle() : SetThinBorders(bodyPlain)

            Dim grpFont = wb.CreateFont() : grpFont.IsBold = True : grpFont.Color = IndexedColors.RoyalBlue.Index

            Dim sh = wb.CreateSheet(SafeSheetName(If(sheetName, "Sheet1")))

            ' 헤더
            Dim r0 = sh.CreateRow(0)
            For ci = 0 To table.Columns.Count - 1
                Dim c = r0.CreateCell(ci)
                c.SetCellValue(table.Columns(ci).ColumnName)
                c.CellStyle = headStyle
            Next

            If table.Columns.Count > 0 Then
                Dim lastCol As Integer = table.Columns.Count - 1
                Dim range As New CellRangeAddress(0, 0, 0, lastCol)
                sh.SetAutoFilter(range)
            End If

            Dim grpIx As Integer = -1
            If Not String.IsNullOrWhiteSpace(groupColumnName) AndAlso table.Columns.Contains(groupColumnName) Then
                grpIx = table.Columns(groupColumnName).Ordinal
            End If

            Dim lastGroup As String = Nothing
            Dim flagA As Boolean = True

            For ri = 0 To table.Rows.Count - 1
                Dim rr = sh.CreateRow(ri + 1)
                Dim curGroup As String = Nothing

                If grpIx >= 0 Then
                    curGroup = If(table.Rows(ri)(grpIx), "").ToString()
                    If lastGroup Is Nothing OrElse Not lastGroup.Equals(curGroup, StringComparison.Ordinal) Then
                        flagA = Not flagA
                        lastGroup = curGroup
                    End If
                End If

                For ci = 0 To table.Columns.Count - 1
                    Dim cc = rr.CreateCell(ci)
                    Dim v = If(table.Rows(ri)(ci), "").ToString()
                    cc.SetCellValue(v)

                    Dim st = If(grpIx >= 0, If(flagA, bodyA, bodyB), bodyPlain)
                    If grpIx = ci AndAlso grpIx >= 0 Then
                        Dim s2 = wb.CreateCellStyle()
                        s2.CloneStyleFrom(st)
                        s2.SetFont(grpFont)
                        cc.CellStyle = s2
                    Else
                        cc.CellStyle = st
                    End If
                Next
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_WRITE", "엑셀 데이터 작성", ri + 1, totalRows)
            Next

            If doAutoFit Then
                AutoSizeAll(sh, table.Columns.Count)
            End If

            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_SAVE", "파일 저장 중", totalRows, totalRows, Nothing, True)
            SaveWorkbookToFile(wb, outPath)

            Dim autoFitMessage As String = If(doAutoFit, "열 너비 자동 조정 중…", "빠른 모드: 열 너비 자동 조정 생략")
            If doAutoFit Then
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", autoFitMessage, totalRows, totalRows, Nothing, True)
                Global.KKY_Tool_Revit.Infrastructure.ExcelCore.TryAutoFitWithExcel(outPath)
            Else
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", autoFitMessage, totalRows, totalRows, Nothing, True)
            End If

            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "DONE", "엑셀 내보내기 완료", totalRows, totalRows, 100.0R, True)
            wb.Close()
        End Sub

        Private Sub SaveWorkbookToFile(wb As IWorkbook, outPath As String)
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

        ' 공통: 얇은 테두리 + 자동열너비
        Private Sub SetThinBorders(st As ICellStyle)
            st.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin
            st.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin
            st.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin
            st.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin
        End Sub

        Private Sub AutoSizeAll(sh As ISheet, colCount As Integer)
            If sh Is Nothing Then
                Return
            End If
            If sh.LastRowNum > 10000 Then
                Return
            End If
            For ci = 0 To colCount - 1
                sh.AutoSizeColumn(ci, False)
                Dim cur = sh.GetColumnWidth(ci)
                sh.SetColumnWidth(ci, Math.Min(cur + 512, 255 * 256))
            Next
        End Sub

        Private Function SafeSheetName(name As String) As String
            Dim bad = New Char() {"/"c, "\"c, "?"c, "*"c, "["c, "]"c, ":"c}
            Dim s = name
            For Each ch In bad : s = s.Replace(ch, "-"c) : Next
            If s.Length > 31 Then s = s.Substring(0, 31)
            If String.IsNullOrWhiteSpace(s) Then s = "Sheet1"
            Return s
        End Function

        ' Excel COM AutoFit (선택적)
        Public Sub TryAutoFitWithExcel(filePath As String)
            If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then Return
            Dim excelApp As Object = Nothing
            Dim wb As Object = Nothing
            Try
                excelApp = CreateObject("Excel.Application")
                If excelApp Is Nothing Then Return
                excelApp.DisplayAlerts = False
                wb = excelApp.Workbooks.Open(filePath)
                If wb Is Nothing Then Return
                For Each sh In wb.Worksheets
                    Try
                        sh.Cells.EntireColumn.AutoFit()
                    Catch
                    End Try
                Next
                wb.Save()
            Catch
            Finally
                Try
                    If wb IsNot Nothing Then wb.Close(False)
                Catch
                End Try
                Try
                    If excelApp IsNot Nothing Then excelApp.Quit()
                Catch
                End Try
                TryReleaseCom(wb)
                TryReleaseCom(excelApp)
            End Try
        End Sub

        Private Sub TryReleaseCom(o As Object)
            Try
                If o IsNot Nothing AndAlso Type.GetTypeFromProgID("Excel.Application") IsNot Nothing Then
                    Dim t = GetType(System.Runtime.InteropServices.Marshal)
                    Dim rel = t.GetMethod("ReleaseComObject", {GetType(Object)})
                    If rel IsNot Nothing Then
                        rel.Invoke(Nothing, New Object() {o})
                    End If
                End If
            Catch
            End Try
        End Sub

        ' ----------------------------
        ' 공통 시트 스타일 적용 유틸
        ' ----------------------------
        Public Sub ApplyStandardSheetStyle(wb As IWorkbook,
                                           sheet As ISheet,
                                           Optional headerRowIndex As Integer = 0,
                                           Optional autoFilter As Boolean = True,
                                           Optional freezeTopRow As Boolean = True,
                                           Optional borderAll As Boolean = True,
                                           Optional autoFit As Boolean = True,
                                           Optional headerFillColor As Short = -1S)
            If wb Is Nothing OrElse sheet Is Nothing Then
                Return
            End If

            Dim headerRow = sheet.GetRow(headerRowIndex)
            If headerRow Is Nothing Then
                Return
            End If

            Dim lastRow As Integer = sheet.LastRowNum
            Dim lastCol As Integer = headerRow.LastCellNum - 1
            If lastCol < 0 Then
                Return
            End If

            ' 헤더 스타일
            Dim headFont = wb.CreateFont()
            headFont.IsBold = True
            Dim headStyle = wb.CreateCellStyle()
            headStyle.SetFont(headFont)
            Dim resolvedHeaderFill As Short = If(headerFillColor < 0S, IndexedColors.Grey25Percent.Index, headerFillColor)
            headStyle.FillPattern = FillPattern.SolidForeground
            headStyle.FillForegroundColor = resolvedHeaderFill
            headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left
            SetThinBorders(headStyle)

            For ci As Integer = 0 To lastCol
                Dim c = headerRow.GetCell(ci)
                If c Is Nothing Then
                    c = headerRow.CreateCell(ci)
                End If
                c.CellStyle = headStyle
            Next

            If autoFilter Then
                Dim range As New CellRangeAddress(headerRowIndex, headerRowIndex, 0, lastCol)
                sheet.SetAutoFilter(range)
            End If

            If freezeTopRow Then
                sheet.CreateFreezePane(0, headerRowIndex + 1)
            End If

            If borderAll Then
                Dim styleCache As New Dictionary(Of Short, ICellStyle)()
                For r As Integer = headerRowIndex + 1 To lastRow
                    Dim row = sheet.GetRow(r)
                    If row Is Nothing Then
                        Continue For
                    End If
                    Dim rowLastCol As Integer = row.LastCellNum - 1
                    If rowLastCol < 0 Then
                        Continue For
                    End If
                    For ci As Integer = 0 To rowLastCol
                        Dim cell = row.GetCell(ci)
                        If cell Is Nothing Then
                            Continue For
                        End If
                        Dim srcStyle As ICellStyle = cell.CellStyle
                        Dim key As Short = If(srcStyle Is Nothing, -1S, srcStyle.Index)
                        Dim styled As ICellStyle = Nothing
                        If key >= 0 AndAlso styleCache.TryGetValue(key, styled) Then
                            cell.CellStyle = styled
                        Else
                            Dim newStyle = wb.CreateCellStyle()
                            If srcStyle IsNot Nothing Then
                                newStyle.CloneStyleFrom(srcStyle)
                            End If
                            SetThinBorders(newStyle)
                            If key >= 0 Then
                                styleCache(key) = newStyle
                            End If
                            cell.CellStyle = newStyle
                        End If
                    Next
                Next
            End If

            If autoFit AndAlso lastRow <= 10000 Then
                For ci As Integer = 0 To lastCol
                    sheet.AutoSizeColumn(ci, False)
                    Dim cur = sheet.GetColumnWidth(ci)
                    Dim padded = Math.Min(cur + 512, 255 * 256)
                    sheet.SetColumnWidth(ci, padded)
                Next
            End If
        End Sub

        Public Sub ApplyNumberFormatByHeader(wb As IWorkbook,
                                             sheet As ISheet,
                                             headerRowIndex As Integer,
                                             headers As IEnumerable(Of String),
                                             format As String)
            If wb Is Nothing OrElse sheet Is Nothing OrElse headers Is Nothing OrElse String.IsNullOrWhiteSpace(format) Then
                Return
            End If
            Dim headerRow = sheet.GetRow(headerRowIndex)
            If headerRow Is Nothing Then
                Return
            End If
            Dim targetSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each h In headers
                If Not String.IsNullOrWhiteSpace(h) Then
                    targetSet.Add(h.Trim())
                End If
            Next
            If targetSet.Count = 0 Then
                Return
            End If
            Dim cols As New List(Of Integer)()
            For ci As Integer = 0 To headerRow.LastCellNum - 1
                Dim name = headerRow.GetCell(ci)?.StringCellValue
                If Not String.IsNullOrWhiteSpace(name) AndAlso targetSet.Contains(name.Trim()) Then
                    cols.Add(ci)
                End If
            Next
            If cols.Count = 0 Then
                Return
            End If
            Dim df As Short = wb.CreateDataFormat().GetFormat(format)
            Dim styleCache As New Dictionary(Of Short, ICellStyle)()
            For r As Integer = headerRowIndex + 1 To sheet.LastRowNum
                Dim row = sheet.GetRow(r)
                If row Is Nothing Then
                    Continue For
                End If
                For Each ci In cols
                    Dim cell = row.GetCell(ci)
                    If cell Is Nothing Then
                        Continue For
                    End If
                    Dim srcStyle As ICellStyle = cell.CellStyle
                    Dim key As Short = If(srcStyle Is Nothing, -1S, srcStyle.Index)
                    Dim cached As ICellStyle = Nothing
                    If key >= 0 AndAlso styleCache.TryGetValue(key, cached) Then
                        cell.CellStyle = cached
                    Else
                        Dim newStyle = wb.CreateCellStyle()
                        If srcStyle IsNot Nothing Then
                            newStyle.CloneStyleFrom(srcStyle)
                        End If
                        newStyle.DataFormat = df
                        If key >= 0 Then
                            styleCache(key) = newStyle
                        End If
                        cell.CellStyle = newStyle
                    End If
                Next
            Next
        End Sub

        Public Sub ApplyResultFillByHeader(wb As IWorkbook,
                                           sheet As ISheet,
                                           headerRowIndex As Integer)
            If wb Is Nothing OrElse sheet Is Nothing Then
                Return
            End If
            Dim headerRow = sheet.GetRow(headerRowIndex)
            If headerRow Is Nothing Then
                Return
            End If
            Dim targetCol As Integer = -1
            For ci As Integer = 0 To headerRow.LastCellNum - 1
                Dim name = headerRow.GetCell(ci)?.StringCellValue
                If String.IsNullOrWhiteSpace(name) Then
                    Continue For
                End If
                Dim lower = name.ToLowerInvariant()
                If lower.Contains("result") OrElse lower.Contains("status") OrElse lower.Contains("검토") Then
                    targetCol = ci
                    Exit For
                End If
            Next
            If targetCol < 0 Then
                Return
            End If

            Dim okStyle = CreateFillStyle(wb, IndexedColors.LightGreen.Index)
            Dim mismatchStyle = CreateFillStyle(wb, IndexedColors.Rose.Index)
            Dim missingStyle = CreateFillStyle(wb, IndexedColors.LightYellow.Index)

            For r As Integer = headerRowIndex + 1 To sheet.LastRowNum
                Dim row = sheet.GetRow(r)
                If row Is Nothing Then Continue For
                Dim cell = row.GetCell(targetCol)
                If cell Is Nothing Then Continue For
                Dim txt = cell.ToString()
                If String.IsNullOrWhiteSpace(txt) Then
                    cell.CellStyle = MergeStyles(cell.CellStyle, missingStyle, wb)
                ElseIf txt.IndexOf("Mismatch", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    cell.CellStyle = MergeStyles(cell.CellStyle, mismatchStyle, wb)
                ElseIf txt.IndexOf("Missing", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    cell.CellStyle = MergeStyles(cell.CellStyle, missingStyle, wb)
                ElseIf txt.IndexOf("OK", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    cell.CellStyle = MergeStyles(cell.CellStyle, okStyle, wb)
                End If
            Next
        End Sub

        Private Function CreateFillStyle(wb As IWorkbook, colorIndex As Short) As ICellStyle
            Dim st = wb.CreateCellStyle()
            st.FillPattern = FillPattern.SolidForeground
            st.FillForegroundColor = colorIndex
            Return st
        End Function

        Private Function MergeStyles(baseStyle As ICellStyle, fillStyle As ICellStyle, wb As IWorkbook) As ICellStyle
            Dim newStyle = wb.CreateCellStyle()
            If baseStyle IsNot Nothing Then
                newStyle.CloneStyleFrom(baseStyle)
            End If
            If fillStyle IsNot Nothing Then
                newStyle.FillPattern = fillStyle.FillPattern
                newStyle.FillForegroundColor = fillStyle.FillForegroundColor
            End If
            Return newStyle
        End Function

    End Module
End Namespace
