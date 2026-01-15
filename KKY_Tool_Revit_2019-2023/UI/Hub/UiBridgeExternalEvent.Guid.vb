Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Services

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private _guidProject As DataTable = Nothing
        Private _guidFamilyDetail As DataTable = Nothing
        Private _guidMode As Integer = 1
        Private _guidRunId As String = String.Empty
        Private _guidFamilyIndex As DataTable = Nothing
        Private _guidIncludeFamily As Boolean = False
        Private _guidLastPct As Double = -1.0R
        Private _guidLastText As String = String.Empty

        Private NotInheritable Class TablePayload
            Public Property columns As List(Of String)
            Public Property rows As List(Of Object())
        End Class

        ' -----------------------------
        ' 핸들러
        ' -----------------------------
        Private Sub HandleGuidAddFiles(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim pick As String = ""
            Try
                If pd IsNot Nothing AndAlso pd.ContainsKey("pick") Then
                    pick = Convert.ToString(pd("pick"))
                End If
            Catch
                pick = ""
            End Try

            If String.Equals(pick, "folder", StringComparison.OrdinalIgnoreCase) Then
                HandleGuidAddFolder()
                Return
            End If

            Using dlg As New OpenFileDialog()
                dlg.Filter = "Revit Project (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.Title = "검토할 RVT 파일 선택"
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> DialogResult.OK Then Return

                Dim files As New List(Of String)()
                For Each p In dlg.FileNames
                    If Not String.IsNullOrWhiteSpace(p) Then files.Add(p)
                Next
                SendToWeb("guid:files", New With {.paths = files})
            End Using
        End Sub

        Private Sub HandleGuidRun(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim mode As Integer = SafeIntObj(GetProp(pd, "mode"), 1)
            If mode <> 1 AndAlso mode <> 2 Then mode = 1
            Dim rvtPaths = ParseStringList(pd, "rvtPaths")
            Dim includeFamily As Boolean = SafeBoolObj(GetProp(pd, "includeFamily"), mode = 2)
            Dim includeAnnotation As Boolean = SafeBoolObj(GetProp(pd, "includeAnnotation"), False)

            Try
                Dim sharedStatus = SharedParameterStatusService.GetStatus(app)
                If sharedStatus Is Nothing OrElse Not String.Equals(sharedStatus.Status, "ok", StringComparison.OrdinalIgnoreCase) Then
                    Dim msg = If(String.IsNullOrWhiteSpace(sharedStatus?.WarningMessage), "Shared Parameter 파일 상태가 올바르지 않습니다.", sharedStatus.WarningMessage)
                    SendToWeb("sharedparam:status", New With {
                        .path = sharedStatus?.Path,
                        .isSet = sharedStatus?.IsSet,
                        .existsOnDisk = sharedStatus?.ExistsOnDisk,
                        .canOpen = sharedStatus?.CanOpen,
                        .status = sharedStatus?.Status,
                        .statusLabel = sharedStatus?.StatusLabel,
                        .warning = sharedStatus?.WarningMessage,
                        .errorMessage = sharedStatus?.ErrorMessage
                    })
                    SendToWeb("guid:error", New With {.message = msg})
                    SendToWeb("revit:error", New With {.message = "GUID 검토 실패: " & msg})
                    Return
                End If

                _guidProject = Nothing
                _guidFamilyDetail = Nothing
                _guidFamilyIndex = Nothing
                _guidRunId = String.Empty
                _guidIncludeFamily = includeFamily
                _guidMode = mode
                _guidLastPct = -1.0R
                _guidLastText = String.Empty

                Dim res = GuidAuditService.Run(app, mode, rvtPaths, AddressOf ReportGuidProgress,
                                               Sub(msg As String)
                                                   If Not String.IsNullOrWhiteSpace(msg) Then
                                                       SendToWeb("guid:warn", New With {.message = msg})
                                                   End If
                                               End Sub,
                                               includeFamily:=includeFamily,
                                               includeAnnotation:=includeAnnotation)
                _guidProject = res.Project
                _guidFamilyDetail = res.FamilyDetail
                _guidFamilyIndex = res.FamilyIndex
                _guidRunId = res.RunId
                _guidIncludeFamily = res.IncludeFamily

                Dim payloadProject As TablePayload = ShapeTable(res.Project, New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {"RvtPath"})
                Dim payloadFamily As TablePayload = ShapeTable(res.FamilyIndex, Nothing)

                Dim donePayload = New With {
                    .mode = mode,
                    .runId = _guidRunId,
                    .includeFamily = _guidIncludeFamily,
                    .includeAnnotation = includeAnnotation,
                    .project = payloadProject,
                    .family = payloadFamily
                }

                Try
                    SendToWeb("guid:done", donePayload)
                Catch exSend As Exception
                    Dim hr As Integer = 0
                    Try : hr = exSend.HResult : Catch : hr = 0 : End Try
                    SendToWeb("guid:error", New With {
                        .message = $"guid:done 전송 실패: {exSend.Message}",
                        .hResult = hr,
                        .projectRows = If(res.Project Is Nothing, 0, res.Project.Rows.Count),
                        .familyDetailRows = If(res.FamilyDetail Is Nothing, 0, res.FamilyDetail.Rows.Count),
                        .familyIndexRows = If(res.FamilyIndex Is Nothing, 0, res.FamilyIndex.Rows.Count)
                    })
                End Try
            Catch ex As Exception
                SendToWeb("guid:error", New With {.message = ex.Message})
            Finally
                ReportGuidProgress(0, String.Empty)
            End Try
        End Sub

        Private Sub HandleGuidExport(app As UIApplication, payload As Object)
            Dim which As String = ""
            Try
                which = Convert.ToString(GetProp(payload, "which"))
            Catch
                which = ""
            End Try
            which = If(which, "").ToLowerInvariant()
            Dim excelMode As String = "fast"
            Try
                Dim em = Convert.ToString(GetProp(payload, "excelMode"))
                If Not String.IsNullOrWhiteSpace(em) Then excelMode = em
            Catch
            End Try

            Try
                Dim projectTable As DataTable = EnsureNoRvtPath(GuidAuditService.PrepareExportTable(_guidProject, 1))
                Dim familyTable As DataTable = EnsureNoRvtPath(GuidAuditService.PrepareExportTable(_guidFamilyDetail, 2))
                Dim sheetList As New List(Of KeyValuePair(Of String, DataTable))()

                If which = "family" Then
                    sheetList.Add(New KeyValuePair(Of String, DataTable)("Family 검토결과", familyTable))
                ElseIf which = "all" Then
                    sheetList.Add(New KeyValuePair(Of String, DataTable)("RVT 검토결과", projectTable))
                    sheetList.Add(New KeyValuePair(Of String, DataTable)("Family 검토결과", familyTable))
                Else
                    sheetList.Add(New KeyValuePair(Of String, DataTable)("RVT 검토결과", projectTable))
                End If

                Dim requestedAutoFit As Boolean = String.Equals(excelMode, "normal", StringComparison.OrdinalIgnoreCase)
                LogAutoFitDecision(requestedAutoFit, "GuidAuditExport")
                Dim saved = GuidAuditService.ExportMulti(sheetList, excelMode, "guid:progress")
                If String.IsNullOrWhiteSpace(saved) Then
                    SendToWeb("guid:error", New With {.message = "엑셀 내보내기가 취소되었습니다."})
                    Return
                End If
                SendToWeb("guid:exported", New With {.path = saved, .which = which})
            Catch ex As Exception
                SendToWeb("guid:error", New With {.message = "엑셀 내보내기 실패: " & ex.Message})
            End Try
        End Sub

        Private Sub HandleGuidRequestFamilyDetail(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim runId As String = ""
            Dim rvtPath As String = ""
            Dim familyName As String = ""
            Try : runId = Convert.ToString(GetProp(pd, "runId")) : Catch : runId = "" : End Try
            Try : rvtPath = Convert.ToString(GetProp(pd, "rvtPath")) : Catch : rvtPath = "" : End Try
            Try : familyName = Convert.ToString(GetProp(pd, "familyName")) : Catch : familyName = "" : End Try

            If Not _guidIncludeFamily Then
                SendToWeb("guid:error", New With {.message = "Family 검토 결과가 없습니다.", .runId = runId})
                Return
            End If

            If String.IsNullOrWhiteSpace(runId) OrElse Not String.Equals(runId, _guidRunId, StringComparison.OrdinalIgnoreCase) Then
                SendToWeb("guid:error", New With {.message = "이전 실행 결과 요청(runId mismatch)", .runId = runId})
                Return
            End If

            If _guidFamilyDetail Is Nothing OrElse _guidFamilyDetail.Rows.Count = 0 Then
                SendToWeb("guid:error", New With {.message = "가져올 패밀리 상세 결과가 없습니다.", .runId = runId})
                Return
            End If

            Dim filtered = FilterFamilyDetail(_guidFamilyDetail, rvtPath, familyName)
            Dim shaped As TablePayload = ShapeTable(filtered, New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {"RvtPath"})

            SendToWeb("guid:family-detail", New With {
                .runId = _guidRunId,
                .rvtPath = rvtPath,
                .familyName = familyName,
                .columns = shaped.columns,
                .rows = shaped.rows
            })
        End Sub

        ' -----------------------------
        ' 유틸
        ' -----------------------------
        Private Sub ReportGuidProgress(pct As Double, text As String)
            Dim changed As Boolean = (pct <> _guidLastPct) OrElse (Not String.Equals(text, _guidLastText, StringComparison.Ordinal))
            If Not changed Then Return
            SendToWeb("guid:progress", New With {.pct = pct, .text = text})
            _guidLastPct = pct
            _guidLastText = text
        End Sub

        Private Function ShapeTable(dt As DataTable, skipCols As HashSet(Of String)) As TablePayload
            If dt Is Nothing Then Return New TablePayload With {.columns = New List(Of String)(), .rows = New List(Of Object())()}

            Dim cols As New List(Of String)()
            For Each c As DataColumn In dt.Columns
                If skipCols IsNot Nothing AndAlso skipCols.Contains(c.ColumnName) Then Continue For
                cols.Add(c.ColumnName)
            Next

            Dim rows As New List(Of Object())()
            For Each r As DataRow In dt.Rows
                Dim arr(cols.Count - 1) As Object
                For i As Integer = 0 To cols.Count - 1
                    arr(i) = SafeStrGuid(r(cols(i)))
                Next
                rows.Add(arr)
            Next

            Return New TablePayload With {.columns = cols, .rows = rows}
        End Function

        Private Function CloneWithoutColumn(dt As DataTable, columnName As String) As DataTable
            If dt Is Nothing Then Return Nothing
            Dim clone As DataTable = dt.Clone()
            If clone.Columns.Contains(columnName) Then clone.Columns.Remove(columnName)
            For Each r As DataRow In dt.Rows
                Dim nr = clone.NewRow()
                For Each c As DataColumn In clone.Columns
                    nr(c.ColumnName) = r(c.ColumnName)
                Next
                clone.Rows.Add(nr)
            Next
            Return clone
        End Function

        Private Function EnsureNoRvtPath(dt As DataTable) As DataTable
            If dt Is Nothing Then Return Nothing
            If dt.Columns.Contains("RvtPath") Then
                Return CloneWithoutColumn(dt, "RvtPath")
            End If
            Return dt
        End Function

        Private Shared Function SafeStrGuid(o As Object) As String
            If o Is Nothing OrElse o Is DBNull.Value Then Return String.Empty
            Return Convert.ToString(o)
        End Function

        Private Function FilterFamilyDetail(source As DataTable, rvtPath As String, familyName As String) As DataTable
            If source Is Nothing Then Return New DataTable()
            Dim clone As DataTable = source.Clone()
            Dim path As String = If(rvtPath, "")
            Dim fam As String = If(familyName, "")
            For Each r As DataRow In source.Rows
                Dim rp As String = SafeStrGuid(r("RvtPath"))
                Dim fn As String = SafeStrGuid(r("FamilyName"))
                If String.Equals(rp, path, StringComparison.OrdinalIgnoreCase) AndAlso
                   String.Equals(fn, fam, StringComparison.OrdinalIgnoreCase) Then
                    clone.ImportRow(r)
                End If
            Next
            Return clone
        End Function

        Private Shared Function SafeBoolObj(o As Object, Optional def As Boolean = False) As Boolean
            Try
                If o Is Nothing Then Return def
                If TypeOf o Is Boolean Then Return DirectCast(o, Boolean)
                Dim s = Convert.ToString(o)
                If String.IsNullOrWhiteSpace(s) Then Return def
                s = s.Trim().ToLowerInvariant()
                If s = "true" OrElse s = "1" OrElse s = "y" OrElse s = "yes" Then Return True
                If s = "false" OrElse s = "0" OrElse s = "n" OrElse s = "no" Then Return False
            Catch
            End Try
            Return def
        End Function

        Private Sub HandleGuidAddFolder()
            Using dlg As New FolderBrowserDialog()
                dlg.Description = "RVT 폴더 선택"
                dlg.ShowNewFolderButton = False

                If dlg.ShowDialog() <> DialogResult.OK Then
                    Return
                End If

                Dim root As String = dlg.SelectedPath
                Dim files As New List(Of String)()
                Const MaxFiles As Integer = 2000

                Try
                    If Directory.Exists(root) Then
                        Dim found = Directory.EnumerateFiles(root, "*.rvt", SearchOption.TopDirectoryOnly).
                            Select(Function(p) New With {.Path = p, .Name = TryCast(Path.GetFileName(p), String)}).
                            OrderBy(Function(x) x.Name, StringComparer.OrdinalIgnoreCase).
                            ToList()

                        For Each item In found
                            Dim fp As String = item.Path
                            If String.IsNullOrWhiteSpace(fp) Then Continue For
                            If files.Count >= MaxFiles Then Exit For
                            files.Add(fp)
                        Next

                        If found.Count > MaxFiles Then
                            SendToWeb("guid:warn", New With {.message = $"RVT 파일이 {found.Count:#,0}개 있습니다. 상위 {MaxFiles:#,0}개만 추가합니다."})
                        End If
                    End If
                Catch ex As Exception
                    SendToWeb("guid:warn", New With {.message = $"폴더를 읽는 중 오류가 발생했습니다: {ex.Message}"})
                End Try

                If files.Count = 0 Then
                    SendToWeb("guid:warn", New With {.message = "선택한 폴더에 RVT 파일이 없습니다."})
                    Return
                End If

                Dim unique As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim deduped As New List(Of String)()
                For Each f In files
                    If unique.Add(f) Then deduped.Add(f)
                Next

                SendToWeb("guid:files", New With {.paths = deduped})
            End Using
        End Sub

    End Class

End Namespace
