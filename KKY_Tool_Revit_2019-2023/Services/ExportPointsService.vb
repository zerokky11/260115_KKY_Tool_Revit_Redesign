Option Explicit On
Option Strict On

Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Globalization
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Infrastructure ' ExcelCore 사용




Namespace Services

    Public Class ExportPointsService

        Public Class Row
            Public Property [File] As String
            Public Property ProjectE As Double
            Public Property ProjectN As Double
            Public Property ProjectZ As Double
            Public Property SurveyE As Double
            Public Property SurveyN As Double
            Public Property SurveyZ As Double
            Public Property TrueNorth As Double
        End Class

        Public Class ProgressInfo
            Public Property Phase As String
            Public Property Message As String
            Public Property Current As Integer
            Public Property Total As Integer
            Public Property PhaseProgress As Double
        End Class

        Public Shared Function Run(uiapp As UIApplication, files As Object, Optional progress As Action(Of ProgressInfo) = Nothing) As IList(Of Row)
            Dim app = uiapp.Application
            Dim list As New List(Of Row)()

            Dim paths As New List(Of String)()
            If TypeOf files Is IEnumerable(Of Object) Then
                For Each o In CType(files, IEnumerable(Of Object))
                    Dim s = TryCast(o, String)
                    If Not String.IsNullOrWhiteSpace(s) AndAlso File.Exists(s) Then paths.Add(s)
                Next
            ElseIf TypeOf files Is String AndAlso File.Exists(CStr(files)) Then
                paths.Add(CStr(files))
            End If
            paths = paths.Distinct().ToList()

            Dim total As Integer = paths.Count
            ReportProgress(progress, "COLLECT", "파일 목록 준비 중", 0, total, 0.0)

            If paths.Count = 0 Then
                ReportProgress(progress, "DONE", "대상 파일이 없습니다.", 0, 0, 1.0)
                Return list
            End If

            For i As Integer = 0 To paths.Count - 1
                Dim p As String = paths(i)
                Dim stageProgress As Double = If(total > 0, CDbl(i) / CDbl(total), 0.0)
                ReportProgress(progress, "EXTRACT", $"파일 열기: {Path.GetFileName(p)}", i, total, stageProgress)

                Dim doc As Document = Nothing
                Try
                    Dim opt As OpenOptions = BuildOpenOptions(p)
                    Dim mp As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(p)
                    doc = app.OpenDocumentFile(mp, opt)

                    Dim row As New Row() With {.File = Path.GetFileName(p)}
                    Extract(doc, row)
                    list.Add(row)
                    Dim afterProgress As Double = If(total > 0, CDbl(i + 1) / CDbl(total), 1.0)
                    ReportProgress(progress, "EXTRACT", $"포인트 추출: {Path.GetFileName(p)}", i + 1, total, afterProgress)
                Catch ex As Exception
                    Dim afterProgress As Double = If(total > 0, CDbl(i + 1) / CDbl(total), 1.0)
                    ReportProgress(progress, "EXTRACT", $"오류로 건너뜀: {Path.GetFileName(p)}", i + 1, total, afterProgress)
                    ' 개별 파일 실패는 무시하고 다음으로 진행
                Finally
                    If doc IsNot Nothing Then
                        Try
                            doc.Close(False)
                        Catch
                        End Try
                    End If
                End Try
            Next
            ReportProgress(progress, "DONE", "포인트 추출 완료", total, total, 1.0)
            Return list
        End Function

        Public Shared Function RunOnDocument(doc As Document, fileName As String, Optional progress As Action(Of ProgressInfo) = Nothing) As IList(Of Row)
            Dim list As New List(Of Row)()
            If doc Is Nothing Then Return list
            Dim row As New Row() With {.File = If(String.IsNullOrWhiteSpace(fileName), doc.Title, fileName)}
            Try
                ReportProgress(progress, "EXTRACT", $"포인트 추출: {row.File}", 0, 1, 0.0)
                Extract(doc, row)
                list.Add(row)
                ReportProgress(progress, "DONE", "포인트 추출 완료", 1, 1, 1.0)
            Catch
                ReportProgress(progress, "ERROR", "포인트 추출 실패", 0, 1, 1.0)
            End Try
            Return list
        End Function

        Public Shared Function ExportToExcel(uiapp As UIApplication, files As Object, Optional unit As String = "ft", Optional doAutoFit As Boolean = False) As String
            Dim rows = Run(uiapp, files)

            Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Dim outPath As String = Path.Combine(desktop, $"ExportPoints_{Date.Now:yyyyMMdd_HHmmss}.xlsx")

            Dim normalizedUnit As String = NormalizeUnit(unit)
            Dim headers = BuildHeaders(normalizedUnit)
            Dim data = rows.Select(Function(r) New Object() {
                r.File,
                RoundCoord(ToUnitValue(r.ProjectE, normalizedUnit)),
                RoundCoord(ToUnitValue(r.ProjectN, normalizedUnit)),
                RoundCoord(ToUnitValue(r.ProjectZ, normalizedUnit)),
                RoundCoord(ToUnitValue(r.SurveyE, normalizedUnit)),
                RoundCoord(ToUnitValue(r.SurveyN, normalizedUnit)),
                RoundCoord(ToUnitValue(r.SurveyZ, normalizedUnit)),
                Math.Round(r.TrueNorth, 3)
            })

            Dim dt As DataTable = BuildTable(headers, data)
            ExcelCore.SaveXlsx(outPath, "Points", dt, doAutoFit)

            Return outPath
        End Function

        ' ───────────────────────── 내부 유틸 ─────────────────────────

        ''' <summary>
        ''' 헤더/데이터(열 배열) 시퀀스를 DataTable로 변환
        ''' </summary>
        Private Shared Function BuildTable(headers As IEnumerable(Of String),
                                           rows As IEnumerable(Of IEnumerable(Of Object))) As DataTable
            Dim dt As New DataTable("ExportedPoints")

            ' 헤더 보정 및 컬럼 생성
            Dim headArr = If(headers, Enumerable.Empty(Of String)()) _
                          .Select(Function(h, i) If(String.IsNullOrWhiteSpace(h), $"Col{i + 1}", h.Trim())) _
                          .ToArray()
            For Each h In headArr
                dt.Columns.Add(h)
            Next

            ' 데이터 행 추가
            If rows IsNot Nothing Then
                For Each r In rows
                    Dim vals = If(r, Enumerable.Empty(Of Object)()).ToArray()
                    Dim dr = dt.NewRow()
                    For i = 0 To Math.Min(vals.Length, dt.Columns.Count) - 1
                        dr(i) = If(vals(i), String.Empty).ToString()
                    Next
                    dt.Rows.Add(dr)
                Next
            End If

            Return dt
        End Function

        Private Shared Sub Extract(doc As Document, row As Row)
            Dim basePt As BasePoint = New FilteredElementCollector(doc).
                OfClass(GetType(BasePoint)).
                Cast(Of BasePoint)().
                FirstOrDefault(Function(bp) bp.IsShared = False)

            Dim surveyPt As BasePoint = New FilteredElementCollector(doc).
                OfClass(GetType(BasePoint)).
                Cast(Of BasePoint)().
                FirstOrDefault(Function(bp) bp.IsShared = True)

            Dim project As XYZ = If(basePt IsNot Nothing, basePt.Position, XYZ.Zero)
            Dim survey As XYZ = If(surveyPt IsNot Nothing, surveyPt.Position, XYZ.Zero)

            ' 내부(ft) 값을 그대로 유지하여 단위 변환을 나중에 적용
            row.ProjectE = TryGetParamDouble(basePt, BuiltInParameter.BASEPOINT_EASTWEST_PARAM, project.X)
            row.ProjectN = TryGetParamDouble(basePt, BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM, project.Y)
            row.ProjectZ = TryGetParamDouble(basePt, BuiltInParameter.BASEPOINT_ELEVATION_PARAM, project.Z)

            row.SurveyE = TryGetParamDouble(surveyPt, BuiltInParameter.BASEPOINT_EASTWEST_PARAM, survey.X)
            row.SurveyN = TryGetParamDouble(surveyPt, BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM, survey.Y)
            row.SurveyZ = TryGetParamDouble(surveyPt, BuiltInParameter.BASEPOINT_ELEVATION_PARAM, survey.Z)

            row.TrueNorth = GetTrueNorthDegrees(doc, basePt)
        End Sub

        Private Shared Function GetTrueNorthDegrees(doc As Document, basePt As BasePoint) As Double
            Dim deg As Double
            If TryGetBasePointAngle(basePt, deg) Then
                Return NormalizeAngleDegrees(deg)
            End If
            If TryGetProjectLocationAngle(doc, deg) Then
                Return NormalizeAngleDegrees(deg)
            End If
            Return 0.0
        End Function

        Private Shared Function TryGetBasePointAngle(basePt As BasePoint, ByRef deg As Double) As Boolean
            If basePt Is Nothing Then Return False
            Try
                Dim p = basePt.Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM)
                If p IsNot Nothing Then
                    deg = p.AsDouble() * (180.0 / Math.PI)
                    Return True
                End If
            Catch
            End Try
            Try
                Dim p = basePt.LookupParameter("Angle to True North")
                If p IsNot Nothing Then
                    deg = p.AsDouble() * (180.0 / Math.PI)
                    Return True
                End If
            Catch
            End Try
            Return False
        End Function

        Private Shared Function TryGetProjectLocationAngle(doc As Document, ByRef deg As Double) As Boolean
            Try
                Dim pl As ProjectLocation = doc.ActiveProjectLocation
                If pl Is Nothing Then Return False
                Dim pp As ProjectPosition = pl.GetProjectPosition(XYZ.Zero)
                If pp IsNot Nothing Then
                    deg = pp.Angle * (180.0 / Math.PI)
                    Return True
                End If
            Catch
            End Try
            Return False
        End Function

        Private Shared Function NormalizeAngleDegrees(deg As Double) As Double
            If Double.IsNaN(deg) OrElse Double.IsInfinity(deg) Then Return 0.0
            Dim v As Double = deg Mod 360.0
            If v < 0.0 Then v += 360.0
            Return v
        End Function

        Private Shared Function TryGetParamDouble(el As Element, bip As BuiltInParameter, fallback As Double) As Double
            If el Is Nothing Then Return fallback
            Try
                Dim p = el.Parameter(bip)
                If p IsNot Nothing Then
                    Return p.AsDouble()
                End If
            Catch
            End Try
            Return fallback
        End Function

        Private Shared Function BuildOpenOptions(path As String) As OpenOptions
            Dim info As BasicFileInfo = Nothing
            Try
                info = BasicFileInfo.Extract(path)
            Catch
            End Try

            Dim ws As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
            Dim opt As New OpenOptions() With {
                .Audit = False,
                .DetachFromCentralOption = If(info IsNot Nothing AndAlso info.IsCentral, DetachFromCentralOption.DetachAndPreserveWorksets, DetachFromCentralOption.DoNotDetach)
            }
            opt.SetOpenWorksetsConfiguration(ws)
            Return opt
        End Function

        Private Shared Function NormalizeUnit(unit As String) As String
            Dim u As String = If(unit, "").Trim().ToLowerInvariant()
            If u = "m" OrElse u = "meter" OrElse u = "meters" Then Return "m"
            If u = "mm" OrElse u = "millimeter" OrElse u = "millimeters" Then Return "mm"
            Return "ft"
        End Function

        Private Shared Function BuildHeaders(unit As String) As String()
            Dim suffix As String = "(ft)"
            If unit = "m" Then
                suffix = "(m)"
            ElseIf unit = "mm" Then
                suffix = "(mm)"
            End If
            Return New String() {
                "File",
                $"ProjectPoint_E{suffix}", $"ProjectPoint_N{suffix}", $"ProjectPoint_Z{suffix}",
                $"SurveyPoint_E{suffix}", $"SurveyPoint_N{suffix}", $"SurveyPoint_Z{suffix}",
                "TrueNorthAngle(deg)"
            }
        End Function

        Private Shared Function ToUnitValue(valueFt As Double, unit As String) As Double
            If unit = "m" Then Return valueFt * 0.3048
            If unit = "mm" Then Return valueFt * 304.8
            Return valueFt
        End Function

        Private Shared Function RoundCoord(v As Double) As Double
            Return Math.Round(v, 4)
        End Function

        Private Shared Sub ReportProgress(cb As Action(Of ProgressInfo),
                                          phase As String,
                                          message As String,
                                          current As Integer,
                                          total As Integer,
                                          phaseProgress As Double)
            If cb Is Nothing Then Return
            Try
                cb(New ProgressInfo() With {
                    .Phase = phase,
                    .Message = message,
                    .Current = current,
                    .Total = total,
                    .PhaseProgress = Clamp01(phaseProgress)
                })
            Catch
            End Try
        End Sub

        Private Shared Function Clamp01(v As Double) As Double
            If Double.IsNaN(v) OrElse Double.IsInfinity(v) Then Return 0.0
            If v < 0.0 Then Return 0.0
            If v > 1.0 Then Return 1.0
            Return v
        End Function

    End Class

End Namespace
