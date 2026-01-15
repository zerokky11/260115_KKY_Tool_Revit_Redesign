Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Exports
Imports KKY_Tool_Revit.Infrastructure

' ✅ WPF 팝업(별칭 import로 Color 충돌 회피)
Imports WPF = System.Windows
Imports WControls = System.Windows.Controls
Imports WMedia = System.Windows.Media

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

#Region "상태 (세션/결과 보관)"

        ' 삭제 묶음 스택 (직전 삭제 → 되돌리기 대상)
        Private ReadOnly _deleteOps As New Stack(Of List(Of Integer))()

        ' 마지막 스캔 결과(엑셀 내보내기 및 UI 상태 동기화용)
        Private _lastRows As New List(Of DupRowDto)

        ' 중첩(shared) 패밀리 인스턴스(서브컴포넌트) ID 집합
        Private Shared _nestedSharedIds As HashSet(Of Integer) = Nothing

        Private Class DupRowDto
            Public Property ElementId As Integer
            Public Property Category As String
            Public Property Family As String
            Public Property [Type] As String
            Public Property ConnectedCount As Integer
            Public Property ConnectedIds As String
            Public Property Candidate As Boolean
            Public Property Deleted As Boolean
        End Class

#End Region

#Region "핸들러"

        ' ====== 중복 스캔 ======
        Private Sub HandleDupRun(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                SendToWeb("host:error", New With {.message = "활성 문서가 없습니다."})
                Return
            End If

            Dim doc As Document = uiDoc.Document

            ' 중첩 Shared 컴포넌트 목록 캐시
            _nestedSharedIds = New HashSet(Of Integer)()
            Try
                Dim famCol As New FilteredElementCollector(doc)
                famCol.OfClass(GetType(FamilyInstance)).WhereElementIsNotElementType()
                For Each o As Element In famCol
                    Dim fi As FamilyInstance = TryCast(o, FamilyInstance)
                    If fi Is Nothing Then Continue For
                    Try
                        Dim subs = fi.GetSubComponentIds()
                        If subs Is Nothing Then Continue For
                        For Each sid As ElementId In subs
                            _nestedSharedIds.Add(sid.IntegerValue)
                        Next
                    Catch
                    End Try
                Next
            Catch
            End Try

            Dim tolFeet As Double = 1.0 / 64.0
            Try
                Dim tolObj = GetProp(payload, "tolFeet")
                If tolObj IsNot Nothing Then tolFeet = Math.Max(0.000001, Convert.ToDouble(tolObj))
            Catch
            End Try

            Dim rows As New List(Of DupRowDto)()
            Dim total As Integer = 0
            Dim groupsWithDup As Integer = 0
            Dim candidates As Integer = 0

            Dim collector As New FilteredElementCollector(doc)
            collector.WhereElementIsNotElementType()

            Dim q = Function(x As Double) As Long
                        Return CLng(Math.Round(x / tolFeet))
                    End Function

            Dim buckets As New Dictionary(Of String, List(Of ElementId))(StringComparer.Ordinal)
            Dim catCache As New Dictionary(Of Integer, String)
            Dim famCache As New Dictionary(Of Integer, String)
            Dim typCache As New Dictionary(Of Integer, String)

            For Each e As Element In collector
                total += 1
                If ShouldSkipForQuantity(e) Then Continue For
                If e Is Nothing OrElse e.Category Is Nothing Then Continue For

                Dim center As XYZ = TryGetCenter(e)
                If center Is Nothing Then Continue For

                Dim catName As String = SafeCategoryName(e, catCache)
                Dim famName As String = SafeFamilyName(e, famCache)
                Dim typName As String = SafeTypeName(e, typCache)
                Dim lvl As Integer = TryGetLevelId(e)

                Dim oriKey As String = GetOrientationKey(e)

                Dim key As String =
                  String.Concat(catName, "|",
                                famName, "|",
                                typName, "|",
                                "O", oriKey, "|",
                                "L", lvl.ToString(), "|",
                                "Q(", q(center.X).ToString(), ",", q(center.Y).ToString(), ",", q(center.Z).ToString(), ")")

                Dim list As List(Of ElementId) = Nothing
                If Not buckets.TryGetValue(key, list) Then
                    list = New List(Of ElementId)()
                    buckets.Add(key, list)
                End If
                list.Add(e.Id)
            Next

            For Each kv In buckets
                Dim ids As List(Of ElementId) = kv.Value
                If ids.Count <= 1 Then Continue For

                groupsWithDup += 1

                For Each id As ElementId In ids
                    Dim e As Element = doc.GetElement(id)
                    If e Is Nothing Then Continue For

                    Dim catName As String = SafeCategoryName(e, catCache)
                    Dim famName As String = SafeFamilyName(e, famCache)
                    Dim typName As String = SafeTypeName(e, typCache)

                    Dim connIds = ids.
                      Where(Function(x) x.IntegerValue <> id.IntegerValue).
                      Select(Function(x) x.IntegerValue.ToString()).
                      ToArray()

                    rows.Add(New DupRowDto With {
                      .ElementId = id.IntegerValue,
                      .Category = catName,
                      .Family = famName,
                      .Type = typName,
                      .ConnectedCount = connIds.Length,
                      .ConnectedIds = String.Join(", ", connIds),
                      .Candidate = True,
                      .Deleted = False
                    })
                    candidates += 1
                Next
            Next

            _lastRows = rows

            Dim wireRows = rows.Select(Function(r) New With {
              .elementId = r.ElementId,
              .category = r.Category,
              .family = r.Family,
              .type = r.Type,
              .connectedCount = r.ConnectedCount,
              .connectedIds = r.ConnectedIds,
              .candidate = r.Candidate,
              .deleted = r.Deleted
            }).ToList()

            SendToWeb("dup:list", wireRows)
            SendToWeb("dup:result", New With {.scan = total, .groups = groupsWithDup, .candidates = candidates})
        End Sub

        ' ====== 선택/줌 ======
        Private Sub HandleDuplicateSelect(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing Then Return

            Dim idVal As Integer = SafeInt(GetProp(payload, "id"))
            If idVal <= 0 Then Return

            Dim elId As New ElementId(idVal)
            Dim el As Element = uiDoc.Document.GetElement(elId)
            If el Is Nothing Then
                SendToWeb("host:warn", New With {.message = $"요소 {idVal} 을(를) 찾을 수 없습니다."})
                Return
            End If

            Try
                uiDoc.Selection.SetElementIds(New List(Of ElementId) From {elId})
            Catch
            End Try

            Dim bb As BoundingBoxXYZ = GetBoundingBox(el)
            Try
                If bb IsNot Nothing Then
                    Dim views = uiDoc.GetOpenUIViews()
                    Dim target = views.FirstOrDefault(Function(v) v.ViewId.IntegerValue = uiDoc.ActiveView.Id.IntegerValue)
                    If target IsNot Nothing Then
                        target.ZoomAndCenterRectangle(bb.Min, bb.Max)
                    Else
                        uiDoc.ShowElements(elId)
                    End If
                Else
                    uiDoc.ShowElements(elId)
                End If
            Catch
            End Try
        End Sub

        ' ====== 삭제(트랜잭션 1회 커밋) ======
        Private Sub HandleDuplicateDelete(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                SendToWeb("revit:error", New With {.message = "활성 문서를 찾을 수 없습니다."})
                Return
            End If

            Dim doc As Document = uiDoc.Document
            Dim ids As List(Of Integer) = ExtractIds(payload)
            If ids Is Nothing OrElse ids.Count = 0 Then
                SendToWeb("revit:error", New With {.message = "잘못된 요청입니다(id 누락/형식 오류)."})
                Return
            End If

            Dim eidList As New List(Of ElementId)
            For Each i In ids
                If i > 0 Then
                    Dim eid As New ElementId(i)
                    If doc.GetElement(eid) IsNot Nothing Then eidList.Add(eid)
                End If
            Next
            If eidList.Count = 0 Then
                SendToWeb("host:warn", New With {.message = "삭제할 유효한 요소가 없습니다."})
                Return
            End If

            Dim actuallyDeleted As New List(Of Integer)

            Using t As New Transaction(doc, $"KKY Dup Delete ({eidList.Count})")
                t.Start()
                Try
                    doc.Delete(eidList)
                    t.Commit()
                Catch ex As Exception
                    t.RollBack()
                    SendToWeb("revit:error", New With {.message = $"삭제 실패({eidList.Count}개): {ex.Message}"})
                    Return
                End Try
            End Using

            For Each eid In eidList
                If doc.GetElement(eid) Is Nothing Then
                    actuallyDeleted.Add(eid.IntegerValue)
                    Dim row = _lastRows.FirstOrDefault(Function(r) r.ElementId = eid.IntegerValue)
                    If row IsNot Nothing Then row.Deleted = True
                    SendToWeb("dup:deleted", New With {.id = eid.IntegerValue})
                End If
            Next

            If actuallyDeleted.Count > 0 Then
                _deleteOps.Push(actuallyDeleted)
            End If
        End Sub

        ' ====== 되돌리기(직전 삭제 묶음 Undo) ======
        Private Sub HandleDuplicateRestore(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                SendToWeb("revit:error", New With {.message = "활성 문서를 찾을 수 없습니다."})
                Return
            End If

            If _deleteOps.Count = 0 Then
                SendToWeb("host:warn", New With {.message = "되돌릴 수 있는 최신 삭제가 없습니다."})
                Return
            End If

            ' 요청으로 들어온 id (현재 UI는 단일 id 기준)
            Dim requestIds As List(Of Integer) = ExtractIds(payload)
            Dim lastPack As List(Of Integer) = _deleteOps.Peek()

            ' 요청 id 집합이 직전 삭제 묶음과 동일한지 확인
            Dim same As Boolean =
              requestIds IsNot Nothing AndAlso
              requestIds.Count = lastPack.Count AndAlso
              Not requestIds.Except(lastPack).Any()

            If Not same Then
                SendToWeb("host:warn", New With {.message = "되돌리기는 직전 삭제 묶음만 가능합니다."})
                Return
            End If

            Try
                ' 🔁 Revit 공식 Undo 포스터블 커맨드 사용
                Dim cmdId As RevitCommandId =
                  RevitCommandId.LookupPostableCommandId(PostableCommand.Undo)

                If cmdId Is Nothing Then
                    Throw New InvalidOperationException("Undo 명령을 찾을 수 없습니다.")
                End If

                uiDoc.Application.PostCommand(cmdId)
            Catch ex As Exception
                SendToWeb("revit:error", New With {.message = $"되돌리기 실패: {ex.Message}"})
                Return
            End Try

            ' 스택에서 제거 후 상태/UI 동기화
            _deleteOps.Pop()

            For Each i In lastPack
                Dim r = _lastRows.FirstOrDefault(Function(x) x.ElementId = i)
                If r IsNot Nothing Then r.Deleted = False
                SendToWeb("dup:restored", New With {.id = i})
            Next
        End Sub

        ' ====== 엑셀 내보내기 ======
        Private Sub HandleDuplicateExport(app As UIApplication, Optional payload As Object = Nothing)
            If _lastRows Is Nothing OrElse _lastRows.Count = 0 Then
                SendToWeb("host:warn", New With {.message = "내보낼 데이터가 없습니다."})
                Return
            End If

            Dim token As String = TryCast(GetProp(payload, "token"), String)

            Try
                ' ⭐ 중복 그룹 수를 먼저 계산해서 파일 이름과 팝업 모두에 사용
                Dim groupsCount As Integer = CountGroups(_lastRows)

                Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                Dim todayToken As String = Date.Now.ToString("yyMMdd")
                Dim defaultFileName As String = $"{todayToken}_중복객체 검토결과_{groupsCount}개.xlsx"
                Dim defaultPath As String = Path.Combine(desktop, defaultFileName)

                Dim sfd As New Microsoft.Win32.SaveFileDialog() With {
                  .Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                  .FileName = Path.GetFileName(defaultPath),
                  .AddExtension = True,
                  .DefaultExt = "xlsx",
                  .OverwritePrompt = True
                }

                If sfd.ShowDialog() <> True Then
                    Exit Sub
                End If

                Dim outPath As String = sfd.FileName

                ' 엑셀 내보내기
                Dim doAutoFit As Boolean = ParseExcelMode(payload)
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset("dup:progress")
                Exports.DuplicateExport.Save(outPath, _lastRows.Cast(Of Object)(), doAutoFit, "dup:progress")

                SendToWeb("dup:exported", New With {.path = outPath, .ok = True, .token = token})
            Catch ioEx As IOException
                Dim msg As String =
                  "해당 파일이 열려 있어 저장에 실패했습니다. 엑셀에서 파일을 닫은 뒤 다시 시도해 주세요."
                SendToWeb("dup:exported", New With {.ok = False, .message = msg, .token = token})
            Catch ex As Exception
                SendToWeb("dup:exported", New With {.ok = False, .message = $"엑셀 내보내기에 실패했습니다: {ex.Message}", .token = token})
            End Try
        End Sub

#End Region

#Region "필터/유틸(물량 필터 강화)"

        ' ⭐ 실제 시공 물량으로 보는 객체만 남기기 위한 필터
        Private Shared Function ShouldSkipForQuantity(e As Element) As Boolean
            ' 0) 기본 예외: 널 / 임포트
            If e Is Nothing Then Return True
            If TypeOf e Is ImportInstance Then Return True

            ' 1) 카테고리 기반 필터
            Try
                If e.Category Is Nothing Then Return True
                If e.Category.CategoryType <> CategoryType.Model Then Return True

                Dim n As String = If(e.Category.Name, "").ToLowerInvariant()

                ' 뷰/표현/기준 관련
                If n.Contains("view") OrElse n.Contains("viewport") Then Return True
                If n.Contains("level") OrElse n.Contains("grid") Then Return True
                If n.Contains("reference plane") OrElse n.Contains("work plane") Then Return True
                If n.Contains("scope box") OrElse n.Contains("matchline") Then Return True
                If n.Contains("section line") OrElse n.Contains("callout") Then Return True

                ' 시트 / 선 / 스케치 / 영역 / 디테일 계열
                If n.Contains("sheet") OrElse n.Contains("시트") Then Return True
                If n.Contains("line") OrElse n.Contains("선") Then Return True
                If n.Contains("sketch") OrElse n.Contains("area boundary") Then Return True
                If n.Contains("filled region") OrElse n.Contains("detail item") Then Return True
                If n.Contains("detail line") OrElse n.Contains("symbol") Then Return True

                ' 텍스트/치수/태그
                If n.Contains("text note") OrElse n.Contains("dimension") Then Return True
                If n.Contains("room tag") OrElse n.Contains("space tag") OrElse n.Contains("area tag") Then Return True

                ' 중심선 / 해석 모델
                If n.Contains("center line") OrElse n.Contains("centerline") OrElse n.Contains("중심선") Then Return True
                If n.StartsWith("analytical") Then Return True
            Catch
                ' 카테고리 이름 읽다가 터지면 보수적으로 제외
                Return True
            End Try

            ' 2) 중첩 패밀리(패밀리 안 패밀리) → 상위만 남기고 서브는 제외
            Dim fi = TryCast(e, FamilyInstance)
            If fi IsNot Nothing Then
                Try
                    ' 상위 패밀리(호스트)에 붙은 서브컴포넌트는 스킵
                    If fi.SuperComponent IsNot Nothing Then Return True
                    If _nestedSharedIds IsNot Nothing AndAlso _nestedSharedIds.Contains(fi.Id.IntegerValue) Then Return True
                Catch
                End Try
            End If

            ' 3) 실제 3D 형상 여부 (솔리드)
            Try
                Dim opts As New Options() With {
                    .ComputeReferences = False,
                    .IncludeNonVisibleObjects = False
                }

                If Not HasPositiveSolid(e, opts) Then
                    Return True
                End If
            Catch
                ' 지오메트리 조회 중 예외 → 보수적으로 제외
                Return True
            End Try

            ' 여기까지 살아남으면 물량 대상
            Return False
        End Function

        Private Shared Function HasPositiveSolid(el As Element, opts As Options) As Boolean
            Dim geom As GeometryElement = el.Geometry(opts)
            If geom Is Nothing Then Return False
            For Each g As GeometryObject In geom
                Dim s As Solid = TryCast(g, Solid)
                If s IsNot Nothing AndAlso s.Volume > 0 Then Return True
                Dim inst As GeometryInstance = TryCast(g, GeometryInstance)
                If inst IsNot Nothing Then
                    Dim instGeom As GeometryElement = inst.GetInstanceGeometry()
                    If instGeom IsNot Nothing Then
                        For Each gi As GeometryObject In instGeom
                            Dim si As Solid = TryCast(gi, Solid)
                            If si IsNot Nothing AndAlso si.Volume > 0 Then Return True
                        Next
                    End If
                End If
            Next
            Return False
        End Function

        Private Shared Function QOri(x As Double) As Long
            Return CLng(Math.Round(x * 1000.0R))
        End Function

        Private Shared Function GetOrientationKey(e As Element) As String
            Try
                Dim fi = TryCast(e, FamilyInstance)
                If fi IsNot Nothing Then
                    Dim mirrored As Boolean = False
                    Dim hand As Boolean = False
                    Dim facing As Boolean = False

                    Try
                        mirrored = fi.Mirrored
                    Catch
                    End Try

                    Try
                        hand = fi.HandFlipped
                    Catch
                    End Try

                    Try
                        facing = fi.FacingFlipped
                    Catch
                    End Try

                    Dim t As Transform = Nothing
                    Try
                        t = fi.GetTransform()
                    Catch
                    End Try

                    Dim keyParts As New List(Of String)()
                    keyParts.Add("M" & If(mirrored, "1", "0"))
                    keyParts.Add("H" & If(hand, "1", "0"))
                    keyParts.Add("F" & If(facing, "1", "0"))

                    If t IsNot Nothing Then
                        Dim ox = t.BasisX
                        Dim oy = t.BasisY
                        Dim oz = t.BasisZ
                        keyParts.Add("OX(" & QOri(ox.X) & "," & QOri(ox.Y) & "," & QOri(ox.Z) & ")")
                        keyParts.Add("OY(" & QOri(oy.X) & "," & QOri(oy.Y) & "," & QOri(oy.Z) & ")")
                        keyParts.Add("OZ(" & QOri(oz.X) & "," & QOri(oz.Y) & "," & QOri(oz.Z) & ")")
                    End If

                    Return String.Join("|", keyParts)
                End If

                Dim loc As Location = Nothing
                Try
                    loc = e.Location
                Catch
                End Try

                Dim lc = TryCast(loc, LocationCurve)
                If lc IsNot Nothing AndAlso lc.Curve IsNot Nothing Then
                    Dim c = lc.Curve
                    Dim dir As XYZ = Nothing
                    Try
                        dir = (c.GetEndPoint(1) - c.GetEndPoint(0))
                    Catch
                    End Try

                    If dir IsNot Nothing Then
                        Dim len As Double = dir.GetLength()
                        If len > 0.000001 Then
                            dir = dir / len
                        End If

                        Return "LC(" & QOri(dir.X) & "," & QOri(dir.Y) & "," & QOri(dir.Z) & ")"
                    End If
                End If
            Catch
            End Try

            Return String.Empty
        End Function

        Private Shared Function SafeCategoryName(e As Element, cache As Dictionary(Of Integer, String)) As String
            If e Is Nothing OrElse e.Category Is Nothing Then Return ""
            Dim id As Integer = e.Category.Id.IntegerValue
            Dim s As String = Nothing
            If cache.TryGetValue(id, s) Then Return s
            s = e.Category.Name
            cache(id) = s
            Return s
        End Function

        Private Shared Function SafeFamilyName(e As Element, cache As Dictionary(Of Integer, String)) As String
            Dim fi = TryCast(e, FamilyInstance)
            If fi Is Nothing OrElse fi.Symbol Is Nothing OrElse fi.Symbol.Family Is Nothing Then Return ""
            Dim id As Integer = fi.Symbol.Family.Id.IntegerValue
            Dim s As String = Nothing
            If cache.TryGetValue(id, s) Then Return s
            s = fi.Symbol.Family.Name
            cache(id) = s
            Return s
        End Function

        Private Shared Function SafeTypeName(e As Element, cache As Dictionary(Of Integer, String)) As String
            Dim fi = TryCast(e, FamilyInstance)
            If fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing Then
                Dim id As Integer = fi.Symbol.Id.IntegerValue
                Dim s As String = Nothing
                If cache.TryGetValue(id, s) Then Return s
                s = fi.Symbol.Name
                cache(id) = s
                Return s
            End If
            Return e.Name
        End Function

        Private Shared Function TryGetLevelId(e As Element) As Integer
            Try
                Dim p As Parameter = e.Parameter(BuiltInParameter.LEVEL_PARAM)
                If p IsNot Nothing Then
                    Dim lvid As ElementId = p.AsElementId()
                    If lvid IsNot Nothing AndAlso lvid <> ElementId.InvalidElementId Then
                        Return lvid.IntegerValue
                    End If
                End If
            Catch
            End Try
            Try
                Dim pi = e.GetType().GetProperty("LevelId")
                If pi IsNot Nothing Then
                    Dim id = TryCast(pi.GetValue(e, Nothing), ElementId)
                    If id IsNot Nothing AndAlso id <> ElementId.InvalidElementId Then
                        Return id.IntegerValue
                    End If
                End If
            Catch
            End Try
            Return -1
        End Function

        Private Shared Function TryGetCenter(e As Element) As XYZ
            If e Is Nothing Then Return Nothing
            Try
                Dim loc As Location = e.Location
                If TypeOf loc Is LocationPoint Then
                    Return CType(loc, LocationPoint).Point
                ElseIf TypeOf loc Is LocationCurve Then
                    Dim crv = CType(loc, LocationCurve).Curve
                    If crv IsNot Nothing Then
                        Return crv.Evaluate(0.5, True)
                    End If
                End If
            Catch
            End Try

            Dim bb = GetBoundingBox(e)
            If bb IsNot Nothing Then
                Return (bb.Min + bb.Max) * 0.5
            End If
            Return Nothing
        End Function

        Private Shared Function GetBoundingBox(e As Element) As BoundingBoxXYZ
            Try
                Dim bb As BoundingBoxXYZ = e.BoundingBox(Nothing)
                If bb IsNot Nothing Then Return bb
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function SafeInt(o As Object) As Integer
            If o Is Nothing Then Return 0
            Try
                Return Convert.ToInt32(o)
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function ExtractIds(payload As Object) As List(Of Integer)
            Dim result As New List(Of Integer)

            Dim singleObj = GetProp(payload, "id")
            Dim v As Integer = SafeToInt(singleObj)
            If v > 0 Then
                result.Add(v)
                Return result
            End If

            Dim arr = GetProp(payload, "ids")
            If arr Is Nothing Then Return result

            Dim enumerable = TryCast(arr, System.Collections.IEnumerable)
            If enumerable IsNot Nothing Then
                For Each o In enumerable
                    Dim iv = SafeToInt(o)
                    If iv > 0 Then result.Add(iv)
                Next
            End If

            Return result
        End Function

        Private Shared Function SafeToInt(o As Object) As Integer
            If o Is Nothing Then Return 0
            Try
                If TypeOf o Is Integer Then Return CInt(o)
                If TypeOf o Is Long Then Return CInt(CLng(o))
                If TypeOf o Is Double Then Return CInt(CDbl(o))
                If TypeOf o Is String Then
                    Dim s As String = CStr(o)
                    Dim iv As Integer
                    If Integer.TryParse(s, iv) Then Return iv
                End If
            Catch
            End Try
            Return 0
        End Function

        Private Function CountGroups(rows As IEnumerable(Of DupRowDto)) As Integer
            If rows Is Nothing Then Return 0
            Dim bucket As New HashSet(Of String)(StringComparer.Ordinal)
            For Each r In rows
                Dim id As String = r.ElementId.ToString()
                Dim cat As String = If(r.Category, "")
                Dim fam As String = If(r.Family, "")
                Dim typ As String = If(r.Type, "")
                Dim conStr As String = If(r.ConnectedIds, "")

                Dim cluster As New List(Of String)
                If Not String.IsNullOrWhiteSpace(id) Then cluster.Add(id)
                cluster.AddRange(SplitIds(conStr))

                Dim norm = cluster.
                  Where(Function(x) Not String.IsNullOrWhiteSpace(x)).
                  [Select](Function(x) x.Trim()).
                  Distinct().
                  OrderBy(Function(x) x).
                  ToList()

                Dim clusterKey As String = If(norm.Count > 1, String.Join(",", norm), "")
                Dim famOut As String = If(String.IsNullOrWhiteSpace(fam), If(String.IsNullOrWhiteSpace(cat), "", cat & " Type"), fam)
                Dim key = String.Join("|", {cat, famOut, typ, clusterKey})
                bucket.Add(key)
            Next
            Return bucket.Count
        End Function

        Private Function SplitIds(s As String) As IEnumerable(Of String)
            If String.IsNullOrWhiteSpace(s) Then Return Array.Empty(Of String)()
            Return s.Split(New Char() {","c, " "c, ";"c, "|"c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf},
                           StringSplitOptions.RemoveEmptyEntries)
        End Function

#End Region

#Region "WPF 팝업 (실패 시 TaskDialog 폴백)"

        Private Function IsSystemDark() As Boolean
            Try
                Dim key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                If key IsNot Nothing Then
                    Dim v = key.GetValue("AppsUseLightTheme", 1)
                    Return (Convert.ToInt32(v) = 0)
                End If
            Catch
            End Try
            Return False
        End Function

        ''' <summary>
        ''' 엑셀 내보내기 안내를 WPF로 시도하고 실패하면 TaskDialog로 폴백한다.
        ''' 반환값: True = 파일 열기
        ''' </summary>
        Private Function ShowExcelSavedDialog(outPath As String, groupsCount As Integer,
                                              Optional chipLabel As String = "중복 그룹",
                                              Optional dialogTitle As String = "중복검토 내보내기",
                                              Optional headerText As String = "엑셀로 저장했습니다.",
                                              Optional questionText As String = "지금 파일을 열어보시겠어요?") As Boolean
            Try
                Dim win As New WPF.Window()
                win.Title = "KKY Tool_Revit - " & dialogTitle
                win.SizeToContent = WPF.SizeToContent.WidthAndHeight
                win.WindowStartupLocation = WPF.WindowStartupLocation.CenterOwner
                win.ResizeMode = WPF.ResizeMode.NoResize
                win.Topmost = True
                win.Content = BuildExcelSavedContent(outPath, groupsCount, chipLabel, headerText, questionText, win)

                ' Owner 연결(가능하면 Revit 메인 윈도우에 붙이기)
                Try
                    Dim t = Type.GetType("Autodesk.Windows.ComponentManager, AdWindows")
                    If t IsNot Nothing Then
                        Dim p = t.GetProperty("ApplicationWindow",
                                 Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static)
                        If p IsNot Nothing Then
                            Dim hwnd = CType(p.GetValue(Nothing, Nothing), IntPtr)
                            Dim helper = New WPF.Interop.WindowInteropHelper(win)
                            helper.Owner = hwnd
                        End If
                    End If
                Catch
                End Try

                Dim res As Boolean? = win.ShowDialog()
                Return If(res.HasValue AndAlso res.Value, True, False)
            Catch
                ' 폴백: TaskDialog
                Dim td As New TaskDialog(dialogTitle)
                td.MainIcon = TaskDialogIcon.TaskDialogIconInformation
                td.MainInstruction = headerText
                td.MainContent = $"{chipLabel}: {groupsCount}개{Environment.NewLine}파일: {outPath}"
                td.CommonButtons = TaskDialogCommonButtons.Yes Or TaskDialogCommonButtons.No
                td.DefaultButton = TaskDialogResult.Yes
                td.FooterText = questionText
                Dim r = td.Show()
                Return r = TaskDialogResult.Yes
            End Try
        End Function

        Private Function BuildExcelSavedContent(outPath As String, groupsCount As Integer,
                                               chipLabel As String,
                                               headerText As String,
                                               questionText As String,
                                               host As WPF.Window) As WPF.UIElement
            Dim isDark = IsSystemDark()

            ' 테마별 색상 정의 (Byte 캐스팅으로 Option Strict 대응)
            Dim bgPanel As WMedia.Color =
              If(isDark,
                 WMedia.Color.FromRgb(CByte(&H12), CByte(&H16), CByte(&H1C)),
                 WMedia.Color.FromRgb(CByte(&HFF), CByte(&HFF), CByte(&HFF)))

            Dim bgCard As WMedia.Color =
              If(isDark,
                 WMedia.Color.FromRgb(CByte(&H18), CByte(&H1C), CByte(&H24)),
                 WMedia.Color.FromRgb(CByte(&HF7), CByte(&HF8), CByte(&HFA)))

            Dim headG1 As WMedia.Color =
              If(isDark,
                 WMedia.Color.FromRgb(CByte(&H1F), CByte(&H5A), CByte(&HFF)),
                 WMedia.Color.FromRgb(CByte(&H66), CByte(&H99), CByte(&HFF)))

            Dim headG2 As WMedia.Color =
              If(isDark,
                 WMedia.Color.FromRgb(CByte(&H78), CByte(&H9B), CByte(&HFF)),
                 WMedia.Color.FromRgb(CByte(&H9F), CByte(&HBE), CByte(&HFF)))

            Dim fgMain As WMedia.Color =
              If(isDark,
                 WMedia.Color.FromRgb(CByte(&HE8), CByte(&HEA), CByte(&HED)),
                 WMedia.Color.FromRgb(CByte(&H11), CByte(&H11), CByte(&H11)))

            Dim fgSub As WMedia.Color =
              If(isDark,
                 WMedia.Color.FromRgb(CByte(&HC7), CByte(&HC9), CByte(&HCC)),
                 WMedia.Color.FromRgb(CByte(&H55), CByte(&H55), CByte(&H55)))

            Dim chipBg As WMedia.Color =
              If(isDark,
                 WMedia.Color.FromRgb(CByte(&H21), CByte(&H26), CByte(&H32)),
                 WMedia.Color.FromRgb(CByte(&HEE), CByte(&HF1), CByte(&HF5)))

            Dim accent As WMedia.Color =
              If(isDark,
                 WMedia.Color.FromRgb(CByte(&H7A), CByte(&HA2), CByte(&HFF)),
                 WMedia.Color.FromRgb(CByte(&H38), CByte(&H67), CByte(&HFF)))

            Dim bdLine As WMedia.Color =
              If(isDark,
                 WMedia.Color.FromArgb(CByte(&H33), CByte(&HFF), CByte(&HFF), CByte(&HFF)),
                 WMedia.Color.FromArgb(CByte(&H22), CByte(&H0), CByte(&H0), CByte(&H0)))

            Dim root As New WControls.Border() With {
              .Background = New WMedia.SolidColorBrush(bgPanel),
              .Padding = New WPF.Thickness(16)
            }

            Dim card As New WControls.Border() With {
              .Padding = New WPF.Thickness(0),
              .CornerRadius = New WPF.CornerRadius(14),
              .BorderThickness = New WPF.Thickness(1),
              .BorderBrush = New WMedia.SolidColorBrush(bdLine),
              .Background = New WMedia.SolidColorBrush(bgCard),
              .Effect = New WMedia.Effects.DropShadowEffect() With {
                .Opacity = 0.25,
                .BlurRadius = 16,
                .ShadowDepth = 0
              }
            }

            Dim wrap As New WControls.StackPanel() With {
              .Width = 560
            }

            ' 헤더 (그라데이션 바)
            Dim header As New WControls.Border() With {
              .CornerRadius = New WPF.CornerRadius(14, 14, 0, 0),
              .Background = New WMedia.LinearGradientBrush(headG1, headG2, 0),
              .Padding = New WPF.Thickness(20, 14, 20, 16)
            }

            Dim hTitle As New WControls.TextBlock() With {
              .Text = headerText,
              .FontSize = 18,
              .FontWeight = WPF.FontWeights.SemiBold,
              .Foreground = WMedia.Brushes.White
            }
            header.Child = hTitle

            ' 바디 패딩
            Dim bodyPad As New WControls.Border() With {
              .Padding = New WPF.Thickness(20),
              .Background = New WMedia.SolidColorBrush(bgCard),
              .CornerRadius = New WPF.CornerRadius(0, 0, 14, 14)
            }

            Dim body As New WControls.StackPanel() With {
              .Orientation = WControls.Orientation.Vertical
            }

            ' 중복 그룹 수 칩
            Dim chip As New WControls.Border() With {
              .CornerRadius = New WPF.CornerRadius(999),
              .Background = New WMedia.SolidColorBrush(chipBg),
              .Padding = New WPF.Thickness(12, 6, 12, 6),
              .Margin = New WPF.Thickness(0, 8, 0, 10)
            }

            Dim chipText As New WControls.TextBlock() With {
              .Text = $"{chipLabel} {groupsCount}개",
              .Foreground = New WMedia.SolidColorBrush(accent),
              .FontWeight = WPF.FontWeights.SemiBold
            }
            chip.Child = chipText

            ' 파일 경로
            Dim pathTb As New WControls.TextBlock() With {
              .Text = $"파일: {outPath}",
              .TextWrapping = WPF.TextWrapping.Wrap,
              .Foreground = New WMedia.SolidColorBrush(fgSub),
              .Margin = New WPF.Thickness(0, 0, 0, 14)
            }

            ' 질문 텍스트
            Dim question As New WControls.TextBlock() With {
              .Text = questionText,
              .Foreground = New WMedia.SolidColorBrush(fgMain),
              .Margin = New WPF.Thickness(0, 4, 0, 10)
            }

            ' 버튼 바
            Dim btnBar As New WControls.StackPanel() With {
              .Orientation = WControls.Orientation.Horizontal,
              .HorizontalAlignment = WPF.HorizontalAlignment.Right
            }

            Dim yesBtn As New WControls.Button() With {
              .Content = "예(Y)",
              .MinWidth = 88,
              .Padding = New WPF.Thickness(14, 7, 14, 7),
              .Margin = New WPF.Thickness(0, 0, 8, 0),
              .Foreground = WMedia.Brushes.White,
              .Background = New WMedia.SolidColorBrush(accent),
              .BorderBrush = WMedia.Brushes.Transparent
            }

            Dim noBtn As New WControls.Button() With {
              .Content = "아니오(N)",
              .MinWidth = 88,
              .Padding = New WPF.Thickness(14, 7, 14, 7),
              .Foreground = New WMedia.SolidColorBrush(fgMain),
              .Background = New WMedia.SolidColorBrush(chipBg),
              .BorderBrush = WMedia.Brushes.Transparent
            }

            AddHandler yesBtn.Click,
              Sub(sender As Object, e As WPF.RoutedEventArgs)
                  host.DialogResult = True
                  host.Close()
              End Sub

            AddHandler noBtn.Click,
              Sub(sender As Object, e As WPF.RoutedEventArgs)
                  host.DialogResult = False
                  host.Close()
              End Sub

            btnBar.Children.Add(yesBtn)
            btnBar.Children.Add(noBtn)

            body.Children.Add(chip)
            body.Children.Add(pathTb)
            body.Children.Add(question)
            body.Children.Add(btnBar)

            bodyPad.Child = body

            wrap.Children.Add(header)
            wrap.Children.Add(bodyPad)

            card.Child = wrap
            root.Child = card

            Return root
        End Function

#End Region

    End Class

End Namespace
