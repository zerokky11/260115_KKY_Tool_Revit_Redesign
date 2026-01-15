Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.DB.Mechanical
Imports Autodesk.Revit.DB.Plumbing
Imports Autodesk.Revit.DB.Electrical
Imports Autodesk.Revit.UI

Namespace Services

    Public Class ConnectorDiagnosticsService

        Private Class ParamInfo
            Public Property Exists As Boolean
            Public Property HasValue As Boolean
            Public Property Text As String
        End Class

        ' === 디버그 로그 (호출자가 읽음) ===
        Public Shared Property LastDebug As List(Of String)

        Private Shared Sub Log(msg As String)
            If LastDebug Is Nothing Then LastDebug = New List(Of String)()
            LastDebug.Add($"{DateTime.Now:HH\:mm\:ss.fff} {msg}")
        End Sub

        Private Class TargetFilter
            Public Property Evaluator As Func(Of Element, Boolean)
            Public Property PrimaryParam As String = String.Empty
        End Class

        ' ============================
        ' Public Entry Points
        ' ============================

        ' 3-인자: tolFt 는 피트 단위 (ft)
        Public Shared Function Run(app As UIApplication, tolFt As Double, param As String,
                                   Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Return Run(app, tolFt, param, CType(Nothing, IEnumerable(Of String)), Nothing, False, False, progress)
        End Function

        Public Shared Function Run(app As UIApplication, tolFt As Double, param As String,
                                   extraParams As IEnumerable(Of String),
                                   Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Return Run(app, tolFt, param, extraParams, Nothing, False, False, progress)
        End Function

        Public Shared Function Run(app As UIApplication, tolFt As Double, param As String,
                                   extraParams As IEnumerable(Of String),
                                   targetFilter As String,
                                   excludeEndDummy As Boolean,
                                   Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Return Run(app, tolFt, param, extraParams, targetFilter, excludeEndDummy, False, progress)
        End Function

        ' ✅ includeOkRows: OK까지 포함하여 추출할지 여부 (기본 False)
        Public Shared Function Run(app As UIApplication, tolFt As Double, param As String,
                                   extraParams As IEnumerable(Of String),
                                   targetFilter As String,
                                   excludeEndDummy As Boolean,
                                   includeOkRows As Boolean,
                                   Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))

            LastDebug = New List(Of String)()
            Dim rows As New List(Of Dictionary(Of String, Object))()

            Dim uidoc = app.ActiveUIDocument
            If uidoc Is Nothing OrElse uidoc.Document Is Nothing Then
                Log("ActiveUIDocument 없음")
                Return rows
            End If

            Dim doc = uidoc.Document
            Dim fileLabel = BuildFileLabel(doc)
            Return RunCore(doc, tolFt, param, extraParams, targetFilter, excludeEndDummy, includeOkRows, progress, fileLabel)
        End Function

        Public Shared Function RunOnDocument(doc As Document,
                                             tolFt As Double,
                                             param As String,
                                             extraParams As IEnumerable(Of String),
                                             targetFilter As String,
                                             excludeEndDummy As Boolean,
                                             Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Return RunOnDocument(doc, tolFt, param, extraParams, targetFilter, excludeEndDummy, False, progress)
        End Function

        Public Shared Function RunOnDocument(doc As Document,
                                             tolFt As Double,
                                             param As String,
                                             extraParams As IEnumerable(Of String),
                                             targetFilter As String,
                                             excludeEndDummy As Boolean,
                                             includeOkRows As Boolean,
                                             Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            LastDebug = New List(Of String)()
            If doc Is Nothing Then
                Log("Document 없음")
                Return New List(Of Dictionary(Of String, Object))()
            End If

            Dim fileLabel = BuildFileLabel(doc)
            Return RunCore(doc, tolFt, param, extraParams, targetFilter, excludeEndDummy, includeOkRows, progress, fileLabel)
        End Function

        Public Shared Function RunOnDocument(doc As Document,
                                             tol As Double,
                                             unit As String,
                                             paramName As String,
                                             extraParams As IEnumerable(Of String),
                                             targetFilter As String,
                                             excludeEndDummy As Boolean,
                                             Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Dim tolFt As Double = ToTolFt(tol, unit)
            Return RunOnDocument(doc, tolFt, paramName, extraParams, targetFilter, excludeEndDummy, False, progress)
        End Function

        ' 4-인자: tol 은 unit 기준(mm/inch/ft) → 내부에서 ft 로 환산 후 3-인자 호출
        Public Shared Function Run(app As UIApplication, tol As Double, unit As String, paramName As String,
                                   Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Return Run(app, ToTolFt(tol, unit), paramName, CType(Nothing, IEnumerable(Of String)), Nothing, False, False, progress)
        End Function

        Public Shared Function Run(app As UIApplication, tol As Double, unit As String, paramName As String,
                                   extraParams As IEnumerable(Of String),
                                   Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Return Run(app, ToTolFt(tol, unit), paramName, extraParams, Nothing, False, False, progress)
        End Function

        Public Shared Function Run(app As UIApplication, tol As Double, unit As String, paramName As String,
                                   extraParams As IEnumerable(Of String),
                                   targetFilter As String,
                                   excludeEndDummy As Boolean,
                                   Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Return Run(app, ToTolFt(tol, unit), paramName, extraParams, targetFilter, excludeEndDummy, False, progress)
        End Function

        ' ============================
        ' Core
        ' ============================

        Private Shared Function RunCore(doc As Document,
                                        tolFt As Double,
                                        param As String,
                                        extraParams As IEnumerable(Of String),
                                        targetFilter As String,
                                        excludeEndDummy As Boolean,
                                        includeOkRows As Boolean,
                                        progress As Action(Of Double, String),
                                        fileLabel As String) As List(Of Dictionary(Of String, Object))

            Dim rows As New List(Of Dictionary(Of String, Object))()

            If doc Is Nothing Then
                Log("Document 없음")
                Return rows
            End If

            Dim normalizedExtras As List(Of String) = Nothing

            Try
                normalizedExtras = NormalizeExtraParams(extraParams)
                Dim extraCache As New Dictionary(Of Integer, Dictionary(Of String, String))()
                Dim filter = ParseTargetFilter(targetFilter)

                Log($"DOC={fileLabel}, tolFt={tolFt:0.###}, param='{param}', extra={String.Join(",", normalizedExtras)}, targetFilter='{targetFilter}', excludeEndDummy={excludeEndDummy}, includeOkRows={includeOkRows}")

                Dim elems = CollectElementsWithConnectors(doc, filter, excludeEndDummy)
                Log($"수집 요소: {elems.Count}")

                If elems.Count = 0 Then
                    Log("커넥터를 가진 요소가 없습니다.")
                    Return rows
                End If

                Dim allowedIds As HashSet(Of Integer) = New HashSet(Of Integer)(elems.Select(Function(e) e.Id.IntegerValue))

                Dim elemConns As New Dictionary(Of Integer, List(Of Connector))()
                For Each el In elems
                    elemConns(el.Id.IntegerValue) = GetConnectors(el)
                Next

                Dim totalElem As Integer = Math.Max(1, elems.Count)
                Dim lastSentPct As Double = -1
                Dim seenPairs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                ' 모든 커넥터 좌표 버킷 구성 (1ft 셀)
                Dim allConnPoints As New List(Of Tuple(Of Integer, XYZ, Connector))()
                For Each kv In elemConns
                    For Each c In kv.Value
                        If c Is Nothing Then Continue For
                        Dim org As XYZ = Nothing
                        Try
                            org = c.Origin
                        Catch
                            Continue For
                        End Try
                        If org Is Nothing Then Continue For
                        allConnPoints.Add(Tuple.Create(kv.Key, org, c))
                    Next
                Next

                Dim buckets = BuildGrid(allConnPoints)
                Log($"버킷 수: {buckets.Count}")

                For i As Integer = 0 To elems.Count - 1
                    Dim el = elems(i)
                    Dim baseId = el.Id.IntegerValue
                    Dim conns = elemConns(baseId)

                    Dim connTotal As Integer = 1
                    If conns IsNot Nothing Then connTotal = Math.Max(1, conns.Count)

                    Dim j As Integer = 0
                    For Each c In conns
                        j += 1

                        Try
                            If c Is Nothing Then Continue For

                            Dim found As Element = Nothing
                            Dim distFt As Double = 0
                            Dim connType As String = ""

                            ' 1) 실제 연결 (Physical)
                            If c.IsConnected Then
                                For Each r As Connector In c.AllRefs.Cast(Of Connector)()
                                    If r?.Owner Is Nothing Then Continue For
                                    If r.Owner.Id.IntegerValue = baseId Then Continue For
                                    If TypeOf r.Owner Is MEPSystem Then Continue For

                                    ' ✅ 활성문서 검토 로직과 동일하게: 수집 대상 밖은 제외
                                    If Not allowedIds.Contains(r.Owner.Id.IntegerValue) Then Continue For

                                    found = r.Owner
                                    connType = "Physical(커넥터 연결 됨)"
                                    Exit For
                                Next
                            End If

                            ' 2) 근접 후보(Proximity) - type match 우선, 실패 시 type 무시
                            If found Is Nothing Then
                                Dim best = FindProximityCandidate(c, buckets, allowedIds, baseId, tolFt, True)
                                If best.Item1 = 0 Then
                                    best = FindProximityCandidate(c, buckets, allowedIds, baseId, tolFt, False)
                                End If
                                If best.Item1 <> 0 Then
                                    found = doc.GetElement(New ElementId(best.Item1))
                                    distFt = best.Item2
                                    connType = "Proximity(커넥터 연결 필요)"
                                End If
                            End If

                            If String.IsNullOrEmpty(connType) Then connType = "연결 대상 객체 없음"

                            Dim distInch As Double = Math.Round(distFt * 12.0, 2)

                            Dim info1 = GetParamInfo(el, param)
                            Dim info2 As ParamInfo = If(found IsNot Nothing,
                                                        GetParamInfo(found, param),
                                                        New ParamInfo() With {.Exists = False, .HasValue = False, .Text = ""})

                            Dim paramCompare As String = "N/A"
                            Dim issueStatus As String = Nothing

                            If found Is Nothing Then
                                issueStatus = "연결 대상 객체 없음"
                            Else
                                If Not info1.Exists OrElse Not info2.Exists Then
                                    issueStatus = "Shared Parameter 등록 필요"
                                Else
                                    If Not info1.HasValue AndAlso Not info2.HasValue Then
                                        paramCompare = "BothEmpty"
                                    ElseIf String.Equals(info1.Text, info2.Text, StringComparison.OrdinalIgnoreCase) Then
                                        paramCompare = "Match"
                                    Else
                                        paramCompare = "Mismatch"
                                    End If
                                End If

                                If issueStatus Is Nothing Then
                                    If connType.IndexOf("Proximity", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                        issueStatus = "연결 필요(Proximity)"
                                    ElseIf String.Equals(paramCompare, "Mismatch", StringComparison.OrdinalIgnoreCase) Then
                                        issueStatus = "Mismatch"
                                    Else
                                        issueStatus = "OK"
                                    End If
                                End If
                            End If

                            Dim v1 As String = If(info1.Exists, info1.Text, "(미등록)")
                            Dim v2 As String = If(info2.Exists, info2.Text, "(미등록)")
                            If found Is Nothing Then v2 = ""

                            Dim extras1 = GetExtraValues(el, normalizedExtras, extraCache)
                            Dim extras2 = GetExtraValues(found, normalizedExtras, extraCache)

                            Dim shouldAdd As Boolean =
                                String.Equals(issueStatus, "Mismatch", StringComparison.OrdinalIgnoreCase) OrElse
                                String.Equals(issueStatus, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) OrElse
                                String.Equals(issueStatus, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) OrElse
                                String.Equals(issueStatus, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) OrElse
                                (includeOkRows AndAlso String.Equals(issueStatus, "OK", StringComparison.OrdinalIgnoreCase))

                            If shouldAdd Then
                                Dim pairKey As String
                                Dim originKey As String = ""
                                Try
                                    originKey = $"{Math.Round(c.Origin.X, 4)},{Math.Round(c.Origin.Y, 4)},{Math.Round(c.Origin.Z, 4)}"
                                Catch
                                End Try

                                If found IsNot Nothing Then
                                    Dim id1 = baseId
                                    Dim id2 = found.Id.IntegerValue
                                    Dim minId = Math.Min(id1, id2)
                                    Dim maxId = Math.Max(id1, id2)
                                    pairKey = $"{minId}-{maxId}-{connType}-{originKey}"
                                Else
                                    pairKey = $"{baseId}-none-{connType}-{originKey}"
                                End If

                                If seenPairs.Add(pairKey) Then
                                    Dim row = BuildRow(el, found, distInch, connType, param, v1, v2,
                                                       issueStatus, paramCompare,
                                                       normalizedExtras, extras1, extras2, fileLabel)
                                    rows.Add(row)
                                End If
                            End If

                            ' progress
                            If progress IsNot Nothing Then
                                Dim baseFrac As Double = CDbl(i) / CDbl(totalElem)
                                Dim withinFrac As Double = (CDbl(j) / CDbl(connTotal)) / CDbl(totalElem)
                                Dim overall As Double = baseFrac + withinFrac
                                Dim pct As Double = Math.Round(overall * 1000.0R) / 10.0R

                                If (i < totalElem - 1) OrElse (j < connTotal) Then
                                    If pct >= 100.0R Then pct = 99.9R
                                End If

                                If pct >= lastSentPct + 0.1R OrElse (i = totalElem - 1 AndAlso j = connTotal) Then
                                    lastSentPct = pct
                                    progress(pct, $"커넥터 진단 중... ({i + 1}/{totalElem})  커넥터 {j}/{connTotal}")
                                End If
                            End If

                        Catch ex As Exception
                            Dim originText As String = ""
                            Try
                                originText = $"{Math.Round(c.Origin.X, 4)},{Math.Round(c.Origin.Y, 4)},{Math.Round(c.Origin.Z, 4)}"
                            Catch
                            End Try

                            Log($"오류: DOC={fileLabel}, Id={baseId}, Origin={originText}, {ex.Message}")

                            Dim errRow As New Dictionary(Of String, Object)(StringComparer.Ordinal) From {
                                {"File", fileLabel},
                                {"Id1", baseId.ToString()},
                                {"Id2", ""},
                                {"Category1", SafeCategoryName(el)},
                                {"Category2", ""},
                                {"Family1", GetFamilyName(el)},
                                {"Family2", ""},
                                {"Distance (inch)", ""},
                                {"ConnectionType", "ERROR"},
                                {"ParamName", param},
                                {"Value1", ""},
                                {"Value2", ""},
                                {"ParamCompare", "N/A"},
                                {"Status", "ERROR"},
                                {"ErrorMessage", ex.Message}
                            }

                            ' ✅ extra 컬럼 스키마 유지
                            If normalizedExtras IsNot Nothing Then
                                For Each name In normalizedExtras
                                    errRow($"{name}(ID1)") = ""
                                    errRow($"{name}(ID2)") = ""
                                Next
                            End If

                            rows.Add(errRow)
                        End Try
                    Next
                Next

                If progress IsNot Nothing Then
                    progress(100.0R, "완료")
                End If

                ' 정렬: Distance 빈값/에러는 아래로 (ToDouble에서 MaxValue 처리)
                rows = rows.OrderBy(Function(r) ToDouble(GetDictValue(r, "Distance (inch)"))) _
                           .ThenBy(Function(r) ToInt(GetDictValue(r, "Id1"))) _
                           .ThenBy(Function(r) ToInt(GetDictValue(r, "Id2"))) _
                           .ToList()

                If rows.Count > 0 Then
                    Dim s = rows(0)
                    Log($"샘플: Id1={GetDictValue(s, "Id1")}, Id2={GetDictValue(s, "Id2")}, d(in)={GetDictValue(s, "Distance (inch)")}, type={GetDictValue(s, "ConnectionType")}, v1='{GetDictValue(s, "Value1")}', v2='{GetDictValue(s, "Value2")}', status={GetDictValue(s, "Status")}")
                Else
                    Log("최종 rows=0 (근접도/연결 모두 해당 없음)")
                End If

            Catch ex As Exception
                Log($"RunCore 실패: {ex.Message}")

                Dim fatal As New Dictionary(Of String, Object)(StringComparer.Ordinal) From {
                    {"File", fileLabel},
                    {"Id1", ""},
                    {"Id2", ""},
                    {"Category1", ""},
                    {"Category2", ""},
                    {"Family1", ""},
                    {"Family2", ""},
                    {"Distance (inch)", ""},
                    {"ConnectionType", "ERROR"},
                    {"ParamName", param},
                    {"Value1", ""},
                    {"Value2", ""},
                    {"ParamCompare", "N/A"},
                    {"Status", "ERROR"},
                    {"ErrorMessage", ex.Message}
                }

                If normalizedExtras IsNot Nothing Then
                    For Each name In normalizedExtras
                        fatal($"{name}(ID1)") = ""
                        fatal($"{name}(ID2)") = ""
                    Next
                End If

                rows.Add(fatal)
            End Try

            Return rows
        End Function

        ' ============================
        ' Row builder
        ' ============================

        Private Shared Function BuildRow(e1 As Element,
                                         e2 As Element,
                                         distInch As Double,
                                         connType As String,
                                         param As String,
                                         v1 As String,
                                         v2 As String,
                                         status As String,
                                         paramCompare As String,
                                         extraNames As IList(Of String),
                                         extraVals1 As Dictionary(Of String, String),
                                         extraVals2 As Dictionary(Of String, String),
                                         fileLabel As String) As Dictionary(Of String, Object)

            Dim cat1 As String = If(e1?.Category Is Nothing, "", e1.Category.Name)
            Dim cat2 As String = If(e2?.Category Is Nothing, "", e2.Category.Name)
            Dim fam1 As String = GetFamilyName(e1)
            Dim fam2 As String = GetFamilyName(e2)

            Dim row As New Dictionary(Of String, Object)(StringComparer.Ordinal) From {
                {"File", fileLabel},
                {"Id1", If(e1 IsNot Nothing, e1.Id.IntegerValue.ToString(), "0")},
                {"Id2", If(e2 IsNot Nothing, e2.Id.IntegerValue.ToString(), "")},
                {"Category1", cat1},
                {"Category2", cat2},
                {"Family1", fam1},
                {"Family2", fam2},
                {"Distance (inch)", distInch},
                {"ConnectionType", connType},
                {"ParamName", param},
                {"Value1", v1},
                {"Value2", v2},
                {"ParamCompare", paramCompare},
                {"Status", status},
                {"ErrorMessage", ""}
            }

            If extraNames IsNot Nothing Then
                For Each name In extraNames
                    Dim vId1 As String = ""
                    Dim vId2 As String = ""
                    If extraVals1 IsNot Nothing AndAlso extraVals1.ContainsKey(name) Then vId1 = extraVals1(name)
                    If extraVals2 IsNot Nothing AndAlso extraVals2.ContainsKey(name) Then vId2 = extraVals2(name)
                    row($"{name}(ID1)") = vId1
                    row($"{name}(ID2)") = vId2
                Next
            End If

            Return row
        End Function

        ' ============================
        ' Collect / Connector Utilities
        ' ============================

        Private Shared Function CollectElementsWithConnectors(doc As Document, filter As TargetFilter, excludeEndDummy As Boolean) As List(Of Element)
            Dim elems As New List(Of Element)()

            ' FamilyInstance (MEPModel)
            For Each fi As FamilyInstance In New FilteredElementCollector(doc).OfClass(GetType(FamilyInstance))
                Try
                    If fi.MEPModel IsNot Nothing AndAlso
                       fi.MEPModel.ConnectorManager IsNot Nothing AndAlso
                       fi.MEPModel.ConnectorManager.Connectors IsNot Nothing AndAlso
                       fi.MEPModel.ConnectorManager.Connectors.Cast(Of Connector)().Any() Then

                        If IsElementAllowed(fi, filter, excludeEndDummy) Then elems.Add(fi)
                    End If
                Catch
                End Try
            Next

            ' Curves/Fittings/etc
            Dim cats = New BuiltInCategory() {
                BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_DuctAccessory
            }

            For Each cat In cats
                For Each el As Element In New FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()
                    If HasConnectors(el) AndAlso IsElementAllowed(el, filter, excludeEndDummy) Then elems.Add(el)
                Next
            Next

            Return elems.
                GroupBy(Function(e) e.Id.IntegerValue).
                Select(Function(g) g.First()).
                ToList()
        End Function

        Private Shared Function HasConnectors(el As Element) As Boolean
            Try
                Dim fi = TryCast(el, FamilyInstance)
                If fi?.MEPModel IsNot Nothing AndAlso fi.MEPModel.ConnectorManager?.Connectors IsNot Nothing Then
                    Return fi.MEPModel.ConnectorManager.Connectors.Cast(Of Connector)().Any()
                End If

                Dim mc = TryCast(el, MEPCurve)
                If mc?.ConnectorManager?.Connectors IsNot Nothing Then
                    Return mc.ConnectorManager.Connectors.Cast(Of Connector)().Any()
                End If
            Catch
            End Try
            Return False
        End Function

        Private Shared Function GetConnectors(el As Element) As List(Of Connector)
            Try
                Dim fi = TryCast(el, FamilyInstance)
                If fi?.MEPModel IsNot Nothing AndAlso fi.MEPModel.ConnectorManager IsNot Nothing Then
                    Return fi.MEPModel.ConnectorManager.Connectors.Cast(Of Connector)().ToList()
                End If

                Dim mc = TryCast(el, MEPCurve)
                If mc?.ConnectorManager IsNot Nothing Then
                    Return mc.ConnectorManager.Connectors.Cast(Of Connector)().ToList()
                End If
            Catch
            End Try
            Return New List(Of Connector)()
        End Function

        Private Shared Function GetFamilyName(e As Element) As String
            Try
                If TypeOf e Is FamilyInstance Then
                    Dim fi = DirectCast(e, FamilyInstance)
                    If fi.Symbol IsNot Nothing AndAlso fi.Symbol.Family IsNot Nothing Then
                        Return fi.Symbol.Family.Name
                    End If
                ElseIf e IsNot Nothing Then
                    Dim et = TryCast(e.Document.GetElement(e.GetTypeId()), ElementType)
                    If et IsNot Nothing Then
                        Return et.FamilyName
                    End If
                End If
            Catch
            End Try
            Return ""
        End Function

        Private Shared Function SafeCategoryName(el As Element) As String
            Try
                If el Is Nothing OrElse el.Category Is Nothing Then Return ""
                Return el.Category.Name
            Catch
                Return ""
            End Try
        End Function

        ' ============================
        ' Param / Extra Values
        ' ============================

        Private Shared Function GetParamInfo(el As Element, name As String) As ParamInfo
            Dim info As New ParamInfo() With {.Exists = False, .HasValue = False, .Text = ""}

            If el Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then
                Return info
            End If

            Dim p As Parameter = Nothing
            Try
                p = el.LookupParameter(name)
            Catch
            End Try

            If p Is Nothing Then
                Return info
            End If

            info.Exists = True
            info.HasValue = p.HasValue
            info.Text = ResolveParamText(p)
            Return info
        End Function

        Private Shared Function ResolveParamText(el As Element, name As String) As String
            If el Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return ""

            Dim p As Parameter = Nothing
            Try
                p = el.LookupParameter(name)
            Catch
            End Try

            If p Is Nothing Then Return ""
            Return ResolveParamText(p)
        End Function

        Private Shared Function ResolveParamText(p As Parameter) As String
            If p Is Nothing OrElse Not p.HasValue Then Return ""
            Dim raw As String = Nothing
            Try
                If p.StorageType = StorageType.[String] Then
                    raw = p.AsString()
                Else
                    raw = p.AsValueString()
                    If String.IsNullOrWhiteSpace(raw) Then raw = p.AsString()
                End If
            Catch
            End Try

            If raw Is Nothing Then raw = ""
            Return raw.Trim()
        End Function

        Private Shared Function GetExtraValues(el As Element,
                                              names As IList(Of String),
                                              cache As Dictionary(Of Integer, Dictionary(Of String, String))) As Dictionary(Of String, String)

            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If names Is Nothing OrElse names.Count = 0 Then Return result

            If el Is Nothing Then
                For Each n In names
                    result(n) = ""
                Next
                Return result
            End If

            Dim id = el.Id.IntegerValue
            Dim perElem As Dictionary(Of String, String) = Nothing
            If Not cache.TryGetValue(id, perElem) Then
                perElem = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                cache(id) = perElem
            End If

            For Each n In names
                If perElem.ContainsKey(n) Then
                    result(n) = perElem(n)
                Else
                    Dim text = ResolveParamText(el, n)
                    perElem(n) = text
                    result(n) = text
                End If
            Next

            Return result
        End Function

        ' ============================
        ' Proximity Grid
        ' ============================

        Private Shared Function BuildGrid(items As List(Of Tuple(Of Integer, XYZ, Connector))) As Dictionary(Of Tuple(Of Integer, Integer, Integer), List(Of Tuple(Of Integer, XYZ, Connector)))
            Dim grid As New Dictionary(Of Tuple(Of Integer, Integer, Integer), List(Of Tuple(Of Integer, XYZ, Connector)))()
            For Each tup In items
                Dim key = BucketKey(tup.Item2)
                If Not grid.ContainsKey(key) Then
                    grid(key) = New List(Of Tuple(Of Integer, XYZ, Connector))()
                End If
                grid(key).Add(tup)
            Next
            Return grid
        End Function

        Private Shared Function FindProximityCandidate(c As Connector,
                                                       buckets As Dictionary(Of Tuple(Of Integer, Integer, Integer), List(Of Tuple(Of Integer, XYZ, Connector))),
                                                       allowedIds As HashSet(Of Integer),
                                                       baseId As Integer,
                                                       tolFt As Double,
                                                       requireTypeMatch As Boolean) As Tuple(Of Integer, Double)

            If c Is Nothing OrElse buckets Is Nothing OrElse allowedIds Is Nothing Then
                Return Tuple.Create(0, 0.0)
            End If

            Dim org As XYZ = Nothing
            Try
                org = c.Origin
            Catch
                Return Tuple.Create(0, 0.0)
            End Try
            If org Is Nothing Then Return Tuple.Create(0, 0.0)

            Dim key = BucketKey(org)
            Dim bestOtherId As Integer = 0
            Dim bestDistFt As Double = 0.0

            For dx = -1 To 1
                For dy = -1 To 1
                    For dz = -1 To 1
                        Dim nbKey = Tuple.Create(key.Item1 + dx, key.Item2 + dy, key.Item3 + dz)
                        If Not buckets.ContainsKey(nbKey) Then Continue For

                        For Each nb In buckets(nbKey)
                            Dim otherId = nb.Item1
                            If otherId = 0 Then Continue For
                            If otherId = baseId Then Continue For
                            If Not allowedIds.Contains(otherId) Then Continue For

                            ' Domain 은 항상 맞춰주고
                            If c.Domain <> nb.Item3.Domain Then Continue For

                            ' 타입 매칭 요구 시 ConnectorType도 체크
                            If requireTypeMatch AndAlso c.ConnectorType <> nb.Item3.ConnectorType Then Continue For

                            Dim d As Double
                            Try
                                d = org.DistanceTo(nb.Item2)
                            Catch
                                Continue For
                            End Try

                            If d > tolFt Then Continue For

                            If bestOtherId = 0 OrElse d < bestDistFt Then
                                bestOtherId = otherId
                                bestDistFt = d
                            End If
                        Next
                    Next
                Next
            Next

            Return Tuple.Create(bestOtherId, bestDistFt)
        End Function

        Private Shared Function BucketKey(p As XYZ) As Tuple(Of Integer, Integer, Integer)
            Return Tuple.Create(CInt(Math.Floor(p.X)), CInt(Math.Floor(p.Y)), CInt(Math.Floor(p.Z)))
        End Function

        ' ============================
        ' Converters / Helpers
        ' ============================

        Private Shared Function ToTolFt(tol As Double, unit As String) As Double
            If String.Equals(unit, "mm", StringComparison.OrdinalIgnoreCase) Then
                Return tol / 304.8
            End If
            If String.Equals(unit, "inch", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(unit, "in", StringComparison.OrdinalIgnoreCase) Then
                Return tol / 12.0
            End If
            Return tol
        End Function

        Private Shared Function BuildFileLabel(doc As Document) As String
            If doc Is Nothing Then Return String.Empty
            If Not String.IsNullOrWhiteSpace(doc.PathName) Then
                Return Path.GetFileName(doc.PathName)
            End If
            Return doc.Title
        End Function

        Private Shared Function NormalizeExtraParams(extraParams As IEnumerable(Of String)) As List(Of String)
            Dim result As New List(Of String)()
            If extraParams Is Nothing Then Return result

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each raw In extraParams
                Dim name = If(raw, "")
                name = name.Trim()
                If String.IsNullOrEmpty(name) Then Continue For
                If seen.Add(name) Then result.Add(name)
            Next
            Return result
        End Function

        Private Shared Function ToDouble(o As Object) As Double
            Try
                If o Is Nothing Then Return 0.0

                Dim s = TryCast(o, String)
                If s IsNot Nothing Then
                    s = s.Trim()
                    If s = "" Then Return Double.MaxValue ' ✅ 빈 값(에러 등)은 맨 아래로
                End If

                Return Convert.ToDouble(o)
            Catch
                Return Double.MaxValue
            End Try
        End Function

        Private Shared Function ToInt(o As Object) As Integer
            Try
                If o Is Nothing Then Return 0
                Dim s = o.ToString().Trim()
                If String.IsNullOrEmpty(s) Then Return 0
                Return Convert.ToInt32(s)
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function GetDictValue(dict As Dictionary(Of String, Object), key As String) As Object
            If dict Is Nothing Then Return Nothing
            Dim v As Object = Nothing
            If dict.TryGetValue(key, v) Then Return v
            Return Nothing
        End Function

        ' ============================
        ' Filter Parser
        ' ============================

        Private Class FilterToken
            Public Property Kind As String
            Public Property Text As String
        End Class

        Private Class FilterParser
            Private ReadOnly _tokens As List(Of FilterToken)
            Private _pos As Integer = 0
            Public Property FirstParam As String = String.Empty

            Public Sub New(raw As String)
                _tokens = Tokenize(raw)
            End Sub

            ' ✅ 최상위에서 "A=1, B=2" 혹은 "A=1;B=2" 입력을 AND로 처리
            Public Function Parse() As Func(Of Element, Boolean)
                If _tokens.Count = 0 Then Return Nothing

                Dim first = ParseExpr()
                If first Is Nothing Then Return Nothing

                Dim exprList As New List(Of Func(Of Element, Boolean)) From {first}

                While PeekIs("comma")
                    [Next]()
                    Dim nextExpr = ParseExpr()
                    If nextExpr IsNot Nothing Then
                        exprList.Add(nextExpr)
                    Else
                        Exit While
                    End If
                End While

                If exprList.Count = 1 Then
                    Return Function(el As Element) first(el)
                End If

                Return Function(el As Element)
                           For Each part In exprList
                               If part IsNot Nothing AndAlso Not part(el) Then Return False
                           Next
                           Return True
                       End Function
            End Function

            Private Function ParseExpr() As Func(Of Element, Boolean)
                If AtEnd() Then Return Nothing
                Dim tok = Peek()
                If tok Is Nothing Then Return Nothing

                If tok.Kind = "ident" Then
                    If NextIs("lparen", 1) Then
                        Return ParseFunc()
                    End If
                    Return ParseComparison()
                End If

                Return Nothing
            End Function

            Private Function ParseFunc() As Func(Of Element, Boolean)
                Dim nameTok = Expect("ident")
                If nameTok Is Nothing Then Return Nothing

                Dim funcName = nameTok.Text.ToLowerInvariant()
                Expect("lparen")

                Dim args As New List(Of Func(Of Element, Boolean))()

                While Not AtEnd()
                    If PeekIs("rparen") Then Exit While

                    Dim arg = ParseExpr()
                    If arg Is Nothing Then Exit While
                    args.Add(arg)

                    If PeekIs("comma") Then
                        [Next]()
                    ElseIf PeekIs("rparen") Then
                        Exit While
                    ElseIf PeekIs("ident") OrElse PeekIs("lparen") Then
                        ' 허용: or(and(...)and(...)) 같이 콤마 생략된 경우
                        Continue While
                    Else
                        Exit While
                    End If
                End While

                Expect("rparen")

                Select Case funcName
                    Case "and"
                        Return Function(el As Element)
                                   For Each a In args
                                       If a IsNot Nothing AndAlso Not a(el) Then Return False
                                   Next
                                   Return True
                               End Function

                    Case "or"
                        Return Function(el As Element)
                                   For Each a In args
                                       If a IsNot Nothing AndAlso a(el) Then Return True
                                   Next
                                   Return False
                               End Function

                    Case "not"
                        Dim inner As Func(Of Element, Boolean) = If(args.Count > 0, args(0), Nothing)
                        Return Function(el As Element)
                                   If inner Is Nothing Then Return True
                                   Return Not inner(el)
                               End Function

                    Case Else
                        Return Nothing
                End Select
            End Function

            Private Function ParseComparison() As Func(Of Element, Boolean)
                Dim left = Expect("ident")
                If left Is Nothing Then Return Nothing

                Expect("eq")

                Dim right = ExpectValue()
                If right Is Nothing Then Return Nothing

                If String.IsNullOrEmpty(FirstParam) Then FirstParam = left.Text

                Dim expected As String = right.Text
                Dim paramName As String = left.Text

                Return Function(el As Element)
                           Dim actual As String = ResolveParamText(el, paramName)
                           Return String.Equals(actual.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase)
                       End Function
            End Function

            Private Function Expect(kind As String) As FilterToken
                If PeekIs(kind) Then Return [Next]()
                Return Nothing
            End Function

            Private Function ExpectValue() As FilterToken
                If PeekIs("string") OrElse PeekIs("ident") Then Return [Next]()
                Return Nothing
            End Function

            Private Function Peek() As FilterToken
                If _pos >= _tokens.Count Then Return Nothing
                Return _tokens(_pos)
            End Function

            Private Function PeekIs(kind As String, Optional offset As Integer = 0) As Boolean
                Dim idx = _pos + offset
                If idx < 0 OrElse idx >= _tokens.Count Then Return False
                Return String.Equals(_tokens(idx).Kind, kind, StringComparison.OrdinalIgnoreCase)
            End Function

            Private Function NextIs(kind As String, Optional offset As Integer = 0) As Boolean
                Return PeekIs(kind, offset)
            End Function

            Private Function [Next]() As FilterToken
                Dim t = Peek()
                _pos += 1
                Return t
            End Function

            Private Function AtEnd() As Boolean
                Return _pos >= _tokens.Count
            End Function

            Private Shared Function Tokenize(raw As String) As List(Of FilterToken)
                Dim list As New List(Of FilterToken)()
                If String.IsNullOrWhiteSpace(raw) Then Return list

                Dim i As Integer = 0
                While i < raw.Length
                    Dim ch = raw(i)

                    If Char.IsWhiteSpace(ch) Then
                        i += 1
                        Continue While
                    End If

                    If ch = "("c Then
                        list.Add(New FilterToken With {.Kind = "lparen", .Text = "("})
                        i += 1
                        Continue While
                    End If
                    If ch = ")"c Then
                        list.Add(New FilterToken With {.Kind = "rparen", .Text = ")"})
                        i += 1
                        Continue While
                    End If
                    If ch = ","c OrElse ch = ";"c Then
                        list.Add(New FilterToken With {.Kind = "comma", .Text = ","})
                        i += 1
                        Continue While
                    End If
                    If ch = "="c Then
                        list.Add(New FilterToken With {.Kind = "eq", .Text = "="})
                        i += 1
                        Continue While
                    End If

                    If ch = "'"c OrElse ch = """"c Then
                        Dim quoteCh As Char = ch
                        i += 1
                        Dim start = i
                        While i < raw.Length AndAlso raw(i) <> quoteCh
                            i += 1
                        End While

                        Dim content As String = raw.Substring(start, i - start)
                        list.Add(New FilterToken With {.Kind = "string", .Text = content})

                        If i < raw.Length AndAlso raw(i) = quoteCh Then i += 1
                        Continue While
                    End If

                    ' ✅ ident 스캔에서 ; 도 끊어야 함 (여기 핵심)
                    Dim startWord = i
                    While i < raw.Length AndAlso Not Char.IsWhiteSpace(raw(i)) _
                          AndAlso raw(i) <> "("c AndAlso raw(i) <> ")"c _
                          AndAlso raw(i) <> ","c AndAlso raw(i) <> ";"c _
                          AndAlso raw(i) <> "="c
                        i += 1
                    End While

                    Dim word = raw.Substring(startWord, i - startWord)
                    If Not String.IsNullOrEmpty(word) Then
                        list.Add(New FilterToken With {.Kind = "ident", .Text = word})
                    End If
                End While

                Return list
            End Function
        End Class

        Private Shared Function ParseTargetFilter(raw As String) As TargetFilter
            Dim result As New TargetFilter()
            If String.IsNullOrWhiteSpace(raw) Then Return result

            Try
                Dim parser As New FilterParser(raw)
                Dim evaluator = parser.Parse()
                If evaluator Is Nothing Then Return result
                result.Evaluator = evaluator
                result.PrimaryParam = parser.FirstParam
            Catch ex As Exception
                Log($"필터 파싱 실패: {ex.Message}")
            End Try

            Return result
        End Function

        Private Shared Function IsElementAllowed(el As Element, filter As TargetFilter, excludeEndDummy As Boolean) As Boolean
            If el Is Nothing Then Return False

            If excludeEndDummy Then
                Dim fam As String = GetFamilyName(el)
                If fam.IndexOf("End_", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso fam.IndexOf("Dummy", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return False
                End If
            End If

            If filter Is Nothing OrElse filter.Evaluator Is Nothing Then Return True
            Return filter.Evaluator(el)
        End Function

    End Class

End Namespace
