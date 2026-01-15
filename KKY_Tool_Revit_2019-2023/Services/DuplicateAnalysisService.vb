Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.DB.Mechanical
Imports Autodesk.Revit.DB.Plumbing
Imports Autodesk.Revit.DB.Electrical
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Infrastructure   ' ElementIdCompat.IntValue / FromInt 사용


Namespace Services

    ''' <summary>
    ''' 중복검토 데이터 수집 서비스 (fix2 규격)
    ''' - 그룹 키: Category + Family + Type + 위치(패밀리) / 시작·끝(커브)
    ''' - 연결수(실 연결): Connector의 AllRefs에서 자기 자신 제외, 소유 ElementId 고유 개수
    ''' - 연결 객체: 연결된 소유 ElementId들을 ','로 연결한 문자열(중복 제거)
    ''' - candidate: 그룹 내 개수가 2개 이상이면 True
    ''' </summary>
    Public Class DuplicateAnalysisService

        ' ====== 공개 엔드포인트 ======

        ''' <summary>
        ''' 허브에서 호출하는 메인 수집 함수.
        ''' 반환: List(Of Dictionary) — fix2 스키마 필드명
        '''   - groupId, id, category, family, type, connectedCount, connectedIds, candidate
        ''' </summary>
        Public Shared Function Run(app As UIApplication) As List(Of Dictionary(Of String, Object))
            Dim uidoc = app.ActiveUIDocument
            If uidoc Is Nothing OrElse uidoc.Document Is Nothing Then
                Return New List(Of Dictionary(Of String, Object))()
            End If
            Dim doc = uidoc.Document

            ' 대상: 커넥터가 있을 법한 주요 MEP 요소 + 기계/전기 패밀리(패밀리 인스턴스)
            Dim elems As New List(Of Element)()

            elems.AddRange(New FilteredElementCollector(doc) _
                           .OfClass(GetType(MEPCurve)) _
                           .WhereElementIsNotElementType() _
                           .ToElements())

            elems.AddRange(New FilteredElementCollector(doc) _
                           .OfClass(GetType(FamilyInstance)) _
                           .WhereElementIsNotElementType() _
                           .ToElements() _
                           .Where(Function(fi) HasAnyConnector(fi)))

            ' 그룹핑 키 생성
            Dim groups As New Dictionary(Of String, List(Of Element))(StringComparer.Ordinal)
            For Each e In elems
                Dim key As String = BuildGroupKey(e)
                If Not groups.ContainsKey(key) Then groups(key) = New List(Of Element)()
                groups(key).Add(e)
            Next

            ' 결과로 평탄화
            Dim rows As New List(Of Dictionary(Of String, Object))()
            Dim gno As Integer = 1

            For Each kv In groups
                Dim list = kv.Value
                If list Is Nothing OrElse list.Count = 0 Then Continue For

                Dim isCandidate As Boolean = (list.Count >= 2)
                For Each e In list
                    Dim cat = If(e.Category Is Nothing, "", e.Category.Name)
                    Dim fam As String = ""
                    Dim typ As String = ""

                    If TypeOf e Is FamilyInstance Then
                        Dim fi = DirectCast(e, FamilyInstance)
                        typ = If(fi.Symbol IsNot Nothing, fi.Symbol.Name, "")
                        fam = If(fi.Symbol IsNot Nothing AndAlso fi.Symbol.Family IsNot Nothing, fi.Symbol.Family.Name, "")
                    Else
                        ' MEPCurve
                        typ = e.Name
                        Dim es = TryCast(doc.GetElement(e.GetTypeId()), ElementType)
                        If es IsNot Nothing Then
                            fam = es.FamilyName
                        End If
                    End If

                    Dim connected As HashSet(Of Integer) = GetConnectedOwnerIds(e)
                    Dim connectedIdsStr As String = String.Join(",", connected.Select(Function(x) x.ToString()).ToArray())

                    Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                        {"groupId", gno},
                        {"id", e.Id.IntegerValue.ToString()},
                        {"category", cat},
                        {"family", fam},
                        {"type", typ},
                        {"connectedCount", connected.Count},
                        {"connectedIds", connectedIdsStr},
                        {"candidate", isCandidate}
                    }
                    rows.Add(d)
                Next
                gno += 1
            Next

            ' groupId 순/ElementId 순으로 정렬(일관된 출력)
            rows = rows.OrderBy(Function(r) Convert.ToInt32(r("groupId"))) _
                       .ThenBy(Function(r) Convert.ToInt32(r("id"))) _
                       .ToList()

            Return rows
        End Function

        ''' <summary>
        ''' (선택) 서비스 자체에서 엑셀을 저장하려면 True 반환.
        ''' 현재는 허브가 공통 로직으로 저장하도록 False를 반환하여 넘어가게 함.
        ''' </summary>
        Public Shared Function Export(app As UIApplication) As Boolean
            Return False
        End Function

        ''' <summary>
        ''' (선택) 삭제 복원 서비스. 구현 시 true/false 반환.
        ''' 현재는 허브 쪽에서 안내 메시지 후 미구현 처리.
        ''' </summary>
        Public Shared Function Restore(app As UIApplication, payload As Dictionary(Of String, Object)) As Boolean
            Return False
        End Function

        ' ====== 내부 유틸 ======

        Private Shared Function HasAnyConnector(e As Element) As Boolean
            Try
                If TypeOf e Is MEPCurve Then
                    Dim cm = DirectCast(e, MEPCurve).ConnectorManager
                    Return cm IsNot Nothing AndAlso cm.Connectors IsNot Nothing AndAlso cm.Connectors.Size > 0
                End If
                Dim fi = TryCast(e, FamilyInstance)
                If fi IsNot Nothing AndAlso fi.MEPModel IsNot Nothing AndAlso fi.MEPModel.ConnectorManager IsNot Nothing Then
                    Dim cms = fi.MEPModel.ConnectorManager.Connectors
                    Return cms IsNot Nothing AndAlso cms.Size > 0
                End If
            Catch
            End Try
            Return False
        End Function

        ''' <summary>요소의 위치 기반 그룹 키. 방향성/자릿수 차이로 인한 미세 오차 대비하여 반올림 처리.</summary>
        Private Shared Function BuildGroupKey(e As Element) As String
            Dim cat = If(e.Category Is Nothing, "", e.Category.Name)
            Dim fam As String = ""
            Dim typ As String = ""

            Dim doc = e.Document

            If TypeOf e Is FamilyInstance Then
                Dim fi = DirectCast(e, FamilyInstance)
                typ = If(fi.Symbol IsNot Nothing, fi.Symbol.Name, "")
                fam = If(fi.Symbol IsNot Nothing AndAlso fi.Symbol.Family IsNot Nothing, fi.Symbol.Family.Name, "")
                Dim lp = TryCast(e.Location, LocationPoint)
                If lp Is Nothing OrElse lp.Point Is Nothing Then
                    Return $"{cat}|{fam}|{typ}|NOLOC"
                End If
                Dim p = lp.Point
                Dim keyp = $"{R2(p.X)}_{R2(p.Y)}_{R2(p.Z)}"
                Return $"{cat}|{fam}|{typ}|{keyp}"
            Else
                ' MEPCurve
                typ = e.Name
                Dim es = TryCast(doc.GetElement(e.GetTypeId()), ElementType)
                If es IsNot Nothing Then fam = es.FamilyName

                Dim lc = TryCast(e.Location, LocationCurve)
                If lc Is Nothing OrElse lc.Curve Is Nothing Then
                    Return $"{cat}|{fam}|{typ}|NOCURVE"
                End If
                Dim c = lc.Curve
                Dim s = c.GetEndPoint(0)
                Dim t = c.GetEndPoint(1)
                ' 방향성 제거 위해 소트
                Dim k1 = $"{R2(s.X)}_{R2(s.Y)}_{R2(s.Z)}"
                Dim k2 = $"{R2(t.X)}_{R2(t.Y)}_{R2(t.Z)}"
                Dim a = New String() {k1, k2}
                Array.Sort(a, StringComparer.Ordinal)
                Return $"{cat}|{fam}|{typ}|{a(0)}|{a(1)}"
            End If
        End Function

        ''' <summary>피트 단위를 소수점 4자리까지 반올림한 문자열</summary>
        Private Shared Function R2(v As Double) As String
            ' 1e-4 ft ≈ 0.3048 mm — 중복 판단엔 충분
            Return Math.Round(v, 4, MidpointRounding.AwayFromZero).ToString("0.####")
        End Function

        ''' <summary>요소의 커넥터 기준 연결 상대 ElementId 모음(자기 자신 제외, 중복 제거)</summary>
        Private Shared Function GetConnectedOwnerIds(e As Element) As HashSet(Of Integer)
            Dim setIds As New HashSet(Of Integer)()

            Try
                Dim cm As ConnectorManager = Nothing

                If TypeOf e Is MEPCurve Then
                    cm = DirectCast(e, MEPCurve).ConnectorManager
                Else
                    Dim fi = TryCast(e, FamilyInstance)
                    If fi IsNot Nothing AndAlso fi.MEPModel IsNot Nothing Then
                        cm = fi.MEPModel.ConnectorManager
                    End If
                End If

                If cm Is Nothing OrElse cm.Connectors Is Nothing OrElse cm.Connectors.Size = 0 Then
                    Return setIds
                End If

                For Each o In cm.Connectors
                    Dim c As Connector = TryCast(o, Connector)
                    If c Is Nothing Then Continue For
                    Dim refs = c.AllRefs
                    If refs Is Nothing OrElse refs.Size = 0 Then Continue For
                    For Each ro In refs
                        Dim rc As Connector = TryCast(ro, Connector)
                        If rc Is Nothing OrElse rc.Owner Is Nothing Then Continue For
                        Dim oid As Integer = rc.Owner.Id.IntegerValue
                        If oid <> e.Id.IntegerValue Then setIds.Add(oid)
                    Next
                Next
            Catch
            End Try

            Return setIds
        End Function

    End Class

End Namespace
