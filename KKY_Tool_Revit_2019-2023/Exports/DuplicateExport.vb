Imports System.Data
Imports System.Linq
Imports KKY_Tool_Revit.Infrastructure

Namespace Exports

    Public Class DupRowDto
        Public Property Id As String
        Public Property Category As String
        Public Property Family As String
        Public Property Type As String
        Public Property ConnectedIds As System.Collections.Generic.List(Of String)
    End Class

    Public Module DuplicateExport

        Public Function Save(rows As System.Collections.IEnumerable, Optional doAutoFit As Boolean = False, Optional progressChannel As String = Nothing) As String
            Dim mapped = MapRows(rows)
            Dim dt = BuildSimpleTable(mapped)
            Return ExcelCore.PickAndSaveXlsx("Duplicates (Simple)", dt, "Duplicates.xlsx", doAutoFit, progressChannel)
        End Function

        Public Sub Save(outPath As String, rows As System.Collections.IEnumerable, Optional doAutoFit As Boolean = False, Optional progressChannel As String = Nothing)
            Export(outPath, rows, doAutoFit, progressChannel)
        End Sub

        Public Sub Export(outPath As String, rows As System.Collections.IEnumerable, Optional doAutoFit As Boolean = False, Optional progressChannel As String = Nothing)
            Dim mapped = MapRows(rows)
            Dim dt = BuildSimpleTable(mapped)
            ExcelCore.SaveStyledSimple(outPath, "Duplicates (Simple)", dt, "Group", doAutoFit, progressChannel)
        End Sub

        Private Function MapRows(rows As System.Collections.IEnumerable) As System.Collections.Generic.List(Of DupRowDto)
            Dim list As New System.Collections.Generic.List(Of DupRowDto)
            If rows Is Nothing Then Return list
            For Each o In rows
                Dim it As New DupRowDto()
                it.Id = ReadProp(o, "Id", "ID", "ElementId", "ElementID", "elementId")
                it.Category = ReadProp(o, "Category", "category")
                it.Family = ReadProp(o, "Family", "family")
                it.Type = ReadProp(o, "Type", "type")
                it.ConnectedIds = ReadList(o, "ConnectedIds", "connectedIds", "Links", "links", "connected", "Connected", "ConnectedElements")
                list.Add(it)
            Next
            Return list
        End Function

        Private Function BuildSimpleTable(rows As System.Collections.Generic.List(Of DupRowDto)) As DataTable
            Dim dt As New DataTable("simple")
            dt.Columns.Add("Group")
            dt.Columns.Add("ID")
            dt.Columns.Add("Category")
            dt.Columns.Add("Family")
            dt.Columns.Add("Type")

            Dim groupList = GroupByLogic(rows)
            For i = 0 To groupList.Count - 1
                Dim gName = $"Group{i + 1}"
                For Each r In groupList(i)
                    Dim famOut As String = If(String.IsNullOrWhiteSpace(r.Family),
                                              If(String.IsNullOrWhiteSpace(r.Category), "", r.Category & " Type"),
                                              r.Family)
                    Dim dr = dt.NewRow()
                    dr("Group") = gName
                    dr("ID") = Nz(r.Id)
                    dr("Category") = Nz(r.Category)
                    dr("Family") = Nz(famOut)
                    dr("Type") = Nz(r.Type)
                    dt.Rows.Add(dr)
                Next
            Next
            Return dt
        End Function

        Private Function GroupByLogic(items As System.Collections.Generic.List(Of DupRowDto)) As System.Collections.Generic.List(Of System.Collections.Generic.List(Of DupRowDto))
            Dim buckets As New System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.List(Of DupRowDto))()

            For Each r In items
                Dim fam As String = If(String.IsNullOrWhiteSpace(r.Family),
                                       If(String.IsNullOrWhiteSpace(r.Category), "", r.Category & " Type"),
                                       r.Family)
                Dim typ As String = If(String.IsNullOrWhiteSpace(r.Type), "", r.Type)
                Dim cat As String = If(String.IsNullOrWhiteSpace(r.Category), "", r.Category)

                Dim clusterSrc As New System.Collections.Generic.List(Of String)
                If Not String.IsNullOrWhiteSpace(r.Id) Then clusterSrc.Add(r.Id)
                If r.ConnectedIds IsNot Nothing Then clusterSrc.AddRange(r.ConnectedIds)

                Dim cluster = clusterSrc _
                    .SelectMany(Function(s) SplitIds(s)) _
                    .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                    .Select(Function(x) x.Trim()) _
                    .Distinct() _
                    .OrderBy(Function(x) PadNum(x)) _
                    .ToList()

                Dim clusterKey As String = If(cluster.Count > 1, String.Join(",", cluster), "")
                Dim key = String.Join("|", {cat, fam, typ, clusterKey})
                If Not buckets.ContainsKey(key) Then buckets(key) = New System.Collections.Generic.List(Of DupRowDto)()
                buckets(key).Add(r)
            Next

            Return buckets.Values.ToList()
        End Function

        Private Function SplitIds(s As String) As System.Collections.Generic.IEnumerable(Of String)
            If String.IsNullOrWhiteSpace(s) Then Return Array.Empty(Of String)()
            Return s.Split(New Char() {","c, " "c, ";"c, "|"c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf},
                           StringSplitOptions.RemoveEmptyEntries)
        End Function

        Private Function PadNum(s As String) As String
            Dim n As Integer
            If Integer.TryParse(s, n) Then Return n.ToString("D10")
            Return s
        End Function

        Private Function Nz(s As String) As String
            If String.IsNullOrWhiteSpace(s) Then Return ""
            Return s
        End Function

        Private Function ReadProp(obj As Object, ParamArray names() As String) As String
            If obj Is Nothing Then Return ""
            For Each nm In names
                If String.IsNullOrEmpty(nm) Then Continue For
                Dim p = obj.GetType().GetProperty(nm)
                If p IsNot Nothing Then
                    Dim v = p.GetValue(obj, Nothing)
                    If v IsNot Nothing Then Return v.ToString()
                End If
            Next
            Return ""
        End Function

        Private Function ReadList(obj As Object, ParamArray names() As String) As System.Collections.Generic.List(Of String)
            Dim res As New System.Collections.Generic.List(Of String)
            If obj Is Nothing Then Return res
            For Each nm In names
                Dim p = obj.GetType().GetProperty(nm)
                If p Is Nothing Then Continue For
                Dim v = p.GetValue(obj, Nothing)
                If v Is Nothing Then Continue For

                If TypeOf v Is String Then
                    res.AddRange(SplitIds(DirectCast(v, String)))
                    Exit For
                End If

                If TypeOf v Is System.Collections.IEnumerable AndAlso Not TypeOf v Is String Then
                    For Each x In DirectCast(v, System.Collections.IEnumerable)
                        If x IsNot Nothing Then res.Add(x.ToString())
                    Next
                    Exit For
                End If
            Next
            Return res
        End Function


    End Module
End Namespace
