Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Text.Encodings.Web
Imports System.Text.Json
Imports System.Text.Json.Nodes
Imports System.Text.Json.Serialization

' RootNamespace 영향 안 받게 Global 붙여줌
Namespace Global.System.Web.Script.Serialization

    ''' <summary>
    ''' Minimal shim for JavaScriptSerializer to keep legacy code intact on .NET 8.
    ''' Uses System.Text.Json under the hood while returning Dictionary/List primitives
    ''' compatible with the existing payload handling code.
    ''' </summary>
    Public Class JavaScriptSerializer

        Private Shared ReadOnly _options As New JsonSerializerOptions With {
            .PropertyNameCaseInsensitive = True,
            .NumberHandling = JsonNumberHandling.AllowReadingFromString,
            .Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        }

        Public Function Deserialize(Of T)(input As String) As T
            If String.IsNullOrWhiteSpace(input) Then
                Return Nothing
            End If

            If GetType(T) Is GetType(Dictionary(Of String, Object)) Then
                Dim root As JsonNode = JsonNode.Parse(input)
                If root Is Nothing Then
                    Return Nothing
                End If

                Dim converted = TryCast(ConvertNode(root), Dictionary(Of String, Object))
                If converted Is Nothing Then
                    Return Nothing
                End If

                Return CType(CType(converted, Object), T)
            End If

            Return JsonSerializer.Deserialize(Of T)(input, _options)
        End Function

        Public Function Serialize(obj As Object) As String
            Return JsonSerializer.Serialize(obj, _options)
        End Function

        Private Shared Function ConvertNode(node As JsonNode) As Object
            If node Is Nothing Then Return Nothing

            Select Case node.GetType()
                Case GetType(JsonValue)
                    Return ConvertValue(CType(node, JsonValue))
                Case GetType(JsonObject)
                    Dim dict As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                    For Each kv In CType(node, JsonObject)
                        dict(kv.Key) = ConvertNode(kv.Value)
                    Next
                    Return dict
                Case GetType(JsonArray)
                    Dim list As New List(Of Object)()
                    For Each child In CType(node, JsonArray)
                        list.Add(ConvertNode(child))
                    Next
                    Return list
                Case Else
                    Return node.ToJsonString()
            End Select
        End Function

        Private Shared Function ConvertValue(value As JsonValue) As Object
            Dim je As JsonElement
            If value.TryGetValue(je) Then
                Select Case je.ValueKind
                    Case JsonValueKind.String
                        Return je.GetString()
                    Case JsonValueKind.Number
                        Dim l As Long
                        If je.TryGetInt64(l) Then Return l
                        Dim d As Double
                        If je.TryGetDouble(d) Then Return d
                        Return je.GetRawText()
                    Case JsonValueKind.True
                        Return True
                    Case JsonValueKind.False
                        Return False
                    Case JsonValueKind.Null, JsonValueKind.Undefined
                        Return Nothing
                    Case Else
                        Return je.GetRawText()
                End Select
            End If

            Dim s As String = Nothing
            If value.TryGetValue(s) Then Return s
            Dim b As Boolean
            If value.TryGetValue(b) Then Return b
            Dim n As Double
            If value.TryGetValue(n) Then Return n

            Return value.ToJsonString()
        End Function

    End Class

End Namespace
