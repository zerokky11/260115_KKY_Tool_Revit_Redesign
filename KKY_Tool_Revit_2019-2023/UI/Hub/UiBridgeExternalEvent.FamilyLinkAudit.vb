Option Explicit On
Option Strict On

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Exports
Imports KKY_Tool_Revit.Services

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private _familyLinkLastRows As List(Of FamilyLinkAuditRow) = Nothing

        Private Sub HandleFamilyLinkInit(app As UIApplication)
            Try
                Dim sourcePath As String = ""
                Try
                    sourcePath = app.Application.SharedParametersFilename
                Catch
                End Try

                If String.IsNullOrWhiteSpace(sourcePath) Then
                    SendToWeb("familylink:error", New With {
                        .message = "Shared Parameters 파일이 설정되어 있지 않습니다.",
                        .detail = "Revit 옵션에서 Shared Parameters 파일 경로를 설정하세요."
                    })
                    Return
                End If

                Dim defFile As DefinitionFile = Nothing
                Try
                    defFile = app.Application.OpenSharedParameterFile()
                Catch ex As Exception
                    SendToWeb("familylink:error", New With {
                        .message = "Shared Parameters 파일을 열 수 없습니다.",
                        .detail = ex.Message
                    })
                    Return
                End Try

                If defFile Is Nothing Then
                    SendToWeb("familylink:error", New With {
                        .message = "Shared Parameters 파일을 읽지 못했습니다.",
                        .detail = sourcePath
                    })
                    Return
                End If

                Dim items As New List(Of Object)()
                For Each g As DefinitionGroup In defFile.Groups
                    If g Is Nothing Then Continue For
                    For Each def As Definition In g.Definitions
                        Dim ext As ExternalDefinition = TryCast(def, ExternalDefinition)
                        If ext Is Nothing Then Continue For
                        items.Add(New With {
                            .name = ext.Name,
                            .guid = ext.GUID.ToString("D"),
                            .groupName = g.Name,
                            .dataTypeToken = SafeDefTypeToken(ext)
                        })
                    Next
                Next

                SendToWeb("familylink:sharedparams", New With {
                    .sourcePath = sourcePath,
                    .items = items
                })

            Catch ex As Exception
                SendToWeb("familylink:error", New With {
                    .message = "Shared Parameters 목록 로드 실패",
                    .detail = ex.Message
                })
            End Try
        End Sub

        Private Sub HandleFamilyLinkPickRvts()
            Using dlg As New OpenFileDialog()
                dlg.Filter = "Revit Project (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.Title = "패밀리 연동 검토 대상 RVT 선택"
                dlg.RestoreDirectory = True

                If dlg.ShowDialog() <> DialogResult.OK Then Return

                Dim files As New List(Of String)()
                Dim dedup As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each p In dlg.FileNames
                    If String.IsNullOrWhiteSpace(p) Then Continue For
                    If dedup.Add(p) Then files.Add(p)
                Next

                SendToWeb("familylink:rvts-picked", New With {.paths = files})
            End Using
        End Sub

        Private Sub HandleFamilyLinkRun(app As UIApplication, payload As Object)
            _familyLinkLastRows = Nothing

            Dim rvtPaths As List(Of String) = ExtractStringList(payload, "rvtPaths")
            Dim targets As List(Of FamilyLinkTargetParam) = ExtractTargets(payload)

            If rvtPaths.Count = 0 Then
                SendToWeb("familylink:error", New With {
                    .message = "검토할 RVT 파일이 없습니다.",
                    .detail = "RVT 목록을 추가하세요."
                })
                Return
            End If

            If targets.Count = 0 Then
                SendToWeb("familylink:error", New With {
                    .message = "검토할 파라미터가 없습니다.",
                    .detail = "Shared Parameters 목록에서 대상 파라미터를 선택하세요."
                })
                Return
            End If

            Try
                Dim rows As List(Of FamilyLinkAuditRow) = FamilyLinkAuditService.Run(app, rvtPaths, targets, AddressOf ReportFamilyLinkProgress)
                _familyLinkLastRows = rows

                Dim schema As String() = FamilyLinkAuditExport.Schema
                Dim payloadRows As List(Of Dictionary(Of String, Object)) = rows.Select(AddressOf ToRowDict).ToList()

                SendToWeb("familylink:result", New With {
                    .schema = schema,
                    .rows = payloadRows
                })

            Catch ex As Exception
                SendToWeb("familylink:error", New With {
                    .message = "패밀리 연동 검토 실행 실패",
                    .detail = ex.Message
                })
            End Try
        End Sub

        Private Sub HandleFamilyLinkExport(payload As Object)
            If _familyLinkLastRows Is Nothing OrElse _familyLinkLastRows.Count = 0 Then
                SendToWeb("familylink:error", New With {
                    .message = "내보낼 결과가 없습니다.",
                    .detail = "먼저 스캔을 실행하세요."
                })
                Return
            End If

            Dim fastExport As Boolean = ExtractBool(payload, "fastExport", True)
            Dim autoFit As Boolean = ExtractBool(payload, "autoFit", False)

            Try
                Dim savedPath As String = FamilyLinkAuditExport.Export(_familyLinkLastRows, fastExport, autoFit)
                If String.IsNullOrWhiteSpace(savedPath) Then
                    SendToWeb("familylink:exported", New With {
                        .ok = False,
                        .message = "엑셀/CSV 내보내기가 취소되었습니다."
                    })
                    Return
                End If

                SendToWeb("familylink:exported", New With {
                    .ok = True,
                    .path = savedPath
                })

            Catch ex As Exception
                SendToWeb("familylink:exported", New With {
                    .ok = False,
                    .message = ex.Message
                })
                SendToWeb("familylink:error", New With {
                    .message = "엑셀/CSV 내보내기 실패",
                    .detail = ex.Message
                })
            End Try
        End Sub

        Private Sub ReportFamilyLinkProgress(percent As Integer, message As String)
            Try
                SendToWeb("familylink:progress", New With {
                    .percent = Math.Max(0, Math.Min(100, percent)),
                    .message = If(message, "")
                })
            Catch
            End Try
        End Sub

        Private Function ExtractStringList(payload As Object, key As String) As List(Of String)
            Dim res As New List(Of String)()

            Dim payloadValue As Object = GetProp(payload, key)
            If payloadValue Is Nothing Then Return res

            ' 문자열은 IEnumerable(문자열 자체가 IEnumerable)로 잡히므로 예외 처리 필요
            Dim payloadItems As System.Collections.IEnumerable = TryCast(payloadValue, System.Collections.IEnumerable)

            If payloadItems Is Nothing OrElse TypeOf payloadValue Is String Then
                Dim singleValue As String = TryCast(payloadValue, String) ' ✅ single(예약어) 금지
                If Not String.IsNullOrWhiteSpace(singleValue) Then
                    res.Add(singleValue) ' ✅ Add(Of String) 같은 잘못된 제네릭 호출 금지
                End If
                Return res
            End If

            For Each o As Object In payloadItems
                If o Is Nothing Then Continue For
                Dim s As String = o.ToString()
                If Not String.IsNullOrWhiteSpace(s) Then
                    res.Add(s)
                End If
            Next

            Return res
        End Function

        Private Function ExtractTargets(payload As Object) As List(Of FamilyLinkTargetParam)
            Dim list As New List(Of FamilyLinkTargetParam)()
            Dim payloadValue As Object = GetProp(payload, "targets")
            Dim payloadItems = TryCast(payloadValue, IEnumerable)
            If payloadItems Is Nothing OrElse TypeOf payloadValue Is String Then Return list

            For Each o In payloadItems
                Dim name As String = TryCast(GetProp(o, "name"), String)
                Dim guidStr As String = TryCast(GetProp(o, "guid"), String)
                If String.IsNullOrWhiteSpace(name) OrElse String.IsNullOrWhiteSpace(guidStr) Then Continue For
                Dim g As Guid
                If Not Guid.TryParse(guidStr, g) Then Continue For

                Dim item As New FamilyLinkTargetParam With {
                    .Name = name.Trim(),
                    .Guid = g
                }
                list.Add(item)
            Next

            Return list
        End Function

        Private Function ToRowDict(row As FamilyLinkAuditRow) As Dictionary(Of String, Object)
            Dim d As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            d("FileName") = row.FileName
            d("HostFamilyName") = row.HostFamilyName
            d("HostFamilyCategory") = row.HostFamilyCategory
            d("NestedFamilyName") = row.NestedFamilyName
            d("NestedTypeName") = row.NestedTypeName
            d("NestedCategory") = row.NestedCategory
            d("TargetParamName") = row.TargetParamName
            d("ExpectedGuid") = row.ExpectedGuid
            d("FoundScope") = row.FoundScope
            d("NestedParamGuid") = row.NestedParamGuid
            d("NestedParamDataType") = row.NestedParamDataType
            d("AssocHostParamName") = row.AssocHostParamName
            d("HostParamGuid") = row.HostParamGuid
            d("HostParamIsShared") = row.HostParamIsShared
            d("Issue") = row.Issue
            d("Notes") = row.Notes
            Return d
        End Function

        Private Function ExtractBool(payload As Object, key As String, defaultValue As Boolean) As Boolean
            Try
                If payload Is Nothing Then Return defaultValue
                Dim raw As Object = GetProp(payload, key)
                If raw Is Nothing Then Return defaultValue
                Return Convert.ToBoolean(raw)
            Catch
                Return defaultValue
            End Try
        End Function

        Private Shared Function SafeDefTypeToken(defn As Definition) As String
            If defn Is Nothing Then Return ""
            Try
#If REVIT2023 = 1 Then
                Dim dt As ForgeTypeId = defn.GetDataType()
                If dt IsNot Nothing Then Return SafeStr(dt.TypeId)
                Return ""
#Else
                Return SafeStr(defn.ParameterType.ToString())
#End If
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SafeStr(s As String) As String
            Return If(s, "")
        End Function

    End Class

End Namespace
