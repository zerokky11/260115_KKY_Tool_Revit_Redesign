Option Explicit On
Option Strict On

Imports System
Imports System.Linq
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Services

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Shared _lastParamResult As ParamPropagateService.SharedParamRunResult

        ' === sharedparam:list ===
        Private Sub HandleSharedParamList(app As UIApplication, payload As Object)
            Try
                Dim res = ParamPropagateService.GetSharedParameterDefinitions(app)
                Dim shaped As Object = New With {
                    .ok = res IsNot Nothing AndAlso res.Ok,
                    .message = If(res Is Nothing, Nothing, res.Message),
                    .definitions = If(res?.Definitions, New List(Of ParamPropagateService.SharedParamDefinitionDto)()).Select(Function(d) New With {
                        .groupName = d.GroupName,
                        .name = d.Name,
                        .paramType = d.ParamType,
                        .visible = d.Visible
                    }).ToList(),
                    .targetGroups = If(res?.TargetGroups, New List(Of ParamPropagateService.ParameterGroupOption)()).Select(Function(g) New With {
                        .id = g.Id,
                        .name = g.Name
                    }).ToList()
                }
                SendToWeb("sharedparam:list", shaped)
            Catch ex As Exception
                SendToWeb("sharedparam:list", New With {.ok = False, .message = ex.Message})
                SendToWeb("revit:error", New With {.message = ex.Message})
            End Try
        End Sub

        ' === sharedparam:run / paramprop:run ===
        Private Sub HandleSharedParamRun(app As UIApplication, payload As Object)
            Try
                Dim req As ParamPropagateService.SharedParamRunRequest = ParamPropagateService.SharedParamRunRequest.FromPayload(payload)
                Dim res = ParamPropagateService.Run(app, req, AddressOf ReportParamPropProgress)
                _lastParamResult = res

                Dim status = If(res Is Nothing, ParamPropagateService.RunStatus.Failed, res.Status)
                Dim ok As Boolean = (status = ParamPropagateService.RunStatus.Succeeded)

                Dim responsePayload As New Dictionary(Of String, Object)()
                responsePayload("ok") = ok
                responsePayload("status") = status.ToString().ToLowerInvariant()
                responsePayload("message") = If(res IsNot Nothing, res.Message, Nothing)
                If res IsNot Nothing Then
                    responsePayload("report") = res.Report
                    responsePayload("details") = If(res.Details, New List(Of ParamPropagateService.SharedParamDetailRow)()).Select(Function(d) New With {
                        .kind = d.Kind,
                        .family = d.Family,
                        .detail = d.Detail
                    }).ToList()
                End If

                SendToWeb("sharedparam:done", responsePayload)
                SendToWeb("paramprop:done", responsePayload)

                If Not ok Then
                    Dim msg As String = If(res Is Nothing, "실패", res.Message)
                    SendToWeb("revit:error", New With {.message = "공유 파라미터 연동 실패: " & msg})
                End If

            Catch ex As Exception
                SendToWeb("sharedparam:done", New With {.ok = False, .status = "failed", .message = ex.Message})
                SendToWeb("paramprop:done", New With {.ok = False, .status = "failed", .message = ex.Message})
                SendToWeb("revit:error", New With {.message = "공유 파라미터 연동 실패: " & ex.Message})
            End Try
        End Sub

        Private Sub ReportParamPropProgress(phase As String,
                                            phaseProgress As Double,
                                            current As Integer,
                                            total As Integer,
                                            message As String,
                                            target As String)
            SendToWeb("paramprop:progress", New With {
                .phase = phase,
                .phaseProgress = phaseProgress,
                .current = current,
                .total = total,
                .message = message,
                .target = target
            })
        End Sub

        ' === sharedparam:export-excel ===
        Private Sub HandleSharedParamExport(app As UIApplication, payload As Object)
            Try
                If _lastParamResult Is Nothing OrElse _lastParamResult.Details Is Nothing OrElse _lastParamResult.Details.Count = 0 Then
                    SendToWeb("sharedparam:exported", New With {.ok = False, .message = "최근 실행 결과가 없습니다."})
                    Return
                End If

                Dim doAutoFit As Boolean = ParseExcelMode(payload)
                Dim saved As String = ParamPropagateService.ExportResultToExcel(_lastParamResult, doAutoFit)
                If String.IsNullOrWhiteSpace(saved) Then
                    SendToWeb("sharedparam:exported", New With {.ok = False, .message = "엑셀 내보내기가 취소되었습니다."})
                    Return
                End If

                SendToWeb("sharedparam:exported", New With {.ok = True, .path = saved})
            Catch ex As Exception
                SendToWeb("sharedparam:exported", New With {.ok = False, .message = ex.Message})
                SendToWeb("revit:error", New With {.message = "엑셀 내보내기 실패: " & ex.Message})
            End Try
        End Sub

    End Class

End Namespace
