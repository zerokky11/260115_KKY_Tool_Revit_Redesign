Imports System
Imports System.Collections.Generic

Namespace UI.Hub

    Public Module ExcelProgressReporter

        Private Class ProgressState
            Public Property LastSent As DateTime = DateTime.MinValue
            Public Property LastRow As Integer = 0
        End Class

        Private ReadOnly States As New Dictionary(Of String, ProgressState)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly Gate As New Object()

        Public Sub Reset(channel As String)
            If String.IsNullOrWhiteSpace(channel) Then
                Return
            End If
            SyncLock Gate
                States(channel) = New ProgressState()
            End SyncLock
        End Sub

        Public Sub Report(channel As String,
                          phase As String,
                          message As String,
                          current As Integer,
                          total As Integer,
                          Optional percentOverride As Double? = Nothing,
                          Optional force As Boolean = False)
            If String.IsNullOrWhiteSpace(channel) Then
                Return
            End If

            Dim shouldSend As Boolean = force
            Dim now As DateTime = DateTime.UtcNow

            SyncLock Gate
                Dim st As ProgressState = Nothing
                If Not States.TryGetValue(channel, st) Then
                    st = New ProgressState()
                    States(channel) = st
                End If

                Dim deltaMs As Double = (now - st.LastSent).TotalMilliseconds
                Dim deltaRows As Integer = Math.Abs(current - st.LastRow)

                If force OrElse deltaMs >= 200.0 OrElse deltaRows >= 200 Then
                    st.LastSent = now
                    st.LastRow = current
                    shouldSend = True
                End If
            End SyncLock

            If Not shouldSend Then
                Return
            End If

            Dim phaseProgress As Double = ComputePhaseProgress(phase, current, total)
            UiBridgeExternalEvent.SendToWeb(channel, New With {
                .phase = phase,
                .message = message,
                .current = current,
                .total = total,
                .phaseProgress = phaseProgress,
                .percent = If(percentOverride.HasValue, percentOverride.Value, CType(Nothing, Double?))
            })
        End Sub

        Private Function ComputePhaseProgress(phase As String, current As Integer, total As Integer) As Double
            Dim norm As String = If(phase, String.Empty).Trim().ToUpperInvariant()
            If norm = "EXCEL_WRITE" OrElse norm = "AUTOFIT" Then
                If total <= 0 Then Return 0.0
                Return Math.Min(1.0, Math.Max(0.0, CDbl(current) / CDbl(total)))
            End If
            If norm = "DONE" Then Return 1.0
            If norm = "EXCEL_SAVE" Then Return 1.0
            Return 0.0
        End Function

    End Module

End Namespace
