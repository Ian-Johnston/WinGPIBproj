' User customizeable tab - Timers


Partial Class Formtest


    ' --- Timer15 (QUERIESTOFILE) state ---
    Private Auto15DeviceName As String = ""
    Private Auto15ScriptBoxName As String = ""
    Private Auto15FilePathControl As String = ""
    Private Auto15ResultControl As String = ""


    ' Timer15 - Used to send block of GPIB commands from a TEXTAREA and logs the results to CSV
    Private Sub Timer15_Tick(sender As Object, e As EventArgs) Handles Timer15.Tick
        Dim autoCb As CheckBox = GetCheckboxFor(Auto15ScriptBoxName, "FuncAuto")
        If autoCb Is Nothing OrElse Not autoCb.Checked Then
            Timer15.Enabled = False
            Exit Sub
        End If

        RunQueriesToFileFromTextArea(Auto15DeviceName,
                                     Auto15ScriptBoxName,
                                     Auto15FilePathControl,
                                     Auto15ResultControl)
    End Sub


    ' Timer 5 - Used for dev1 FuncAuto
    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        If AutoJobs5.Count = 0 Then
            Timer5.Enabled = False
            Exit Sub
        End If

        Dim nowT As Integer = NowTick()
        Dim keys = AutoJobs5.Keys.ToList()

        For Each k In keys
            Dim j As AutoJob = Nothing
            If Not AutoJobs5.TryGetValue(k, j) Then Continue For
            If j.InFlight Then Continue For
            If Not Due(nowT, j.NextDue) Then Continue For

            j.InFlight = True
            Try
                RunQueryToResult(j.Device, j.Command, j.Result, Nothing, j.OverloadToken)
            Finally
                j.InFlight = False
                j.NextDue = nowT + j.IntervalMs
            End Try
        Next
    End Sub


    ' Timer 16 - Used for dev2 FuncAuto
    Private Sub Timer16_Tick(sender As Object, e As EventArgs) Handles Timer16.Tick
        If AutoJobs16.Count = 0 Then
            Timer16.Enabled = False
            Exit Sub
        End If

        Dim nowT As Integer = NowTick()
        Dim keys = AutoJobs16.Keys.ToList()

        For Each k In keys
            Dim j As AutoJob = Nothing
            If Not AutoJobs16.TryGetValue(k, j) Then Continue For
            If j.InFlight Then Continue For
            If Not Due(nowT, j.NextDue) Then Continue For

            j.InFlight = True
            Try
                RunQueryToResult(j.Device, j.Command, j.Result, Nothing, j.OverloadToken)
            Finally
                j.InFlight = False
                j.NextDue = nowT + j.IntervalMs
            End Try
        Next
    End Sub


End Class