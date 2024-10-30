
' Device 1 & 2 run code including the timers for each

Imports IODevices
Imports System.Threading
Imports System.Diagnostics
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Threading.Tasks

Partial Class Formtest

    Dim Dev1TimerDuration As Double
    Dim Dev2TimerDuration As Double
    Dim Dev12TimerDuration As Double
    Dim Dev1SampleCount As Integer  ' Max =  2,147,483,647
    Dim Dev2SampleCount As Integer  ' Max =  2,147,483,647
    Dim txtr1alogged As String
    Dim txtr2alogged As String
    Dim inst_value1F As Double
    Dim inst_value2F As Double
    Dim inst_value3F As Double
    Dim lineCount As Integer
    Dim TestLen As Integer

    Dim Dev2_3457A As Boolean
    Dim Dev1_3457A As Boolean

    Private Dev1stopWatch As New Stopwatch()    ' for interrupt
    Private Dev1elapsedTime As Integer = 0 ' Variable to keep track of elapsed time
    Private Dev1runStopwatch As Boolean = False ' Flag to indicate if the stopwatch should be run
    Private Dev1isStopwatchRunning As Boolean = False ' Flag to indicate if the stopwatch is running
    Dim Dev1Timer2codeRunning As Boolean = True
    Private Dev1elapsedSeconds As Double

    Private Dev2stopWatch As New Stopwatch()    ' for interrupt
    Private Dev2elapsedTime As Integer = 0 ' Variable to keep track of elapsed time
    Private Dev2runStopwatch As Boolean = False ' Flag to indicate if the stopwatch should be run
    Private Dev2isStopwatchRunning As Boolean = False ' Flag to indicate if the stopwatch is running
    Dim Dev2Timer3codeRunning As Boolean = True
    Private Dev2elapsedSeconds As Double

    Private Dev1waitingForText As Boolean = False
    Private Dev2waitingForText As Boolean = False


    Private Sub Dev12Run_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDev12Run.Click

        Dev1INT.ForeColor = Color.Black
        Dev1INTb.ForeColor = Color.Black
        Dev2INT.ForeColor = Color.Black
        Dev2INTb.ForeColor = Color.Black

        If (ButtonDev12Run.Text = "Run") Then
            EnableChart1.Enabled = True
            EnableChart2.Enabled = True
            CheckBoxDevice1Hide.Enabled = True
            CheckBoxDevice2Hide.Enabled = True

            CSVfilename.ReadOnly = True
            CheckboxEnableCSV.Enabled = False
            'EnableResetCSV.Enabled = True
            TextFilenameAppend.ReadOnly = True
            'ResetCSV.Enabled = True
        Else
            EnableChart1.Checked = False
            EnableChart2.Checked = False
            EnableChart1.Enabled = False
            EnableChart2.Enabled = False
            CheckBoxDevice1Hide.Checked = False
            CheckBoxDevice1Hide.Enabled = False
            CheckBoxDevice2Hide.Checked = False
            CheckBoxDevice2Hide.Enabled = False

            CSVfilename.ReadOnly = False
            CheckboxEnableCSV.Enabled = True
            'EnableResetCSV.Enabled = False
            TextFilenameAppend.ReadOnly = False
            'ResetCSV.Enabled = False
        End If

        ' Start
        If (ButtonDev12Run.Text = "Run") Then

            ' Start HRS:MINS stopwatch
            'Me.Timer8.Start()
            stopwatch.Reset()
            stopwatch.Start()

            'Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & " Dev 1 & 2 Running")

            Dev1Samples.Text = 0
            Dev1SampleCount = 0
            ButtonDev12Run.Text = "Stop"
            Dev2Samples.Text = 0
            Dev2SampleCount = 0

            IndexCount = 1  ' reset index count for CSV file


            ' Dev1 Send all lines from command PRE-RUN text box
            lineCount = CommandStart1.Lines.Count

            For i = 0 To (lineCount - 1)
                'For i = 0 To (lineCount - 1)

                If IgnoreErrors1.Checked = False Then
                    dev1.SendAsync(CommandStart1.Lines(i), True)
                Else
                    dev1.SendAsync(CommandStart1.Lines(i), False)
                End If

                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Pre-Run Dev.1 " & CommandStart1.Lines(i))

            Next i

            ' Dev2 Send all lines from command start text box
            lineCount = CommandStart2.Lines.Count

            For i = 0 To (lineCount - 1)
                'For i = 0 To (lineCount - 1)

                If IgnoreErrors2.Checked = False Then
                    dev2.SendAsync(CommandStart2.Lines(i) & TermStr(), True)       ' & TermStr() added by IanJ for terminator check button option, see Formtest.vb
                Else
                    dev2.SendAsync(CommandStart2.Lines(i) & TermStr(), False)      ' & TermStr() added by IanJ for terminator check button option, see Formtest.vb
                End If

                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Pre-Run Dev.2 " & CommandStart2.Lines(i))
            Next i

            ' Set Timer 2 & 3 duration in mS
            Dev12TimerDuration = Dev12SampleRate.Text
            Me.Timer2.Interval = Dev12TimerDuration * 1000
            Me.Timer2.Start()
            Me.Timer3.Interval = Dev12TimerDuration * 1000
            Me.Timer3.Start()
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Dev.1 & 2 Running")


            ' Stop
        ElseIf (ButtonDev12Run.Text = "Stop" And Timer2.Enabled And Timer3.Enabled) Then

            ' Reset interrupt command settings
            Dev1elapsedTime = 0
            Dev1runStopwatch = False
            Dev1stopWatch.Reset()
            Dev1isStopwatchRunning = False
            Dev1Timer2codeRunning = True
            Dev2elapsedTime = 0
            Dev2runStopwatch = False
            Dev2stopWatch.Reset()
            Dev2isStopwatchRunning = False
            Dev2Timer3codeRunning = True

            Me.Timer2.Stop()
            Me.Timer3.Stop()
            ButtonDev12Run.Text = "Run"
            'gboxdev.Enabled = True

            ' Dev1 Send all lines from command stop text box
            lineCount = CommandStop1.Lines.Count
            For i = 0 To (lineCount - 1)

                If IgnoreErrors1.Checked = False Then
                    dev1.SendAsync(CommandStop1.Lines(i), True)
                    ' dev1.QueryAsync(CommandStop1.Lines(i), AddressOf cbdev1, True)
                Else
                    dev1.SendAsync(CommandStop1.Lines(i), False)
                    'dev1.QueryAsync(CommandStop1.Lines(i), AddressOf cbdev1, False)
                End If

                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Halt Dev.1 GPIB with " & CommandStop1.Lines(i))
            Next i

            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Dev.1 Stopped")

            ' full dispose
            gbox1.Enabled = True   'enable sending new commands
            'dev1.Dispose()

            ' Dev2 Send all lines from command stop text box
            lineCount = CommandStop2.Lines.Count
            For i = 0 To (lineCount - 1)

                If IgnoreErrors2.Checked = False Then
                    dev2.SendAsync(CommandStop2.Lines(i), True)
                    'dev2.QueryAsync(CommandStop2.Lines(i), AddressOf cbdev2, True)
                Else
                    dev2.SendAsync(CommandStop2.Lines(i), False)
                    'dev2.QueryAsync(CommandStop2.Lines(i), AddressOf cbdev2, False)
                End If

                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Halt Dev.2 GPIB with " & CommandStop2.Lines(i))
            Next i

            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Dev.2 Stopped")

            ' full dispose
            gbox2.Enabled = True   'enable sending new commands
            'dev2.Dispose()
        End If

    End Sub


    Private Sub Dev1Run_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDev1Run.Click

        Dev1INT.ForeColor = Color.Black
        Dev1INTb.ForeColor = Color.Black

        If (ButtonDev1Run.Text = "Run") Then

            EnableChart1.Enabled = True
            CheckBoxDevice1Hide.Enabled = True

            CSVfilename.ReadOnly = True
            CheckboxEnableCSV.Enabled = False
            'EnableResetCSV.Enabled = True
            TextFilenameAppend.ReadOnly = True
            'ResetCSV.Enabled = True
        Else
            EnableChart1.Checked = False
            EnableChart1.Enabled = False
            CheckBoxDevice1Hide.Checked = False
            CheckBoxDevice1Hide.Enabled = False

            CSVfilename.ReadOnly = False
            CheckboxEnableCSV.Enabled = True
            'EnableResetCSV.Enabled = False
            TextFilenameAppend.ReadOnly = False
            'ResetCSV.Enabled = False
        End If

        ' Start
        If (ButtonDev1Run.Text = "Run") Then

            ' Start HRS:MINS stopwatch
            'Me.Timer8.Start()
            stopwatch.Reset()
            stopwatch.Start()

            'Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & " Dev 1 Running")

            Dev1Samples.Text = 0
            Dev1SampleCount = 0
            ButtonDev1Run.Text = "Stop"
            'gboxdev.Enabled = True

            ' Set up options - Added by IanJ
            If Dev1PollingEnable.Checked = True Then
                dev1.enablepoll = True
            Else
                dev1.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
            End If

            If Dev1STBMask.Text = "" Then
                Dev1STBMask.Text = "16"
            End If
            dev1.MAVmask = Val(Dev1STBMask.Text)
            If Dev1STBMask.Text = "0" Then
                dev1.enablepoll = False
                Dev1PollingEnable.Checked = False
            End If


            'If Dev1CheckEOI.Checked = True Then
            'dev1.checkEOI = True      'use EOI information: if true repeat read if EOI not detected (eg. buffer too small); default=true,
            'Else
            'dev1.checkEOI = False     'use EOI information: if true repeat read if EOI not detected (eg. buffer too small); default=true,
            'End If


            IndexCount = 1  ' reset index count for CSV file


            ' Send all lines from command PRE-RUN text box
            lineCount = CommandStart1.Lines.Count

            For i = 0 To (lineCount - 1)
                'For i = 0 To (lineCount - 1)

                If IgnoreErrors1.Checked = False Then
                    dev1.SendAsync(CommandStart1.Lines(i), True)
                    'dev1.QueryAsync(CommandStart1.Lines(i), AddressOf cbdev1, True)
                Else
                    dev1.SendAsync(CommandStart1.Lines(i), False)
                    'dev1.QueryAsync(CommandStart1.Lines(i), AddressOf cbdev1, False)
                End If

                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Pre-Run Dev.1 " & CommandStart1.Lines(i))
            Next i

            ' Set Timer2 duration in mS
            Dev1TimerDuration = Dev1SampleRate.Text
            Me.Timer2.Interval = Dev1TimerDuration * 1000
            Me.Timer2.Start()
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Dev.1 Running")


            ' Stop
        ElseIf (ButtonDev1Run.Text = "Stop" And Timer2.Enabled) Then

            ' Reset interrupt command settings
            Dev1elapsedTime = 0
            Dev1runStopwatch = False
            Dev1stopWatch.Reset()
            Dev1isStopwatchRunning = False
            Dev1Timer2codeRunning = True

            Me.Timer2.Stop()
            ButtonDev1Run.Text = "Run"

            ' Send all lines from command stop text box
            lineCount = CommandStop1.Lines.Count
            For i = 0 To (lineCount - 1)

                If IgnoreErrors1.Checked = False Then
                    dev1.SendAsync(CommandStop1.Lines(i), True)
                    'dev1.QueryAsync(CommandStop1.Lines(i), AddressOf cbdev1, True)
                Else
                    dev1.SendAsync(CommandStop1.Lines(i), False)
                    'dev1.QueryAsync(CommandStop1.Lines(i), AddressOf cbdev1, False)
                End If

                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Halt Dev.1 GPIB with " & CommandStop1.Lines(i))
            Next i

            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Dev.1 Stopped")

            ' full dispose
            gbox1.Enabled = True   'enable sending new commands
            'dev1.Dispose()
        End If


        ' If no RUN command then abort, ideal for running PRE-RUN commands only
        If (CommandStart1run.Text = "") Then
            Me.Timer2.Stop()
            ButtonDev1Run.Text = "Run"
            gbox1.Enabled = True   'enable sending new commands
        End If

    End Sub


    Private Sub ButtonDev1PreRun_Click(sender As Object, e As EventArgs) Handles ButtonDev1PreRun.Click

        ' Set up options - Added by IanJ
        If Dev1PollingEnable.Checked = True Then
            dev1.enablepoll = True
        Else
            dev1.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
        End If

        If Dev1STBMask.Text = "" Then
            Dev1STBMask.Text = "16"
        End If
        dev1.MAVmask = Val(Dev1STBMask.Text)
        If Dev1STBMask.Text = "0" Then
            dev1.enablepoll = False
            Dev1PollingEnable.Checked = False
        End If

        ' Send all lines from command PRE-RUN text box
        lineCount = CommandStart1.Lines.Count

        For i = 0 To (lineCount - 1)
            'For i = 0 To (lineCount - 1)

            If IgnoreErrors1.Checked = False Then
                dev1.SendAsync(CommandStart1.Lines(i), True)
                'dev1.QueryAsync(CommandStart1.Lines(i), AddressOf cbdev1, True)
            Else
                dev1.SendAsync(CommandStart1.Lines(i), False)
                'dev1.QueryAsync(CommandStart1.Lines(i), AddressOf cbdev1, False)
            End If

            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "PRE-RUN Dev.1 GPIB with " & CommandStart1.Lines(i))
        Next i

    End Sub


    Private Sub Dev2Run_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDev2Run.Click

        Dev2INT.ForeColor = Color.Black
        Dev2INTb.ForeColor = Color.Black

        If (ButtonDev2Run.Text = "Run") Then
            EnableChart2.Enabled = True
            CheckBoxDevice2Hide.Enabled = True

            CSVfilename.ReadOnly = True
            CheckboxEnableCSV.Enabled = False
            'EnableResetCSV.Enabled = True
            TextFilenameAppend.ReadOnly = True
            'ResetCSV.Enabled = True
        Else
            EnableChart2.Checked = False
            EnableChart2.Enabled = False
            CheckBoxDevice2Hide.Checked = False
            CheckBoxDevice2Hide.Enabled = False

            CSVfilename.ReadOnly = False
            CheckboxEnableCSV.Enabled = True
            'EnableResetCSV.Enabled = False
            TextFilenameAppend.ReadOnly = False
            'ResetCSV.Enabled = False
        End If

        ' Start
        If (ButtonDev2Run.Text = "Run") Then

            ' Start HRS:MINS stopwatch
            'Me.Timer8.Start()
            stopwatch.Reset()
            stopwatch.Start()

            'Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & " Dev 2 Running")

            Dev2Samples.Text = 0
            Dev2SampleCount = 0
            ButtonDev2Run.Text = "Stop"
            'gboxdev.Enabled = True

            ' Set up options - Added by IanJ
            If Dev2PollingEnable.Checked = True Then
                dev2.enablepoll = True      'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
            Else
                dev2.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
            End If

            If Dev2STBMask.Text = "" Then
                Dev2STBMask.Text = "16"
            End If
            dev2.MAVmask = Val(Dev2STBMask.Text)
            If Dev2STBMask.Text = "0" Then
                dev2.enablepoll = False
                Dev2PollingEnable.Checked = False
            End If


            'If Dev2CheckEOI.Checked = True Then
            'dev2.checkEOI = True      'use EOI information: if true repeat read if EOI not detected (eg. buffer too small); default=true,
            'Else
            'dev2.checkEOI = False     'use EOI information: if true repeat read if EOI not detected (eg. buffer too small); default=true,
            'End If

            gbox2.Enabled = True

            ' Send all lines from command PRE-RUN text box
            lineCount = CommandStart2.Lines.Count

            For i = 0 To (lineCount - 1)

                If IgnoreErrors2.Checked = False Then
                    dev2.SendAsync(CommandStart2.Lines(i) & TermStr(), True)         ' & TermStr() added by IanJ for terminator check button option, see Formtest.vb
                    'dev2.QueryAsync(CommandStart2.Lines(i), AddressOf cbdev2, True)
                Else
                    dev2.SendAsync(CommandStart2.Lines(i) & TermStr(), False)        ' & TermStr() added by IanJ for terminator check button option, see Formtest.vb
                    'dev2.QueryAsync(CommandStart2.Lines(i), AddressOf cbdev2, False)
                End If

                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Pre-Run Dev.2 " & CommandStart2.Lines(i))
            Next i

            ' Set Timer2 duration in mS
            Dev2TimerDuration = Dev2SampleRate.Text
            Me.Timer3.Interval = Dev2TimerDuration * 1000
            Me.Timer3.Start()
            'Log(DateTime.Now & " Device 2 Running.......")
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Dev.2 Running")


            ' Stop
        ElseIf (ButtonDev2Run.Text = "Stop" And Timer3.Enabled) Then

            ' Reset interrupt command settings
            Dev2elapsedTime = 0
            Dev2runStopwatch = False
            Dev2stopWatch.Reset()
            Dev2isStopwatchRunning = False
            Dev2Timer3codeRunning = True

            Me.Timer3.Stop()
            ButtonDev2Run.Text = "Run"

            ' Send all lines from command stop text box
            lineCount = CommandStop2.Lines.Count
            For i = 0 To (lineCount - 1)

                If IgnoreErrors2.Checked = False Then
                    dev2.SendAsync(CommandStop2.Lines(i), True)
                    'dev2.QueryAsync(CommandStop2.Lines(i), AddressOf cbdev2, True)
                Else
                    dev2.SendAsync(CommandStop2.Lines(i), False)
                    'dev2.QueryAsync(CommandStop2.Lines(i), AddressOf cbdev2, False)
                End If

                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Halt Dev.2 GPIB with " & CommandStop2.Lines(i))
            Next i

            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "Dev.2 Stopped")

            ' full dispose
            gbox2.Enabled = True   'enable sending new commands
            'dev2.Dispose()

        End If


        ' If no RUN command then abort, ideal for running PRE-RUN commands only
        If (CommandStart2run.Text = "") Then
            Me.Timer3.Stop()
            ButtonDev2Run.Text = "Run"
            gbox2.Enabled = True   'enable sending new commands
        End If


    End Sub


    Private Sub ButtonDev2PreRun_Click(sender As Object, e As EventArgs) Handles ButtonDev2PreRun.Click

        ' Set up options - Added by IanJ
        If Dev2PollingEnable.Checked = True Then
            dev2.enablepoll = True      'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
        Else
            dev2.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
        End If

        If Dev2STBMask.Text = "" Then
            Dev2STBMask.Text = "16"
        End If
        dev2.MAVmask = Val(Dev2STBMask.Text)
        If Dev2STBMask.Text = "0" Then
            dev2.enablepoll = False
            Dev2PollingEnable.Checked = False
        End If

        ' Send all lines from command PRE-RUN text box
        lineCount = CommandStart2.Lines.Count

        For i = 0 To (lineCount - 1)

            If IgnoreErrors2.Checked = False Then
                dev2.SendAsync(CommandStart2.Lines(i) & TermStr(), True)         ' & TermStr() added by IanJ for terminator check button option, see Formtest.vb
                'dev2.QueryAsync(CommandStart2.Lines(i), AddressOf cbdev2, True)
            Else
                dev2.SendAsync(CommandStart2.Lines(i) & TermStr(), False)        ' & TermStr() added by IanJ for terminator check button option, see Formtest.vb
                'dev2.QueryAsync(CommandStart2.Lines(i), AddressOf cbdev2, False)
            End If

            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "PRE-RUN Dev.2 GPIB with " & CommandStart2.Lines(i))
        Next i

    End Sub


    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        ' Dev 1 Logging

        ' Interrupt commands via stopwatch
        If Dev1IntEnable.Checked = True Then

            ' Increment the elapsed time counter with the interval of the timer
            ' Elapsed time = Timer 2 (sample rate * 1000) / 1000, so say increments every 5 secs (5, 10, 15, 20 etc) which gives the elapsed time in secs
            Dev1elapsedTime += Timer2.Interval / 1000                           ' Convert the logging interval from milliseconds to seconds and add to existing

            'Label297.Text = Dev1elapsedTime

            ' Need to check if elapsed time in mins is >= period in mins
            'If Dev1elapsedTime >= (Val(Dev1runStopwatchEveryInMins.Text) * 60) - Val(Dev1pauseDurationInSeconds.Text) - Val(Dev1SampleRate.Text) Then          ' also includes compensation for missing secs
            If Dev1elapsedTime >= (Val(Dev1runStopwatchEveryInMins.Text) * 60) And Dev1SendQuery.Checked = False Then         ' Send Async - If period set to 2mins and 10sec duration then you will get 2mins, then 10sec interrupt, then another 2mins etc.
                Dev1Timer2codeRunning = False

                Dev1INT.ForeColor = Color.Red
                Dev1INTb.ForeColor = Color.Red
                Refresh()
                Dev1elapsedTime = 0 ' Reset the elapsed time counter
                Dev1isStopwatchRunning = True
                Dev1Interrupt() ' Call the method to run additional code during the stopwatch run
                System.Threading.Thread.Sleep(100)       ' delay to give time for int sb to run

                'Timer2.Enabled = False ' Pause the timer
                'Label295.Text = "Interrupt Command Called"
                Dev1stopWatch.Reset()
                Dev1stopWatch.Start() ' Start the Stopwatch
            End If

            If Dev1elapsedTime >= (Val(Dev1runStopwatchEveryInMins.Text) * 60) And Dev1SendQuery.Checked = True Then         ' Query Async - No duration necessary since query async will reply and the period can continue

                Dev1Timer2codeRunning = False       ' stop logging timer

                Dev1INT.ForeColor = Color.Red
                Dev1INTb.ForeColor = Color.Red
                Refresh()
                Dev1elapsedTime = 0 ' Reset the elapsed time counter
                Dev1Interrupt() ' Call the method to run additional code during the stopwatch run
                System.Threading.Thread.Sleep(100)       ' delay to give time for int sb to run

                Dev1INT.ForeColor = Color.Black
                Dev1INTb.ForeColor = Color.Black

                Dev1Timer2codeRunning = True        ' restart logging timer here because no need for stopwatch reset etc

            End If

            ' Get the number of seconds the Stopwatch has been running
            Dev1elapsedSeconds = Dev1stopWatch.Elapsed.TotalSeconds

            If Dev1isStopwatchRunning And (Dev1elapsedSeconds >= Val(Dev1pauseDurationInSeconds.Text)) And Dev1SendQuery.Checked = False Then           ' only if Send Async mode
                ' Stop the Stopwatch and resume the Timer code
                Dev1isStopwatchRunning = False
                Dev1stopWatch.Stop()
                Dev1stopWatch.Reset()
                Dev1Timer2codeRunning = True

                Dev1INT.ForeColor = Color.Black
                Dev1INTb.ForeColor = Color.Black

                Dev1elapsedTime = 0 ' Reset the elapsed time counter
                'Timer2.Enabled = True ' Resume the timer
            End If

        End If


        If Dev1Timer2codeRunning = True Then

            ' Dev 1 timer
            If dev1.PendingTasks(txtq1a.Text) <= 3 Then
                'example of using PendingTasks() method
                'dev1.QueryAsync(txtq1a.Text, AddressOf cbdev1, True)

                ' Read last line from starting command list and repeat it via this timer
                lineCount = CommandStart1.Lines.Count
                'CommandStartRun1.Text = (CommandStart1.Lines(lineCount - 1))

                If IgnoreErrors1.Checked = False Then

                    If CheckBoxSendBlockingDev1.Checked = True Then
                        Dim q As IOQuery = Nothing
                        dev1.QueryBlocking(CommandStart1run.Text, q, True) 'simpler version with string parameter
                        Cbdev1(q)
                    Else
                        dev1.QueryAsync(CommandStart1run.Text, AddressOf Cbdev1, True)      ' 3457A mode is OFF so execute normal command

                        ' Read 7th digit from 3457A and tag it's happened
                        If Dev13457Aseven.Checked = True Then
                            Dev1_3457A = True
                            dev1.QueryAsync("RMATH HIRES", AddressOf Cbdev1, True)
                        End If

                    End If
                    'dev1.QueryAsync(CommandStart1run.Text, AddressOf cbdev1, True)

                Else

                    If CheckBoxSendBlockingDev1.Checked = True Then
                        Dim q As IOQuery = Nothing
                        dev1.QueryBlocking(CommandStart1run.Text, q, False) 'simpler version with string parameter
                        Cbdev1(q)
                    Else
                        dev1.QueryAsync(CommandStart1run.Text, AddressOf Cbdev1, False)

                        ' Read 7th digit from 3457A and tag it's happened
                        If Dev13457Aseven.Checked = True Then
                            Dev1_3457A = True
                            dev1.QueryAsync("RMATH HIRES", AddressOf Cbdev1, False)
                        End If

                    End If
                    'dev1.QueryAsync(CommandStart1run.Text, AddressOf cbdev1, False)

                End If

            Else
                txtr1astat.Text = "already 3 '" & txtq1a.Text & "' commands pending"
            End If

        End If

    End Sub


    Private Sub Dev1Interrupt()

        ' Run Interrupt command, this is sent one-off

        If Dev1SendQuery.Checked = False Then

            ' Send Async
            If Not String.IsNullOrEmpty(txtq1d.Text) Then
                dev1.SendAsync(txtq1d.Text, True)
                txtr1astat.Text = "Send Async '" & txtq1d.Text
                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & " Dev 1 Interrupt = " & txtq1d.Text)
            End If

        Else

            ' Query Async
            If Not String.IsNullOrEmpty(txtq1d.Text) Then

                Me.Timer2.Stop()
                Dev1Timer2codeRunning = False       ' stop logging timer

                txtr1a.Text = ""

                dev1.QueryAsync(txtq1d.Text, AddressOf Cbdev1, True)
                txtr1astat.Text = "Query Async '" & txtq1d.Text
                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & " Dev 1 Interrupt = " & txtq1d.Text)

                'Do While txtr1a.Text = ""               ' there is a better way, but can't find it.
                'Application.DoEvents()
                'Loop

                ' Now set up waiting for txtr1a.Text to contain something, i.e. rely is received.
                ' Set the flag and start the CheckTimer
                Dev1waitingForText = True
                Me.Timer11.Interval = 100
                Timer11.Start()

                ' Reply has been received
                'Me.Timer2.Start()
                'Dev1Timer2codeRunning = True       ' restart logging timer

            End If

        End If

    End Sub

    Private Sub CheckTimer11_Tick(sender As Object, e As EventArgs) Handles Timer11.Tick
        If Dev1waitingForText AndAlso txtr1a.Text <> "" Then
            ' Reply has been received
            Timer11.Stop()
            Dev1waitingForText = False

            Me.Timer2.Start()
            Dev1Timer2codeRunning = True       ' restart logging timer

            ' Additional code to handle the updated text can be added here if needed
        End If
    End Sub


    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick

        ' Dev 2 Logging

        ' Interrupt commands via stopwatch
        If Dev2IntEnable.Checked = True Then

            ' Increment the elapsed time counter with the interval of the timer
            ' Elapsed time = Timer 2 (sample rate * 1000) / 1000, so say increments every 5 secs (5, 10, 15, 20 etc) which gives the elapsed time in secs
            Dev2elapsedTime += Timer3.Interval / 1000                           ' Convert the logging interval from milliseconds to seconds and add to existing

            'Label297.Text = Dev2elapsedTime

            ' Need to check if elapsed time in mins is >= period in mins
            'If Dev2elapsedTime >= Val(Dev2runStopwatchEveryInSeconds.Text) * 60 - Val(Dev2pauseDurationInSeconds.Text) - Val(Dev2SampleRate.Text) Then          ' also includes compensation for missing secs
            If Dev2elapsedTime >= Val(Dev2runStopwatchEveryInMins.Text) * 60 And Dev2SendQuery.Checked = False Then         ' Send Async

                Dev2Timer3codeRunning = False

                Dev2INT.ForeColor = Color.Red
                Dev2INTb.ForeColor = Color.Red
                Refresh()
                Dev2elapsedTime = 0 ' Reset the elapsed time counter
                Dev2isStopwatchRunning = True
                Dev2Interrupt() ' Call the method to run additional code during the stopwatch run
                System.Threading.Thread.Sleep(100)       ' delay to give time for int sb to run

                'Timer2.Enabled = False ' Pause the timer
                'Label295.Text = "Interrupt Command Called"
                Dev2stopWatch.Reset()
                Dev2stopWatch.Start() ' Start the Stopwatch
            End If

            If Dev2elapsedTime >= (Val(Dev2runStopwatchEveryInMins.Text) * 60) And Dev2SendQuery.Checked = True Then         ' Query Async - No duration necessary since query async will reply and the period can continue

                Dev2Timer3codeRunning = False       ' stop logging timer

                Dev2INT.ForeColor = Color.Red
                Dev2INTb.ForeColor = Color.Red
                Refresh()
                Dev2elapsedTime = 0 ' Reset the elapsed time counter
                Dev2Interrupt() ' Call the method to run additional code during the stopwatch run
                System.Threading.Thread.Sleep(100)       ' delay to give time for int sb to run

                Dev2INT.ForeColor = Color.Black
                Dev2INTb.ForeColor = Color.Black

                Dev2Timer3codeRunning = True        ' restart logging timer here because no need for stopwatch reset etc

            End If

            ' Get the number of seconds the Stopwatch has been running
            Dev2elapsedSeconds = Dev2stopWatch.Elapsed.TotalSeconds

            If Dev2isStopwatchRunning AndAlso Dev2elapsedSeconds >= Val(Dev2pauseDurationInSeconds.Text) And Dev2SendQuery.Checked = False Then           ' only if Send Async mode
                ' Stop the Stopwatch and resume the Timer code
                Dev2isStopwatchRunning = False
                Dev2stopWatch.Stop()
                Dev2stopWatch.Reset()
                Dev2Timer3codeRunning = True

                Dev2INT.ForeColor = Color.Black
                Dev2INTb.ForeColor = Color.Black

                Dev2elapsedTime = 0 ' Reset the elapsed time counter
                'Timer3.Enabled = True ' Resume the timer
            End If

        End If

        If Dev2Timer3codeRunning = True Then

            ' Dev 2 timer
            If dev2.PendingTasks(txtq2a.Text) <= 3 Then
                'example of using PendingTasks() method
                'dev2.QueryAsync(txtq2a.Text, AddressOf cbdev2, True)

                ' Read last line from starting command list and repeat it via this timer
                lineCount = CommandStart2.Lines.Count
                'CommandStartRun2.Text = (CommandStart2.Lines(lineCount - 1))

                If IgnoreErrors2.Checked = False Then

                    If CheckBoxSendBlockingDev2.Checked = True Then
                        Dim q As IOQuery = Nothing
                        dev2.QueryBlocking(CommandStart2run.Text, q, True) 'simpler version with string parameter
                        Cbdev2(q)
                    Else

                        Dev2_3457A = False
                        dev2.QueryAsync(CommandStart2run.Text, AddressOf Cbdev2, True)      ' 3457A mode is OFF so execute normal command

                        ' Read 7th digit from 3457A and tag it's happened.
                        If Dev23457Aseven.Checked = True Then
                            Dev2_3457A = True
                            dev2.QueryAsync("RMATH HIRES", AddressOf Cbdev2, True)
                        End If

                    End If
                    'dev2.QueryAsync(CommandStart2run.Text, AddressOf cbdev2, True)

                Else

                    If CheckBoxSendBlockingDev2.Checked = True Then
                        Dim q As IOQuery = Nothing
                        dev2.QueryBlocking(CommandStart2run.Text, q, False) 'simpler version with string parameter
                        Cbdev2(q)
                    Else

                        Dev2_3457A = False
                        dev2.QueryAsync(CommandStart2run.Text, AddressOf Cbdev2, False)

                        ' Read 7th digit from 3457A and tag it's happened.
                        If Dev23457Aseven.Checked = True Then
                            Dev2_3457A = True
                            dev2.QueryAsync("RMATH HIRES", AddressOf Cbdev2, False)
                        End If

                    End If
                    'dev2.QueryAsync(CommandStart2run.Text, AddressOf cbdev2, False)

                End If
                'dev2.QueryAsync(CommandStart2.Lines(lineCount - 1), AddressOf cbdev2, True)

            Else
                txtr2astat.Text = "already 3 '" & txtq2a.Text & "' commands pending"
            End If

        End If

    End Sub

    Private Sub Dev2Interrupt()

        ' Run Interrupt command, this is sent one-off

        If Dev2SendQuery.Checked = False Then

            ' Send Async
            If Not String.IsNullOrEmpty(txtq2d.Text) Then
                dev2.SendAsync(txtq2d.Text, True)
                txtr2astat.Text = "Send Async '" & txtq2d.Text
                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & " Dev 2 Interrupt = " & txtq2d.Text)
            End If

        Else

            ' Query Async
            If Not String.IsNullOrEmpty(txtq2d.Text) Then

                Me.Timer3.Stop()
                Dev2Timer3codeRunning = False       ' stop logging timer

                txtr2a.Text = ""

                dev2.QueryAsync(txtq2d.Text, AddressOf Cbdev2, True)
                txtr2astat.Text = "Query Async '" & txtq2d.Text
                Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & " Dev 2 Interrupt = " & txtq2d.Text)

                'Do While txtr2a.Text = ""               ' there is a better way, but can't find it.
                'Application.DoEvents()
                'Loop

                ' Now set up waiting for txtr2a.Text to contain something, i.e. rely is received.
                ' Set the flag and start the CheckTimer
                Dev2waitingForText = True
                Me.Timer12.Interval = 100
                Timer12.Start()

                ' Reply has been received
                'Me.Timer3.Start()
                'Dev2Timer3codeRunning = True       ' restart logging timer

            End If

        End If

    End Sub


    Private Sub CheckTimer12_Tick(sender As Object, e As EventArgs) Handles Timer12.Tick
        If Dev2waitingForText AndAlso txtr2a.Text <> "" Then
            ' Reply has been received
            Timer12.Stop()
            Dev2waitingForText = False

            Me.Timer3.Start()
            Dev2Timer3codeRunning = True       ' restart logging timer

            ' Additional code to handle the updated text can be added here if needed
        End If
    End Sub


    ' GPIB Running Time
    Private stopwatch As New Stopwatch()

    Private Sub UpdateStopwatchLabel()
        Dim elapsedTime As TimeSpan = stopwatch.Elapsed

        ' Extract hours and minutes from the elapsed time.
        Dim days As Integer = elapsedTime.Days
        Dim hours As Integer = elapsedTime.Hours
        Dim minutes As Integer = elapsedTime.Minutes
        Dim seconds As Integer = elapsedTime.Seconds

        ' Format the label text to display hours and minutes.
        RunningTimeLogging.Text = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", days, hours, minutes, seconds)

    End Sub


    Private Sub Dev1IntEnable_CheckedChanged(sender As Object, e As EventArgs) Handles Dev1IntEnable.CheckedChanged

        ' Manually unchecked Interrupt check box so reset vars

        If Dev1IntEnable.Checked = False Then
            'Dev1IntEnable.Checked = False
            ' Stop the Stopwatch and resume the Timer code
            Dev1isStopwatchRunning = False
            Dev1stopWatch.Stop()
            Dev1stopWatch.Reset()
            Dev1Timer2codeRunning = True
            Dev1INT.ForeColor = Color.Black
            Dev1INTb.ForeColor = Color.Black
        End If

    End Sub


    Private Sub Dev1SendQuery_CheckedChanged(sender As Object, e As EventArgs) Handles Dev1SendQuery.CheckedChanged

        If Dev1SendQuery.Checked = True Then
            Dev1pauseDurationInSeconds.Enabled = False
            Label294.Enabled = False
            Label296.Enabled = False
        Else
            Dev1pauseDurationInSeconds.Enabled = True
            Label294.Enabled = True
            Label296.Enabled = True
        End If

    End Sub


    Private Sub Dev2IntEnable_CheckedChanged(sender As Object, e As EventArgs) Handles Dev2IntEnable.CheckedChanged

        ' Manually unchecked Interrupt check box so reset vars

        If Dev2IntEnable.Checked = False Then
            'Dev2IntEnable.Checked = False
            ' Stop the Stopwatch and resume the Timer code
            Dev2isStopwatchRunning = False
            Dev2stopWatch.Stop()
            Dev2stopWatch.Reset()
            Dev2Timer3codeRunning = True
            Dev2INT.ForeColor = Color.Black
            Dev2INTb.ForeColor = Color.Black
        End If

    End Sub


    Private Sub Dev2SendQuery_CheckedChanged(sender As Object, e As EventArgs) Handles Dev2SendQuery.CheckedChanged

        If Dev2SendQuery.Checked = True Then
            Dev2pauseDurationInSeconds.Enabled = False
            Label293.Enabled = False
            Label299.Enabled = False
        Else
            Dev2pauseDurationInSeconds.Enabled = True
            Label293.Enabled = True
            Label299.Enabled = True
        End If

    End Sub

    Private Sub Timer3_Disposed(sender As Object, e As EventArgs) Handles Timer3.Disposed

    End Sub
End Class