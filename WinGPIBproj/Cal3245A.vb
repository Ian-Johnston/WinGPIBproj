
' Cal 3245A Universal Source

Imports IODevices
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.IO

Partial Class Formtest

    Dim CalStart As Integer = 1
    Dim CalEnd As Integer = 47
    Dim CalNum As Integer = 1

    Dim Abort3245A As Boolean = False
    Dim Calunderway As Boolean = False

    Dim TimeoutSave As String

    Dim result As Double

    Private Sub ButtonCal3245A_Click(sender As Object, e As EventArgs) Handles ButtonCal3245A.Click

        '3245A
        ' run appropriate routine depending on what needs calibrated. DCV is VOLTS only, DCV & DCI is VOLTS and CURRENT

        If RadioButton3245ADCV.Checked = True Then    ' DCV
            CalNum = 1
            Cal3245status.Text = "STARTING CALIBRATION - DCV"
            LabelRDG.Text = "#"
            Label3458ARDG.Text = "#"
            Cal3245A()
        Else                                          ' DCV & DCI
            CalNum = 1
            Cal3245status.Text = "STARTING CALIBRATION - DCV & DCI"
            LabelRDG.Text = "#"
            Label3458ARDG.Text = "#"
            Cal3245A()
        End If

    End Sub


    Private Sub CheckBoxChA_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxChA.CheckedChanged

        If CheckBoxChA.Checked = True Then
            CheckBoxChB.Checked = False
        Else
            CheckBoxChB.Checked = True
        End If

    End Sub


    Private Sub CheckBoxChB_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxChB.CheckedChanged

        If CheckBoxChB.Checked = True Then
            CheckBoxChA.Checked = False
        Else
            CheckBoxChA.Checked = True
        End If

    End Sub

    Private Sub Cal3245A()

        ' 3245A Calibration

        Cal3245status.Text = "STARTING......"
        Me.Refresh()
        System.Threading.Thread.Sleep(1000)     ' 1000mS delay


        If ButtonDev12Run.Enabled = True Then      ' Device 1 & 2 is started

            Button3245Aabort.Enabled = True

            Calunderway = True
            Abort3245A = False                      ' reset abort flag

            TimeoutSave = Dev1Timeout.Text          ' save off timout on DEVICE tab
            If Val(Timeout3458A.Text) < 5000 Then   ' check limits of timeout (ms)
                Timeout3458A.Text = "5000"
            End If
            If Val(Timeout3458A.Text) > 60000 Then
                Timeout3458A.Text = "60000"
            End If
            dev1.readtimeout = Val(Timeout3458A.Text)   ' set user adjustable timeout for Dev 1 3458A
            Dev1Timeout.Text = Timeout3458A.Text        ' copy to DEVICE 1 tab so they match.....this will be returned back to original once Cal is completed

            If CheckBoxChA.Checked = True Then
                Dialog2.Warning1 = "3245A Calibration - User intervention req'd"
                Dialog2.Warning2 = "Ensure the following connections:"
                Dialog2.Warning3 = " - Source to Ch A on 3245A"
                Dialog2.Warning4 = " - Input to voltage on DMM"
                Dialog2.Warning5 = ".......Then hit OK"
                Dialog2.ShowDialog(Me)  ' this method positions centre of parent form, and requires to hit OK to return back to parent
                CalStart = 1
                CalEnd = 47
                Cal3245A_Channel("0")       ' 3245A Channel 0
            End If

            If CheckBoxChB.Checked = True Then

                If RadioButton3458A.Checked = True Then
                    dev1.SendAsync("BEEP", True)
                    dev1.SendAsync("BEEP", True)
                End If
                If RadioButton344XXA.Checked = True Then
                    dev1.SendAsync("SYST:BEEP", True)
                    dev1.SendAsync("SYST:BEEP", True)
                End If
                If RadioButtonR6581.Checked = True Then
                    dev1.SendAsync("SYST:BEEP", True)
                    dev1.SendAsync("SYST:BEEP", True)
                End If

                Dialog2.Warning1 = "3245A Calibration - User intervention req'd"
                Dialog2.Warning2 = "Ensure the following connections:"
                Dialog2.Warning3 = " - Source to Ch B on 3245A"
                Dialog2.Warning4 = " - Input to voltage on DMM"
                Dialog2.Warning5 = ".......Then hit OK"
                Dialog2.ShowDialog(Me)  ' this method positions centre of parent form, and requires to hit OK to return back to parent
                CalStart = 1
                CalEnd = 47
                Cal3245A_Channel("100")       ' 3245A Channel 1
            End If

            Cal3245status.Text = "COMPLETED - WAIT FOR 3245A"
            Abort3245A = False
            CheckBoxAZERO.Enabled = True
            RadioButton3245ADCV.Enabled = True
            RadioButton3245ADCVDCI.Enabled = True

            Dev1Timeout.Text = TimeoutSave          ' return timout on DEVICE tab back to original
            Calunderway = False
            Button3245Aabort.Enabled = False

            Me.Refresh()

        Else

            ' GPIB Dev 1 has not been started
            Cal3245status.Text = "DEVICE 1 & 2 NOT SETUP/CONNECTED"
            Abort3245A = False
            Calunderway = False
            Me.Refresh()

        End If

    End Sub

    Private Sub Cal3245A_Channel(Channel As String)
        ' 3245A Calibration
        Me.Refresh()
        Cal3245status.Text = "SETTING UP 3458A...."
        System.Threading.Thread.Sleep(500)     ' delay
        Me.Refresh()

        If RadioButton3458A.Checked = True Then
            dev1.SendAsync("RESET", True)                   ' DCV
            Cal3245status.Text = "3458A - RESET"
            dev1.SendAsync("FUNC DCV", True)                ' DCV
            Cal3245status.Text = "3458A - FUNC DCV"

            dev1.SendAsync("RANGE AUTO", True)              ' RANGE AUTO
            Cal3245status.Text = "3458A - AUTO RANGE"

            dev1.SendAsync("END ALWAYS", True)              ' END ALWAYS
            Cal3245status.Text = "3458A - END ALWAYS"
            System.Threading.Thread.Sleep(500)              ' delay
        End If
        If RadioButton344XXA.Checked = True Then
            dev1.SendAsync("*RST", True)                   ' DCV
            Cal3245status.Text = "344xxA - *RST"
            System.Threading.Thread.Sleep(500)          ' delay
            dev1.SendAsync("VOLT:DC:NPLC " & TextBoxNPLC.Text, True)                ' DCV
            Cal3245status.Text = "344xxA - VOLT:DC:NPLC " & TextBoxNPLC.Text
            System.Threading.Thread.Sleep(500)          ' delay
            dev1.SendAsync("VOLT:DC:RANG:AUTO ON", True)              ' RANGE AUTO
            Cal3245status.Text = "344xxA - AUTO RANGE"
            System.Threading.Thread.Sleep(500)
            dev1.SendAsync("VOLT:IMP:AUTO ON", True)              ' INPUT Z AUTO
            Cal3245status.Text = "344xxA - Input Z AUTO"
            System.Threading.Thread.Sleep(500)              ' delay
        End If
        If RadioButtonR6581.Checked = True Then
            dev1.SendAsync("*RST", True)                   ' DCV
            Cal3245status.Text = "R6581 - *RST"
            System.Threading.Thread.Sleep(500)          ' delay
            dev1.SendAsync("VOLT:DC:NPLC " & TextBoxNPLC.Text, True)                ' DCV
            Cal3245status.Text = "R6581 - VOLT:DC:NPLC " & TextBoxNPLC.Text
            System.Threading.Thread.Sleep(500)          ' delay
            dev1.SendAsync("VOLT:DC:RANG:AUTO ON", True)              ' RANGE AUTO
            Cal3245status.Text = "R6581 - AUTO RANGE"
            System.Threading.Thread.Sleep(500)              ' delay
        End If

        If CheckBoxAZERO.Checked = False Then

            If RadioButton3458A.Checked = True Then
                dev1.SendAsync("AZERO OFF", True)           ' AZERO OFF - Disabled autozero, speeds up readings......this helps if the GPIB fails, sometimes at RDG 45!
                Cal3245status.Text = "3458A - AZERO OFF"
                System.Threading.Thread.Sleep(500)          ' delay
            End If
            If RadioButton344XXA.Checked = True Then
                dev1.SendAsync("VOLT:DC:ZERO:AUTO OFF", True)
                Cal3245status.Text = "344xxA - AUTO ZERO OFF"
                System.Threading.Thread.Sleep(500)          ' delay
            End If
            If RadioButtonR6581.Checked = True Then
                dev1.SendAsync("ZERO:AUTO OFF", True)
                Cal3245status.Text = "344xxA - AUTO ZERO OFF"
                System.Threading.Thread.Sleep(500)          ' delay
            End If

        Else

            If RadioButton3458A.Checked = True Then
                dev1.SendAsync("AZERO ON", True)            ' AZERO ON (default 3458A)
                Cal3245status.Text = "3458A - AZERO ON"
                System.Threading.Thread.Sleep(500)          ' delay
            End If
            If RadioButton344XXA.Checked = True Then
                dev1.SendAsync("VOLT:DC:ZERO:AUTO ON", True)
                Cal3245status.Text = "344xxA - AUTO ZERO ON"
                System.Threading.Thread.Sleep(500)          ' delay
            End If
            If RadioButtonR6581.Checked = True Then
                dev1.SendAsync("ZERO:AUTO ON", True)
                Cal3245status.Text = "R6581 - AUTO ZERO ON"
                System.Threading.Thread.Sleep(500)          ' delay
            End If

        End If

        If RadioButton3458A.Checked = True Then
            dev1.SendAsync("NRDGS 1", True)                 ' NRDGS 1
            Cal3245status.Text = "3458A - NRDGS 1"
            System.Threading.Thread.Sleep(500)              ' delay

            dev1.SendAsync("NPLC " & TextBoxNPLC.Text, True)
            Cal3245status.Text = "3458A - NPLC " & TextBoxNPLC.Text
            System.Threading.Thread.Sleep(500)
        End If
        If RadioButton344XXA.Checked = True Then
            'dev1.SendAsync("NRDGS 1", True)                 ' NRDGS 1                  not required for 344XXA
            'Cal3245status.Text = "3458A - NRDGS 1"
            'System.Threading.Thread.Sleep(500)              ' delay

            'dev1.SendAsync("NPLC 100", True)                ' NPLC 100                 already set earlier for 344XXA
            'Cal3245status.Text = "344XXA - NPLC 100"
            'System.Threading.Thread.Sleep(500)              ' delay
        End If
        If RadioButtonR6581.Checked = True Then
            'dev1.SendAsync("NRDGS 1", True)                 ' NRDGS 1                  not required for 344XXA
            'Cal3245status.Text = "3458A - NRDGS 1"
            'System.Threading.Thread.Sleep(500)              ' delay

            'dev1.SendAsync("NPLC 100", True)                ' NPLC 100                 already set earlier for R6581
            'Cal3245status.Text = "344XXA - NPLC 100"
            'System.Threading.Thread.Sleep(500)              ' delay
        End If

        Cal3245status.Text = "SETTING UP 3245A...."
        System.Threading.Thread.Sleep(500)              ' delay
        Cal3245status.Text = "STARTING......"
        System.Threading.Thread.Sleep(1000)             ' 1000mS delay
        dev2.SendAsync("USE " & Channel, True)          ' USE 0
        Cal3245status.Text = "3245A - USE" & Channel
        System.Threading.Thread.Sleep(500)              ' delay

        'If Code3245A.Text = "" Then     ' if code is blank then set it to zero. This would probably only happen if the hardware lock is in place and the user thought to just clear the code entry box
        'Code3245A.Text = "0"        ' in which case the code immediately below will ignore it anyways.....I think.
        'End If

        If RadioButton3245ADCV.Checked = True Then
            dev2.SendAsync("CAL VOLTS " & Code3245A.Text, True) ' CAL VOLTS 3245
            Cal3245status.Text = "3245A - CAL DCV " & Code3245A.Text
        Else
            dev2.SendAsync("CAL " & Code3245A.Text, True) ' CAL VOLTS 3245
            Cal3245status.Text = "3245A - CAL DCV && DCI " & Code3245A.Text
        End If

        'Label264.Text = "of 47"

        System.Threading.Thread.Sleep(500)     ' delay

        Cal3245status.Text = "CALIBRATING......."

        ' Now loop for DCV calibration - 1 to 47
        For Calnum As Integer = CalStart To CalEnd      ' Loop

            If Calnum = 45 Then     ' RDG 45 requires 3458A is put into 100Vdc range

                If RadioButton3458A.Checked = True Then
                    dev1.SendAsync("RANGE 100", True)
                    Cal3245status.Text = "3458A - Set to 100Vdc RANGE"
                    Me.Refresh()
                    Thread.Sleep(500)     ' delay
                End If
                If RadioButton344XXA.Checked = True Then
                    dev1.SendAsync("VOLT:DC:RANG 100", True)
                    Cal3245status.Text = "344XXA - Set to 100Vdc RANGE"
                    Me.Refresh()
                    Thread.Sleep(500)     ' delay
                End If
                If RadioButtonR6581.Checked = True Then
                    dev1.SendAsync("VOLT:DC:RANG 100", True)
                    Cal3245status.Text = "R6581 - Set to 100Vdc RANGE"
                    Me.Refresh()
                    Thread.Sleep(500)     ' delay
                End If

            End If

            If Calnum = 46 Then     ' and back to AUTO range again
                If RadioButton3458A.Checked = True Then
                    dev1.SendAsync("RANGE AUTO", True)
                    Cal3245status.Text = "3458A - AUTO RANGE"
                    Me.Refresh()
                    Thread.Sleep(500)     ' delay
                End If
                If RadioButton344XXA.Checked = True Then
                    dev1.SendAsync("VOLT:DC:RANG:AUTO ON", True)
                    Cal3245status.Text = "344xxA - AUTO RANGE"
                    Me.Refresh()
                    Thread.Sleep(500)     ' delay
                End If
                If RadioButtonR6581.Checked = True Then
                    dev1.SendAsync("VOLT:DC:RANG:AUTO ON", True)
                    Cal3245status.Text = "R6581 - AUTO RANGE"
                    Me.Refresh()
                    Thread.Sleep(500)     ' delay
                End If

            End If

            If Abort3245A = True Then   ' abort calibration, reset vars and exit sub
                Abort3245A = False
                Calnum = 1
                Cal3245status.Text = "USER ABORTED......"
                Exit Sub
            End If

            LabelRDG.Text = Calnum
            Read3458A()
            Application.DoEvents()
            Write3245A()
            Application.DoEvents()
        Next

        ' Get ready for DCI calibration
        If RadioButton3245ADCVDCI.Checked = True Then

            If RadioButton3458A.Checked = True Then
                dev1.SendAsync("BEEP", True)
                dev1.SendAsync("BEEP", True)
            End If
            If RadioButton344XXA.Checked = True Then
                dev1.SendAsync("SYST:BEEP", True)
                dev1.SendAsync("SYST:BEEP", True)
            End If
            If RadioButtonR6581.Checked = True Then
                dev1.SendAsync("SYST:BEEP", True)
                dev1.SendAsync("SYST:BEEP", True)
            End If

            Dialog2.Warning1 = "3245A Calibration - User intervention req'd"
            Dialog2.Warning2 = "Switch cables on DMM to 'I' and 'LO' for DCI operation"
            Dialog2.Warning3 = ".......Then hit OK"
            Dialog2.Warning3 = ""
            Dialog2.Warning4 = ""
            Dialog2.Warning5 = ""
            Dialog2.ShowDialog(Me)  ' this method positions centre of parent form, and requires to hit OK to return back to parent

            CalStart = 48
            CalEnd = 71
            'Label264.Text = "of  71"

            If RadioButton3458A.Checked = True Then
                dev1.SendAsync("FUNC DCI", True)       ' DCI
                System.Threading.Thread.Sleep(500)          ' delay
                dev1.SendAsync("NPLC " & TextBoxNPLCDCI.Text, True)       ' DCI NPLC 10
                Cal3245status.Text = "3458A - NPLC " & TextBoxNPLCDCI.Text
            End If
            If RadioButton344XXA.Checked = True Then
                dev1.SendAsync("CONF:CURR:DC", True)       ' DCI
                Cal3245status.Text = "344xxA - DCI OPERATION"
                System.Threading.Thread.Sleep(500)          ' delay
                dev1.SendAsync("CURR:DC:NPLC " & TextBoxNPLCDCI.Text, True)       ' DCI NPLC 10
                Cal3245status.Text = "344xxA - NPLC " & TextBoxNPLCDCI.Text
            End If
            If RadioButtonR6581.Checked = True Then
                dev1.SendAsync("CURR:DC:NPLC " & TextBoxNPLCDCI.Text, True)       ' DCI
                Cal3245status.Text = "R6581 - DCI OPERATION"
            End If

            Cal3245status.Text = "DMM - FUNC DCI"
            Me.Refresh()
            Thread.Sleep(500)     ' delay

            Cal3245status.Text = "DMM - STARTING DCI CALIBRATION"
            Me.Refresh()
            Thread.Sleep(500)     ' delay

        End If

        ' now loop for DCI calibration - 48 to 71
        If RadioButton3245ADCVDCI.Checked = True Then

            For Calnum As Integer = CalStart To CalEnd      ' Loop

                If Abort3245A = True Then   ' abort calibration, reset vars and exit sub
                    Abort3245A = False
                    Calnum = 1
                    Cal3245status.Text = "USER ABORTED......"
                    Exit Sub
                End If

                LabelRDG.Text = Calnum
                Read3458A()
                Application.DoEvents()
                Write3245A()
                Application.DoEvents()
            Next

        End If
    End Sub

    Private Sub RadioButton3245ADCV_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3245ADCV.CheckedChanged

        'RadioButton3245ADCV.Checked = True
        If RadioButton3245ADCV.Checked = True Then
            Label264.Text = "of  47"
        End If

    End Sub

    Private Sub RadioButton3245ADCVDCI_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3245ADCVDCI.CheckedChanged

        'RadioButton3245ADCV.Checked = True
        If RadioButton3245ADCVDCI.Checked = True Then
            Label264.Text = "of  71"
        End If

    End Sub

    Private Sub Read3458A()

        ' Read from 3458A and wait for reply, do it twice so that reading is stable given the NPLC 100
        Label3458123.Text = "Read 1"

        If RadioButton3458A.Checked = True Then
            txtq1b.Text = "TARM SGL"
        End If
        If RadioButton344XXA.Checked = True Then
            txtq1b.Text = "READ?"
        End If
        If RadioButtonR6581.Checked = True Then
            txtq1b.Text = "READ?"
        End If

        query_blockingDMM()

        Label3458ARDG.Text = txtr1b.Text                                ' E-notation

        Application.DoEvents()

        Label3458123.Text = "Read 2"

        If RadioButton3458A.Checked = True Then
            txtq1b.Text = "TARM SGL"
        End If
        If RadioButton344XXA.Checked = True Then
            txtq1b.Text = "READ?"
        End If
        If RadioButtonR6581.Checked = True Then
            txtq1b.Text = "READ?"
        End If

        query_blockingDMM()

        Label3458ARDG.Text = txtr1b.Text                                ' E-notation

        Application.DoEvents()

        Thread.Sleep(500)     ' delay

    End Sub

    Private Sub Write3245A()

        Label3458123.Text = ""
        Cal3245status.Text = "3245A - SENDING"

        'dev2.SendAsync("CAL VALUE " & Val(txtr1a_disp.Text), True)     ' decimal        Comment out this line for testing without writing to 3245A
        'Label3245AWRI.Text = Val(txtr1a_disp.Text)

        dev2.SendAsync("CAL VALUE " & txtr1b.Text, True)                ' E-notation        Comment out this line for testing without writing to 3245A
        Label3245AWRI.Text = txtr1b.Text

        Application.DoEvents()

    End Sub

    Private Sub Button3245Aabort_Click(sender As Object, e As EventArgs) Handles Button3245Aabort.Click

        Abort3245A = True
        Dev1Timeout.Text = TimeoutSave          ' return timout on DEVICE tab back to original
        Exit Sub

    End Sub

    Private Sub query_blockingDMM()

        ' This is basically a copy of the btnq1b_Click SUB from Formtest.vb

        Dim q As IOQuery = Nothing
        Dim result As Double

        Dim s As String

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

        'result = dev1.QueryBlocking(txtq1b.Text & TermStr2(), q, True)      'standard version with IOQuery parameter
        result = dev1.QueryBlocking(txtq1b.Text & TermStr2(), q, False)      'standard version with IOQuery parameter       PDVS2mini sub uses TRUE

        s = "blocking command:'" & q.cmd & "'" & vbCrLf

        If result = 0 Then

            txtr1b.Text = q.ResponseAsString
            s &= "device response time:" & Str(q.timeend.Subtract(q.timestart).TotalSeconds) & " s" & vbCrLf
            s &= "thread wait time:" & Str(q.timestart.Subtract(q.timecall).TotalSeconds) & " s" & vbCrLf

        Else
            s &= "status: " & result & vbCrLf
            s &= q.errmsg
        End If

        txtr1astat.Text = s

    End Sub

    Private Sub RadioButtonDMM3458A_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3458A.CheckedChanged
        If RadioButton3458A.Checked = True Then
            PictureBox8.Visible = True
            PictureBox7.Visible = False
            PictureBox2.Visible = False
            Label274.Text = "3458A Read"
            Me.Refresh()
        End If
    End Sub

    Private Sub RadioButtonDMM344XXA_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton344XXA.CheckedChanged
        If RadioButton344XXA.Checked = True Then
            PictureBox8.Visible = False
            PictureBox7.Visible = True
            PictureBox2.Visible = False
            Label274.Text = "344xxA Read"
            Me.Refresh()
        End If
    End Sub

    Private Sub RadioButtonR6581_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonR6581.CheckedChanged
        If RadioButtonR6581.Checked = True Then
            PictureBox8.Visible = False
            PictureBox7.Visible = False
            PictureBox2.Visible = True
            Label274.Text = "R6581(T) Read"
            Me.Refresh()
        End If
    End Sub


End Class