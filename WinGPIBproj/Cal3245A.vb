﻿
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
                Dialog2.Warning4 = " - Input to voltage on 3458A"
                Dialog2.Warning5 = ".......Then hit OK"
                Dialog2.ShowDialog(Me)  ' this method positions centre of parent form, and requires to hit OK to return back to parent
                CalStart = 1
                CalEnd = 47
                Cal3245A_Channel("0")       ' 3245A Channel 0
            End If

            If CheckBoxChB.Checked = True Then
                dev1.SendAsync("BEEP", True)
                dev1.SendAsync("BEEP", True)
                Dialog2.Warning1 = "3245A Calibration - User intervention req'd"
                Dialog2.Warning2 = "Ensure the following connections:"
                Dialog2.Warning3 = " - Source to Ch B on 3245A"
                Dialog2.Warning4 = " - Input to voltage on 3458A"
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

        dev1.SendAsync("RESET", True)                   ' DCV
        Cal3245status.Text = "3458A - RESET"
        dev1.SendAsync("FUNC DCV", True)                ' DCV
        Cal3245status.Text = "3458A - FUNC DCV"

        dev1.SendAsync("RANGE AUTO", True)              ' RANGE AUTO
        Cal3245status.Text = "3458A - AUTO RANGE"

        dev1.SendAsync("END ALWAYS", True)              ' END ALWAYS
        Cal3245status.Text = "3458A - END ALWAYS"
        System.Threading.Thread.Sleep(500)              ' delay

        If CheckBoxAZERO.Checked = False Then
            dev1.SendAsync("AZERO OFF", True)           ' AZERO OFF - Disabled autozero, speeds up readings......this helps if the GPIB fails, sometimes at RDG 45!
            Cal3245status.Text = "3458A - AZERO OFF"
            System.Threading.Thread.Sleep(500)          ' delay
        Else
            dev1.SendAsync("AZERO ON", True)            ' AZERO ON (default 3458A)
            Cal3245status.Text = "3458A - AZERO ON"
            System.Threading.Thread.Sleep(500)          ' delay
        End If

        dev1.SendAsync("NRDGS 1", True)                 ' NRDGS 1
        Cal3245status.Text = "3458A - NRDGS 1"
        System.Threading.Thread.Sleep(500)              ' delay

        dev1.SendAsync("NPLC 100", True)                ' NPLC 100
        Cal3245status.Text = "3458A - NPLC 100"
        System.Threading.Thread.Sleep(500)              ' delay

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

        Label264.Text = "of 47"

        System.Threading.Thread.Sleep(500)     ' delay

        Cal3245status.Text = "CALIBRATING......."

        ' Now loop for DCV calibration - 1 to 47
        For Calnum As Integer = CalStart To CalEnd      ' Loop

            If Calnum = 45 Then     ' RDG 45 requires 3458A is put into 100Vdc range
                dev1.SendAsync("RANGE 100", True)       ' END ALWAYS
                Cal3245status.Text = "3458A - Set to 100Vdc RANGE"
                Me.Refresh()
                Thread.Sleep(500)     ' delay
            End If
            If Calnum = 46 Then     ' and back to AUTO range again
                dev1.SendAsync("RANGE AUTO", True)       ' END ALWAYS
                Cal3245status.Text = "3458A - AUTO RANGE"
                Me.Refresh()
                Thread.Sleep(500)     ' delay
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

            dev1.SendAsync("BEEP", True)
            dev1.SendAsync("BEEP", True)
            Dialog2.Warning1 = "3245A Calibration - User intervention req'd"
            Dialog2.Warning2 = "Switch cables on 3458A to 'I' and 'LO' for DCI operation"
            Dialog2.Warning3 = ".......Then hit OK"
            Dialog2.Warning3 = ""
            Dialog2.Warning4 = ""
            Dialog2.Warning5 = ""
            Dialog2.ShowDialog(Me)  ' this method positions centre of parent form, and requires to hit OK to return back to parent

            CalStart = 48
            CalEnd = 71
            Label264.Text = "of 71"

            dev1.SendAsync("FUNC DCI", True)       ' DCI
            Cal3245status.Text = "3458A - FUNC DCI"
            Me.Refresh()
            Thread.Sleep(500)     ' delay

            Cal3245status.Text = "3458A - STARTING DCI CALIBRATION"
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

    End Sub

    Private Sub RadioButton3245ADCVDCI_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3245ADCVDCI.CheckedChanged

        'RadioButton3245ADCV.Checked = True

    End Sub

    Private Sub Read3458A()

        ' Read from 3458A and wait for reply, do it twice so that reading is stable given the NPLC 100
        Label3458123.Text = "Read 1"
        txtq1b.Text = "TARM SGL"
        query_blocking3458A()

        Label3458ARDG.Text = txtr1b.Text                                ' E-notation

        Application.DoEvents()

        Label3458123.Text = "Read 2"
        txtq1b.Text = "TARM SGL"
        query_blocking3458A()

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

    Private Sub query_blocking3458A()

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


End Class