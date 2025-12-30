Imports IODevices
'Imports System.Threading
'Imports System.Runtime.InteropServices
'Imports System.IO
'Imports System
'Imports System.IO.Ports


Partial Class Formtest

    Dim KeyVoltageTemp As Double
    Dim Dev1SampleCountCopy As Integer

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Try
            ' Format of code to send: <Getdata,1>
            SerialPort1.WriteLine("<Getdata,1>")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub KeyVoltage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles KeyVoltage.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
            SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
            KeyVoltage.Text = ""
        End If
    End Sub

    Private Sub DacZero_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DacZero0.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ' Format of code to send: <DacZero,0>
            DacZero0.Text = "<DacZero," + DacZero0.Text + ">"
            SerialPort1.WriteLine(DacZero0.Text)    ' Write data to the send buffer 
            DacZero0.Text = ""
        End If
    End Sub

    Private Sub DacSpan_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DacSpan0.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ' Format of code to send: <DacSpan,0>
            DacSpan0.Text = "<DacSpan," + DacSpan0.Text + ">"
            SerialPort1.WriteLine(DacSpan0.Text)    ' Write data to the send buffer 
            DacSpan0.Text = ""
        End If
    End Sub

    Private Sub OutVdc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles OutVdc.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ' Format of code to send: <OutVdc,0>
            OutVdc.Text = "<OutVdc," + OutVdc.Text + ">"
            SerialPort1.WriteLine(OutVdc.Text)    ' Write data to the send buffer 
            OutVdc.Text = ""
        End If
    End Sub

    Private Sub BattVdc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles BattVdc.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ' Format of code to send: <BattVdc,0>
            BattVdc.Text = "<BattVdc," + BattVdc.Text + ">"
            SerialPort1.WriteLine(BattVdc.Text)    ' Write data to the send buffer 
            BattVdc.Text = ""
        End If
    End Sub

    Private Sub ChargeI_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ChargeI.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ' Format of code to send: <BattVdc,0>
            ChargeI.Text = "<ChargeI," + ChargeI.Text + ">"
            SerialPort1.WriteLine(ChargeI.Text)    ' Write data to the send buffer 
            ChargeI.Text = ""
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles ButtonKeyVoltage.Click

        'If KeyVoltage.Text.Length = 0 Then            ' Error if there is no send data 
        'MsgBox("Enter a value first", MessageBoxButtons.OK, "SET OUTPUT")
        'Exit Sub                ' Break out of processing 
        'End If
        Try

            If ButtonDev1Run.Text = "Stop" Then         ' Dev 1 must be connected
                LabelCalCount.Text = "Working"
                txtr1aBIG.Text = " Wait......."
                LabelDeltaV.Text = "Wait..."
            End If

            KeyVoltageTemp = CDbl(Val(KeyVoltage.Text))

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
            Call TxLEDon()
            SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
            KeyVoltage.Text = KeyVoltageTemp
        Catch ex As Exception           ' Exception handling 
            KeyVoltage.Text = KeyVoltageTemp
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        If ButtonDev1Run.Text = "Stop" Then             ' Dev 1 must be connected
            ' Request reading from 3458A
            txtq1b.Text = "TARM SGL"
            query_blocking()

            LabelCalCount.Text = "2nd Read"
            System.Threading.Thread.Sleep(1000)

            ' Request reading from 3458A again
            txtq1b.Text = "TARM SGL"
            query_blocking()

            ' Now Update display
            txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

            LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - KeyVoltageTemp) * 1000000
            LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
            'KeyVoltage.Text = ""

            LabelCalCount.Text = "Done"
        End If

    End Sub

    Private Sub ButtonOutVdc_Click(sender As Object, e As EventArgs) Handles ButtonOutVdc.Click
        Try
            ' Format of code to send: <OutVdc,0>
            ' Send 10V to output first
            KeyVoltage.Text = "<KeyVoltage,0," + "10.00000" + ">"
            SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
            KeyVoltage.Text = ""

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish

            OutVdc.Text = LabelOutputVFeedMult.Text
            OutVdc.Text = "<OutVdc," + OutVdc.Text + ">"
            SerialPort1.WriteLine(OutVdc.Text)    ' Write data to the send buffer 
            OutVdc.Text = ""
        Catch ex As Exception           ' Exception handling 
            OutVdc.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        'voltsout.Text = txtr1a_disp.Text
        OutVdc.Text = LabelOutputVFeedMult.Text
    End Sub

    Private Sub ButtonBattVdc_Click(sender As Object, e As EventArgs) Handles ButtonBattVdc.Click
        Try
            ' Format of code to send: <BattVdc,0>
            BattVdc.Text = LabelBatteryVFeedMult.Text
            BattVdc.Text = "<BattVdc," + BattVdc.Text + ">"
            SerialPort1.WriteLine(BattVdc.Text)    ' Write data to the send buffer 
            BattVdc.Text = ""
        Catch ex As Exception           ' Exception handling 
            BattVdc.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        'voltasbatt.Text = txtr1a_disp.Text
        BattVdc.Text = LabelBatteryVFeedMult.Text
    End Sub

    Private Sub ButtonChargeI_Click(sender As Object, e As EventArgs) Handles ButtonChargeI.Click

        Try
            ' Format of code to send: <ChargeI,0>
            BattVdc.Text = LabelBatteryMonICMult.Text
            ChargeI.Text = "<ChargeI," + ChargeI.Text + ">"
            SerialPort1.WriteLine(ChargeI.Text)    ' Write data to the send buffer 
            ChargeI.Text = ""
        Catch ex As Exception           ' Exception handling 
            ChargeI.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        'chargema.Text = txtr1a_disp.Text
        ChargeI.Text = LabelBatteryMonICMult.Text
    End Sub

    Private Sub ButtonOutVdcdown_Click(sender As Object, e As EventArgs) Handles ButtonOutVdcdown.Click

        number = Val(LabelOutputVFeedMult.Text) ' string to int
        number = number - 1
        OutVdc.Text = System.Convert.ToString(number)
        OutVdc.Text = "<OutVdc," + OutVdc.Text + ">"
        SerialPort1.WriteLine(OutVdc.Text)    ' Write data to the send buffer 
        OutVdc.Text = ""

    End Sub

    Private Sub ButtonBattVdcdown_Click(sender As Object, e As EventArgs) Handles ButtonBattVdcdown.Click

        number = Val(LabelBatteryVFeedMult.Text) ' string to int
        number = number - 1
        BattVdc.Text = System.Convert.ToString(number)
        BattVdc.Text = "<BattVdc," + BattVdc.Text + ">"
        SerialPort1.WriteLine(BattVdc.Text)    ' Write data to the send buffer 
        BattVdc.Text = ""

    End Sub

    Private Sub ButtonChargeIdown_Click(sender As Object, e As EventArgs) Handles ButtonChargeIdown.Click

        number = Val(LabelBatteryMonICMult.Text) ' string to int
        number = number - 1
        ChargeI.Text = System.Convert.ToString(number)
        ChargeI.Text = "<ChargeI," + ChargeI.Text + ">"
        SerialPort1.WriteLine(ChargeI.Text)    ' Write data to the send buffer 
        ChargeI.Text = ""

    End Sub

    Private Sub ButtonOutVdcUp_Click(sender As Object, e As EventArgs) Handles ButtonOutVdcUp.Click
        number = Val(LabelOutputVFeedMult.Text) ' string to int
        number = number + 1
        OutVdc.Text = System.Convert.ToString(number)
        OutVdc.Text = "<OutVdc," + OutVdc.Text + ">"
        SerialPort1.WriteLine(OutVdc.Text)    ' Write data to the send buffer 
        OutVdc.Text = ""
    End Sub

    Private Sub ButtonBattVdcUp_Click(sender As Object, e As EventArgs) Handles ButtonBattVdcUp.Click
        number = Val(LabelBatteryVFeedMult.Text) ' string to int
        number = number + 1
        BattVdc.Text = System.Convert.ToString(number)
        BattVdc.Text = "<BattVdc," + BattVdc.Text + ">"
        SerialPort1.WriteLine(BattVdc.Text)    ' Write data to the send buffer 
        BattVdc.Text = ""
    End Sub

    Private Sub ButtonChargeIUp_Click(sender As Object, e As EventArgs) Handles ButtonChargeIUp.Click
        number = Val(LabelBatteryMonICMult.Text) ' string to int
        number = number + 1
        ChargeI.Text = System.Convert.ToString(number)
        ChargeI.Text = "<ChargeI," + ChargeI.Text + ">"
        SerialPort1.WriteLine(ChargeI.Text)    ' Write data to the send buffer 
        ChargeI.Text = ""
    End Sub




    ' SET mV buttons

    Private Sub Button000010_Click(sender As Object, e As EventArgs) Handles Button000010.Click

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.00010"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts000010.Text = txtr1aBIG.Text
        volts000010.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

    End Sub

    Private Sub Button000100_Click(sender As Object, e As EventArgs) Handles Button000100.Click

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.00100"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts000100.Text = txtr1aBIG.Text
        volts000100.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

    End Sub

    Private Sub Button001000_Click(sender As Object, e As EventArgs) Handles Button001000.Click

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.01000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts001000.Text = txtr1aBIG.Text
        volts001000.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

    End Sub

    Private Sub Button010000_Click(sender As Object, e As EventArgs) Handles Button010000.Click

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.10000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts010000.Text = txtr1aBIG.Text
        volts010000.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

    End Sub

    Private Sub Button020000_Click(sender As Object, e As EventArgs) Handles Button020000.Click

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.20000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts020000.Text = txtr1aBIG.Text
        volts020000.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

    End Sub

    Private Sub Button030000_Click(sender As Object, e As EventArgs) Handles Button030000.Click

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.30000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts030000.Text = txtr1aBIG.Text
        volts030000.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

    End Sub

    Private Sub Button050000_Click(sender As Object, e As EventArgs) Handles Button050000.Click

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.50000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts050000.Text = txtr1aBIG.Text
        volts050000.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

    End Sub







    Private Sub ButtonDacZero0down_Click(sender As Object, e As EventArgs) Handles ButtonDacZero0down.Click
        number = Val(LabeldacZero0Cal.Text) ' string to int
        number = number - 1
        DacZero0.Text = System.Convert.ToString(number)
        DacZero0.Text = "<DacZero0," + DacZero0.Text + ">"
        SerialPort1.WriteLine(DacZero0.Text)    ' Write data to the send buffer 
        DacZero0.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "0.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacZero0Up_Click(sender As Object, e As EventArgs) Handles ButtonDacZero0Up.Click
        number = Val(LabeldacZero0Cal.Text) ' string to int
        number = number + 1
        DacZero0.Text = System.Convert.ToString(number)
        DacZero0.Text = "<DacZero0," + DacZero0.Text + ">"
        SerialPort1.WriteLine(DacZero0.Text)    ' Write data to the send buffer 
        DacZero0.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "0.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan0down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan0down.Click
        number = Val(LabeldacSpan0Cal.Text) ' string to int
        number = number - 1
        DacSpan0.Text = System.Convert.ToString(number)
        DacSpan0.Text = "<DacSpan0," + DacSpan0.Text + ">"
        SerialPort1.WriteLine(DacSpan0.Text)    ' Write data to the send buffer 
        DacSpan0.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "1.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan0Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan0Up.Click
        number = Val(LabeldacSpan0Cal.Text) ' string to int
        number = number + 1
        DacSpan0.Text = System.Convert.ToString(number)
        DacSpan0.Text = "<DacSpan0," + DacSpan0.Text + ">"
        SerialPort1.WriteLine(DacSpan0.Text)    ' Write data to the send buffer 
        DacSpan0.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "1.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan1down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan1down.Click
        number = Val(LabeldacSpan1Cal.Text) ' string to int
        number = number - 1
        DacSpan1.Text = System.Convert.ToString(number)
        DacSpan1.Text = "<DacSpan1," + DacSpan1.Text + ">"
        SerialPort1.WriteLine(DacSpan1.Text)    ' Write data to the send buffer 
        DacSpan1.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "2.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan1Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan1Up.Click
        number = Val(LabeldacSpan1Cal.Text) ' string to int
        number = number + 1
        DacSpan1.Text = System.Convert.ToString(number)
        DacSpan1.Text = "<DacSpan1," + DacSpan1.Text + ">"
        SerialPort1.WriteLine(DacSpan1.Text)    ' Write data to the send buffer 
        DacSpan1.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "2.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan2down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan2down.Click
        number = Val(LabeldacSpan2Cal.Text) ' string to int
        number = number - 1
        DacSpan2.Text = System.Convert.ToString(number)
        DacSpan2.Text = "<DacSpan2," + DacSpan2.Text + ">"
        SerialPort1.WriteLine(DacSpan2.Text)    ' Write data to the send buffer 
        DacSpan2.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "3.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan2Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan2Up.Click
        number = Val(LabeldacSpan2Cal.Text) ' string to int
        number = number + 1
        DacSpan2.Text = System.Convert.ToString(number)
        DacSpan2.Text = "<DacSpan2," + DacSpan2.Text + ">"
        SerialPort1.WriteLine(DacSpan2.Text)    ' Write data to the send buffer 
        DacSpan2.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "3.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan3down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan3down.Click
        number = Val(LabeldacSpan3Cal.Text) ' string to int
        number = number - 1
        DacSpan3.Text = System.Convert.ToString(number)
        DacSpan3.Text = "<DacSpan3," + DacSpan3.Text + ">"
        SerialPort1.WriteLine(DacSpan3.Text)    ' Write data to the send buffer 
        DacSpan3.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "4.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan3Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan3Up.Click
        number = Val(LabeldacSpan3Cal.Text) ' string to int
        number = number + 1
        DacSpan3.Text = System.Convert.ToString(number)
        DacSpan3.Text = "<DacSpan3," + DacSpan3.Text + ">"
        SerialPort1.WriteLine(DacSpan3.Text)    ' Write data to the send buffer 
        DacSpan3.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "4.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan4down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan4down.Click
        number = Val(LabeldacSpan4Cal.Text) ' string to int
        number = number - 1
        DacSpan4.Text = System.Convert.ToString(number)
        DacSpan4.Text = "<DacSpan4," + DacSpan4.Text + ">"
        SerialPort1.WriteLine(DacSpan4.Text)    ' Write data to the send buffer 
        DacSpan4.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "5.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan4Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan4Up.Click
        number = Val(LabeldacSpan4Cal.Text) ' string to int
        number = number + 1
        DacSpan4.Text = System.Convert.ToString(number)
        DacSpan4.Text = "<DacSpan4," + DacSpan4.Text + ">"
        SerialPort1.WriteLine(DacSpan4.Text)    ' Write data to the send buffer 
        DacSpan4.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "5.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan5down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan5down.Click
        number = Val(LabeldacSpan5Cal.Text) ' string to int
        number = number - 1
        DacSpan5.Text = System.Convert.ToString(number)
        DacSpan5.Text = "<DacSpan5," + DacSpan5.Text + ">"
        SerialPort1.WriteLine(DacSpan5.Text)    ' Write data to the send buffer 
        DacSpan5.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "6.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan5Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan5Up.Click
        number = Val(LabeldacSpan5Cal.Text) ' string to int
        number = number + 1
        DacSpan5.Text = System.Convert.ToString(number)
        DacSpan5.Text = "<DacSpan5," + DacSpan5.Text + ">"
        SerialPort1.WriteLine(DacSpan5.Text)    ' Write data to the send buffer 
        DacSpan5.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "6.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan6down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan6down.Click
        number = Val(LabeldacSpan6Cal.Text) ' string to int
        number = number - 1
        DacSpan6.Text = System.Convert.ToString(number)
        DacSpan6.Text = "<DacSpan6," + DacSpan6.Text + ">"
        SerialPort1.WriteLine(DacSpan6.Text)    ' Write data to the send buffer 
        DacSpan6.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "7.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan6Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan6Up.Click
        number = Val(LabeldacSpan6Cal.Text) ' string to int
        number = number + 1
        DacSpan6.Text = System.Convert.ToString(number)
        DacSpan6.Text = "<DacSpan6," + DacSpan6.Text + ">"
        SerialPort1.WriteLine(DacSpan6.Text)    ' Write data to the send buffer 
        DacSpan6.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "7.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan7down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan7down.Click
        number = Val(LabeldacSpan7Cal.Text) ' string to int
        number = number - 1
        DacSpan7.Text = System.Convert.ToString(number)
        DacSpan7.Text = "<DacSpan7," + DacSpan7.Text + ">"
        SerialPort1.WriteLine(DacSpan7.Text)    ' Write data to the send buffer 
        DacSpan7.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "8.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan7Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan7Up.Click
        number = Val(LabeldacSpan7Cal.Text) ' string to int
        number = number + 1
        DacSpan7.Text = System.Convert.ToString(number)
        DacSpan7.Text = "<DacSpan7," + DacSpan7.Text + ">"
        SerialPort1.WriteLine(DacSpan7.Text)    ' Write data to the send buffer 
        DacSpan7.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "8.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan8down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan8down.Click
        number = Val(LabeldacSpan8Cal.Text) ' string to int
        number = number - 1
        DacSpan8.Text = System.Convert.ToString(number)
        DacSpan8.Text = "<DacSpan8," + DacSpan8.Text + ">"
        SerialPort1.WriteLine(DacSpan8.Text)    ' Write data to the send buffer 
        DacSpan8.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "9.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan8Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan8Up.Click
        number = Val(LabeldacSpan8Cal.Text) ' string to int
        number = number + 1
        DacSpan8.Text = System.Convert.ToString(number)
        DacSpan8.Text = "<DacSpan8," + DacSpan8.Text + ">"
        SerialPort1.WriteLine(DacSpan8.Text)    ' Write data to the send buffer 
        DacSpan8.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "9.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan9down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan9down.Click
        number = Val(LabeldacSpan9Cal.Text) ' string to int
        number = number - 1
        DacSpan9.Text = System.Convert.ToString(number)
        DacSpan9.Text = "<DacSpan9," + DacSpan9.Text + ">"
        SerialPort1.WriteLine(DacSpan9.Text)    ' Write data to the send buffer 
        DacSpan9.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "10.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan10down_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan10down.Click
        number = Val(LabeldacSpan10Cal.Text) ' string to int
        number = number - 1
        DacSpan10.Text = System.Convert.ToString(number)
        DacSpan10.Text = "<DacSpan10," + DacSpan10.Text + ">"
        SerialPort1.WriteLine(DacSpan10.Text)    ' Write data to the send buffer 
        DacSpan10.Text = ""

        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "10.22222"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""
    End Sub

    Private Sub ButtonDacSpan9Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan9Up.Click
        number = Val(LabeldacSpan9Cal.Text) ' string to int
        number = number + 1
        DacSpan9.Text = System.Convert.ToString(number)
        DacSpan9.Text = "<DacSpan9," + DacSpan9.Text + ">"
        SerialPort1.WriteLine(DacSpan9.Text)    ' Write data to the send buffer 
        DacSpan9.Text = ""
        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "10.00000"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""

    End Sub

    Private Sub ButtonDacSpan10Up_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan10Up.Click
        number = Val(LabeldacSpan10Cal.Text) ' string to int
        number = number + 1
        DacSpan10.Text = System.Convert.ToString(number)
        DacSpan10.Text = "<DacSpan10," + DacSpan10.Text + ">"
        SerialPort1.WriteLine(DacSpan10.Text)    ' Write data to the send buffer 
        DacSpan10.Text = ""
        ' Format of code to send: <KeyVoltage,0,1.23456>
        KeyVoltage.Text = "10.22222"
        KeyVoltage.Text = "<KeyVoltage,0," + KeyVoltage.Text + ">"
        SerialPort1.WriteLine(KeyVoltage.Text)    ' Write data to the send buffer 
        KeyVoltage.Text = ""

    End Sub


    ' SET buttons

    Private Sub ButtonButtonDacZero0_Click(sender As Object, e As EventArgs) Handles ButtonDacZero0.Click

        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacZero0,0>
            DacZero0.Text = LabeldacZero0Cal.Text
            SerialPort1.WriteLine("<DacZero0," + DacZero0.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacZero0.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts0.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacZero0.Text = LabeldacZero0Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts0.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacZero0.Text = LabeldacZero0Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts0.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacZero0.Text = LabeldacZero0Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)

        LabeldacZero0Delta.Text = LabelDeltaV.Text

        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan0_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan0.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "1.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan0,0>
            DacSpan0.Text = LabeldacSpan0Cal.Text
            SerialPort1.WriteLine("<DacSpan0," + DacSpan0.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan0.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts1.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan0.Text = LabeldacSpan0Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts1.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan0.Text = LabeldacSpan0Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts1.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan0.Text = LabeldacSpan0Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan0Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan1_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan1.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "2.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan1,0>
            DacSpan1.Text = LabeldacSpan1Cal.Text
            SerialPort1.WriteLine("<DacSpan1," + DacSpan1.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan1.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts2.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan1.Text = LabeldacSpan1Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts2.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan1.Text = LabeldacSpan1Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts2.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan1.Text = LabeldacSpan1Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan1Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan2_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan2.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "3.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan2,0>
            DacSpan2.Text = LabeldacSpan2Cal.Text
            SerialPort1.WriteLine("<DacSpan2," + DacSpan2.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan2.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts3.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan2.Text = LabeldacSpan2Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts3.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan2.Text = LabeldacSpan2Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts3.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan2.Text = LabeldacSpan2Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan2Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan3_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan3.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "4.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan3,0>
            DacSpan3.Text = LabeldacSpan3Cal.Text
            SerialPort1.WriteLine("<DacSpan3," + DacSpan3.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan3.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts4.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan3.Text = LabeldacSpan3Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts4.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan3.Text = LabeldacSpan3Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts4.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan3.Text = LabeldacSpan3Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan3Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan4_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan4.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "5.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan4,0>
            DacSpan4.Text = LabeldacSpan4Cal.Text
            SerialPort1.WriteLine("<DacSpan4," + DacSpan4.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan4.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts5.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan4.Text = LabeldacSpan4Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts5.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan4.Text = LabeldacSpan4Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts5.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan4.Text = LabeldacSpan4Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan4Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan5_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan5.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "6.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan5,0>
            DacSpan5.Text = LabeldacSpan5Cal.Text
            SerialPort1.WriteLine("<DacSpan5," + DacSpan5.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan5.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts6.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan5.Text = LabeldacSpan5Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts6.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan5.Text = LabeldacSpan5Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts6.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan5.Text = LabeldacSpan5Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan5Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan6_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan6.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "7.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan6,0>
            DacSpan6.Text = LabeldacSpan6Cal.Text
            SerialPort1.WriteLine("<DacSpan6," + DacSpan6.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan6.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts7.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan6.Text = LabeldacSpan6Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts7.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan6.Text = LabeldacSpan6Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts7.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan6.Text = LabeldacSpan6Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan6Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan7_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan7.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "8.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan7,0>
            DacSpan7.Text = LabeldacSpan7Cal.Text
            SerialPort1.WriteLine("<DacSpan7," + DacSpan7.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan7.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts8.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan7.Text = LabeldacSpan7Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts8.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan7.Text = LabeldacSpan7Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts8.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan7.Text = LabeldacSpan7Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan7Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan8_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan8.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "9.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan8,0>
            DacSpan8.Text = LabeldacSpan8Cal.Text
            SerialPort1.WriteLine("<DacSpan8," + DacSpan8.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan8.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts9.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan8.Text = LabeldacSpan8Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts9.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan8.Text = LabeldacSpan8Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts9.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan8.Text = LabeldacSpan8Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan8Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan9_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan9.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "10.00000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan9,0>
            DacSpan9.Text = LabeldacSpan9Cal.Text
            SerialPort1.WriteLine("<DacSpan9," + DacSpan9.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan9.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts10.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan9.Text = LabeldacSpan9Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts10.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan9.Text = LabeldacSpan9Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts10.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan9.Text = LabeldacSpan9Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan9Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub ButtonButtonDacSpan10_Click(sender As Object, e As EventArgs) Handles ButtonDacSpan10.Click
        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "10.22222"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan10,0>
            DacSpan9.Text = LabeldacSpan10Cal.Text
            SerialPort1.WriteLine("<DacSpan10," + DacSpan10.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan10.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts11.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan10.Text = LabeldacSpan10Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts11.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan10.Text = LabeldacSpan10Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts11.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan10.Text = LabeldacSpan10Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan10Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub SavePDVS2Eprom_Click(sender As Object, e As EventArgs) Handles SavePDVS2Eprom.Click

        ' There is a bug in the PDVS2mini uP code which means this eprom save doesn't work properly. Only works with PDVS2mini code V1.03 (Dec. 2020) onwards.

        LabelCalCount.Text = "SAVING CAL TO EEPROM"
        Me.Refresh()

        ' Send SAVE command to PDVS2 - Format of code to send: <OutVdc,0>
        Call TxLEDon()
        SerialPort1.WriteLine("<SaveCal," + "0" + ">")    ' Write data to the send buffer 

        System.Threading.Thread.Sleep(1000) ' delay 1sec

        LabelCalCount.Text = "Done"
    End Sub

    Private Sub query_blocking()

        ' This is basically a copy of the btnq1b_Click SUB from Formtest.vb

        Dim q As IOQuery = Nothing
        Dim result As Integer

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

        result = dev1.QueryBlocking(txtq1b.Text, q, True) 'standard version with IOQuery parameter
        'result = dev1.QueryBlocking(txtq1b.Text & TermStr2(), q, True) 'simpler version with string parameter, modified for BB3 operation

        s = "blocking command:'" & q.cmd & "'" & vbCrLf

        If result = 0 Then

            txtr1b.Text = respNorm
            s &= "device response time:" & Str(q.timeend.Subtract(q.timestart).TotalSeconds) & " s" & vbCrLf
            s &= "thread wait time:" & Str(q.timestart.Subtract(q.timecall).TotalSeconds) & " s" & vbCrLf

        Else
            s &= "status: " & result & vbCrLf
            s &= q.errmsg
        End If

        txtr1astat.Text = s

    End Sub


    ' Auto mV button (multi)

    Private Sub ButtonAutomV_Click(sender As Object, e As EventArgs) Handles ButtonAutomV.Click

        Try
            LabelCalCount.Text = "Starting Auto mV..."
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.00010"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts000010.Text = txtr1aBIG.Text
        volts000010.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.00100"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts000100.Text = txtr1aBIG.Text
        volts000100.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.01000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts001000.Text = txtr1aBIG.Text
        volts001000.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.10000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts010000.Text = txtr1aBIG.Text
        volts010000.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.20000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts020000.Text = txtr1aBIG.Text
        volts020000.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.30000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts030000.Text = txtr1aBIG.Text
        volts030000.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

        Try
            LabelCalCount.Text = "Working"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.50000"
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
        Catch ex As Exception           ' Exception handling 
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        'volts050000.Text = txtr1aBIG.Text
        volts050000.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

        ' Play sound to indicate finished
        My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Asterisk)

    End Sub


    ' Auto run through all X-Y SET buttons

    Private Sub ButtonAutoSET_Click(sender As Object, e As EventArgs) Handles ButtonAutoSET.Click

        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "0.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacZero0,0>
            DacZero0.Text = LabeldacZero0Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacZero0," + DacZero0.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacZero0.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts0.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacZero0.Text = LabeldacZero0Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts0.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacZero0.Text = LabeldacZero0Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts0.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacZero0.Text = LabeldacZero0Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)

        LabeldacZero0Delta.Text = LabelDeltaV.Text

        'KeyVoltage.Text = ""

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "1.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan0,0>
            DacSpan0.Text = LabeldacSpan0Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan0," + DacSpan0.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan0.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts1.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan0.Text = LabeldacSpan0Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts1.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan0.Text = LabeldacSpan0Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts1.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan0.Text = LabeldacSpan0Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan0Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "2.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan1,0>
            DacSpan1.Text = LabeldacSpan1Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan1," + DacSpan1.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan1.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts2.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan1.Text = LabeldacSpan1Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts2.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan1.Text = LabeldacSpan1Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts2.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan1.Text = LabeldacSpan1Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan1Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "3.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan2,0>
            DacSpan2.Text = LabeldacSpan2Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan2," + DacSpan2.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan2.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts3.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan2.Text = LabeldacSpan2Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts3.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan2.Text = LabeldacSpan2Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts3.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan2.Text = LabeldacSpan2Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan2Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "4.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan3,0>
            DacSpan3.Text = LabeldacSpan3Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan3," + DacSpan3.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan3.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts4.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan3.Text = LabeldacSpan3Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts4.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan3.Text = LabeldacSpan3Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts4.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan3.Text = LabeldacSpan3Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan3Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "5.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan4,0>
            DacSpan4.Text = LabeldacSpan4Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan4," + DacSpan4.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan4.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts5.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan4.Text = LabeldacSpan4Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts5.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan4.Text = LabeldacSpan4Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts5.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan4.Text = LabeldacSpan4Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan4Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "6.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan5,0>
            DacSpan5.Text = LabeldacSpan5Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan5," + DacSpan5.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan5.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts6.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan5.Text = LabeldacSpan5Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts6.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan5.Text = LabeldacSpan5Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts6.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan5.Text = LabeldacSpan5Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan5Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "7.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan6,0>
            DacSpan6.Text = LabeldacSpan6Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan6," + DacSpan6.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan6.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts7.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan6.Text = LabeldacSpan6Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts7.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan6.Text = LabeldacSpan6Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts7.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan6.Text = LabeldacSpan6Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan6Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "8.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan7,0>
            DacSpan7.Text = LabeldacSpan7Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan7," + DacSpan7.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan7.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts8.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan7.Text = LabeldacSpan7Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts8.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan7.Text = LabeldacSpan7Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts8.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan7.Text = LabeldacSpan7Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan7Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "9.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan8,0>
            DacSpan8.Text = LabeldacSpan8Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan8," + DacSpan8.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan8.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts9.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan8.Text = LabeldacSpan8Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts9.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan8.Text = LabeldacSpan8Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts9.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan8.Text = LabeldacSpan8Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan8Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"

        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "10.00000"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan9,0>
            DacSpan9.Text = LabeldacSpan9Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan9," + DacSpan9.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan9.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts10.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan9.Text = LabeldacSpan9Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts10.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan9.Text = LabeldacSpan9Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts10.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan9.Text = LabeldacSpan9Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan9Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"


        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))



        Try
            LabelCalCount.Text = "1st Read"
            txtr1aBIG.Text = " Wait......."
            LabelDeltaV.Text = "Wait..."

            ' Stop the GPIB timer if it's running
            ButtonDev1Run.Text = "Run"
            Me.Timer2.Stop()

            ' Format of code to send: <KeyVoltage,0,1.23456>
            KeyVoltage.Text = "10.22222"
            Call TxLEDon()
            SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer 
            'KeyVoltage.Text = ""
            System.Threading.Thread.Sleep(Val(PDVS2delay.Text) * 8)  ' give PDVS2mini time to finish
            ' Format of code to send: <DacSpan9,0>
            DacSpan10.Text = LabeldacSpan10Cal.Text
            Call TxLEDon()
            SerialPort1.WriteLine("<DacSpan10," + DacSpan10.Text + ">")    ' Write data to the send buffer 
        Catch ex As Exception           ' Exception handling 
            DacSpan9.Text = ""
            MsgBox("Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'System.Threading.Thread.Sleep(500)

        ' Request reading from 3458A
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts11.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan10.Text = LabeldacSpan10Cal.Text

        LabelCalCount.Text = "2nd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts11.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan10.Text = LabeldacSpan10Cal.Text

        LabelCalCount.Text = "3rd Read"
        System.Threading.Thread.Sleep(1000)

        ' Request reading from 3458A again
        txtq1b.Text = "TARM SGL"
        query_blocking()

        ' Now Update vars
        volts11.Text = Format(Val(txtr1b.Text), "#00.0000000") ' Convert from E-notation
        DacSpan10.Text = LabeldacSpan10Cal.Text

        txtr1aBIG.Text = Format(Val(txtr1b.Text), "#00.000000000") ' Convert from E-notation

        LabelDeltaV.Text = (CDbl(Val(txtr1aBIG.Text)) - CDbl(Val(KeyVoltage.Text))) * 1000000
        LabelDeltaV.Text = Math.Round(CDbl(Val(LabelDeltaV.Text)), 1)
        'KeyVoltage.Text = ""

        LabeldacSpan10Delta.Text = LabelDeltaV.Text

        LabelCalCount.Text = "Done"






        ' Play sound to indicate finished
        My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Asterisk)

    End Sub

End Class