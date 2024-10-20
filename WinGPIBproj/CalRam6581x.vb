﻿
' CalRam Advantest R6581 DMM

Imports IODevices
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.IO

Partial Class Formtest

    ' General info..........................
    ' ENTER SERVICE MODE
    ' CAL:EXT:EEPROM:PROTECTION 1

    ' GET ALL CURRENT DCV CALRAM CONSTANTS
    ' CAL:INT:DCV:EEPROM:NEW?

    ' GET ALL CURRENT OHM CALRAM CONSTANTS
    ' CAL:INT:OHM:EEPROM:NEW?

    ' GET ALL CURRENT AC CALRAM CONSTANTS
    ' CAL:INT:AC:EEPROM:NEW?

    ' EXIT SERVICE MODE
    ' CAL:EXT:EEPROM:PROTECTION 0 

    ' :INIT:CONT ON 
    ' CONF:VOLT:DC 
    ' VOLT:DC:RANG 10;NPLC 10 
    ' :SENS:VOLT:DC:DIG 8 

    Dim fs6581 As System.IO.FileStream
    Dim fs26581 As System.IO.FileStream
    Dim CalRamPathfile6581 As System.IO.BinaryWriter
    Dim CalRamPathfile26581 As System.IO.BinaryWriter
    Dim c6581 As Char
    Dim CalramStore6581(500) As String
    'Dim Counter6581 As Integer = 1

    '3458A
    Dim Abort6581 As Boolean = False
    Dim Stepsize6581 As Integer = 2
    Dim RamType6581 As String = ""
    Dim RamType26581 As String = ""
    Dim RAMfilename6581 As String
    Dim RAMfilename26581 As String
    Dim CalramAddress6581 As Integer
    Dim Calrambytefordisplay6581 As String
    Dim CalramStoreTemp16581 As String = ""
    Dim CalramStoreTemp26581 As String = ""
    Dim Counter65812 As Integer = 1
    Dim CalramValue6581 As String = ""
    Dim CalAddrStart6581 As Integer = 393216
    Dim CalAddrEnd6581 As Integer = 397311


    Private Sub ButtonCalramDumpR6581_Click(sender As Object, e As EventArgs) Handles ButtonCalramDumpR6581.Click

        'R6581

        ' run appropriate routine

        If AllFactoryConstantsReadR6581.Checked = True Then    ' All factory calibration constants
            TextBoxCalRamFile6581.Text = ""
            CalFactoryramextractR6581()
        End If

        If AllRegularConstantsReadR6581.Checked = True Then    ' All regular calibration constants
            TextBoxCalRamFile6581.Text = ""
            CalRegularramextractR6581()
        End If


    End Sub


    Private Sub CalFactoryramextractR6581()

        ' R6581 all FACTORY calibration constants read to file



    End Sub


    Private Sub CalRegularramextractR6581()

        ' R6581 all REGULAR calibration constants read to file

        Abort6581 = False

        CalramStatus6581.Text = "CHECKING SETUP"

        Me.Refresh()

        If ButtonDev1Run.Enabled = True Then      ' Device 1 is started

            Dim i As Integer
            Dim s As String
            Dim RAMfilename6581 As String

            CalramStatus6581.Text = "SETTING UP GPIB"
            System.Threading.Thread.Sleep(500)     ' 500mS delay
            Me.Refresh()

            dev1.SendAsync("*RST", True)         ' NPLC 0
            CalramStatus6581.Text = "*RST"
            System.Threading.Thread.Sleep(3000)     ' 3sec delay
            Me.Refresh()

            dev1.SendAsync(":VOLT:DC:DIG MAX", True)         ' NPLC 0
            CalramStatus6581.Text = ":VOLT:DC:DIG MAX"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            dev1.SendAsync(":SENS:VOLT:DC:RANG 10", True)         ' NPLC 0
            CalramStatus6581.Text = ":SENS:VOLT:DC:RANG 10"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            dev1.SendAsync(":VOLT:DC:NPLC 1", True)         ' NPLC 0
            CalramStatus6581.Text = ":VOLT:DC:NPLC 1"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            txtr1a.Text = ""                       ' Prepare reply textbox as empty

            RAMfilename6581 = strPath & "\" & "R6581_Regular_Calibration_Constants" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".txt"

            'LabelCounter6581.Text = "0"
            'Counter6581 = 0
            TextBoxCalRamFile6581.Text = RAMfilename6581

            If Not IO.File.Exists(RAMfilename6581) Then
                IO.File.Create(RAMfilename6581).Dispose()
            End If

            dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 1", True)         ' Enable Service EXT protection mode.....not required for INT commands
            Dev1TextResponse.Checked = True

            ' Write header to text file
            IO.File.AppendAllText(RAMfilename6581, " Regular EXT:ZERO:FRONT:EEPROM Calibration Constants:" & Environment.NewLine)

            Threading.Thread.Sleep(200)

            ' FRONT INPUTS ZERO CALIBRATION
            For i = 0 To 46

                CalramStatus6581.Text = "READING Regular EXT:ZERO:FRONT:EEPROM Calibration Constants"

                ' Blind command
                dev1.SendAsync("CAL:EXT:ZERO:FRONT:NUMBER " & i & "," & i, True)      ' I.E. Limit the range to the 5th constant only  CAL:INT:DCV:HOSEI:NUMBER 5,5

                Threading.Thread.Sleep(10)

                ' New commmand, get reply, store it in array
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("CAL:EXT:ZERO:FRONT:EEPROM:NEW?", q, False)                               ' Get the contents of the range specified above example, so just the 5th
                Cbdev1(q)

                ' Have got the reply, store it in array
                LabelCalRamByte6581.Text = q.ResponseAsString

                ' Write to text file
                IO.File.AppendAllText(RAMfilename6581, q.ResponseAsString & Environment.NewLine)

                Threading.Thread.Sleep(10)

            Next

            Threading.Thread.Sleep(100)

            IO.File.AppendAllText(RAMfilename6581, Environment.NewLine)
            IO.File.AppendAllText(RAMfilename6581, " Regular EXT:ZERO:REAR:EEPROM Calibration Constants:" & Environment.NewLine)

            ' REAR INPUTS ZERO CALIBRATION
            For i = 100 To 146

                CalramStatus6581.Text = "READING Regular EXT:ZERO:REAR:EEPROM Calibration Constants"

                ' Blind command
                dev1.SendAsync("CAL:EXT:ZERO:REAR:NUMBER " & i & "," & i, True)      ' I.E. Limit the range to the 5th constant only  CAL:INT:DCV:HOSEI:NUMBER 5,5

                Threading.Thread.Sleep(10)

                ' New commmand, get reply, store it in array
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("CAL:EXT:ZERO:REAR:EEPROM:NEW?", q, False)                               ' Get the contents of the range specified above example, so just the 5th
                Cbdev1(q)

                ' Have got the reply, store it in array
                LabelCalRamByte6581.Text = q.ResponseAsString

                Me.Refresh()

                ' Write to text file
                IO.File.AppendAllText(RAMfilename6581, txtr1a.Text & Environment.NewLine)

                Threading.Thread.Sleep(10)

            Next

            Threading.Thread.Sleep(100)

            IO.File.AppendAllText(RAMfilename6581, Environment.NewLine)
            IO.File.AppendAllText(RAMfilename6581, " Regular EXT:DCV:EEPROM Calibration Constants:" & Environment.NewLine)

            ' DCV CALIBRATION
            For i = 200 To 203

                CalramStatus6581.Text = "READING Regular EXT:DCV:EEPROM Calibration Constants"

                ' Blind command
                dev1.SendAsync("CAL:EXT:DCV:NUMBER " & i & "," & i, True)      ' I.E. Limit the range to the 5th constant only  CAL:INT:DCV:HOSEI:NUMBER 5,5

                Threading.Thread.Sleep(10)

                ' New commmand, get reply, store it in array
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("CAL:EXT:DCV:EEPROM:NEW?", q, False)                               ' Get the contents of the range specified above example, so just the 5th
                Cbdev1(q)

                ' Have got the reply, store it in array
                LabelCalRamByte6581.Text = q.ResponseAsString

                Me.Refresh()

                ' Write to text file
                IO.File.AppendAllText(RAMfilename6581, txtr1a.Text & Environment.NewLine)

                Threading.Thread.Sleep(10)

            Next

            Threading.Thread.Sleep(100)

            IO.File.AppendAllText(RAMfilename6581, Environment.NewLine)
            IO.File.AppendAllText(RAMfilename6581, " Regular INT:DCV:EEPROM Calibration Constants:" & Environment.NewLine)

            ' DCV CALIBRATION
            For i = 400 To 406

                CalramStatus6581.Text = "READING Regular INT:DCV:EEPROM Calibration Constants"

                ' Blind command
                dev1.SendAsync("CAL:INT:DCV:NUMBER " & i & "," & i, True)      ' I.E. Limit the range to the 5th constant only  CAL:INT:DCV:HOSEI:NUMBER 5,5

                Threading.Thread.Sleep(10)

                ' New commmand, get reply, store it in array
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("CAL:INT:DCV:EEPROM:NEW?", q, False)                               ' Get the contents of the range specified above example, so just the 5th
                Cbdev1(q)

                ' Have got the reply, store it in array
                LabelCalRamByte6581.Text = q.ResponseAsString

                Me.Refresh()

                ' Write to text file
                IO.File.AppendAllText(RAMfilename6581, txtr1a.Text & Environment.NewLine)

                Threading.Thread.Sleep(10)

            Next

            Dev1TextResponse.Checked = False

            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)         ' Disable Service EXT protection mode.....not required for INT commands

            ' Finished
            CalramStatus6581.Text = "DONE!"

        Else

            ' GPIB Dev 1 has not been started
            CalramStatus6581.Text = "DEVICE 1 IS NOT STARTED"

        End If

    End Sub


    Private Sub ShowFilesCalRamR6581_Click(sender As Object, e As EventArgs) Handles ShowFilesCalRamR6581.Click
        'Process.Start("explorer.exe", String.Format("/n, /e, {0}", CSVfilepath.Text))
        Process.Start("explorer.exe", String.Format("/n, /e, {0}", strPath))
    End Sub


    Private Sub ButtonR6581abort_Click(sender As Object, e As EventArgs) Handles ButtonR6581abort.Click

        'dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)         ' Disable Service EXT protection mode.....not required for INT commands

        Abort6581 = True
        TextBoxCalRamFile6581.Text = ""
        Dev1TextResponse.Checked = False

    End Sub

End Class