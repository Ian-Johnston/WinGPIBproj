
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

    '3458A
    Dim Abort6581 As Boolean = False
    Dim Stepsize6581 As Integer = 2
    Dim RamType6581 As String = ""
    Dim RamType26581 As String = ""
    Dim RAMfilename6581 As String
    Dim RAMfilename26581 As String
    Dim CalramAddress6581 As Integer
    Dim Calrambytefordisplay6581 As String
    Dim CalramStore6581(32768) As String
    Dim CalramStoreTemp16581 As String = ""
    Dim CalramStoreTemp26581 As String = ""
    Dim Counter6581 As Integer = 1
    Dim Counter65812 As Integer = 1
    Dim CalramValue6581 As String = ""
    Dim CalAddrStart6581 As Integer = 393216
    Dim CalAddrEnd6581 As Integer = 397311


    Private Sub ButtonCalramDumpR6581_Click(sender As Object, e As EventArgs) Handles ButtonCalramDumpR6581.Click

        'R6581

        ' run appropriate routine

        If AllConstantsReadR6581.Checked = True Then    ' All calibration constants
            TextBoxCalRamFile6581.Text = ""
            CalramextractR6581()
        End If


    End Sub


    Private Sub CalramextractR6581()

        ' R6581 all calibration constants read to file

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

            'dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 1", True)         ' Enable Service EXT protection mode.....not required for INT commands

            dev1.SendAsync("*RST", True)         ' NPLC 0
            CalramStatus6581.Text = "*RST"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
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

            RAMfilename6581 = strPath & "\" & "R6581_Calibration_Constants" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".txt"

            LabelCounter6581.Text = "0"
            Counter6581 = 0
            TextBoxCalRamFile6581.Text = RAMfilename6581

            If Not IO.File.Exists(RAMfilename6581) Then
                IO.File.Create(RAMfilename6581).Dispose()
            End If

            For i = 0 To 24

                CalramStatus6581.Text = "READING DCV Calibration Constants"
                LabelCounter6581.Text = i

                ' Blind command
                dev1.SendAsync("CAL:INT:DCV:HOSEI:NUMBER " & i & "," & i, True)

                ' New commmand, get reply, store it in array
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("CAL:INT:DCV:HOSEI?", q, False)
                Cbdev1(q)

                ' Have got the reply, store it in array
                CalramStore6581(Counter6581) = Val(txtr1a.Text)
                LabelCalRamByte6581.Text = Val(txtr1a.Text)

                ' Write to text file
                IO.File.AppendAllText(RAMfilename6581, CalramStore6581(Counter6581).ToString() & Environment.NewLine)

                Counter6581 = Counter6581 + 1

                Threading.Thread.Sleep(100)
            Next

            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            For i = 0 To 28

                CalramStatus6581.Text = "READING AC Calibration Constants"
                LabelCounter6581.Text = i

                ' Blind command
                dev1.SendAsync("CAL:INT:AC:HOSEI:NUMBER " & i & "," & i, True)

                ' New commmand, get reply, store it in array
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("CAL:INT:AC:HOSEI?", q, False)
                Cbdev1(q)

                ' Have got the reply, store it in array
                CalramStore6581(Counter6581) = Val(txtr1a.Text)
                LabelCalRamByte6581.Text = Val(txtr1a.Text)

                ' Write to text file
                IO.File.AppendAllText(RAMfilename6581, CalramStore6581(Counter6581).ToString() & Environment.NewLine)

                Counter6581 = Counter6581 + 1

                Threading.Thread.Sleep(100)
            Next

            'dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)         ' Disable Service EXT protection mode.....not required for INT commands

            ' Finished
            CalramStatus.Text = "DONE!"

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

    End Sub

End Class