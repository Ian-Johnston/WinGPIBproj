
' CalRam MREAD (3458A)

Imports IODevices
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.IO

Partial Class Formtest

    Dim fs As System.IO.FileStream
    Dim fs2 As System.IO.FileStream
    Dim CalRamPathfile As System.IO.BinaryWriter
    Dim CalRamPathfile2 As System.IO.BinaryWriter
    Dim c As Char

    '3458A
    Dim Abort3458A As Boolean = False
    Dim Stepsize As Integer = 2
    Dim RamType As String = ""
    Dim RamType2 As String = ""
    Dim RAMfilename As String
    Dim RAMfilename2 As String
    Dim CalramAddress As Integer
    Dim Calrambytefordisplay As String
    Dim CalramStore(32768) As String
    Dim CalramStoreTemp1 As String = ""
    Dim CalramStoreTemp2 As String = ""
    Dim Counter As Integer = 1
    Dim Counter2 As Integer = 1
    Dim CalramValue As String = ""
    Dim CalAddrStart As Integer = 393216
    Dim CalAddrEnd As Integer = 397311

    '3457A
    Dim Abort3457A As Boolean = False
    Dim RAMfilename3457A As String
    Dim CalramAddress3457A As Integer
    'Dim CalramAddress3457AHex As String
    Dim CalramStore3457A(32768) As String
    Dim CalramStore3457Abyte1(32768) As String
    Dim CalramStore3457Abyte2(32768) As String
    Dim Counter3457A As Integer = 1
    Dim CalramValue3457A As String = ""
    Dim CalAddrStart3457A As Integer = 64
    Dim CalAddrEnd3457A As Integer = 511


    Private Sub ButtonCalramDump_Click(sender As Object, e As EventArgs) Handles ButtonCalramDump3458A.Click

        '3458A

        ' run appropriate routine

        If AddressRangeC.Checked = True Then    ' Calram
            ' 0x60000...0x60fff, so issuing 2048 GPIB commands
            CalAddrStart = 393216
            CalAddrEnd = 397311
            Stepsize = 2
            RamType = "3458A_Cal_ram_"
            TextBoxCalRamFile.Text = ""
            TextBoxCalRamFile2.Text = ""
            Calramextract3458A()
        End If

        If AddressRangeD.Checked = True Then    ' Settings ram 1 & 2
            ' 0x120000...0x12ffff (1179648...1245182 decimal)
            CalAddrStart = 1179648
            CalAddrEnd = 1245183        '1245183 1212415
            Stepsize = 2
            RamType = "3458A_Settings_ram_L_U121_"
            RamType2 = "3458A_Settings_ram_U_U122_"
            TextBoxCalRamFile.Text = ""
            TextBoxCalRamFile2.Text = ""
            Settingsramextract3458A()
        End If

    End Sub


    Private Sub Calramextract3458A()

        ' 3458A

        Abort3458A = False

        CalramStatus.Text = "CHECKING SETUP"

        Me.Refresh()

        If ButtonDev1Run.Enabled = True Then      ' Device 1 is started

            'RAMfilename = CSVfilepath.Text & "\" & RamType & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            RAMfilename = strPath & "\" & RamType & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            c = Chr(9)
            fs = New System.IO.FileStream(RAMfilename, IO.FileMode.OpenOrCreate)
            'fs = New System.IO.FileStream(RAMfilename, IO.FileMode.Append)
            CalRamPathfile = New System.IO.BinaryWriter(fs)
            CalRamPathfile.Seek(0, System.IO.SeekOrigin.Begin)

            LabelCounter.Text = "0"
            Counter = 0
            Counter2 = 0

            TextBoxCalRamFile.Text = RAMfilename

            CalramStatus.Text = "SETTING UP GPIB"
            System.Threading.Thread.Sleep(500)     ' 500mS delay
            Me.Refresh()

            dev1.SendAsync("NPLC 0", True)         ' NPLC 0
            CalramStatus.Text = "NPLC 0"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            dev1.SendAsync("TRIG HOLD", True)      ' TRIG HOLD
            CalramStatus.Text = "TRIG HOLD"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            dev1.SendAsync("QFORMAT NUM", True)    ' QFORMAT NUM - Query responses will be in number format only, no headers. The 3458A cannot return HEX formatted data.
            CalramStatus.Text = "QFORMAT NUM"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            txtr1a.Text = ""                       ' Prepare reply as empty

            ' Retrieve the data

            For CalAddr As Integer = CalAddrStart To CalAddrEnd Step Stepsize      ' Stepsize 2 enables reading of a 16bit word

                If Abort3458A = True Then
                    Exit For
                End If

                CalramStatus.Text = "READING 2048 BYTES (1024 16bit)"

                ' Send MREAD command with address and wait for reply
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("MREAD " & CalAddr, q, False)
                Cbdev1(q)   ' Process reply which stores value in txtr1a.Text (see Formtest.vb)

                ' Got reply, store it in array
                CalramStore(Counter) = Hex(Val(txtr1a.Text))

                If (Len(CalramStore(Counter)) > 4) Then     ' originally a negative number i.e. FFFFE3B9 so need to strip FFFF of beginning
                    CalramStore(Counter) = CalramStore(Counter).Remove(0, 4)  ' remove first 4 characters if 5 or more bytes long
                End If

                If (Len(CalramStore(Counter)) = 3) Then     ' FB9 should be 0FB9 so need to add a 0 to beginning
                    CalramStore(Counter) = "0" & CalramStore(Counter)
                End If

                If (Len(CalramStore(Counter)) = 2) Then     ' B9 should be 00B9 so need to add a 00 to beginning
                    CalramStore(Counter) = "00" & CalramStore(Counter)
                End If

                If (Len(CalramStore(Counter)) = 1) Then     ' B should be 000B so need to add a 000 to beginning
                    CalramStore(Counter) = "000" & CalramStore(Counter)
                End If

                ' Calram - Strip off the last two characters
                If (Len(CalramStore(Counter)) = 4) Then     ' E3B9 so need to strip of 2nd hex number
                    CalramStore(Counter) = CalramStore(Counter).Remove(2, 2)  ' remove last two characters, i.e. strip low byte
                End If

                ' Write to binary file
                fs.WriteByte(Convert.ToByte(CalramStore(Counter), 16))

                ' Update display
                LabelCalRamAddress.Text = CalAddr
                LabelCalRamAddressHex.Text = String.Join(",", LabelCalRamAddress.Text.Split(","c).        ' Hex conversion
                              Select(Function(x) _
                              Convert.ToInt32(x).ToString("X")))

                LabelCalRamByte.Text = Calrambytefordisplay

                CalramStatus.Text = CalAddr & "=" & Int(Val(txtr1a.Text))     ' display
                Counter = Counter + 1   ' prepare for next loop
                Counter2 = Counter2 + 2   ' prepare for next loop

                LabelCounter.Text = Counter

            Next

            ' Close file
            fs.Close()

            ' QFORMAT NORM, TRIG AUTO - set back to 3458A defaults
            dev1.SendAsync("QFORMAT NORM", True)
            dev1.SendAsync("TRIG AUTO", True)

            ' Abort display update
            If Abort3458A = True Then
                Abort3458A = False
                CalramStatus.Text = "ABORTED!"
                TextBoxCalRamFile.Text = ""
                TextBoxCalRamFile2.Text = ""
                fs.Close()
            Else
                ' Finished
                CalramStatus.Text = "DONE!"
            End If

        Else

            ' GPIB Dev 1 has not been started
            CalramStatus.Text = "DEVICE 1 IS NOT STARTED"

        End If

    End Sub


    Private Sub Settingsramextract3458A()

        ' 3458A

        Abort3458A = False

        CalramStatus.Text = "CHECKING SETUP"

        Me.Refresh()

        If ButtonDev1Run.Enabled = True Then      ' Device 1 is started

            ' RAM0L Lower (U121)
            RAMfilename = strPath & "\" & RamType & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            c = Chr(9)
            fs = New System.IO.FileStream(RAMfilename, IO.FileMode.OpenOrCreate)
            'fs = New System.IO.FileStream(RAMfilename, IO.FileMode.Append)
            CalRamPathfile = New System.IO.BinaryWriter(fs)
            CalRamPathfile.Seek(0, System.IO.SeekOrigin.Begin)

            ' RAM0H Upper (U122)
            RAMfilename2 = strPath & "\" & RamType2 & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            c = Chr(9)
            fs2 = New System.IO.FileStream(RAMfilename2, IO.FileMode.OpenOrCreate)
            'fs = New System.IO.FileStream(RAMfilename, IO.FileMode.Append)
            CalRamPathfile2 = New System.IO.BinaryWriter(fs2)
            CalRamPathfile2.Seek(0, System.IO.SeekOrigin.Begin)


            LabelCounter.Text = "0"
            Counter = 0
            Counter2 = 0

            TextBoxCalRamFile.Text = RAMfilename    ' L
            TextBoxCalRamFile2.Text = RAMfilename2  ' U

            CalramStatus.Text = "SETTING UP GPIB"
            System.Threading.Thread.Sleep(500)     ' 500mS delay
            Me.Refresh()

            dev1.SendAsync("NPLC 0", True)         ' NPLC 0
            CalramStatus.Text = "NPLC 0"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            dev1.SendAsync("TRIG HOLD", True)      ' TRIG HOLD
            CalramStatus.Text = "TRIG HOLD"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            dev1.SendAsync("QFORMAT NUM", True)    ' QFORMAT NUM - Query responses will be in number format only, no headers. The 3458A cannot return HEX formatted data.
            CalramStatus.Text = "QFORMAT NUM"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            txtr1a.Text = ""                       ' Prepare reply as empty


            ' Retrieve the data
            For CalAddr As Integer = CalAddrStart To CalAddrEnd Step Stepsize      ' Stepsize 2 enables reading of a 16bit word

                If Abort3458A = True Then
                    Exit For
                End If

                CalramStatus.Text = "READING 2 LOTS 32768 BYTES (2 LOTS 16384 16bit)"

                ' Send MREAD command with address and wait for reply
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("MREAD " & CalAddr, q, False)
                Cbdev1(q)   ' Process reply which stores value in txtr1a.Text (see Formtest.vb)

                ' Got reply, store it in array
                CalramStore(Counter) = Hex(Val(txtr1a.Text))

                If (Len(CalramStore(Counter)) > 4) Then     ' originally a negative number i.e. FFFFE3B9 so need to strip FFFF of beginning
                    CalramStore(Counter) = CalramStore(Counter).Remove(0, 4)  ' remove first 4 characters if 5 or more bytes long
                End If

                If (Len(CalramStore(Counter)) = 3) Then     ' FB9 should be 0FB9 so need to add a 0 to beginning
                    CalramStore(Counter) = "0" & CalramStore(Counter)
                End If

                If (Len(CalramStore(Counter)) = 2) Then     ' B9 should be 00B9 so need to add a 00 to beginning
                    CalramStore(Counter) = "00" & CalramStore(Counter)
                End If

                If (Len(CalramStore(Counter)) = 1) Then     ' B should be 000B so need to add a 000 to beginning
                    CalramStore(Counter) = "000" & CalramStore(Counter)
                End If


                ' Settings 1/2 - Split the two bytes and save them....so xABCD becomes two bytes AB (high byte) and CD (low byte)
                If (Len(CalramStore(Counter)) = 4) Then     ' E3B9 so need to split into two bytes
                    CalramStoreTemp1 = CalramStore(Counter).Remove(2, 2)  ' remove last two characters, i.e. strip low byte
                    CalramStoreTemp2 = CalramStore(Counter).Remove(0, 2)  ' remove first two characters, i.e. strip high byte
                End If
                ' Write low byte to binary file
                fs.WriteByte(Convert.ToByte(CalramStoreTemp2, 16))
                ' Write high byte to binary file
                fs2.WriteByte(Convert.ToByte(CalramStoreTemp1, 16))


                ' Update display
                LabelCalRamAddress.Text = CalAddr
                LabelCalRamAddressHex.Text = String.Join(",", LabelCalRamAddress.Text.Split(","c).        ' Hex conversion
                              Select(Function(x) _
                              Convert.ToInt32(x).ToString("X")))

                LabelCalRamByte.Text = Calrambytefordisplay

                CalramStatus.Text = CalAddr & "=" & Int(Val(txtr1a.Text))     ' display
                Counter = Counter + 1   ' prepare for next loop
                Counter2 = Counter2 + 2   ' prepare for next loop

                LabelCounter.Text = Counter2

            Next

            ' Close both file
            fs.Close()
            fs2.Close()

            ' QFORMAT NORM, TRIG AUTO - set back to 3458A defaults
            dev1.SendAsync("QFORMAT NORM", True)
            dev1.SendAsync("TRIG AUTO", True)

            ' Abort display update
            If Abort3458A = True Then
                Abort3458A = False
                CalramStatus.Text = "ABORTED!"
                TextBoxCalRamFile.Text = ""
                TextBoxCalRamFile2.Text = ""
                fs.Close()
                fs2.Close()
            Else
                ' Finished
                CalramStatus.Text = "DONE!"
            End If

        Else

            ' GPIB Dev 1 has not been started
            CalramStatus.Text = "DEVICE 1 IS NOT STARTED"

        End If

    End Sub


    Private Sub ShowFilesCalRam_Click(sender As Object, e As EventArgs) Handles ShowFilesCalRam.Click
        'Process.Start("explorer.exe", String.Format("/n, /e, {0}", CSVfilepath.Text))
        Process.Start("explorer.exe", String.Format("/n, /e, {0}", strPath))
    End Sub


    Private Sub ButtonCalramDump3457A_Click(sender As Object, e As EventArgs) Handles ButtonCalramDump3457A.Click

        ' 3457A

        If AddressRangeA.Checked = True Then
            CalAddrStart3457A = 64
            CalAddrEnd3457A = 511
        End If
        If AddressRangeB.Checked = True Then
            CalAddrStart3457A = 20480
            CalAddrEnd3457A = 22527
        End If
        If AddressRangeF.Checked = True Then
            CalAddrStart3457A = Val(TextBox3457AFrom.Text)
            CalAddrEnd3457A = Val(TextBox3457ATo.Text)

            If (CalAddrStart3457A < 0) Then
                CalAddrStart3457A = 0
            End If

            If (CalAddrEnd3457A < 0) Then
                CalAddrEnd3457A = 0
            End If

            If (CalAddrStart3457A > 32767) Then
                CalAddrStart3457A = 32767
            End If

            If (CalAddrEnd3457A > 32767) Then
                CalAddrEnd3457A = 32767
            End If

            If (CalAddrEnd3457A < CalAddrStart3457A) Then
                CalAddrStart3457A = 0
                CalAddrEnd3457A = 32767
            End If
        End If

        LabelCounter3457A.Text = "0"
        Counter3457A = 0

        Abort3457A = False

        CalramStatus3457A.Text = "CHECKING SETUP"

        Me.Refresh()

        If ButtonDev1Run.Enabled = True Then      ' Device 1 is started

            'RAMfilename3457A = CSVfilepath.Text & "\" & "3457ACalram_" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            RAMfilename3457A = strPath & "\" & "3457ACalram_" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            c = Chr(9)
            fs = New System.IO.FileStream(RAMfilename3457A, IO.FileMode.OpenOrCreate)
            'fs = New System.IO.FileStream(RAMfilename, IO.FileMode.Append)
            CalRamPathfile = New System.IO.BinaryWriter(fs)
            CalRamPathfile.Seek(0, System.IO.SeekOrigin.Begin)

            TextBoxCalRamFile3457A.Text = RAMfilename3457A

            CalramStatus3457A.Text = "SETTING UP GPIB"
            System.Threading.Thread.Sleep(500)     ' 500mS delay
            Me.Refresh()

            dev1.SendAsync("TRIG 4", True)      ' TRIG HOLD
            CalramStatus3457A.Text = "TRIG 4"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            txtr1a.Text = ""                       ' Prepare reply as empty

            ' Retrieve the data
            For CalAddr3457A As Integer = CalAddrStart3457A To CalAddrEnd3457A Step 2      ' step 2 so even addresses only

                If Abort3457A = True Then
                    Exit For
                End If

                CalramStatus3457A.Text = "READING........"

                ' Send MREAD command with address and wait for reply
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("PEEK " & CalAddr3457A, q, False)
                Cbdev1(q)   ' Process reply which stores value in txtr1a.Text (see Formtest.vb)

                ' Got reply, store it in array
                CalramStore3457A(Counter3457A) = Hex(Val(txtr1a.Text))

                'Label127.Text = CalramStore3457A(Counter3457A)
                'Me.Refresh()

                If (Len(CalramStore3457A(Counter3457A)) > 4) Then     ' originally a negative number i.e. FFFFE3B9 so need to strip FFFF of beginning
                    CalramStore3457A(Counter3457A) = CalramStore3457A(Counter3457A).Remove(0, 4)  ' remove first 4 characters if 5 or more bytes long
                End If

                If (Len(CalramStore3457A(Counter3457A)) = 3) Then     ' FB9 should be 0FB9 so need to add a 0 to beginning
                    CalramStore3457A(Counter3457A) = "0" & CalramStore3457A(Counter3457A)
                End If

                If (Len(CalramStore3457A(Counter3457A)) = 1) Then     ' B should be 000B so need to add a 000 to beginning
                    CalramStore3457A(Counter3457A) = "000" & CalramStore3457A(Counter3457A)
                End If

                If (Len(CalramStore3457A(Counter3457A)) = 2) Then     ' B9 should be 00B9 so need to add a 00 to beginning
                    CalramStore3457A(Counter3457A) = "00" & CalramStore3457A(Counter3457A)
                End If

                If (Len(CalramStore3457A(Counter3457A)) = 3) Then     ' B should be 000B so need to add a 00 to beginning
                    CalramStore3457A(Counter3457A) = "0" & CalramStore3457A(Counter3457A)
                End If

                ' Now strip into two bytes
                If (Len(CalramStore3457A(Counter3457A)) = 4) Then     ' E3B9 so need to strip into two bytes
                    CalramStore3457Abyte1(Counter3457A) = CalramStore3457A(Counter3457A).Remove(2, 2)  ' remove last two characters for byte 1
                    CalramStore3457Abyte2(Counter3457A) = CalramStore3457A(Counter3457A).Remove(0, 2)  ' remove first two characters for byte 2....so xABCD becomes two bytes AB and CD
                End If

                'Label129.Text = CalramStore3457Abyte1(Counter3457A)
                'Label131.Text = CalramStore3457Abyte2(Counter3457A)

                ' Write to text box
                'Me.ListCalRam.Items.Insert(0, CalramStore3457A(Counter3457A))

                ' Write to binary file
                fs.WriteByte(Convert.ToByte(CalramStore3457Abyte1(Counter3457A), 16))
                fs.WriteByte(Convert.ToByte(CalramStore3457Abyte2(Counter3457A), 16))

                LabelCounter3457A.Text = Counter3457A
                LabelCalRamAddress3457A.Text = CalAddr3457A
                LabelCalRamAddress3457AHex.Text = String.Join(",", LabelCalRamAddress3457A.Text.Split(","c).        ' Hex conversion
                              Select(Function(x) _
                              Convert.ToInt32(x).ToString("X")))

                LabelCalRamByte3457A.Text = CalramStore3457A(Counter3457A)
                CalramStatus3457A.Text = CalAddr3457A & "=" & Int(Val(txtr1a.Text))     ' display
                Counter3457A = Counter3457A + 1   ' prepare for next loop

            Next

            ' Close file
            fs.Close()

            ' Abort display update
            If Abort3457A = True Then
                Abort3457A = False
                CalramStatus3457A.Text = "ABORTED!"
                TextBoxCalRamFile3457A.Text = ""
            Else
                ' Finished
                LabelCalRamAddress3457A.Text = CalAddrEnd3457A
                CalramStatus3457A.Text = "DONE!"
            End If

        Else

            ' GPIB Dev 1 has not been started
            CalramStatus3457A.Text = "DEVICE 1 IS NOT STARTED"

        End If
    End Sub
    Private Sub Button3458Aabort_Click(sender As Object, e As EventArgs) Handles Button3458Aabort.Click

        Abort3458A = True
        TextBoxCalRamFile.Text = ""
        TextBoxCalRamFile2.Text = ""

    End Sub
    Private Sub Button3457Aabort_Click(sender As Object, e As EventArgs) Handles Button3457Aabort.Click

        Abort3457A = True
        TextBoxCalRamFile3457A.Text = ""

    End Sub

End Class