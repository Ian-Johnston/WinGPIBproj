
' CalRam HP 3457A

'Imports System.Threading
'Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports IODevices
Imports PdfSharp.Pdf.Content.Objects

Partial Class Formtest

    Dim fs3457 As System.IO.FileStream
    'Dim fs23457 As System.IO.FileStream
    Dim CalRamPathfile3457 As System.IO.BinaryWriter
    'Dim CalRamPathfile23457 As System.IO.BinaryWriter
    Dim c3457 As Char

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


    Private Sub ShowFilesCalRam2_Click(sender As Object, e As EventArgs) Handles ShowFilesCalRam2.Click
        'Process.Start("explorer.exe", String.Format("/n, /e, {0}", CSVfilepath.Text))
        Process.Start("explorer.exe", String.Format("/n, /e, {0}", strPath))
    End Sub


    Private Sub ButtonCalramDump3457A_Click(sender As Object, e As EventArgs) Handles ButtonCalramDump3457A.Click

        ' 3457A
        respUSERTABonly = False

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
            c3457 = Chr(9)
            fs3457 = New System.IO.FileStream(RAMfilename3457A, IO.FileMode.OpenOrCreate)
            'fs3457 = New System.IO.FileStream(RAMfilename, IO.FileMode.Append)
            CalRamPathfile3457 = New System.IO.BinaryWriter(fs3457)
            CalRamPathfile3457.Seek(0, System.IO.SeekOrigin.Begin)

            TextBoxCalRamFile3457A.Text = RAMfilename3457A

            CalramStatus3457A.Text = "SETTING UP GPIB"
            System.Threading.Thread.Sleep(500)     ' 500mS delay
            Me.Refresh()

            dev1.SendAsync("TRIG 4", True)      ' TRIG HOLD
            CalramStatus3457A.Text = "TRIG 4"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            txtr1a.Text = ""                       ' Prepare reply as empty


            ' 10 dummy reads to set the interface up (some take a read or two to start getting valid data, buffer flush maybe)
            CalramStatus3457A.Text = "DUMMY READ - BUFFER FLUSH"
            For CalAddrtemp As Integer = 1 To 10 Step 1
                Dim r As IOQuery = Nothing
                dev1.QueryBlocking("PEEK " & CalAddrStart3457A, r, False)
                Debug.WriteLine("BLOCKING DetermineQuery: ")

                Cbdev1(r)
                System.Threading.Thread.Sleep(50)     ' 50mS delay
            Next

            System.Threading.Thread.Sleep(250)     ' 250mS delay


            ' Retrieve the data
            For CalAddr3457A As Integer = CalAddrStart3457A To CalAddrEnd3457A Step 2      ' step 2 so even addresses only

                If Abort3457A = True Then
                    Exit For
                End If

                CalramStatus3457A.Text = "READING........"

                ' Send MREAD command with address and wait for reply
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("PEEK " & CalAddr3457A, q, False)
                Debug.WriteLine("BLOCKING DetermineQuery: ")

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
                fs3457.WriteByte(Convert.ToByte(CalramStore3457Abyte1(Counter3457A), 16))
                fs3457.WriteByte(Convert.ToByte(CalramStore3457Abyte2(Counter3457A), 16))

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
            'fs3457.Close()
            CalRamPathfile3457.Flush()
            CalRamPathfile3457.Close()
            CalRamPathfile3457 = Nothing
            fs3457 = Nothing

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


    Private Sub Button3457Aabort_Click(sender As Object, e As EventArgs) Handles Button3457Aabort.Click

        Abort3457A = True
        TextBoxCalRamFile3457A.Text = ""
        respUSERTABonly = False

    End Sub

End Class