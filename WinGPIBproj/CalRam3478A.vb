
' CalRam HP 3478A

Imports System.Drawing.Text
Imports System.Globalization
Imports System.IO.Ports
Imports System.Linq.Expressions
Imports System.Net
Imports System.Resources
Imports System.Runtime.ConstrainedExecution
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Threading
Imports IODevices
Imports Ivi.Visa
Imports Microsoft.Office.Interop.Word
Imports MigraDoc.DocumentObjectModel.Shapes
Imports Newtonsoft.Json.Serialization

Partial Class Formtest
    '3478A
    Dim Abort3478A As Boolean = False
    Dim RAMfilename3478A As String
    Dim CalramAddress3478A As Integer
    Dim CalramStore3478A(255) As String
    Dim Counter3478A As Integer = 1
    Dim CalramValue3478A As String = ""
    Dim CalAddrStart3478A As Integer = 0
    Dim CalAddrEnd3478A As Integer = 255
    Dim OpenFileDialog1 As New OpenFileDialog()
    Dim DataToLoad(255) As Byte
    Dim DataReadFromRam(255) As Byte
    Dim ReadCommand As String
    Dim ReadCommandArray(2) As Byte
    Dim ReadResult As Byte
    Dim ReadAddrByte As Byte
    Dim StartAddrInconsistance(1) As Integer
    Dim EndAddrInconsistance(1) As Integer
    Dim q As IOQuery = Nothing
    Dim Offset As String
    Dim GainValue As String
    Dim IterationRawOffset As String
    Dim IterationRawGain As String
    Dim IterationChecksum As String
    Dim EditedCalibrationData(18) As String
    Dim NewOffset(18) As String
    Dim NewGain(18) As String
    Dim NewCheckSum3478A(18) As String
    Dim CalibrationRangeStore3478A(18) As String
    Dim CalibrationRangeBackup3478A(18) As String
    Dim CalibrationOffsetBackup(18) As String
    Dim CalibrationGainBackup(18) As String
    Dim CalibrationValidBackup(18) As String
    Dim SerialNumberBackup As String
    Dim DeviceID As String
    Dim CalibrationSwitchDisabled As Boolean
    Dim WriteSuccess3478A As Boolean
    Dim FileContainsSerNr As Boolean
    Dim Caller3478A As String
    Dim EditOccurred As Boolean
    Dim QuickAndDirty As Boolean
    Dim EditCount As Integer
    Dim SuppressMsgs As Boolean
    Private Sub CheckBoxSupressMsg_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSuppressMsg.CheckedChanged
        'Suppress not error related messages at user request
        SuppressMsgs = Not SuppressMsgs
    End Sub

    'This subroutine calls the process that will download the calibration data from a connected HP3478A's RAM to a binary file
    Private Sub ButtonCalramDump3478A_Click(sender As Object, e As EventArgs) Handles ButtonCalramDump3478A.Click

        'Check if device 1 is connected
        If Not btnq1b.Enabled = True Then
            MessageBox.Show("The device has not been connected")
            Exit Sub
        End If

        'Initialize abort flag
        Abort3478A = False
        Button3478Aabort.Enabled = True

        'Initialize QUICK-AND-DIRTS flag
        QuickAndDirty = False

        'Initialize calibration data edit and rollback buffer
        For I = 0 To 18
            CalibrationRangeBackup3478A(I) = Nothing
            EditedCalibrationData(I) = Nothing
        Next

        'Buffer the calling subrouting
        Caller3478A = "Read"

        'Manage HP3478A serial number options
        CheckSerialNumber()
        If Not q.status = 0 Then                                                      'check if communication with the HP3478A is working
            Exit Sub
        ElseIf Abort3478A Then                                                        'abort if serial number is invalid
            MessageBox.Show("The entered serial number is invalid (must be 4 numeric 1 alpha and 5 numberic - NNNNANNNNN")
            Exit Sub
        ElseIf TextBoxSerialNumber3478A.Focus = True And TextBoxSerialNumber3478A.Text = Nothing Or Abort3478A = True Then
            Exit Sub                                                                  'user decided to enter 3478A unit's serial number
        End If

        'Adjust program flow options
        ButtonCalramDump3478A.Enabled = False
        ButtonCalramLoad3478A.Enabled = False
        ButtonRestartTab3478A.Enabled = True
        TextBoxSerialNumber3478A.Enabled = False
        TextBox3478AFrom.Enabled = False
        TextBox3478ATo.Enabled = False
        TextBoxCalramFile3478A.Enabled = False
        TextBoxCalRamLoadFile3478A.Enabled = False
        ComboBoxCalRange3478A.Enabled = False

        'Perform the dump of the connected unit's calibration ram data
        DumpCalram3478A()

        'Adjust program flow options
        If Abort3478A Then
            ClearTab3478A()
            ButtonCalramDump3478A.Enabled = True
            Button3478Aabort.Enabled = False
            ButtonRestartTab3478A.Enabled = False
        ElseIf Not EditOccurred Then
            ButtonCalramLoad3478A.Enabled = False
        Else
            'Allow the option to immediately upload the entered serial number into the HP3478A calibration memory range 19 (otherwise unused calibration range)
            LabelCalRange19.Text = "HP serial number"
            ButtonCalramLoad3478A.Enabled = True
            ButtonCalramLoad3478A.Text = "3478A Upload"
        End If

    End Sub
    Private Sub DumpCalram3478A()

        'Initialize fields
        LabelCounter3478A.Text = "0"
        Counter3478A = 0
        Abort3478A = False

        'Always dump the entire memory range regardless of user address range selection
        CalAddrStart3478A = 0
        CalAddrEnd3478A = 255

        'Display the current process to the user
        CalramStatus3478A.Text = "CHECKING SETUP"
        Me.Refresh()

        'Start the download process
        If btnq1b.Enabled = True Then      ' Device 1 is connected

            'Create and open a binary file where the calibration ram data read from the unit will be saved as backup from which to restore the unit's calibration
            RAMfilename3478A = strPath & "\" & "3478ACalram_" & dev1.devname & "_" & "SerNr_" & TextBoxSerialNumber3478A.Text & "_" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            fs = New System.IO.FileStream(RAMfilename3478A, IO.FileMode.OpenOrCreate)
            CalRamPathfile = New System.IO.BinaryWriter(fs)
            CalRamPathfile.Seek(0, System.IO.SeekOrigin.Begin)

            'Display the path and file name to the user
            TextBoxCalramFile3478A.Text = RAMfilename3478A

            'Block other processing paths while a download is running
            ButtonCalramLoad3478A.Enabled = False
            ButtonRestartTab3478A.Enabled = False

            'Display the current processing status to the user
            CalramStatus3478A.Text = "SETTING UP GPIB"
            System.Threading.Thread.Sleep(500)     ' 500mS delay
            Me.Refresh()

            'Initialize the variable which recieves the 3478A unit's response to a data request
            txtr1a.Text = ""

            ' Retrieve the data from the unit's calibration ram
            For CalAddr3478A As Integer = CalAddrStart3478A To CalAddrEnd3478A Step 1   'loop through the calibration ram addresses

                'Abort the download at user's request
                If Abort3478A = True Then
                    MessageBox.Show("The generated binary download file is incomplete and should not be used for uploading into an HP3478A calibration memory!")

                    'Re-allow other processing paths that were bloacked while the download was running
                    ButtonCalramLoad3478A.Enabled = True
                    ButtonRestartTab3478A.Enabled = True
                    Exit For
                End If

                'Display the current processing status
                CalramStatus3478A.Text = "READING........"

                'Get data byte from the currently processed address of the HP3478A calibration data memory
                If Not GetCalRamByte(CalAddr3478A) = 0 Then
                    Exit Sub
                End If

                'Store the received byte to the data array which buffers the calibration ram data for later output
                CalramStore3478A(Counter3478A) = CalramValue3478A

                'Special handling of the last, normally unused, calibration range data in the event it is used to store the unit's serial number
                If Counter3478A > 239 And Counter3478A < 248 Then                       'address range that must be restructured if serial number is found
                    CalramStore3478A(Counter3478A + 1) = CalramValue3478A               'prepare to replace address 240 (6th position in the offset of range 19) with a space
                    If CalramStore3478A(239) = "@" Then                                 'if the 5th position of the offset address range is not "empty", then a serial number has been detected
                        CalramStore3478A(240) = "@"                                     'no serial number detected, so continue to process the address 240 normally
                    Else
                        CalramStore3478A(240) = " "                                     'replace address 240 with a space to enable the display of a properly structured serial number
                    End If
                End If

                'Write the recieved value to binary output file
                ReadResult = q.ResponseAsByteArray(0)
                fs.WriteByte(ReadResult)

                'Display the processed byte of the unit's memory data
                LabelCounter3478A.Text = Counter3478A
                LabelCalramAddress3478A.Text = CalAddr3478A
                LabelCalRamAddress3478AHex.Text = String.Join(",", LabelCalramAddress3478A.Text.Split(","c).        ' Hex conversion
                              Select(Function(x) _
                              Convert.ToInt32(x).ToString("X")))
                LabelCalRamByte3478A.Text = CalramStore3478A(Counter3478A)
                CalramStatus3478A.Text = CalAddr3478A & "=" & Int(Val(txtr1a.Text))

                Counter3478A = Counter3478A + 1                                         'increment iteration counter 
            Next

            'Enable Calibration Range selection for optional editing
            ComboBoxCalRange3478A.Enabled = True

            'Close the binary output file
            fs.Close()


            'Calculate checksums to verify calibration data is consistent
            VerifyCalibrationFileData()

            'Buffer the calibration range data
            If CalibrationRangeBackup3478A(0) Is Nothing Then
                For i = 0 To 18
                    CalibrationRangeBackup3478A(i) = CalibrationRangeStore3478A(i)
                Next
            End If

            'Abort display update at user's request
            If Abort3478A Then
                CalramStatus3478A.Text = "ABORTED!"
                TextBoxCalramFile3478A.Text = ""

                'Rename the binary output file to indicate it is incomplete since the download process was aborted during processing
                TextBoxCalramFile3478A.Text = Mid(RAMfilename3478A, 1, RAMfilename3478A.Length - 4) & "_Aborted.bin"
                System.IO.File.Move(RAMfilename3478A, TextBoxCalramFile3478A.Text)
            Else
                'Finished
                LabelCalramAddress3478A.Text = CalAddrEnd3478A
                CalramStatus3478A.Text = "DONE!"

            End If

            'Adjust program flow options
            ButtonRestartTab3478A.Enabled = True
            Button3478Aabort.Enabled = False

        Else

            'Dev 1 has not been connected
            CalramStatus3478A.Text = "DEVICE 1 IS NOT CONNECTED"

        End If

    End Sub

    'This subroutine reads calibration data from a binary file and uploads it into the 3478A's calibration RAM
    Private Sub ButtonCalramLoad3478A_Click(sender As Object, e As EventArgs) Handles ButtonCalramLoad3478A.Click

        Dim ReadBackValue As String
        Dim LoadCommand As String
        Dim CalAddrHex As Byte
        Dim LoadData(4) As Byte
        Dim ReadBackData(2) As Byte
        Dim CalAddrByte As Byte
        Dim AsciiValue As Integer
        Dim SerialNumber As String
        Dim SerNrChangeOK As Boolean

        'Initialize abort flag
        Abort3478A = False
        Button3478Aabort.Enabled = True

        If Caller3478A = "Read" And ComboBoxCalRange3478A.Text = Nothing And Not EditOccurred Then

            'Prevent unnecessary upload of identical data previously downloaded
            ClearTab3478A()
        End If

        'Buffer the calling subrouting
        Caller3478A = "Write"

        'Update buffer for calibration range edits
        For i = 0 To 18
            If EditedCalibrationData(i) <> Nothing And CalibrationRangeBackup3478A(i) <> Nothing Then
                CalibrationRangeBackup3478A(i) = EditedCalibrationData(i)
                EditOccurred = True
            End If
        Next

        'Prevent unnecessary upload of calibration ram data if it would not change the HP3478A RAM memory content
        For i = 0 To 18
            If Not CalibrationRangeBackup3478A(i) = CalibrationRangeStore3478A(i) Then
                EditOccurred = True
            End If
        Next
        If Not EditOccurred And Not CalibrationRangeStore3478A(0) = Nothing Then
            MessageBox.Show("The data to load is identical to the unit's current calibration ram data. No need to continue upload!")
            ButtonCalramLoad3478A.Enabled = False
            Exit Sub
        End If

        'Check if the serlial number is being intentionally deleted (using the editing option)
        If Not NewOffset(18) = "0" And Not NewGain(18) = "1" Then
            CheckSerialNumber()

            'Check if communication with the HP3478A is working
            If Not q.status = 0 Then
                'Error message will have already been displayed during the CheckSerialNumber() process
                Exit Sub
            End If

            'Check if the serial number is formally valid
            If Abort3478A Then
                MessageBox.Show("The serial number is invalid (must be 4 numberic 1 alpha and 5 numeric - NNNNANNNN)")
                LabelCalRange19.Text = "Not used"
                Exit Sub
            End If
        End If


        'User decides whether to enter 3478A unit's serial number or to continue without using serial number
        If TextBoxSerialNumber3478A.Focus = True And TextBoxSerialNumber3478A.Text = Nothing Or Abort3478A = True Then

            'User has decided to enter the HP3478A serial number
            LabelCalRange19.Text = "HP serial number"
            Exit Sub
        End If

        'Adjust program flow options
        ButtonCalramDump3478A.Enabled = False
        ButtonCalramLoad3478A.Enabled = False
        TextBoxSerialNumber3478A.Enabled = False
        TextBox3478AFrom.Enabled = False
        TextBox3478ATo.Enabled = False

        'Make sure memory range is between 0 and 255 and set the user selected address range
        If AddressRange3478A.Checked = True Then
            CalAddrStart3478A = Val(TextBox3478AFrom.Text)
            CalAddrEnd3478A = Val(TextBox3478ATo.Text)

            'Warn the user if address range selection does not correspond to a valid calibration range
            If Not CalAddrStart3478A = 0 Or Not CalAddrEnd3478A = 255 Then
                If ((CalAddrStart3478A + 12) Mod 13 <> 0 And CalAddrStart3478A <> 0) Or (CalAddrEnd3478A Mod 13 <> 0 And CalAddrEnd3478A < 248) Then
                    Dim result As DialogResult = MessageBox.Show("selected address range is inconsistent with calibration ranges. Allow correction?", "Confirm Action", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                    If result = DialogResult.Yes Then

                        'Adjust selected calibration address range to nearest consistent calibration range
                        CalAddrStart3478A = CalAddrStart3478A - CalAddrStart3478A Mod 13 + 1
                        CalAddrEnd3478A = CalAddrEnd3478A - CalAddrEnd3478A Mod 13
                    Else

                        'Store the inconsistance values for later use
                        StartAddrInconsistance(0) = (CalAddrStart3478A - 1) Mod 13
                        StartAddrInconsistance(1) = CalAddrStart3478A
                        If CalAddrEnd3478A < 248 Then                                    'ram addresses beyond 248 are irrelevant for the HP3478A calibration
                            EndAddrInconsistance(0) = CalAddrEnd3478A Mod 13
                            EndAddrInconsistance(1) = CalAddrEnd3478A
                        End If
                    End If
                End If
            End If

            'Insure selected address range is within ram address range
            If (CalAddrStart3478A < 0) Then
                CalAddrStart3478A = 0
            End If

            If (CalAddrEnd3478A < 0) Then
                CalAddrEnd3478A = 0
            End If

            If (CalAddrStart3478A > 255) Then
                CalAddrStart3478A = 255
            End If

            If (CalAddrEnd3478A > 255) Then
                CalAddrEnd3478A = 255
            End If

            If (CalAddrEnd3478A < CalAddrStart3478A) Then
                CalAddrStart3478A = 0
                CalAddrEnd3478A = 255
            End If
        End If

        'Initialize memory address variables
        AddressRange3478A.Checked = True
        CalAddrStart3478A = Val(TextBox3478AFrom.Text)
        CalAddrEnd3478A = Val(TextBox3478ATo.Text)
        LabelCounter3478A.Text = "0"
        Counter3478A = 0
        Abort3478A = False
        CalramStatus3478A.Text = "CHECKING SETUP"
        Me.Refresh()

        'Initialize calibration ranges of previously selected upload ram address range
        For i = 0 To 18
            SetCalibrationRangeColor(i, Color.Black, 0, 0, True)
        Next

        If btnq1b.Enabled = True Then                                             'device 1 is connected

            'Check if a backup file to be restored has been selected
            If CalramStore3478A(1) = Nothing And EditOccurred = False Or TextBoxSerialNumber3478A.BackColor = Color.LightYellow Then  'a backup file has not yet been loaded

                ' User gets a dialog to choose the appropriate file containing calibration data for uploading to device
                If RAMfilename3478A = Nothing Then
                    OpenFileDialog1.ShowDialog()
                    Try
                        'Open the selected backup file
                        RAMfilename3478A = OpenFileDialog1.FileNames(0)                  'buffer the selected file name and path
                    Catch ex As Exception
                        ButtonRestartTab3478A.Enabled = True                             'allow user to restart the process if dialog to choose file is aborted
                        Exit Sub
                    End Try
                    fs = New System.IO.FileStream(RAMfilename3478A, IO.FileMode.Open)    'open the selected file
                    TextBoxCalRamLoadFile3478A.Text = RAMfilename3478A                   'display the selected file to the user
                    CalRamPathfile = New System.IO.BinaryWriter(fs)                      'prepare file for binary reading
                    CalRamPathfile.Seek(0, System.IO.SeekOrigin.Begin)                   'set the begin of the file as the start position for reading

                    'Read entire file regardless of user address choises
                    ReadCalibrationFile(fs, 0, 255, DataToLoad)

                    'Check if an attempt was made to read an invalid calibration ram data file
                    If RAMfilename3478A = Nothing Then
                        TextBoxCalRamLoadFile3478A.Text = Nothing                        'clear the invalid file name and path
                        ButtonRestartTab3478A.Enabled = True                             'allow user to restart after choosing an invalid file
                        Me.Refresh()
                        Exit Sub
                    End If

                    'Check if the 3478A serial number has been entered or if the datafile contains a serial number
                    If TextBoxSerialNumber3478A.Text <> Nothing Or DataToLoad(235) <> 64 Then

                        'Check if calibration range 19 contains a serial number
                        For i = 235 To 245                                                                      'adddress range of the 19th calibration range 
                            If i <> 239 And DataToLoad(i) > 63 Or SerialNumber = "0000" Then                    'no special serial number data detected
                                SerialNumber = SerialNumber & Convert.ToString(Chr(Asc(DataToLoad(i) - 64)))    'process the data byte normally
                            ElseIf DataToLoad(i) > 63 Then                                                      'a "country code" character detected at address 239 indicating serial number present   
                                SerialNumber = SerialNumber & Convert.ToString(Chr(DataToLoad(i)))              'use the ascii character as country code instead of the ascii number
                            End If
                        Next
                        If Mid(SerialNumber, 6, 1) = "0" Then                                                   'the 6th offset position of a valid serial number will be empty
                            SerialNumber = Mid(SerialNumber, 1, 5) & Mid(SerialNumber, 7, 5)                    'construct the serial number excluding the empty position 6 of the offset
                            If Mid(SerialNumber, 1, 6) <> "000000" Then
                                FileContainsSerNr = True                                                        'set flag indicating that a serial number is contained in the binary file being uploaded
                            End If
                        End If

                        'Check for serial number mismatch
                        If Not SerialNumber = TextBoxSerialNumber3478A.Text Then
                            Dim DecisionText As String = "File S/N " & SerialNumber & " and Unit's S/N " & TextBoxSerialNumber3478A.Text & " do NOT match!"
                            Dim decision As DialogResult = MessageBox.Show(DecisionText, "Continue anyway?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

                            'User decides whether to continue or to stop and clarify serial number issue
                            If decision = DialogResult.Yes Then
                                If SerialNumber = "0000000000" And ComboBoxCalRange3478A.Text <> Nothing Then
                                    DecisionText = "Answer YES to delete the serial number and continue the upload. Answer NO to continue the upload without deleting the serial number."
                                    decision = MessageBox.Show(DecisionText, "You have the option to delete the stored serial number.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                                    If decision = DialogResult.Yes Then

                                        'Delete the stored serial number at user's request and continue the upload
                                        TextBoxSerialNumber3478A.Text = Nothing
                                        LabelCalRange19.Text = "Not used"
                                    Else
                                        'continue without deleting the stored serial number
                                    End If
                                End If
                                SerNrChangeOK = True
                                LabelCalRange19.Text = "HP serial number"
                            Else
                                Button3478Aabort.Enabled = False
                                ButtonRestartTab3478A.Enabled = True
                                LabelCalRange19.Text = "HP serial number"
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                'Insure that the file contains valid calibration data
                VerifyCalibrationFileData()
            Else

                'If serial number has been added, allow immediate upload of the new serial number to the 3478A unit's memory
                If EditOccurred And EditedCalibrationData(18) = Nothing And CalibrationRangeStore3478A(18) <> "@@@@@@@@@@@OO" And ComboBoxCalRange3478A.Text = Nothing Then
                    EditedCalibrationData(18) = CalibrationRangeStore3478A(18)
                    LabelCalRange19.Text = "HP serial number"
                End If

                'Initialize edit flags that indicat data has been edited
                EditOccurred = False
                EditCount = 0

                'Check if calibration range data has been edited
                For i = 0 To 18
                    If Not EditedCalibrationData(i) = Nothing Then
                        For j = 1 To 13                                                   'transfer edited calibration range values into the data to load
                            CalramAddress3478A = i * 13 + j
                            CalramStore3478A(CalramAddress3478A) = Mid(EditedCalibrationData(i), j, 1)
                        Next
                        EditCount = EditCount + 1                                         'count how many calibration ranges have been edited
                        EditOccurred = True                                               'set flag indicating the calibration data has been edited
                    End If
                Next

                'Insure that the file contains valid calibration data after edit modifications
                VerifyCalibrationFileData()

                'Highlight range values which were modified by user editing
                For i = 0 To 18
                    If Not EditedCalibrationData(i) = Nothing Then
                        SetCalibrationRangeColor(i, Color.LightSeaGreen, 1, 13, 2)
                    End If
                Next

                'Convert the calibration string data into bytes and fill the data to load byte array
                For CalramAddress3478A = 0 To 255
                    If Not CalramStore3478A(CalramAddress3478A) = Nothing Then            'skip any empty addresses resulting from QUICK-AND-DIRTY editing
                        AsciiValue = Asc(CalramStore3478A(CalramAddress3478A))
                        DataToLoad(CalramAddress3478A) = Convert.ToByte(AsciiValue)
                    End If
                Next
                If DataToLoad(0) = 0 Then
                    QuickAndDirty = True                                                  'set flag indicating QUICK-AND-DIRTY processing
                End If
                If EditCount = 1 And Not QuickAndDirty Then
                    For i = 0 To 18
                        If Not EditedCalibrationData(i) = Nothing Then                    'check if the current calibration range data has been edited
                            If Not SuppressMsgs Then                                      'check if suppression of messages is requested
                                If i = 18 And EditedCalibrationData(i) = CalibrationRangeBackup3478A(i) Then
                                    'Suppress the messagebox if only changing the serial number
                                Else
                                    Dim answer As DialogResult = MessageBox.Show("A single calibration range has ben edited. Upload only this range?", "Confirm Action", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                                    If Not answer = DialogResult.Yes Then
                                        Exit For
                                    End If
                                End If
                            End If
                            'Adjust start and end addresses to match the address range of the edited calibration range
                            CalAddrStart3478A = ((i * 13) + 1)                        'calibration ranges are organized in block of 13 nibbles (6 offset, 5 gain and 2 checksum)
                            CalAddrEnd3478A = CalAddrStart3478A + 12                  'end address is equal to start address + 12
                            TextBox3478AFrom.Text = CalAddrStart3478A                 'display the edit start address
                            TextBox3478ATo.Text = CalAddrEnd3478A                     'display the edit end address
                            EditOccurred = True                                        'flag to signal generating new backup file of changeed calibration data when upload completed
                            VerifyCalibrationFileData()                               'insure that the file contains valid calibration data after edit modifications
                            SetCalibrationRangeColor(i, Color.Chocolate, 1, 13, 2)    'highlight the modified range values
                        End If
                    Next
                ElseIf QuickAndDirty Or EditCount > 1 Then

                    'Adjust the start and end addresses to encompass only addresses relevant to QUICK-AND-DIRTY process
                    Dim StartAddr As Integer
                    Dim EndAddr As Integer
                    For i = 0 To 18
                        If Not EditedCalibrationData(i) = Nothing Then
                            CalAddrStart3478A = ((i * 13) + 1)                           'calibration ranges are organized in block of 13 nibbles (6 offset, 5 gain and 2 checksum)
                            If StartAddr = 0 Then
                                StartAddr = CalAddrStart3478A                            'set the lowest needed start address
                            End If
                            CalAddrEnd3478A = CalAddrStart3478A + 12
                            If CalAddrEnd3478A > EndAddr Then
                                EndAddr = CalAddrEnd3478A                                'set the highest needed end address
                            End If
                        End If
                    Next
                    CalAddrStart3478A = StartAddr                                        'adjust the start address for  QUICK-AND-DIRTY process
                    CalAddrEnd3478A = EndAddr                                            'adjust the end address for  QUICK-AND-DIRTY process
                    TextBox3478AFrom.Text = StartAddr                                    'display the edit start address
                    TextBox3478ATo.Text = EndAddr                                        'display the edit end address
                End If
            End If

            'User gets a dialog to review data and confirm or cancel uploading the calibration data to the device
            If Not SuppressMsgs Then                                                     'check if suppression of messages is requested
                Dim result As DialogResult = MessageBox.Show("Continue to upload to the device? Review the displayed calibration data before deciding!", "Confirm Action", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                If result = DialogResult.Yes Then
                    'continue with the upload
                Else

                    'Abort the upload process and allow user to make new selection, edit calibration data, restart or leave program
                    ComboBoxCalRange3478A.Enabled = True
                    ButtonCalramLoad3478A.Enabled = True
                    ButtonCalramLoad3478A.Text = "3478A Upload"
                    Button3478Aabort.Enabled = False
                    ButtonRestartTab3478A.Enabled = True
                    Exit Sub
                End If
            End If
            'Check if a serial number is contained in the data to load
            If SerialNumber = Nothing Then
                For i = 235 To 245
                    If (i <> 239 And i <> 240) Or SerialNumber = "0000" Then
                        SerialNumber = SerialNumber & Convert.ToString(Chr(Asc(DataToLoad(i) - 64)))   'numeric "prefix" and "suffix" parts of the 3478A serial number
                    ElseIf i = 239 Then
                        SerialNumber = SerialNumber & CalramStore3478A(i)                              'charachter separating "prefix" and "suffix" denoting country of manufacture
                    End If
                Next
            End If

            If Not QuickAndDirty Then                                                                  'no serial number processing by QUICK-AND-DIRTY editing
                'Check for serial number mismatch
                If SerialNumber <> TextBoxSerialNumber3478A.Text And SerialNumber <> "0000000000" Then

                    'The serial number has been assigned but the upload file contains a different serial number or the user is changing the serial number
                    If Not SerNrChangeOK Then

                        'The change of serial number has not yet been confirmed
                        Dim DecisionText As String = "File S/N " & SerialNumber & " and Unit's S/N " & TextBoxSerialNumber3478A.Text & " do NOT match!"
                        Dim decision As DialogResult = MessageBox.Show(DecisionText, "Are you certain that you want to continue anyway?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

                        'User decides whether or not to continue with uploading the calibration ram data with the associated serial number
                        If decision = DialogResult.Yes Then
                            If FileContainsSerNr Then                                                          'the new serial number was taken from the upload-file which contained the serial number data

                                'Display the accepted new serial number
                                TextBoxSerialNumber3478A.Text = SerialNumber
                                LabelCalRange19.Text = "HP serial number"
                            Else                                                                               'the new serial number was entered by the user

                                'Transfer the new serial number to the data to upload
                                Dim j As Integer
                                Dim NewData As String
                                NewData = TextBoxSerialNumber3478A.Text

                                'First transfer the 4-digit "prefix" part of the serial number to the first four offset data fields of calibration range 19
                                For i = 235 To 238
                                    j = i - 234
                                    DataToLoad(i) = Convert.ToByte(Asc(Mid(NewData, j, 1)))
                                Next

                                'Next, append the "country code" character of the serial number to position 5 of the offset data field followed by a "space" character
                                DataToLoad(239) = Asc(Mid(NewData, 5, 1))
                                DataToLoad(240) = 32

                                'Finally, append the 5-digit "suffix" part of the serial number to the calibration gain data fields
                                For I = 241 To 245
                                    j = I - 235
                                    DataToLoad(I) = Convert.ToByte(Asc(Mid(NewData, j, 1)))
                                Next
                                LabelCalRange19.Text = "HP serial number"
                            End If
                        Else
                            'User has decided to abort and clarify the serial number issue
                            Button3478Aabort.Enabled = False
                            ButtonRestartTab3478A.Enabled = True
                            Exit Sub
                        End If
                    Else
                        TextBoxSerialNumber3478A.Text = SerialNumber
                        LabelCalRange19.Text = "HP serial number"
                    End If

                    'Check if the calibration ram data to be uploaded did not have a serial number associated with it
                ElseIf SerialNumber = "0000000000" And TextBoxSerialNumber3478A.Text <> Nothing Then
                    Dim NewData As String

                    'Restructure the serial number in the format needed to upload to the unit's RAM
                    TextBoxCalOffsetRaw.Text = Mid(TextBoxSerialNumber3478A.Text, 1, 5) & "@"
                    TextBoxCalGainRaw.Text = Mid(TextBoxSerialNumber3478A.Text, 6, 5)
                    GenerateNewChecksum()
                    NewData = Mid(TextBoxSerialNumber3478A.Text, 1, 5) & "@" & Mid(TextBoxSerialNumber3478A.Text, 6, 5) & TextBoxNewChecksum.Text
                    TextBoxCalOffsetRaw.Text = Nothing
                    TextBoxCalGainRaw.Text = Nothing
                    TextBoxNewChecksum.Text = Nothing

                    'Transfer the restructured serial number to the data to load array
                    For i = 235 To 245
                        Dim j As Integer
                        j = i - 234
                        DataToLoad(i) = Convert.ToByte(Asc(Mid(NewData, j, 1)))
                    Next
                End If
            End If

            'Check to see if the CAL ENABLE switch is in the calibrate position
            CheckCalEnableSwitch(False)

            'Check if the connection to the HP3478A is working
            If Not q.status = 0 Then

                'Error message will have already been displayed in the CheckCalEnableSwith() process
                Exit Sub
            End If

            If CalibrationSwitchDisabled Then

                'Confirm whether or not use has enabled the CAL ENABLE switch
                CalibrationSwitchDisabled = False
                CheckCalEnableSwitch(True)

                'Check if the connection to the HP3478A is working
                If Not q.status = 0 Then

                    'Error message will have already been displayed in the CheckCalEnableSwith() process
                    Exit Sub
                End If
                If CalibrationSwitchDisabled Then
                    MessageBox.Show("The selected calibration backup file content will now be compared with the current calibration ram memory content of the connected HP3478A.")
                End If
            End If

            'Display the current processing status
            CalramStatus3478A.Text = "SETTING UP GPIB"
            Me.Refresh()

            'Retrieve the data to load to the calibration memory
            For CalAddr3478A As Integer = CalAddrStart3478A To CalAddrEnd3478A

                'Skip empty addresses resulting from QUICK-And-DIRTY calibration value editing
                If DataToLoad(CalAddr3478A) = 0 Then
                    Continue For
                End If
                Me.Refresh()

                'Abort the process at user's request
                If Abort3478A Then
                    Dim answer As DialogResult = MessageBox.Show("Aborting the upload process may render the unit UNCALIBRATED! Are you sure you want to Abort?", "Confirm Action", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                    If answer = DialogResult.Yes Then
                        MessageBox.Show("The upload has been aborted. Use 3478A Read to verify calibration or 3478A Write to perform a new upload attempt, or edit as needed and repreat upload.")
                        Exit For
                    Else
                        Abort3478A = False
                    End If
                End If

                'Display the current processing status
                If CalibrationSwitchDisabled = False Then
                    CalramStatus3478A.Text = "WRITING........"
                Else
                    CalramStatus3478A.Text = "COMPARING......"
                End If

                dev1.enablepoll = False
                CalAddrHex = Convert.ToByte(CalAddr3478A)                               'setup address to upload the next byte to the 3478A calibration memory

                'Setup the command string that will be sent to the 3478A to update the appropriate memory byte
                CalAddrByte = Convert.ToByte(CalAddr3478A)                              'convert the integer address value to byte
                LoadData(0) = Convert.ToByte(88)                                        'convert ascii "X" into byte array
                LoadData(1) = CalAddrByte                                               'enter calibration ram address into byte array
                LoadData(2) = DataToLoad(CalAddr3478A)                                  'enter calibration data byte into byte array
                LoadData(3) = Nothing                                                   'initialize Prologix handling extra array byte
                LoadData(4) = Nothing                                                   'initialize Prologix handling extra array byte

                'Special handling for Prologix GPIB adapters (prefix specific special characters with ESC-character)
                If lstIntf1.Text = "Prologix Serial" Or lstIntf1.Text = "Prologix Ethernet" Then
                    Select Case CalAddr3478A
                        Case 10, 13, 27, 43                                             'ASCII characters that must be prefixed with ESC for Prologix adapters
                            LoadData(3) = LoadData(2)                                   'push the data byte down one array position to make room for the address byte
                            LoadData(2) = LoadData(1)                                   'push the address byte down one array position to make room for the ESC character
                            LoadData(1) = Convert.ToByte(27)                            'insert the ESC charcter to array position preceding the address byte
                    End Select
                    If LoadData(3) = Nothing Then
                        Select Case LoadData(2)
                            Case 10, 13, 27, 43                                         'ASCII characters that must be prefixed with ESC for Prologix adapters
                                LoadData(3) = LoadData(2)                               'push the data byte down one array position to make room for the ESC character
                                LoadData(2) = Convert.ToByte(27)                        'insert the ESC charcter to array position preceding the data byte
                        End Select
                    Else
                        Select Case LoadData(3)
                            Case 10, 13, 27, 43                                         'ASCII characters that must be prefixed with ESC for Prologix adapters
                                LoadData(4) = LoadData(3)                               'address and data bytes already pushed for ESC preceding address - now push data byte for it's ESC
                                LoadData(3) = Convert.ToByte(27)                        'insert the ESC charcter to array position preceding the pushed data byte
                        End Select
                    End If
                End If

                'Configure the load command and upload the current data byte
                LoadCommand = System.Text.Encoding.Default.GetString(LoadData)          'setup the upload command
                dev1.SendBlocking(LoadCommand, False)                                   'perform upload to the HP3478A calibration ram memory
                System.Threading.Thread.Sleep(50)                                       '50mS delay to allow device time to process command

                'Read back the data from the address just uploaded
                If Not GetCalRamByte(CalAddr3478A) = 0 Then
                    Exit Sub
                Else
                    ReadBackValue = CalramValue3478A
                End If

                'If the read back does not match the data to load, repeat the load operaton and check again
                If ReadBackValue <> (Convert.ToChar(DataToLoad(CalAddr3478A))).ToString() And CalAddr3478A <> 0 Then
                    CalramStatus3478A.Text = "1ST RETRY WRITING TO ADDRESS " & Convert.ToString(CalAddr3478A)
                    System.Threading.Thread.Sleep(100)
                    dev1.SendBlocking(LoadCommand, False)                                'retry writing to ram address after short delay (1st attempt)
                    System.Threading.Thread.Sleep(100)

                    'Readback processing for addresses up to the last (19th) calibration range
                    If GetResponse() <> (Convert.ToChar(DataToLoad(CalAddr3478A))).ToString() And CalAddr3478A < 235 Then
                        CalramStatus3478A.Text = "2ND RETRY WRITING TO ADDRESS " & Convert.ToString(CalAddr3478A)
                        System.Threading.Thread.Sleep(100)
                        dev1.SendBlocking(LoadCommand, False)                            'retry writing to ram address after short delay (2nd attempt)
                        System.Threading.Thread.Sleep(100)

                        'Readback processing for addresses which belong to the 19th calibration range (which might contain specially structured serial number data)
                    ElseIf GetResponse() <> Convert.ToChar(DataToLoad(CalAddr3478A) + 16) And q.ResponseAsString <> Convert.ToChar(DataToLoad(CalAddr3478A)) And CalAddr3478A <> 240 Then
                        CalramStatus3478A.Text = "2ND RETRY WRITING TO ADDRESS " & Convert.ToString(CalAddr3478A)
                        System.Threading.Thread.Sleep(100)
                        dev1.SendBlocking(LoadCommand, False)                            'retry writing to ram address after short delay (2nd attempt)
                        System.Threading.Thread.Sleep(100)
                    ElseIf Not DataToLoad(CalAddr3478A) = 32 Then                       'ignore "space" character the formatted serial number data, stored in RAM as ascii value 64 ("@")
                        CalramStatus3478A.Text = "2ND RETRY WRITING TO ADDRESS " & Convert.ToString(CalAddr3478A)
                        System.Threading.Thread.Sleep(100)
                        dev1.SendBlocking(LoadCommand, False)                            'retry writing to ram address after short delay (2nd attempt)
                        System.Threading.Thread.Sleep(100)
                    End If

                    Dim MsgBuffer As String
                    MsgBuffer = ControlChars.NewLine & ControlChars.NewLine & Convert.ToChar(DataToLoad(CalAddr3478A)) & " = File content" & ControlChars.NewLine & CalramValue3478A & " = HP3478A calibration ram content" & ControlChars.NewLine & ControlChars.NewLine

                    'Final check if addresses up to calibration range 19 were successful
                    If GetResponse() <> (Convert.ToChar(DataToLoad(CalAddr3478A))).ToString() And CalAddr3478A < 235 Then

                        'Setup error message text
                        If CalibrationSwitchDisabled = False Then
                            CalramStatus3478A.Text = "write error at ram address " & Convert.ToString(CalAddr3478A) & "!" & MsgBuffer
                        Else
                            CalramStatus3478A.Text = "compare error at ram address " & Convert.ToString(CalAddr3478A) & "!" & MsgBuffer
                        End If
                        'User gets a dialog informing that a write to calibration ram was unsuccessful
                        Dim result1 As DialogResult = MessageBox.Show(CalramStatus3478A.Text & " Continue anyway?", "Confirm Action", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                        If result1 = DialogResult.Yes Then
                            'continue with the upload
                            CalibrationSwitchDisabled = True                            'set the calibration switch flag to disabled
                        Else
                            ButtonRestartTab3478A.Enabled = True                        'enable button allowing restart 
                            Exit Sub                                                    'abort the upload process and allow user to make new selection or leave program
                        End If

                        'Final check if addresses used for storage of the serial number (calibration range 19) were successfull
                    ElseIf GetResponse() <> Convert.ToChar(DataToLoad(CalAddr3478A) + 16) And q.ResponseAsString <> Convert.ToChar(DataToLoad(CalAddr3478A)) And CalAddr3478A <> 240 Then

                        'Setup error message text
                        If CalibrationSwitchDisabled = False Then
                            CalramStatus3478A.Text = "write error at ram address " & Convert.ToString(CalAddr3478A) & "!" & MsgBuffer
                        Else
                            CalramStatus3478A.Text = "compare error at ram address " & Convert.ToString(CalAddr3478A) & "!" & MsgBuffer
                        End If

                        'User gets a dialog informing that a write to calibration ram was unsuccessful
                        Dim result1 As DialogResult = MessageBox.Show(CalramStatus3478A.Text & " Continue anyway?", "Confirm Action", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                        If result1 = DialogResult.Yes Then

                            'Continue with the upload
                            CalibrationSwitchDisabled = True                            'set the calibration switch flag to disabled
                        Else
                            ButtonRestartTab3478A.Enabled = True                        'enable button allowing restart 
                            Exit Sub                                                    'abort the upload process and allow user to make new selection or leave program
                        End If
                    ElseIf DataToLoad(CalAddr3478A) <> 32 And CalAddr3478A = 240 Then 'if address 240 is not a space (no serial number) then the write should be successful

                        'Setup error message text
                        If CalibrationSwitchDisabled = False Then
                            CalramStatus3478A.Text = "write error at ram address " & Convert.ToString(CalAddr3478A) & "!" & MsgBuffer
                        Else
                            CalramStatus3478A.Text = "compare error at ram address " & Convert.ToString(CalAddr3478A) & "!" & MsgBuffer
                        End If

                        'User gets a dialog informing that a write to calibration ram was unsuccessful
                        Dim result1 As DialogResult = MessageBox.Show(CalramStatus3478A.Text & " Continue anyway?", "Confirm Action", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                        If result1 = DialogResult.Yes Then

                            'Set the calibration switch flag to disabled and continue 
                            CalibrationSwitchDisabled = True
                        Else

                            'Abort the upload process and allow user to make new backup file selection, edit the calibration data, restart or leave program
                            ButtonCalramLoad3478A.Enabled = True
                            ButtonCalramLoad3478A.Text = "3478A Upload"
                            ButtonRestartTab3478A.Enabled = True                        'enable button allowing restart 
                            Exit Sub
                        End If
                    End If
                End If


                LabelCalramAddress3478A.Text = CalAddr3478A                             'display the current memory address
                LabelCounter3478A.Text = Counter3478A                                   'display the number of addresses processed

                'Hex conversion of the memory address
                LabelCalRamAddress3478AHex.Text = String.Join(",", LabelCalramAddress3478A.Text.Split(","c).
                              Select(Function(x) _
                              Convert.ToInt32(x).ToString("X")))

                'Display the calibration memory byte as an ASCII letter to match to values found in the upload file as shown in a hex-editor
                LabelCalRamByte3478A.Text = (Convert.ToChar(DataToLoad(CalAddr3478A))).ToString()
                LabelCalramAddress3478A.Text = Convert.ToChar(DataToLoad(CalAddr3478A))
                Counter3478A = Counter3478A + 1                                        'increment the counter for the next iteration
            Next

            RAMfilename3478A = ""

            'Display upload aborted at user request
            If Abort3478A = True Then
                Abort3478A = False
                CalramStatus3478A.Text = "ABORTED!"
                ButtonRestartTab3478A.Enabled = True                                'allow the user to perform a restart of the 3478A Cal tab functionality
                ComboBoxCalRange3478A.Enabled = True                                'allow the user the option to edit the calibration data
            Else
                'Display finished uploading or comparing calibration ram data from file
                LabelCalramAddress3478A.Text = CalAddrEnd3478A
                CalramStatus3478A.Text = "DONE!"

                'If the calibration ram has been changed, generate a new backup file containing the new updated calibration ram data
                If EditOccurred Then
                    If CalibrationSwitchDisabled = False Then
                        If DataToLoad(235) = 0 Or QuickAndDirty Then
                            If Not SuppressMsgs Then                                'check if suppression of messages is requested
                                MessageBox.Show("A QUICK-AND-DIRTY calibration edit was performed. No new binary file containing the changes has been generated. Use 3478A Read to generate an updated file if needed.")
                            End If
                        Else

                            'Generate new backup file containing the unit's new calibration ram data and display the file name
                            Button3478Aabort.Enabled = False                        'prevent aborting the backup 
                            DumpCalram3478A()
                            If Not SuppressMsgs Then                                'check if suppression of messages is requested
                                MessageBox.Show("New backup of changed calibration ram data has been created: file " & TextBoxCalramFile3478A.Text)
                            End If
                            'Prevent needless generation of identical backup files
                            ButtonCalramDump3457A.Enabled = False
                        End If
                        WriteSuccess3478A = True

                        'Clear the editing buffer of the uploaded changes
                        For i = 0 To 18
                            EditedCalibrationData(i) = Nothing
                        Next
                    End If
                Else
                    'Prevent generating a binary file containing data identical to the data contained in the uploaded file
                    ButtonCalramDump3478A.Enabled = False
                End If

                'Adjust program flow options
                ButtonCalramLoad3478A.Enabled = False                               'prevent uploading the same data to the unit's calibration memory more than once
                ButtonDeleteSerNr.Enabled = False                                   'prevent repeating deletion process of a deleted serial number
                ButtonDeleteSerNr.Visible = False                                   'hide the Delete Serial Number button until needed again
                ButtonRestartTab3478A.Enabled = True                                'allow the user to perform a restart of the 3478A Cal tab functionality
                ComboBoxCalRange3478A.Enabled = True                                'allow the user the option to edit the calibration data
                Button3478Aabort.Enabled = False                                    'no sense aborting if no process is running

            End If
        Else

            'GPIB Dev 1 has not been connected
            CalramStatus3478A.Text = "DEVICE 1 MUST BE CONNECTED TO ENABLE UPLOADING CALIBRATION DATA"
        End If


    End Sub


    Private Sub Button3478Aabort_Click(sender As Object, e As EventArgs) Handles Button3478Aabort.Click

        'Flag user's request to abort the current process
        Abort3478A = True
        TextBoxCalramFile3478A.Text = ""
    End Sub
    Private Sub ShowFilesCalRam3478A_Click(sender As Object, e As EventArgs) Handles ShowFilesCalRam3478A.Click

        'Open Windows Explorer with path for selection of a backup file to be restored to the HP3478A
        Process.Start("explorer.exe", String.Format("/n, /e, {0}", CSVfilepath.Text))
    End Sub

    Private Sub ReadCalibrationFile(fs As IO.FileStream, StartAddress As Integer, EndAddress As Integer, DataToLoad As Byte())

        'Read the selected binary file into the upload data array buffer
        For CalAddr3478A = StartAddress To EndAddress
            Try
                DataToLoad(CalAddr3478A) = fs.ReadByte()                                         'read the next byte from the upload file
                CalramStore3478A(CalAddr3478A) = Convert.ToChar(DataToLoad(CalAddr3478A))        'store it in the ram calibration data array buffer
            Catch ex As Exception
                MessageBox.Show(RAMfilename3478A & " does not contain valid HP3478A calibration data for uploading with this utility")
                fs.Close()
                OpenFileDialog1.Reset()
                RAMfilename3478A = Nothing
                For i = 0 To CalAddr3478A
                    CalramStore3478A(i) = Nothing
                    DataToLoad(i) = Nothing
                Next
                ButtonCalramLoad3478A.Enabled = True
                Exit Sub
            End Try
        Next
        'Close file
        fs.Close()
    End Sub

    Private Sub VerifyCalibrationFileData()

        'Calculate checksums to Verify the calibration range data being processed 
        CalramStatus3478A.Text = "Validating checksums..."
        System.Threading.Thread.Sleep(500)     ' 500mS delay

        'Transfer the data read from the binary file into 19 calibration range data arrays
        For CalibrationRange = 0 To 18
            Dim ValueSum As Integer
            Dim CheckSumArray(1) As Byte
            Dim Address As Integer
            Dim RamValue As Char
            Dim CheckSum As Byte
            Dim RawOffset As String
            Dim RawGain As String
            Dim ChecksumChar As String
            Dim ChecksumValid As String
            Dim GainBuffer As String
            Dim OffsetBuffer As String
            Dim Multiplier As Integer
            Dim PosVal As Integer
            Dim OffsetValue As Integer
            Dim gain As Double
            Dim temp As Double
            Dim Buffer As String

            If Abort3478A = True Then
                Exit Sub
            End If

            'Calibration values are organized in blocks of 11 nibbles + 2 checksum nibbles starting at 2nd ram address (address 01)
            Address = 1 + CalibrationRange * 13

            'Skip addresses left empty as a result of a "QUICK-AND-DIRTY" calibration data edit 
            If CalramStore3478A(Address) = Nothing Then
                Continue For
            End If


            ValueSum = 0
            'Accumulate raw offset, raw gain and checksum values
            For CheckSumValues = 1 To 11
                RamValue = Convert.ToChar(CalramStore3478A(Address))
                ValueSum = ValueSum + Convert.ToInt32(RamValue) - 64                                    'accumulate checksum
                If CheckSumValues < 7 Then                                                              'first 6 nibbles of ram data represent the raw offset
                    RawOffset = RawOffset & RamValue
                ElseIf CheckSumValues < 12 Then                                                         'next 5 nibbles of ram data represent the raw gain
                    RawGain = RawGain & RamValue
                End If
                Address = Address + 1
            Next

            If CalibrationRange = 18 And Not NewOffset(18) = "0" And Not NewGain(18) = "1" Then         'check if the serial number is being intentionally deleted

                'Treat the last calibration range (the 19th - unused) as possible storage area for the HP3478A's serial number
                ChecksumChar = CalramStore3478A(246) & CalramStore3478A(247)
                OffsetBuffer = RawOffset
                GainBuffer = RawGain

                If (TextBoxSerialNumber3478A.Text <> Nothing And TextBoxSerialNumber3478A.Text <> "0000000000") Or FileContainsSerNr Then

                    'Terminate editiing the serial number 
                    TextBoxSerialNumber3478A.BackColor = Color.WhiteSmoke
                    TextBoxSerialNumber3478A.Enabled = False

                    'Integrate the serial number into the calibration range data fields
                    RawOffset = Mid(TextBoxSerialNumber3478A.Text, 1, 5) & " "                          'serial number "prefix" plus "country code"
                    Offset = RawOffset
                    If Mid(TextBoxSerialNumber3478A.Text, 6, 1) = " " Then                              '6th position of the offset is empty when the offset contains a valid serial number
                        RawGain = Mid(TextBoxSerialNumber3478A.Text, 7, 5)                              'serial number "suffix" is found in the "gain" portion of the 19th calibration range
                        TextBoxSerialNumber3478A.Text = Mid(Offset, 1, 5) & RawGain                     'display the "prefix" and "country code" in the serial number textbox
                    Else
                        RawGain = Mid(TextBoxSerialNumber3478A.Text, 6, 5)                              'append the "suffix" to the serial number textbox
                    End If
                    GainValue = RawGain
                    RawOffset = Mid(RawOffset, 1, 5) & "@"                                                 'replace space in 6th offset position with "@" as zero value offset place holder

                    'Buffer the serial number data
                    CalibrationRangeStore3478A(CalibrationRange) = RawOffset & RawGain & ChecksumChar

                    If FileContainsSerNr Then
                        'Translate raw offset and raw gain values into serial number found in the binary upload file into plain text format
                        Offset = Nothing
                        GainValue = Nothing
                        For i = 235 To 238                                                              'serial number "prefix" storage area in the 19th calibration range data
                            Offset = Offset & Convert.ToString(DataToLoad(i) - 64)                      'plain text "prefix" translated from the binary storage format
                        Next
                        Offset = Offset & CalramStore3478A(239)                                         'serial number "country code" storage positon in the 19th calibration range data
                        For i = 241 To 245                                                              'serial number "suffix" storage area in the 19th calibration range data
                            GainValue = GainValue & Convert.ToString(DataToLoad(i) - 64)                'plain text "suffix" translated from the binary storage format
                        Next
                        Offset = Offset & " "                                                           'fill position 6 of the offset (unused by the serial number) with a "space" character 
                        If ChecksumChar = "OO" Then
                            ChecksumValid = "invalid"                                                   'a serial with checksum 255 ("OO") is invalid
                        End If
                    Else
                        'Translate raw offset and raw gain values from the user editing data into binary storage format
                        OffsetBuffer = Nothing
                        For i = 1 To 4                                                                                 'plain text serial number "prefix"
                            OffsetBuffer = OffsetBuffer & Convert.ToString(Chr(Asc(Mid(RawOffset, i, 1)) + 16))        'translate plain text "prefix" into binary storage format
                        Next
                        OffsetBuffer = OffsetBuffer & Mid(RawOffset, 5, 1)                                             'serial number "country code" must not be translated
                        GainBuffer = Nothing
                        For i = 1 To 5                                                                                 'plain text serial number "suffix"
                            GainBuffer = GainBuffer & Convert.ToString(Chr(Asc(Mid(RawGain, i, 1)) + 16))              'translate plain text "suffix" into binary storage format
                        Next
                        OffsetBuffer = OffsetBuffer & "@"                                                              'append the offsetbuffer with 6th position place holder
                        CalibrationRangeStore3478A(CalibrationRange) = OffsetBuffer & GainBuffer & ChecksumChar        'transfer the binary serial number data and checksum to calibration range 19
                    End If

                    If ChecksumChar = "OO" And Not FileContainsSerNr Then

                        'Generate a checksum for the serial number
                        Buffer = TextBoxCalOffsetRaw.Text & TextBoxCalGainRaw.Text & TextBoxNewChecksum.Text                'buffer any current values
                        TextBoxCalOffsetRaw.Text = OffsetBuffer                                                             'setup offset for checksum generation
                        TextBoxCalGainRaw.Text = GainBuffer                                                                 'setup gain for checksum generation
                        GenerateNewChecksum()                                                                               'generate the checksum
                        ChecksumValid = "valid"                                                                             'a gernerated checksum is always valid
                        ChecksumChar = TextBoxNewChecksum.Text                                                              'buffer the new checksum
                        CalibrationRangeStore3478A(CalibrationRange) = OffsetBuffer & GainBuffer & ChecksumChar             'transfer the binary serial number data and checksum to calibration range 19

                        'Restore any buffered values for the temporaily used offset, gain and checksum textboxes
                        If Not Buffer.Length = 0 Then
                            TextBoxCalOffsetRaw.Text = Mid(Buffer, 1, 6)
                            TextBoxCalGainRaw.Text = Mid(Buffer, 7, 5)
                            TextBoxNewChecksum.Text = Mid(Buffer, 12, 2)
                        Else
                            TextBoxCalOffsetRaw.Text = Nothing
                            TextBoxCalGainRaw.Text = Nothing
                            TextBoxNewChecksum.Text = Nothing
                        End If
                    End If

                    'Display the serial number in the 19th calibration range data to the user
                    SetCalibrationRangeColor(CalibrationRange, Color.Blue, StartAddrInconsistance(0), 0, False)
                    DisplayCalibration(CalibrationRange, OffsetBuffer, Offset, GainBuffer, GainValue, ChecksumChar, ChecksumValid)

                    'Insure a buffer of the calibratition range data is available for reverting edits if needed
                    If CalibrationRangeBackup3478A(CalibrationRange) = Nothing Then
                        CalibrationRangeBackup3478A(CalibrationRange) = CalibrationRangeStore3478A(CalibrationRange)
                        CalibrationOffsetBackup(CalibrationRange) = Offset
                        CalibrationGainBackup(CalibrationRange) = GainValue
                        CalibrationValidBackup(CalibrationRange) = ChecksumValid
                        SerialNumberBackup = TextBoxSerialNumber3478A.Text
                    End If
                    Exit For
                Else
                    'Initialize the 19th calibration range since no serial number data is available
                    Offset = "0"
                    RawGain = "@@@@@"
                    GainValue = "1"
                    SetCalibrationRangeColor(CalibrationRange, Color.Blue, StartAddrInconsistance(0), 0, False)
                End If

            Else
                'Convert BCD offset to integer from the ram's raw offset sequence (nibbles 1 thru 6 of calibration range data)
                OffsetValue = 0
                Multiplier = 100000
                For Position = 1 To 6                                                                       'each successive digit represents an order of magnitude smaller value
                    RamValue = Convert.ToChar(Mid(RawOffset, Position, 1))                                  'go through the 6 digit offset string one by one
                    PosVal = Convert.ToInt32(RamValue) - 64                                                 'convert the byte value into an integer value
                    OffsetValue = OffsetValue + (PosVal * Multiplier)                                       'accumulate the individual digit value to the total offset value
                    Multiplier = Multiplier / 10                                                            'adjust the order of magnitude for the next digit
                Next

                'Adjust the offset total for negative offset values
                If OffsetValue > 999 Then
                    OffsetValue = OffsetValue - 1000000
                End If
                Offset = Convert.ToString(OffsetValue)


                'Only the correction component of the gain is coded in the calibration data as the value after the decimal delimiter
                gain = 1.0

                'Calculate the gain value from the ram's gain string sequence (nibbles 7 thru 11 of calibration range data)
                For cur = 1 To 5                                                                            'go through the 5 digit gain string one by one
                    RamValue = Convert.ToChar(Mid(RawGain, cur, 1))                                         'convert the current gain digit value to char
                    temp = Convert.ToInt32(RamValue) - 64                                                   'convert the gain digit char value to integer
                    If (temp >= 8) Then                                                                     'adjust the value for negative gain
                        temp = temp - 16
                    End If
                    gain = gain + (temp / (10 ^ (cur + 1)))                                                 'adjust the gain digit value by successive order of magnitude and accumulate total
                Next
                GainValue = Convert.ToString(gain)

                'Reformat the checksum values for display
                RamValue = Convert.ToChar(CalramStore3478A(Address))
                CheckSumArray(0) = Convert.ToInt32(RamValue) - 64
                ChecksumChar = RamValue
                Address = Address + 1
                RamValue = Convert.ToChar(CalramStore3478A(Address))
                CheckSumArray(1) = Convert.ToInt32(RamValue) - 64
                ChecksumChar = ChecksumChar & RamValue
                CheckSum = CheckSumArray(0) * 16 + CheckSumArray(1)

                'Adjust the color of the calibration range displayed data according to whether or not the checksum is valid
                If CheckSum + ValueSum = 255 Then
                    ChecksumValid = "valid"
                    Select Case CalibrationRange
                        Case 5, 16, 18

                            'Highlight the unused calibration ranges blue
                            SetCalibrationRangeColor(CalibrationRange, Color.Blue, 0, 0, False)
                        Case Else

                            'Indicate normal valid calibration ranges with green color
                            If Not StartAddrInconsistance(0) = 0 And (CalibrationRange + 1) * 13 >= StartAddrInconsistance(1) Then
                                SetCalibrationRangeColor(CalibrationRange, Color.Green, StartAddrInconsistance(0), 0, False)
                                StartAddrInconsistance(0) = 0
                            Else
                                SetCalibrationRangeColor(CalibrationRange, Color.Green, 0, 0, False)
                            End If

                            If (CalibrationRange + 1) * 13 > EndAddrInconsistance(1) Then
                                SetCalibrationRangeColor(CalibrationRange, Color.Green, 0, EndAddrInconsistance(0), False)
                                EndAddrInconsistance(0) = 0
                            End If
                    End Select
                Else

                    'Highlight calibration ranges with invalid checksums red (except the unused calibration ranges)
                    ChecksumValid = "invalid"
                    Select Case CalibrationRange
                        Case 5, 16, 18

                            'Highlight the unused calibration ranges blue (even if having invalid checksums)
                            SetCalibrationRangeColor(CalibrationRange, Color.Blue, 0, 0, False)
                        Case Else

                            'Highlight normal calibration ranges with invalid checksums red
                            If Not StartAddrInconsistance(0) = 0 And (CalibrationRange + 1) * 13 >= StartAddrInconsistance(1) Then
                                SetCalibrationRangeColor(CalibrationRange, Color.Red, StartAddrInconsistance(0), 0, False)
                                StartAddrInconsistance(0) = 0
                            Else
                                SetCalibrationRangeColor(CalibrationRange, Color.Red, 0, 0, False)
                            End If

                            If (CalibrationRange + 1) * 13 > EndAddrInconsistance(1) Then
                                SetCalibrationRangeColor(CalibrationRange, Color.Red, 0, EndAddrInconsistance(0), False)
                                EndAddrInconsistance(0) = 0
                            End If
                    End Select
                End If
            End If

            'Buffer the calibration range data for possible use when rolling back editing changes
            CalibrationRangeStore3478A(CalibrationRange) = RawOffset & RawGain & ChecksumChar
            NewCheckSum3478A(CalibrationRange) = ChecksumChar
            If CalibrationRangeBackup3478A(CalibrationRange) = Nothing Then
                CalibrationRangeBackup3478A(CalibrationRange) = CalibrationRangeStore3478A(CalibrationRange)
                CalibrationOffsetBackup(CalibrationRange) = Offset
                CalibrationGainBackup(CalibrationRange) = GainValue
                CalibrationValidBackup(CalibrationRange) = ChecksumValid
            End If

            'Display the calibration data to the user
            DisplayCalibration(CalibrationRange, RawOffset, Offset, RawGain, GainValue, ChecksumChar, ChecksumValid)

            'Clean up the serial number variables
            RawOffset = ""
            RawGain = ""
            ChecksumChar = ""
        Next

    End Sub
    Private Sub DisplayCalibration(CalibrationRange As Integer, RawOffset As String, Offset As String, RawGain As String, GainValue As String, ChecksumChar As String, ChecksumValid As String)

        'Display the calibration ram data to the user (there must be a better way of doing this!)
        Select Case CalibrationRange
            Case 0
                Label01RawOffset.Text = RawOffset
                Label01RawGain.Text = RawGain
                Label01Checksum.Text = ChecksumChar
                Label01Valid.Text = ChecksumValid
                Label01Offset.Text = Offset
                Label01Gain.Text = GainValue
            Case 1
                Label02RawOffset.Text = RawOffset
                Label02RawGain.Text = RawGain
                Label02Checksum.Text = ChecksumChar
                Label02Valid.Text = ChecksumValid
                Label02Offset.Text = Offset
                Label02Gain.Text = GainValue
            Case 2
                Label03RawOffset.Text = RawOffset
                Label03RawGain.Text = RawGain
                Label03Checksum.Text = ChecksumChar
                Label03Valid.Text = ChecksumValid
                Label03Offset.Text = Offset
                Label03Gain.Text = GainValue
            Case 3
                Label04RawOffset.Text = RawOffset
                Label04RawGain.Text = RawGain
                Label04Checksum.Text = ChecksumChar
                Label04Valid.Text = ChecksumValid
                Label04Offset.Text = Offset
                Label04Gain.Text = GainValue
            Case 4
                Label05RawOffset.Text = RawOffset
                Label05RawGain.Text = RawGain
                Label05Checksum.Text = ChecksumChar
                Label05Valid.Text = ChecksumValid
                Label05Offset.Text = Offset
                Label05Gain.Text = GainValue
            Case 5
                Label06RawOffset.Text = RawOffset
                Label06RawGain.Text = RawGain
                Label06Checksum.Text = ChecksumChar
                Label06Valid.Text = ChecksumValid
                Label06Offset.Text = Offset
                Label06Gain.Text = GainValue
            Case 6
                Label07RawOffset.Text = RawOffset
                Label07RawGain.Text = RawGain
                Label07Checksum.Text = ChecksumChar
                Label07Valid.Text = ChecksumValid
                Label07Offset.Text = Offset
                Label07Gain.Text = GainValue
            Case 7
                Label08RawOffset.Text = RawOffset
                Label08RawGain.Text = RawGain
                Label08Checksum.Text = ChecksumChar
                Label08Valid.Text = ChecksumValid
                Label08Offset.Text = Offset
                Label08Gain.Text = GainValue
            Case 8
                Label09RawOffset.Text = RawOffset
                Label09RawGain.Text = RawGain
                Label09Checksum.Text = ChecksumChar
                Label09Valid.Text = ChecksumValid
                Label09Offset.Text = Offset
                Label09Gain.Text = GainValue
            Case 9
                Label10RawOffset.Text = RawOffset
                Label10RawGain.Text = RawGain
                Label10Checksum.Text = ChecksumChar
                Label10Valid.Text = ChecksumValid
                Label10Offset.Text = Offset
                Label10Gain.Text = GainValue
            Case 10
                Label11RawOffset.Text = RawOffset
                Label11RawGain.Text = RawGain
                Label11Checksum.Text = ChecksumChar
                Label11Valid.Text = ChecksumValid
                Label11Offset.Text = Offset
                Label11Gain.Text = GainValue
            Case 11
                Label12RawOffset.Text = RawOffset
                Label12RawGain.Text = RawGain
                Label12Checksum.Text = ChecksumChar
                Label12Valid.Text = ChecksumValid
                Label12Offset.Text = Offset
                Label12Gain.Text = GainValue
            Case 12
                Label13RawOffset.Text = RawOffset
                Label13RawGain.Text = RawGain
                Label13Checksum.Text = ChecksumChar
                Label13Valid.Text = ChecksumValid
                Label13Offset.Text = Offset
                Label13Gain.Text = GainValue
            Case 13
                Label14RawOffset.Text = RawOffset
                Label14RawGain.Text = RawGain
                Label14Checksum.Text = ChecksumChar
                Label14Valid.Text = ChecksumValid
                Label14Offset.Text = Offset
                Label14Gain.Text = GainValue
            Case 14
                Label15RawOffset.Text = RawOffset
                Label15RawGain.Text = RawGain
                Label15Checksum.Text = ChecksumChar
                Label15Valid.Text = ChecksumValid
                Label15Offset.Text = Offset
                Label15Gain.Text = GainValue
            Case 15
                Label16RawOffset.Text = RawOffset
                Label16RawGain.Text = RawGain
                Label16Checksum.Text = ChecksumChar
                Label16Valid.Text = ChecksumValid
                Label16Offset.Text = Offset
                Label16Gain.Text = GainValue
            Case 16
                Label17RawOffset.Text = RawOffset
                Label17RawGain.Text = RawGain
                Label17Checksum.Text = ChecksumChar
                Label17Valid.Text = ChecksumValid
                Label17Offset.Text = Offset
                Label17Gain.Text = GainValue
            Case 17
                Label18RawOffset.Text = RawOffset
                Label18RawGain.Text = RawGain
                Label18Checksum.Text = ChecksumChar
                Label18Valid.Text = ChecksumValid
                Label18Offset.Text = Offset
                Label18Gain.Text = GainValue
            Case 18
                Label19RawOffset.Text = RawOffset
                Label19RawGain.Text = RawGain
                Label19Checksum.Text = ChecksumChar
                Label19Valid.Text = ChecksumValid
                Label19Offset.Text = Offset
                Label19Gain.Text = GainValue
        End Select

        'Buffer the calibration validity flag
        CalibrationValidBackup(CalibrationRange) = ChecksumValid

    End Sub

    Private Sub SetCalibrationRangeColor(CalibrationRange As Integer, RangeColor As Color, StartInconsistance As Integer, EndInconsistance As Integer, Mode As Integer)

        Dim Color(2) As Color
        Dim color2 As Color

        Color(0) = RangeColor
        Color(1) = RangeColor
        Color(2) = RangeColor
        color2 = RangeColor

        If Mode < 2 Then

            'Highlight selected ram addresses that do not match to exactly corresponding calibration ranges
            If StartInconsistance > 0 Then
                If StartInconsistance > 10 Then
                    Color(1) = color2.Orange
                    Color(0) = color2.Black
                    Color(2) = color2.Black
                ElseIf StartInconsistance < 6 Then
                    Color(0) = color2.Orange
                    Color(2) = RangeColor
                    Color(1) = RangeColor
                Else
                    Color(2) = color2.Orange
                    Color(0) = color2.Black
                    Color(1) = RangeColor
                End If
                If Mode = 0 Then
                    Mode = 1
                    RangeColor = color2.Black
                End If
            End If
            If EndInconsistance > 0 Then
                If EndInconsistance > 10 Then
                    Color(1) = color2.Orange
                    Color(0) = RangeColor
                    Color(2) = RangeColor
                ElseIf EndInconsistance < 7 Then
                    Color(0) = color2.Orange
                    Color(2) = color2.Black
                    Color(1) = color2.Black
                Else
                    Color(2) = color2.Orange
                    Color(1) = color2.Black
                    Color(0) = RangeColor
                End If
                If Mode = 0 Then
                    Mode = 1
                    RangeColor = color2.Black
                End If
            End If
        End If

        'Highlight the selected calibration ram data 
        If (CalAddrStart3478A <= (1 + CalibrationRange) * 13 And CalAddrEnd3478A >= (1 + CalibrationRange * 13)) Or Mode = 1 Then
            Select Case CalibrationRange
                Case 0
                    Label01RawOffset.ForeColor = Color(0)
                    Label01RawGain.ForeColor = Color(2)
                    Label01Checksum.ForeColor = Color(1)
                    Label01Valid.ForeColor = RangeColor
                    Label01Offset.ForeColor = RangeColor
                    Label01Gain.ForeColor = RangeColor
                Case 1
                    Label02RawOffset.ForeColor = Color(0)
                    Label02RawGain.ForeColor = Color(2)
                    Label02Checksum.ForeColor = Color(1)
                    Label02Valid.ForeColor = RangeColor
                    Label02Offset.ForeColor = RangeColor
                    Label02Gain.ForeColor = RangeColor
                Case 2
                    Label03RawOffset.ForeColor = Color(0)
                    Label03RawGain.ForeColor = Color(2)
                    Label03Checksum.ForeColor = Color(1)
                    Label03Valid.ForeColor = RangeColor
                    Label03Offset.ForeColor = RangeColor
                    Label03Gain.ForeColor = RangeColor
                Case 3
                    Label04RawOffset.ForeColor = Color(0)
                    Label04RawGain.ForeColor = Color(2)
                    Label04Checksum.ForeColor = Color(1)
                    Label04Valid.ForeColor = RangeColor
                    Label04Offset.ForeColor = RangeColor
                    Label04Gain.ForeColor = RangeColor
                Case 4
                    Label05RawOffset.ForeColor = Color(0)
                    Label05RawGain.ForeColor = Color(2)
                    Label05Checksum.ForeColor = Color(1)
                    Label05Valid.ForeColor = RangeColor
                    Label05Offset.ForeColor = RangeColor
                    Label05Gain.ForeColor = RangeColor
                Case 5
                    Label06RawOffset.ForeColor = Color(0)
                    Label06RawGain.ForeColor = Color(2)
                    Label06Checksum.ForeColor = Color(1)
                    Label06Valid.ForeColor = RangeColor
                    Label06Offset.ForeColor = RangeColor
                    Label06Gain.ForeColor = RangeColor
                Case 6
                    Label07RawOffset.ForeColor = Color(0)
                    Label07RawGain.ForeColor = Color(2)
                    Label07Checksum.ForeColor = Color(1)
                    Label07Valid.ForeColor = RangeColor
                    Label07Offset.ForeColor = RangeColor
                    Label07Gain.ForeColor = RangeColor
                Case 7
                    Label08RawOffset.ForeColor = Color(0)
                    Label08RawGain.ForeColor = Color(2)
                    Label08Checksum.ForeColor = Color(1)
                    Label08Valid.ForeColor = RangeColor
                    Label08Offset.ForeColor = RangeColor
                    Label08Gain.ForeColor = RangeColor
                Case 8
                    Label09RawOffset.ForeColor = Color(0)
                    Label09RawGain.ForeColor = Color(2)
                    Label09Checksum.ForeColor = Color(1)
                    Label09Valid.ForeColor = RangeColor
                    Label09Offset.ForeColor = RangeColor
                    Label09Gain.ForeColor = RangeColor
                Case 9
                    Label10RawOffset.ForeColor = Color(0)
                    Label10RawGain.ForeColor = Color(2)
                    Label10Checksum.ForeColor = Color(1)
                    Label10Valid.ForeColor = RangeColor
                    Label10Offset.ForeColor = RangeColor
                    Label10Gain.ForeColor = RangeColor
                Case 10
                    Label11RawOffset.ForeColor = Color(0)
                    Label11RawGain.ForeColor = Color(2)
                    Label11Checksum.ForeColor = Color(1)
                    Label11Valid.ForeColor = RangeColor
                    Label11Offset.ForeColor = RangeColor
                    Label11Gain.ForeColor = RangeColor
                Case 11
                    Label12RawOffset.ForeColor = Color(0)
                    Label12RawGain.ForeColor = Color(2)
                    Label12Checksum.ForeColor = Color(1)
                    Label12Valid.ForeColor = RangeColor
                    Label12Offset.ForeColor = RangeColor
                    Label12Gain.ForeColor = RangeColor
                Case 12
                    Label13RawOffset.ForeColor = Color(0)
                    Label13RawGain.ForeColor = Color(2)
                    Label13Checksum.ForeColor = Color(1)
                    Label13Valid.ForeColor = RangeColor
                    Label13Offset.ForeColor = RangeColor
                    Label13Gain.ForeColor = RangeColor
                Case 13
                    Label14RawOffset.ForeColor = Color(0)
                    Label14RawGain.ForeColor = Color(2)
                    Label14Checksum.ForeColor = Color(1)
                    Label14Valid.ForeColor = RangeColor
                    Label14Offset.ForeColor = RangeColor
                    Label14Gain.ForeColor = RangeColor
                Case 14
                    Label15RawOffset.ForeColor = Color(0)
                    Label15RawGain.ForeColor = Color(2)
                    Label15Checksum.ForeColor = Color(1)
                    Label15Valid.ForeColor = RangeColor
                    Label15Offset.ForeColor = RangeColor
                    Label15Gain.ForeColor = RangeColor
                Case 15
                    Label16RawOffset.ForeColor = Color(0)
                    Label16RawGain.ForeColor = Color(2)
                    Label16Checksum.ForeColor = Color(1)
                    Label16Valid.ForeColor = RangeColor
                    Label16Offset.ForeColor = RangeColor
                    Label16Gain.ForeColor = RangeColor
                Case 16
                    Label17RawOffset.ForeColor = Color(0)
                    Label17RawGain.ForeColor = Color(2)
                    Label17Checksum.ForeColor = Color(1)
                    Label17Valid.ForeColor = RangeColor
                    Label17Offset.ForeColor = RangeColor
                    Label17Gain.ForeColor = RangeColor
                Case 17
                    Label18RawOffset.ForeColor = Color(0)
                    Label18RawGain.ForeColor = Color(2)
                    Label18Checksum.ForeColor = Color(1)
                    Label18Valid.ForeColor = RangeColor
                    Label18Offset.ForeColor = RangeColor
                    Label18Gain.ForeColor = RangeColor
                Case 18
                    Label19RawOffset.ForeColor = Color(0)
                    Label19RawGain.ForeColor = Color(2)
                    Label19Checksum.ForeColor = Color(1)
                    Label19Valid.ForeColor = RangeColor
                    Label19Offset.ForeColor = RangeColor
                    Label19Gain.ForeColor = RangeColor
            End Select
        End If
    End Sub

    Private Sub ButtonCommitEdit3478A_Click(sender As Object, e As EventArgs) Handles ButtonCommitEdit3478A.Click
        Dim Buffer As String

        'Update the buffered calibration range data with the edited calibration range data
        For i = 0 To 18
            If EditedCalibrationData(i) <> Nothing Then

                'Warn the user if serial number is being deleted
                If i = 18 And EditedCalibrationData(i) = "@@@@@@@@@@@OO" And (TextBoxEditGain.Text = "1" Or TextBoxEditGain.Text = "@@@@@") Then
                    If Not TextBoxSerialNumber3478A.Text = Nothing Then
                        Dim result1 As DialogResult = MessageBox.Show("Are you certain that you want to delete the serial number?", "Confirm Action", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                        If result1 = DialogResult.Yes Then
                            TextBoxSerialNumber3478A.Text = Nothing
                            ButtonDeleteSerNr.Enabled = False
                            LabelCalRange19.Text = "Not used"
                            Me.Refresh()
                            NewOffset(i) = 0
                            NewGain(i) = 1
                        Else

                            'Abort the upload process and allow user to make new range selection, edit the calibration data, restart or leave program
                            ButtonCalramLoad3478A.Enabled = True
                            ButtonCalramLoad3478A.Text = "3478A Upload"
                            RollbackEdits()
                            Exit Sub
                        End If
                    End If
                ElseIf i = 18 Then

                    'Warn the user if serial number is being edited
                    If Not TextBoxSerialNumber3478A.Text = Nothing Then
                        Try
                            For j = 1 To 11
                                Buffer = Convert.ToString(Asc(Mid(TextBoxSerialNumber3478A.Text, j, 1)))
                                Select Case j
                                    Case 5
                                        If Buffer < 65 Or Buffer > 90 Then
                                            MessageBox.Show(EditedCalibrationData(i) & " is not a valid serial number for an HP3478A (country code must be alphabetic")
                                            Exit Sub
                                        End If
                                    Case Else
                                        If Asc(Buffer) < 48 Or Asc(Buffer) > 57 Then
                                            MessageBox.Show(EditedCalibrationData(i) & " is not a valid serial number for an HP3478A (prefix and suffix must be numeric")
                                            Exit Sub
                                        End If
                                End Select
                            Next
                        Catch ex As Exception

                        End Try
                        Dim result1 As DialogResult = MessageBox.Show("Are you certain that you want to change the serial number?", "Confirm Action", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                        If Not result1 = DialogResult.Yes Then

                            'Abort the upload process and allow user to make new range selection, edit the calibration data, restart or leave program
                            ButtonCalramLoad3478A.Enabled = True
                            ButtonCalramLoad3478A.Text = "3478A Upload"
                            RollbackEdits()
                            Exit Sub
                        Else
                            'Disable and hide the Delete Serial Number button until needed
                            ButtonDeleteSerNr.Enabled = False
                            ButtonDeleteSerNr.Visible = False

                            'Display the new serial number in the serial number textbox
                            TextBoxSerialNumber3478A.Text = NewOffset(18) & NewGain(18)
                        End If
                    Else

                        'Display the serial number in the serial number textbox in condensed plain text format (the "space" character in the 6th position of the offset removed)
                        If Not EditedCalibrationData(i) = "@@@@@@@@@@@OO" Then
                            Buffer = EditedCalibrationData(i)
                            EditedCalibrationData(i) = Buffer
                            TextBoxSerialNumber3478A.Text = NewOffset(18) & NewGain(18)
                        End If
                    End If
                End If

                'Buffer and display the calibration range values
                Buffer = EditedCalibrationData(i)
                CalibrationRangeStore3478A(i) = EditedCalibrationData(i)
                DisplayCalibration(i, Mid(Buffer, 1, 6), NewOffset(i), Mid(Buffer, 7, 5), NewGain(i), Mid(Buffer, 12, 2), "valid")
                SetCalibrationRangeColor(i, Color.LightSeaGreen, 1, 13, 2)

                'Adjust the program flow options
                TextBoxEditGain.Enabled = False                                         'prevent altering gain value after a commit
                TextBoxEditOffset.Enabled = False                                       'prevent altering offset value after a commit
                ButtonRevertEdit3478A.Enabled = True                                    'provide the option to roll back all edits
                ButtonCommitEdit3478A.Enabled = False                                   'disallow furthre commits until further edits occur
                ButtonCalramLoad3478A.Enabled = True                                    'allow option to load an alternative binary calibration data file
                ButtonCalramLoad3478A.Text = "3478A Upload"                             'indicate that calibration data is ready for upload to the HP3478A
                Abort3478A = False                                                      'a commit obviously negates an abort
                EditOccurred = True                                                     'set the flag indicating that editing has occurred
            End If
        Next
    End Sub
    Private Sub ButtonEditCal3478A_Click(sender As Object, e As EventArgs) Handles ButtonEditCal3478A.Click

        'Filter out and highlight any invalid entries to the calibration range selection field
        If ComboBoxCalRange3478A.SelectedIndex < 0 Or ComboBoxCalRange3478A.SelectedIndex > 18 Then
            MessageBox.Show("Invalid calibration range selection")
            ComboBoxCalRange3478A.BackColor = Color.Yellow
            TextBoxEditOffset.Enabled = False
            TextBoxEditGain.Enabled = False
            Exit Sub
        End If

        'Filter for calibration ranges used by HP3478A
        Select Case ComboBoxCalRange3478A.SelectedIndex
            Case 5, 16, 18                                                              'calbration ranges unsued by the HP3478A

            Case Else
                'Insure the gain value contains 6 decimal digits
                If TextBoxEditGain.Text.Length < 8 Then
                    If Not SuppressMsgs Then                                            'check if suppression of messages is requested
                        MessageBox.Show("The gain value must contain 6 digits following the decimal point. Use 0 placeholders if necessary.")
                    End If
                    TextBoxEditGain.Focus()
                    Exit Sub
                End If
        End Select

        'Special handlinng for serial number editing
        If ComboBoxCalRange3478A.SelectedIndex = 18 Then
            If TextBoxEditOffset.Text = "0" Then

                'Zero offset represents an initial unused calibration range offset value
                TextBoxCalOffsetRaw.Text = "@@@@@@"
                TextBoxEditOffset.Text = "@@@@@@"
            Else

                'Convert the entered offset to raw offset values for display in the edit fields
                TextBoxCalOffsetRaw.Text = Nothing
                NewOffset(18) = Nothing

                'Append the hexadecimal representation of each of the first 4 offset values to the raw offset
                For i = 1 To 4
                    TextBoxCalOffsetRaw.Text = TextBoxCalOffsetRaw.Text & Convert.ToString(Chr(Asc(Mid(TextBoxEditOffset.Text, i, 1)) + 16))   '0="@", 1="A", 2="B", etc.
                    NewOffset(18) = NewOffset(18) & Mid(TextBoxEditOffset.Text, i, 1)                                                          'buffer the new offset value
                Next

                'The 5th positioin of the offset entry must be an alphabetic charachter 
                TextBoxCalOffsetRaw.Text = TextBoxCalOffsetRaw.Text & Mid(TextBoxEditOffset.Text, 5, 1)                       '(country code: "A" = manufactured in the U.S.)
                NewOffset(18) = NewOffset(18) & Mid(TextBoxEditOffset.Text, 5, 1)                                             'buffer the new offset value
            End If
            If TextBoxEditGain.Text = "1" Then

                'A gain of 1 represents an initial unused calibration range gain value
                TextBoxEditGain.Text = "@@@@@"
                TextBoxCalGainRaw.Text = "@@@@@"
            Else

                'Convert the entered plain text gain value to raw gain values for display in the edit raw display fields
                TextBoxCalGainRaw.Text = Nothing
                NewGain(18) = Nothing

                'Append the hexadecimal representation of each of the individual gain digits to the raw gain
                For i = 1 To 5
                    TextBoxCalGainRaw.Text = TextBoxCalGainRaw.Text & Convert.ToString(Chr(Asc(Mid(TextBoxEditGain.Text, i, 1)) + 16))          '0="@", 1="A", 2="B", etc.
                    NewGain(18) = NewGain(18) & Mid(TextBoxEditGain.Text, i, 1)                                                                 'buffer the new offset value
                Next
            End If
            If TextBoxCalOffsetRaw.Text = Mid(IterationRawOffset, 1, 5) And TextBoxCalGainRaw.Text = IterationRawGain And TextBoxNewChecksum.Text = IterationChecksum Then
                MessageBox.Show("The loaded calibration data has not changed.")
                SetCalibrationRangeColor(ComboBoxCalRange3478A.SelectedIndex, Color.Green, 1, 13, 2) 'reset the original highlighting
                EditedCalibrationData(ComboBoxCalRange3478A.SelectedIndex) = Nothing                 'clear any buffered editing for this calibration range
                Exit Sub
            End If
        End If


        'Check if edit fields differ from existing ram file data
        If TextBoxCalOffsetRaw.Text = IterationRawOffset And TextBoxCalGainRaw.Text = IterationRawGain And TextBoxNewChecksum.Text = IterationChecksum Then
            MessageBox.Show("The loaded calibration data has not changed.")
            SetCalibrationRangeColor(ComboBoxCalRange3478A.SelectedIndex, Color.Green, 1, 13, 2) 'reset the original highlighting
            EditedCalibrationData(ComboBoxCalRange3478A.SelectedIndex) = Nothing                 'clear any buffered editing for this calibration range
            Exit Sub
        End If

        'Highlight the calibration data which has been edited
        SetCalibrationRangeColor(ComboBoxCalRange3478A.SelectedIndex, Color.Orange, 1, 13, 2)

        'Buffer the edited calibration range data
        If ComboBoxCalRange3478A.SelectedIndex = 18 And TextBoxCalOffsetRaw.Text <> "@@@@@@" Then

            'Serial number offset needs a spaceholder in the 6th position since the "prefix" + "country code" total to 5 and the offset length must be 6
            EditedCalibrationData(ComboBoxCalRange3478A.SelectedIndex) = TextBoxCalOffsetRaw.Text & "@" & TextBoxCalGainRaw.Text & TextBoxNewChecksum.Text
        Else
            EditedCalibrationData(ComboBoxCalRange3478A.SelectedIndex) = TextBoxCalOffsetRaw.Text & TextBoxCalGainRaw.Text & TextBoxNewChecksum.Text
        End If

        'Enable the Commit Edit button
        ButtonCommitEdit3478A.Enabled = True


    End Sub
    'This subrouting processes the values entered into the calibration range offset and gain fields into binary storage format
    Private Function EditCalibration3478A(CalibrationRange As Integer, Offset As Integer, Gain As Double, RawOffset As String, RawGain As String, CheckSum As String)
        Dim Multiplier As Integer
        Dim PosVal As Integer
        Dim OffsetValue As Integer
        Dim temp As Double
        Dim RamValue As Char

        'Buffer the current calibration values
        IterationRawOffset = RawOffset
        IterationRawGain = RawGain
        If CheckSum <> Nothing Then
            IterationChecksum = CheckSum
        Else
            IterationChecksum = Mid(CalibrationRangeStore3478A(CalibrationRange), 12, 2)
        End If

        'Display the unconverted raw values in the appropriate TextBox
        TextBoxCalOffsetRaw.Text = RawOffset
        TextBoxCalGainRaw.Text = RawGain
        Gain = 1.0

        'Translate the entered plain text gain value to the format as stored in the unit's memory
        For cur = 1 To 5                                                                            'go through the 5 digit gain string one by one
            RamValue = Convert.ToChar(Mid(RawGain, cur, 1))                                         'convert the current gain digit value to char
            temp = Convert.ToInt32(RamValue) - 64                                                   'convert the gain digit char value to integer
            If (temp >= 8) Then                                                                     'adjust the value for negative gain
                temp = temp - 16
            End If
            Gain = Gain + (temp / (10 ^ (cur + 1)))                                                 'adjust the gain digit value by receding order of magnitude and accumulate total
        Next

        'Convert BCD offset to integer from the ram's offset sequence (nibbles 1 thru 6)
        OffsetValue = 0
        Multiplier = 100000
        For Position = 1 To RawOffset.Length                                                        'each successive digit represents an order of magnitude smaller value
            RamValue = Convert.ToChar(Mid(RawOffset, Position, 1))                                  'go through the 6 digit offset string one by one
            PosVal = Convert.ToInt32(RamValue) - 64                                                 'convert the byte value into an integer value
            OffsetValue = OffsetValue + (PosVal * Multiplier)                                       'accumulate the individual digit value to the total offset value
            Multiplier = Multiplier / 10                                                            'adjust the order of magnitude for the next digit
        Next

        'Adjust the offset total for negative offset values
        If OffsetValue > 999 Then
            OffsetValue = OffsetValue - 1000000
        End If
        Offset = Convert.ToString(OffsetValue)

        'Display the offset and gain values in the edit fields
        TextBoxEditOffset.Text = Convert.ToString(Offset)
        TextBoxEditGain.Text = Convert.ToString(Gain)

        'Enable edit mode for calibration offset and gain values
        TextBoxEditOffset.Enabled = True
        TextBoxEditGain.Enabled = True

        'Disable and hide Delete Serial Number button until needed
        ButtonDeleteSerNr.Enabled = False
        ButtonDeleteSerNr.Visible = False

    End Function
    Private Sub TextBoxEditGain_TextChanged(sender As Object, e As EventArgs) Handles TextBoxEditGain.TextChanged
        Dim KFactor As Double
        Dim GainDigit As Integer
        Dim GainDigitValue(6) As Integer
        Dim Carry As Integer
        Dim ReverseGain As String
        Dim GainString As String
        Dim KString As String
        Dim KInteger As Integer
        Dim KDouble As Double

        'Check if a valid gain value has been entered
        Select Case ComboBoxCalRange3478A.SelectedIndex
            Case 5, 16                                                                                   'unused calibration ranges
            Case 18                                                                                      'unused calibration range, but optional storge area for unit's serial number

                'Buffer the raw gain value
                GainString = TextBoxCalGainRaw.Text

                'Move the resulting new gain string value to the gain edit TextBox
                TextBoxCalGainRaw.Text = TextBoxEditGain.Text

                'Calculate new checksum for the edited calibration range
                If TextBoxCalGainRaw.Text.Length = 5 Then
                    ' TextBoxCalGainRaw.Text = TextBoxCalGainRaw.Text & " "
                    GenerateNewChecksum()
                End If
                Exit Sub
            Case Else

                'Allow entering up to 8 digits of gain value
                If TextBoxEditGain.Text.Length < 8 Then
                    Exit Sub
                End If
        End Select
        If Mid(TextBoxEditGain.Text, 1, 3) <> "1.0" And Mid(TextBoxEditGain.Text, 1, 3) <> "0.9" Then
            If Mid(TextBoxEditGain.Text, 1, 3) <> "1,0" And Mid(TextBoxEditGain.Text, 1, 3) <> "0,9" Then   'check for other culture settings, e.g. Germany
                Select Case ComboBoxCalRange3478A.SelectedIndex
                    Case 5, 16, 18
                        TextBoxEditGain.Text = "1"                                                          'unused calibration range initial gain value always 1
                        TextBoxCalGainRaw.Text = "@@@@@"                                                    'unused calibration range initial raw gain always '@@@@@'
                        Exit Sub
                    Case Else

                        'Guide the user as to valid gain values for an HP3478A
                        If Not SuppressMsgs Then                                                            'check if suppression of messages is requested
                            MessageBox.Show("The gain value must be between 0.911112 and 1.077777")
                        End If
                        Exit Sub
                End Select
            End If
        End If

        'Get the edited gain value from the edit TextBox
        If Mid(TextBoxEditGain.Text, 8, 1) = Nothing Then
            Select Case ComboBoxCalRange3478A.SelectedIndex
                Case 5, 16, 18                                                                               'unused calibration ranges
                Case Else
                    TextBoxEditGain.Text = TextBoxEditGain.Text & "0"                                        'add trailing 0 if necessary
            End Select
        End If

        'Get the edited gain value
        GainString = TextBoxEditGain.Text

        'Buffer the new gain value for later display
        NewGain(ComboBoxCalRange3478A.SelectedIndex) = GainString

        'Get the gain correction faktor for the edited gain value
        Try
            KFactor = Convert.ToDouble(Mid(TextBoxEditGain.Text, 2, 7))                 'KFactor is the decimal portion of the calibration gain value
        Catch ex As Exception
            'My rather inept solution for different keyboard/language settings
            If Mid(TextBoxEditGain.Text, 2, 1) = "," Then
                Mid(TextBoxEditGain.Text, 2, 1) = "."
            ElseIf Mid(TextBoxEditGain.Text, 2, 1) = "." Then
                Mid(TextBoxEditGain.Text, 2, 1) = ","
            End If
            KFactor = Convert.ToDouble(Mid(TextBoxEditGain.Text, 2, 7))
        End Try

        'Extract individual digit from correction faktors between 10^-2 and 10^-6
        For i = 2 To 6
            KDouble = KFactor / 10 ^ -i                                                 'get the appropriate order of magnitude value for the current iteration
            KInteger = Convert.ToInt32(KDouble)                                         'round the calculated magnitude to nearest integer value
            KString = Convert.ToString(KInteger)                                        'covert the rounded value to string format
            GainDigit = Convert.ToInt32(Strings.Right(KString, 1))                      'convert rightmost digit to gain digit for the current iteration

            'Compensate for "double-rounding"
            If Mid(GainString, i + 3, 1) = "5" And i < 5 Then                           'do not compensate rounding for the final digit
                If Convert.ToInt32(Mid(GainString, i + 4, 1) < 5) Then                  'if the digit following the current iteration is "5"...
                    GainDigit = GainDigit - 1                                           '...prevent it from being rounded again
                End If
            End If

            'First the non-rollover (no borrow or carry) digits
            If GainDigit > -1 And GainDigit < 6 Then                                    'digits with values from 0 to 5
                GainDigitValue(i) = GainDigit

            ElseIf GainDigit < 0 And GainDigit > -9 Then                                'digits with values from -1 to -8
                GainDigitValue(i) = GainDigit

                'Test for the rollover (borrow carry) digits
            ElseIf GainDigit > 5 Then                                                   'digits with values from 6 to 9
                GainDigitValue(i) = GainDigit

                For Carry = i To 2 Step -1                                              'recursive carry.
                    If GainDigitValue(Carry) > 5 Then
                        GainDigitValue(Carry) = GainDigitValue(Carry) - 10              'subtract 10 from the current digit
                    End If
                Next

            ElseIf GainDigit = -9 Then                                                  'digits with value -9
                GainDigitValue(i) = GainDigit

                For Carry = i To 2 Step -1                                              'recursive borrow
                    If GainDigitValue(Carry) < -8 Then
                        GainDigitValue(Carry) = GainDigitValue(Carry) + 10              'add 10 to the current digit
                        GainDigitValue(Carry - 1) = GainDigitValue(Carry - 1) - 1       'decrement left digit (borrow)
                    End If
                Next
            End If
        Next

        'Build compressed gain hex string from digits array
        For i = 2 To 6
            If GainDigitValue(i) < 0 Then
                ReverseGain = ReverseGain & Convert.ToChar(GainDigitValue(i) + 80)
            Else
                ReverseGain = ReverseGain & Convert.ToChar(GainDigitValue(i) + 64)
            End If
        Next

        'Move the resulting new gain string value to the gain edit TextBox
        TextBoxCalGainRaw.Text = ReverseGain

        'Calculate new checksum for the edited calibration range
        GenerateNewChecksum()

        'Enable the "Edit Calibration" button
        ButtonEditCal3478A.Enabled = True


    End Sub
    Private Sub TextBoxEditOffset_TextChanged(sender As Object, e As EventArgs) Handles TextBoxEditOffset.TextChanged
        Dim OffsetVal As Integer
        Dim OffsetRaw As String
        Dim RamValue As String
        Dim RamChar As Char
        Dim CharValue As Integer
        Dim MinusSign As Boolean
        Dim BufferText As String

        'Check if a valid offset value has been entered
        If TextBoxEditOffset.Text = Nothing Then
            Exit Sub

            'Check for need to process negative offset value
        ElseIf Mid(TextBoxEditOffset.Text, 1, 1) = "-" Then
            If Mid(TextBoxEditOffset.Text, 2, 1) = Nothing Then
                Exit Sub                                                                         'wait for entry of the next digit of the negative offset
            Else
                BufferText = Mid(TextBoxEditOffset.Text, 2, 5)                                   'buffer the digits following the "minus" sign (-)
                Try
                    OffsetVal = 0 - Convert.ToInt32(BufferText)                                  'buffer the negaive offset value
                    MinusSign = True                                                             'flag the offset entry as a negative value
                Catch

                    'Allow only numberic entries (except for the 19th calibration range, which may contain an alphanumeric serial number
                    If ComboBoxCalRange3478A.SelectedIndex <> 18 Then
                        MessageBox.Show("The offset value must contain only numeric values.")
                        Exit Sub
                    End If
                End Try
            End If
        ElseIf ComboBoxCalRange3478A.SelectedIndex = 18 Then

            'Serial number processing
            TextBoxCalOffsetRaw.Text = TextBoxEditOffset.Text
            If TextBoxCalOffsetRaw.Text.Length = 5 Then                                           'a serial number must occupy the first 5 positions of the offset
                ButtonEditCal3478A.Enabled = True                                                 'allow processing the offset when the first 5 positions have been entered
                GenerateNewChecksum()
            ElseIf TextBoxEditOffset.Text = "0" Then
                ButtonEditCal3478A.Enabled = True                                                 'allow deleting a serial number (offset value is "0" for an unused calibration range)
            End If
            Exit Sub
        Else
            Try
                OffsetVal = Convert.ToInt32(TextBoxEditOffset.Text)                               'buffer the numeric plain text offset value
            Catch
                MessageBox.Show("The offset value must contain only numeric values.")
                Exit Sub
            End Try
        End If

        If OffsetVal < -100000 And OffsetVal > 899999 Then
            Select Case ComboBoxCalRange3478A.SelectedIndex
                Case 5, 16                                                                          'do not apply normal range limits to unused ranges
                Case 18
                    TextBoxCalOffsetRaw.Text = TextBoxEditOffset.Text                               'buffer the "prefix" and "country code" part of the serial number (offset 1 thru 5)
                    Exit Sub
                Case Else

                    'Guide the user to "legal" range for HP3478A calibration offset values
                    MessageBox.Show("The offset value must be between from -100000 to 899999.")
                    Exit Sub
            End Select
        End If
        If Not TextBoxEditOffset.Text = Nothing Then
            If MinusSign = False Then

                'Generate raw offset character string from the entered positive offset value
                For i = 1 To TextBoxEditOffset.TextLength
                    CharValue = Convert.ToInt32(Mid(TextBoxEditOffset.Text, i, 1)) + 64
                    RamChar = Convert.ToChar(CharValue)
                    RamValue = RamValue & RamChar
                Next

                'Pad the raw offst string with traiing zeros (@)
                For i = 1 To (6 - TextBoxEditOffset.TextLength)
                    RamValue = "@" & RamValue
                Next

                'Display the new values into the appropriate TextBoxes
                TextBoxEditOffset.Text = OffsetVal
                TextBoxCalOffsetRaw.Text = RamValue
            Else

                'Generate raw offset character string from the entered negative offset value
                OffsetVal = OffsetVal + 1000000
                BufferText = Convert.ToString(OffsetVal)
                For i = 1 To 6
                    CharValue = Convert.ToInt32(Mid(BufferText, i, 1)) + 64
                    RamChar = Convert.ToChar(CharValue)
                    RamValue = RamValue & RamChar
                Next

                'Display the new values into the appropriate TextBoxes
                If ComboBoxCalRange3478A.SelectedIndex = 18 Then
                    RamValue = Mid(RamValue, 2, 6)
                End If
                TextBoxCalOffsetRaw.Text = RamValue
                OffsetVal = OffsetVal - 1000000

            End If

            'Calculate new checksum for the edited calibration range
            GenerateNewChecksum()

            'Enable the "Edit Calibration" button
            ButtonEditCal3478A.Enabled = True
        End If

        'Buffer the new offset value for later display
        NewOffset(ComboBoxCalRange3478A.SelectedIndex) = Convert.ToString(OffsetVal)

    End Sub
    Private Sub GenerateNewChecksum()        'Calculate new checksum for the edited calibration range

        Dim CalibrationString As String
        Dim Range As Integer

        'Initialze raw offset textbox
        If TextBoxCalOffsetRaw.Text = "0" Then
            TextBoxCalOffsetRaw.Text = "@@@@@@"
        End If

        'Setup the calibration string and range variables
        CalibrationString = TextBoxCalOffsetRaw.Text & TextBoxCalGainRaw.Text
        If ComboBoxCalRange3478A.SelectedIndex < 0 Then
            Range = 18
        Else
            Range = ComboBoxCalRange3478A.SelectedIndex
        End If

        'Special handling for serial number
        If Range = 18 And CalibrationString.Length = 10 Then
            CalibrationString = Mid(CalibrationString, 1, 5) & "@" & Mid(CalibrationString, 6, 5)
        End If

        'Generate the appropriate checksum for the given offset and gain values
        For CalibrationRange = Range + 1 To Range + 1
            Dim ValueSum As Integer
            Dim Address As Integer
            Dim RamValue As Char
            Dim CheckSum As Byte
            Dim RawOffset As String
            Dim RawGain As String
            Dim ChecksumChar As Char
            Dim checksumchar2 As Char
            Dim NewChecksum As String
            Dim checksumstr As String
            Dim checksumstr2 As String
            Dim checksumint As Integer
            Dim checksumint2 As Integer
            Dim AsciiValue2 As Integer

            'Check if a valid gain value has been entered
            If Mid(TextBoxEditGain.Text, 1, 3) <> "1.0" And Mid(TextBoxEditGain.Text, 1, 3) <> "0.9" Then
                If Mid(TextBoxEditGain.Text, 1, 3) <> "1,0" And Mid(TextBoxEditGain.Text, 1, 3) <> "0,9" Then 'check for other culture settings, e.g. Germany
                    Select Case CalibrationRange
                        Case 6, 17, 19
                            'Not used calibration ranges need not be within the normal values
                        Case Else

                            'Inform user of valid gain values for an HP3478A
                            If Not SuppressMsgs Then                                      'check if suppression of messages is requested
                                MessageBox.Show("The gain value must be between 0.911112 and 1.077777")
                            End If
                            Exit Sub
                    End Select
                End If
            End If

            'Calibration values are organized in blocks of 11 nibbles + 2 checksum nibbles starting at 2nd ram address (address 01)
            Address = 1

            'Accumulate raw offset, raw gain and checksum values
            ValueSum = 0
            For CheckSumValues = 1 To 11                                            'offset and gain values for which a valid checksum is to be generated
                Try
                    RamValue = Convert.ToChar(Mid(CalibrationString, Address, 1))
                Catch
                    If Mid(CalibrationString, CheckSumValues, 1) = Nothing Then
                        RamValue = ""                                               'allow for the "missing" 6th offset position value for a serial number offset
                    End If
                End Try
                ValueSum = ValueSum + Convert.ToInt32(RamValue) - 64                'accumulate checksum from binary data ("@"=0, "A"=1, "B"=2 etc.)
                If CheckSumValues < 7 Then                                          'first 6 nibbles of ram data represent the raw offset
                    RawOffset = RawOffset & RamValue
                ElseIf CheckSumValues < 12 Then                                     'next 5 nibbles of ram data represent the raw gain
                    RawGain = RawGain & RamValue
                End If
                Address = Address + 1
            Next

            'Generate the checksum for edited calibration values
            Select Case Range
                Case 5, 16                                                          'do not recalculate checksum for unused calibration ranges
                    TextBoxNewChecksum.Text = "OO"                                  'checksum for unused calibration ranges should always be "OO" (0xFF)
                Case Else
                    If ValueSum < 0 Then
                        ValueSum = ValueSum + 144
                    End If
                    CheckSum = Convert.ToByte(255 - ValueSum)                       'calculate the new checksum value
                    NewChecksum = Hex(CheckSum)                                     'checksum must be split into 2 separate values for calibration ram
                    ChecksumChar = Convert.ToChar(Mid(NewChecksum, 1, 1))           'split the hex value of the checksum into 2 separte characters
                    checksumchar2 = Convert.ToChar(Mid(NewChecksum, 2, 1))          'rightmost hex value to 2nd character
                    checksumstr = Convert.ToString(ChecksumChar)                    'convert first checksum character to string
                    checksumstr2 = Convert.ToString(checksumchar2)                  'convert second checksum character to string
                    checksumint = Asc(checksumstr) + 9                              'generate ascii value for the first checksum character string
                    checksumint2 = Asc(checksumstr2) + 9                            'generate ascii value for the sectond checksum character string
                    AsciiValue2 = Asc(checksumstr2)                                 'move the ascii value of the second character to a testing variable

                    'Second ascii value must be corrected for characters 0 thru 9
                    If AsciiValue2 > 47 And AsciiValue2 < 58 Then                   'check if the current character is a digit within the range of 0 to 9
                        checksumint2 = Asc(checksumstr2) + 16                       'correct the ascii value 
                    End If
                    checksumstr = Chr(checksumint) & Chr(checksumint2)              'generate the string for the calibration ram from the two split checksum characters
                    TextBoxNewChecksum.Text = checksumstr                           'display the new checksum in the edit TextBox
            End Select
        Next

        'Buffer the original calibration range offset and gain data prior to editing
        If CalibrationOffsetBackup(Range) = Nothing Then
            CalibrationOffsetBackup(Range) = TextBoxEditOffset.Text
            CalibrationGainBackup(Range) = TextBoxEditGain.Text
        End If


    End Sub
    Private Sub ComboBoxCalRange3478A_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxCalRange3478A.SelectedIndexChanged
        Dim Offset As String
        Dim Gain As String
        Dim OffsetGain As String

        'Filter out any invalid entry into the selection box
        If ComboBoxCalRange3478A.SelectedIndex < 0 Or ComboBoxCalRange3478A.SelectedIndex > 18 Then
            Exit Sub
        End If

        'Insure calibration data buffers are available for possible editing rollback
        If Not QuickAndDirty Then
            For i = 0 To 18
                If CalibrationRangeBackup3478A(i) = Nothing Then
                    CalibrationRangeBackup3478A(i) = CalibrationRangeStore3478A(i)
                End If
            Next
        End If

        'Reset the standard background color (may have been changed by a highlighted erroneous entry)
        ComboBoxCalRange3478A.BackColor = Color.WhiteSmoke

        'Check for QUICK-AND-DIRTY processing flag
        If Caller3478A = Nothing Then
            QuickAndDirty = True
        End If

        'Display the calibration ram data for the calibration range selected for editing
        Select Case (ComboBoxCalRange3478A.SelectedIndex)
            Case 0
                Offset = Label01RawOffset.Text
                Gain = Label01RawGain.Text
            Case 1
                Offset = Label02RawOffset.Text
                Gain = Label02RawGain.Text
            Case 2
                Offset = Label03RawOffset.Text
                Gain = Label03RawGain.Text
            Case 3
                Offset = Label04RawOffset.Text
                Gain = Label04RawGain.Text
            Case 4
                Offset = Label05RawOffset.Text
                Gain = Label05RawGain.Text
            Case 5
                Offset = Label06RawOffset.Text
                Gain = Label06RawGain.Text
            Case 6
                Offset = Label07RawOffset.Text
                Gain = Label07RawGain.Text
            Case 7
                Offset = Label08RawOffset.Text
                Gain = Label08RawGain.Text
            Case 8
                Offset = Label09RawOffset.Text
                Gain = Label09RawGain.Text
            Case 9
                Offset = Label10RawOffset.Text
                Gain = Label10RawGain.Text
            Case 10
                Offset = Label11RawOffset.Text
                Gain = Label11RawGain.Text
            Case 11
                Offset = Label12RawOffset.Text
                Gain = Label12RawGain.Text
            Case 12
                Offset = Label13RawOffset.Text
                Gain = Label13RawGain.Text
            Case 13
                Offset = Label14RawOffset.Text
                Gain = Label14RawGain.Text
            Case 14
                Offset = Label15RawOffset.Text
                Gain = Label15RawGain.Text
            Case 15
                Offset = Label16RawOffset.Text
                Gain = Label16RawGain.Text
            Case 16
                Offset = Label17RawOffset.Text
                Gain = Label17RawGain.Text
            Case 17
                Offset = Label18RawOffset.Text
                Gain = Label18RawGain.Text
            Case 18
                Offset = Label19RawOffset.Text
                Gain = Label19RawGain.Text
        End Select

        If QuickAndDirty Then

            'Read the current calibration data directly from the HP3478A since no buffer data has been generated through a backup or restore process
            OffsetGain = GetCalibrationRange(ComboBoxCalRange3478A.SelectedIndex)

            'Check for communication failure while reading from the HP3478A calibration range memory
            If OffsetGain <> "Error" Then
                Offset = Mid(OffsetGain, 1, 6)
                Gain = Mid(OffsetGain, 7, 5)
                TextBoxCalOffsetRaw.Text = Offset
                TextBoxCalGainRaw.Text = Gain
            Else
                MessageBox.Show("Error reading calibration range. Please check device connectrions.")
                Exit Sub
            End If
        End If

        'Display the edited calibration range data
        EditCalibration3478A(ComboBoxCalRange3478A.SelectedIndex, 0, 0, Offset, Gain, Nothing)

        'Enable the calibration edit button
        ButtonEditCal3478A.Enabled = True

        'Place the cursor in the offset edit textbox
        TextBoxEditOffset.Focus()

        If ComboBoxCalRange3478A.SelectedIndex = 18 And TextBoxSerialNumber3478A.Text <> Nothing Then

            'Offer the user the option to delete the serial number
            ButtonDeleteSerNr.Visible = True
            ButtonDeleteSerNr.Enabled = True

        Else

            'Hide the serial number delete button until needed
            ButtonDeleteSerNr.Visible = False
            ButtonDeleteSerNr.Enabled = False
        End If

        'Prevent editing unused calibration ranges
        Select Case (ComboBoxCalRange3478A.SelectedIndex)
            Case 5, 16
                TextBoxEditGain.Enabled = False
                TextBoxEditOffset.Enabled = False
                ButtonEditCal3478A.Enabled = False
        End Select


    End Sub

    Private Sub CheckCalEnableSwitch(Reconfirming As Boolean)
        Dim CalAddrByte As Byte
        Dim LoadData(4) As Byte
        Dim LoadCommand As String
        Dim ReadBackData(1) As Byte

        'Check to see if the calibration enable switch is in the calibrate position
        dev1.enablepoll = False
        CalAddrByte = Convert.ToByte(0)                                         'convert the integer address value to byte
        LoadData(0) = Convert.ToByte(88)                                        'convert ascii "X" into byte array
        LoadData(1) = CalAddrByte                                               'enter calibration ram address into byte array
        LoadData(2) = Convert.ToByte(Asc("O"))                                  'enter test byte into byte array
        LoadData(3) = Nothing                                                   'initialize Prologix handling extra array byte
        LoadData(4) = Nothing                                                   'initialize Prologix handling extra array byte
        LoadCommand = System.Text.Encoding.Default.GetString(LoadData)          'setup the load command
        dev1.SendBlocking(LoadCommand, False)                                   'perform upload to the HP3478A calibration ram address 0
        System.Threading.Thread.Sleep(100)                                      '100mS delay to allow device time to process command

        'Read back the current memory address and verify the written is correct
        ReadBackData(0) = Convert.ToByte(87)                                    'convert ascii "W" into byte array to enable read back
        ReadBackData(1) = CalAddrByte                                           'enter calibration ram address into read back byte array
        dev1.enablepoll = False
        ReadAddrByte = Convert.ToByte(0)                                         'convert calibration ram address 0 to byte
        ReadCommandArray(0) = Convert.ToByte(87)                                 'convert ascii "W" into read command byte array
        ReadCommandArray(1) = ReadAddrByte                                       'enter calibration ram address 0 into read command byte array
        ReadCommandArray(2) = Nothing                                            'Third byte needed only for Prologix adapter escape sequence characters - not relevant by address 0
        ReadCommand = System.Text.Encoding.Default.GetString(ReadCommandArray)   'generate the command to read calibration ram data from the unit
        ReadResult = SendQuery(ReadCommand)                                      'read the first address in the calibration ram (address 0)
        If Not q.status = 0 Then                                                 'check if the connection to the HP3478A is working
            Exit Sub
        End If
        System.Threading.Thread.Sleep(100)                                       '100mS delay to allow device time to process command
        If GetResponse() <> "?" Then
            CalramValue3478A = q.ResponseAsString
        Else
            MessageBox.Show("Device 1 does not appear to be responding. Please check the device connections.")
            Exit Sub
        End If
        If CalramValue3478A <> "O" And CalramValue3478A <> "@" Then              'if not 'O' or "@", then the switch is turned off (horizontal position)
            CalibrationSwitchDisabled = True                                     'set the calibration switch flag to disabled
            If Not Reconfirming Then
                'Display message to allow user to enable the CAL ENABLE switch position 
                MessageBox.Show(Me, "The CAL ENABLE switch appears to be off! You must activate the CAL ENABLE switch now or only a test write (compare) will be performed.", MessageBoxButtons.OK)
                Button3478Aabort.Enabled = True                                  'allow aborting the process
                ButtonRestartTab3478A.Enabled = True                             'allow restarting after abort
            End If
            Exit Sub
        Else
            'First match could have been coincidence, so repeat the write/read test with changed data
            LoadData(2) = Convert.ToByte(Asc("L"))                                'enter 2nd test byte into byte array
            LoadCommand = System.Text.Encoding.Default.GetString(LoadData)        'setup the load command
            dev1.SendBlocking(LoadCommand, False)                                 'perform load to ram operation
            System.Threading.Thread.Sleep(100)                                    '100mS delay to allow device time to process command
            ReadResult = SendQuery(ReadCommand)                                   'read the first address in the calibration ram
            If Not q.status = 0 Then                                              'check if the connection to the HP3478A is working
                Exit Sub
            End If
            If GetResponse() <> "?" Then
                CalramValue3478A = q.ResponseAsString
            Else
                MessageBox.Show("Device 1 does not appear to be responding. Please check the device connections.")
                Exit Sub
            End If
            If CalramValue3478A <> "L" And CalramValue3478A <> "C" Then           'if not 'L' or "C", then the switch is turned off (horizontal position)
                CalibrationSwitchDisabled = True                                  'set the calibration switch flag to disabled
                'Display message to allow user to enable the CAL ENABLE switch position 
                If Not Reconfirming Then
                    MessageBox.Show(Me, "The CAL ENABLE switch appears to be off! You must activate the CAL ENABLE switch now or only a test write (compare) will be performed.", MessageBoxButtons.OK)
                    Button3478Aabort.Enabled = True                               'allow aborting the process
                    ButtonRestartTab3478A.Enabled = True                          'allow restarting after abort
                End If
                Exit Sub
            End If
        End If
        If Not SuppressMsgs Then                                                  'check if suppression of messages is requested
            MessageBox.Show("The unit's calibration ram data will now be overwritten!")
        End If
        CalibrationSwitchDisabled = False

        'Prevent other processes from being initiated during the upload process
        TextBox3478AFrom.Enabled = False
        TextBox3478ATo.Enabled = False
        TextBoxEditGain.Enabled = False
        TextBoxEditOffset.Enabled = False
        TextBoxCalramFile3478A.Enabled = False
        TextBoxCalRamLoadFile3478A.Enabled = False
        ButtonCalramDump3478A.Enabled = False
        ButtonRestartTab3478A.Enabled = False
        ButtonRevertEdit3478A.Enabled = False
        ButtonEditCal3478A.Enabled = False
        ComboBoxCalRange3478A.Enabled = False

    End Sub
    Private Sub ButtonRevertEdit3478A_Click(sender As Object, e As EventArgs) Handles ButtonRevertEdit3478A.Click
        'Roll back any calibration range editing changes
        RollbackEdits()
    End Sub

    Private Sub RollbackEdits()
        Dim Buffer As String
        Dim UpdateColor As Color

        'Roll back any calibration range editing changes
        TextBoxSerialNumber3478A.Text = SerialNumberBackup
        For i = 0 To 18
            If CalibrationRangeStore3478A(i) <> CalibrationRangeBackup3478A(i) Or QuickAndDirty Or EditCount = 0 Then
                If CalibrationRangeBackup3478A(i) = Nothing Or CalibrationRangeBackup3478A(i) = "@@@@@@@@@@@@@" Then
                    CalibrationRangeBackup3478A(i) = "@@@@@@@@@@@@@"
                    CalibrationGainBackup(i) = "0.000000"
                    CalibrationOffsetBackup(i) = "0000000"
                    CalibrationValidBackup(i) = "invalid"
                    UpdateColor = Color.Black
                Else
                    Select Case i
                        Case 5, 16, 18
                            UpdateColor = Color.Blue
                        Case Else
                            UpdateColor = Color.Green
                    End Select
                End If
                NewGain(i) = Nothing
                NewOffset(i) = Nothing
                NewCheckSum3478A(i) = Nothing
                EditedCalibrationData(i) = Nothing
                CalibrationRangeStore3478A(i) = CalibrationRangeBackup3478A(i)
                Buffer = CalibrationRangeStore3478A(i)
                SetCalibrationRangeColor(i, UpdateColor, 0, 0, 2)
                DisplayCalibration(i, Mid(Buffer, 1, 6), CalibrationOffsetBackup(i), Mid(Buffer, 7, 5), CalibrationGainBackup(i), Mid(Buffer, 12, 2), CalibrationValidBackup(i))
            End If
        Next

        'Initialize editing fields
        Abort3478A = False
        EditOccurred = False
        EditCount = 0
        QuickAndDirty = False
        ComboBoxCalRange3478A.Text = Nothing
        ComboBoxCalRange3478A.Enabled = True
        TextBoxEditGain.Text = Nothing
        TextBoxEditOffset.Text = Nothing
        TextBoxCalOffsetRaw.Text = Nothing
        TextBoxCalGainRaw.Text = Nothing
        TextBoxNewChecksum.Text = Nothing
        TextBoxEditGain.Enabled = False
        TextBoxEditOffset.Enabled = False
        ButtonCommitEdit3478A.Enabled = False
        ButtonRevertEdit3478A.Enabled = False
        ButtonCalramLoad3478A.Enabled = False
        ButtonCalramLoad3478A.Text = "3478A Restore"
        ButtonEditCal3478A.Enabled = False
        ButtonDeleteSerNr.Enabled = False
        ButtonDeleteSerNr.Visible = False
        Button3478Aabort.Enabled = False

    End Sub

    Private Sub ButtonRestartTab3478A_Click(sender As Object, e As EventArgs) Handles ButtonRestartTab3478A.Click
        'Initialize 3478A calibration datafields
        ClearTab3478A()
        ButtonCalramDump3478A.Enabled = True
    End Sub

    Private Sub ClearTab3478A()
        'Perform the initialization of 3478A calibration datafields
        RAMfilename3478A = ""
        Abort3478A = False
        EditOccurred = False
        QuickAndDirty = False
        CalibrationSwitchDisabled = False
        EditCount = 0
        CalAddrStart3478A = 0
        CalAddrEnd3478A = 255
        Offset = Nothing
        GainValue = Nothing
        IterationRawGain = Nothing
        IterationRawOffset = Nothing
        IterationChecksum = Nothing
        WriteSuccess3478A = False
        FileContainsSerNr = False
        TextBoxSerialNumber3478A.Text = Nothing
        TextBoxSerialNumber3478A.Enabled = False
        TextBoxSerialNumber3478A.BackColor = Color.WhiteSmoke
        TextBoxEditGain.Text = Nothing
        TextBoxEditOffset.Text = Nothing
        TextBoxEditGain.Enabled = False
        TextBoxEditOffset.Enabled = False
        TextBoxNewChecksum.Text = Nothing
        TextBoxCalGainRaw.Text = Nothing
        TextBoxCalOffsetRaw.Text = Nothing
        TextBoxCalramFile3478A.Text = Nothing
        TextBoxCalRamLoadFile3478A.Text = Nothing
        ComboBoxCalRange3478A.Text = Nothing
        ComboBoxCalRange3478A.Enabled = True
        Caller3478A = Nothing
        TextBox3478AFrom.Text = 0
        TextBox3478ATo.Text = 255
        TextBox3478AFrom.Enabled = True
        TextBox3478ATo.Enabled = True
        ButtonCalramLoad3478A.Enabled = True
        ButtonCalramDump3478A.Enabled = True
        ButtonDeleteSerNr.Enabled = False
        ButtonDeleteSerNr.Visible = False
        ButtonEditCal3478A.Enabled = False
        ButtonRevertEdit3478A.Enabled = False
        ButtonCalramLoad3478A.Text = "3478A Restore"
        Button3478Aabort.Enabled = False
        ButtonRestartTab3478A.Enabled = False
        LabelCalRange19.Text = "Not used"

        For i = 0 To 255
            CalramStore3478A(i) = Nothing
            DataReadFromRam(i) = 0
            DataToLoad(i) = 0
        Next
        For i = 0 To 18
            CalibrationRangeStore3478A(i) = "@@@@@@@@@@@@@"
            CalibrationRangeBackup3478A(i) = Nothing
            CalibrationOffsetBackup(i) = Nothing
            CalibrationGainBackup(1) = Nothing
            CalibrationValidBackup(i) = Nothing
            EditedCalibrationData(i) = Nothing
            NewGain(i) = Nothing
            NewOffset(i) = Nothing
            NewCheckSum3478A(i) = Nothing
            DisplayCalibration(i, "@@@@@@", "000000", "@@@@@", "0.000000", "@@", "invalid")
            SetCalibrationRangeColor(i, Color.Black, 0, 0, 2)
        Next
        For i = 0 To 18
            CalibrationRangeStore3478A(i) = Nothing
        Next

    End Sub
    Private Sub CheckSerialNumber()
        Dim Buffer As String

        CalramStatus3478A.Text = "CHECKING SERIAL NUMBER"
        Me.Refresh()

        'Check if the device connection has changed
        If Not DeviceID = Nothing Then
            If Not DeviceID = EventlogEntries(EventlogEntries.Count - 1) Then

                'Buffer the calling subroutine
                Buffer = Caller3478A

                'Clear all data from previous device connection
                ClearTab3478A()

                'Reinstate the calling subroutine
                Caller3478A = Buffer

                'Re-assign the current device ID
                DeviceID = EventlogEntries(EventlogEntries.Count - 1)
            End If
        Else

            'Assign the current connection as the device ID
            DeviceID = EventlogEntries(EventlogEntries.Count - 1)
        End If

        'Associate the calibration ram data with the serial number of the HP3478A
        If TextBoxSerialNumber3478A.Text = Nothing And Not EditOccurred Then
            Dim SerNr As String

            'Prepare the command to read serial number from unit's calibration ram data stored in the 19th (otherwise unused) calibration range
            ReadCommandArray(0) = Convert.ToByte(87)                                     'convert ascii "W" into read command byte array
            ReadCommandArray(2) = Nothing                                                'last byte only required when a Prologix adapter is connected
            dev1.enablepoll = False
            dev1.showmessages = False
            For i = 235 To 245                                                           'serial number is assumed to be stored in the first ten byte of calibration range 19
                ReadAddrByte = Convert.ToByte(i)                                         'convert calibration ram address to byte
                ReadCommandArray(1) = ReadAddrByte                                       'enter calibration ram address into read command byte array
                ReadCommand = System.Text.Encoding.Default.GetString(ReadCommandArray)   'generate the command to read calibration ram data from the unit
                ReadResult = SendQuery(ReadCommand)                                      'send the read command with address and wait for reply
                If Not q.status = 0 Then                                                 'check if the connection to the HP3478A is working
                    Exit Sub
                End If
                If i = 240 Then
                    'Skip over address 240 - the 6th position of the offset data is not part of the stored serial number
                ElseIf i <> 239 Or SerNr = "0000" Then
                    SerNr = SerNr & Asc(GetResponse()) - 64                              'serial number prefix (digits 1 thru 4) and serial number suffix (digits 6 thru 10)
                Else
                    SerNr = SerNr & GetResponse()                                        'fifth poition of serial number is country character separating prefix from suffix
                End If
            Next

            If Not SerNr = "0000000000" Then                                             'the unit's serial number has been entered

                'Display the serial number in the serial number textbox
                TextBoxSerialNumber3478A.Text = SerNr
                LabelCalRange19.Text = "HP serial number"
            Else
                Dim result As DialogResult = MessageBox.Show("Continue without serial number?", "To improve safety against mishaps, enter the unit's serial number.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

                'Check for serial number match
                If result = DialogResult.Yes Then

                    'Set standard background color to the serial number textbox
                    TextBoxSerialNumber3478A.BackColor = DefaultBackColor                'continue without serial number
                Else

                    'Adjust the program flow options
                    If Caller3478A = "Read" Then

                        'Allow the user to upload the serial number to the unit's memory
                        ButtonCalramLoad3478A.Enabled = True
                        EditOccurred = True
                    ElseIf Caller3478A = "Write" Then

                        'Prevent unnecessarily uploading the identical data again
                        ButtonCalramDump3478A.Enabled = False
                        EditOccurred = True
                    End If
                    Me.Refresh()

                    'Allow the user enter the unit's serial number
                    TextBoxSerialNumber3478A.Enabled = True                     'Allow serial number textbox to be edited
                    TextBoxSerialNumber3478A.BackColor = Color.LightYellow      'Highlight the serial number textbox
                    TextBoxSerialNumber3478A.Focus()                            'Place the cursor into the serial number textbox for user to provide the unit's serial number

                    'Adjust program flow options
                    TextBox3478AFrom.Enabled = False                            'prevent changing the ram start address
                    TextBox3478ATo.Enabled = False                              'prevent changing the ram end address
                    TextBoxCalramFile3478A.Enabled = False                      'prevent editing download file path and name
                    TextBoxCalRamLoadFile3478A.Enabled = False                  'prevent editing upload file path and name
                    ComboBoxCalRange3478A.Enabled = False                       'prevent editing the calibration data
                    ButtonRestartTab3478A.Enabled = True                        'allow user to restart if desired
                    If Caller3478A = "Read" Then
                        ButtonCalramLoad3478A.Enabled = False                   'prevent the upload process from being started
                    ElseIf Caller3478A = "Write" Then
                        ButtonCalramDump3478A.Enabled = False                   'prevent the download process from being started
                    End If
                    Exit Sub
                End If
            End If
        ElseIf TextBoxSerialNumber3478A.Text <> Nothing And TextBoxSerialNumber3478A.Text.Length < 10 Then
            Abort3478A = True
        End If

    End Sub
    Private Sub ButtonDeleteSerNr_Click(sender As Object, e As EventArgs) Handles ButtonDeleteSerNr.Click

        'Insure that all unedited calibration ranges have been buffered
        VerifyCalibrationFileData()

        'Set initialization values for unused calibration range 19 (delete serial number)
        TextBoxEditGain.Text = 1
        TextBoxEditOffset.Text = 0
        TextBoxCalGainRaw.Text = "@@@@@"
        TextBoxCalOffsetRaw.Text = "@@@@@@"
        ButtonCalramDump3478A.Enabled = False
    End Sub
    Private Sub TextBoxSerialNumber3478A_TextChanged(sender As Object, e As EventArgs) Handles TextBoxSerialNumber3478A.TextChanged
        Dim Buffer As String
        Dim Value As Integer

        'Check serial number validity while it is being entered and display invalid results
        If TextBoxSerialNumber3478A.Text = Nothing Then
            Exit Sub
        End If

        Buffer = TextBoxSerialNumber3478A.Text
        Value = Asc(Mid(Buffer, Buffer.Length, 1))

        If Buffer.Length > 10 Then
            MessageBox.Show("The length of the serial number must be 10 (a 4-digit prefix followed by an alphabetic country code and a 5-digit suffix")
            TextBoxSerialNumber3478A.Text = Mid(Buffer, 1, Buffer.Length - 1)
            Exit Sub
        End If
        Select Case Buffer.Length
            Case 5
                If Value < 65 Or Value > 90 Then
                    MessageBox.Show("The 5th position in the serial must be an alphabetic character (country code)")
                    TextBoxSerialNumber3478A.Text = Mid(Buffer, 1, Buffer.Length - 1)
                    Exit Sub
                End If
            Case 10
                If Caller3478A = "Read" Then
                    ButtonCalramDump3478A.Focus()
                ElseIf Caller3478A = "Write" Then
                    ButtonCalramLoad3478A.Focus()
                End If
            Case Else
                If Value < 48 Or Value > 57 Then
                    MessageBox.Show("The serial must be all numeric, except for the 5th position, which must be an alphabetic character (country code)")
                    TextBoxSerialNumber3478A.Text = Mid(Buffer, 1, Buffer.Length - 1)
                    Exit Sub
                End If
        End Select
    End Sub
    Private Sub ButtonUG3478A_Click(sender As Object, e As EventArgs) Handles ButtonUG3478A.Click

        'Setup the path and file for the 3478A Tab user guide
        Dim filePath As String = System.IO.Path.Combine(CSVfilepath.Text, "3478A_UserGuide.pdf")

        ' Check if the user guide file exists
        If Not System.IO.File.Exists(filePath) Then
            MessageBox.Show("File not found: " & filePath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Attempt to open the user guide
        Try
            Process.Start(filePath)
        Catch ex As Exception
            MessageBox.Show("An error occurred while trying to open the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetResponse() As String

        'Check if a query was success 
        If q.status = 0 Then

            'Return the value of the response to the query
            Return q.ResponseAsString
        Else

            'Report a connection problem after unsuccessful query
            MessageBox.Show("Device 1 is not responding. Please check the device connections and start over")
            IODevice.DisposeAll()
            IODevice.ShowDevices()
            If Not RAMfilename3478A = Nothing Then

                'Rename the binary output file to indicate it is incomplete since the download process was interrupted during processing
                Try
                    fs.Close()
                Catch ex As Exception
                End Try
                TextBoxCalramFile3478A.Text = Mid(RAMfilename3478A, 1, RAMfilename3478A.Length - 4) & "_defect.bin"
                System.IO.File.Move(RAMfilename3478A, TextBoxCalramFile3478A.Text)
            End If
            ClearTab3478A()
            Return "?"
        End If
    End Function
    Private Function SendQuery(ReadCommand As String) As Byte
        Try
            ReadResult = dev1.QueryBlocking(ReadCommand, q, False)                   'send the read command with address and wait for reply
        Catch ex As Exception

            'Report a connection problem after unsuccessful query
            IODevice.DisposeAll()
            MessageBox.Show("Device 1 is not responding. Please check the device connections and start over.")
            IODevice.ShowDevices()
            If Not RAMfilename3478A = Nothing Then

                'Rename the binary output file to indicate it is incomplete since the download process was interrupted during processing
                Try
                    fs.Close()
                Catch exception As Exception
                End Try
                TextBoxCalramFile3478A.Text = Mid(RAMfilename3478A, 1, RAMfilename3478A.Length - 4) & "_defect.bin"
                System.IO.File.Move(RAMfilename3478A, TextBoxCalramFile3478A.Text)
            End If
            ClearTab3478A()
            Exit Function
        End Try
        If Not q.status = 0 Then

            'Report a connection problem after unsuccessful query
            IODevice.DisposeAll()
            MessageBox.Show("Device 1 is not responding. Please check the device connections and start over.")
            IODevice.ShowDevices()
            If Not RAMfilename3478A = Nothing Then

                'Rename the binary output file to indicate it is incomplete since the download process was interrupted during processing
                Try
                    fs.Close()
                Catch ex As Exception
                End Try
                TextBoxCalramFile3478A.Text = Mid(RAMfilename3478A, 1, RAMfilename3478A.Length - 4) & "_defect.bin"
                System.IO.File.Move(RAMfilename3478A, TextBoxCalramFile3478A.Text)
            End If
            ClearTab3478A()
            Exit Function
        End If
        Return ReadResult
    End Function

    Private Function GetCalibrationRange(RangeIndex As Integer) As String
        Dim StartAddress As Integer
        Dim EndAddress As Integer
        Dim OffsetGain As String

        'Calibration ranges are organized in blocks of 11 nibbles + 2 checksum nibbles starting at 2nd ram address (address 01)
        StartAddress = 1 + RangeIndex * 13
        EndAddress = StartAddress + 12

        'Initialize return parameter
        OffsetGain = Nothing

        'Read the calibration range data from the HP3478A calibration memory
        For i = StartAddress To EndAddress
            'Read calibration data byte of current address
            If GetCalRamByte(i) = 0 Then

                'Convert the ASCII byte to character and append the return parameter
                OffsetGain = OffsetGain & Convert.ToChar(CalramValue3478A)
            Else
                Return "Error"
                Exit Function
            End If
        Next
        'Return the calibration range data
        Return OffsetGain
    End Function
    Private Function GetCalRamByte(CalAddr3478A As Integer) As Integer

        ' Prepare the command to read a byte of calibration ram data from the HP3478A unit's memory
        dev1.enablepoll = False
        dev1.showmessages = False
        ReadAddrByte = Convert.ToByte(CalAddr3478A)                              'convert calibration ram address to byte
        ReadCommandArray(0) = Convert.ToByte(87)                                 'convert ascii "W" into read command byte array
        ReadCommandArray(1) = ReadAddrByte                                       'enter calibration ram address into read command byte array
        ReadCommandArray(2) = Nothing                                            'last byte only required when a Prologix adapter is connected

        'Special handling for Prologix GPIB adapters (prefix specific special characters with ESC-character)
        If lstIntf1.Text = "Prologix Serial" Or lstIntf1.Text = "Prologix Ethernet" Then
            Select Case CalAddr3478A
                Case 10, 13, 27, 43                                              'ASCII characters that must be prefixed with ESC for Prologix adapters
                    ReadCommandArray(2) = ReadCommandArray(1)                    'move the data byte one position down the array to make spece for prefix ESC
                    ReadCommandArray(1) = Convert.ToByte(27)                     'prefix the data byte with ESC charachter
            End Select
        End If

        'Generate the command to read calibration ram data from the unit
        ReadCommand = System.Text.Encoding.Default.GetString(ReadCommandArray)

        'Send the read command with address and wait for reply
        ReadResult = SendQuery(ReadCommand)
        If Not q.status = 0 Then                                                 'check if the connection to the HP3478A is working
            Return -1
            Exit Function
        End If

        'Get value contained in the ram address from read command
        If GetResponse() <> "?" Then
            CalramValue3478A = GetResponse()
        Else
            MessageBox.Show("Device 1 does not appear to be responding. Please check the device connections.")
            Return -1
            Exit Function
        End If
        Return 0
    End Function
End Class