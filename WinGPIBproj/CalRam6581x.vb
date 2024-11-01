
' CalRam Advantest R6581 DMM

Imports IODevices
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.IO
Imports Newtonsoft.Json
Imports System.Text.RegularExpressions
Imports System.IO.File
Imports Newtonsoft.Json.Linq

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
    Private Abort6581 As Boolean = False
    Dim RAMfilename6581 As String
    Dim JSONfilename6581 As String
    Dim stopTime As DateTime
    'Dim abortLoop As Boolean = False


    Private Sub AllRegularConstantsReadR6581_CheckedChanged(sender As Object, e As EventArgs) Handles AllRegularConstantsReadR6581.CheckedChanged

        If AllRegularConstantsReadR6581.Checked = True Then

            ButtonCalramDumpR6581.Enabled = True
            ButtonOpenR6581fileSelectJson.Enabled = False

        End If

    End Sub


    Private Sub ButtonCalramDumpR6581_Click(sender As Object, e As EventArgs) Handles ButtonCalramDumpR6581.Click

        'R6581

        If AllRegularConstantsReadR6581.Checked = True Then    ' All regular calibration constants

            Abort6581 = False

            ButtonCalramDumpR6581.Enabled = False
            ButtonOpenR6581file.Enabled = False
            ButtonOpenR6581fileJson.Enabled = False
            ButtonJsonViewer2.Enabled = False

            TextBoxCalRamFile6581.Text = ""
            TextBoxCalRamFileJson6581.Text = ""
            CalRegularramextractR6581()

        End If


    End Sub


    Private Sub CalFactoryramextractR6581()

        ' R6581 all FACTORY calibration constants read to file

    End Sub


    Private Sub ButtonOpenR6581file_Click(sender As Object, e As EventArgs) Handles ButtonOpenR6581file.Click

        Dim filePath As String = TextBoxCalRamFile6581.Text

        If Not String.IsNullOrEmpty(filePath) AndAlso IO.File.Exists(filePath) Then
            Try
                Process.Start("notepad.exe", filePath)
            Catch ex As Exception
                MessageBox.Show("Unable to open the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("The specified file does not exist. Please check the path and try again.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub


    Private Sub ButtonJsonViewer2_Click(sender As Object, e As EventArgs) Handles ButtonJsonViewer2.Click

        ' Example usage to show JSON content in the tree viewer pop-up
        Dim filePath As String = TextBoxCalRamFileJson6581.Text

        If Not String.IsNullOrEmpty(filePath) AndAlso IO.File.Exists(filePath) Then

            If System.IO.File.Exists(filePath) Then
                Dim jsonText As String = System.IO.File.ReadAllText(filePath)
                Dim viewer As New JsonViewer()
                viewer.LoadJson(jsonText)  ' Prepare the JSON data in the viewer
                viewer.ShowDialog()        ' Show as a pop-up
            Else
                MessageBox.Show("File not found: " & filePath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("The specified file does not exist. Please check the path and try again.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub


    Private Sub ButtonOpenR6581fileJson_Click(sender As Object, e As EventArgs) Handles ButtonOpenR6581fileJson.Click

        Dim filePath As String = TextBoxCalRamFileJson6581.Text

        If Not String.IsNullOrEmpty(filePath) AndAlso IO.File.Exists(filePath) Then
            Try
                Process.Start("notepad.exe", filePath)
            Catch ex As Exception
                MessageBox.Show("Unable to open the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("The specified file does not exist. Please check the path and try again.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub


    Private Sub CalRegularramextractR6581()

        ' R6581 all REGULAR calibration constants read to file

        Abort6581 = False
        CalramStatus6581.Text = "CHECKING SETUP"
        Me.Refresh()

        If ButtonDev1Run.Enabled = True Then ' Device 1 is started

            Dim i As Integer
            Dim s As String
            Dim jsonData As New Dictionary(Of String, Object) ' For storing JSON data
            Dim sections As New Dictionary(Of String, List(Of Object)) ' Sections for JSON

            CalramStatus6581.Text = "SETTING UP GPIB"
            'System.Threading.Thread.Sleep(500) ' 500ms delay
            stopTime = DateTime.Now.AddMilliseconds(500)
            Do While DateTime.Now < stopTime
                Application.DoEvents()
            Loop

            dev1.SendAsync("*RST", True) ' NPLC 0
            CalramStatus6581.Text = "*RST"
            'System.Threading.Thread.Sleep(3000) ' 3sec delay
            stopTime = DateTime.Now.AddMilliseconds(3000)
            Do While DateTime.Now < stopTime
                Application.DoEvents()
            Loop

            dev1.SendAsync(":VOLT:DC:DIG MAX", True) ' NPLC 0
            CalramStatus6581.Text = "VOLT:DC:DIG MAX"
            'System.Threading.Thread.Sleep(250) ' 250ms delay
            stopTime = DateTime.Now.AddMilliseconds(250)
            Do While DateTime.Now < stopTime
                Application.DoEvents()
            Loop

            dev1.SendAsync(":SENS:VOLT:DC:RANG 10", True) ' NPLC 0
            CalramStatus6581.Text = "SENS:VOLT:DC:RANG 10"
            'System.Threading.Thread.Sleep(250) ' 250ms delay
            stopTime = DateTime.Now.AddMilliseconds(250)
            Do While DateTime.Now < stopTime
                Application.DoEvents()
            Loop

            dev1.SendAsync(":VOLT:DC:NPLC 1", True) ' NPLC 1
            CalramStatus6581.Text = "VOLT:DC:NPLC 1"
            'System.Threading.Thread.Sleep(250) ' 250ms delay
            stopTime = DateTime.Now.AddMilliseconds(250)
            Do While DateTime.Now < stopTime
                Application.DoEvents()
            Loop

            txtr1a.Text = "" ' Prepare reply textbox as empty

            RAMfilename6581 = strPath & "\" & "R6581_Regular_Factory_Cal_Constants_" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".txt"
            JSONfilename6581 = strPath & "\" & "R6581_Regular_Factory_Cal_Constants_" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".json"

            TextBoxCalRamFile6581.Text = RAMfilename6581
            TextBoxCalRamFileJson6581.Text = JSONfilename6581

            If Not IO.File.Exists(RAMfilename6581) Then
                IO.File.Create(RAMfilename6581).Dispose()
            End If

            dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 1", True) ' Enable Service EXT protection mode
            Dev1TextResponse.Checked = True

            Threading.Thread.Sleep(100)

            ' Write Main header to text file
            IO.File.AppendAllText(RAMfilename6581, "# ADVANTEST R6581 - CALIBRATION CONSTANTS - " & DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") & Environment.NewLine)

            Threading.Thread.Sleep(100)

            ' Collect calibration constants and add to JSON
            sections("EXT:ZERO:FRONT:EEPROM") = GetCalConstants(0, 46, "READING Regular EXT:ZERO:FRONT:EEPROM", "CAL:EXT:ZERO:FRONT:NUMBER ", "CAL:EXT:ZERO:FRONT:EEPROM:NEW?", "# REGULAR EXT:ZERO:FRONT:EEPROM:")
            sections("EXT:ZERO:REAR:EEPROM") = GetCalConstants(100, 146, "READING Regular EXT:ZERO:REAR:EEPROM", "CAL:EXT:ZERO:REAR:NUMBER ", "CAL:EXT:ZERO:REAR:EEPROM:NEW?", "# REGULAR EXT:ZERO:REAR:EEPROM:")
            sections("EXT:DCV:EEPROM") = GetCalConstants(200, 203, "READING Regular EXT:DCV:EEPROM", "CAL:EXT:DCV:NUMBER ", "CAL:EXT:DCV:EEPROM:NEW?", "# REGULAR EXT:DCV:EEPROM:")
            sections("EXT:OHM:EEPROM") = GetCalConstants(300, 303, "READING Regular EXT:OHM:EEPROM", "CAL:EXT:OHM:NUMBER ", "CAL:EXT:OHM:EEPROM:NEW?", "# REGULAR EXT:OHM:EEPROM:")
            sections("INT:OHM:EEPROM") = GetCalConstants(500, 518, "READING Regular INT:OHM:EEPROM", "CAL:INT:OHM:NUMBER ", "CAL:INT:OHM:EEPROM:NEW?", "# REGULAR INT:OHM:EEPROM:")
            sections("INT:AC:EEPROM") = GetCalConstants(600, 646, "READING Regular INT:AC:EEPROM", "CAL:INT:AC:NUMBER ", "CAL:INT:AC:EEPROM:NEW?", "# REGULAR INT:AC:EEPROM:")
            sections("INT:DCV:EEPROM") = GetCalConstants(400, 406, "READING Regular INT:DCV:EEPROM", "CAL:INT:DCV:NUMBER ", "CAL:INT:DCV:EEPROM:NEW?", "# REGULAR INT:DCV:EEPROM:")

            Threading.Thread.Sleep(100)
            IO.File.AppendAllText(RAMfilename6581, Environment.NewLine)
            IO.File.AppendAllText(RAMfilename6581, "#############################################################" & Environment.NewLine)
            IO.File.AppendAllText(RAMfilename6581, Environment.NewLine)
            Threading.Thread.Sleep(100)

            sections("INT:DCV:HOSEI") = GetCalConstants(0, 25, "READING Factory INT:DCV:HOSIE", "CAL:INT:DCV:HOSEI:NUMBER ", "CAL:INT:DCV:HOSEI?", "# FACTORY INT:DCV:HOSEI:")
            sections("INT:AC:HOSEI") = GetCalConstants(0, 29, "READING Factory INT:AC:HOSIE", "CAL:INT:AC:HOSEI:NUMBER ", "CAL:INT:AC:HOSEI?", "# FACTORY INT:AC:HOSEI:")

            If CheckBoxR6581RetrieveREF.Checked = True Then
                Threading.Thread.Sleep(100)
                IO.File.AppendAllText(RAMfilename6581, Environment.NewLine)
                IO.File.AppendAllText(RAMfilename6581, "#############################################################" & Environment.NewLine)
                IO.File.AppendAllText(RAMfilename6581, Environment.NewLine)
                Threading.Thread.Sleep(100)

                GetCalConstants(0, 0, "READING REF DATA EXT:DCV:EEPROM", "CAL:EXT:DCV:NUMBER ", "CAL:EXT:DCV:EEPROM:REF?", "# REGULAR EXT:DCV:EEPROM: 7.2Vdc INTERNAL REFERENCE DATA HISTORY (20)")        ' works but get out of range error and beep on R6581....?
                GetCalConstants(0, 0, "READING REF DATA EXT:OHM:EEPROM", "CAL:EXT:OHM:NUMBER ", "CAL:EXT:OHM:EEPROM:REF?", "# REGULAR EXT:OHM:EEPROM: 10kohm INTERNAL REFERENCE DATA HISTORY (20)")      ' works but get out of range error and beep on R6581....?
            End If

            ' Strip out the INDEX from the VALUE
            For Each section In sections.Keys

                If Abort6581 = True Then Exit For

                Dim sectionList As List(Of Object) = sections(section)
                For Each item In sectionList

                    If Abort6581 = True Then Exit For

                    ' Use reflection to get the "Value" property dynamically
                    Dim valueProperty = item.GetType().GetProperty("Value")
                    If valueProperty IsNot Nothing Then
                        Dim valueString As String = valueProperty.GetValue(item).ToString().Trim()

                        ' Regular expressions for scientific notation and date/time patterns
                        Dim sciNotationMatch As Match = Regex.Match(valueString, "[+-]?\d+\.\d+E[+-]?\d+")
                        Dim dateTimeMatch As Match = Regex.Match(valueString, "\d{4}/\d{2}/\d{2} \d{2}:\d{2}")

                        If sciNotationMatch.Success Then
                            ' If scientific notation is found, set it as the value
                            valueProperty.SetValue(item, sciNotationMatch.Value)
                        ElseIf dateTimeMatch.Success Then
                            ' If date/time pattern is found, retain it as is
                            valueProperty.SetValue(item, dateTimeMatch.Value)
                        Else
                            ' If none of the patterns match, clear or set to a default
                            valueProperty.SetValue(item, "")
                        End If

                    End If

                    ' Keep the UI responsive
                    Application.DoEvents()

                Next
            Next

            ' Add sections to JSON data
            jsonData("Header") = "ADVANTEST R6581 - CALIBRATION CONSTANTS - " & DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")
            jsonData("Sections") = sections

            ' Write JSON file
            Try
                Dim jsonContent As String = JsonConvert.SerializeObject(jsonData, Formatting.Indented)
                IO.File.WriteAllText(JSONfilename6581, jsonContent)
                'MessageBox.Show("JSON file created successfully at: " & JSONfilename6581, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                'MessageBox.Show("Failed to create JSON file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ButtonCalramDumpR6581.Enabled = True
            End Try

            Dev1TextResponse.Checked = False
            LabelCalRamByte6581.Text = ""

            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)         ' Disable Service EXT protection mode.....not required for INT commands

            Threading.Thread.Sleep(100)

            dev1.SendAsync(":VOLT:DC:NPLC 30", True)                    ' Reset NPLC to 30
            CalramStatus6581.Text = "VOLT:DC:NPLC 30"

            ' Finished
            If Abort6581 = True Then
                CalramStatus6581.Text = "Download aborted!"
                MessageBox.Show("Download aborted!", "Aborted", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                CalramStatus6581.Text = "Download complete!"
                MessageBox.Show("Download complete!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If


            ButtonCalramDumpR6581.Enabled = True
            ButtonOpenR6581file.Enabled = True
            ButtonOpenR6581fileJson.Enabled = True
            ButtonJsonViewer2.Enabled = True

        Else
            ' GPIB Dev 1 has not been started
            'CalramStatus6581.Text = "DEVICE 1 IS NOT STARTED"
            MessageBox.Show("Device 1 is not started", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ButtonCalramDumpR6581.Enabled = True
            ButtonOpenR6581file.Enabled = True
            ButtonOpenR6581fileJson.Enabled = True
            ButtonJsonViewer2.Enabled = True
        End If

    End Sub



    Private Function GetCalConstants(startIndex As Integer, endIndex As Integer, statusText As String, commandPrefix As String, queryCommand As String, headerText As String) As List(Of Object)

        ' Prepare to return JSON data
        Dim constantsList As New List(Of Object)

        ' Write header to text file
        IO.File.AppendAllText(RAMfilename6581, Environment.NewLine)
        IO.File.AppendAllText(RAMfilename6581, headerText & Environment.NewLine)

        For i = startIndex To endIndex

            If Abort6581 = True Then Exit For

            CalramStatus6581.Text = statusText

            If startIndex = 0 And endIndex = 0 Then
                ' do nothing
            Else
                ' Send commands
                dev1.SendAsync(commandPrefix & i & "," & i, True)
                Threading.Thread.Sleep(100)
            End If

            ' Query and get response
            Dim q As IOQuery = Nothing
            dev1.QueryBlocking(queryCommand, q, False)
            Cbdev1(q)

            ' Update status and text file
            LabelCalRamByte6581.Text = q.ResponseAsString
            Me.Refresh()
            IO.File.AppendAllText(RAMfilename6581, q.ResponseAsString & Environment.NewLine)

            ' Add data to JSON structure
            constantsList.Add(New With {.Index = i, .Value = q.ResponseAsString})

            Threading.Thread.Sleep(10)
        Next

        Return constantsList

    End Function



    Private Sub ShowFilesCalRamR6581_Click(sender As Object, e As EventArgs) Handles ShowFilesCalRamR6581.Click

        Process.Start("explorer.exe", String.Format("/n, /e, {0}", strPath))

    End Sub


    Private Sub ButtonR6581abort_Click(sender As Object, e As EventArgs) Handles ButtonR6581abort.Click

        'dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)         ' Disable Service EXT protection mode.....not required for INT commands

        Abort6581 = True
        TextBoxCalRamFile6581.Text = ""
        TextBoxCalRamFileJson6581.Text = ""
        Dev1TextResponse.Checked = False

        ButtonCalramDumpR6581.Enabled = True
        ButtonOpenR6581file.Enabled = True
        ButtonOpenR6581fileJson.Enabled = True
        ButtonJsonViewer2.Enabled = True

    End Sub



    ' ###############################################################################


    ' Checkbox change
    Private Sub SendRegularConstantsReadR6581_CheckedChanged(sender As Object, e As EventArgs) Handles SendRegularConstantsReadR6581.CheckedChanged

        If SendRegularConstantsReadR6581.Checked = True Then

            ButtonCalramDumpR6581.Enabled = False
            ButtonOpenR6581fileSelectJson.Enabled = True

        End If

    End Sub


    ' Button to select JSON file
    Private Sub ButtonOpenR6581fileSelectJson_Click(sender As Object, e As EventArgs) Handles ButtonOpenR6581fileSelectJson.Click

        If SendRegularConstantsReadR6581.Checked = True Then

            ButtonJsonViewer.Enabled = True

            Dim fd As New OpenFileDialog()
            fd.Title = "Select Calibration JSON File"
            fd.InitialDirectory = strPath ' Set the initial directory
            fd.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
            fd.FilterIndex = 1
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                ' Set the selected file path to the text box
                TextBoxCalRamFileJson6581Select.Text = fd.FileName
                ButtonR6581upload.Enabled = True

                ' Call the subroutine to check JSON file contents
                CheckJsonFileContents(fd.FileName)
            Else
                MessageBox.Show("No file selected. Please choose a valid JSON file.", "File Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ButtonR6581upload.Enabled = False
                ButtonR6581commitEEprom.Enabled = False
            End If

        End If

    End Sub


    ' Subroutine to check and display JSON file contents after the user browses to one
    Private Sub CheckJsonFileContents(jsonFilePath As String)
        Try
            ' Ensure the file exists
            If String.IsNullOrEmpty(jsonFilePath) OrElse Not IO.File.Exists(jsonFilePath) Then
                MessageBox.Show("The selected file does not exist or the path is invalid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Load and parse JSON data from the file
            Dim jsonData As String = IO.File.ReadAllText(jsonFilePath)
            Dim calibrationData As JObject = JObject.Parse(jsonData)

            ' Display Header if available
            If calibrationData.ContainsKey("Header") Then
                Dim header As String = calibrationData("Header").ToString()
                MessageBox.Show("For your confirmation:" & Environment.NewLine & Environment.NewLine & "Header:" & Environment.NewLine & header, "JSON File Contents", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("No header found in the JSON file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            ' Check and list all sections available in the JSON
            If calibrationData.ContainsKey("Sections") Then
                Dim sections As JObject = calibrationData("Sections")
                Dim sectionNames As String = String.Join(Environment.NewLine, sections.Properties().Select(Function(p) p.Name).ToArray())

                ' Identify the sections that are selected for updating
                Dim sectionMapping As New Dictionary(Of CheckBox, String) From {
                {CheckBoxR6581Upload1, "EXT:ZERO:FRONT:EEPROM"},
                {CheckBoxR6581Upload2, "EXT:ZERO:REAR:EEPROM"},
                {CheckBoxR6581Upload3, "EXT:DCV:EEPROM"},
                {CheckBoxR6581Upload4, "EXT:OHM:EEPROM"},
                {CheckBoxR6581Upload5, "INT:OHM:EEPROM"},
                {CheckBoxR6581Upload6, "INT:AC:EEPROM"},
                {CheckBoxR6581Upload7, "INT:DCV:EEPROM"},
                {CheckBoxR6581Upload8, "INT:DCV:HOSEI"},
                {CheckBoxR6581Upload9, "INT:AC:HOSEI"}
            }

                ' Prepare a list of selected sections for updating
                Dim selectedSections As New List(Of String)
                For Each kvp In sectionMapping
                    If kvp.Key.Checked Then
                        selectedSections.Add(kvp.Value)
                    End If
                Next

                Dim selectedSectionNames As String = String.Join(Environment.NewLine, selectedSections)

                ' Display message with sections found and sections selected for updating
                MessageBox.Show("For your confirmation:" & Environment.NewLine & Environment.NewLine &
                            "Sections Found:" & Environment.NewLine & sectionNames & Environment.NewLine &
                            Environment.NewLine & "Sections Selected for Updating:" & Environment.NewLine & selectedSectionNames,
                            "JSON File Contents", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else
                MessageBox.Show("No sections found in the JSON file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            ' Handle any parsing or file reading errors
            MessageBox.Show("An error occurred while checking the JSON file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' Button to initiate the upload process
    Private Sub ButtonR6581upload_Click(sender As Object, e As EventArgs) Handles ButtonR6581upload.Click

        If ButtonDev1Run.Enabled = True Then ' Device 1 is started

            TextBoxR6581GPIBlist.Clear()

            CalramStatus6581upload.Text = "SETTING UP GPIB"
            System.Threading.Thread.Sleep(500) ' 500ms delay
            Me.Refresh()

            'dev1.SendAsync("*RST", True) ' NPLC 0
            CalramStatus6581upload.Text = "*RST"
            System.Threading.Thread.Sleep(3000) ' 3sec delay
            Me.Refresh()

            'dev1.SendAsync(":VOLT:DC:DIG MAX", True) ' NPLC 0
            CalramStatus6581upload.Text = "VOLT:DC:DIG MAX"
            System.Threading.Thread.Sleep(250) ' 250ms delay
            Me.Refresh()

            'dev1.SendAsync(":SENS:VOLT:DC:RANG 10", True) ' NPLC 0
            CalramStatus6581upload.Text = "SENS:VOLT:DC:RANG 10"
            System.Threading.Thread.Sleep(250) ' 250ms delay
            Me.Refresh()

            'dev1.SendAsync(":VOLT:DC:NPLC 1", True) ' NPLC 1
            CalramStatus6581upload.Text = "VOLT:DC:NPLC 1"
            System.Threading.Thread.Sleep(250) ' 250ms delay
            Me.Refresh()

            Dim jsonFilePath As String = TextBoxCalRamFileJson6581Select.Text

            ' Ensure a valid JSON file path is provided
            If String.IsNullOrEmpty(jsonFilePath) OrElse Not System.IO.File.Exists(jsonFilePath) Then
                MessageBox.Show("Please select a valid JSON file before proceeding.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ButtonR6581commitEEprom.Enabled = False
                Return
            End If


            ' Call the upload function, passing in the selected JSON file path and checkboxes
            Dim checkboxes As New List(Of CheckBox) From {CheckBoxR6581Upload1, CheckBoxR6581Upload2, CheckBoxR6581Upload3, CheckBoxR6581Upload4, CheckBoxR6581Upload5, CheckBoxR6581Upload6, CheckBoxR6581Upload7, CheckBoxR6581Upload8, CheckBoxR6581Upload9}
            UploadSelectedCalibrationDataFromJson(jsonFilePath, checkboxes)

            ButtonR6581commitEEprom.Enabled = True

        Else
            ' GPIB Dev 1 has not been started
            'CalramStatus6581upload.Text = "DEVICE 1 IS NOT STARTED"
            MessageBox.Show("Device 1 is not started", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub


    ' Upload selected calibration data from the JSON file based on checkbox selection
    Private Sub UploadSelectedCalibrationDataFromJson(jsonFilePath As String, checkboxes As List(Of CheckBox))
        Try
            ' Load JSON data from file
            Dim jsonData As String = IO.File.ReadAllText(jsonFilePath)
            Dim calibrationData As JObject = JObject.Parse(jsonData)

            ' Assuming the calibration sections are structured under "Sections"
            If calibrationData.ContainsKey("Sections") Then
                ' Enable service EXT protection mode if needed
                ' dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 1", True)

                Dim sections As JObject = calibrationData("Sections")

                ' Define mapping between checkboxes and the section names
                Dim sectionMapping As New Dictionary(Of CheckBox, String) From {
                {CheckBoxR6581Upload1, "EXT:ZERO:FRONT:EEPROM"},
                {CheckBoxR6581Upload2, "EXT:ZERO:REAR:EEPROM"},
                {CheckBoxR6581Upload3, "EXT:DCV:EEPROM"},
                {CheckBoxR6581Upload4, "EXT:OHM:EEPROM"},
                {CheckBoxR6581Upload5, "INT:OHM:EEPROM"},
                {CheckBoxR6581Upload6, "INT:AC:EEPROM"},
                {CheckBoxR6581Upload7, "INT:DCV:EEPROM"},
                {CheckBoxR6581Upload8, "INT:DCV:HOSEI"},
                {CheckBoxR6581Upload9, "INT:AC:HOSEI"}
            }

                ' Loop through each checkbox and process selected sections
                For Each kvp In sectionMapping
                    Dim checkBox As CheckBox = kvp.Key
                    Dim eepromSection As String = kvp.Value

                    If checkBox.Checked AndAlso sections.ContainsKey(eepromSection) Then
                        ' Extract the section as a JArray
                        Dim calibrationList As JArray = CType(sections(eepromSection), JArray)

                        ' Determine the GPIB command prefix based on the section
                        Dim ramCommandPrefix As String = GetRamCommandPrefix(eepromSection)

                        ' Loop through each calibration data entry and send it
                        For Each calibrationItem As JObject In calibrationList
                            Dim index As Integer = Convert.ToInt32(calibrationItem("Index"))
                            Dim value As String = calibrationItem("Value").ToString().Trim()

                            ' Construct and send GPIB command
                            Dim command As String = $"{ramCommandPrefix} {index},{value}"
                            LabelCalRamByte6581upload.Text = command

                            ' Add the command to the TextBox for logging
                            TextBoxR6581GPIBlist.AppendText(command & Environment.NewLine)

                            ' Force the UI to update and process pending events
                            Application.DoEvents()

                            ' dev1.SendAsync(command, True)

                            ' Add a sizeable delay after each command to give the R6581 time
                            Threading.Thread.Sleep(200)
                        Next
                    End If
                Next

                CalramStatus6581upload.Text = "Ram upload complete!"

                ' Disable service EXT protection mode if needed
                ' dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)

            Else
                MessageBox.Show("Sections key not found in the JSON file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ' dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)
            End If

        Catch ex As Exception
            ' Handle JSON parsing errors or other issues
            MessageBox.Show("An error occurred while reading or processing the JSON file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)
        End Try
    End Sub


    ' Button to commit RAM contents to EEPROM
    Private Sub ButtonR6581commitEEprom_Click(sender As Object, e As EventArgs) Handles ButtonR6581commitEEprom.Click

        If ButtonDev1Run.Enabled = True Then ' Device 1 is started

            ' Show confirmation dialog
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to transfer RAM contents to EEPROM?", "Confirm Action", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If result = DialogResult.Yes Then

                ButtonR6581upload.Enabled = False
                ButtonR6581commitEEprom.Enabled = False

                'dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 1", True)         ' Enable Service EXT protection mode.....not required for INT commands

                TextBoxR6581GPIBlist.AppendText("CAL:EXT:EEPROM:PROTECTION 1" & Environment.NewLine)                        ' Add the command to the TextBox for logging


                ' REGULAR

                If CheckBoxR6581Upload1.Checked = True Or CheckBoxR6581Upload2.Checked = True Or CheckBoxR6581Upload3.Checked = True Or CheckBoxR6581Upload4.Checked = True Or CheckBoxR6581Upload5.Checked = True Or CheckBoxR6581Upload6.Checked = True Or CheckBoxR6581Upload7.Checked = True Then
                    ' Send GPIB command to commit RAM to EEPROM
                    Dim commandon As String = "CAL:EXT:EEPROM:STORE 1"
                    'dev1.SendAsync(commandon, True)
                    TextBoxR6581GPIBlist.AppendText("CAL:EXT:EEPROM:STORE 1" & Environment.NewLine)                            ' Add the command to the TextBox for logging

                    CalramStatus6581upload.Text = "Regular RAM contents have been committed to EEPROM"

                    Threading.Thread.Sleep(100)
                End If


                ' FACTORY DCV

                If CheckBoxR6581Upload8.Checked = True Then
                    ' Send GPIB command to commit RAM to EEPROM
                    Dim command8on As String = "CAL:INT:DCV:HOSEI:STORE 1"
                    'dev1.SendAsync(command8on, True)
                    TextBoxR6581GPIBlist.AppendText("CAL:INT:DCV:HOSEI:STORE 1" & Environment.NewLine)                            ' Add the command to the TextBox for logging

                    CalramStatus6581upload.Text = "Factory DCV RAM contents have been committed to EEPROM"

                    Threading.Thread.Sleep(100)
                End If


                ' FACTORY AC

                If CheckBoxR6581Upload9.Checked = True Then
                    ' Send GPIB command to commit RAM to EEPROM
                    Dim command9on As String = "CAL:INT:AC:HOSEI:STORE 1"
                    'dev1.SendAsync(command9on, True)
                    TextBoxR6581GPIBlist.AppendText("CAL:INT:AC:HOSEI:STORE 1" & Environment.NewLine)                            ' Add the command to the TextBox for logging

                    CalramStatus6581upload.Text = "Factory AC RAM contents have been committed to EEPROM"

                    Threading.Thread.Sleep(100)
                End If


                ' FINISH UP

                'dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)         ' Disable Service EXT protection mode.....not required for INT commands
                TextBoxR6581GPIBlist.AppendText("CAL:EXT:EEPROM:PROTECTION 0" & Environment.NewLine)                        ' Add the command to the TextBox for logging

                Threading.Thread.Sleep(100)

                'dev1.SendAsync(":VOLT:DC:NPLC 30", True)                    ' Reset NPLC to 30
                CalramStatus6581upload.Text = "VOLT:DC:NPLC 30"

                ' Pause for 2000ms without freezing the GUI - Delay to give the R6581 time to mass copy EEprom contents
                'Threading.Thread.Sleep(2000)    ' delay to give the R6581 time to mass copy EEprom contents
                stopTime = DateTime.Now.AddMilliseconds(2500)
                Do While DateTime.Now < stopTime
                    Application.DoEvents()
                Loop

                CalramStatus6581upload.Text = "Commit to EEprom complete!"
                MessageBox.Show("All RAM contents have been committed to EEPROM.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else

                'dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)         ' Disable Service EXT protection mode.....not required for INT commands

            End If

        Else
            ' GPIB Dev 1 has not been started
            'CalramStatus6581upload.Text = "DEVICE 1 IS NOT STARTED"
            MessageBox.Show("Device 1 is not started", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub


    ' Helper function to map EEPROM section to corresponding RAM GPIB command
    Private Function GetRamCommandPrefix(eepromSection As String) As String

        Select Case eepromSection
            Case "EXT:ZERO:FRONT:EEPROM"
                Return "CAL:EXT:ZERO:FRONT:RAM"
            Case "EXT:ZERO:REAR:EEPROM"
                Return "CAL:EXT:ZERO:REAR:RAM"
            Case "EXT:DCV:EEPROM"
                Return "CAL:EXT:DCV:RAM"
            Case "EXT:OHM:EEPROM"
                Return "CAL:EXT:OHM:RAM"
            Case "INT:OHM:EEPROM"
                Return "CAL:INT:OHM:RAM"
            Case "INT:AC:EEPROM"
                Return "CAL:INT:AC:RAM"
            Case "INT:DCV:EEPROM"
                Return "CAL:INT:DCV:RAM"
            Case "INT:DCV:HOSEI"
                Return "CAL:INT:DCV:HOSEI"
            Case "INT:AC:HOSEI"
                Return "CAL:INT:AC:HOSEI"
            Case Else
                Return String.Empty
        End Select

    End Function


    Private Sub ButtonJsonViewer_Click(sender As Object, e As EventArgs) Handles ButtonJsonViewer.Click
        ' Example usage to show JSON content in the tree viewer pop-up
        Dim filePath As String = TextBoxCalRamFileJson6581Select.Text

        If System.IO.File.Exists(filePath) Then
            Dim jsonText As String = System.IO.File.ReadAllText(filePath)
            Dim viewer As New JsonViewer()
            viewer.LoadJson(jsonText)  ' Prepare the JSON data in the viewer
            viewer.ShowDialog()        ' Show as a pop-up
        Else
            MessageBox.Show("File not found: " & filePath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub


End Class