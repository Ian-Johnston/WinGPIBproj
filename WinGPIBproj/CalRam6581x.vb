﻿
' CalRam Advantest R6581 DMM

Imports IODevices
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.IO
Imports Newtonsoft.Json
Imports System.Text.RegularExpressions

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
    Dim Abort6581 As Boolean = False
    Dim RAMfilename6581 As String
    Dim JSONfilename6581 As String


    Private Sub ButtonCalramDumpR6581_Click(sender As Object, e As EventArgs) Handles ButtonCalramDumpR6581.Click

        'R6581

        If AllRegularConstantsReadR6581.Checked = True Then    ' All regular calibration constants
            TextBoxCalRamFile6581.Text = ""
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
            System.Threading.Thread.Sleep(500) ' 500ms delay
            Me.Refresh()

            dev1.SendAsync("*RST", True) ' NPLC 0
            CalramStatus6581.Text = "*RST"
            System.Threading.Thread.Sleep(3000) ' 3sec delay
            Me.Refresh()

            dev1.SendAsync(":VOLT:DC:DIG MAX", True) ' NPLC 0
            CalramStatus6581.Text = "VOLT:DC:DIG MAX"
            System.Threading.Thread.Sleep(250) ' 250ms delay
            Me.Refresh()

            dev1.SendAsync(":SENS:VOLT:DC:RANG 10", True) ' NPLC 0
            CalramStatus6581.Text = "SENS:VOLT:DC:RANG 10"
            System.Threading.Thread.Sleep(250) ' 250ms delay
            Me.Refresh()

            dev1.SendAsync(":VOLT:DC:NPLC 1", True) ' NPLC 1
            CalramStatus6581.Text = "VOLT:DC:NPLC 1"
            System.Threading.Thread.Sleep(250) ' 250ms delay
            Me.Refresh()

            txtr1a.Text = "" ' Prepare reply textbox as empty

            RAMfilename6581 = strPath & "\" & "R6581_Regular_Calibration_Constants_" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".txt"
            JSONfilename6581 = strPath & "\" & "R6581_Regular_Calibration_Constants_" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".json"

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
                Dim sectionList As List(Of Object) = sections(section)
                For Each item In sectionList
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
            End Try

            Dev1TextResponse.Checked = False
            LabelCalRamByte6581.Text = ""

            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            dev1.SendAsync("CAL:EXT:EEPROM:PROTECTION 0", True)         ' Disable Service EXT protection mode.....not required for INT commands

            dev1.SendAsync(":VOLT:DC:NPLC 30", True)                    ' Reset NPLC to 30
            CalramStatus6581.Text = "VOLT:DC:NPLC 30"

            ' Finished
            CalramStatus6581.Text = "DONE!"

        Else
            ' GPIB Dev 1 has not been started
            CalramStatus6581.Text = "DEVICE 1 IS NOT STARTED"
        End If

    End Sub



    Private Function GetCalConstants(startIndex As Integer, endIndex As Integer, statusText As String, commandPrefix As String, queryCommand As String, headerText As String) As List(Of Object)
        ' Prepare to return JSON data
        Dim constantsList As New List(Of Object)

        ' Write header to text file
        IO.File.AppendAllText(RAMfilename6581, Environment.NewLine)
        IO.File.AppendAllText(RAMfilename6581, headerText & Environment.NewLine)

        For i = startIndex To endIndex
            CalramStatus6581.Text = statusText

            If startIndex = 0 And endIndex = 0 Then
                ' do nothing
            Else
                ' Send commands
                dev1.SendAsync(commandPrefix & i & "," & i, True)
                Threading.Thread.Sleep(10)
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
        Dev1TextResponse.Checked = False

    End Sub

End Class