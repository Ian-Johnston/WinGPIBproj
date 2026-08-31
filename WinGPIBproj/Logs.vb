' Data log, event log, CSV generation and main app chart control

Partial Class Formtest


    Dim file As System.IO.StreamWriter
    Dim filepath As String
    Dim filename As String

    Dim CSVupdate As Boolean = False

    Dim IndexCount As Double = 1      ' for CSV file

    Dim CSVlength As Long

    Dim CSVfilestarting As Boolean = False
    Dim AllowMetaDataWrite As Boolean = True
    Dim CSVfilenameprevious As String = "somedummyfilename"

    ' Create a list to store log entries
    Private EventlogEntries As New List(Of String)

    Dim StartCSVLogClicked As Boolean = False


    Private Sub Log(ByVal str As String)

        ' Add the new log entry to the logEntries list
        EventlogEntries.Add(str)

        ' Clear the ListBox and add updated entries
        ListLog.Items.Clear()
        For Each entry As String In EventlogEntries
            ListLog.Items.Add(entry)
        Next

        ' Scroll to the bottom of the ListBox
        ListLog.TopIndex = ListLog.Items.Count - 1

    End Sub


    Private Sub ClearEventLOG_Click(sender As Object, e As EventArgs) Handles ClearEventLOG.Click

        EventlogEntries.Clear()
        ListLog.Items.Clear()       ' Clear the log data from display

    End Sub


    ' Create a list to store log entries
    ' Private logEntries As New List(Of String)


    Private Sub LogData(device As String,
                    logDate As String,
                    logTime As String,
                    value As String,
                    temperature As String,
                    humidity As String)

        If CheckboxEnableLOG.Checked = False Then Exit Sub


        ' ==========================================================
        ' Device 1 statistics
        ' ==========================================================

        Dim dev1Samples As String = ""
        Dim dev1Mean As String = ""
        Dim dev1Stdev As String = ""
        Dim dev1SEM As String = ""
        Dim dev1Gain As String = ""
        Dim dev1MaxDiff As String = ""
        Dim dev1Deviation As String = ""

        If CheckBoxStats1Enable.Checked = True AndAlso Stats1Count > 0 Then

            dev1Samples = Stats1Count.ToString()

            If ENotationDecimal.Checked = True Then

                dev1Mean =
                Stats1Mean.ToString("#0.0000000000")

                dev1Stdev =
                Stats1StdevCurrent.ToString("#0.0000000000")

                dev1SEM =
                Stats1SEMCurrent.ToString("#0.0000000000")

                dev1MaxDiff =
                (Stats1Max - Stats1Min).ToString("#0.0000000000")

                dev1Deviation =
                Stats1DeviationCurrent.ToString("#0.0000")

            Else

                dev1Mean =
                Stats1Mean.ToString("0.000E+00")

                dev1Stdev =
                Stats1StdevCurrent.ToString("0.000E+00")

                dev1SEM =
                Stats1SEMCurrent.ToString("0.000E+00")

                dev1MaxDiff =
                (Stats1Max - Stats1Min).ToString("0.000E+00")

                dev1Deviation =
                Stats1DeviationCurrent.ToString("0.0000")

            End If

            dev1Gain =
            (0.5 * Math.Log10(Stats1Count)).ToString("0.00")

        End If


        ' ==========================================================
        ' Device 2 statistics
        ' ==========================================================

        Dim dev2Samples As String = ""
        Dim dev2Mean As String = ""
        Dim dev2Stdev As String = ""
        Dim dev2SEM As String = ""
        Dim dev2Gain As String = ""
        Dim dev2MaxDiff As String = ""
        Dim dev2Deviation As String = ""

        If CheckBoxStats2Enable.Checked = True AndAlso Stats2Count > 0 Then

            dev2Samples = Stats2Count.ToString()

            If ENotationDecimal.Checked = True Then

                dev2Mean =
                Stats2Mean.ToString("#0.0000000000")

                dev2Stdev =
                Stats2StdevCurrent.ToString("#0.0000000000")

                dev2SEM =
                Stats2SEMCurrent.ToString("#0.0000000000")

                dev2MaxDiff =
                (Stats2Max - Stats2Min).ToString("#0.0000000000")

                dev2Deviation =
                Stats2DeviationCurrent.ToString("#0.0000")

            Else

                dev2Mean =
                Stats2Mean.ToString("0.000E+00")

                dev2Stdev =
                Stats2StdevCurrent.ToString("0.000E+00")

                dev2SEM =
                Stats2SEMCurrent.ToString("0.000E+00")

                dev2MaxDiff =
                (Stats2Max - Stats2Min).ToString("0.000E+00")

                dev2Deviation =
                Stats2DeviationCurrent.ToString("0.0000")

            End If

            dev2Gain =
            (0.5 * Math.Log10(Stats2Count)).ToString("0.00")

        End If


        ' ==========================================================
        ' Add row
        ' ==========================================================

        DataGridViewLogData.Rows.Add(
        device,
        logDate,
        logTime,
        value,
        temperature,
        humidity,
        dev1Samples,
        dev1Mean,
        dev1Stdev,
        dev1SEM,
        dev1Gain,
        dev1MaxDiff,
        dev1Deviation,
        dev2Samples,
        dev2Mean,
        dev2Stdev,
        dev2SEM,
        dev2Gain,
        dev2MaxDiff,
        dev2Deviation)


        ' Keep maximum of 31 displayed entries
        While DataGridViewLogData.Rows.Count > 31
            DataGridViewLogData.Rows.RemoveAt(0)
        End While


        ' Scroll to newest entry
        If DataGridViewLogData.Rows.Count > 0 Then

            DataGridViewLogData.FirstDisplayedScrollingRowIndex =
            DataGridViewLogData.Rows.Count - 1

        End If

    End Sub


    Private Sub ClearLOGdisp_Click(sender As Object, e As EventArgs) Handles ClearLOGdisp.Click

        DataGridViewLogData.Rows.Clear()

        Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " LOG Cleared")

    End Sub


    Private Sub LOGdisplay()

        ' ==========================================================
        ' Device 1 Log display
        ' ==========================================================

        If ((ButtonDev1Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop") AndAlso
        CheckboxEnableLOG.Checked = True AndAlso
        txtr1a.Text <> "" AndAlso
        Dev1GPIBActivity = True AndAlso
        StartCSVLogClicked = True) Then

            TestLen = Len(txtr1a.Text)

            Dim valueText As String

            If ENotationDecimal.Checked = True AndAlso TestLen > 4 Then

                valueText =
                Format(
                    CDbl(Val(NormalizeNumericResponse(txtr1a.Text))),
                    "#0.0000000000")

            Else

                valueText = txtr1a.Text

            End If


            Dim temperatureText As String = "0.0"
            Dim humidityText As String = "0.0"


            If TempHumLogs.Checked = True Then

                If LabelTemperature.Text = "NaN" OrElse
               LabelTemperature.Text = "nil" OrElse
               LabelHumidity.Text = "NaN" OrElse
               LabelHumidity.Text = "nil" Then

                    LabelTemperature.Text = "0.0"
                    LabelHumidity.Text = "0.0"

                End If


                Dim valueTemp As Double = CDbl(LabelTemperature.Text)
                temperatureText = valueTemp.ToString("0.00")

                Dim valueHum As Double = CDbl(LabelHumidity.Text)
                humidityText = valueHum.ToString("0.00")

            End If


            LogData(
            txtname1.Text,
            DateTime.Now.ToString("yyyy-MM-dd"),
            DateTime.Now.ToString("HH:mm:ss"),
            valueText,
            temperatureText,
            humidityText)

        End If


        ' ==========================================================
        ' Device 2 Log display
        ' ==========================================================

        If ((ButtonDev2Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop") AndAlso
        CheckboxEnableLOG.Checked = True AndAlso
        txtr2a.Text <> "" AndAlso
        Dev2GPIBActivity = True AndAlso
        StartCSVLogClicked = True) Then

            TestLen = Len(txtr2a.Text)

            Dim valueText As String

            If ENotationDecimal.Checked = True AndAlso TestLen > 4 Then

                valueText =
                Format(
                    CDbl(Val(NormalizeNumericResponse(txtr2a.Text))),
                    "#0.0000000000")

            Else

                valueText = txtr2a.Text

            End If


            Dim temperatureText As String = "0.0"
            Dim humidityText As String = "0.0"


            If TempHumLogs.Checked = True Then

                If LabelTemperature.Text = "NaN" OrElse
               LabelTemperature.Text = "nil" OrElse
               LabelHumidity.Text = "NaN" OrElse
               LabelHumidity.Text = "nil" Then

                    LabelTemperature.Text = "0.0"
                    LabelHumidity.Text = "0.0"

                End If


                Dim valueTemp As Double = CDbl(LabelTemperature.Text)
                temperatureText = valueTemp.ToString("0.00")

                Dim valueHum As Double = CDbl(LabelHumidity.Text)
                humidityText = valueHum.ToString("0.00")

            End If


            LogData(
            txtname2.Text,
            DateTime.Now.ToString("yyyy-MM-dd"),
            DateTime.Now.ToString("HH:mm:ss"),
            valueText,
            temperatureText,
            humidityText)

        End If

    End Sub


    Private Function GetLiveStatisticsCSV() As String

        ' Blank by default means statistics were not enabled.
        Dim dev1CountText As String = ""
        Dim dev1MeanText As String = ""
        Dim dev1StdevText As String = ""
        Dim dev1SEMText As String = ""
        Dim dev1GainText As String = ""
        Dim dev1MaxDiffText As String = ""
        Dim dev1DeviationText As String = ""

        Dim dev2CountText As String = ""
        Dim dev2MeanText As String = ""
        Dim dev2StdevText As String = ""
        Dim dev2SEMText As String = ""
        Dim dev2GainText As String = ""
        Dim dev2MaxDiffText As String = ""
        Dim dev2DeviationText As String = ""


        ' ==========================================================
        ' Device 1 statistics
        ' ==========================================================

        If CheckBoxStats1Enable.Checked = True AndAlso Stats1Count > 0 Then

            dev1CountText =
            Stats1Count.ToString(Globalization.CultureInfo.InvariantCulture)

            If ENotationDecimal.Checked = True Then

                dev1MeanText =
                Stats1Mean.ToString("#0.0000000000",
                                    Globalization.CultureInfo.InvariantCulture)

                dev1StdevText =
                Stats1StdevCurrent.ToString("#0.0000000000",
                                            Globalization.CultureInfo.InvariantCulture)

                dev1SEMText =
                Stats1SEMCurrent.ToString("#0.0000000000",
                                          Globalization.CultureInfo.InvariantCulture)

                dev1MaxDiffText =
                (Stats1Max - Stats1Min).ToString("#0.0000000000",
                                    Globalization.CultureInfo.InvariantCulture)

                dev1DeviationText =
                Stats1DeviationCurrent.ToString("#0.0000",
                                    Globalization.CultureInfo.InvariantCulture)

            Else

                dev1MeanText =
                Stats1Mean.ToString("0.0000000000E+00",
                                    Globalization.CultureInfo.InvariantCulture)

                dev1StdevText =
                Stats1StdevCurrent.ToString("0.0000000000E+00",
                                            Globalization.CultureInfo.InvariantCulture)

                dev1SEMText =
                Stats1SEMCurrent.ToString("0.0000000000E+00",
                                          Globalization.CultureInfo.InvariantCulture)

                dev1MaxDiffText =
                (Stats1Max - Stats1Min).ToString("0.0000000000E+00",
                                    Globalization.CultureInfo.InvariantCulture)

                dev1DeviationText =
                Stats1DeviationCurrent.ToString("0.0000",
                                    Globalization.CultureInfo.InvariantCulture)

            End If

            dev1GainText =
            (0.5 * Math.Log10(Stats1Count)).
            ToString("0.00", Globalization.CultureInfo.InvariantCulture)

        End If


        ' ==========================================================
        ' Device 2 statistics
        ' ==========================================================

        If CheckBoxStats2Enable.Checked = True AndAlso Stats2Count > 0 Then

            dev2CountText =
            Stats2Count.ToString(Globalization.CultureInfo.InvariantCulture)

            If ENotationDecimal.Checked = True Then

                dev2MeanText =
                Stats2Mean.ToString("#0.0000000000",
                                    Globalization.CultureInfo.InvariantCulture)

                dev2StdevText =
                Stats2StdevCurrent.ToString("#0.0000000000",
                                            Globalization.CultureInfo.InvariantCulture)

                dev2SEMText =
                Stats2SEMCurrent.ToString("#0.0000000000",
                                          Globalization.CultureInfo.InvariantCulture)

                dev2MaxDiffText =
                (Stats2Max - Stats2Min).ToString("#0.0000000000",
                                    Globalization.CultureInfo.InvariantCulture)

                dev2DeviationText =
                Stats2DeviationCurrent.ToString("#0.0000",
                                    Globalization.CultureInfo.InvariantCulture)

            Else

                dev2MeanText =
                Stats2Mean.ToString("0.0000000000E+00",
                                    Globalization.CultureInfo.InvariantCulture)

                dev2StdevText =
                Stats2StdevCurrent.ToString("0.0000000000E+00",
                                            Globalization.CultureInfo.InvariantCulture)

                dev2SEMText =
                Stats2SEMCurrent.ToString("0.0000000000E+00",
                                          Globalization.CultureInfo.InvariantCulture)

                dev2MaxDiffText =
                (Stats2Max - Stats2Min).ToString("0.0000000000E+00",
                                    Globalization.CultureInfo.InvariantCulture)

                dev2DeviationText =
                Stats2DeviationCurrent.ToString("0.0000",
                                    Globalization.CultureInfo.InvariantCulture)

            End If

            dev2GainText =
            (0.5 * Math.Log10(Stats2Count)).
            ToString("0.00", Globalization.CultureInfo.InvariantCulture)

        End If


        ' ==========================================================
        ' Return fixed-position statistics fields
        '
        ' Max Diff / Deviation are appended AFTER the original 10
        ' fields (5 per device), not interleaved, so existing CSVs
        ' and the Playback reader's V5 field positions (indices
        ' 6-15) are completely unaffected.
        ' ==========================================================

        Return dev1CountText & CSVdelimit &
           dev1MeanText & CSVdelimit &
           dev1StdevText & CSVdelimit &
           dev1SEMText & CSVdelimit &
           dev1GainText & CSVdelimit &
           dev2CountText & CSVdelimit &
           dev2MeanText & CSVdelimit &
           dev2StdevText & CSVdelimit &
           dev2SEMText & CSVdelimit &
           dev2GainText & CSVdelimit &
           dev1MaxDiffText & CSVdelimit &
           dev1DeviationText & CSVdelimit &
           dev2MaxDiffText & CSVdelimit &
           dev2DeviationText

    End Function


    Private Sub CSVfile()

        ' Write data to CSV file, for both Dev1 & Dev2

        If (CheckboxEnableCSV.Checked = True) Then

            'CSVfilename.ReadOnly = True
            'CSVfilepath.ReadOnly = True
            ButtonExportCSV.Enabled = False
            CSVdelimiterComma.Enabled = False
            CSVdelimiterSemiColon.Enabled = False

            ' if folder is left blank on exit (via saved settings) then on startup it will generate the path as the existing folder where the program is being run from
            If (CSVfilepath.Text = "") Then
                CSVfilepath.Text = strPath
            End If

            ' When manually editing the folder path, check path and name, remove characters which shouldn't be there
            ' remove / or \ if appear at end of filepath string
            CSVfilepath.Text = CSVfilepath.Text.Replace("/", "\")
            CSVfilename.Text = CSVfilename.Text.Replace("/", "")
            CSVfilename.Text = CSVfilename.Text.Replace("\", "")
            If (CSVfilepath.Text.Substring(CSVfilepath.Text.Length - 1)) = "/" Or (CSVfilepath.Text.Substring(CSVfilepath.Text.Length - 1)) = "\" Then
                CSVfilepath.Text = CSVfilepath.Text.Substring(0, CSVfilepath.Text.Length - 1)
            End If

            ' check that the CSV file specified exists, if it doesn't then create a new blank CSV file based on the name entered by the user
            If System.IO.File.Exists(CSVfilepath.Text & "\" & CSVfilename.Text) Then
                'the file exists
                AllowMetaDataWrite = False
            Else
                'the file doesn't exist
                System.IO.File.Create(CSVfilepath.Text & "\" & CSVfilename.Text).Dispose()
                AllowMetaDataWrite = True
            End If

            ' If file was cleared then allow MetaData
            Dim fileCheckInfo As New System.IO.FileInfo(CSVfilepath.Text & "\" & CSVfilename.Text)
            If fileCheckInfo.Length = 0 Then
                AllowMetaDataWrite = True
            End If

            ' update file length (lines)
            LabelCSVfilesize.Enabled = True
            CSVsize.Enabled = True
            CSVlength = System.IO.File.ReadAllLines(CSVfilepath.Text & "\" & CSVfilename.Text).Length
            CSVsize.Text = CSVlength + 1

        Else
            'CSVfilename.ReadOnly = False
            'CSVfilepath.ReadOnly = False
            ButtonExportCSV.Enabled = True
            CSVsize.Text = "##"
            LabelCSVfilesize.Enabled = False
            CSVsize.Enabled = False
            CSVdelimiterComma.Enabled = True
            CSVdelimiterSemiColon.Enabled = True
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter(CSVfilepath.Text & "\" & CSVfilename.Text, True)

        ' Write metadata if logging has just started and either Dev 1 or Dev2
        If (AllowMetaDataWrite = True And CSVfilestarting = True And CheckboxEnableCSV.Checked = True) Then
            CSVfilestarting = False     ' set so that this is only done when Enable Start CSV box is first checked
            AllowMetaDataWrite = False  ' set so that for this CSV the metadata should not be written again

            ' Split the text by lines and prefix each line with "//"
            Dim textToSave As String = LogFileMetadata.Text
            Dim lines As String() = textToSave.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            For Each line As String In lines
                file.WriteLine($"// {line}")
            Next

            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Metadata to CSV Written")

        End If


        ' Fixed Live Statistics fields appended to every CSV record
        Dim statsCSV As String = GetLiveStatisticsCSV()

        ' Device 1 updates
        If ((ButtonDev1Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop") And CheckboxEnableCSV.Checked = True And txtr1a.Text <> "" And Dev1GPIBActivity = True) Then

            If (TempHumLogs.Checked = True And TempHumConnected = True) Then
                If LabelTemperature.Text = "NaN" Or LabelTemperature.Text = "nil" Or LabelHumidity.Text = "NaN" Or LabelHumidity.Text = "nil" Then
                    LabelTemperature.Text = "0.0"
                    LabelHumidity.Text = "0.0"
                End If

                If (ENotationDecimal.Checked = True) Then
                    ' decimal
                    Dim v1 As Double = Val(NormalizeNumericResponse(txtr1a.Text))
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit &
                           DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                           v1.ToString("#0.0000000000", Globalization.CultureInfo.InvariantCulture) & CSVdelimit &
                           LabelTemperature.Text & CSVdelimit & LabelHumidity.Text & CSVdelimit & statsCSV)

                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit &
                            DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                            v1.ToString("#0.0000000000", Globalization.CultureInfo.InvariantCulture) & CSVdelimit &
                            LabelTemperature.Text & CSVdelimit & LabelHumidity.Text & CSVdelimit & statsCSV
                Else
                    ' e-notation (numeric-only, from normalized value)
                    Dim v1 As Double = Val(NormalizeNumericResponse(txtr1a.Text))
                    Dim eStr1 As String = v1.ToString("0.0000000000E+00", Globalization.CultureInfo.InvariantCulture)

                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit &
                           DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                           eStr1 & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text & CSVdelimit & statsCSV)

                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit &
                            DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                            eStr1 & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text & CSVdelimit & statsCSV
                End If

            Else
                ' No Temp/Hum logging → use 0.0,0.0
                If (ENotationDecimal.Checked = True) Then
                    ' decimal
                    Dim v1 As Double = Val(NormalizeNumericResponse(txtr1a.Text))
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit &
                           DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                           v1.ToString("#0.0000000000", Globalization.CultureInfo.InvariantCulture) & CSVdelimit &
                           "0.0" & CSVdelimit & "0.0" & CSVdelimit & statsCSV)

                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit &
                            DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                            v1.ToString("#0.0000000000", Globalization.CultureInfo.InvariantCulture) & CSVdelimit &
                            "0.0" & CSVdelimit & "0.0" & CSVdelimit & statsCSV
                Else
                    ' e-notation
                    Dim v1 As Double = Val(NormalizeNumericResponse(txtr1a.Text))
                    Dim eStr1 As String = v1.ToString("0.0000000000E+00", Globalization.CultureInfo.InvariantCulture)

                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit &
                           DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                           eStr1 & CSVdelimit & "0.0" & CSVdelimit & "0.0" & CSVdelimit & statsCSV)

                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit &
                            DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                            eStr1 & CSVdelimit & "0.0" & CSVdelimit & "0.0" & CSVdelimit & statsCSV
                End If
            End If

        End If



        ' Device 2 updates
        If ((ButtonDev2Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop") And CheckboxEnableCSV.Checked = True And txtr2a.Text <> "" And Dev2GPIBActivity = True) Then

            If LabelTemperature.Text = "NaN" Or LabelTemperature.Text = "nil" Or LabelHumidity.Text = "NaN" Or LabelHumidity.Text = "nil" Then
                LabelTemperature.Text = "0.0"
                LabelHumidity.Text = "0.0"
            End If

            If (TempHumLogs.Checked = True And TempHumConnected = True) Then
                If (ENotationDecimal.Checked = True) Then
                    ' decimal
                    Dim v2 As Double = Val(NormalizeNumericResponse(txtr2a.Text))
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit &
                           DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                           v2.ToString("#0.0000000000", Globalization.CultureInfo.InvariantCulture) & CSVdelimit &
                           LabelTemperature.Text & CSVdelimit & LabelHumidity.Text & CSVdelimit & statsCSV)

                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit &
                            DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                            v2.ToString("#0.0000000000", Globalization.CultureInfo.InvariantCulture) & CSVdelimit &
                            LabelTemperature.Text & CSVdelimit & LabelHumidity.Text & CSVdelimit & statsCSV
                Else
                    ' e-notation (numeric-only, from normalized value)
                    Dim v2 As Double = Val(NormalizeNumericResponse(txtr2a.Text))
                    Dim eStr2 As String = v2.ToString("0.0000000000E+00", Globalization.CultureInfo.InvariantCulture)

                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit &
                           DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                           eStr2 & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text & CSVdelimit & statsCSV)

                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit &
                            DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                            eStr2 & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text & CSVdelimit & statsCSV
                End If
            Else
                ' No Temp/Hum logging → 0.0,0.0
                If (ENotationDecimal.Checked = True) Then
                    ' decimal
                    Dim v2 As Double = Val(NormalizeNumericResponse(txtr2a.Text))
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit &
                           DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                           v2.ToString("#0.0000000000", Globalization.CultureInfo.InvariantCulture) & CSVdelimit &
                           "0.0" & CSVdelimit & "0.0" & CSVdelimit & statsCSV)

                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit &
                            DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                            v2.ToString("#0.0000000000", Globalization.CultureInfo.InvariantCulture) & CSVdelimit &
                            "0.0" & CSVdelimit & "0.0" & CSVdelimit & statsCSV
                Else
                    ' e-notation
                    Dim v2 As Double = Val(NormalizeNumericResponse(txtr2a.Text))
                    Dim eStr2 As String = v2.ToString("0.0000000000E+00", Globalization.CultureInfo.InvariantCulture)

                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit &
                           DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                           eStr2 & CSVdelimit & "0.0" & CSVdelimit & "0.0" & CSVdelimit & statsCSV)

                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit &
                            DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit &
                            eStr2 & CSVdelimit & "0.0" & CSVdelimit & "0.0" & CSVdelimit & statsCSV
                End If
            End If

        End If



        ' CSV entries per device - Dual logging
        If (ButtonDev12Run.Text = "Stop" And CheckboxEnableCSV.Checked = True) Then
            CSVcounts.Text = Format(Math.Floor(IndexCount), "0")     ' Update entries in CSV per device
            'IndexCount = IndexCount + 0.5
            IndexCount += 0.5
        End If

        ' CSV entries per device - Dev 1 logging only
        If (ButtonDev1Run.Text = "Stop" And CheckboxEnableCSV.Checked = True) Then
            CSVcounts.Text = IndexCount     ' Update entries in CSV per device
            'IndexCount = IndexCount + 1
            IndexCount += 1
        End If

        ' CSV entries per device - Dev 2 logging only
        If (ButtonDev2Run.Text = "Stop" And CheckboxEnableCSV.Checked = True) Then
            CSVcounts.Text = IndexCount     ' Update entries in CSV per device
            'IndexCount = IndexCount + 1
            IndexCount += 1
        End If


        ' Manual limit on CSV entries if its running
        If ((Val(CSVcounts.Text) >= Val(CSVEntryLimit.Text)) And CheckboxCSVlimit.Checked = True And CheckboxEnableCSV.Checked = True) Then
            CheckboxEnableCSV.Checked = False
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to CSV Ended")
        End If

        ' Dev 1 Manual limit on CSV entries MINS if its running
        ' If CSV Entries >= (Entry Mins * 60) / Dev1SampleRate Then......
        If ((Val(CSVcounts.Text) >= ((Val(CSVEntryLimitMins.Text) * 60) / Val(Dev1SampleRate.Text))) And CheckboxCSVlimitMins.Checked = True And CheckboxEnableCSV.Checked = True And ButtonDev1Run.Text = "Stop") Then
            CheckboxEnableCSV.Checked = False
            'CSVfilename.ReadOnly = True
            CSVfilepath.ReadOnly = True
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to CSV Ended")
        End If

        ' Dev 2 Manual limit on CSV entries MINS if its running
        ' If CSV Entries >= (Entry Mins * 60) / Dev2SampleRate Then......
        If ((Val(CSVcounts.Text) >= ((Val(CSVEntryLimitMins.Text) * 60) / Val(Dev2SampleRate.Text))) And CheckboxCSVlimitMins.Checked = True And CheckboxEnableCSV.Checked = True And ButtonDev2Run.Text = "Stop") Then
            CheckboxEnableCSV.Checked = False
            'CSVfilename.ReadOnly = True
            CSVfilepath.ReadOnly = True
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to CSV Ended")
        End If

        ' Dev 1&2 Manual limit on CSV entries MINS if its running
        ' If CSV Entries >= (Entry Mins * 60) / Dev12SampleRate Then......
        If ((Val(CSVcounts.Text) >= ((Val(CSVEntryLimitMins.Text) * 60) / Val(Dev12SampleRate.Text))) And CheckboxCSVlimitMins.Checked = True And CheckboxEnableCSV.Checked = True And ButtonDev12Run.Text = "Stop") Then
            CheckboxEnableCSV.Checked = False
            CSVfilename.ReadOnly = True
            'CSVfilepath.ReadOnly = True
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to CSV Ended")
        End If


        file.Close()

    End Sub


    Private Sub ButtonExportCSV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExportCSV.Click

        Dim DateandTime As String
        DateandTime = DateTime.Now
        DateandTime = DateandTime.Replace(" ", "_")
        DateandTime = DateandTime.Replace("/", "_")
        DateandTime = DateandTime.Replace(":", "_")

        If (CSVfilepath.Text <> "" And CSVfilename.Text <> "") Then
            My.Computer.FileSystem.CopyFile((CSVfilepath.Text & "\" & CSVfilename.Text), (CSVfilepath.Text & "\" & DateandTime & "_" & TextFilenameAppend.Text & "_" & CSVfilename.Text), Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
        End If

        Dialog2.Warning1 = "Exported file has been created:"
        Dialog2.Warning2 = "File = " & DateandTime & "_" & TextFilenameAppend.Text & "_" & CSVfilename.Text
        Dialog2.Warning3 = ""
        Dialog2.ShowDialog(Me)  ' this method positions centre of parent form, and requires to hit OK to return back to parent

    End Sub


    Private Sub ResetCSV_Click(sender As Object, e As EventArgs) Handles ResetCSV.Click

        'If CheckboxEnableCSV.Checked = False Then

        Dim CSVpath As String = CSVfilepath.Text & "\" & CSVfilename.Text
        System.IO.File.WriteAllText(CSVpath, "")
        IndexCount = 1  ' reset index count for CSV file
        CSVcounts.Text = "0"     ' Update entries in CSV per device
        CSVsize.Text = "0"
        AllowMetaDataWrite = True
        Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " CSV file contents cleared")
        CSVwrite.Text = ""

        My.Settings.data11 = CSVfilename.Text
        My.Settings.data12 = CSVfilepath.Text

        'End If

    End Sub


    Private Sub CheckboxEnableCSV_CheckedChanged(sender As Object, e As EventArgs) Handles CheckboxEnableCSV.CheckedChanged

        If CheckboxEnableCSV.Checked = True Then
            CSVcounts.Text = "0"
            IndexCount = 1
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to CSV Enabled")
            'CSVfilename.ReadOnly = True
            'CSVfilepath.ReadOnly = True
            CSVfilestarting = True

            'ResetCSV.Enabled = False
            CSVwrite.Text = ""
        Else
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to CSV Disabled")
            'CSVfilename.ReadOnly = False
            'CSVfilepath.ReadOnly = False

            CSVfilestarting = False

            'ResetCSV.Enabled = True
        End If

    End Sub


    Private Sub CheckboxCSVlimit_CheckedChanged(sender As Object, e As EventArgs) Handles CheckboxCSVlimit.CheckedChanged

        If (CheckboxCSVlimit.Checked = True) Then
            CheckboxCSVlimitMins.Checked = False
        End If

    End Sub


    Private Sub CheckboxCSVlimitMins_CheckedChanged(sender As Object, e As EventArgs) Handles CheckboxCSVlimitMins.CheckedChanged

        If (CheckboxCSVlimitMins.Checked = True) Then
            CheckboxCSVlimit.Checked = False
        End If

    End Sub


    Private Sub CheckboxEnableLOG_CheckedChanged(sender As Object, e As EventArgs) Handles CheckboxEnableLOG.CheckedChanged

        If CheckboxEnableLOG.Checked = True Then
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to LOG Enabled")
        Else
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to LOG Stopped")
        End If

    End Sub


    Private Sub LogFileMetadata_TextChanged(sender As Object, e As EventArgs) Handles LogFileMetadata.TextChanged

        ' Limit number of lines in the textbox to 3

        ' Split the text into lines
        Dim lines() As String = LogFileMetadata.Text.Split({vbCrLf, vbLf}, StringSplitOptions.None)

        ' Check if the number of lines exceeds 3
        If lines.Length > 3 Then
            ' Remove the last line if there are more than 3 lines
            Dim newText As String = String.Join(vbCrLf, lines.Take(3))
            LogFileMetadata.Text = newText
            LogFileMetadata.Select(LogFileMetadata.TextLength, 0) ' Move the cursor to the end of the text
        End If
    End Sub


    Private Sub CSVfilename_TextChanged(sender As Object, e As EventArgs) Handles CSVfilename.TextChanged

        ' filename has changed so allow metadata writing
        AllowMetaDataWrite = True

    End Sub


    Private Sub StartCSVLog_Click(sender As Object, e As EventArgs) Handles StartCSVLog.Click

        'Check that at least one logging option is enabled
        If CheckboxEnableLOG.Checked = False And CheckboxEnableCSV.Checked = False Then
            MsgBox("You must check 'Enable DATA LOG' and/or 'Enable CSV'")
            Exit Sub
        End If

        StartCSVLogClicked = Not StartCSVLogClicked

        If StartCSVLogClicked = True Then
            StartCSVLog.Text = "Stop Log/CSV"
            CSVfilename.ReadOnly = True
            CSVfilepath.ReadOnly = True
        Else
            StartCSVLog.Text = "Start Log/CSV"
            CSVfilename.ReadOnly = False
            CSVfilepath.ReadOnly = False
        End If

    End Sub

End Class
