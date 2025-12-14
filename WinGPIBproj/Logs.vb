
' Data log, event log, CSV generation and main app chart control

'Imports System.IO
'Imports System.IO.File
'Imports System.Threading
'Imports Microsoft.Office.Interop.Word

Partial Class Formtest

    Dim file As System.IO.StreamWriter
    Dim filepath As String
    Dim filename As String

    Dim inst_value1FChart As Double
    Dim inst_value2FChart As Double
    Dim inst_value3FChart As Double
    Dim txtr1achart As String
    Dim txtr2achart As String
    Dim txtr3achart As String

    Dim inst_value1FChartMin As Double = 0
    Dim inst_value1FChartMax As Double = 10
    Dim inst_value2FChartMin As Double = 0
    Dim inst_value2FChartMax As Double = 10

    Dim inst_value1FChartMinActual As Double = 0
    Dim inst_value1FChartMaxActual As Double = 10
    Dim inst_value2FChartMinActual As Double = 0
    Dim inst_value2FChartMaxActual As Double = 10

    Dim inst_value1FChartMinRecordedDisplay As Double
    Dim inst_value1FChartMaxRecordedDisplay As Double
    Dim inst_value2FChartMinRecordedDisplay As Double
    Dim inst_value2FChartMaxRecordedDisplay As Double

    Dim inst_TemperatureChartMinRecordedDisplay As Double
    Dim inst_TemperatureChartMaxRecordedDisplay As Double

    Dim AutoScale1Flag As Boolean = False
    Dim AutoScale2Flag As Boolean = False
    Dim Scalerange As Double
    Dim rangeDev1 As Double
    Dim CSVupdate As Boolean = False

    Dim IndexCount As Double = 1      ' for CSV file

    Dim AutoScaleFirst1Flag As Boolean = True
    Dim AutoScaleFirst2Flag As Boolean = True

    Dim value1F As Double
    Dim value2F As Double
    Dim valueTempF As Double

    Dim CSVlength As Long

    Dim RunChart As Boolean = False
    Dim ChartPoints1 As Integer = 0
    Dim ChartPoints2 As Integer = 0

    Dim CSVfilestarting As Boolean = False
    Dim AllowMetaDataWrite As Boolean = True
    Dim CSVfilenameprevious As String = "somedummyfilename"

    ' Create a list to store log entries
    Private EventlogEntries As New List(Of String)

    ' === NEW: lightweight rolling average state ===
    Private q1 As New Queue(Of Double)  ' Dev1 samples
    Private q2 As New Queue(Of Double)  ' Dev2 samples
    Private sum1 As Double = 0          ' Dev1 running sum
    Private sum2 As Double = 0          ' Dev2 running sum


    ' Expect these controls to exist on the form:
    ' CheckBoxAvgEnable  (checkbox to enable averaging)
    ' TextBoxAvgWindow   (N value for rolling window)

    ' === NEW: returns averaged value if enabled, else raw ===
    Private Function AvgVal(v As Double, dev As Integer) As Double
        If Not CheckBoxAvgEnable.Checked Then Return v

        Dim n As Integer
        If Not Integer.TryParse(TextBoxAvgWindow.Text, n) OrElse n <= 1 Then Return v

        If dev = 1 Then
            q1.Enqueue(v) : sum1 += v
            If q1.Count > n Then sum1 -= q1.Dequeue()
            Return sum1 / q1.Count
        Else
            q2.Enqueue(v) : sum2 += v
            If q2.Count > n Then sum2 -= q2.Dequeue()
            Return sum2 / q2.Count
        End If
    End Function



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


    Private Sub LiveChart()

        ' Chart 1 - Device 1 only
        If (EnableChart1.Checked = True And EnableChart2.Checked = False And RunChart = True And Dev1GPIBActivity = True) Then

            inst_value1FChart = CDbl(Val(txtr1a.Text))
            inst_value1FChart = AvgVal(inst_value1FChart, 1)
            Dev1ChartValue.Text = Format(inst_value1FChart, "0.#########")
            txtr1achart = Format(inst_value1FChart, "#0.00000000")

            ' plot to chart Device 1
            If DisableRollingChart.Checked = False Then
                Chart1.Series(0).Points.AddY(txtr1achart)
                If Chart1.Series(0).Points.Count > XaxisPoints.Text Then  'sliding graph: last n points
                    Chart1.Series(0).Points.RemoveAt(0)
                End If
            Else
                Chart1.Series(0).Points.AddY(txtr1achart)
            End If

            ' Chart 3 - Temperature
            If (EnableChart3.Checked = True And RunChart = True) Then
                ' set up max and min for temperature
                If Val(LCTempMax.Text) > Val(LCTempMin.Text) Then
                    UpdateChartTemperatureYAxisMinMaxInterval()
                End If

                inst_value3FChart = gCurrTemp
                'inst_value3FChart = inst_value3FChart + Val(TempOffset.Text)    ' integrate offset
                inst_value3FChart += Val(TempOffset.Text)    ' integrate offset
                txtr3achart = Format(inst_value3FChart, "#0.00000000")

                ' plot to chart

                If DisableRollingChart.Checked = False Then
                    Chart1.Series(2).Points.AddY(txtr3achart)
                    If Chart1.Series(2).Points.Count > XaxisPoints.Text Then  'sliding graph: last n points
                        Chart1.Series(2).Points.RemoveAt(0)
                    End If
                Else
                    Chart1.Series(2).Points.AddY(txtr3achart)
                End If

                ' Temp - record min & max for display (resettable)
                Resetmaxdiffrecorded_temp()
            End If

            If (EnableChart3.Checked = False And RunChart = True) Then      ' dummy data so chart vertical data can align if temperature is checked later
                Chart1.Series(2).Points.AddY(0.0)
            End If

        End If


        ' Chart 2 - Device 2 only
        If (EnableChart2.Checked = True And EnableChart1.Checked = False And RunChart = True And Dev2GPIBActivity = True) Then

            inst_value2FChart = CDbl(Val(txtr2a.Text))
            inst_value2FChart = AvgVal(inst_value2FChart, 2)
            Dev2ChartValue.Text = Format(inst_value2FChart, "0.#########")
            txtr2achart = Format(inst_value2FChart, "#0.00000000")

            ' plot to chart Device 2

            If DisableRollingChart.Checked = False Then
                Chart1.Series(1).Points.AddY(txtr2achart)
                If Chart1.Series(1).Points.Count > XaxisPoints.Text Then  'sliding graph: last n points
                    Chart1.Series(1).Points.RemoveAt(0)
                End If
            Else
                Chart1.Series(1).Points.AddY(txtr2achart)
            End If

            ' Chart 3 - Temperature
            If (EnableChart3.Checked = True And RunChart = True) Then
                ' set up max and min for temperature
                If Val(LCTempMax.Text) > Val(LCTempMin.Text) Then
                    UpdateChartTemperatureYAxisMinMaxInterval()
                End If

                inst_value3FChart = gCurrTemp
                'inst_value3FChart = inst_value3FChart + Val(TempOffset.Text)    ' integrate offset
                inst_value3FChart += Val(TempOffset.Text)   ' integrate offset
                txtr3achart = Format(inst_value3FChart, "#0.00000000")

                ' plot to chart

                If DisableRollingChart.Checked = False Then
                    Chart1.Series(2).Points.AddY(txtr3achart)
                    If Chart1.Series(2).Points.Count > XaxisPoints.Text Then  'sliding graph: last n points
                        Chart1.Series(2).Points.RemoveAt(0)
                    End If
                Else
                    Chart1.Series(2).Points.AddY(txtr3achart)
                End If


                ' Temp - record min & max for display (resettable)
                Resetmaxdiffrecorded_temp()
            End If

            If (EnableChart3.Checked = False And RunChart = True) Then      ' dummy data so chart vertical data can align if temperature is checked later
                Chart1.Series(2).Points.AddY(0.0)
            End If

        End If


        ' Chart 1 & 2 - Device 1 & Device 2
        If (EnableChart1.Checked = True And EnableChart2.Checked = True And RunChart = True And Dev2GPIBActivity = True) Then

            inst_value1FChart = CDbl(Val(txtr1a.Text))
            inst_value1FChart = AvgVal(inst_value1FChart, 1)
            Dev1ChartValue.Text = Format(inst_value1FChart, "0.#########")
            inst_value2FChart = CDbl(Val(txtr2a.Text))
            inst_value2FChart = AvgVal(inst_value2FChart, 2)
            Dev2ChartValue.Text = Format(inst_value2FChart, "0.#########")
            txtr1achart = Format(inst_value1FChart, "#0.00000000")
            txtr2achart = Format(inst_value2FChart, "#0.00000000")

            ' plot to chart Device 1
            If DisableRollingChart.Checked = False Then
                Chart1.Series(0).Points.AddY(txtr1achart)
                Chart1.Series(1).Points.AddY(txtr2achart)
                If Chart1.Series(0).Points.Count > XaxisPoints.Text Then  'sliding graph: last n points
                    Chart1.Series(0).Points.RemoveAt(0)
                    Chart1.Series(1).Points.RemoveAt(0)
                End If
            Else
                Chart1.Series(0).Points.AddY(txtr1achart)
                Chart1.Series(1).Points.AddY(txtr2achart)
            End If


            ' Chart 3 - Temperature
            If (EnableChart3.Checked = True And RunChart = True) Then
                ' set up max and min for temperature
                If Val(LCTempMax.Text) > Val(LCTempMin.Text) Then
                    UpdateChartTemperatureYAxisMinMaxInterval()
                End If

                inst_value3FChart = gCurrTemp
                'inst_value3FChart = inst_value3FChart + Val(TempOffset.Text)    ' integrate offset
                inst_value3FChart += Val(TempOffset.Text)   ' integrate offset
                txtr3achart = Format(inst_value3FChart, "#0.00000000")

                ' plot to chart

                If DisableRollingChart.Checked = False Then
                    Chart1.Series(2).Points.AddY(txtr3achart)
                    If Chart1.Series(2).Points.Count > XaxisPoints.Text Then  'sliding graph: last n points
                        Chart1.Series(2).Points.RemoveAt(0)
                    End If
                Else
                    Chart1.Series(2).Points.AddY(txtr3achart)
                End If


                ' Temp - record min & max for display (resettable)
                Resetmaxdiffrecorded_temp()
            End If

            If (EnableChart3.Checked = False And RunChart = True) Then      ' dummy data so chart vertical data can align if temperature is checked later
                Chart1.Series(2).Points.AddY(0.0)
            End If

        End If


        ' ============================
        ' Fixed X window so trace appears
        ' at the right and scrolls left
        ' ============================
        If DisableRollingChart.Checked = False Then

            If Chart1.ChartAreas.Count > 0 Then
                Dim ca = Chart1.ChartAreas(0)

                ' How many points wide should the visible window be?
                Dim windowN As Integer
                If Not Integer.TryParse(XaxisPoints.Text, windowN) OrElse windowN < 2 Then
                    windowN = 100
                End If

                ' Pick the first series that actually has data
                Dim sRef As DataVisualization.Charting.Series = Nothing
                For si As Integer = 0 To Chart1.Series.Count - 1
                    If Chart1.Series(si).Points.Count > 0 Then
                        sRef = Chart1.Series(si)
                        Exit For
                    End If
                Next

                If sRef IsNot Nothing Then
                    ' Use the point index as X (0,1,2,...) instead of XValue
                    Dim lastIndex As Integer = sRef.Points.Count - 1
                    Dim window As Integer = windowN - 1

                    Dim xmin As Double = lastIndex - window
                    Dim xmax As Double = lastIndex

                    ca.AxisX.Minimum = xmin
                    ca.AxisX.Maximum = xmax

                    Dim domain As Double = window
                    If domain <= 0 Then domain = 10.0R
                    ca.AxisX.Interval = domain / 10.0R
                Else
                    ' No points yet – let chart decide
                    ca.AxisX.Minimum = Double.NaN
                    ca.AxisX.Maximum = Double.NaN
                    ca.AxisX.Interval = Double.NaN
                End If
            End If

        Else
            ' Rolling disabled – let chart auto-manage X axis
            If Chart1.ChartAreas.Count > 0 Then
                Dim ca = Chart1.ChartAreas(0)
                ca.AxisX.Minimum = Double.NaN
                ca.AxisX.Maximum = Double.NaN
                ca.AxisX.Interval = Double.NaN
            End If
        End If

    End Sub



    ' Create a list to store log entries
    Private logEntries As New List(Of String)


    Private Sub LogData(ByVal str As String)
        ' Updates the display ListBoxData log display and the Live Chart.
        ' Called from LOGdisplay()

        ' Check if logging is enabled
        If CheckboxEnableLOG.Checked Then
            ' Add the new log entry to the logEntries list
            logEntries.Add(str)

            ' Ensure the logEntries list contains no more than 30 items
            If logEntries.Count > 31 Then
                logEntries.RemoveAt(0) ' Remove the oldest entry from logEntries
            End If

            ' Clear the ListBox and add updated entries
            ListBoxData.Items.Clear()
            For Each entry As String In logEntries
                ListBoxData.Items.Add(entry)
            Next
        End If
    End Sub


    Private Sub ClearLOGdisp_Click(sender As Object, e As EventArgs) Handles ClearLOGdisp.Click
        ' Clear the log data from both display and the logEntries list
        ListBoxData.Items.Clear()
        logEntries.Clear()
        LogData(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " LOG Cleared")
    End Sub


    Private Sub LOGdisplay()

        'Dev 1 Log display
        If ((ButtonDev1Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop") And CheckboxEnableLOG.Checked = True And txtr1a.Text <> "" And Dev1GPIBActivity = True) Then

            TestLen = Len(txtr1a.Text)

            If (ENotationDecimal.Checked = True And TestLen > 4) Then

                txtr1alogged = Format(CDbl(Val(txtr1a.Text)), "#0.0000000000")

                Dim txtname1fixed As String = txtname1.Text.PadRight(10)
                Dim txtr1aloggedfixed As String = txtr1alogged.PadRight(14)

                If (TempHumLogs.Checked = True) Then
                    If LabelTemperature.Text = "NaN" Or LabelTemperature.Text = "nil" Or LabelHumidity.Text = "NaN" Or LabelHumidity.Text = "nil" Then
                        LabelTemperature.Text = "0.0"
                        LabelHumidity.Text = "0.0"
                    End If

                    Dim valueTemp As Double = CDbl(LabelTemperature.Text())             ' Convert the string to a double
                    Dim formattedValueTemp As String = valueTemp.ToString("0.00")       ' Format the value with two decimal places (XX.XX format)
                    Dim valueHum As Double = CDbl(LabelHumidity.Text())                 ' Convert the string to a double
                    Dim formattedValueHum As String = valueHum.ToString("0.00")         ' Format the value with two decimal places (XX.XX format)

                    LogData(txtname1fixed & "  " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "   " & txtr1aloggedfixed & "  " & formattedValueTemp & "   " & formattedValueHum)
                Else
                    LogData(txtname1fixed & "  " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "   " & txtr1aloggedfixed & "  " & "0.0" & "   " & "0.0")
                End If

            Else

                Dim txtname1fixed As String = txtname1.Text.PadRight(10)
                Dim txtr1afixed As String = txtr1a.Text.PadRight(14)

                If (TempHumLogs.Checked = True) Then
                    If LabelTemperature.Text = "NaN" Or LabelTemperature.Text = "nil" Or LabelHumidity.Text = "NaN" Or LabelHumidity.Text = "nil" Then
                        LabelTemperature.Text = "0.0"
                        LabelHumidity.Text = "0.0"
                    End If

                    Dim valueTemp As Double = CDbl(LabelTemperature.Text())             ' Convert the string to a double
                    Dim formattedValueTemp As String = valueTemp.ToString("0.00")       ' Format the value with two decimal places (XX.XX format)
                    Dim valueHum As Double = CDbl(LabelHumidity.Text())                 ' Convert the string to a double
                    Dim formattedValueHum As String = valueHum.ToString("0.00")         ' Format the value with two decimal places (XX.XX format)

                    LogData(txtname1fixed & "  " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "   " & txtr1afixed & "  " & formattedValueTemp & "   " & formattedValueHum)
                Else
                    LogData(txtname1fixed & "  " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "   " & txtr1afixed & "  " & "0.0" & "   " & "0.0")
                End If

            End If


        End If




        'Dev 2 Log display
        If ((ButtonDev2Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop") And CheckboxEnableLOG.Checked = True And txtr2a.Text <> "" And Dev2GPIBActivity = True) Then

            TestLen = Len(txtr2a.Text)

            If (ENotationDecimal.Checked = True And TestLen > 4) Then

                txtr2alogged = Format(CDbl(Val(txtr2a.Text)), "#0.0000000000")

                Dim txtname2fixed As String = txtname2.Text.PadRight(10)
                Dim txtr2aloggedfixed As String = txtr2alogged.PadRight(14)

                If (TempHumLogs.Checked = True) Then
                    If LabelTemperature.Text = "NaN" Or LabelTemperature.Text = "nil" Or LabelHumidity.Text = "NaN" Or LabelHumidity.Text = "nil" Then
                        LabelTemperature.Text = "0.0"
                        LabelHumidity.Text = "0.0"
                    End If

                    Dim valueTemp As Double = CDbl(LabelTemperature.Text())             ' Convert the string to a double
                    Dim formattedValueTemp As String = valueTemp.ToString("0.00")       ' Format the value with two decimal places (XX.XX format)
                    Dim valueHum As Double = CDbl(LabelHumidity.Text())                 ' Convert the string to a double
                    Dim formattedValueHum As String = valueHum.ToString("0.00")         ' Format the value with two decimal places (XX.XX format)

                    LogData(txtname2fixed & "  " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "   " & txtr2aloggedfixed & "  " & formattedValueTemp & "   " & formattedValueHum)
                Else
                    LogData(txtname2fixed & "  " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "   " & txtr2aloggedfixed & "  " & "0.0" & "   " & "0.0")
                End If

            Else

                Dim txtname2fixed As String = txtname2.Text.PadRight(10)
                Dim txtr2afixed As String = txtr2a.Text.PadRight(14)

                If (TempHumLogs.Checked = True) Then
                    If LabelTemperature.Text = "NaN" Or LabelTemperature.Text = "nil" Or LabelHumidity.Text = "NaN" Or LabelHumidity.Text = "nil" Then
                        LabelTemperature.Text = "0.0"
                        LabelHumidity.Text = "0.0"
                    End If

                    Dim valueTemp As Double = CDbl(LabelTemperature.Text())             ' Convert the string to a double
                    Dim formattedValueTemp As String = valueTemp.ToString("0.00")       ' Format the value with two decimal places (XX.XX format)
                    Dim valueHum As Double = CDbl(LabelHumidity.Text())                 ' Convert the string to a double
                    Dim formattedValueHum As String = valueHum.ToString("0.00")         ' Format the value with two decimal places (XX.XX format)

                    LogData(txtname2fixed & "  " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "   " & txtr2afixed & "  " & formattedValueTemp & "   " & formattedValueHum)
                Else
                    LogData(txtname2fixed & "  " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") & "   " & txtr2afixed & "  " & "0.0" & "   " & "0.0")
                End If

            End If


        End If


    End Sub


    Private Sub ButtonDiffRecorded1Reset_Click(sender As Object, e As EventArgs) Handles ButtonDiffRecorded1Reset.Click

        ' Reset Device 1 Max Diff Recorded
        inst_value1FChartMinRecordedDisplay = inst_value1FChart
        inst_value1FChartMaxRecordedDisplay = inst_value1FChart

        Resetmaxdiffrecorded_value1()

    End Sub


    Private Sub ButtonDiffRecorded2Reset_Click(sender As Object, e As EventArgs) Handles ButtonDiffRecorded2Reset.Click

        ' Reset Device 2 Max Diff Recorded
        inst_value2FChartMinRecordedDisplay = inst_value2FChart
        inst_value2FChartMaxRecordedDisplay = inst_value2FChart

        Resetmaxdiffrecorded_value2()

    End Sub


    Private Sub ButtonDiffRecordedTempReset_Click(sender As Object, e As EventArgs) Handles ButtonDiffRecordedTempReset.Click

        ' Reset Temperature Max Diff Recorded
        inst_TemperatureChartMinRecordedDisplay = inst_value3FChart
        inst_TemperatureChartMaxRecordedDisplay = inst_value3FChart
        Resetmaxdiffrecorded_temp()

    End Sub


    Private Sub Resetmaxdiffrecorded_value1()

        ' Device 1 - record min & max for display (resettable)
        If (inst_value1FChart > inst_value1FChartMaxRecordedDisplay) Then
            inst_value1FChartMaxRecordedDisplay = inst_value1FChart
        End If
        If (inst_value1FChart < inst_value1FChartMinRecordedDisplay) Then
            inst_value1FChartMinRecordedDisplay = inst_value1FChart
        End If
        ' invert it if diff will be negative
        If (inst_value1FChartMaxRecordedDisplay < inst_value1FChartMinRecordedDisplay) Then
            Dim temprecorded1 As Double = inst_value1FChartMaxRecordedDisplay
            inst_value1FChartMaxRecordedDisplay = inst_value1FChartMinRecordedDisplay
            inst_value1FChartMinRecordedDisplay = temprecorded1
        End If
        ' now value converted to decimal from E notation
        value1F = CDbl(Val(inst_value1FChartMaxRecordedDisplay - inst_value1FChartMinRecordedDisplay))
        inst_value1FDiffRecorded.Text = Format(value1F, "#0.00000000")

    End Sub


    Private Sub Resetmaxdiffrecorded_value2()

        ' Device 2 - record min & max for display (resettable)
        If (inst_value2FChart > inst_value2FChartMaxRecordedDisplay) Then
            inst_value2FChartMaxRecordedDisplay = inst_value2FChart
        End If
        If (inst_value2FChart < inst_value2FChartMinRecordedDisplay) Then
            inst_value2FChartMinRecordedDisplay = inst_value2FChart
        End If
        ' invert it if diff will be negative
        If (inst_value2FChartMaxRecordedDisplay < inst_value2FChartMinRecordedDisplay) Then
            Dim temprecorded2 As Double = inst_value2FChartMaxRecordedDisplay
            inst_value2FChartMaxRecordedDisplay = inst_value2FChartMinRecordedDisplay
            inst_value2FChartMinRecordedDisplay = temprecorded2
        End If
        ' now value converted to decimal from E notation
        value2F = CDbl(Val(inst_value2FChartMaxRecordedDisplay - inst_value2FChartMinRecordedDisplay))
        inst_value2FDiffRecorded.Text = Format(value2F, "#0.00000000")

    End Sub


    Private Sub Resetmaxdiffrecorded_temp()

        ' Temp - record min & max for display (resettable)
        If (inst_value3FChart > inst_TemperatureChartMaxRecordedDisplay) Then
            inst_TemperatureChartMaxRecordedDisplay = inst_value3FChart
        End If
        If (inst_value3FChart < inst_TemperatureChartMinRecordedDisplay) Then
            inst_TemperatureChartMinRecordedDisplay = inst_value3FChart
        End If
        ' invert it if diff will be negative
        If (inst_TemperatureChartMaxRecordedDisplay < inst_TemperatureChartMinRecordedDisplay) Then
            Dim temperaturerecorded1 As Double = inst_TemperatureChartMaxRecordedDisplay
            inst_TemperatureChartMaxRecordedDisplay = inst_TemperatureChartMinRecordedDisplay
            inst_TemperatureChartMinRecordedDisplay = temperaturerecorded1
        End If
        ' now value converted to decimal from E notation
        valueTempF = CDbl(Val(inst_TemperatureChartMaxRecordedDisplay - inst_TemperatureChartMinRecordedDisplay))
        TemperatureDiffRecorded.Text = Format(valueTempF, "#00.00")

    End Sub


    Private Sub ButtonClearChart_Click(sender As Object, e As EventArgs) Handles ButtonClearChart.Click

        Chart1.Visible = False
        StartChartMessage.Visible = True

        ' Clear charts
        Chart1.Series(0).Points.Clear()
        Chart1.Series(1).Points.Clear()
        Chart1.Series(2).Points.Clear()

        ' Reset saved max/min values for auto-scale
        inst_value1FChartMax = inst_value1FChart
        inst_value1FChartMin = inst_value1FChart - 0.0000000001
        inst_value2FChartMax = inst_value2FChart
        inst_value2FChartMin = inst_value2FChart - 0.0000000001

        ' Reset saved max/min for recorded display
        inst_value1FChartMinRecordedDisplay = 0
        inst_value1FChartMaxRecordedDisplay = 0
        inst_value2FChartMinRecordedDisplay = 0
        inst_value2FChartMaxRecordedDisplay = 0
        inst_TemperatureChartMaxRecordedDisplay = 0
        inst_TemperatureChartMinRecordedDisplay = 0

        ChartPoints1 = 0
        LabelChartPoints1.Text = "0"
        ChartPoints2 = 0
        LabelChartPoints2.Text = "0"

        q1.Clear() : q2.Clear() : sum1 = 0 : sum2 = 0

        RunChart = False

        YaxisDiff.Text = "0"

        LabeChartMinutes.Text = "0hrs 00mins 00secs"

    End Sub


    Private Sub ButtonPauseChart_Click(sender As Object, e As EventArgs) Handles ButtonPauseChart.Click

        RunChart = Not RunChart

        ' Chart currently running and user just hit pause
        If (RunChart = False) Then
            ButtonPauseChart.Text = "Start Chart"
            ButtonClearChart.Enabled = True
        End If

        ' Chart currently paused and user just hit run 
        If (RunChart = True) Then
            ButtonPauseChart.Text = "Pause Chart"
            ButtonClearChart.Enabled = False

            ' Set Y-scale of chart based on Min/Max ensuring at least 1DP and number of DP's set in Min/Max
            ' Parse values from textboxes
            Dim minValue As Double
            Dim maxValue As Double
            If Double.TryParse(Dev1Min.Text, minValue) AndAlso Double.TryParse(Dev1Max.Text, maxValue) And EnableAutoYChart1.Checked = False Then
                ' Determine the number of decimal places based on maximum precision
                Dim decimalPlaces As Integer = Math.Max(GetMaxPrecision(minValue, maxValue), 1)

                UpdateChartYAxisMinMaxInterval()

                ' Set the number of decimal places for Y-axis labels
                Chart1.ChartAreas(0).AxisY.LabelStyle.Format = "F8"

                ' Disable auto-fit to prevent automatic scaling
                Chart1.ChartAreas(0).AxisY.IsLabelAutoFit = False

                ' Ensure that auto-fit is turned off to prevent automatic scaling
                Chart1.ChartAreas(0).AxisY.IsStartedFromZero = False
            End If

            Chart1.Visible = True
            StartChartMessage.Visible = False

        End If

    End Sub

    ' Function to determine the maximum precision (number of decimal places)
    Function GetMaxPrecision(ParamArray values As Double()) As Integer
        Dim maxPrecision As Integer = 0

        For Each value As Double In values
            Dim strValue As String = value.ToString("G") ' Use "G" format to avoid scientific notation
            Dim decimalIndex As Integer = strValue.IndexOf("."c)

            If decimalIndex <> -1 Then
                Dim currentPrecision As Integer = strValue.Length - decimalIndex - 1
                maxPrecision = Math.Max(maxPrecision, currentPrecision)
            End If
        Next

        Return maxPrecision
    End Function


    Private Sub DisableRollingChart_CheckedChanged(sender As Object, e As EventArgs) Handles DisableRollingChart.CheckedChanged

        ' if disable rilling chart is unchecked by user and the currentx-axis scale points is more than as set by the user then purge to currentx-axis scale points setting
        If DisableRollingChart.Checked = False And (ChartPoints1 > Val(XaxisPoints.Text) Or ChartPoints2 > Val(XaxisPoints.Text)) Then

            ' tba....not sure if need to do this. Only when it's been disabled and the scale points is over 2000 does the now scrolling graph move very slowly.....hmmmm
            ' saying that, it's maybe a nice feature!

        End If

    End Sub


    Private Sub ChartControl()

        ' This sub called by 100mS permanent timer4

        ' Chart sample counters
        Dim chartPoints1 As Integer = Chart1.Series(0).Points.Count
        Dim chartPoints2 As Integer = Chart1.Series(1).Points.Count

        Dim points As Integer = 0
        Dim sampleRateText As String = ""
        Dim pointsLabel As Label = Nothing

        ' Dev1 only
        If EnableChart1.Checked And Not EnableChart2.Checked And ButtonDev1Run.Text = "Stop" Then
            points = chartPoints1
            pointsLabel = LabelChartPoints1
            sampleRateText = Dev1SampleRate.Text
        End If

        ' Dev1 only but Dev1/2 channel active
        If EnableChart1.Checked And Not EnableChart2.Checked And ButtonDev12Run.Text = "Stop" Then
            points = chartPoints1
            pointsLabel = LabelChartPoints1
            sampleRateText = Dev1SampleRate.Text
        End If

        ' Dev2 only
        If Not EnableChart1.Checked And EnableChart2.Checked And ButtonDev2Run.Text = "Stop" Then
            points = chartPoints2
            pointsLabel = LabelChartPoints2
            sampleRateText = Dev2SampleRate.Text
        End If

        ' Dev2 only but Dev1/2 channel active
        If Not EnableChart1.Checked And EnableChart2.Checked And ButtonDev12Run.Text = "Stop" Then
            points = chartPoints2
            pointsLabel = LabelChartPoints2
            sampleRateText = Dev2SampleRate.Text
        End If

        ' Both charts active
        If EnableChart1.Checked And EnableChart2.Checked Then
            points = chartPoints1
            pointsLabel = LabelChartPoints1
            sampleRateText = Dev12SampleRate.Text
        End If

        ' Final calcs
        If sampleRateText <> "" AndAlso pointsLabel IsNot Nothing Then
            pointsLabel.Text = points.ToString()

            Dim totalSeconds As Integer = CInt(Val(sampleRateText) * points)
            Dim hours As Integer = totalSeconds \ 3600
            Dim minutes As Integer = (totalSeconds Mod 3600) \ 60
            Dim seconds As Integer = totalSeconds Mod 60

            LabeChartMinutes.Text = $"{hours}hrs {minutes:00}mins {seconds:00}secs"
        End If

        ' Chart controls
        If (EnableChart1.Checked = True Or EnableChart2.Checked = True Or EnableChart3.Checked = True) Then

            ' set up max and min for temperature
            If Val(LCTempMax.Text) > Val(LCTempMin.Text) Then
                UpdateChartTemperatureYAxisMinMaxInterval()
            End If


            ' Autoscale chart y-axis - Device 1 only
            If (EnableAutoYChart1.Checked = True And EnableChart1.Checked = True And EnableChart2.Checked = False) Then
                ' Check if 5samples have been received
                If Chart1.Series(0).Points.Count >= 5 Then
                    ' Autoscale the minimum and maximum of the Y-axis
                    Dim minValue1 As Double = Chart1.Series(0).Points.Min(Function(p) p.YValues(0))
                    Dim maxValue1 As Double = Chart1.Series(0).Points.Max(Function(p) p.YValues(0))

                    ' Ensure the difference is not zero to avoid crashes
                    Dim range As Double = maxValue1 - minValue1
                    If range = 0 Then
                        ' Use a small buffer based on the magnitude of the values
                        Dim buffer As Double = Math.Abs(minValue1) * 0.01 ' 1% of the magnitude as buffer
                        If buffer = 0 Then buffer = 0.00001 ' Fallback to a minimal buffer for very small values
                        minValue1 -= buffer
                        maxValue1 += buffer
                        range = maxValue1 - minValue1 ' Recalculate range
                    End If

                    Chart1.ChartAreas(0).AxisY.Minimum = minValue1
                    Chart1.ChartAreas(0).AxisY.Maximum = maxValue1

                    ' Customize the Y-axis interval to control the tick marks and labels
                    Chart1.ChartAreas(0).AxisY.Interval = (range) / 10 ' Adjust as needed

                    ' Prevent scientific notation (e-notation) on the Y-axis labels
                    Chart1.ChartAreas(0).AxisY.LabelStyle.Format = "#0.########"

                    ' Autoscale the Y-axis
                    Chart1.ChartAreas(0).RecalculateAxesScale()
                    YaxisDiff.Text = Format(range, "#0.00000000")
                Else
                    UpdateChartYAxisMinMaxInterval()
                    YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
                End If
            End If

            If (EnableAutoYChart1.Checked = False And EnableChart1.Checked = True And EnableChart2.Checked = False) Then
                UpdateChartYAxisMinMaxInterval()
                YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
            End If


            ' Autoscale chart y-axis - Device 2 only
            If (EnableAutoYChart1.Checked = True And EnableChart1.Checked = False And EnableChart2.Checked = True) Then
                ' Check if 5 samples have been received
                If Chart1.Series(1).Points.Count >= 5 Then
                    ' Autoscale the minimum and maximum of the Y-axis
                    Dim minValue2 As Double = Chart1.Series(1).Points.Min(Function(p) p.YValues(0))
                    Dim maxValue2 As Double = Chart1.Series(1).Points.Max(Function(p) p.YValues(0))

                    ' Ensure the difference is not zero to avoid crashes
                    Dim range As Double = maxValue2 - minValue2
                    If range = 0 Then
                        ' Use a small buffer based on the magnitude of the values
                        Dim buffer As Double = Math.Abs(minValue2) * 0.01 ' 1% of the magnitude as buffer
                        If buffer = 0 Then buffer = 0.00001 ' Fallback to a minimal buffer for very small values
                        minValue2 -= buffer
                        maxValue2 += buffer
                        range = maxValue2 - minValue2 ' Recalculate range
                    End If

                    Chart1.ChartAreas(0).AxisY.Minimum = minValue2
                    Chart1.ChartAreas(0).AxisY.Maximum = maxValue2

                    ' Customize the Y-axis interval to control the tick marks and labels
                    Chart1.ChartAreas(0).AxisY.Interval = (range) / 10 ' Adjust as needed

                    ' Prevent scientific notation (e-notation) on the Y-axis labels
                    Chart1.ChartAreas(0).AxisY.LabelStyle.Format = "#0.########"

                    ' Autoscale the Y-axis
                    Chart1.ChartAreas(0).RecalculateAxesScale()
                    YaxisDiff.Text = Format(range, "#0.00000000")
                Else
                    UpdateChartYAxisMinMaxInterval()
                    YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
                End If
            End If

            If (EnableAutoYChart1.Checked = False And EnableChart1.Checked = False And EnableChart2.Checked = True) Then
                UpdateChartYAxisMinMaxInterval()
                YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
            End If




            ' Autoscale chart y-axis - Device 1 & Device 2
            If (EnableAutoYChart1.Checked = True And EnableChart1.Checked = True And EnableChart2.Checked = True) Then
                ' Check if 5 samples have been received
                If (Chart1.Series(0).Points.Count >= 5 And Chart1.Series(1).Points.Count >= 5) Then

                    ' Get the minimum and maximum values from both series
                    Dim minValue1 As Double = Chart1.Series(0).Points.Min(Function(p) p.YValues(0))
                    Dim minValue2 As Double = Chart1.Series(1).Points.Min(Function(p) p.YValues(0))

                    Dim maxValue1 As Double = Chart1.Series(0).Points.Max(Function(p) p.YValues(0))
                    Dim maxValue2 As Double = Chart1.Series(1).Points.Max(Function(p) p.YValues(0))

                    ' Calculate the overall minimum and maximum values
                    Dim overallMin As Double = Math.Min(minValue1, minValue2)
                    Dim overallMax As Double = Math.Max(maxValue1, maxValue2)

                    ' Ensure the difference is not zero to avoid crashes
                    Dim range As Double = overallMax - overallMin
                    If range = 0 Then
                        ' Use a small buffer based on the magnitude of the values
                        Dim buffer As Double = Math.Abs(overallMin) * 0.01 ' 1% of the magnitude as buffer
                        If buffer = 0 Then buffer = 0.00001 ' Fallback to a minimal buffer for very small values
                        overallMin -= buffer
                        overallMax += buffer
                        range = overallMax - overallMin ' Recalculate range
                    End If

                    ' Set the minimum and maximum values for both Y-axes
                    Chart1.ChartAreas(0).AxisY.Minimum = overallMin
                    Chart1.ChartAreas(0).AxisY.Maximum = overallMax

                    ' Customize the Y-axis interval to control the tick marks and labels
                    Chart1.ChartAreas(0).AxisY.Interval = (range) / 10 ' Adjust as needed

                    ' Prevent scientific notation (e-notation) on the Y-axis labels
                    Chart1.ChartAreas(0).AxisY.LabelStyle.Format = "#0.########"

                    ' Recalculate the scale of the Y-axis
                    Chart1.ChartAreas(0).RecalculateAxesScale()

                    YaxisDiff.Text = Format(range, "#0.00000000")
                Else
                    UpdateChartYAxisMinMaxInterval()
                    YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
                End If
            End If

            If (EnableAutoYChart1.Checked = False And EnableChart1.Checked = True And EnableChart2.Checked = True) Then
                UpdateChartYAxisMinMaxInterval()
                YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
            End If


            If (XaxisPoints.Text < 100) Then
                XaxisPoints.Text = 100
            End If

        Else
            Dev1Min.ReadOnly = False
            Dev1Max.ReadOnly = False
            'ButtonClearChart.Enabled = True
            If (XaxisPoints.Text < 100) Then
                XaxisPoints.Text = 100
            End If

        End If


        If (EnableChart1.Checked = True) Then
            ' Device 1 - record min & max for display (resettable)
            Resetmaxdiffrecorded_value1()
        Else
            inst_value1FDiffRecorded.Text = "0.00000000"
        End If

        If (EnableChart2.Checked = True) Then
            ' Device 2 - record min & max for display (resettable)
            Resetmaxdiffrecorded_value2()
        Else
            inst_value2FDiffRecorded.Text = "0.00000000"
        End If


        If CheckBoxDevice1Hide.Checked = True Then
            Chart1.Series(0).Enabled = False
        Else
            Chart1.Series(0).Enabled = True
        End If

        If CheckBoxDevice2Hide.Checked = True Then
            Chart1.Series(1).Enabled = False
        Else
            Chart1.Series(1).Enabled = True
        End If

        If CheckBoxTempHide.Checked = True Then
            Chart1.Series(2).Enabled = False
        Else
            Chart1.Series(2).Enabled = True
        End If

    End Sub


    Private Sub CSVfile()

        ' Write data to CSV file, for both Dev1 & Dev2

        If (CheckboxEnableCSV.Checked = True) Then

            CSVfilename.ReadOnly = True
            CSVfilepath.ReadOnly = True
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
            CSVfilename.ReadOnly = False
            CSVfilepath.ReadOnly = False
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


        ' Device 1 updates
        If ((ButtonDev1Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop") And CheckboxEnableCSV.Checked = True And txtr1a.Text <> "" And Dev1GPIBActivity = True) Then

            If (TempHumLogs.Checked = True And TempHumConnected = True) Then
                If LabelTemperature.Text = "NaN" Or LabelTemperature.Text = "nil" Or LabelHumidity.Text = "NaN" Or LabelHumidity.Text = "nil" Then
                    LabelTemperature.Text = "0.0"
                    LabelHumidity.Text = "0.0"
                End If

                If (ENotationDecimal.Checked = True) Then
                    'decimal
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & Format(Val(txtr1a.Text), "#0.0000000000") & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text)
                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & Format(Val(txtr1a.Text), "#0.0000000000") & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text
                Else
                    'e-notation
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & txtr1a.Text & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text)
                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & txtr1a.Text & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text
                End If
            Else
                If (ENotationDecimal.Checked = True) Then
                    'decimal
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & Format(Val(txtr1a.Text), "#0.0000000000") & CSVdelimit & "0.0" & CSVdelimit & "0.0")
                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & Format(Val(txtr1a.Text), "#0.0000000000") & CSVdelimit & "0.0" & CSVdelimit & "0.0"
                Else
                    'e-notation
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & txtr1a.Text & CSVdelimit & "0.0" & CSVdelimit & "0.0")
                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname1.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & txtr1a.Text & CSVdelimit & "0.0" & CSVdelimit & "0.0"
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
                    'decimal
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & Format(Val(txtr2a.Text), "#0.0000000000") & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text)
                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & Format(Val(txtr2a.Text), "#0.0000000000") & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text
                Else
                    'e-notation
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & txtr2a.Text & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text)
                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & txtr2a.Text & CSVdelimit & LabelTemperature.Text & CSVdelimit & LabelHumidity.Text
                End If
            Else
                If (ENotationDecimal.Checked = True) Then
                    'decimal
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & Format(Val(txtr2a.Text), "#0.0000000000") & CSVdelimit & "0.0" & CSVdelimit & "0.0")
                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & Format(Val(txtr2a.Text), "#0.0000000000") & CSVdelimit & "0.0" & CSVdelimit & "0.0"
                Else
                    'e-notation
                    file.WriteLine(Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & txtr2a.Text & CSVdelimit & "0.0" & CSVdelimit & "0.0")
                    CSVwrite.Text = Format(Math.Floor(IndexCount), "0") & CSVdelimit & txtname2.Text & CSVdelimit & DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss") & CSVdelimit & txtr2a.Text & CSVdelimit & "0.0" & CSVdelimit & "0.0"
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
            CSVfilename.ReadOnly = True
            CSVfilepath.ReadOnly = True
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to CSV Ended")
        End If

        ' Dev 2 Manual limit on CSV entries MINS if its running
        ' If CSV Entries >= (Entry Mins * 60) / Dev2SampleRate Then......
        If ((Val(CSVcounts.Text) >= ((Val(CSVEntryLimitMins.Text) * 60) / Val(Dev2SampleRate.Text))) And CheckboxCSVlimitMins.Checked = True And CheckboxEnableCSV.Checked = True And ButtonDev2Run.Text = "Stop") Then
            CheckboxEnableCSV.Checked = False
            CSVfilename.ReadOnly = True
            CSVfilepath.ReadOnly = True
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to CSV Ended")
        End If

        ' Dev 1&2 Manual limit on CSV entries MINS if its running
        ' If CSV Entries >= (Entry Mins * 60) / Dev12SampleRate Then......
        If ((Val(CSVcounts.Text) >= ((Val(CSVEntryLimitMins.Text) * 60) / Val(Dev12SampleRate.Text))) And CheckboxCSVlimitMins.Checked = True And CheckboxEnableCSV.Checked = True And ButtonDev12Run.Text = "Stop") Then
            CheckboxEnableCSV.Checked = False
            CSVfilename.ReadOnly = True
            CSVfilepath.ReadOnly = True
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

        If CheckboxEnableCSV.Checked = False Then
            Dim CSVpath As String = CSVfilepath.Text & "\" & CSVfilename.Text
            System.IO.File.WriteAllText(CSVpath, "")
            IndexCount = 1  ' reset index count for CSV file
            CSVcounts.Text = "0"     ' Update entries in CSV per device
            CSVsize.Text = "0"
            AllowMetaDataWrite = True
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " CSV file contents cleared")
            CSVwrite.Text = ""
        End If

    End Sub


    Private Sub CheckboxEnableCSV_CheckedChanged(sender As Object, e As EventArgs) Handles CheckboxEnableCSV.CheckedChanged

        If CheckboxEnableCSV.Checked = True Then
            CSVcounts.Text = "0"
            IndexCount = 1
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to CSV Enabled")
            CSVfilename.ReadOnly = True
            CSVfilepath.ReadOnly = True
            CSVfilestarting = True

            ResetCSV.Enabled = False
            CSVwrite.Text = ""
        Else
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to CSV Disabled")
            CSVfilename.ReadOnly = False
            CSVfilepath.ReadOnly = False

            CSVfilestarting = False

            ResetCSV.Enabled = True
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
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to LOG Started")
        Else
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Writing to LOG Stopped")
        End If

    End Sub


    Private Sub EnableAutoYChart1_CheckedChanged(sender As Object, e As EventArgs) Handles EnableAutoYChart1.CheckedChanged

        If EnableAutoYChart1.Checked = True Then
            Dev1Max.Enabled = False
            Dev1Min.Enabled = False
        Else
            Dev1Max.Enabled = True
            Dev1Min.Enabled = True
        End If

    End Sub

    Private Sub Dev1Max_TextChanged(sender As Object, e As EventArgs) Handles Dev1Max.TextChanged

        If Val(Dev1Max.Text) <= Val(Dev1Min.Text) Then
            Dev1Max.Text = Val(Dev1Min.Text) + 1
        End If

    End Sub

    Private Sub Dev1Min_TextChanged(sender As Object, e As EventArgs) Handles Dev1Min.TextChanged

        If Val(Dev1Min.Text) >= Val(Dev1Max.Text) Then
            Dev1Min.Text = Val(Dev1Max.Text) - 1
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

    Private Sub UpdateChartYAxisMinMaxInterval()

        ' Parse the minimum and maximum values from the text inputs
        Dim minVal As Double = Val(Dev1Min.Text)
        Dim maxVal As Double = Val(Dev1Max.Text)

        ' Set the minimum and maximum values for the Y-axis
        Chart1.ChartAreas(0).AxisY.Minimum = minVal
        Chart1.ChartAreas(0).AxisY.Maximum = maxVal

        ' Calculate the range of the Y-axis
        Dim scalerange As Double = maxVal - minVal

        ' Calculate the interval to have 11 labels
        Dim interval As Double = scalerange / 10

        ' Set the interval for the Y-axis
        Chart1.ChartAreas(0).AxisY.Interval = interval

        ' Configure major grid lines
        With Chart1.ChartAreas(0).AxisY.MajorGrid
            .LineColor = Color.Gray
            .LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot
            .LineWidth = 1
        End With

    End Sub


    Private Sub UpdateChartTemperatureYAxisMinMaxInterval()

        ' Parse the minimum and maximum values from the text inputs
        Dim TminVal As Double = Val(LCTempMin.Text)
        Dim TmaxVal As Double = Val(LCTempMax.Text)

        ' Set the minimum and maximum values for the Y-axis
        Chart1.ChartAreas(0).AxisY2.Minimum = TminVal
        Chart1.ChartAreas(0).AxisY2.Maximum = TmaxVal

        ' Calculate the range of the Y-axis
        Dim Tscalerange As Double = TmaxVal - TminVal

        ' Calculate the interval to have 11 labels
        Dim interval As Double = Tscalerange / 10

        ' Set the interval for the Y-axis
        Chart1.ChartAreas(0).AxisY2.Interval = interval

        ' Force labels to show 1 decimal place
        Chart1.ChartAreas(0).AxisY2.LabelStyle.Format = "0.0"

        ' Configure major grid lines
        With Chart1.ChartAreas(0).AxisY2.MajorGrid
            .LineColor = Color.Gray
            .LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot
            .LineWidth = 1
        End With

    End Sub


    Private Sub EnableChart1_CheckedChanged(sender As Object, e As EventArgs) Handles EnableChart1.CheckedChanged

        If EnableChart1.Checked = True Then
            EnableChart1.BackColor = Color.Yellow
        Else
            EnableChart1.BackColor = Color.WhiteSmoke
        End If

    End Sub


    Private Sub EnableChart2_CheckedChanged(sender As Object, e As EventArgs) Handles EnableChart2.CheckedChanged

        If EnableChart2.Checked = True Then
            EnableChart2.BackColor = Color.Aqua
        Else
            EnableChart2.BackColor = Color.WhiteSmoke
        End If

    End Sub


    Private Sub EnableChart3_CheckedChanged(sender As Object, e As EventArgs) Handles EnableChart3.CheckedChanged

        If EnableChart3.Checked = True Then
            EnableChart3.BackColor = Color.Red
        Else
            EnableChart3.BackColor = Color.WhiteSmoke
        End If

    End Sub

End Class