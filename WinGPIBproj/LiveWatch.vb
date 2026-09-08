' Live Watch

Partial Class Formtest


    Dim inst_value1FChart As Double = Double.NaN
    Dim inst_value2FChart As Double = Double.NaN
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

    Dim inst_value1FChartMinRecordedDisplay As Double = Double.NaN
    Dim inst_value1FChartMaxRecordedDisplay As Double = Double.NaN
    Dim inst_value2FChartMinRecordedDisplay As Double = Double.NaN
    Dim inst_value2FChartMaxRecordedDisplay As Double = Double.NaN
    Dim inst_TemperatureChartMinRecordedDisplay As Double = Double.NaN
    Dim inst_TemperatureChartMaxRecordedDisplay As Double = Double.NaN

    Dim AutoScale1Flag As Boolean = False
    Dim AutoScale2Flag As Boolean = False
    Dim Scalerange As Double
    Dim rangeDev1 As Double

    Dim AutoScaleFirst1Flag As Boolean = True
    Dim AutoScaleFirst2Flag As Boolean = True

    Dim value1F As Double
    Dim value2F As Double
    Dim valueTempF As Double

    Dim RunChart As Boolean = False
    Dim ChartPoints1 As Integer = 0
    Dim ChartPoints2 As Integer = 0

    ' Lightweight rolling average state
    Private q1 As New Queue(Of Double)  ' Dev1 samples
    Private q2 As New Queue(Of Double)  ' Dev2 samples
    Private sum1 As Double = 0          ' Dev1 running sum
    Private sum2 As Double = 0          ' Dev2 running sum

    ' Live Statistics - Device 1
    Private Stats1Count As Long = 0
    Private Stats1Mean As Double = 0.0
    Private Stats1M2 As Double = 0.0

    ' Live Statistics - Device 2
    Private Stats2Count As Long = 0
    Private Stats2Mean As Double = 0.0
    Private Stats2M2 As Double = 0.0

    ' Live Analysis Pop-out Chart
    Private LiveAnalysisForm As Form = Nothing
    Private LiveAnalysisChart As DataVisualization.Charting.Chart = Nothing
    Private LiveAnalysisTimeLabel As Label
    Private LiveAnalysisSample As Long = 0
    Private Stats1StdevCurrent As Double = 0.0
    Private Stats1SEMCurrent As Double = 0.0
    Private Stats2StdevCurrent As Double = 0.0
    Private Stats2SEMCurrent As Double = 0.0

    Dim inst_value1FChartRaw As Double
    Dim inst_value2FChartRaw As Double

    Private LiveAnalysisLastStats1Count As Long = 0
    Private LiveAnalysisLastStats2Count As Long = 0

    Private Stats1Max As Double = Double.MinValue
    Private Stats1Min As Double = Double.MaxValue

    Private Stats2Max As Double = Double.MinValue
    Private Stats2Min As Double = Double.MaxValue

    Private Stats1FirstValue As Double = Double.NaN
    Private Stats2FirstValue As Double = Double.NaN

    Private Stats1DeviationCurrent As Double = 0.0
    Private Stats2DeviationCurrent As Double = 0.0


    ' Returns averaged value if enabled, else raw
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


    Private Sub LiveChart()

        ' Chart 1 - Device 1 only
        If (EnableChart1.Checked = True And EnableChart2.Checked = False And RunChart = True And Dev1GPIBActivity = True) Then

            inst_value1FChartRaw = CDbl(Val(NormalizeNumericResponse(txtr1a.Text)))

            ' Live statistics use RAW reading
            'UpdateStats1(inst_value1FChartRaw)

            ' Existing Live Watch chart may use rolling averaging
            inst_value1FChart = AvgVal(inst_value1FChartRaw, 1)
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

            inst_value2FChartRaw = CDbl(Val(NormalizeNumericResponse(txtr2a.Text)))

            ' Live statistics use RAW reading
            'UpdateStats2(inst_value2FChartRaw)

            ' Existing Live Watch chart may use rolling averaging
            inst_value2FChart = AvgVal(inst_value2FChartRaw, 2)
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

            ' ==========================================================
            ' Device 1
            ' ==========================================================

            ' Raw incoming reading
            inst_value1FChartRaw =
    CDbl(Val(NormalizeNumericResponse(txtr1a.Text)))

            ' Device 1 statistics use RAW reading
            'UpdateStats1(inst_value1FChartRaw)

            ' Existing Live Watch chart may use rolling averaging
            inst_value1FChart =
    AvgVal(inst_value1FChartRaw, 1)

            Dev1ChartValue.Text =
    Format(inst_value1FChart, "0.#########")


            ' ==========================================================
            ' Device 2
            ' ==========================================================

            ' Raw incoming reading
            inst_value2FChartRaw =
    CDbl(Val(NormalizeNumericResponse(txtr2a.Text)))

            ' Device 2 statistics use RAW reading
            'UpdateStats2(inst_value2FChartRaw)

            ' Existing Live Watch chart may use rolling averaging
            inst_value2FChart =
    AvgVal(inst_value2FChartRaw, 2)

            Dev2ChartValue.Text =
    Format(inst_value2FChart, "0.#########")


            txtr1achart =
    Format(inst_value1FChart, "#0.00000000")

            txtr2achart =
    Format(inst_value2FChart, "#0.00000000")


            ' ==========================================================
            ' Plot Device 1 & Device 2 to normal Live Watch chart
            ' ==========================================================

            If DisableRollingChart.Checked = False Then

                Chart1.Series(0).Points.AddY(txtr1achart)
                Chart1.Series(1).Points.AddY(txtr2achart)

                If Chart1.Series(0).Points.Count > Val(XaxisPoints.Text) Then

                    Chart1.Series(0).Points.RemoveAt(0)
                    Chart1.Series(1).Points.RemoveAt(0)

                End If

            Else

                Chart1.Series(0).Points.AddY(txtr1achart)
                Chart1.Series(1).Points.AddY(txtr2achart)

            End If


            ' ==========================================================
            ' Chart 3 - Temperature
            ' ==========================================================

            If (EnableChart3.Checked = True And RunChart = True) Then

                ' Set up max and min for temperature
                If Val(LCTempMax.Text) > Val(LCTempMin.Text) Then

                    UpdateChartTemperatureYAxisMinMaxInterval()

                End If

                inst_value3FChart = gCurrTemp
                inst_value3FChart += Val(TempOffset.Text)

                txtr3achart =
        Format(inst_value3FChart, "#0.00000000")


                ' Plot temperature
                If DisableRollingChart.Checked = False Then

                    Chart1.Series(2).Points.AddY(txtr3achart)

                    If Chart1.Series(2).Points.Count > Val(XaxisPoints.Text) Then

                        Chart1.Series(2).Points.RemoveAt(0)

                    End If

                Else

                    Chart1.Series(2).Points.AddY(txtr3achart)

                End If


                ' Temp - record min & max for display (resettable)
                Resetmaxdiffrecorded_temp()

            End If


            ' Dummy data so chart vertical data can align
            ' if temperature is enabled later
            If (EnableChart3.Checked = False And RunChart = True) Then

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

                    ' Determine applicable sample rate for time labels.
                    Dim liveChartSampleRateText As String = ""

                    If EnableChart1.Checked = True AndAlso EnableChart2.Checked = True Then
                        liveChartSampleRateText = Dev12SampleRate.Text
                    ElseIf EnableChart1.Checked = True Then
                        liveChartSampleRateText = Dev1SampleRate.Text
                    ElseIf EnableChart2.Checked = True Then
                        liveChartSampleRateText = Dev2SampleRate.Text
                    End If

                    Dim liveChartSampleRateSeconds As Double = Val(liveChartSampleRateText)

                    ' Replace numeric sample-index labels with elapsed time.
                    ca.AxisX.CustomLabels.Clear()

                    If liveChartSampleRateSeconds > 0 Then

                        Dim tickCount As Integer = 10
                        Dim tickStep As Double = (xmax - xmin) / tickCount

                        For i As Integer = 0 To tickCount

                            Dim tickPos As Double = xmin + (i * tickStep)
                            Dim tickSeconds As Integer = CInt(Math.Max(tickPos, 0) * liveChartSampleRateSeconds)

                            Dim tickHours As Integer = tickSeconds \ 3600
                            Dim tickMinutes As Integer = (tickSeconds Mod 3600) \ 60
                            Dim tickSecs As Integer = tickSeconds Mod 60

                            Dim tickLabel As String = $"{tickHours:00}:{tickMinutes:00}:{tickSecs:00}"

                            Dim labelLow As Double = tickPos - (tickStep / 2)
                            Dim labelHigh As Double = tickPos + (tickStep / 2)

                            ca.AxisX.CustomLabels.Add(labelLow, labelHigh, tickLabel)

                        Next

                    End If

                Else
                    ' No points yet – let chart decide
                    ca.AxisX.Minimum = Double.NaN
                    ca.AxisX.Maximum = Double.NaN
                    ca.AxisX.Interval = Double.NaN
                    ca.AxisX.CustomLabels.Clear()
                End If
            End If

        Else
            ' Rolling disabled – let chart auto-manage X axis
            If Chart1.ChartAreas.Count > 0 Then
                Dim ca = Chart1.ChartAreas(0)

                ca.AxisX.Minimum = Double.NaN
                ca.AxisX.Maximum = Double.NaN
                ca.AxisX.Interval = Double.NaN
                ca.AxisX.CustomLabels.Clear()

                ca.RecalculateAxesScale()

                Dim disabledMin As Double = ca.AxisX.Minimum
                Dim disabledMax As Double = ca.AxisX.Maximum

                Dim liveChartSampleRateTextDisabled As String = ""

                If EnableChart1.Checked = True AndAlso EnableChart2.Checked = True Then
                    liveChartSampleRateTextDisabled = Dev12SampleRate.Text
                ElseIf EnableChart1.Checked = True Then
                    liveChartSampleRateTextDisabled = Dev1SampleRate.Text
                ElseIf EnableChart2.Checked = True Then
                    liveChartSampleRateTextDisabled = Dev2SampleRate.Text
                End If

                Dim liveChartSampleRateSecondsDisabled As Double = Val(liveChartSampleRateTextDisabled)

                If liveChartSampleRateSecondsDisabled > 0 AndAlso disabledMax > disabledMin Then

                    Dim tickCount As Integer = 10
                    Dim tickStep As Double = (disabledMax - disabledMin) / tickCount

                    ca.AxisX.Interval = tickStep

                    For i As Integer = 0 To tickCount

                        Dim tickPos As Double = disabledMin + (i * tickStep)
                        Dim tickSeconds As Integer = CInt(Math.Max(tickPos, 0) * liveChartSampleRateSecondsDisabled)

                        Dim tickHours As Integer = tickSeconds \ 3600
                        Dim tickMinutes As Integer = (tickSeconds Mod 3600) \ 60
                        Dim tickSecs As Integer = tickSeconds Mod 60

                        Dim tickLabel As String = $"{tickHours:00}:{tickMinutes:00}:{tickSecs:00}"

                        Dim labelLow As Double = tickPos - (tickStep / 2)
                        Dim labelHigh As Double = tickPos + (tickStep / 2)

                        ca.AxisX.CustomLabels.Add(labelLow, labelHigh, tickLabel)

                    Next

                End If

            End If
        End If

        'UpdateLiveAnalysisChart()

    End Sub


    Private Sub ButtonStatsInfo_Click(sender As Object, e As EventArgs) Handles ButtonStatsInfo.Click

        Dim frm As New Form With {
        .Text = "Live Statistics Help / Info",
        .StartPosition = FormStartPosition.CenterParent,
        .FormBorderStyle = FormBorderStyle.FixedDialog,
        .ShowIcon = False,
        .ShowInTaskbar = False,
        .Width = 700,
        .Height = 620,
        .MinimizeBox = False,
        .MaximizeBox = False
    }

        Dim txt As New RichTextBox With {
    .ReadOnly = True,
    .WordWrap = True,
    .Dock = DockStyle.Fill,
    .Font = New Font("Segoe UI", 9),
    .BackColor = Color.White,
    .ScrollBars = RichTextBoxScrollBars.Vertical,
    .BorderStyle = BorderStyle.Fixed3D,
    .Text =
    "LIVE CHART STATISTICS" & vbCrLf & vbCrLf &
    "WinGPIB can calculate live statistics independently for Device 1 and Device 2 using each raw incoming measurement reading." & vbCrLf &
    "Statistics run automatically whenever a device is actively acquiring and its Enable Statistics checkbox is checked - independent of whether the Live Chart or Data Log/CSV logging is running." & vbCrLf &
    "The Enable Statistics checkbox for each device is located in that device's configuration box on the Cmd Line tab, alongside its other settings. It can only be changed while the device is stopped - tick it before pressing Run, as it locks once the device starts." & vbCrLf &
    "The optional Live Chart rolling average does not affect these statistics." & vbCrLf & vbCrLf &
    "SAMPLES (N)" & vbCrLf &
    "The total number of individual readings included in the current statistics calculation." & vbCrLf &
    "The sample count starts from zero when RESET STAT is pressed." & vbCrLf & vbCrLf &
    "MEAN" & vbCrLf &
    "The arithmetic average of all readings collected since the statistics were started or reset." & vbCrLf & vbCrLf &
    "MAX / MIN RECORDED" & vbCrLf &
    "The highest and lowest individual raw readings seen since the statistics were started or reset." & vbCrLf &
    "These update on every new reading and are cleared back to a fresh state by RESET STAT." & vbCrLf & vbCrLf &
    "STDEV - STANDARD DEVIATION" & vbCrLf &
    "Shows how much the individual readings vary or scatter around the calculated mean." & vbCrLf &
    "A smaller STDEV generally indicates less variation or noise in the readings." & vbCrLf & vbCrLf &
    "WinGPIB calculates sample standard deviation using Welford's running algorithm. This is mathematically equivalent to the conventional sample STDEV calculation but avoids having to store every individual reading." & vbCrLf & vbCrLf &
    "Sample STDEV = Sqrt(Sum((Xi - Mean)^2) / (N - 1))" & vbCrLf & vbCrLf &
    "SEM - STANDARD ERROR OF THE MEAN" & vbCrLf &
    "Shows how precisely the mean has been determined from the accumulated readings." & vbCrLf &
    "SEM is derived from the STDEV and decreases as more independent readings are averaged." & vbCrLf & vbCrLf &
    "SEM = STDEV / Sqrt(N)" & vbCrLf & vbCrLf &
    "AVERAGING GAIN (DIGITS)" & vbCrLf &
    "Shows the theoretical increase in resolution obtained by averaging N independent readings." & vbCrLf & vbCrLf &
    "Averaging Gain = 0.5 x Log10(N)" & vbCrLf & vbCrLf &
    "Examples:" & vbCrLf &
    "  10 readings     = 0.50 digits" & vbCrLf &
    "  100 readings    = 1.00 digits" & vbCrLf &
    "  1,000 readings  = 1.50 digits" & vbCrLf &
    "  10,000 readings = 2.00 digits" & vbCrLf & vbCrLf &
    "LIVE ANALYSIS CHART" & vbCrLf &
    "The Live Analysis chart displays the raw Device 1 and Device 2 readings together with their running Mean, STDEV and SEM, plus temperature when enabled." & vbCrLf &
    "The analysis chart uses the same raw readings as the statistics calculations and is not affected by the optional Live Chart rolling average." & vbCrLf &
    "The Live Analysis chart does not require the Live Chart to be started - it runs from the same live statistics as soon as a device is running with Enable Statistics checked." & vbCrLf &
    "When both Device 1 and Device 2 are running together, the chart advances once per matched pair of readings rather than once per device, so the two devices share a common position on the chart instead of doubling the update rate." & vbCrLf &
    "The X-axis shows elapsed time (HH:mm:ss), calculated from the sample rate of whichever device(s) are running." & vbCrLf & vbCrLf &
    "DATA LOG / CSV" & vbCrLf &
    "When statistics are enabled, the current Samples, Mean, STDEV, SEM and Averaging Gain values are also available in the Data Log and CSV output." & vbCrLf &
    "Statistics fields remain in fixed positions for Device 1 and Device 2. When statistics are disabled for a device, those fields are left blank." & vbCrLf & vbCrLf &
    "IMPORTANT" & vbCrLf &
    "• The statistics are calculated independently for Device 1 and Device 2." & vbCrLf &
    "• RESET STAT clears the accumulated statistics for that device and starts again from zero." & vbCrLf &
    "• Statistics are calculated from individual raw readings before any optional Live Chart rolling averaging is applied." & vbCrLf &
    "• Statistics run whenever the device is active and Enable Statistics is checked - the Live Chart and Data Log/CSV logging do not need to be running." & vbCrLf &
    "• The Enable Statistics checkbox locks while its device is running - enable it before pressing Run, not after." & vbCrLf &
    "• STDEV includes all variation present in the readings, including random noise, drift, temperature effects and other changes." & vbCrLf &
    "• SEM is most meaningful when the readings are independent and the underlying measured value is stable." & vbCrLf &
    "• A small SEM does not by itself represent the total measurement uncertainty or accuracy." & vbCrLf &
    "• Averaging reduces random noise but does not remove systematic errors or long-term drift."
}

        ' Make headings bold.
        Dim headings() As String = {
    "LIVE CHART STATISTICS",
    "SAMPLES (N)",
    "MEAN",
    "MAX / MIN RECORDED",
    "STDEV - STANDARD DEVIATION",
    "SEM - STANDARD ERROR OF THE MEAN",
    "AVERAGING GAIN (DIGITS)",
    "LIVE ANALYSIS CHART",
    "DATA LOG / CSV",
    "IMPORTANT"
}

        For Each heading As String In headings

            Dim start As Integer =
            txt.Text.IndexOf(heading, StringComparison.Ordinal)

            If start >= 0 Then
                txt.Select(start, heading.Length)
                txt.SelectionFont = New Font(txt.Font, FontStyle.Bold)
            End If

        Next

        ' Return cursor to beginning and remove selection.
        txt.Select(0, 0)

        Dim btn As New Button With {
    .Text = "OK",
    .Width = 100,
    .Height = 30,
    .Anchor = AnchorStyles.Bottom
}

        AddHandler btn.Click,
    Sub()
        frm.Close()
    End Sub

        Dim panel As New Panel With {
    .Dock = DockStyle.Bottom,
    .Height = 45
}

        panel.Controls.Add(btn)

        AddHandler panel.Resize,
    Sub()
        btn.Left = (panel.ClientSize.Width - btn.Width) \ 2
        btn.Top = 7
    End Sub

        frm.Controls.Add(txt)
        frm.Controls.Add(panel)

        frm.AcceptButton = btn

        frm.Show()

    End Sub

    Private Sub UpdateStats1(value As Double)

        If CheckBoxStats1Enable.Checked = False Then Exit Sub

        Stats1Count += 1

        If Stats1Count = 1 Then
            Stats1FirstValue = value
        End If

        Dim delta As Double = value - Stats1Mean
        Stats1Mean += delta / Stats1Count

        Dim delta2 As Double = value - Stats1Mean
        Stats1M2 += delta * delta2

        ' Max / Min recorded
        If value > Stats1Max Then Stats1Max = value
        If value < Stats1Min Then Stats1Min = value

        ' PPM Deviation from first sample since last reset
        'Stats1DeviationCurrent = (value - Stats1FirstValue) * 1000000
        Stats1DeviationCurrent = If(Stats1FirstValue <> 0, (value - Stats1FirstValue) / Stats1FirstValue * 1000000, 0)

        ' Number of samples
        LabelStats1Samples.Text = Stats1Count.ToString()

        ' Mean
        LabelStats1Mean.Text = Stats1Mean.ToString("0.0000000000")

        ' Max / Min recorded
        LabelStats1Max.Text = Stats1Max.ToString("0.0000000000")
        LabelStats1Min.Text = Stats1Min.ToString("0.0000000000")

        ' Max Diff and PPM Deviation from first sample
        LabelStats1MaxDiff.Text = (Stats1Max - Stats1Min).ToString("0.0000000000")
        LabelStats1Deviation.Text = Stats1DeviationCurrent.ToString("0.0000")

        ' Need at least 2 readings for STDEV
        If Stats1Count >= 2 Then

            Dim variance As Double = Stats1M2 / (Stats1Count - 1)
            Dim stdev As Double = Math.Sqrt(variance)
            Dim sem As Double = stdev / Math.Sqrt(Stats1Count)

            Stats1StdevCurrent = stdev
            Stats1SEMCurrent = sem

            LabelStats1Stdev.Text = stdev.ToString("0.0000000000")
            LabelStats1SEM.Text = sem.ToString("0.0000000000")

        Else

            Stats1StdevCurrent = 0.0
            Stats1SEMCurrent = 0.0

            LabelStats1Stdev.Text = "0.0000000000"
            LabelStats1SEM.Text = "0.0000000000"

        End If

        ' Theoretical averaging gain in digits
        If Stats1Count > 0 Then
            Dim digitsGained As Double = 0.5 * Math.Log10(Stats1Count)
            LabelStats1Digits.Text = digitsGained.ToString("0.00")
        Else
            LabelStats1Digits.Text = "0.00"
        End If

    End Sub


    Private Sub UpdateStats2(value As Double)

        If CheckBoxStats2Enable.Checked = False Then Exit Sub

        Stats2Count += 1

        If Stats2Count = 1 Then
            Stats2FirstValue = value
        End If

        Dim delta As Double = value - Stats2Mean
        Stats2Mean += delta / Stats2Count

        Dim delta2 As Double = value - Stats2Mean
        Stats2M2 += delta * delta2

        ' Max / Min recorded
        If value > Stats2Max Then Stats2Max = value
        If value < Stats2Min Then Stats2Min = value

        ' PPM Deviation from first sample since last reset
        'Stats2DeviationCurrent = (value - Stats2FirstValue) * 1000000
        Stats2DeviationCurrent = If(Stats2FirstValue <> 0, (value - Stats2FirstValue) / Stats2FirstValue * 1000000, 0)

        ' Number of samples
        LabelStats2Samples.Text = Stats2Count.ToString()

        ' Mean
        LabelStats2Mean.Text = Stats2Mean.ToString("0.0000000000")

        ' Max / Min recorded
        LabelStats2Max.Text = Stats2Max.ToString("0.0000000000")
        LabelStats2Min.Text = Stats2Min.ToString("0.0000000000")

        ' Max Diff and PPM Deviation from first sample
        LabelStats2MaxDiff.Text = (Stats2Max - Stats2Min).ToString("0.0000000000")
        LabelStats2Deviation.Text = Stats2DeviationCurrent.ToString("0.0000")

        ' Need at least 2 readings for STDEV
        If Stats2Count >= 2 Then

            Dim variance As Double = Stats2M2 / (Stats2Count - 1)
            Dim stdev As Double = Math.Sqrt(variance)
            Dim sem As Double = stdev / Math.Sqrt(Stats2Count)

            Stats2StdevCurrent = stdev
            Stats2SEMCurrent = sem

            LabelStats2Stdev.Text = stdev.ToString("0.0000000000")
            LabelStats2SEM.Text = sem.ToString("0.0000000000")

        Else

            Stats2StdevCurrent = 0.0
            Stats2SEMCurrent = 0.0

            LabelStats2Stdev.Text = "0.0000000000"
            LabelStats2SEM.Text = "0.0000000000"

        End If

        ' Theoretical averaging gain in digits
        If Stats2Count > 0 Then
            Dim digitsGained As Double = 0.5 * Math.Log10(Stats2Count)
            LabelStats2Digits.Text = digitsGained.ToString("0.00")
        Else
            LabelStats2Digits.Text = "0.00"
        End If

    End Sub


    Private Sub ProcessLiveStatistics(deviceNumber As Integer, value As Double)

        ' Statistics now live on the Meters tab and are shown
        ' whenever the device is running, independent of whether
        ' Live Chart or CSV logging is active. The per-device
        ' Enable Statistics checkbox (checked inside UpdateStats1/2)
        ' is the only gate.

        Select Case deviceNumber

            Case 1

                ' UpdateStats1 checks the Device 1
                ' Enable Statistics checkbox itself.
                UpdateStats1(value)


            Case 2

                ' UpdateStats2 checks the Device 2
                ' Enable Statistics checkbox itself.
                UpdateStats2(value)

        End Select

    End Sub


    Private Sub ButtonStats1Reset_Click(sender As Object, e As EventArgs) Handles ButtonStats1Reset.Click

        Stats1Count = 0
        Stats1Mean = 0.0
        Stats1M2 = 0.0
        Stats1Max = Double.MinValue
        Stats1Min = Double.MaxValue
        Stats1FirstValue = Double.NaN
        Stats1DeviationCurrent = 0.0

        LabelStats1Samples.Text = "-"
        LabelStats1Mean.Text = "-"
        LabelStats1Stdev.Text = "-"
        LabelStats1SEM.Text = "-"
        LabelStats1Digits.Text = "-"
        LabelStats1Max.Text = "-"
        LabelStats1Min.Text = "-"
        LabelStats1MaxDiff.Text = "-"
        LabelStats1Deviation.Text = "-"

    End Sub


    Private Sub ButtonStats2Reset_Click(sender As Object, e As EventArgs) Handles ButtonStats2Reset.Click

        Stats2Count = 0
        Stats2Mean = 0.0
        Stats2M2 = 0.0
        Stats2Max = Double.MinValue
        Stats2Min = Double.MaxValue
        Stats2FirstValue = Double.NaN
        Stats2DeviationCurrent = 0.0

        LabelStats2Samples.Text = "-"
        LabelStats2Mean.Text = "-"
        LabelStats2Stdev.Text = "-"
        LabelStats2SEM.Text = "-"
        LabelStats2Digits.Text = "-"
        LabelStats2Max.Text = "-"
        LabelStats2Min.Text = "-"
        LabelStats2MaxDiff.Text = "-"
        LabelStats2Deviation.Text = "-"

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

        ' Don't seed or update until Device 1's chart value has
        ' actually been assigned a real reading.
        If Double.IsNaN(inst_value1FChart) Then Exit Sub

        ' Seed both Min and Max from the very first reading instead
        ' of waiting for RESET to be pressed manually.
        If Double.IsNaN(inst_value1FChartMaxRecordedDisplay) Then
            inst_value1FChartMaxRecordedDisplay = inst_value1FChart
            inst_value1FChartMinRecordedDisplay = inst_value1FChart
        End If

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

        ' Don't seed or update until Device 2's chart value has
        ' actually been assigned a real reading.
        If Double.IsNaN(inst_value2FChart) Then Exit Sub

        ' Seed both Min and Max from the very first reading instead
        ' of waiting for RESET to be pressed manually.
        If Double.IsNaN(inst_value2FChartMaxRecordedDisplay) Then
            inst_value2FChartMaxRecordedDisplay = inst_value2FChart
            inst_value2FChartMinRecordedDisplay = inst_value2FChart
        End If

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

        ' Seed both Min and Max from the very first reading instead
        ' of waiting for RESET to be pressed manually.
        If Double.IsNaN(inst_TemperatureChartMaxRecordedDisplay) Then
            inst_TemperatureChartMaxRecordedDisplay = inst_value3FChart
            inst_TemperatureChartMinRecordedDisplay = inst_value3FChart
        End If

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
        'ButtonLiveChartPopout.Enabled = False

        YaxisDiff.Text = "0"

        LabeChartMinutes.Text = "0hrs 00mins 00secs"

    End Sub


    Private Sub ButtonPauseChart_Click(sender As Object, e As EventArgs) Handles ButtonPauseChart.Click

        RunChart = Not RunChart

        ' Chart currently running and user just hit pause
        If (RunChart = False) Then

            ButtonPauseChart.Text = "Start Chart"
            ButtonClearChart.Enabled = True

            ' Disable Live Analysis while Live Chart is paused
            'ButtonLiveChartPopout.Enabled = False

        End If

        ' Chart currently paused and user just hit run
        If (RunChart = True) Then

            ButtonPauseChart.Text = "Pause Chart"
            ButtonClearChart.Enabled = False

            ' Enable Live Analysis only if at least one
            ' statistics function is enabled
            'ButtonLiveChartPopout.Enabled = CheckBoxStats1Enable.Checked OrElse CheckBoxStats2Enable.Checked

            ' Set Y-scale of chart based on Min/Max ensuring at least 1DP and number of DP's set in Min/Max
            ' Parse values from textboxes
            Dim minValue As Double
            Dim maxValue As Double

            If Double.TryParse(Dev1Min.Text, minValue) AndAlso
           Double.TryParse(Dev1Max.Text, maxValue) AndAlso
           EnableAutoYChart1.Checked = False Then

                ' Determine the number of decimal places based on maximum precision
                Dim decimalPlaces As Integer =
                Math.Max(GetMaxPrecision(minValue, maxValue), 1)

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
            LabelChartPoints2.Text = chartPoints2.ToString()
        End If

        ' Final calcs
        If sampleRateText <> "" AndAlso pointsLabel IsNot Nothing Then
            pointsLabel.Text = points.ToString()

            Dim totalSeconds As Integer = CInt(Val(sampleRateText) * points)
            Dim hours As Integer = totalSeconds \ 3600
            Dim minutes As Integer = (totalSeconds Mod 3600) \ 60
            Dim seconds As Integer = totalSeconds Mod 60

            LabeChartMinutes.Text = $"{hours}hrs {minutes:00}mins {seconds:00}secs"         ' Live Chart tab

            'If LiveAnalysisTimeLabel IsNot Nothing AndAlso LiveAnalysisTimeLabel.IsDisposed = False Then    ' Live Analysis tab
            'LiveAnalysisTimeLabel.Text = "Visible Chart =  " & $"{hours}hrs {minutes:00}mins {seconds:00}secs"
            'End If

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


    Private Sub EnableAutoYChart1_CheckedChanged(sender As Object, e As EventArgs) Handles EnableAutoYChart1.CheckedChanged

        If EnableAutoYChart1.Checked = True Then
            Dev1Max.Enabled = False
            Dev1Min.Enabled = False
        Else
            Dev1Max.Enabled = True
            Dev1Min.Enabled = True
        End If

    End Sub


    Private Sub Dev1Max_Leave(sender As Object, e As EventArgs) Handles Dev1Max.Leave

        If Val(Dev1Max.Text) <= Val(Dev1Min.Text) Then
            Dev1Max.Text = Val(Dev1Min.Text) + 1
        End If

    End Sub


    Private Sub Dev1Min_Leave(sender As Object, e As EventArgs) Handles Dev1Min.Leave

        If Val(Dev1Min.Text) >= Val(Dev1Max.Text) Then
            Dev1Min.Text = Val(Dev1Max.Text) - 1
        End If

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


    Private Sub ButtonLiveChartPopout_Click(sender As Object, e As EventArgs) Handles ButtonLiveChartPopout.Click

        ' If already open, simply bring it to the front.
        If LiveAnalysisForm IsNot Nothing AndAlso LiveAnalysisForm.IsDisposed = False Then

            LiveAnalysisForm.BringToFront()
            LiveAnalysisForm.Activate()
            Exit Sub

        End If


        LiveAnalysisForm = New Form With {
        .Text = "WinGPIB Live Analysis Chart",
        .StartPosition = FormStartPosition.CenterParent,
        .Width = 1100,
        .Height = 840,
        .MinimumSize = New Size(1100, 840),
        .ShowIcon = False,
        .ShowInTaskbar = True,
        .BackColor = Color.WhiteSmoke
    }


        LiveAnalysisChart = New DataVisualization.Charting.Chart With {
        .Dock = DockStyle.Fill,
        .BackColor = Color.WhiteSmoke
    }

        ' ==========================================================
        ' Chart Area 1 - Device readings and running means
        ' ==========================================================
        Dim areaMeasurement As New DataVisualization.Charting.ChartArea("Measurement")

        areaMeasurement.Position = New DataVisualization.Charting.ElementPosition(4, 8, 88, 48)
        areaMeasurement.InnerPlotPosition = New DataVisualization.Charting.ElementPosition(12, 5, 85, 90)

        areaMeasurement.BackColor = Color.Black

        areaMeasurement.AxisY.IsStartedFromZero = False
        areaMeasurement.AxisY.LabelStyle.Format = "#0.##########"

        areaMeasurement.AxisX.LabelStyle.ForeColor = Color.Black
        areaMeasurement.AxisY.LabelStyle.ForeColor = Color.Black

        areaMeasurement.AxisX.LineColor = Color.Gray
        areaMeasurement.AxisY.LineColor = Color.Gray

        areaMeasurement.AxisX.MajorGrid.LineColor = Color.DimGray
        areaMeasurement.AxisY.MajorGrid.LineColor = Color.DimGray

        areaMeasurement.AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        areaMeasurement.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        areaMeasurement.AxisX.LabelStyle.Enabled = True
        areaMeasurement.AxisX.LabelStyle.Format = "0"
        areaMeasurement.AxisX.Title = ""
        areaMeasurement.AxisX.TitleForeColor = Color.Black

        'areaMeasurement.AxisY.Title = "DEVICE 1 & 2 / RUNNING MEAN"
        areaMeasurement.AxisY.Title = ""
        areaMeasurement.AxisY.TitleForeColor = Color.Black

        ' Minor grid lines - Measurement
        areaMeasurement.AxisX.MinorGrid.Enabled = True
        areaMeasurement.AxisY.MinorGrid.Enabled = True
        areaMeasurement.AxisX.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
        areaMeasurement.AxisY.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
        areaMeasurement.AxisX.MinorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot
        areaMeasurement.AxisY.MinorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        areaMeasurement.AxisX.IsLabelAutoFit = False
        areaMeasurement.AxisX.LabelStyle.Font = New Font("Microsoft Sans Serif", 8)
        areaMeasurement.AxisY.IsLabelAutoFit = False
        areaMeasurement.AxisY.LabelStyle.Font = New Font("Microsoft Sans Serif", 8)

        'areaMeasurement.AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount


        ' ==========================================================
        ' Chart Area 2 - STDEV / SEM
        ' ==========================================================
        Dim areaStatistics As New DataVisualization.Charting.ChartArea("Statistics")

        areaStatistics.Position =
        New DataVisualization.Charting.ElementPosition(4, 57, 88, 23)
        areaStatistics.InnerPlotPosition = New DataVisualization.Charting.ElementPosition(12, 5, 85, 90)

        areaStatistics.BackColor = Color.Black

        areaStatistics.AxisY.IsStartedFromZero = False
        'areaStatistics.AxisY.LabelStyle.Format = "0.###E+00"
        areaStatistics.AxisY.LabelStyle.Format = "0.0000000"

        areaStatistics.AxisX.LabelStyle.ForeColor = Color.Black
        areaStatistics.AxisY.LabelStyle.ForeColor = Color.Black

        areaStatistics.AxisX.LineColor = Color.Gray
        areaStatistics.AxisY.LineColor = Color.Gray

        areaStatistics.AxisX.MajorGrid.LineColor = Color.DimGray
        areaStatistics.AxisY.MajorGrid.LineColor = Color.DimGray

        areaStatistics.AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        areaStatistics.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        areaStatistics.AxisX.LabelStyle.Enabled = False

        'areaStatistics.AxisY.Title = "DEVICE 1 & 2 / STDEV / SEM"
        areaStatistics.AxisY.Title = ""
        areaStatistics.AxisY.TitleForeColor = Color.Black

        ' Minor grid lines - Statistics
        areaStatistics.AxisX.MinorGrid.Enabled = True
        areaStatistics.AxisY.MinorGrid.Enabled = True
        areaStatistics.AxisX.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
        areaStatistics.AxisY.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
        areaStatistics.AxisX.MinorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot
        areaStatistics.AxisY.MinorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        areaStatistics.AxisY.IsLabelAutoFit = False
        areaStatistics.AxisY.LabelStyle.Font = New Font("Microsoft Sans Serif", 8)
        areaStatistics.AxisY2.IsLabelAutoFit = False
        areaStatistics.AxisY2.LabelStyle.Font = New Font("Microsoft Sans Serif", 8)

        'areaStatistics.AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount

        ' Secondary Y-axis for PPM Deviation.
        '
        ' PPM Deviation is already scaled to human-friendly units
        ' (multiplied by 1,000,000), whereas STDEV/SEM on the
        ' primary axis are raw, unscaled values several orders of
        ' magnitude smaller. Sharing one axis would make one or
        ' the other unreadable, so PPM Deviation gets its own
        ' independent scale on the right while staying in the
        ' same "measurement stability" panel.
        areaStatistics.AxisY2.Enabled = DataVisualization.Charting.AxisEnabled.True
        areaStatistics.AxisY2.IsStartedFromZero = False
        areaStatistics.AxisY2.LabelStyle.Enabled = True
        areaStatistics.AxisY2.LabelStyle.Format = "0.0000"
        areaStatistics.AxisY2.LabelStyle.ForeColor = Color.Black
        areaStatistics.AxisY2.LineColor = Color.Gray
        areaStatistics.AxisY2.Title = ""
        areaStatistics.AxisY2.TitleForeColor = Color.Black
        areaStatistics.AxisY2.MajorGrid.Enabled = False
        areaStatistics.AxisY2.MinorGrid.Enabled = False


        ' ==========================================================
        ' Chart Area 3 - Temperature
        ' ==========================================================
        Dim areaTemperature As New DataVisualization.Charting.ChartArea("Temperature")

        areaTemperature.Position = New DataVisualization.Charting.ElementPosition(4, 82, 88, 15)
        areaTemperature.InnerPlotPosition = New DataVisualization.Charting.ElementPosition(12, 5, 85, 90)

        areaTemperature.BackColor = Color.Black

        areaTemperature.AxisY.IsStartedFromZero = False
        areaTemperature.AxisY.LabelStyle.Format = "0.00"

        areaTemperature.AxisX.LabelStyle.ForeColor = Color.Black
        areaTemperature.AxisY.LabelStyle.ForeColor = Color.Black

        areaTemperature.AxisX.LineColor = Color.Gray
        areaTemperature.AxisY.LineColor = Color.Gray

        areaTemperature.AxisX.MajorGrid.LineColor = Color.DimGray
        areaTemperature.AxisY.MajorGrid.LineColor = Color.DimGray

        areaTemperature.AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        areaTemperature.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        'areaTemperature.AxisX.Title = "SAMPLES"
        'areaTemperature.AxisX.TitleForeColor = Color.Black

        'areaTemperature.AxisY.Title = "TEMPERATURE"
        areaTemperature.AxisY.Title = ""
        areaTemperature.AxisY.TitleForeColor = Color.Black

        areaTemperature.AxisX.LabelStyle.Enabled = False

        ' Minor grid lines - Temperature
        areaTemperature.AxisX.MinorGrid.Enabled = True
        areaTemperature.AxisY.MinorGrid.Enabled = True
        areaTemperature.AxisX.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
        areaTemperature.AxisY.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
        areaTemperature.AxisX.MinorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot
        areaTemperature.AxisY.MinorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        areaTemperature.AxisY.IsLabelAutoFit = False
        areaTemperature.AxisY.LabelStyle.Font = New Font("Microsoft Sans Serif", 8)

        'areaTemperature.AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount


        LiveAnalysisChart.ChartAreas.Add(areaMeasurement)
        LiveAnalysisChart.ChartAreas.Add(areaStatistics)
        LiveAnalysisChart.ChartAreas.Add(areaTemperature)


        ' ==========================================================
        ' Chart Area Titles
        ' ==========================================================
        Dim titleMeasurement As New DataVisualization.Charting.Title
        titleMeasurement.Text = "DEVICE 1 & 2 DATA" & vbCrLf & "/ RUNNING MEAN"
        titleMeasurement.DockedToChartArea = "Measurement"
        titleMeasurement.Docking = DataVisualization.Charting.Docking.Left
        titleMeasurement.IsDockedInsideChartArea = False
        titleMeasurement.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        titleMeasurement.TextOrientation = DataVisualization.Charting.TextOrientation.Rotated270
        titleMeasurement.Position.Auto = False
        titleMeasurement.Position = New DataVisualization.Charting.ElementPosition(2.0F, 19.0F, 4.0F, 25.0F)
        LiveAnalysisChart.Titles.Add(titleMeasurement)


        Dim titleStatistics As New DataVisualization.Charting.Title
        titleStatistics.Text = "STDEV / SEM"
        titleStatistics.DockedToChartArea = "Statistics"
        titleStatistics.Docking = DataVisualization.Charting.Docking.Left
        titleStatistics.IsDockedInsideChartArea = False
        titleStatistics.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        titleStatistics.TextOrientation = DataVisualization.Charting.TextOrientation.Rotated270
        titleStatistics.Position.Auto = False
        titleStatistics.Position = New DataVisualization.Charting.ElementPosition(2.0F, 60.0F, 4.0F, 18.0F)
        LiveAnalysisChart.Titles.Add(titleStatistics)


        Dim titleStatisticsPPM As New DataVisualization.Charting.Title
        titleStatisticsPPM.Text = "PPM DEVIATION"
        titleStatisticsPPM.DockedToChartArea = "Statistics"
        titleStatisticsPPM.Docking = DataVisualization.Charting.Docking.Right
        titleStatisticsPPM.IsDockedInsideChartArea = False
        titleStatisticsPPM.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        titleStatisticsPPM.TextOrientation = DataVisualization.Charting.TextOrientation.Rotated90
        titleStatisticsPPM.Position.Auto = False
        titleStatisticsPPM.Position = New DataVisualization.Charting.ElementPosition(94.0F, 59.0F, 4.0F, 18.0F)
        LiveAnalysisChart.Titles.Add(titleStatisticsPPM)


        Dim titleTemperature As New DataVisualization.Charting.Title
        titleTemperature.Text = "TEMPERATURE"
        titleTemperature.DockedToChartArea = "Temperature"
        titleTemperature.Docking = DataVisualization.Charting.Docking.Left
        titleTemperature.IsDockedInsideChartArea = False
        titleTemperature.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        titleTemperature.TextOrientation = DataVisualization.Charting.TextOrientation.Rotated270
        titleTemperature.Position.Auto = False
        titleTemperature.Position = New DataVisualization.Charting.ElementPosition(2.0F, 83.0F, 4.0F, 12.0F)
        LiveAnalysisChart.Titles.Add(titleTemperature)


        ' ==========================================================
        ' Series
        ' ==========================================================

        AddLiveAnalysisSeries(
        "Device 1",
        "Measurement",
        Color.Yellow)

        AddLiveAnalysisSeries(
        "Device 2",
        "Measurement",
        Color.Aqua)

        AddLiveAnalysisSeries(
        "Dev 1 Mean",
        "Measurement",
        Color.Orange)

        AddLiveAnalysisSeries(
        "Dev 2 Mean",
        "Measurement",
        Color.Lime)

        AddLiveAnalysisSeries(
        "Dev 1 STDEV",
        "Statistics",
        Color.LightGray)

        AddLiveAnalysisSeries(
        "Dev 1 SEM",
        "Statistics",
        Color.DeepSkyBlue)

        AddLiveAnalysisSeries(
        "Dev 2 STDEV",
        "Statistics",
        Color.Magenta)

        AddLiveAnalysisSeries(
        "Dev 2 SEM",
        "Statistics",
        Color.LimeGreen)

        AddLiveAnalysisSeries(
        "Dev 1 PPM Deviation",
        "Statistics",
        Color.White)

        LiveAnalysisChart.Series("Dev 1 PPM Deviation").YAxisType =
        DataVisualization.Charting.AxisType.Secondary

        AddLiveAnalysisSeries(
        "Dev 2 PPM Deviation",
        "Statistics",
        Color.LightGray)

        LiveAnalysisChart.Series("Dev 2 PPM Deviation").YAxisType =
        DataVisualization.Charting.AxisType.Secondary

        AddLiveAnalysisSeries(
        "Temperature",
        "Temperature",
        Color.Red)


        ' ==========================================================
        ' Trace enable/disable checkboxes - Becomes Legends also
        ' ==========================================================
        Dim liveToggles As New List(Of CheckBox)

        Dim gbDev1 As New GroupBox With {.Text = "Device 1 - " & txtname1.Text, .BackColor = Color.WhiteSmoke, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim gbDev2 As New GroupBox With {.Text = "Device 2 - " & txtname2.Text, .BackColor = Color.WhiteSmoke, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim gbTemp As New GroupBox With {.Text = "Temperature", .BackColor = Color.WhiteSmoke, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}

        LiveAnalysisChart.Controls.Add(gbDev1)
        LiveAnalysisChart.Controls.Add(gbDev2)
        LiveAnalysisChart.Controls.Add(gbTemp)

        Dim AddTraceToggle = Function(parent As GroupBox, seriesName As String, displayText As String, color As Color) As CheckBox
                                 Dim cb As New CheckBox With {
                      .Text = displayText,
                      .ForeColor = Color.Black,
                      .BackColor = color,
                      .AutoSize = False,
                      .Checked = True,
                      .Height = 18,
                      .Tag = color,
                      .Font = New Font("Segoe UI", 8, FontStyle.Regular)
                  }
                                 AddHandler cb.CheckedChanged, Sub(s, ev)
                                                                   LiveAnalysisChart.Series(seriesName).Enabled = cb.Checked
                                                               End Sub
                                 liveToggles.Add(cb)
                                 parent.Controls.Add(cb)
                                 Return cb
                             End Function

        Dim tglDevice1 = AddTraceToggle(gbDev1, "Device 1", "Data", Color.Yellow)
        Dim tglDev1Mean = AddTraceToggle(gbDev1, "Dev 1 Mean", "Mean", Color.Orange)
        Dim tglDev1Stdev = AddTraceToggle(gbDev1, "Dev 1 STDEV", "STDEV", Color.LightGray)
        Dim tglDev1SEM = AddTraceToggle(gbDev1, "Dev 1 SEM", "SEM", Color.DeepSkyBlue)
        Dim tglDev1PPM = AddTraceToggle(gbDev1, "Dev 1 PPM Deviation", "PPM Dev.", Color.White)

        Dim tglDevice2 = AddTraceToggle(gbDev2, "Device 2", "Data", Color.Aqua)
        Dim tglDev2Mean = AddTraceToggle(gbDev2, "Dev 2 Mean", "Mean", Color.Lime)
        Dim tglDev2Stdev = AddTraceToggle(gbDev2, "Dev 2 STDEV", "STDEV", Color.Magenta)
        Dim tglDev2SEM = AddTraceToggle(gbDev2, "Dev 2 SEM", "SEM", Color.LimeGreen)
        Dim tglDev2PPM = AddTraceToggle(gbDev2, "Dev 2 PPM Deviation", "PPM Dev.", Color.LightGray)

        Dim tglTemp = AddTraceToggle(gbTemp, "Temperature", "Temp.", Color.Red)

        Dim dev1Boxes = {tglDevice1, tglDev1Mean, tglDev1Stdev, tglDev1SEM, tglDev1PPM}
        Dim dev2Boxes = {tglDevice2, tglDev2Mean, tglDev2Stdev, tglDev2SEM, tglDev2PPM}
        Dim tempBoxes = {tglTemp}

        Dim RefreshDeviceAvailability = Sub()
                                            Dim dev1Active As Boolean = (ButtonDev1Run.Text = "Stop") OrElse (ButtonDev12Run.Text = "Stop")
                                            Dim dev2Active As Boolean = (ButtonDev2Run.Text = "Stop") OrElse (ButtonDev12Run.Text = "Stop")
                                            Dim tempActive As Boolean = Timer1.Enabled

                                            Dim SetGroup = Sub(boxes As CheckBox(), active As Boolean)
                                                               For Each cb As CheckBox In boxes
                                                                   cb.Enabled = active
                                                                   If active Then
                                                                       cb.BackColor = CType(cb.Tag, Color)
                                                                   Else
                                                                       cb.Checked = False
                                                                       cb.BackColor = Color.LightGray
                                                                   End If
                                                               Next
                                                           End Sub

                                            SetGroup(dev1Boxes, dev1Active)
                                            SetGroup(dev2Boxes, dev2Active)
                                            SetGroup(tempBoxes, tempActive)
                                        End Sub

        RefreshDeviceAvailability()

        Dim RunButtonHandler = Sub(s As Object, ev As EventArgs) RefreshDeviceAvailability()

        AddHandler ButtonDev1Run.Click, RunButtonHandler
        AddHandler ButtonDev2Run.Click, RunButtonHandler
        AddHandler ButtonDev12Run.Click, RunButtonHandler
        AddHandler ButtonStart.Click, RunButtonHandler
        AddHandler ButtonEnd.Click, RunButtonHandler

        Dim RepositionLiveToggles = Sub()
                                        Dim cw As Double = LiveAnalysisChart.ClientSize.Width
                                        Dim ch As Double = LiveAnalysisChart.ClientSize.Height

                                        Dim rowHeightPx As Integer = 22
                                        Dim colWidthPx As Integer = 110
                                        Dim pad As Integer = 5         ' margin around the checkboxes
                                        Dim gap As Integer = 20        ' space between the three group boxes
                                        Dim topPct As Double = 0       ' % from top of form

                                        Dim GroupWidth = Function(boxes As CheckBox()) As Integer
                                                             Dim numCols As Integer = CInt(Math.Ceiling(boxes.Length / 2.0))
                                                             Return (colWidthPx * numCols) + (pad * 2)
                                                         End Function

                                        Dim widthDev1 As Integer = GroupWidth(dev1Boxes)
                                        Dim widthDev2 As Integer = GroupWidth(dev2Boxes)
                                        Dim widthTemp As Integer = GroupWidth(tempBoxes)

                                        Dim totalWidth As Integer = widthDev1 + widthDev2 + widthTemp + (gap * 2)
                                        Dim startX As Integer = CInt((cw - totalWidth) / 2.0)

                                        Dim leftDev1 As Integer = startX
                                        Dim leftDev2 As Integer = leftDev1 + widthDev1 + gap
                                        Dim leftTemp As Integer = leftDev2 + widthDev2 + gap

                                        Dim topPx As Integer = CInt(topPct / 100.0 * ch)

                                        Dim PlaceGroupBox = Sub(gb As GroupBox, boxes As CheckBox(), leftPx As Integer)
                                                                Dim numCols As Integer = CInt(Math.Ceiling(boxes.Length / 2.0))
                                                                Dim numRows As Integer = Math.Min(2, boxes.Length)
                                                                Dim titleAllowance As Integer = TextRenderer.MeasureText(gb.Text, gb.Font).Height + 2

                                                                Dim contentHeight As Integer = rowHeightPx * numRows

                                                                gb.Location = New Point(leftPx, topPx)
                                                                gb.Size = New Size((colWidthPx * numCols) + (pad * 2), contentHeight + (pad * 2) + titleAllowance)

                                                                Dim dr As Rectangle = gb.DisplayRectangle
                                                                Dim hOffset As Integer = dr.Left + pad
                                                                Dim vOffset As Integer = dr.Top + Math.Max(pad, (dr.Height - contentHeight) \ 2)

                                                                For i As Integer = 0 To boxes.Length - 1
                                                                    Dim col As Integer = i \ 2
                                                                    Dim row As Integer = i Mod 2
                                                                    boxes(i).Location = New Point(hOffset + (col * colWidthPx), vOffset + (row * rowHeightPx))
                                                                    boxes(i).Width = colWidthPx - 2
                                                                Next
                                                            End Sub

                                        PlaceGroupBox(gbDev1, dev1Boxes, leftDev1)
                                        PlaceGroupBox(gbDev2, dev2Boxes, leftDev2)
                                        PlaceGroupBox(gbTemp, tempBoxes, leftTemp)
                                    End Sub

        RepositionLiveToggles()
        AddHandler LiveAnalysisChart.Resize, Sub(s, ev) RepositionLiveToggles()




        LiveAnalysisForm.Controls.Add(LiveAnalysisChart)

        LiveAnalysisTimeLabel = New Label With {
.Dock = DockStyle.Bottom,
.Height = 22,
.TextAlign = ContentAlignment.MiddleCenter,
.Font = New Font("Segoe UI", 12),
.BackColor = Color.WhiteSmoke,
.ForeColor = Color.Black
}

        ' These still work fine with this approach - keep them
        LiveAnalysisChart.Series("Dev 1 PPM Deviation").LegendText = "Dev 1 PPM Dev"
        LiveAnalysisChart.Series("Dev 2 PPM Deviation").LegendText = "Dev 2 PPM Dev"




        LiveAnalysisForm.Controls.Add(LiveAnalysisTimeLabel)
        LiveAnalysisTimeLabel.BringToFront()

        AddHandler LiveAnalysisForm.FormClosed,
        Sub()
            RemoveHandler ButtonDev1Run.Click, RunButtonHandler
            RemoveHandler ButtonDev2Run.Click, RunButtonHandler
            RemoveHandler ButtonDev12Run.Click, RunButtonHandler

            LiveAnalysisChart = Nothing
            LiveAnalysisForm = Nothing
        End Sub


        ' Start chart at sample zero whenever opened.
        LiveAnalysisSample = 0

        ' Remember current statistics counts so that only NEW
        ' readings received after opening the chart are plotted.
        LiveAnalysisLastStats1Count = Stats1Count
        LiveAnalysisLastStats2Count = Stats2Count

        LiveAnalysisForm.Show()

    End Sub


    Private Sub AddLiveAnalysisSeries(seriesName As String,
                                  chartAreaName As String,
                                  seriesColor As Color)

        Dim s As New DataVisualization.Charting.Series(seriesName)

        ' Match existing Live Watch chart trace style.
        s.ChartType = DataVisualization.Charting.SeriesChartType.FastLine

        s.ChartArea = chartAreaName

        s.Color = seriesColor

        s.Legend = "Legend"

        LiveAnalysisChart.Series.Add(s)

    End Sub


    Private Sub UpdateLiveAnalysisChart()

        If LiveAnalysisForm Is Nothing Then Exit Sub
        If LiveAnalysisForm.IsDisposed Then Exit Sub
        If LiveAnalysisChart Is Nothing Then Exit Sub

        ' Determine whether each device has a fresh (unflushed)
        ' reading waiting since the last time it was plotted.
        Dim newDev1Sample As Boolean = CheckBoxStats1Enable.Checked = True AndAlso Stats1Count <> LiveAnalysisLastStats1Count
        Dim newDev2Sample As Boolean = CheckBoxStats2Enable.Checked = True AndAlso Stats2Count <> LiveAnalysisLastStats2Count

        ' Decide whether to advance the chart yet.
        Dim dev1CurrentlyRunning As Boolean = ButtonDev1Run.Text = "Stop" OrElse ButtonDev12Run.Text = "Stop"
        Dim dev2CurrentlyRunning As Boolean = ButtonDev2Run.Text = "Stop" OrElse ButtonDev12Run.Text = "Stop"

        Dim dev1Contributing As Boolean = dev1CurrentlyRunning AndAlso CheckBoxStats1Enable.Checked = True AndAlso Stats1Count > 0
        Dim dev2Contributing As Boolean = dev2CurrentlyRunning AndAlso CheckBoxStats2Enable.Checked = True AndAlso Stats2Count > 0

        Dim readyToAdvance As Boolean

        If dev1Contributing AndAlso dev2Contributing Then
            readyToAdvance = newDev1Sample AndAlso newDev2Sample
        Else
            readyToAdvance = newDev1Sample OrElse newDev2Sample
        End If

        If readyToAdvance = False Then Exit Sub

        ' Advance ONE X-axis sample
        LiveAnalysisSample += 1

        Dim x As Double = LiveAnalysisSample

        ' Device 1
        If newDev1Sample = True Then

            Dim dev1Raw As Double = CDbl(Val(NormalizeNumericResponse(txtr1a.Text)))

            LiveAnalysisChart.Series("Device 1").Points.AddXY(x, dev1Raw)
            LiveAnalysisChart.Series("Dev 1 Mean").Points.AddXY(x, Stats1Mean)
            LiveAnalysisChart.Series("Dev 1 STDEV").Points.AddXY(x, Stats1StdevCurrent)
            LiveAnalysisChart.Series("Dev 1 SEM").Points.AddXY(x, Stats1SEMCurrent)
            LiveAnalysisChart.Series("Dev 1 PPM Deviation").Points.AddXY(x, Stats1DeviationCurrent)

            LiveAnalysisLastStats1Count = Stats1Count

        End If

        ' Device 2
        If newDev2Sample = True Then

            Dim dev2Raw As Double = CDbl(Val(NormalizeNumericResponse(txtr2a.Text)))

            LiveAnalysisChart.Series("Device 2").Points.AddXY(x, dev2Raw)
            LiveAnalysisChart.Series("Dev 2 Mean").Points.AddXY(x, Stats2Mean)
            LiveAnalysisChart.Series("Dev 2 STDEV").Points.AddXY(x, Stats2StdevCurrent)
            LiveAnalysisChart.Series("Dev 2 SEM").Points.AddXY(x, Stats2SEMCurrent)
            LiveAnalysisChart.Series("Dev 2 PPM Deviation").Points.AddXY(x, Stats2DeviationCurrent)

            LiveAnalysisLastStats2Count = Stats2Count

        End If

        ' Temperature
        If newDev1Sample = True Or newDev2Sample = True Then

            ' Temperature
            Dim currentTemp As Double = gCurrTemp + Val(TempOffset.Text)
            LiveAnalysisChart.Series("Temperature").Points.AddXY(x, currentTemp)

            ' X-axis scale
            Dim laSampleRateText As String = ""

            If ButtonDev12Run.Text = "Stop" Then
                laSampleRateText = Dev12SampleRate.Text
            ElseIf ButtonDev1Run.Text = "Stop" Then
                laSampleRateText = Dev1SampleRate.Text
            ElseIf ButtonDev2Run.Text = "Stop" Then
                laSampleRateText = Dev2SampleRate.Text
            End If

            If laSampleRateText <> "" Then

                Dim laTotalSeconds As Integer = CInt(Val(laSampleRateText) * LiveAnalysisSample)
                Dim laHours As Integer = laTotalSeconds \ 3600
                Dim laMinutes As Integer = (laTotalSeconds Mod 3600) \ 60
                Dim laSeconds As Integer = laTotalSeconds Mod 60

                'LiveAnalysisTimeLabel.Text = "Visible Chart =  " & $"{laHours}hrs {laMinutes:00}mins {laSeconds:00}secs"

            End If

        End If

        ' X-axis behaviour
        If DisableRollingChartLiveA.Checked = False Then

            Dim windowN As Integer

            If Not Integer.TryParse(XaxisPointsLiveA.Text, windowN) OrElse windowN < 2 Then
                windowN = 100
            End If


            ' Remove points that have moved outside the rolling window.
            For Each s As DataVisualization.Charting.Series In LiveAnalysisChart.Series

                While s.Points.Count > 0 AndAlso s.Points(0).XValue < LiveAnalysisSample - windowN + 1
                    s.Points.RemoveAt(0)
                End While

            Next


            ' Keep newest reading at the RIGHT edge.
            ' This matches the normal Live Watch behaviour.
            Dim xmax As Double = LiveAnalysisSample
            Dim xmin As Double = xmax - (windowN - 1)


            ' Sample rate for converting sample number to elapsed time.
            Dim axisSampleRateText As String = ""

            If ButtonDev12Run.Text = "Stop" Then
                axisSampleRateText = Dev12SampleRate.Text
            ElseIf ButtonDev1Run.Text = "Stop" Then
                axisSampleRateText = Dev1SampleRate.Text
            ElseIf ButtonDev2Run.Text = "Stop" Then
                axisSampleRateText = Dev2SampleRate.Text
            End If

            Dim axisSampleRateSeconds As Double = Val(axisSampleRateText)


            For Each area As DataVisualization.Charting.ChartArea In LiveAnalysisChart.ChartAreas

                area.AxisX.Minimum = xmin
                area.AxisX.Maximum = xmax

                Dim domain As Double = windowN - 1
                If domain <= 0 Then domain = 10.0R

                area.AxisX.Interval = domain / 10.0R


                ' Replace numeric sample-index labels with elapsed time.
                area.AxisX.CustomLabels.Clear()

                If axisSampleRateSeconds > 0 Then

                    Dim tickCount As Integer = 10
                    Dim tickStep As Double = (xmax - xmin) / tickCount

                    For i As Integer = 0 To tickCount

                        Dim tickPos As Double = xmin + (i * tickStep)
                        Dim tickSeconds As Integer = CInt(Math.Max(tickPos, 0) * axisSampleRateSeconds)

                        Dim tickHours As Integer = tickSeconds \ 3600
                        Dim tickMinutes As Integer = (tickSeconds Mod 3600) \ 60
                        Dim tickSecs As Integer = tickSeconds Mod 60

                        Dim tickLabel As String = $"{tickHours:00}:{tickMinutes:00}:{tickSecs:00}"

                        Dim labelLow As Double = tickPos - (tickStep / 2)
                        Dim labelHigh As Double = tickPos + (tickStep / 2)

                        area.AxisX.CustomLabels.Add(labelLow, labelHigh, tickLabel)

                    Next

                End If

            Next

        Else

            ' Rolling disabled - retain all points, then convert the
            ' auto-resolved axis range into elapsed-time labels the
            ' same way as the rolling-enabled case.
            Dim disabledSampleRateText As String = ""

            If ButtonDev12Run.Text = "Stop" Then
                disabledSampleRateText = Dev12SampleRate.Text
            ElseIf ButtonDev1Run.Text = "Stop" Then
                disabledSampleRateText = Dev1SampleRate.Text
            ElseIf ButtonDev2Run.Text = "Stop" Then
                disabledSampleRateText = Dev2SampleRate.Text
            End If

            Dim disabledSampleRateSeconds As Double = Val(disabledSampleRateText)

            For Each area As DataVisualization.Charting.ChartArea In LiveAnalysisChart.ChartAreas

                area.AxisX.Minimum = Double.NaN
                area.AxisX.Maximum = Double.NaN
                area.AxisX.Interval = Double.NaN
                area.AxisX.CustomLabels.Clear()

                area.RecalculateAxesScale()

                Dim resolvedMin As Double = area.AxisX.Minimum
                Dim resolvedMax As Double = area.AxisX.Maximum

                If disabledSampleRateSeconds > 0 AndAlso resolvedMax > resolvedMin Then

                    Dim tickCount As Integer = 10
                    Dim tickStep As Double = (resolvedMax - resolvedMin) / tickCount

                    area.AxisX.Interval = tickStep

                    For i As Integer = 0 To tickCount

                        Dim tickPos As Double = resolvedMin + (i * tickStep)
                        Dim tickSeconds As Integer = CInt(Math.Max(tickPos, 0) * disabledSampleRateSeconds)

                        Dim tickHours As Integer = tickSeconds \ 3600
                        Dim tickMinutes As Integer = (tickSeconds Mod 3600) \ 60
                        Dim tickSecs As Integer = tickSeconds Mod 60

                        Dim tickLabel As String = $"{tickHours:00}:{tickMinutes:00}:{tickSecs:00}"

                        Dim labelLow As Double = tickPos - (tickStep / 2)
                        Dim labelHigh As Double = tickPos + (tickStep / 2)

                        area.AxisX.CustomLabels.Add(labelLow, labelHigh, tickLabel)

                    Next

                End If

            Next

        End If

        ' Measurement Y-axis autoscale
        Dim measurementMin As Double = Double.MaxValue
        Dim measurementMax As Double = Double.MinValue
        Dim measurementPointsFound As Boolean = False

        Dim measurementSeriesNames() As String = {"Device 1", "Device 2", "Dev 1 Mean", "Dev 2 Mean"}

        For Each seriesName As String In measurementSeriesNames

            Dim s As DataVisualization.Charting.Series = LiveAnalysisChart.Series(seriesName)

            If s.Enabled AndAlso s.Points.Count > 0 Then

                Dim sMin As Double = s.Points.Min(Function(p) p.YValues(0))
                Dim sMax As Double = s.Points.Max(Function(p) p.YValues(0))

                measurementMin = Math.Min(measurementMin, sMin)
                measurementMax = Math.Max(measurementMax, sMax)

                measurementPointsFound = True

            End If

        Next


        If measurementPointsFound = True Then

            Dim range As Double = measurementMax - measurementMin

            ' Protect against identical readings.
            If range <= 0 Then

                Dim buffer As Double = Math.Abs(measurementMin) * 0.000001
                If buffer = 0 Then buffer = 0.000000001

                measurementMin -= buffer
                measurementMax += buffer

                range = measurementMax - measurementMin

            End If


            ' Small margin above and below traces.
            Dim margin As Double = range * 0.05

            With LiveAnalysisChart.ChartAreas("Measurement").AxisY
                .Minimum = measurementMin - margin
                .Maximum = measurementMax + margin
                .Interval = ((measurementMax + margin) - (measurementMin - margin)) / 10.0R
            End With

        End If

        ' STDEV / SEM Y-axis autoscale (primary)
        With LiveAnalysisChart.ChartAreas("Statistics").AxisY
            .Minimum = Double.NaN
            .Maximum = Double.NaN
            .Interval = Double.NaN
        End With

        ' PPM Deviation Y-axis autoscale (secondary)
        With LiveAnalysisChart.ChartAreas("Statistics").AxisY2
            .Minimum = Double.NaN
            .Maximum = Double.NaN
            .Interval = Double.NaN
        End With

        LiveAnalysisChart.ChartAreas("Statistics").RecalculateAxesScale()
        EnsureValidAxisScale(LiveAnalysisChart.ChartAreas("Statistics").AxisY)
        EnsureValidAxisScale(LiveAnalysisChart.ChartAreas("Statistics").AxisY2)

        ' Temperature Y-axis autoscale
        With LiveAnalysisChart.ChartAreas("Temperature").AxisY
            .Minimum = Double.NaN
            .Maximum = Double.NaN
            .Interval = Double.NaN
        End With

        LiveAnalysisChart.ChartAreas("Temperature").RecalculateAxesScale()
        EnsureValidAxisScale(LiveAnalysisChart.ChartAreas("Temperature").AxisY)

    End Sub

    ' A ChartArea's RecalculateAxesScale() can leave Minimum/Maximum as NaN
    ' when every series feeding that axis is currently disabled (e.g. all
    ' Live Analysis trace checkboxes for an axis unchecked at once) - the
    ' next Paint then throws OverflowException trying to render against a
    ' NaN-bounded axis. Snap back to a safe placeholder range instead.
    Private Sub EnsureValidAxisScale(axis As DataVisualization.Charting.Axis)

        If Double.IsNaN(axis.Minimum) OrElse Double.IsNaN(axis.Maximum) OrElse
           Double.IsInfinity(axis.Minimum) OrElse Double.IsInfinity(axis.Maximum) OrElse
           axis.Minimum = axis.Maximum Then

            axis.Minimum = 0
            axis.Maximum = 1
            axis.Interval = 0.5

        End If

    End Sub


    Private Sub StatisticsEnable_CheckedChanged(sender As Object, e As EventArgs) _
    Handles CheckBoxStats1Enable.CheckedChanged,
            CheckBoxStats2Enable.CheckedChanged

        If CheckBoxStats1Enable.Checked = True Or CheckBoxStats2Enable.Checked = True Then

            ButtonLiveChartPopout.Enabled = True

        End If

        If CheckBoxStats1Enable.Checked = False And CheckBoxStats2Enable.Checked = False Then

            ButtonLiveChartPopout.Enabled = False

        End If

        'ButtonLiveChartPopout.Enabled = RunChart AndAlso (CheckBoxStats1Enable.Checked OrElse CheckBoxStats2Enable.Checked)

    End Sub


    Private Sub UpdateProjectedTimeLabel() Handles XaxisPointsLiveA.TextChanged,
                                                DisableRollingChartLiveA.CheckedChanged,
                                                ButtonDev1Run.Click, ButtonDev2Run.Click, ButtonDev12Run.Click,
                                                Dev1SampleRate.TextChanged, Dev2SampleRate.TextChanged, Dev12SampleRate.TextChanged

        If DisableRollingChartLiveA.Checked = True Then
            LabelXaxisProjectedTime.Text = "N/A (rolling disabled)"
            Exit Sub
        End If

        Dim sampleRateText As String = ""

        If ButtonDev12Run.Text = "Stop" Then
            sampleRateText = Dev12SampleRate.Text
        ElseIf ButtonDev1Run.Text = "Stop" Then
            sampleRateText = Dev1SampleRate.Text
        ElseIf ButtonDev2Run.Text = "Stop" Then
            sampleRateText = Dev2SampleRate.Text
        End If

        If sampleRateText = "" Then
            LabelXaxisProjectedTime.Text = "Not Running"
            Exit Sub
        End If

        Dim totalSeconds As Integer = CInt(Val(XaxisPointsLiveA.Text) * Val(sampleRateText))

        Dim hrs As Integer = totalSeconds \ 3600
        Dim mins As Integer = (totalSeconds Mod 3600) \ 60
        Dim secs As Integer = totalSeconds Mod 60

        LabelXaxisProjectedTime.Text = $"{hrs:00}:{mins:00}:{secs:00}"

    End Sub


End Class
