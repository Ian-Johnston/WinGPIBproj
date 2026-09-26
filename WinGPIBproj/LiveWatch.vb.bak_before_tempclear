' Live Watch

Imports System.Runtime.InteropServices

Partial Class Formtest

    ' Hands the resize-grip drag off to Windows' native bottom-right resize handling.
    <DllImport("user32.dll")>
    Private Shared Function ReleaseCapture() As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As Integer, lParam As Integer) As Integer
    End Function

    ' Number format: minimum 3 decimals, trailing zeros trimmed beyond that.
    ' Shared so the stats readouts match the big meter.
    Private Function BuildMinDpFormat(decimalPlaces As Integer) As String
        Dim minDp As Integer = Math.Min(3, decimalPlaces)
        Return "0." & New String("0"c, minDp) & New String("#"c, decimalPlaces - minDp)
    End Function



    Dim inst_value1FChart As Double = Double.NaN
    Dim inst_value2FChart As Double = Double.NaN
    Dim inst_value3FChart As Double
    Dim txtr1achart As String
    Dim txtr2achart As String
    Dim txtr3achart As String

    ' Chart1 data buffers/plottables. Each list is held by reference by its Scatter,
    ' so appending/trimming the list is enough before Refresh().
    ' X is a sample index that is never renumbered when old points are trimmed.
    Dim Chart1Dev1Data As New List(Of ScottPlot.Coordinates)
    Dim Chart1Dev2Data As New List(Of ScottPlot.Coordinates)
    Dim Chart1TempData As New List(Of ScottPlot.Coordinates)
    Dim Chart1Dev1NextX As Integer = 0
    Dim Chart1Dev2NextX As Integer = 0
    Dim Chart1TempNextX As Integer = 0
    Dim Chart1Dev1Series As ScottPlot.Plottables.Scatter
    Dim Chart1Dev2Series As ScottPlot.Plottables.Scatter
    Dim Chart1TempSeries As ScottPlot.Plottables.Scatter
    Dim Chart1TempAxis As ScottPlot.IYAxis
    Dim Chart1Crosshair As ScottPlot.Plottables.Crosshair
    Dim Chart1HighlightMarker As ScottPlot.Plottables.Marker
    Dim Chart1HighlightText As ScottPlot.Plottables.Text
    Dim Chart1LastRightClickPixel As ScottPlot.Pixel
    Dim Chart1LastLeftClickTime As DateTime = DateTime.MinValue
    Dim Chart1LastLeftClickPixel As ScottPlot.Pixel

    ' Two-point delta/measurement tool - see Chart1OnDoubleClick.
    Dim Chart1MeasureMarkerA As ScottPlot.Plottables.Marker
    Dim Chart1MeasureMarkerB As ScottPlot.Plottables.Marker
    Dim Chart1MeasureLine As ScottPlot.Plottables.LinePlot
    Dim Chart1MeasureText As ScottPlot.Plottables.Text
    Dim Chart1MeasureHavePointA As Boolean = False
    Dim Chart1MeasureHavePointB As Boolean = False
    Dim Chart1MeasurePointA As ScottPlot.DataPoint
    Dim Chart1MeasureAxisA As ScottPlot.IYAxis

    ' Data max/min found by autoscale (the view is padded beyond them); used for the Max/Min ticks.
    Dim Chart1DetectedMax As Double = Double.NaN
    Dim Chart1DetectedMin As Double = Double.NaN

    ' Decimal places on the left Y axis labels; the Max/Min boxes use the same.
    Private Const Chart1YDecimals As Integer = 8

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
    Private PauseLiveDisplay1 As Boolean = False
    Private LabelStats1SamplesNormalColor As Color = Color.Empty
    Private LabelStats1SamplesNormalBackColor As Color = Color.Empty

    ' Live Statistics - Device 2
    Private Stats2Count As Long = 0
    Private Stats2Mean As Double = 0.0
    Private Stats2M2 As Double = 0.0
    Private PauseLiveDisplay2 As Boolean = False
    Private LabelStats2SamplesNormalColor As Color = Color.Empty
    Private LabelStats2SamplesNormalBackColor As Color = Color.Empty

    ' Flashes the Samples label on whichever device(s) are display-paused,
    ' as a visual reminder that what's on screen is frozen (the underlying
    ' running stats keep updating regardless).
    Private WithEvents PauseFlashTimer As New Timer With {.Interval = 500}
    Private PauseFlashOn As Boolean = False

    ' Live Analysis Pop-out Chart
    Private LiveAnalysisForm As Form = Nothing
    Private LiveAnalysisChart As DataVisualization.Charting.Chart = Nothing
    Private LiveAnalysisTimeLabel As Label
    Private LiveAnalysisSample As Long = 0
    Private Stats1StdevCurrent As Double = 0.0
    Private Stats1SEMCurrent As Double = 0.0
    Private Stats2StdevCurrent As Double = 0.0
    Private Stats2SEMCurrent As Double = 0.0

    ' Short-Term Mean: display-only rolling average of the last few readings for the Mean trace.
    ' Does not affect Stats1Mean/Stats2Mean, STDEV/SEM/PPM Deviation or the CSV.
    Private chkShortTermMean As CheckBox = Nothing
    Private Const ShortTermMeanWindow As Integer = 30
    Private q1ShortTermMean As New Queue(Of Double)
    Private sum1ShortTermMean As Double = 0.0
    Private q2ShortTermMean As New Queue(Of Double)
    Private sum2ShortTermMean As Double = 0.0

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


    ' Appends a Y value to a series buffer and trims from the front past the window size,
    ' unless DisableRollingChart is checked.
    Private Sub Chart1AddPoint(data As List(Of ScottPlot.Coordinates), ByRef nextX As Integer, y As Double)

        data.Add(New ScottPlot.Coordinates(nextX, y))
        nextX += 1

        If DisableRollingChart.Checked = False Then

            Dim windowN As Integer
            If Not Integer.TryParse(XaxisPoints.Text, windowN) OrElse windowN < 2 Then windowN = 100

            If data.Count > windowN Then data.RemoveAt(0)

        End If

    End Sub

    ' Sets Chart1's primary Y range (autoscale path) and remembers the detected data max/min.
    Private Sub Chart1SetYAxisRange(minV As Double, maxV As Double)

        ' Small margin so the trace doesn't sit flush against the top/bottom
        ' edge of the plot, where a very stable/flat signal can be hard to see.
        Dim margin As Double = (maxV - minV) * 0.05

        FormsPlot1.Plot.Axes.Left.Min = minV - margin
        FormsPlot1.Plot.Axes.Left.Max = maxV + margin

        Chart1DetectedMax = maxV
        Chart1DetectedMin = minV
        Chart1EchoYRange()

    End Sub

    ' Keeps the Y-axis Max/Min boxes equal to the chart's current left-axis limits
    ' (autoscale, mouse pan/zoom, keys). Skipped while the user is typing in either box.
    Private Sub Chart1EchoYRange()

        If Dev1Max.Focused OrElse Dev1Min.Focused Then Exit Sub

        Dim lo As Double = FormsPlot1.Plot.Axes.Left.Min
        Dim hi As Double = FormsPlot1.Plot.Axes.Left.Max
        If Double.IsNaN(lo) OrElse Double.IsNaN(hi) OrElse Double.IsInfinity(lo) OrElse Double.IsInfinity(hi) OrElse hi <= lo Then Exit Sub

        Dim numberFormat As String = "F" & Chart1YDecimals.ToString()
        Dim maxText As String = hi.ToString(numberFormat, Globalization.CultureInfo.InvariantCulture)
        Dim minText As String = lo.ToString(numberFormat, Globalization.CultureInfo.InvariantCulture)
        If Dev1Max.Text <> maxText Then Dev1Max.Text = maxText
        If Dev1Min.Text <> minText Then Dev1Min.Text = minText

    End Sub

    ' Replaces the best hover candidate if this series' nearest point is closer in pixels.
    ' Pixel distance keeps Dev1/2 (left axis) and Temperature (right axis) comparable.
    Private Sub Chart1ConsiderHoverCandidate(series As ScottPlot.Plottables.Scatter, mouseLocation As ScottPlot.Coordinates,
                                              yAxis As ScottPlot.IYAxis, mousePixel As ScottPlot.Pixel,
                                              ByRef found As Boolean, ByRef bestPoint As ScottPlot.DataPoint,
                                              ByRef bestYAxis As ScottPlot.IYAxis, ByRef bestColor As ScottPlot.Color,
                                              ByRef bestDistance As Single)

        ' Hidden traces can't be hovered or measured.
        If Not series.IsVisible Then Exit Sub

        ' Series.Data.GetNearest ignores non-primary axes, so call the utility directly with the real axes.
        Dim dataSource As ScottPlot.IDataSource = DirectCast(series.Data, ScottPlot.IDataSource)
        Dim point As ScottPlot.DataPoint = ScottPlot.DataSourceUtilities.GetNearestSmart(
            dataSource, mouseLocation, FormsPlot1.Plot.LastRender, 15, FormsPlot1.Plot.Axes.Bottom, yAxis)
        If Not point.IsReal Then Exit Sub

        Dim pointPixel As ScottPlot.Pixel = FormsPlot1.Plot.GetPixel(point.Coordinates, FormsPlot1.Plot.Axes.Bottom, yAxis)
        Dim distance As Single = pointPixel.DistanceFrom(mousePixel)

        If distance < bestDistance Then
            found = True
            bestPoint = point
            bestYAxis = yAxis
            bestColor = series.LineStyle.Color
            bestDistance = distance
        End If

    End Sub

    ' Finds whichever of Chart1's three traces has a point nearest the given
    ' pixel (shared by the hover tooltip and the "Copy Value At Cursor" menu
    ' action).
    Private Sub Chart1FindNearestPoint(mousePixel As ScottPlot.Pixel, ByRef found As Boolean, ByRef bestPoint As ScottPlot.DataPoint,
                                        ByRef bestYAxis As ScottPlot.IYAxis, ByRef bestColor As ScottPlot.Color)

        Dim mouseLocationPrimary As ScottPlot.Coordinates = FormsPlot1.Plot.GetCoordinates(mousePixel)
        Dim mouseLocationTemp As ScottPlot.Coordinates = FormsPlot1.Plot.GetCoordinates(mousePixel, FormsPlot1.Plot.Axes.Bottom, Chart1TempAxis)

        Dim bestDistance As Single = Single.MaxValue

        Chart1ConsiderHoverCandidate(Chart1Dev1Series, mouseLocationPrimary, FormsPlot1.Plot.Axes.Left, mousePixel, found, bestPoint, bestYAxis, bestColor, bestDistance)
        Chart1ConsiderHoverCandidate(Chart1Dev2Series, mouseLocationPrimary, FormsPlot1.Plot.Axes.Left, mousePixel, found, bestPoint, bestYAxis, bestColor, bestDistance)
        Chart1ConsiderHoverCandidate(Chart1TempSeries, mouseLocationTemp, Chart1TempAxis, mousePixel, found, bestPoint, bestYAxis, bestColor, bestDistance)

    End Sub

    ' Moves the crosshair/marker/text label to whichever of Chart1's three
    ' traces has a point nearest the mouse, or hides them when nothing is
    ' close enough.
    Private Sub Chart1ShowValueOnHover(sender As Object, e As MouseEventArgs)

        Chart1EchoYRange()

        Dim mousePixel As New ScottPlot.Pixel(CSng(e.X), CSng(e.Y))

        Dim found As Boolean = False
        Dim bestPoint As ScottPlot.DataPoint = Nothing
        Dim bestYAxis As ScottPlot.IYAxis = Nothing
        Dim bestColor As ScottPlot.Color = Nothing

        Chart1FindNearestPoint(mousePixel, found, bestPoint, bestYAxis, bestColor)

        If Not found Then
            If Chart1Crosshair.IsVisible Then
                Chart1Crosshair.IsVisible = False
                Chart1HighlightMarker.IsVisible = False
                Chart1HighlightText.IsVisible = False
                FormsPlot1.Refresh()
            End If
            Exit Sub
        End If

        Chart1Crosshair.IsVisible = True
        Chart1Crosshair.Position = bestPoint.Coordinates
        Chart1Crosshair.Axes.YAxis = bestYAxis
        Chart1Crosshair.LineColor = bestColor

        Chart1HighlightMarker.IsVisible = True
        Chart1HighlightMarker.Location = bestPoint.Coordinates
        Chart1HighlightMarker.Axes.YAxis = bestYAxis
        Chart1HighlightMarker.MarkerStyle.LineColor = bestColor

        Chart1HighlightText.IsVisible = True
        Chart1HighlightText.Location = bestPoint.Coordinates
        Chart1HighlightText.Axes.YAxis = bestYAxis
        Chart1HighlightText.LabelText = bestPoint.Y.ToString("0.########")
        Chart1HighlightText.LabelFontColor = bestColor

        ' Flip the label to whichever side of the point keeps it inside the
        ' plot area, instead of always drawing above-right (which runs off
        ' the top near the top edge, or off the right near the right edge).
        Const edgeMarginPx As Single = 40

        Dim bestPixel As ScottPlot.Pixel = FormsPlot1.Plot.GetPixel(bestPoint.Coordinates, FormsPlot1.Plot.Axes.Bottom, bestYAxis)
        Dim dataRect As ScottPlot.PixelRect = FormsPlot1.Plot.LastRender.DataRect

        Dim nearTop As Boolean = (bestPixel.Y - dataRect.Top) < edgeMarginPx
        Dim nearRight As Boolean = (dataRect.Right - bestPixel.X) < edgeMarginPx

        Chart1HighlightText.OffsetY = If(nearTop, 7, -7)
        Chart1HighlightText.OffsetX = If(nearRight, -7, 7)

        Chart1HighlightText.LabelAlignment =
            If(nearTop,
               If(nearRight, ScottPlot.Alignment.UpperRight, ScottPlot.Alignment.UpperLeft),
               If(nearRight, ScottPlot.Alignment.LowerRight, ScottPlot.Alignment.LowerLeft))

        ' While B isn't locked, live-preview the delta against the nearest point.
        If Chart1MeasureHavePointA AndAlso Not Chart1MeasureHavePointB Then
            Chart1UpdateMeasureDisplay(bestPoint, bestYAxis)
        End If

        FormsPlot1.Refresh()

    End Sub

    ' Any mouse-down unchecks AutoScale Y-axis, records the right-click position for the menu,
    ' and detects left double-clicks by timing consecutive mouse-downs (see Chart1OnDoubleClick).
    Private Sub Chart1OnMouseDown(sender As Object, e As MouseEventArgs)

        Chart1AutoScaleYAxis.Checked = False

        If e.Button = MouseButtons.Right Then
            Chart1LastRightClickPixel = New ScottPlot.Pixel(CSng(e.X), CSng(e.Y))
        End If

        If e.Button = MouseButtons.Left Then

            Dim thisPixel As New ScottPlot.Pixel(CSng(e.X), CSng(e.Y))
            Dim elapsedMs As Double = (DateTime.Now - Chart1LastLeftClickTime).TotalMilliseconds
            Dim dx As Single = thisPixel.X - Chart1LastLeftClickPixel.X
            Dim dy As Single = thisPixel.Y - Chart1LastLeftClickPixel.Y
            Dim distance As Single = CSng(Math.Sqrt(dx * dx + dy * dy))

            If elapsedMs <= SystemInformation.DoubleClickTime AndAlso
               distance <= SystemInformation.DoubleClickSize.Width Then

                ' Consume it, rather than leaving this click available to
                ' pair with a third - so 4 rapid clicks are two separate
                ' double-clicks, not three overlapping ones.
                Chart1LastLeftClickTime = DateTime.MinValue
                Chart1OnDoubleClick(thisPixel)

            Else

                Chart1LastLeftClickTime = DateTime.Now
                Chart1LastLeftClickPixel = thisPixel

            End If

        End If

    End Sub

    ' "Copy Value At Cursor" menu action: copies the Y value of the point nearest the right-click.
    Private Sub Chart1CopyValueAtCursor(plot As ScottPlot.Plot)

        Dim found As Boolean = False
        Dim bestPoint As ScottPlot.DataPoint = Nothing
        Dim bestYAxis As ScottPlot.IYAxis = Nothing
        Dim bestColor As ScottPlot.Color = Nothing

        Chart1FindNearestPoint(Chart1LastRightClickPixel, found, bestPoint, bestYAxis, bestColor)

        If found Then
            Clipboard.SetText(bestPoint.Y.ToString("0.########", Globalization.CultureInfo.InvariantCulture))
        End If

    End Sub

    ' Two-point measurement: 1st double-click sets A, 2nd sets B and locks the delta, 3rd clears.
    ' Detected in Chart1OnMouseDown by timing, as the DoubleClick event fires after the button is released.
    Private Sub Chart1OnDoubleClick(mousePixel As ScottPlot.Pixel)

        Dim found As Boolean = False
        Dim bestPoint As ScottPlot.DataPoint = Nothing
        Dim bestYAxis As ScottPlot.IYAxis = Nothing
        Dim bestColor As ScottPlot.Color = Nothing

        Chart1FindNearestPoint(mousePixel, found, bestPoint, bestYAxis, bestColor)

        If Chart1MeasureHavePointB Then

            Chart1ClearMeasurement(FormsPlot1.Plot)

        ElseIf Chart1MeasureHavePointA Then

            ' No point nearby to lock in as B - clear instead of leaving the measurement stuck.
            If Not found Then
                Chart1ClearMeasurement(FormsPlot1.Plot)
                Exit Sub
            End If

            Chart1MeasureHavePointB = True

            Chart1MeasureMarkerB.IsVisible = True
            Chart1MeasureMarkerB.Location = bestPoint.Coordinates
            Chart1MeasureMarkerB.Axes.YAxis = bestYAxis

            Chart1UpdateMeasureDisplay(bestPoint, bestYAxis)

            FormsPlot1.Refresh()

        Else

            If Not found Then Exit Sub

            Chart1MeasureHavePointA = True
            Chart1MeasurePointA = bestPoint
            Chart1MeasureAxisA = bestYAxis

            Chart1MeasureMarkerA.IsVisible = True
            Chart1MeasureMarkerA.Location = bestPoint.Coordinates
            Chart1MeasureMarkerA.Axes.YAxis = bestYAxis

            FormsPlot1.Refresh()

        End If

    End Sub

    ' Updates the A-B line and label. Line and delta only when both points share a Y axis,
    ' otherwise just the two raw values. Used by the hover preview and the locked B.
    Private Sub Chart1UpdateMeasureDisplay(pointB As ScottPlot.DataPoint, axisB As ScottPlot.IYAxis)

        Dim sameAxis As Boolean = axisB Is Chart1MeasureAxisA

        Chart1MeasureLine.IsVisible = sameAxis
        If sameAxis Then
            Chart1MeasureLine.Axes.YAxis = axisB
            Chart1MeasureLine.Start = Chart1MeasurePointA.Coordinates
            Chart1MeasureLine.[End] = pointB.Coordinates
        End If

        Chart1MeasureText.IsVisible = True
        Chart1MeasureText.Location = pointB.Coordinates
        Chart1MeasureText.Axes.YAxis = axisB

        If sameAxis Then
            Dim deltaX As Double = pointB.X - Chart1MeasurePointA.X
            Dim deltaY As Double = pointB.Y - Chart1MeasurePointA.Y
            Chart1MeasureText.LabelText = "dY " & deltaY.ToString("0.########") & "   dX " & deltaX.ToString("0") & " samples"
        Else
            Chart1MeasureText.LabelText = "A " & Chart1MeasurePointA.Y.ToString("0.########") & "   B " & pointB.Y.ToString("0.########")
        End If

        ' Flip the label to keep it inside the plot; wider right margin as this text is longer.
        Const edgeMarginTopPx As Single = 40
        Const edgeMarginRightPx As Single = 300

        Dim pointBPixel As ScottPlot.Pixel = FormsPlot1.Plot.GetPixel(pointB.Coordinates, FormsPlot1.Plot.Axes.Bottom, axisB)
        Dim dataRect As ScottPlot.PixelRect = FormsPlot1.Plot.LastRender.DataRect

        Dim nearTop As Boolean = (pointBPixel.Y - dataRect.Top) < edgeMarginTopPx
        Dim nearRight As Boolean = (dataRect.Right - pointBPixel.X) < edgeMarginRightPx

        Chart1MeasureText.OffsetY = If(nearTop, 7, -7)
        Chart1MeasureText.OffsetX = If(nearRight, -7, 7)

        Chart1MeasureText.LabelAlignment =
            If(nearTop,
               If(nearRight, ScottPlot.Alignment.UpperRight, ScottPlot.Alignment.UpperLeft),
               If(nearRight, ScottPlot.Alignment.LowerRight, ScottPlot.Alignment.LowerLeft))

    End Sub

    ' Clears the measurement tool. The Plot parameter is unused, it only matches the menu delegate.
    Private Sub Chart1ClearMeasurement(plot As ScottPlot.Plot)

        Chart1MeasureHavePointA = False
        Chart1MeasureHavePointB = False

        Chart1MeasureMarkerA.IsVisible = False
        Chart1MeasureMarkerB.IsVisible = False
        Chart1MeasureLine.IsVisible = False
        Chart1MeasureText.IsVisible = False

        FormsPlot1.Refresh()

    End Sub

    ' Esc clears the measurement tool.
    Private Sub Chart1OnKeyDown(sender As Object, e As KeyEventArgs)

        If e.KeyCode = Keys.Escape Then
            Chart1ClearMeasurement(FormsPlot1.Plot)
        End If

        Chart1EchoYRange()

    End Sub

    ' Builds evenly spaced manual ticks across the axis range and applies them immediately.
    ' Returns the generator so callers can add ticks.
    Private Function Chart1SetFixedDivisionTicks(axis As ScottPlot.IAxis, divisions As Integer, decimals As Integer, edge As ScottPlot.Edge, rp As ScottPlot.RenderPack) As ScottPlot.TickGenerators.NumericManual

        Dim span As Double = axis.Max - axis.Min
        If span = 0 OrElse divisions <= 0 Then Return Nothing

        Dim numberFormat As String = "0." & New String("0"c, decimals)
        Dim stepValue As Double = span / divisions

        Dim positions(divisions) As Double
        Dim labels(divisions) As String

        For i As Integer = 0 To divisions

            Dim tickValue As Double = axis.Min + i * stepValue
            positions(i) = tickValue
            labels(i) = tickValue.ToString(numberFormat)

        Next

        Dim manualTicks As New ScottPlot.TickGenerators.NumericManual(positions, labels)
        manualTicks.Regenerate(New ScottPlot.CoordinateRange(axis.Min, axis.Max), edge,
                                New ScottPlot.PixelLength(100), rp.Paint, axis.TickLabelStyle)

        axis.TickGenerator = manualTicks

        Return manualTicks

    End Function

    ' Fixed 12-division grid; Temperature's tick labels are mapped to the same gridlines as the left axis.
    Private Sub Chart1RenderStarting(sender As Object, rp As ScottPlot.RenderPack)

        ' Mouse pan/zoom also moves the Temperature axis; snap it back to its Max/Min boxes before drawing.
        Dim tempMin As Double = Val(LCTempMin.Text)
        Dim tempMax As Double = Val(LCTempMax.Text)
        If tempMax > tempMin Then
            Chart1TempAxis.Min = tempMin
            Chart1TempAxis.Max = tempMax
        End If

        Const divisions As Integer = 12

        Dim leftAxis As ScottPlot.IAxis = FormsPlot1.Plot.Axes.Left
        Dim leftTicks = Chart1SetFixedDivisionTicks(leftAxis, divisions, Chart1YDecimals, ScottPlot.Edge.Left, rp)

        ' Extra ticks at the detected Dev1 max/min (autoscale only, as the view is padded beyond them).
        If leftTicks IsNot Nothing AndAlso Chart1AutoScaleYAxis.Checked Then

            Dim maxVal As Double
            Dim minVal As Double

            maxVal = Chart1DetectedMax
            If Not Double.IsNaN(maxVal) AndAlso maxVal >= leftAxis.Min AndAlso maxVal <= leftAxis.Max Then
                leftTicks.AddMajor(maxVal, "Max " & maxVal.ToString("0.########"))
            End If

            minVal = Chart1DetectedMin
            If Not Double.IsNaN(minVal) AndAlso minVal >= leftAxis.Min AndAlso minVal <= leftAxis.Max Then
                leftTicks.AddMajor(minVal, "Min " & minVal.ToString("0.########"))
            End If

            leftTicks.Regenerate(New ScottPlot.CoordinateRange(leftAxis.Min, leftAxis.Max), ScottPlot.Edge.Left,
                                  New ScottPlot.PixelLength(100), rp.Paint, leftAxis.TickLabelStyle)

        End If

        Dim leftSpan As Double = leftAxis.Max - leftAxis.Min
        If leftSpan = 0 Then Exit Sub

        Dim tempSpan As Double = Chart1TempAxis.Max - Chart1TempAxis.Min

        Dim tempPositions(divisions) As Double
        Dim tempLabels(divisions) As String

        For i As Integer = 0 To divisions

            Dim fraction As Double = i / divisions
            Dim tempValue As Double = Chart1TempAxis.Min + fraction * tempSpan

            tempPositions(i) = tempValue
            tempLabels(i) = tempValue.ToString("0.0")

        Next

        Dim manualTempTicks As New ScottPlot.TickGenerators.NumericManual(tempPositions, tempLabels)
        manualTempTicks.Regenerate(New ScottPlot.CoordinateRange(Chart1TempAxis.Min, Chart1TempAxis.Max),
                                    ScottPlot.Edge.Right, New ScottPlot.PixelLength(100), rp.Paint, Chart1TempAxis.TickLabelStyle)

        Chart1TempAxis.TickGenerator = manualTempTicks

    End Sub


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
            Chart1AddPoint(Chart1Dev1Data, Chart1Dev1NextX, Val(txtr1achart))

            ' Chart 3 - Temperature
            If (EnableChart3.Checked = True And RunChart = True) Then
                ' set up max and min for temperature
                If Chart1AutoScaleYAxis.Checked AndAlso Val(LCTempMax.Text) > Val(LCTempMin.Text) Then
                    UpdateChartTemperatureYAxisMinMaxInterval()
                End If

                inst_value3FChart = gCurrTemp
                inst_value3FChart += Val(TempOffset.Text)    ' integrate offset
                txtr3achart = Format(inst_value3FChart, "#0.00000000")

                ' plot to chart
                Chart1AddPoint(Chart1TempData, Chart1TempNextX, Val(txtr3achart))

                ' Temp - record min & max for display (resettable)
                Resetmaxdiffrecorded_temp()
            End If

            If (EnableChart3.Checked = False And RunChart = True) Then      ' dummy data so chart vertical data can align if temperature is checked later
                Chart1AddPoint(Chart1TempData, Chart1TempNextX, 0.0)
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
            Chart1AddPoint(Chart1Dev2Data, Chart1Dev2NextX, Val(txtr2achart))

            ' Chart 3 - Temperature
            If (EnableChart3.Checked = True And RunChart = True) Then
                ' set up max and min for temperature
                If Chart1AutoScaleYAxis.Checked AndAlso Val(LCTempMax.Text) > Val(LCTempMin.Text) Then
                    UpdateChartTemperatureYAxisMinMaxInterval()
                End If

                inst_value3FChart = gCurrTemp
                inst_value3FChart += Val(TempOffset.Text)   ' integrate offset
                txtr3achart = Format(inst_value3FChart, "#0.00000000")

                ' plot to chart
                Chart1AddPoint(Chart1TempData, Chart1TempNextX, Val(txtr3achart))

                ' Temp - record min & max for display (resettable)
                Resetmaxdiffrecorded_temp()
            End If

            If (EnableChart3.Checked = False And RunChart = True) Then      ' dummy data so chart vertical data can align if temperature is checked later
                Chart1AddPoint(Chart1TempData, Chart1TempNextX, 0.0)
            End If

        End If


        ' Chart 1 & 2 - Device 1 & Device 2
        If (EnableChart1.Checked = True And EnableChart2.Checked = True And RunChart = True And Dev2GPIBActivity = True) Then

            ' Device 1

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


            ' Device 2

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


            ' Plot Device 1 & Device 2 to normal Live Watch chart

            Chart1AddPoint(Chart1Dev1Data, Chart1Dev1NextX, Val(txtr1achart))
            Chart1AddPoint(Chart1Dev2Data, Chart1Dev2NextX, Val(txtr2achart))


            ' Chart 3 - Temperature

            If (EnableChart3.Checked = True And RunChart = True) Then

                ' Set up max and min for temperature
                If Chart1AutoScaleYAxis.Checked AndAlso Val(LCTempMax.Text) > Val(LCTempMin.Text) Then

                    UpdateChartTemperatureYAxisMinMaxInterval()

                End If

                inst_value3FChart = gCurrTemp
                inst_value3FChart += Val(TempOffset.Text)

                txtr3achart =
        Format(inst_value3FChart, "#0.00000000")


                ' Plot temperature
                Chart1AddPoint(Chart1TempData, Chart1TempNextX, Val(txtr3achart))

                ' Temp - record min & max for display (resettable)
                Resetmaxdiffrecorded_temp()

            End If


            ' Dummy data so chart vertical data can align
            ' if temperature is enabled later
            If (EnableChart3.Checked = False And RunChart = True) Then

                Chart1AddPoint(Chart1TempData, Chart1TempNextX, 0.0)

            End If

        End If


        ' Fixed X window so trace appears
        ' at the right and scrolls left
        ' Skipped while the user has panned/zoomed manually (AutoScale unchecked); new points still draw.
        ' With rolling disabled the X axis always fits the whole trace, whatever the AutoScale setting.
        If Chart1AutoScaleYAxis.Checked = False AndAlso DisableRollingChart.Checked = False Then

            ' Keep XaxisPoints in sync with the manually set view width.
            Dim currentSpan As Double = FormsPlot1.Plot.Axes.Bottom.Max - FormsPlot1.Plot.Axes.Bottom.Min
            If currentSpan > 0 Then
                XaxisPoints.Text = CInt(Math.Round(currentSpan)).ToString()
            End If

        ElseIf DisableRollingChart.Checked = False Then

            ' How many points wide should the visible window be?
            Dim windowN As Integer
            If Not Integer.TryParse(XaxisPoints.Text, windowN) OrElse windowN < 2 Then
                windowN = 100
            End If

            ' Use whichever series has advanced the furthest as "now"
            Dim lastIndex As Integer = Math.Max(Chart1Dev1NextX, Math.Max(Chart1Dev2NextX, Chart1TempNextX)) - 1

            If lastIndex >= 0 Then

                Dim window As Integer = windowN - 1

                Dim xmin As Double = lastIndex - window
                Dim xmax As Double = lastIndex

                FormsPlot1.Plot.Axes.SetLimitsX(xmin, xmax)

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
                If liveChartSampleRateSeconds > 0 Then

                    Dim tickCount As Integer = 10
                    Dim tickStep As Double = (xmax - xmin) / tickCount
                    Dim tickPositions(tickCount) As Double
                    Dim tickLabels(tickCount) As String

                    For i As Integer = 0 To tickCount

                        Dim tickPos As Double = xmin + (i * tickStep)
                        Dim tickSeconds As Integer = CInt(Math.Max(tickPos, 0) * liveChartSampleRateSeconds)

                        Dim tickHours As Integer = tickSeconds \ 3600
                        Dim tickMinutes As Integer = (tickSeconds Mod 3600) \ 60
                        Dim tickSecs As Integer = tickSeconds Mod 60

                        tickPositions(i) = tickPos
                        tickLabels(i) = $"{tickHours:00}:{tickMinutes:00}:{tickSecs:00}"

                    Next

                    FormsPlot1.Plot.Axes.Bottom.SetTicks(tickPositions, tickLabels)

                End If

            Else
                ' No points yet – let ScottPlot decide
                FormsPlot1.Plot.Axes.AutoScaleX()
            End If

        Else
            ' Rolling disabled – let the X axis grow to show all data
            FormsPlot1.Plot.Axes.AutoScaleX()

            Dim disabledMin As Double = FormsPlot1.Plot.Axes.Bottom.Min
            Dim disabledMax As Double = FormsPlot1.Plot.Axes.Bottom.Max

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
                Dim tickPositions(tickCount) As Double
                Dim tickLabels(tickCount) As String

                For i As Integer = 0 To tickCount

                    Dim tickPos As Double = disabledMin + (i * tickStep)
                    Dim tickSeconds As Integer = CInt(Math.Max(tickPos, 0) * liveChartSampleRateSecondsDisabled)

                    Dim tickHours As Integer = tickSeconds \ 3600
                    Dim tickMinutes As Integer = (tickSeconds Mod 3600) \ 60
                    Dim tickSecs As Integer = tickSeconds Mod 60

                    tickPositions(i) = tickPos
                    tickLabels(i) = $"{tickHours:00}:{tickMinutes:00}:{tickSecs:00}"

                Next

                FormsPlot1.Plot.Axes.Bottom.SetTicks(tickPositions, tickLabels)

            End If

        End If

        FormsPlot1.Refresh()

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
    "LIVE CHART STATISTICS" & vbLf &
    "WinGPIB can calculate live statistics independently for Device 1 and Device 2 using each raw incoming measurement reading." & vbLf &
    "Statistics run automatically whenever a device is actively acquiring and its Enable Statistics checkbox is checked - independent of whether the Live Chart or Data Log/CSV logging is running." & vbLf &
    "The Enable Statistics checkbox for each device is located in that device's configuration box on the Cmd Line tab, alongside its other settings. It can only be changed while the device is stopped - tick it before pressing Run, as it locks once the device starts." & vbLf &
    "The optional Live Chart rolling average does not affect these statistics." & vbLf & vbLf &
    "SAMPLES (N)" & vbLf &
    "The total number of individual readings included in the current statistics calculation." & vbLf &
    "The sample count starts from zero when RESET STAT is pressed." & vbLf & vbLf &
    "MEAN" & vbLf &
    "The arithmetic average of all readings collected since the statistics were started or reset." & vbLf & vbLf &
    "MAX / MIN RECORDED" & vbLf &
    "The highest and lowest individual raw readings seen since the statistics were started or reset." & vbLf &
    "These update on every new reading and are cleared back to a fresh state by RESET STAT." & vbLf & vbLf &
    "STDEV - STANDARD DEVIATION" & vbLf &
    "Shows how much the individual readings vary or scatter around the calculated mean." & vbLf &
    "A smaller STDEV generally indicates less variation or noise in the readings." & vbLf & vbLf &
    "WinGPIB calculates sample standard deviation using Welford's running algorithm. This is mathematically equivalent to the conventional sample STDEV calculation but avoids having to store every individual reading." & vbLf & vbLf &
    "Sample STDEV = Sqrt(Sum((Xi - Mean)^2) / (N - 1))" & vbLf & vbLf &
    "SEM - STANDARD ERROR OF THE MEAN" & vbLf &
    "Shows how precisely the mean has been determined from the accumulated readings." & vbLf &
    "SEM is derived from the STDEV and decreases as more independent readings are averaged." & vbLf & vbLf &
    "SEM = STDEV / Sqrt(N)" & vbLf & vbLf &
    "AVERAGING GAIN (DIGITS)" & vbLf &
    "Shows the theoretical increase in resolution obtained by averaging N independent readings." & vbLf & vbLf &
    "Averaging Gain = 0.5 x Log10(N)" & vbLf & vbLf &
    "Examples:" & vbLf &
    "  10 readings     = 0.50 digits" & vbLf &
    "  100 readings    = 1.00 digits" & vbLf &
    "  1,000 readings  = 1.50 digits" & vbLf &
    "  10,000 readings = 2.00 digits" & vbLf & vbLf &
    "LIVE ANALYSIS CHART" & vbLf &
    "The Live Analysis chart displays the raw Device 1 and Device 2 readings together with their running Mean, STDEV and SEM, plus temperature when enabled." & vbLf &
    "The analysis chart uses the same raw readings as the statistics calculations and is not affected by the optional Live Chart rolling average." & vbLf &
    "The Live Analysis chart does not require the Live Chart to be started - it runs from the same live statistics as soon as a device is running with Enable Statistics checked." & vbLf &
    "When both Device 1 and Device 2 are running together, the chart advances once per matched pair of readings rather than once per device, so the two devices share a common position on the chart instead of doubling the update rate." & vbLf &
    "The X-axis shows elapsed time (HH:mm:ss), calculated from the sample rate of whichever device(s) are running." & vbLf & vbLf &
    "LIVE ANALYSIS CHART - MISC. OPTIONS" & vbLf &
    "ANTI-ALIASING - Smooths lines and text on this chart and the main Live Chart. Keep this on to avoid a jagged/moire look on traces with many closely-packed points; only turn it off to compare against unsmoothed rendering." & vbLf &
    "FAST RENDERING - Switches every trace on this chart to FastLine, a stripped-down renderer built for very large point counts that skips anti-aliasing entirely regardless of the Anti-Aliasing setting above. Only useful if the chart becomes slow with a very large or unbounded rolling window. Overrides Smooth Lines while checked." & vbLf &
    "SMOOTH LINES - Draws each trace as a curved spline between points instead of straight segments. Purely cosmetic - a curve can visually suggest values in between samples that were never actually measured, so treat it as a display preference, not a data change." & vbLf &
    $"SHORT-TERM MEAN - Plots the Mean trace as a rolling average of only the last {ShortTermMeanWindow} readings instead of the full cumulative Mean since Reset Stats. Responds faster to recent changes but is noisier. This affects the Mean TRACE on this chart only - it does not change the Mean shown in the DEVICE DATA panel, STDEV/SEM/PPM Deviation, or what is written to the CSV log. Checking or unchecking it only affects the trace from that moment onwards - points already plotted are not redrawn, so you will see a kink in the trace at the point you toggled it." & vbLf & vbLf &
    "DATA LOG / CSV" & vbLf &
    "When statistics are enabled, the current Samples, Mean, STDEV, SEM and Averaging Gain values are also available in the Data Log and CSV output." & vbLf &
    "Statistics fields remain in fixed positions for Device 1 and Device 2. When statistics are disabled for a device, those fields are left blank." & vbLf & vbLf &
    "NOTES" & vbLf &
    "• The statistics are calculated independently for Device 1 and Device 2." & vbLf &
    "• RESET STAT clears the accumulated statistics for that device and starts again from zero." & vbLf &
    "• Statistics are calculated from individual raw readings before any optional Live Chart rolling averaging is applied." & vbLf &
    "• Statistics run whenever the device is active and Enable Statistics is checked - the Live Chart and Data Log/CSV logging do not need to be running." & vbLf &
    "• The Enable Statistics checkbox locks while its device is running - enable it before pressing Run, not after." & vbLf &
    "• STDEV includes all variation present in the readings, including random noise, drift, temperature effects and other changes." & vbLf &
    "• SEM is most meaningful when the readings are independent and the underlying measured value is stable." & vbLf &
    "• A small SEM does not by itself represent the total measurement uncertainty or accuracy." & vbLf &
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
    "LIVE ANALYSIS CHART - MISC. OPTIONS",
    "DATA LOG / CSV",
    "NOTES"
}

        For Each heading As String In headings

            Dim start As Integer =
            txt.Text.IndexOf(heading, StringComparison.Ordinal)

            If start >= 0 Then
                txt.Select(start, heading.Length)
                txt.SelectionFont = New Font(txt.Font, FontStyle.Bold)
            End If

        Next

        ' Indents section body text with SelectionIndent so wrapped lines stay aligned.
        For i As Integer = 0 To headings.Length - 1

            Dim headingStart As Integer =
            txt.Text.IndexOf(headings(i), StringComparison.Ordinal)

            If headingStart < 0 Then Continue For

            ' +1 skips past the vbLf that ends the heading's own line -
            ' starting the selection exactly on that character would still
            ' count as touching the heading's paragraph, indenting it too.
            Dim bodyStart As Integer = headingStart + headings(i).Length + 1
            Dim bodyEnd As Integer = txt.Text.Length

            If i < headings.Length - 1 Then
                Dim nextHeadingStart As Integer =
                txt.Text.IndexOf(headings(i + 1), StringComparison.Ordinal)
                If nextHeadingStart > bodyStart Then bodyEnd = nextHeadingStart
            End If

            If bodyEnd > bodyStart Then
                txt.Select(bodyStart, bodyEnd - bodyStart)
                txt.SelectionIndent = 20
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

        ApplySquareCorners(frm)
        frm.Show()

    End Sub

    Private Sub UpdateStats1(value As Double)

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

        ' PPM Deviation from the first sample since reset; the baseline needs an epsilon check
        ' (near-zero readings give huge ratios).
        'Stats1DeviationCurrent = (value - Stats1FirstValue) * 1000000
        Stats1DeviationCurrent = If(Math.Abs(Stats1FirstValue) > 0.000000001, (value - Stats1FirstValue) / Stats1FirstValue * 1000000, 0)

        ' Need at least 2 readings for STDEV
        If Stats1Count >= 2 Then
            Dim variance As Double = Stats1M2 / (Stats1Count - 1)
            Stats1StdevCurrent = Math.Sqrt(variance)
            Stats1SEMCurrent = Stats1StdevCurrent / Math.Sqrt(Stats1Count)
        Else
            Stats1StdevCurrent = 0.0
            Stats1SEMCurrent = 0.0
        End If

        ' Everything below is display only - the running stats above keep
        ' updating regardless, so "Pause Display" never affects the
        ' underlying data, only what's shown on screen.
        If PauseLiveDisplay1 Then Exit Sub

        ' Match Dev1Meter's own decimal-places setting, so this "repeat of
        ' the large digits" always shows exactly what the meter shows.
        Dim dev1ValueDp As Integer
        Dim dev1ValueFormat As String = If(Integer.TryParse(Dev1DecimalNumDPs.Text, dev1ValueDp), BuildMinDpFormat(dev1ValueDp), "0.000#######")
        LabelStats1Value.Text = value.ToString(dev1ValueFormat, Globalization.CultureInfo.InvariantCulture)

        If Stats1Count = 1 Then
            LabelStats1FirstValue.Text = Stats1FirstValue.ToString("0.000#######")
        End If

        ' Number of samples
        LabelStats1Samples.Text = Stats1Count.ToString()

        ' Mean
        LabelStats1Mean.Text = Stats1Mean.ToString("0.000#######")

        ' Max / Min recorded
        LabelStats1Max.Text = Stats1Max.ToString("0.000#######")
        LabelStats1Min.Text = Stats1Min.ToString("0.000#######")

        ' Max Diff and PPM Deviation from first sample
        LabelStats1MaxDiff.Text = (Stats1Max - Stats1Min).ToString("0.000#######")
        LabelStats1Deviation.Text = Stats1DeviationCurrent.ToString("0.000#")

        LabelStats1Stdev.Text = Stats1StdevCurrent.ToString("0.000#######")
        LabelStats1SEM.Text = Stats1SEMCurrent.ToString("0.000#######")

        ' Theoretical averaging gain in digits
        If Stats1Count > 0 Then
            Dim digitsGained As Double = 0.5 * Math.Log10(Stats1Count)
            LabelStats1Digits.Text = digitsGained.ToString("0.00")
        Else
            LabelStats1Digits.Text = "0.00"
        End If

    End Sub


    Private Sub UpdateStats2(value As Double)

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

        ' PPM Deviation from first sample since last reset - see the same
        ' near-zero-baseline note in UpdateStats1.
        'Stats2DeviationCurrent = (value - Stats2FirstValue) * 1000000
        Stats2DeviationCurrent = If(Math.Abs(Stats2FirstValue) > 0.000000001, (value - Stats2FirstValue) / Stats2FirstValue * 1000000, 0)

        ' Need at least 2 readings for STDEV
        If Stats2Count >= 2 Then
            Dim variance As Double = Stats2M2 / (Stats2Count - 1)
            Stats2StdevCurrent = Math.Sqrt(variance)
            Stats2SEMCurrent = Stats2StdevCurrent / Math.Sqrt(Stats2Count)
        Else
            Stats2StdevCurrent = 0.0
            Stats2SEMCurrent = 0.0
        End If

        ' Everything below is display only - the running stats above keep
        ' updating regardless, so "Pause Display" never affects the
        ' underlying data, only what's shown on screen.
        If PauseLiveDisplay2 Then Exit Sub

        ' Match Dev2Meter's own decimal-places setting, so this "repeat of
        ' the large digits" always shows exactly what the meter shows.
        Dim dev2ValueDp As Integer
        Dim dev2ValueFormat As String = If(Integer.TryParse(Dev2DecimalNumDPs.Text, dev2ValueDp), BuildMinDpFormat(dev2ValueDp), "0.000#######")
        LabelStats2Value.Text = value.ToString(dev2ValueFormat, Globalization.CultureInfo.InvariantCulture)

        If Stats2Count = 1 Then
            LabelStats2FirstValue.Text = Stats2FirstValue.ToString("0.000#######")
        End If

        ' Number of samples
        LabelStats2Samples.Text = Stats2Count.ToString()

        ' Mean
        LabelStats2Mean.Text = Stats2Mean.ToString("0.000#######")

        ' Max / Min recorded
        LabelStats2Max.Text = Stats2Max.ToString("0.000#######")
        LabelStats2Min.Text = Stats2Min.ToString("0.000#######")

        ' Max Diff and PPM Deviation from first sample
        LabelStats2MaxDiff.Text = (Stats2Max - Stats2Min).ToString("0.000#######")
        LabelStats2Deviation.Text = Stats2DeviationCurrent.ToString("0.000#")

        LabelStats2Stdev.Text = Stats2StdevCurrent.ToString("0.000#######")
        LabelStats2SEM.Text = Stats2SEMCurrent.ToString("0.000#######")

        ' Theoretical averaging gain in digits
        If Stats2Count > 0 Then
            Dim digitsGained As Double = 0.5 * Math.Log10(Stats2Count)
            LabelStats2Digits.Text = digitsGained.ToString("0.00")
        Else
            LabelStats2Digits.Text = "0.00"
        End If

    End Sub


    Private Sub ProcessLiveStatistics(deviceNumber As Integer, value As Double)

        ' Statistics show whenever the device is running; only the per-device Enable Statistics checkbox gates them.

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

        Dim confirmResult As DialogResult = MessageBox.Show(
            "Resetting stats starts a new baseline value for PPM Deviation." & Environment.NewLine &
            "This affects PPM Deviation from this point on, including the DEV1_DEVIATION value written to the CSV log." & Environment.NewLine & Environment.NewLine &
            "Continue with the reset?",
            "Reset Device 1 Stats",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning)

        If confirmResult = DialogResult.No Then Exit Sub

        ResetStats1()

    End Sub

    Private Sub ButtonStats2Reset_Click(sender As Object, e As EventArgs) Handles ButtonStats2Reset.Click

        Dim confirmResult As DialogResult = MessageBox.Show(
            "Resetting stats starts a new baseline value for PPM Deviation." & Environment.NewLine &
            "This affects PPM Deviation from this point on, including the DEV2_DEVIATION value written to the CSV log." & Environment.NewLine & Environment.NewLine &
            "Continue with the reset?",
            "Reset Device 2 Stats",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning)

        If confirmResult = DialogResult.No Then Exit Sub

        ResetStats2()

    End Sub

    ' Resets Device 1's running stats and readouts (incl. LabelStats1Value) to the power-up "-" state.
    Private Sub ResetStats1()

        Stats1Count = 0
        Stats1Mean = 0.0
        Stats1M2 = 0.0
        Stats1Max = Double.MinValue
        Stats1Min = Double.MaxValue
        Stats1FirstValue = Double.NaN
        Stats1DeviationCurrent = 0.0
        LiveAnalysisLastStats1Count = Stats1Count   ' keep chart's "last plotted" in sync with the reset

        Dev1Meter.Text = "---------------"
        LabelStats1Value.Text = "-"
        LabelStats1Samples.Text = "-"
        LabelStats1Mean.Text = "-"
        LabelStats1Stdev.Text = "-"
        LabelStats1SEM.Text = "-"
        LabelStats1Digits.Text = "-"
        LabelStats1Max.Text = "-"
        LabelStats1Min.Text = "-"
        LabelStats1MaxDiff.Text = "-"
        LabelStats1Deviation.Text = "-"
        LabelStats1FirstValue.Text = "-"

        ' Resetting stats also un-pauses the display, so the reset is
        ' actually visible rather than sitting frozen behind a paused view.
        If PauseLiveDisplay1 Then
            PauseLiveDisplay1 = False
            ButtonStats1PauseDisplay.Text = "Pause Display"
            LabelStats1Samples.ForeColor = LabelStats1SamplesNormalColor
            LabelStats1Samples.BackColor = LabelStats1SamplesNormalBackColor
            If Not PauseLiveDisplay2 Then PauseFlashTimer.Stop()
        End If

    End Sub

    Private Sub ResetStats2()

        Stats2Count = 0
        Stats2Mean = 0.0
        Stats2M2 = 0.0
        Stats2Max = Double.MinValue
        Stats2Min = Double.MaxValue
        Stats2FirstValue = Double.NaN
        Stats2DeviationCurrent = 0.0
        LiveAnalysisLastStats2Count = Stats2Count   ' keep chart's "last plotted" in sync with the reset

        Dev2Meter.Text = "---------------"
        LabelStats2Value.Text = "-"
        LabelStats2Samples.Text = "-"
        LabelStats2Mean.Text = "-"
        LabelStats2Stdev.Text = "-"
        LabelStats2SEM.Text = "-"
        LabelStats2Digits.Text = "-"
        LabelStats2Max.Text = "-"
        LabelStats2Min.Text = "-"
        LabelStats2MaxDiff.Text = "-"
        LabelStats2Deviation.Text = "-"
        LabelStats2FirstValue.Text = "-"

        ' Resetting stats also un-pauses the display, so the reset is
        ' actually visible rather than sitting frozen behind a paused view.
        If PauseLiveDisplay2 Then
            PauseLiveDisplay2 = False
            ButtonStats2PauseDisplay.Text = "Pause Display"
            LabelStats2Samples.ForeColor = LabelStats2SamplesNormalColor
            LabelStats2Samples.BackColor = LabelStats2SamplesNormalBackColor
            If Not PauseLiveDisplay1 Then PauseFlashTimer.Stop()
        End If

    End Sub


    Private Sub ButtonStats1PauseDisplay_Click(sender As Object, e As EventArgs) Handles ButtonStats1PauseDisplay.Click

        PauseLiveDisplay1 = Not PauseLiveDisplay1
        ButtonStats1PauseDisplay.Text = If(PauseLiveDisplay1, "Resume Display", "Pause Display")

        If PauseLiveDisplay1 Then
            LabelStats1SamplesNormalColor = LabelStats1Samples.ForeColor
            LabelStats1SamplesNormalBackColor = LabelStats1Samples.BackColor
            PauseFlashTimer.Start()
        Else
            LabelStats1Samples.ForeColor = LabelStats1SamplesNormalColor
            LabelStats1Samples.BackColor = LabelStats1SamplesNormalBackColor
        End If

    End Sub


    Private Sub ButtonStats2PauseDisplay_Click(sender As Object, e As EventArgs) Handles ButtonStats2PauseDisplay.Click

        PauseLiveDisplay2 = Not PauseLiveDisplay2
        ButtonStats2PauseDisplay.Text = If(PauseLiveDisplay2, "Resume Display", "Pause Display")

        If PauseLiveDisplay2 Then
            LabelStats2SamplesNormalColor = LabelStats2Samples.ForeColor
            LabelStats2SamplesNormalBackColor = LabelStats2Samples.BackColor
            PauseFlashTimer.Start()
        Else
            LabelStats2Samples.ForeColor = LabelStats2SamplesNormalColor
            LabelStats2Samples.BackColor = LabelStats2SamplesNormalBackColor
        End If

    End Sub


    Private Sub PauseFlashTimer_Tick(sender As Object, e As EventArgs) Handles PauseFlashTimer.Tick

        PauseFlashOn = Not PauseFlashOn

        If PauseLiveDisplay1 Then
            If PauseFlashOn Then
                LabelStats1Samples.ForeColor = LabelStats1SamplesNormalBackColor
                LabelStats1Samples.BackColor = LabelStats1SamplesNormalColor
            Else
                LabelStats1Samples.ForeColor = LabelStats1SamplesNormalColor
                LabelStats1Samples.BackColor = LabelStats1SamplesNormalBackColor
            End If
        End If

        If PauseLiveDisplay2 Then
            If PauseFlashOn Then
                LabelStats2Samples.ForeColor = LabelStats2SamplesNormalBackColor
                LabelStats2Samples.BackColor = LabelStats2SamplesNormalColor
            Else
                LabelStats2Samples.ForeColor = LabelStats2SamplesNormalColor
                LabelStats2Samples.BackColor = LabelStats2SamplesNormalBackColor
            End If
        End If

        ' Nothing left to flash - stop running rather than tick forever in the background.
        If Not PauseLiveDisplay1 AndAlso Not PauseLiveDisplay2 Then
            PauseFlashTimer.Stop()
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

        FormsPlot1.Visible = False
        StartChartMessage.Visible = True

        ' Clear charts
        Chart1Dev1Data.Clear()
        Chart1Dev2Data.Clear()
        Chart1TempData.Clear()
        Chart1Dev1NextX = 0
        Chart1Dev2NextX = 0
        Chart1TempNextX = 0
        Chart1ClearMeasurement(FormsPlot1.Plot)
        FormsPlot1.Refresh()

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

            If Double.TryParse(Dev1Min.Text, minValue) AndAlso
           Double.TryParse(Dev1Max.Text, maxValue) AndAlso
           Chart1AutoScaleYAxis.Checked = False Then

                ' Determine the number of decimal places based on maximum precision
                Dim decimalPlaces As Integer =
                Math.Max(GetMaxPrecision(minValue, maxValue), 1)

                UpdateChartYAxisMinMaxInterval()

            End If

            FormsPlot1.Visible = True
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
        Dim chartPoints1 As Integer = Chart1Dev1Data.Count
        Dim chartPoints2 As Integer = Chart1Dev2Data.Count

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

            ' Capped by the X-axis rolling window: this label shows the visible time span, not total run time.
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


            ' A running device with a hidden trace is excluded from the autoscaled Y range.
            Dim dev1Visible As Boolean = EnableChart1.Checked AndAlso Not CheckBoxDevice1Hide.Checked
            Dim dev2Visible As Boolean = EnableChart2.Checked AndAlso Not CheckBoxDevice2Hide.Checked

            ' Autoscale chart y-axis - Device 1 only
            If (Chart1AutoScaleYAxis.Checked = True And dev1Visible = True And dev2Visible = False) Then
                ' Autoscale as soon as any data has arrived - the range=0
                ' buffer below covers a single repeated value (e.g. min =
                ' max = 1.000000).
                If Chart1Dev1Data.Count >= 1 Then
                    ' Autoscale the minimum and maximum of the Y-axis
                    Dim minValue1 As Double = Chart1Dev1Data.Min(Function(p) p.Y)
                    Dim maxValue1 As Double = Chart1Dev1Data.Max(Function(p) p.Y)

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

                    ' Prevent scientific notation (e-notation) on the Y-axis labels
                    Chart1SetYAxisRange(minValue1, maxValue1)
                    YaxisDiff.Text = Format(range, "#0.00000000")
                Else
                    UpdateChartYAxisMinMaxInterval()
                    YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
                End If
            End If

            ' Not applied every tick as it would fight manual pan/zoom; the textbox handlers apply edits.
            If (Chart1AutoScaleYAxis.Checked = False And EnableChart1.Checked = True And EnableChart2.Checked = False) Then
                YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
            End If


            ' Autoscale chart y-axis - Device 2 only
            If (Chart1AutoScaleYAxis.Checked = True And dev1Visible = False And dev2Visible = True) Then
                ' Autoscale as soon as any data has arrived.
                If Chart1Dev2Data.Count >= 1 Then
                    ' Autoscale the minimum and maximum of the Y-axis
                    Dim minValue2 As Double = Chart1Dev2Data.Min(Function(p) p.Y)
                    Dim maxValue2 As Double = Chart1Dev2Data.Max(Function(p) p.Y)

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

                    ' Prevent scientific notation (e-notation) on the Y-axis labels
                    Chart1SetYAxisRange(minValue2, maxValue2)
                    YaxisDiff.Text = Format(range, "#0.00000000")
                Else
                    UpdateChartYAxisMinMaxInterval()
                    YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
                End If
            End If

            If (Chart1AutoScaleYAxis.Checked = False And EnableChart1.Checked = False And EnableChart2.Checked = True) Then
                YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
            End If




            ' Autoscale chart y-axis - Device 1 & Device 2
            If (Chart1AutoScaleYAxis.Checked = True And dev1Visible = True And dev2Visible = True) Then
                ' Autoscale as soon as both devices have any data.
                If (Chart1Dev1Data.Count >= 1 And Chart1Dev2Data.Count >= 1) Then

                    ' Get the minimum and maximum values from both series
                    Dim minValue1 As Double = Chart1Dev1Data.Min(Function(p) p.Y)
                    Dim minValue2 As Double = Chart1Dev2Data.Min(Function(p) p.Y)

                    Dim maxValue1 As Double = Chart1Dev1Data.Max(Function(p) p.Y)
                    Dim maxValue2 As Double = Chart1Dev2Data.Max(Function(p) p.Y)

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

                    ' Set the minimum and maximum values for the Y-axis,
                    ' preventing scientific notation on its labels
                    Chart1SetYAxisRange(overallMin, overallMax)

                    YaxisDiff.Text = Format(range, "#0.00000000")
                Else
                    UpdateChartYAxisMinMaxInterval()
                    YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
                End If
            End If

            If (Chart1AutoScaleYAxis.Checked = False And EnableChart1.Checked = True And EnableChart2.Checked = True) Then
                YaxisDiff.Text = Format(Val(Dev1Max.Text) - Val(Dev1Min.Text), "#0.00000000")
            End If

        Else
            Dev1Min.ReadOnly = False
            Dev1Max.ReadOnly = False
            'ButtonClearChart.Enabled = True

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
            Chart1Dev1Series.IsVisible = False
        Else
            Chart1Dev1Series.IsVisible = True
        End If

        If CheckBoxDevice2Hide.Checked = True Then
            Chart1Dev2Series.IsVisible = False
        Else
            Chart1Dev2Series.IsVisible = True
        End If

        If CheckBoxTempHide.Checked = True Then
            Chart1TempSeries.IsVisible = False
        Else
            Chart1TempSeries.IsVisible = True
        End If

        FormsPlot1.Refresh()

    End Sub


    Private Sub Chart1AutoScaleYAxis_CheckedChanged(sender As Object, e As EventArgs) Handles Chart1AutoScaleYAxis.CheckedChanged

        ' ReadOnly (not Enabled = False) while autoscaling, so the boxes stay
        ' legible as autoscale continuously writes the detected min/max into
        ' them, instead of greying out.
        If Chart1AutoScaleYAxis.Checked = True Then
            Dev1Max.ReadOnly = True
            Dev1Min.ReadOnly = True
        Else
            Dev1Max.ReadOnly = False
            Dev1Min.ReadOnly = False
        End If

        ' XaxisPoints drives the rolling window width while following live data;
        ' after manual pan/zoom it only reflects the view width, so it's read-only.
        XaxisPoints.ReadOnly = Not Chart1AutoScaleYAxis.Checked

    End Sub


    ' Rejects non-numeric input or Max <= Min and reverts the box to the axis's last good value.
    Private Sub Dev1Max_Leave(sender As Object, e As EventArgs) Handles Dev1Max.Leave

        Dim maxVal As Double
        Dim minVal As Double

        If Not Double.TryParse(Dev1Max.Text, maxVal) OrElse
           Not Double.TryParse(Dev1Min.Text, minVal) OrElse
           maxVal <= minVal Then
            Dev1Max.Text = FormsPlot1.Plot.Axes.Left.Max.ToString(Globalization.CultureInfo.InvariantCulture)
            Exit Sub
        End If

        UpdateChartYAxisMinMaxInterval()

    End Sub


    Private Sub Dev1Min_Leave(sender As Object, e As EventArgs) Handles Dev1Min.Leave

        Dim maxVal As Double
        Dim minVal As Double

        If Not Double.TryParse(Dev1Max.Text, maxVal) OrElse
           Not Double.TryParse(Dev1Min.Text, minVal) OrElse
           minVal >= maxVal Then
            Dev1Min.Text = FormsPlot1.Plot.Axes.Left.Min.ToString(Globalization.CultureInfo.InvariantCulture)
            Exit Sub
        End If

        UpdateChartYAxisMinMaxInterval()

    End Sub


    ' Applies the typed value as soon as the user presses Enter, instead
    ' of only on Leave (tabbing/clicking away).
    Private Sub Dev1Max_KeyDown(sender As Object, e As KeyEventArgs) Handles Dev1Max.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dev1Max_Leave(sender, e)
        End If
    End Sub

    Private Sub Dev1Min_KeyDown(sender As Object, e As KeyEventArgs) Handles Dev1Min.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dev1Min_Leave(sender, e)
        End If
    End Sub


    ' Enforces a numeric value of at least 100 points so later comparisons on .Text are safe.
    Private Sub XaxisPoints_Leave(sender As Object, e As EventArgs) Handles XaxisPoints.Leave

        ' While read-only (AutoScale Y-axis unchecked) this box just
        ' reflects the mouse-set view width, which can legitimately be
        ' under 100 - only clamp actual typed user input.
        If XaxisPoints.ReadOnly Then Exit Sub

        Dim points As Integer

        If Not Integer.TryParse(XaxisPoints.Text, points) OrElse points < 100 Then
            XaxisPoints.Text = "100"
        End If

    End Sub

    Private Sub XaxisPoints_KeyDown(sender As Object, e As KeyEventArgs) Handles XaxisPoints.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            XaxisPoints_Leave(sender, e)
        End If
    End Sub


    ' Same reject-and-revert as Dev1Max/Min, using the temperature axis's last good value.
    Private Sub LCTempMax_Leave(sender As Object, e As EventArgs) Handles LCTempMax.Leave

        Dim maxVal As Double
        Dim minVal As Double

        If Not Double.TryParse(LCTempMax.Text, maxVal) OrElse
           Not Double.TryParse(LCTempMin.Text, minVal) OrElse
           maxVal <= minVal Then
            LCTempMax.Text = Chart1TempAxis.Max.ToString(Globalization.CultureInfo.InvariantCulture)
            Exit Sub
        End If

        UpdateChartTemperatureYAxisMinMaxInterval()

    End Sub

    Private Sub LCTempMin_Leave(sender As Object, e As EventArgs) Handles LCTempMin.Leave

        Dim maxVal As Double
        Dim minVal As Double

        If Not Double.TryParse(LCTempMax.Text, maxVal) OrElse
           Not Double.TryParse(LCTempMin.Text, minVal) OrElse
           minVal >= maxVal Then
            LCTempMin.Text = Chart1TempAxis.Min.ToString(Globalization.CultureInfo.InvariantCulture)
            Exit Sub
        End If

        UpdateChartTemperatureYAxisMinMaxInterval()

    End Sub

    Private Sub LCTempMax_KeyDown(sender As Object, e As KeyEventArgs) Handles LCTempMax.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            LCTempMax_Leave(sender, e)
        End If
    End Sub

    Private Sub LCTempMin_KeyDown(sender As Object, e As KeyEventArgs) Handles LCTempMin.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            LCTempMin_Leave(sender, e)
        End If
    End Sub


    Private Sub UpdateChartYAxisMinMaxInterval()

        ' Runs every 100ms; skips unparsable or inverted ranges (assigning them throws) and keeps the last good range.
        Dim minVal As Double
        Dim maxVal As Double

        If Not Double.TryParse(Dev1Min.Text, minVal) Then Exit Sub
        If Not Double.TryParse(Dev1Max.Text, maxVal) Then Exit Sub
        If maxVal <= minVal Then Exit Sub

        ' Set the minimum and maximum values for the Y-axis
        FormsPlot1.Plot.Axes.Left.Min = minVal
        FormsPlot1.Plot.Axes.Left.Max = maxVal

        FormsPlot1.Refresh()

    End Sub


    Private Sub UpdateChartTemperatureYAxisMinMaxInterval()

        ' Parse the minimum and maximum values from the text inputs
        Dim TminVal As Double = Val(LCTempMin.Text)
        Dim TmaxVal As Double = Val(LCTempMax.Text)

        ' Guards against an inverted range so the temperature axis can never be assigned one.
        If TmaxVal <= TminVal Then Exit Sub

        ' Set the minimum and maximum values for the Y-axis. Tick/label
        ' formatting is applied uniformly every render by
        ' Chart1RenderStarting, so it doesn't need to be set here too.
        Chart1TempAxis.Min = TminVal
        Chart1TempAxis.Max = TmaxVal

        FormsPlot1.Refresh()

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
        .Width = 1000,
        .Height = 800,
        .MinimumSize = New Size(1000, 800),
        .ShowIcon = False,
        .ShowInTaskbar = True,
        .BackColor = Color.WhiteSmoke
    }


        LiveAnalysisChart = New DataVisualization.Charting.Chart With {
        .Dock = DockStyle.Fill,
        .BackColor = Color.WhiteSmoke,
        .AntiAliasing = DataVisualization.Charting.AntiAliasingStyles.All
    }

        ' Chart Area 1 - Device readings and running means
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


        ' Chart Area 2 - STDEV / SEM
        Dim areaStatistics As New DataVisualization.Charting.ChartArea("Statistics")

        areaStatistics.Position =
        New DataVisualization.Charting.ElementPosition(4, 57, 88, 23)
        areaStatistics.InnerPlotPosition = New DataVisualization.Charting.ElementPosition(12, 5, 85, 90)

        areaStatistics.BackColor = Color.Black

        areaStatistics.AxisY.IsStartedFromZero = False
        'areaStatistics.AxisY.LabelStyle.Format = "0.###E+00"
        areaStatistics.AxisY.LabelStyle.Format = "0.00000000"

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

        ' PPM Deviation gets its own right-hand axis: it's in ppm, far larger than the raw STDEV/SEM values.
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


        ' Chart Area 3 - Temperature
        Dim areaTemperature As New DataVisualization.Charting.ChartArea("Temperature")

        areaTemperature.Position = New DataVisualization.Charting.ElementPosition(4, 82, 88, 15)
        areaTemperature.InnerPlotPosition = New DataVisualization.Charting.ElementPosition(12, 5, 85, 90)

        areaTemperature.BackColor = Color.Black

        areaTemperature.AxisY.IsStartedFromZero = False
        areaTemperature.AxisY.LabelStyle.Format = "0.000"

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


        ' Chart Area Titles
        Dim titleMeasurement As New DataVisualization.Charting.Title
        titleMeasurement.Text = "DEVICE 1 & 2 DATA / RUNNING MEAN"
        titleMeasurement.DockedToChartArea = "Measurement"
        titleMeasurement.Docking = DataVisualization.Charting.Docking.Left
        titleMeasurement.IsDockedInsideChartArea = False
        titleMeasurement.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        titleMeasurement.TextOrientation = DataVisualization.Charting.TextOrientation.Rotated270
        titleMeasurement.Position.Auto = False
        titleMeasurement.Position = New DataVisualization.Charting.ElementPosition(2.0F, 9.0F, 4.0F, 45.0F)
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


        ' Series

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


        ' Trace enable/disable checkboxes - Becomes Legends also
        Dim liveToggles As New List(Of CheckBox)

        ' Only show a device name once it's actually connected, not its configured name.
        Dim dev1ActiveAtOpen As Boolean = (ButtonDev1Run.Text = "Stop") OrElse (ButtonDev12Run.Text = "Stop")
        Dim dev2ActiveAtOpen As Boolean = (ButtonDev2Run.Text = "Stop") OrElse (ButtonDev12Run.Text = "Stop")

        Dim gbDev1 As New GroupBox With {.Text = If(dev1ActiveAtOpen, "Device 1 - " & txtname1.Text, "Device 1"), .BackColor = Color.WhiteSmoke, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim gbDev2 As New GroupBox With {.Text = If(dev2ActiveAtOpen, "Device 2 - " & txtname2.Text, "Device 2"), .BackColor = Color.WhiteSmoke, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim gbTemp As New GroupBox With {.Text = "Temperature", .BackColor = Color.WhiteSmoke, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim gbMisc As New GroupBox With {.Text = "Misc.", .BackColor = Color.WhiteSmoke, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}

        LiveAnalysisChart.Controls.Add(gbDev1)
        LiveAnalysisChart.Controls.Add(gbDev2)
        LiveAnalysisChart.Controls.Add(gbTemp)
        LiveAnalysisChart.Controls.Add(gbMisc)

        ' Lets the user manually clear all three traces and restart the
        ' sample counter without needing to Stop/Start a device - placed
        ' directly below the Temperature groupbox in RepositionLiveToggles.
        Dim btnResetLiveCharts As New Button With {
            .Text = "Restart Charts",
            .Height = 20,
            .Font = New Font("Segoe UI", 8, FontStyle.Regular)
        }
        LiveAnalysisChart.Controls.Add(btnResetLiveCharts)

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
        Dim tglDev1PPM = AddTraceToggle(gbDev1, "Dev 1 PPM Deviation", "PPM DEV", Color.White)

        Dim tglDevice2 = AddTraceToggle(gbDev2, "Device 2", "Data", Color.Aqua)
        Dim tglDev2Mean = AddTraceToggle(gbDev2, "Dev 2 Mean", "Mean", Color.Lime)
        Dim tglDev2Stdev = AddTraceToggle(gbDev2, "Dev 2 STDEV", "STDEV", Color.Magenta)
        Dim tglDev2SEM = AddTraceToggle(gbDev2, "Dev 2 SEM", "SEM", Color.LimeGreen)
        Dim tglDev2PPM = AddTraceToggle(gbDev2, "Dev 2 PPM Deviation", "PPM DEV", Color.LightGray)

        Dim tglTemp = AddTraceToggle(gbTemp, "Temperature", "Temp.", Color.Red)

        Dim dev1Boxes = {tglDevice1, tglDev1Mean, tglDev1Stdev, tglDev1SEM, tglDev1PPM}
        Dim dev2Boxes = {tglDevice2, tglDev2Mean, tglDev2Stdev, tglDev2SEM, tglDev2PPM}
        Dim tempBoxes = {tglTemp}

        ' ToolTip1 belongs to Formtest and reliably tracks hover only for
        ' Formtest's own controls - LiveAnalysisForm is a separate top-level
        ' Form, so its controls need their own ToolTip to actually show.
        Dim liveAnalysisToolTip As New ToolTip()

        ' Global rendering toggle - applies to every LiveWatch chart (this
        ' pop-out AND the embedded Live Chart on the Devices tab), not just
        ' one series, so it isn't wired up via AddTraceToggle like the others.
        Dim chkAntiAliasing As New CheckBox With {
            .Text = "Anti-Aliasing",
            .ForeColor = Color.Black,
            .AutoSize = False,
            .Checked = True,
            .Height = 18,
            .Font = New Font("Segoe UI", 8, FontStyle.Regular)
        }
        gbMisc.Controls.Add(chkAntiAliasing)
        liveAnalysisToolTip.SetToolTip(chkAntiAliasing, "Smooths lines and text on this chart and the main Live Chart." & vbCrLf & "Keep this ON to avoid a jagged/moire look on busy traces.")

        AddHandler chkAntiAliasing.CheckedChanged, Sub(s, ev)
                                                       Dim style = If(chkAntiAliasing.Checked,
                                                           DataVisualization.Charting.AntiAliasingStyles.All,
                                                           DataVisualization.Charting.AntiAliasingStyles.None)
                                                       LiveAnalysisChart.AntiAliasing = style
                                                       Chart1Dev1Series.LineStyle.AntiAlias = chkAntiAliasing.Checked
                                                       Chart1Dev2Series.LineStyle.AntiAlias = chkAntiAliasing.Checked
                                                       Chart1TempSeries.LineStyle.AntiAlias = chkAntiAliasing.Checked
                                                       FormsPlot1.Refresh()
                                                   End Sub

        ' Fast Rendering / Smooth Lines only change this pop-out's series chart types, not Chart1.
        Dim chkFastRendering As New CheckBox With {
            .Text = "Fast Rendering",
            .ForeColor = Color.Black,
            .AutoSize = False,
            .Checked = False,
            .Height = 18,
            .Font = New Font("Segoe UI", 8, FontStyle.Regular)
        }
        gbMisc.Controls.Add(chkFastRendering)
        liveAnalysisToolTip.SetToolTip(chkFastRendering, "Switches to a faster, lower-quality line renderer (FastLine) that" & vbCrLf & "skips anti-aliasing. Only worth using if the chart becomes slow" & vbCrLf & "with very large point counts. Overrides Smooth Lines while checked.")

        Dim chkSmoothLines As New CheckBox With {
            .Text = "Smooth Lines",
            .ForeColor = Color.Black,
            .AutoSize = False,
            .Checked = False,
            .Height = 18,
            .Font = New Font("Segoe UI", 8, FontStyle.Regular)
        }
        gbMisc.Controls.Add(chkSmoothLines)
        liveAnalysisToolTip.SetToolTip(chkSmoothLines, "Curves the trace between points (spline) instead of straight" & vbCrLf & "segments. Cosmetic only - can visually suggest values that were" & vbCrLf & "never actually measured between samples.")

        ' Display only: rolling average of the last ShortTermMeanWindow readings on the Mean trace.
        ' Stats1Mean/Stats2Mean, STDEV/SEM/PPM Deviation and the CSV are unaffected.
        chkShortTermMean = New CheckBox With {
            .Text = "Short-Term Mean",
            .ForeColor = Color.Black,
            .AutoSize = False,
            .Checked = False,
            .Height = 18,
            .Font = New Font("Segoe UI", 8, FontStyle.Regular)
        }
        gbMisc.Controls.Add(chkShortTermMean)
        liveAnalysisToolTip.SetToolTip(chkShortTermMean, $"Shows the Mean TRACE as a rolling average of the last {ShortTermMeanWindow} readings" & vbCrLf & "instead of the full cumulative average since Reset Stats - more responsive" & vbCrLf & "to recent changes, but noisier. Display only: does not affect the Mean" & vbCrLf & "shown elsewhere, STDEV/SEM/PPM Deviation, or the CSV log." & vbCrLf & "Only affects the trace from the moment you check/uncheck it onwards -" & vbCrLf & "points already plotted are not redrawn.")

        Dim miscBoxes = {chkAntiAliasing, chkFastRendering, chkSmoothLines, chkShortTermMean}

        ' Fast Rendering (FastLine) wins over Smooth Lines (Spline) since a
        ' series can't be both at once - Smooth Lines is disabled while Fast
        ' Rendering is checked so it's clear which one is actually in effect.
        Dim ApplyLiveAnalysisChartType = Sub()
                                             Dim chartType As DataVisualization.Charting.SeriesChartType

                                             If chkFastRendering.Checked Then
                                                 chartType = DataVisualization.Charting.SeriesChartType.FastLine
                                             ElseIf chkSmoothLines.Checked Then
                                                 chartType = DataVisualization.Charting.SeriesChartType.Spline
                                             Else
                                                 chartType = DataVisualization.Charting.SeriesChartType.Line
                                             End If

                                             For Each s As DataVisualization.Charting.Series In LiveAnalysisChart.Series
                                                 s.ChartType = chartType
                                             Next

                                             chkSmoothLines.Enabled = Not chkFastRendering.Checked
                                         End Sub

        AddHandler chkFastRendering.CheckedChanged, Sub(s, ev) ApplyLiveAnalysisChartType()
        AddHandler chkSmoothLines.CheckedChanged, Sub(s, ev) ApplyLiveAnalysisChartType()

        ' Tracks each device's run state across calls so a fresh Run can be
        ' told apart from a continuing one - see RefreshDeviceAvailability.
        Dim dev1WasActive As Boolean = False
        Dim dev2WasActive As Boolean = False
        Dim tempWasActive As Boolean = False

        Dim ClearLiveAnalysisSeries = Sub(seriesNames As String())
                                          For Each seriesName In seriesNames
                                              If LiveAnalysisChart.Series.IndexOf(seriesName) >= 0 Then
                                                  LiveAnalysisChart.Series(seriesName).Points.Clear()
                                              End If
                                          Next
                                      End Sub

        Dim EnableGroup = Sub(boxes As CheckBox())
                              For Each cb As CheckBox In boxes
                                  ' Also uncheck, so a re-enabled box defaults to checked like a fresh AddTraceToggle box.
                                  cb.Checked = True
                                  cb.Enabled = True
                                  cb.BackColor = CType(cb.Tag, Color)
                              Next
                          End Sub

        Dim DisableGroup = Sub(boxes As CheckBox())
                               For Each cb As CheckBox In boxes
                                   ' Unchecking fires AddTraceToggle's CheckedChanged, which disables the matching series.
                                   cb.Checked = False
                                   cb.Enabled = False
                               Next
                           End Sub

        Dim RefreshDeviceAvailability = Sub()
                                            Dim dev1Active As Boolean = (ButtonDev1Run.Text = "Stop") OrElse (ButtonDev12Run.Text = "Stop")
                                            Dim dev2Active As Boolean = (ButtonDev2Run.Text = "Stop") OrElse (ButtonDev12Run.Text = "Stop")

                                            ' Temperature has its own sensor: ButtonStart/ButtonEnd (TempHumidity.vb) start/stop it,
                                            ' ButtonEnd.Enabled = True means running.
                                            Dim tempActive As Boolean = ButtonEnd.Enabled

                                            ' Stopped -> running clears that device's traces and re-enables its checkboxes.
                                            ' Running -> stopped leaves them frozen so the last run stays on screen.
                                            If dev1Active AndAlso Not dev1WasActive Then
                                                ClearLiveAnalysisSeries({"Device 1", "Dev 1 Mean", "Dev 1 STDEV", "Dev 1 SEM", "Dev 1 PPM Deviation"})
                                                q1ShortTermMean.Clear() : sum1ShortTermMean = 0.0
                                                LiveAnalysisLastStats1Count = Stats1Count
                                                EnableGroup(dev1Boxes)
                                            End If

                                            If dev2Active AndAlso Not dev2WasActive Then
                                                ClearLiveAnalysisSeries({"Device 2", "Dev 2 Mean", "Dev 2 STDEV", "Dev 2 SEM", "Dev 2 PPM Deviation"})
                                                q2ShortTermMean.Clear() : sum2ShortTermMean = 0.0
                                                LiveAnalysisLastStats2Count = Stats2Count
                                                EnableGroup(dev2Boxes)
                                            End If

                                            If tempActive AndAlso Not tempWasActive Then
                                                ClearLiveAnalysisSeries({"Temperature"})
                                                EnableGroup(tempBoxes)
                                            End If

                                            ' Unlike Dev1/Dev2, Temperature is disabled and unchecked when stopped (no sensor behind it).
                                            If Not tempActive AndAlso tempWasActive Then
                                                DisableGroup(tempBoxes)
                                            End If

                                            dev1WasActive = dev1Active
                                            dev2WasActive = dev2Active
                                            tempWasActive = tempActive
                                        End Sub

        RefreshDeviceAvailability()

        ' RefreshDeviceAvailability() never disables a group (so stopped devices stay frozen);
        ' disable never-connected devices here, when the popup first opens.
        If Not dev1ActiveAtOpen Then DisableGroup(dev1Boxes)
        If Not dev2ActiveAtOpen Then DisableGroup(dev2Boxes)
        If Not ButtonEnd.Enabled Then DisableGroup(tempBoxes)

        Dim RunButtonHandler = Sub(s As Object, ev As EventArgs) RefreshDeviceAvailability()

        AddHandler ButtonDev1Run.Click, RunButtonHandler
        AddHandler ButtonDev2Run.Click, RunButtonHandler
        AddHandler ButtonDev12Run.Click, RunButtonHandler
        AddHandler ButtonStart.Click, RunButtonHandler
        AddHandler ButtonEnd.Click, RunButtonHandler

        ' Main Reset: clears the Dev 1/Dev 2/Temperature traces and disables their checkboxes,
        ' marking all three inactive so the next Run/Start re-clears and re-enables them.
        Dim ResetHandler = Sub(s As Object, ev As EventArgs)
                               ClearLiveAnalysisSeries({"Device 1", "Dev 1 Mean", "Dev 1 STDEV", "Dev 1 SEM", "Dev 1 PPM Deviation"})
                               ClearLiveAnalysisSeries({"Device 2", "Dev 2 Mean", "Dev 2 STDEV", "Dev 2 SEM", "Dev 2 PPM Deviation"})
                               ClearLiveAnalysisSeries({"Temperature"})
                               q1ShortTermMean.Clear() : sum1ShortTermMean = 0.0
                               q2ShortTermMean.Clear() : sum2ShortTermMean = 0.0
                               LiveAnalysisLastStats1Count = Stats1Count
                               LiveAnalysisLastStats2Count = Stats2Count
                               DisableGroup(dev1Boxes)
                               DisableGroup(dev2Boxes)
                               DisableGroup(tempBoxes)
                               gbDev1.Text = "Device 1"
                               gbDev2.Text = "Device 2"
                               dev1WasActive = False
                               dev2WasActive = False
                               tempWasActive = False
                           End Sub

        AddHandler ButtonReset.Click, ResetHandler

        ' Manual "Restart Charts" button - clears the plotted traces and
        ' restarts the sample counter, leaving the devices themselves
        ' (and the checkboxes' enabled state) untouched.
        Dim ResetChartsButtonHandler = Sub(s As Object, ev As EventArgs)
                                           ClearLiveAnalysisSeries({"Device 1", "Dev 1 Mean", "Dev 1 STDEV", "Dev 1 SEM", "Dev 1 PPM Deviation"})
                                           ClearLiveAnalysisSeries({"Device 2", "Dev 2 Mean", "Dev 2 STDEV", "Dev 2 SEM", "Dev 2 PPM Deviation"})
                                           ClearLiveAnalysisSeries({"Temperature"})
                                           q1ShortTermMean.Clear() : sum1ShortTermMean = 0.0
                                           q2ShortTermMean.Clear() : sum2ShortTermMean = 0.0
                                           LiveAnalysisSample = 0
                                           LiveAnalysisLastStats1Count = Stats1Count
                                           LiveAnalysisLastStats2Count = Stats2Count
                                       End Sub

        AddHandler btnResetLiveCharts.Click, ResetChartsButtonHandler

        ' Re-enables a reconnected device's checkboxes and restores its groupbox title.
        ' btncreate = both devices, btncreate2 = Device 1, btncreate3 = Device 2.
        Dim ConnectBothHandler = Sub(s As Object, ev As EventArgs)
                                     EnableGroup(dev1Boxes)
                                     EnableGroup(dev2Boxes)
                                     gbDev1.Text = "Device 1 - " & txtname1.Text
                                     gbDev2.Text = "Device 2 - " & txtname2.Text
                                 End Sub
        Dim ConnectDev1Handler = Sub(s As Object, ev As EventArgs)
                                     EnableGroup(dev1Boxes)
                                     gbDev1.Text = "Device 1 - " & txtname1.Text
                                 End Sub
        Dim ConnectDev2Handler = Sub(s As Object, ev As EventArgs)
                                     EnableGroup(dev2Boxes)
                                     gbDev2.Text = "Device 2 - " & txtname2.Text
                                 End Sub

        AddHandler btncreate.Click, ConnectBothHandler
        AddHandler btncreate2.Click, ConnectDev1Handler
        AddHandler btncreate3.Click, ConnectDev2Handler

        ' The stacked ChartAreas use percentage positions, so margins grow with the popup;
        ' hold the original pixel margins constant and let the plot areas absorb the size change.
        Dim liveChartAreasInOrder As DataVisualization.Charting.ChartArea() =
            {areaMeasurement, areaStatistics, areaTemperature}

        Dim originalLiveChartWidth As Double
        Dim originalLiveChartHeight As Double

        Dim originalAreaLeftMarginPx As New Dictionary(Of DataVisualization.Charting.ChartArea, Double)
        Dim originalAreaRightMarginPx As New Dictionary(Of DataVisualization.Charting.ChartArea, Double)
        Dim originalAreaInnerLeftMarginPx As New Dictionary(Of DataVisualization.Charting.ChartArea, Double)
        Dim originalAreaInnerRightMarginPx As New Dictionary(Of DataVisualization.Charting.ChartArea, Double)
        Dim originalAreaHeightPx As New Dictionary(Of DataVisualization.Charting.ChartArea, Double)
        Dim originalAreaInnerTopMarginPx As New Dictionary(Of DataVisualization.Charting.ChartArea, Double)
        Dim originalAreaInnerBottomMarginPx As New Dictionary(Of DataVisualization.Charting.ChartArea, Double)

        Dim originalAreaTopMarginPx As Double         ' above areaMeasurement
        Dim originalGapMeasStatsPx As Double          ' between Measurement and Statistics
        Dim originalGapStatsTempPx As Double          ' between Statistics and Temperature
        Dim originalAreaBottomMarginPx As Double      ' below areaTemperature

        ' Rotated axis titles are separate Chart.Titles positioned as a percentage of the whole chart;
        ' hold their distance from the left/right edges fixed.
        Dim liveChartTitlesInOrder As DataVisualization.Charting.Title() =
            {titleMeasurement, titleStatistics, titleStatisticsPPM, titleTemperature}
        Dim titleOwnerArea As New Dictionary(Of DataVisualization.Charting.Title, DataVisualization.Charting.ChartArea) From {
            {titleMeasurement, areaMeasurement},
            {titleStatistics, areaStatistics},
            {titleStatisticsPPM, areaStatistics},
            {titleTemperature, areaTemperature}
        }

        Dim originalTitleLeftPx As New Dictionary(Of DataVisualization.Charting.Title, Double)
        Dim originalTitleRightPx As New Dictionary(Of DataVisualization.Charting.Title, Double)
        Dim originalTitleWidthPx As New Dictionary(Of DataVisualization.Charting.Title, Double)
        Dim originalTitleYFractionOfArea As New Dictionary(Of DataVisualization.Charting.Title, Double)
        ' Height is left as its original percentage, which scales correctly with the chart.
        Dim originalTitleHeightPct As New Dictionary(Of DataVisualization.Charting.Title, Single)

        ' Extra inner margin for axis value labels, baked into the baseline:
        ' the high-precision readouts need more room than MSChart's auto margin gave.
        Const innerMarginLeftPadPx As Double = 100.0
        Const innerMarginRightPadPx As Double = 80.0

        Dim CaptureOriginalChartAreaMargins = Sub()
                                                  originalLiveChartWidth = LiveAnalysisChart.Width
                                                  originalLiveChartHeight = LiveAnalysisChart.Height

                                                  For Each ca In liveChartAreasInOrder
                                                      Dim pos = ca.Position
                                                      originalAreaLeftMarginPx(ca) = (pos.X / 100.0) * originalLiveChartWidth
                                                      originalAreaRightMarginPx(ca) = ((100.0 - pos.X - pos.Width) / 100.0) * originalLiveChartWidth

                                                      Dim posWidthPx As Double = (pos.Width / 100.0) * originalLiveChartWidth
                                                      Dim inner = ca.InnerPlotPosition
                                                      originalAreaInnerLeftMarginPx(ca) = (inner.X / 100.0) * posWidthPx + innerMarginLeftPadPx
                                                      originalAreaInnerRightMarginPx(ca) = ((100.0 - inner.X - inner.Width) / 100.0) * posWidthPx + innerMarginRightPadPx

                                                      ' No vertical padding: the X time labels are short and fixed-height.
                                                      Dim heightPx As Double = (pos.Height / 100.0) * originalLiveChartHeight
                                                      originalAreaHeightPx(ca) = heightPx
                                                      originalAreaInnerTopMarginPx(ca) = (inner.Y / 100.0) * heightPx
                                                      originalAreaInnerBottomMarginPx(ca) = ((100.0 - inner.Y - inner.Height) / 100.0) * heightPx
                                                  Next

                                                  Dim measPos = areaMeasurement.Position
                                                  Dim statsPos = areaStatistics.Position
                                                  Dim tempPos = areaTemperature.Position

                                                  ' Extra pad in the gap between stacked areas so the upper area's shared X labels
                                                  ' don't overlap the next area's top labels.
                                                  Const gapPadPx As Double = 20.0
                                                  Const topMarginPadPx As Double = 50.0

                                                  originalAreaTopMarginPx = (measPos.Y / 100.0) * originalLiveChartHeight + topMarginPadPx
                                                  originalGapMeasStatsPx = ((statsPos.Y - (measPos.Y + measPos.Height)) / 100.0) * originalLiveChartHeight + gapPadPx
                                                  originalGapStatsTempPx = ((tempPos.Y - (statsPos.Y + statsPos.Height)) / 100.0) * originalLiveChartHeight + gapPadPx
                                                  originalAreaBottomMarginPx = ((100.0 - tempPos.Y - tempPos.Height) / 100.0) * originalLiveChartHeight

                                                  ' Nudges the title in from the popup's edge by roughly one
                                                  ' character's width at its font size, so it isn't sitting
                                                  ' flush against the very edge of the window.
                                                  Const titleEdgeInsetPx As Double = 10.0

                                                  For Each t In liveChartTitlesInOrder
                                                      Dim tp = t.Position
                                                      originalTitleWidthPx(t) = (tp.Width / 100.0) * originalLiveChartWidth

                                                      ' Right-docked titles hold their distance from the right edge instead.
                                                      If t.Docking = DataVisualization.Charting.Docking.Right Then
                                                          originalTitleRightPx(t) = originalLiveChartWidth - ((tp.X + tp.Width) / 100.0) * originalLiveChartWidth + titleEdgeInsetPx
                                                      Else
                                                          originalTitleLeftPx(t) = (tp.X / 100.0) * originalLiveChartWidth + titleEdgeInsetPx
                                                      End If

                                                      Dim ownerPos = titleOwnerArea(t).Position
                                                      originalTitleYFractionOfArea(t) = (tp.Y - ownerPos.Y) / ownerPos.Height
                                                      originalTitleHeightPct(t) = tp.Height
                                                  Next
                                              End Sub

        Dim ApplyChartAreaMargins = Sub()
                                        If LiveAnalysisChart.Width <= 0 OrElse LiveAnalysisChart.Height <= 0 Then Exit Sub

                                        ' Horizontal - each area's own left/right margins held fixed.
                                        For Each ca In liveChartAreasInOrder
                                            Dim leftPct As Single = CSng((originalAreaLeftMarginPx(ca) / LiveAnalysisChart.Width) * 100.0)
                                            Dim rightPct As Single = CSng((originalAreaRightMarginPx(ca) / LiveAnalysisChart.Width) * 100.0)
                                            Dim widthPct As Single = 100.0F - leftPct - rightPct
                                            ca.Position = New DataVisualization.Charting.ElementPosition(leftPct, ca.Position.Y, widthPct, ca.Position.Height)

                                            Dim positionWidthPx As Double = (widthPct / 100.0) * LiveAnalysisChart.Width
                                            If positionWidthPx > 0 Then
                                                Dim innerLeftPct As Single = CSng((originalAreaInnerLeftMarginPx(ca) / positionWidthPx) * 100.0)
                                                Dim innerRightPct As Single = CSng((originalAreaInnerRightMarginPx(ca) / positionWidthPx) * 100.0)
                                                Dim innerWidthPct As Single = 100.0F - innerLeftPct - innerRightPct
                                                ca.InnerPlotPosition = New DataVisualization.Charting.ElementPosition(
                                                    innerLeftPct, ca.InnerPlotPosition.Y, innerWidthPct, ca.InnerPlotPosition.Height)
                                            End If
                                        Next

                                        ' Vertical - top/gap/gap/bottom margins held fixed, the three
                                        ' plot areas splitting the remaining height in their original ratio.
                                        Dim totalOriginalHeightPx As Double =
                                            originalAreaHeightPx(areaMeasurement) + originalAreaHeightPx(areaStatistics) + originalAreaHeightPx(areaTemperature)
                                        Dim remainingHeightPx As Double =
                                            LiveAnalysisChart.Height - originalAreaTopMarginPx - originalGapMeasStatsPx - originalGapStatsTempPx - originalAreaBottomMarginPx
                                        If remainingHeightPx <= 0 Then Exit Sub

                                        Dim measHeightPx As Double = remainingHeightPx * (originalAreaHeightPx(areaMeasurement) / totalOriginalHeightPx)
                                        Dim statsHeightPx As Double = remainingHeightPx * (originalAreaHeightPx(areaStatistics) / totalOriginalHeightPx)
                                        Dim tempHeightPx As Double = remainingHeightPx - measHeightPx - statsHeightPx
                                        If measHeightPx <= 0 OrElse statsHeightPx <= 0 OrElse tempHeightPx <= 0 Then Exit Sub

                                        Dim measYPx As Double = originalAreaTopMarginPx
                                        Dim statsYPx As Double = measYPx + measHeightPx + originalGapMeasStatsPx
                                        Dim tempYPx As Double = statsYPx + statsHeightPx + originalGapStatsTempPx

                                        Dim heights As New Dictionary(Of DataVisualization.Charting.ChartArea, Double) From {
                                            {areaMeasurement, measHeightPx}, {areaStatistics, statsHeightPx}, {areaTemperature, tempHeightPx}
                                        }
                                        Dim tops As New Dictionary(Of DataVisualization.Charting.ChartArea, Double) From {
                                            {areaMeasurement, measYPx}, {areaStatistics, statsYPx}, {areaTemperature, tempYPx}
                                        }

                                        For Each ca In liveChartAreasInOrder
                                            Dim yPct As Single = CSng((tops(ca) / LiveAnalysisChart.Height) * 100.0)
                                            Dim heightPct As Single = CSng((heights(ca) / LiveAnalysisChart.Height) * 100.0)
                                            ca.Position = New DataVisualization.Charting.ElementPosition(ca.Position.X, yPct, ca.Position.Width, heightPct)

                                            Dim innerTopPct As Single = CSng((originalAreaInnerTopMarginPx(ca) / heights(ca)) * 100.0)
                                            Dim innerBottomPct As Single = CSng((originalAreaInnerBottomMarginPx(ca) / heights(ca)) * 100.0)
                                            ca.InnerPlotPosition = New DataVisualization.Charting.ElementPosition(
                                                ca.InnerPlotPosition.X, innerTopPct, ca.InnerPlotPosition.Width, 100.0F - innerTopPct - innerBottomPct)
                                        Next

                                        ' Rotated titles keep a fixed pixel distance from the left/right edges; Y/Height track
                                        ' their owning ChartArea using the fraction captured at baseline.
                                        For Each t In liveChartTitlesInOrder
                                            Dim widthPct As Single = CSng((originalTitleWidthPx(t) / LiveAnalysisChart.Width) * 100.0)

                                            Dim leftPct As Single
                                            If t.Docking = DataVisualization.Charting.Docking.Right Then
                                                Dim leftPx As Double = LiveAnalysisChart.Width - originalTitleRightPx(t) - (widthPct / 100.0) * LiveAnalysisChart.Width
                                                leftPct = CSng((leftPx / LiveAnalysisChart.Width) * 100.0)
                                            Else
                                                leftPct = CSng((originalTitleLeftPx(t) / LiveAnalysisChart.Width) * 100.0)
                                            End If

                                            Dim ownerPos = titleOwnerArea(t).Position
                                            Dim yPct As Single = CSng(ownerPos.Y + originalTitleYFractionOfArea(t) * ownerPos.Height)

                                            t.Position = New DataVisualization.Charting.ElementPosition(leftPct, yPct, widthPct, originalTitleHeightPct(t))
                                        Next
                                    End Sub

        Dim RepositionLiveToggles = Sub()
                                        Dim cw As Double = LiveAnalysisChart.ClientSize.Width
                                        Dim ch As Double = LiveAnalysisChart.ClientSize.Height

                                        Dim rowHeightPx As Integer = 22
                                        Dim colWidthPx As Integer = 80   ' narrower than the original 110, but wide enough for "PPM DEV" to not clip
                                        Dim miscColWidthPx As Integer = 130  ' Misc. labels ("Fast Rendering" etc.) are longer than trace names
                                        Dim pad As Integer = 5         ' margin around the checkboxes
                                        Dim gap As Integer = 20        ' space between the group boxes
                                        Dim topPct As Double = 0       ' % from top of form

                                        ' Wide enough for the checkbox columns or the title, whichever is larger;
                                        ' widthMultiplier adds headroom for titles that grow (editable device names).
                                        Dim GroupWidth = Function(gb As GroupBox, boxes As CheckBox(), colWidth As Integer, widthMultiplier As Double) As Integer
                                                             Dim numCols As Integer = CInt(Math.Ceiling(boxes.Length / 2.0))
                                                             Dim columnWidth As Integer = (colWidth * numCols) + (pad * 2)
                                                             Dim titleWidth As Integer = TextRenderer.MeasureText(gb.Text, gb.Font).Width + (pad * 2) + 20
                                                             Return CInt(Math.Max(columnWidth, titleWidth) * widthMultiplier)
                                                         End Function

                                        Dim widthDev1 As Integer = GroupWidth(gbDev1, dev1Boxes, colWidthPx, 1.0)
                                        Dim widthDev2 As Integer = GroupWidth(gbDev2, dev2Boxes, colWidthPx, 1.0)
                                        Dim widthTemp As Integer = GroupWidth(gbTemp, tempBoxes, colWidthPx, 1.0)
                                        Dim widthMisc As Integer = GroupWidth(gbMisc, miscBoxes, miscColWidthPx, 1.0)

                                        Dim totalWidth As Integer = widthDev1 + widthDev2 + widthTemp + widthMisc + (gap * 3)
                                        Dim startX As Integer = CInt((cw - totalWidth) / 2.0)

                                        Dim leftDev1 As Integer = startX
                                        Dim leftDev2 As Integer = leftDev1 + widthDev1 + gap
                                        Dim leftTemp As Integer = leftDev2 + widthDev2 + gap
                                        Dim leftMisc As Integer = leftTemp + widthTemp + gap

                                        Dim topPx As Integer = CInt(topPct / 100.0 * ch)

                                        Dim PlaceGroupBox = Sub(gb As GroupBox, boxes As CheckBox(), leftPx As Integer, colWidth As Integer, widthMultiplier As Double)
                                                                Dim numRows As Integer = Math.Min(2, boxes.Length)
                                                                Dim titleAllowance As Integer = TextRenderer.MeasureText(gb.Text, gb.Font).Height + 2

                                                                Dim contentHeight As Integer = rowHeightPx * numRows

                                                                gb.Location = New Point(leftPx, topPx)
                                                                gb.Size = New Size(GroupWidth(gb, boxes, colWidth, widthMultiplier), contentHeight + (pad * 2) + titleAllowance)

                                                                Dim dr As Rectangle = gb.DisplayRectangle
                                                                Dim hOffset As Integer = dr.Left + pad
                                                                Dim vOffset As Integer = dr.Top + Math.Max(pad, (dr.Height - contentHeight) \ 2)

                                                                For i As Integer = 0 To boxes.Length - 1
                                                                    Dim col As Integer = i \ 2
                                                                    Dim row As Integer = i Mod 2
                                                                    boxes(i).Location = New Point(hOffset + (col * colWidth), vOffset + (row * rowHeightPx))
                                                                    boxes(i).Width = colWidth - 2
                                                                Next
                                                            End Sub

                                        PlaceGroupBox(gbDev1, dev1Boxes, leftDev1, colWidthPx, 1.0)
                                        PlaceGroupBox(gbDev2, dev2Boxes, leftDev2, colWidthPx, 1.0)
                                        PlaceGroupBox(gbTemp, tempBoxes, leftTemp, colWidthPx, 1.0)
                                        PlaceGroupBox(gbMisc, miscBoxes, leftMisc, miscColWidthPx, 1.0)

                                        ' Bottom-aligned with Dev 2/Misc. since Temp. has fewer checkboxes and a shorter groupbox.
                                        btnResetLiveCharts.Width = gbTemp.Width
                                        btnResetLiveCharts.Location = New Point(gbTemp.Left, gbDev2.Bottom - btnResetLiveCharts.Height)
                                    End Sub

        CaptureOriginalChartAreaMargins()
        ApplyChartAreaMargins()

        RepositionLiveToggles()
        AddHandler LiveAnalysisChart.Resize, Sub(s, ev)
                                                 ApplyChartAreaMargins()
                                                 RepositionLiveToggles()
                                             End Sub




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

        ' Bottom-right resize grip - purely a visual cue that the window can
        ' be resized. Dragging it hands off to Windows' own native resize
        ' (WM_NCLBUTTONDOWN / HTBOTTOMRIGHT) rather than us tracking the drag.
        Dim liveAnalysisGrip As New PictureBox With {
            .Image = My.Resources.grip,
            .SizeMode = PictureBoxSizeMode.StretchImage,
            .Size = New Size(36, 36),
            .BackColor = Color.Transparent,
            .Cursor = Cursors.SizeNWSE,
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        }
        liveAnalysisGrip.Location = New Point(
            LiveAnalysisForm.ClientSize.Width - liveAnalysisGrip.Width,
            LiveAnalysisForm.ClientSize.Height - liveAnalysisGrip.Height)

        AddHandler liveAnalysisGrip.MouseDown,
        Sub(gripSender As Object, gripArgs As MouseEventArgs)
            If gripArgs.Button = MouseButtons.Left Then
                ReleaseCapture()
                SendMessage(LiveAnalysisForm.Handle, &HA1, 17, 0)   ' WM_NCLBUTTONDOWN, HTBOTTOMRIGHT
            End If
        End Sub

        LiveAnalysisForm.Controls.Add(liveAnalysisGrip)
        liveAnalysisGrip.BringToFront()

        AddHandler LiveAnalysisForm.FormClosed,
        Sub()
            RemoveHandler ButtonDev1Run.Click, RunButtonHandler
            RemoveHandler ButtonDev2Run.Click, RunButtonHandler
            RemoveHandler ButtonDev12Run.Click, RunButtonHandler
            RemoveHandler ButtonStart.Click, RunButtonHandler
            RemoveHandler ButtonEnd.Click, RunButtonHandler

            RemoveHandler ButtonReset.Click, ResetHandler
            RemoveHandler btncreate.Click, ConnectBothHandler
            RemoveHandler btncreate2.Click, ConnectDev1Handler
            RemoveHandler btncreate3.Click, ConnectDev2Handler

            LiveAnalysisChart = Nothing
            LiveAnalysisForm = Nothing
        End Sub


        ' Start chart at sample zero whenever opened.
        LiveAnalysisSample = 0

        ' Fresh Short-Term Mean window each time the pop-out is (re)opened.
        q1ShortTermMean.Clear() : sum1ShortTermMean = 0.0
        q2ShortTermMean.Clear() : sum2ShortTermMean = 0.0

        ' Remember current statistics counts so that only NEW
        ' readings received after opening the chart are plotted.
        LiveAnalysisLastStats1Count = Stats1Count
        LiveAnalysisLastStats2Count = Stats2Count

        ApplySquareCorners(LiveAnalysisForm)
        LiveAnalysisForm.Show()

    End Sub


    Private Sub AddLiveAnalysisSeries(seriesName As String,
                                  chartAreaName As String,
                                  seriesColor As Color)

        Dim s As New DataVisualization.Charting.Series(seriesName)

        ' FastLine skips anti-aliasing regardless of chart-level settings, which is
        ' what caused the moire/banding look when many points land close together in
        ' a rolling window - Line respects anti-aliasing and renders cleanly instead.
        s.ChartType = DataVisualization.Charting.SeriesChartType.Line

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
        Dim newDev1Sample As Boolean = Stats1Count <> LiveAnalysisLastStats1Count
        Dim newDev2Sample As Boolean = Stats2Count <> LiveAnalysisLastStats2Count

        ' Decide whether to advance the chart yet.
        Dim dev1CurrentlyRunning As Boolean = ButtonDev1Run.Text = "Stop" OrElse ButtonDev12Run.Text = "Stop"
        Dim dev2CurrentlyRunning As Boolean = ButtonDev2Run.Text = "Stop" OrElse ButtonDev12Run.Text = "Stop"

        Dim dev1Contributing As Boolean = dev1CurrentlyRunning AndAlso Stats1Count > 0
        Dim dev2Contributing As Boolean = dev2CurrentlyRunning AndAlso Stats2Count > 0

        Dim readyToAdvance As Boolean

        If dev1Contributing AndAlso dev2Contributing Then
            readyToAdvance = newDev1Sample AndAlso newDev2Sample
        Else
            readyToAdvance = newDev1Sample OrElse newDev2Sample
        End If

        If readyToAdvance = False Then Exit Sub

        ' Suspend repaint so the chart never paints a half-updated mix of points and axis ranges.
        LiveAnalysisChart.SuspendLayout()

        Try

            ' Advance ONE X-axis sample
            LiveAnalysisSample += 1

            Dim x As Double = LiveAnalysisSample

            ' Device 1
            If newDev1Sample = True Then

                Dim dev1Raw As Double = CDbl(Val(NormalizeNumericResponse(txtr1a.Text)))

                Dim dev1MeanToPlot As Double = Stats1Mean
                If chkShortTermMean IsNot Nothing AndAlso chkShortTermMean.Checked Then
                    q1ShortTermMean.Enqueue(dev1Raw) : sum1ShortTermMean += dev1Raw
                    If q1ShortTermMean.Count > ShortTermMeanWindow Then sum1ShortTermMean -= q1ShortTermMean.Dequeue()
                    dev1MeanToPlot = sum1ShortTermMean / q1ShortTermMean.Count
                End If

                LiveAnalysisChart.Series("Device 1").Points.AddXY(x, dev1Raw)
                LiveAnalysisChart.Series("Dev 1 Mean").Points.AddXY(x, dev1MeanToPlot)
                LiveAnalysisChart.Series("Dev 1 STDEV").Points.AddXY(x, Stats1StdevCurrent)
                LiveAnalysisChart.Series("Dev 1 SEM").Points.AddXY(x, Stats1SEMCurrent)
                ' NaN/Infinity points make MSChart's axis auto-scaling throw on the next repaint, so guard here too.
                LiveAnalysisChart.Series("Dev 1 PPM Deviation").Points.AddXY(
                    x, If(Double.IsNaN(Stats1DeviationCurrent) OrElse Double.IsInfinity(Stats1DeviationCurrent), 0, Stats1DeviationCurrent))

                LiveAnalysisLastStats1Count = Stats1Count

            End If

            ' Device 2
            If newDev2Sample = True Then

                Dim dev2Raw As Double = CDbl(Val(NormalizeNumericResponse(txtr2a.Text)))

                Dim dev2MeanToPlot As Double = Stats2Mean
                If chkShortTermMean IsNot Nothing AndAlso chkShortTermMean.Checked Then
                    q2ShortTermMean.Enqueue(dev2Raw) : sum2ShortTermMean += dev2Raw
                    If q2ShortTermMean.Count > ShortTermMeanWindow Then sum2ShortTermMean -= q2ShortTermMean.Dequeue()
                    dev2MeanToPlot = sum2ShortTermMean / q2ShortTermMean.Count
                End If

                LiveAnalysisChart.Series("Device 2").Points.AddXY(x, dev2Raw)
                LiveAnalysisChart.Series("Dev 2 Mean").Points.AddXY(x, dev2MeanToPlot)
                LiveAnalysisChart.Series("Dev 2 STDEV").Points.AddXY(x, Stats2StdevCurrent)
                LiveAnalysisChart.Series("Dev 2 SEM").Points.AddXY(x, Stats2SEMCurrent)
                ' Defence in depth - see the same note on the Dev 1 PPM point above.
                LiveAnalysisChart.Series("Dev 2 PPM Deviation").Points.AddXY(
                    x, If(Double.IsNaN(Stats2DeviationCurrent) OrElse Double.IsInfinity(Stats2DeviationCurrent), 0, Stats2DeviationCurrent))

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
                SetAdaptiveDecimalFormat(LiveAnalysisChart.ChartAreas("Measurement").AxisY, 10)

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
            EnsureMinimumGridLines(LiveAnalysisChart.ChartAreas("Temperature").AxisY, 5)

        Finally
            LiveAnalysisChart.ResumeLayout()
        End Try

    End Sub

    ' Falls back to a safe range when an axis has no valid Min/Max (all its traces disabled), which throws on Paint.
    Private Sub EnsureValidAxisScale(axis As DataVisualization.Charting.Axis)

        If Double.IsNaN(axis.Minimum) OrElse Double.IsNaN(axis.Maximum) OrElse
           Double.IsInfinity(axis.Minimum) OrElse Double.IsInfinity(axis.Maximum) OrElse
           axis.Minimum = axis.Maximum Then

            axis.Minimum = 0
            axis.Maximum = 1
            axis.Interval = 0.5

        End If

    End Sub

    ' Divides the axis into exactly minDivisions equal gridlines, then
    ' floors/ceils Minimum/Maximum onto that Interval so a gridline always
    ' lands on both the top and bottom edge.
    Private Sub EnsureMinimumGridLines(axis As DataVisualization.Charting.Axis, minDivisions As Integer)

        If minDivisions > 0 AndAlso axis.Maximum > axis.Minimum Then

            axis.Interval = (axis.Maximum - axis.Minimum) / minDivisions

            axis.Minimum = Math.Floor(axis.Minimum / axis.Interval) * axis.Interval
            axis.Maximum = Math.Ceiling(axis.Maximum / axis.Interval) * axis.Interval

        End If

    End Sub

    ' Sets the axis label format to however many decimal places its current
    ' Interval actually needs to show distinct values, never fewer than
    ' minDecimals.
    Private Sub SetAdaptiveDecimalFormat(axis As DataVisualization.Charting.Axis, minDecimals As Integer)

        If axis.Interval > 0 Then

            Dim neededDecimals As Integer = Math.Max(minDecimals, CInt(Math.Ceiling(-Math.Log10(axis.Interval))) + 1)
            axis.LabelStyle.Format = "0." & New String("0"c, neededDecimals)

        End If

    End Sub


    ' Live Analysis is available whenever any device is running, regardless of the DATA tab chart's Pause state.
    Private Sub UpdateLiveChartPopoutAvailability()

        ButtonLiveChartPopout.Enabled =
            ButtonDev1Run.Text = "Stop" OrElse
            ButtonDev2Run.Text = "Stop" OrElse
            ButtonDev12Run.Text = "Stop"

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
