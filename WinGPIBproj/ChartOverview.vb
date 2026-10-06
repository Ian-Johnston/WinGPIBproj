Imports System.Globalization

' Playback Chart "Zoom Overview": a small floating window, owned by the Playback Chart so it stays above it, showing the
' whole run with a rectangle for the main chart's current X range. Drag the rectangle to pan, drag its edges to zoom, click
' elsewhere to jump there, double-click for the full view. The overview is a separate plot of about 1000 min/max points,
' so it is cheap to draw whatever the size of the CSV; the main chart is redrawn throttled while the mouse is dragging.
Partial Public Class Chart

    Private Chart2OverviewForm As Form = Nothing
    Private Chart2OverviewPlot As ScottPlot.WinForms.FormsPlot = Nothing
    Private Chart2OverviewViewSpan As ScottPlot.Plottables.HorizontalSpan = Nothing     ' the main chart's current X range (invisible: holds the range)
    Private Chart2OverviewViewBox As ScottPlot.Plottables.Rectangle = Nothing           ' the closed box drawn for that range
    Private Chart2OverviewRegionSpan As ScottPlot.Plottables.HorizontalSpan = Nothing   ' the Regional Stats band, if showing
    Private Chart2OverviewHoverLabel As Label = Nothing                                 ' time under the mouse pointer
    Private Chart2OverviewHoverTimer As System.Windows.Forms.Timer = Nothing                                 ' hides it when the mouse has not moved over the window for a while
    Private Const ChartOverviewHoverTimeoutMs As Integer = 2000
    Private Chart2OverviewViewText As ScottPlot.Plottables.Text = Nothing               ' the rectangle's start-end minutes, when zoomed in

    Private Chart2OverviewKey As String = ""             ' what the overview was last built from (see Chart2OverviewDataKey)
    Private Chart2OverviewRegionKey As String = ""
    Private Chart2OverviewSync As Boolean = False        ' True while code (not the user) ticks the button or closes the window
    Private Chart2OverviewBounds As Rectangle = Rectangle.Empty    ' where the window was last left (this session)
    Private Chart2OverviewMaxX As Double = 0             ' the overview's X range is 0 to this (the last sample index)

    Private Chart2OverviewViewX1 As Double = Double.NaN  ' the rectangle as drawn
    Private Chart2OverviewViewX2 As Double = Double.NaN
    Private Chart2OverviewDragMode As Integer = 0        ' 0 = not dragging, 1 = move, 2 = left edge, 3 = right edge
    Private Chart2OverviewDragOffset As Double = 0       ' mouse X minus rectangle left, when moving
    Private Chart2OverviewLastClickTime As DateTime = DateTime.MinValue
    Private Chart2OverviewLastClickPixel As ScottPlot.Pixel = ScottPlot.Pixel.NaN
    Private Chart2OverviewClock As Stopwatch = Stopwatch.StartNew()
    Private Chart2OverviewLastApplyMs As Long = 0

    Private Const ChartOverviewBuckets As Integer = 1000       ' min/max points drawn however large the CSV
    Private Const ChartOverviewApplyIntervalMs As Integer = 25 ' main chart redraw rate while dragging
    Private Const ChartOverviewEdgePixels As Single = 6.0F     ' how close to an edge counts as grabbing it

    Private Sub ButtonZoomOverview_CheckedChanged(sender As Object, e As EventArgs) Handles ButtonZoomOverview.CheckedChanged

        If Chart2OverviewSync Then Exit Sub

        If ButtonZoomOverview.Checked Then
            Chart2OpenOverview()
        Else
            Chart2CloseOverview()
        End If

    End Sub

    ' Right-click > Zoom Overview: flips the button, which opens or closes the window.
    Private Sub Chart2MenuZoomOverview(plot As ScottPlot.Plot)

        If ButtonZoomOverview.Enabled Then ButtonZoomOverview.Checked = Not ButtonZoomOverview.Checked

    End Sub

    Private Sub Chart2SetOverviewChecked(value As Boolean)

        Chart2OverviewSync = True
        ButtonZoomOverview.Checked = value
        Chart2OverviewSync = False

    End Sub

    Private Sub Chart2CloseOverview()

        If Chart2OverviewForm IsNot Nothing AndAlso Not Chart2OverviewForm.IsDisposed Then
            Chart2OverviewSync = True
            Chart2OverviewForm.Close()
            Chart2OverviewSync = False
        End If

        Chart2SetOverviewChecked(False)

    End Sub

    Private Sub Chart2OpenOverview()

        If Not (ChartLoaded AndAlso CSVfileok) Then
            Chart2SetOverviewChecked(False)
            Exit Sub
        End If

        If Chart2OverviewForm IsNot Nothing AndAlso Not Chart2OverviewForm.IsDisposed Then
            Chart2OverviewForm.BringToFront()
            Exit Sub
        End If

        Dim light As Boolean = CheckBoxColours.Checked

        ' Fixed size. Where it opens: where it was last left, otherwise the lower right of the main chart. Always kept on the screen.
        Dim w As Integer = 760
        Dim h As Integer = 190
        Dim loc As Point

        If Not Chart2OverviewBounds.IsEmpty Then
            loc = Chart2OverviewBounds.Location
        Else
            loc = Me.PointToScreen(New Point(Math.Max(0, FormsPlot2.Right - w - 14), Math.Max(0, FormsPlot2.Bottom - h - 40)))
        End If

        Dim area As Rectangle = Screen.FromControl(Me).WorkingArea
        w = Math.Min(w, area.Width)
        h = Math.Min(h, area.Height)
        loc.X = Math.Max(area.Left, Math.Min(loc.X, area.Right - w))
        loc.Y = Math.Max(area.Top, Math.Min(loc.Y, area.Bottom - h))

        Dim frm As New Form With {
            .Text = "WinGPIB - Zoom Overview",
            .StartPosition = FormStartPosition.Manual,
            .Location = loc,
            .Size = New Size(w, h),
            .ShowIcon = False,
            .ShowInTaskbar = False,
            .MinimizeBox = False,
            .MaximizeBox = False
        }

        Dim overviewPlot As New ScottPlot.WinForms.FormsPlot With {.Dock = DockStyle.Fill}
        overviewPlot.UserInputProcessor.Disable()       ' no built-in pan/zoom: the mouse handlers below do the work
        frm.Controls.Add(overviewPlot)

        ' Thin title strip in place of the Windows title bar (added after the plot so the plot fills what is left below it)
        AddPopupTitleStrip(frm, "WinGPIB - Zoom Overview")

        AddHandler overviewPlot.MouseDown, AddressOf Chart2OverviewMouseDown
        AddHandler overviewPlot.MouseMove, AddressOf Chart2OverviewMouseMove
        AddHandler overviewPlot.MouseUp, AddressOf Chart2OverviewMouseUp
        AddHandler overviewPlot.MouseLeave, Sub(snd, ev) Chart2OverviewHideHover()

        ' Small time read-out that follows the mouse pointer (a plain label over the plot, so hovering never redraws the plot)
        Dim hover As New Label With {
            .AutoSize = True,
            .Visible = False,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
            .Padding = New Padding(3, 1, 3, 1),
            .BorderStyle = BorderStyle.FixedSingle
        }
        overviewPlot.Controls.Add(hover)
        hover.BringToFront()

        ' The read-out disappears 2 seconds after the mouse last moved over the window (so it does not stay up after the mouse
        ' has left it)
        Dim hoverTimer As New System.Windows.Forms.Timer With {.Interval = ChartOverviewHoverTimeoutMs}
        AddHandler hoverTimer.Tick, Sub(snd, ev) Chart2OverviewHideHover()
        Chart2OverviewHoverTimer = hoverTimer

        AddHandler frm.Shown, Sub(snd, ev) overviewPlot.Refresh()

        AddHandler frm.FormClosing, Sub(snd, ev)
                                        Chart2OverviewBounds = frm.Bounds
                                    End Sub

        ' Closing the window unticks the button (unless the code is just closing it).
        AddHandler frm.FormClosed, Sub(snd, ev)
                                       If Chart2OverviewForm Is frm Then
                                           Chart2OverviewForm = Nothing
                                           Chart2OverviewPlot = Nothing
                                           Chart2OverviewHoverLabel = Nothing
                                           If Chart2OverviewHoverTimer IsNot Nothing Then
                                               Chart2OverviewHoverTimer.Stop()
                                               Chart2OverviewHoverTimer.Dispose()
                                               Chart2OverviewHoverTimer = Nothing
                                           End If
                                           Chart2OverviewViewSpan = Nothing
                                           Chart2OverviewViewBox = Nothing
                                           Chart2OverviewRegionSpan = Nothing
                                           Chart2OverviewDragMode = 0
                                       End If
                                       If Not Chart2OverviewSync Then Chart2SetOverviewChecked(False)
                                   End Sub

        ApplySquareCorners(frm)

        Chart2OverviewForm = frm
        Chart2OverviewPlot = overviewPlot
        Chart2OverviewHoverLabel = hover
        Chart2OverviewKey = ""
        Chart2OverviewRegionKey = ""
        Chart2OverviewViewX1 = Double.NaN
        Chart2OverviewViewX2 = Double.NaN
        Chart2OverviewDragMode = 0

        Chart2SetOverviewChecked(True)
        frm.Show(Me)        ' owned by the Playback Chart: stays above it, closes with it, not above other programs

        Chart2SyncOverview()

    End Sub

    ' ---- Shared pop-up chrome (Zoom Overview, Allan Deviation, Histogram) ----

    Private Const PopupTitleStripHeight As Integer = 18

    Private Class PopupTitleStrip
        Public Bar As Panel
        Public TitleLabel As Label
        Public CloseLabel As Label
    End Class

    ' Replaces the Windows title bar (about twice as tall) with a thin strip in the same colours: drag it to move the window,
    ' X to close. Add it AFTER the form's other docked content so that content fills what is left below it. The form gets a 1px
    ' border. Keeps the strip in frm.Tag so SetPopupTitle can change the title.
    Private Function AddPopupTitleStrip(frm As Form, title As String) As PopupTitleStrip

        Dim strip As New PopupTitleStrip

        frm.FormBorderStyle = FormBorderStyle.None
        frm.Padding = New Padding(1)          ' the form's own colour shows round the edge as a 1px border
        frm.Text = title

        strip.Bar = New Panel With {.Dock = DockStyle.Top, .Height = PopupTitleStripHeight}
        strip.TitleLabel = New Label With {
            .Text = title,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleLeft,
            .Font = New Font("Segoe UI", 9.0F),
            .Padding = New Padding(6, 0, 0, 0)
        }
        strip.CloseLabel = New Label With {
            .Text = ChrW(&H2715),
            .Dock = DockStyle.Right,
            .Width = 28,
            .TextAlign = ContentAlignment.MiddleCenter,
            .Font = New Font("Segoe UI Symbol", 9.0F),
            .Cursor = Cursors.Default
        }
        strip.Bar.Controls.Add(strip.TitleLabel)
        strip.Bar.Controls.Add(strip.CloseLabel)
        frm.Controls.Add(strip.Bar)
        frm.Tag = strip

        Dim dragWindow As MouseEventHandler =
            Sub(snd As Object, ev As MouseEventArgs)
                If ev.Button = MouseButtons.Left Then
                    ReleaseCapture()
                    SendMessage(frm.Handle, &HA1, 2, 0)   ' WM_NCLBUTTONDOWN, HTCAPTION: move the window
                End If
            End Sub
        AddHandler strip.Bar.MouseDown, dragWindow
        AddHandler strip.TitleLabel.MouseDown, dragWindow

        AddHandler strip.CloseLabel.Click, Sub(snd, ev) frm.Close()
        AddHandler strip.CloseLabel.MouseEnter, Sub(snd, ev)
                                                    strip.CloseLabel.BackColor = Color.FromArgb(196, 43, 28)
                                                    strip.CloseLabel.ForeColor = Color.White
                                                End Sub
        AddHandler strip.CloseLabel.MouseLeave, Sub(snd, ev) ApplyPopupTitleColors(frm, strip)

        ApplyPopupTitleColors(frm, strip)

        Return strip

    End Function

    ' The standard light title bar colours (not the app's Light Mode)
    Private Sub ApplyPopupTitleColors(frm As Form, strip As PopupTitleStrip)

        Dim barBack As Color = SystemTitleBarColor()
        Dim barText As Color = TitleTextColorFor(barBack)

        frm.BackColor = SystemTitleBorderColor()
        strip.Bar.BackColor = barBack
        strip.TitleLabel.BackColor = barBack
        strip.TitleLabel.ForeColor = barText
        strip.CloseLabel.BackColor = barBack
        strip.CloseLabel.ForeColor = barText

    End Sub

    ' Title text of a pop-up made with AddPopupTitleStrip (also sets the form's own Text).
    Private Sub SetPopupTitle(frm As Form, title As String)

        If frm Is Nothing OrElse frm.IsDisposed Then Exit Sub

        frm.Text = title

        Dim strip As PopupTitleStrip = TryCast(frm.Tag, PopupTitleStrip)
        If strip IsNot Nothing Then strip.TitleLabel.Text = title

    End Sub

    ' Resize grip for a frameless form (Windows' own resize needs a frame): drag to change the size, never below MinimumSize.
    Private Sub HookPopupResizeGrip(grip As PictureBox, frm As Form)

        Dim dragging As Boolean = False
        Dim startMouse As Point
        Dim startSize As Size

        AddHandler grip.MouseDown,
            Sub(snd As Object, ev As MouseEventArgs)
                If ev.Button = MouseButtons.Left Then
                    dragging = True
                    startMouse = Control.MousePosition
                    startSize = frm.Size
                End If
            End Sub

        AddHandler grip.MouseMove,
            Sub(snd As Object, ev As MouseEventArgs)
                If Not dragging Then Exit Sub
                Dim mouse As Point = Control.MousePosition
                frm.Size = New Size(Math.Max(frm.MinimumSize.Width, startSize.Width + mouse.X - startMouse.X),
                                    Math.Max(frm.MinimumSize.Height, startSize.Height + mouse.Y - startMouse.Y))
            End Sub

        AddHandler grip.MouseUp, Sub(snd, ev) dragging = False

    End Sub

    ' The standard light Windows title bar colour (the other windows' own title bars, which the app does not ask Windows to darken)
    Private Function SystemTitleBarColor() As Color

        Return Color.FromArgb(243, 243, 243)

    End Function

    Private Function SystemTitleBorderColor() As Color

        Return Color.FromArgb(160, 160, 160)

    End Function

    Private Function TitleTextColorFor(back As Color) As Color

        Dim luminance As Double = 0.299 * back.R + 0.587 * back.G + 0.114 * back.B
        Return If(luminance > 150, Color.Black, Color.White)

    End Function

    ' What the overview was built from: both devices' visibility, sample counts and a few sample values (so a different
    ' file, or a changed Avg, is noticed even if the count is the same), the theme and the time scale.
    Private Function Chart2OverviewDataKey() As String

        Dim parts As New List(Of String)

        For pass As Integer = 1 To 2

            Dim data As List(Of ScottPlot.Coordinates) = If(pass = 1, Chart2Dev1Data, Chart2Dev2Data)
            Dim series As ScottPlot.Plottables.Scatter = If(pass = 1, Chart2Dev1Series, Chart2Dev2Series)

            parts.Add(If(series IsNot Nothing AndAlso series.IsVisible, "1", "0"))

            If data Is Nothing Then
                parts.Add("0")
            Else
                parts.Add(data.Count.ToString())
                If data.Count > 0 Then
                    parts.Add(data(0).Y.ToString("R", CultureInfo.InvariantCulture))
                    parts.Add(data(data.Count \ 2).Y.ToString("R", CultureInfo.InvariantCulture))
                    parts.Add(data(data.Count - 1).Y.ToString("R", CultureInfo.InvariantCulture))
                End If
            End If

        Next

        parts.Add(CheckBoxColours.Checked.ToString())
        parts.Add(Chart2MinsPerSample.ToString("R", CultureInfo.InvariantCulture))

        Return String.Join("|", parts)

    End Function

    ' Rebuilds the overview from the main chart's Dev 1 / Dev 2 data: a min/max envelope plus a mean line per visible device.
    Private Sub Chart2RebuildOverview()

        If Chart2OverviewPlot Is Nothing OrElse Chart2OverviewForm Is Nothing Then Exit Sub

        Dim light As Boolean = CheckBoxColours.Checked
        Dim plot As ScottPlot.Plot = Chart2OverviewPlot.Plot

        plot.Clear()

        ' Theme: plot (the window frame and title strip follow the Windows title bar colours, not Light Mode)

        If Chart2OverviewHoverLabel IsNot Nothing Then
            Chart2OverviewHoverLabel.BackColor = If(light, Color.FromArgb(255, 255, 225), Color.FromArgb(50, 50, 50))
            Chart2OverviewHoverLabel.ForeColor = If(light, Color.Black, Color.White)
        End If

        If light Then
            plot.FigureBackground.Color = ScottPlot.Colors.White
            plot.DataBackground.Color = ScottPlot.Colors.White
            plot.Axes.Color(New ScottPlot.Color(Color.Black))
            plot.Grid.MajorLineColor = New ScottPlot.Color(Color.FromArgb(155, 185, 185, 185))
        Else
            plot.FigureBackground.Color = New ScottPlot.Color(Color.FromArgb(30, 30, 30))
            plot.DataBackground.Color = ScottPlot.Colors.Black
            plot.Axes.Color(New ScottPlot.Color(Color.FromArgb(220, 220, 220)))
            plot.Grid.MajorLineColor = New ScottPlot.Color(Color.FromArgb(255, 70, 70, 70))
        End If

        Dim maxX As Double = 0
        Dim yLo As Double = Double.MaxValue
        Dim yHi As Double = Double.MinValue
        Dim drawn As Boolean = False

        For pass As Integer = 1 To 2

            Dim data As List(Of ScottPlot.Coordinates) = If(pass = 1, Chart2Dev1Data, Chart2Dev2Data)
            Dim series As ScottPlot.Plottables.Scatter = If(pass = 1, Chart2Dev1Series, Chart2Dev2Series)

            If data Is Nothing OrElse series Is Nothing OrElse Not series.IsVisible OrElse data.Count < 2 Then Continue For

            Dim n As Integer = data.Count
            Dim b As Integer = Math.Min(ChartOverviewBuckets, n)

            Dim xs(b - 1) As Double
            Dim lows(b - 1) As Double
            Dim highs(b - 1) As Double
            Dim means(b - 1) As Double

            For k As Integer = 0 To b - 1

                Dim i0 As Integer = CInt((CLng(k) * n) \ b)
                Dim i1 As Integer = CInt((CLng(k + 1) * n) \ b) - 1
                If i1 < i0 Then i1 = i0

                Dim lo As Double = Double.MaxValue
                Dim hi As Double = Double.MinValue
                Dim sum As Double = 0
                Dim count As Integer = 0

                For i As Integer = i0 To i1
                    Dim y As Double = data(i).Y
                    If Double.IsNaN(y) Then Continue For
                    If y < lo Then lo = y
                    If y > hi Then hi = y
                    sum += y
                    count += 1
                Next

                If count = 0 Then
                    ' Nothing usable in this slice: carry the previous value along
                    Dim carried As Double = If(k > 0, means(k - 1), 0.0)
                    lo = carried
                    hi = carried
                    sum = carried
                    count = 1
                End If

                xs(k) = (data(i0).X + data(i1).X) / 2.0
                lows(k) = lo
                highs(k) = hi
                means(k) = sum / count

                If lo < yLo Then yLo = lo
                If hi > yHi Then yHi = hi

            Next

            Dim envelope As ScottPlot.Plottables.FillY = plot.Add.FillY(xs, lows, highs)
            envelope.FillColor = series.Color.WithAlpha(0.45)
            envelope.LineColor = series.Color.WithAlpha(0.8)
            envelope.LineWidth = 1
            envelope.MarkerSize = 0

            Dim meanLine As ScottPlot.Plottables.Scatter = plot.Add.Scatter(xs, means)
            meanLine.Color = series.Color
            meanLine.LineWidth = 1
            meanLine.MarkerStyle.IsVisible = False

            maxX = Math.Max(maxX, data(n - 1).X)
            drawn = True

        Next

        Chart2OverviewViewSpan = Nothing
        Chart2OverviewViewBox = Nothing
        Chart2OverviewRegionSpan = Nothing
        Chart2OverviewViewText = Nothing
        Chart2OverviewViewX1 = Double.NaN
        Chart2OverviewViewX2 = Double.NaN
        Chart2OverviewRegionKey = ""

        ' Keep the left axis line as the left-hand border, but with no ticks or labels (Y values are not needed here)
        plot.Axes.Left.IsVisible = True

        ' Equal fixed margins left and right (room for the first and last X labels, which are centred on the ends of the run)
        plot.Layout.Fixed(New ScottPlot.PixelPadding(14, 14, 26, 4))
        plot.Axes.Left.TickGenerator = New ScottPlot.TickGenerators.NumericManual()
        plot.Axes.Left.MajorTickStyle.Length = 0
        plot.Axes.Left.MinorTickStyle.Length = 0

        If Not drawn OrElse maxX <= 0 Then

            Chart2OverviewMaxX = 0
            plot.Axes.SetLimits(0, 1, 0, 1)
            plot.Axes.Bottom.TickGenerator = New ScottPlot.TickGenerators.NumericManual()

            Dim note As ScottPlot.Plottables.Annotation = plot.Add.Annotation("No visible Dev 1 / Dev 2 trace", ScottPlot.Alignment.MiddleCenter)
            note.LabelFontSize = 12
            note.LabelFontColor = If(light, ScottPlot.Colors.Black, ScottPlot.Colors.White)
            note.LabelBackgroundColor = ScottPlot.Colors.Transparent

            Chart2OverviewPlot.Refresh()
            Exit Sub

        End If

        Chart2OverviewMaxX = maxX

        Dim yPad As Double = (yHi - yLo) * 0.08
        If yPad = 0 Then yPad = If(yHi = 0, 0.001, Math.Abs(yHi) / 1000)
        plot.Axes.SetLimits(0, maxX, yLo - yPad, yHi + yPad)

        ' X labels in minutes, as on the main chart (sample numbers if the time scale isn't known)
        Dim ticks As New ScottPlot.TickGenerators.NumericManual()
        Dim minsPerSample As Double = Chart2MinsPerSample
        Dim labelFormat As String = If(maxX * minsPerSample >= 30, "0", "0.0")

        For k As Integer = 0 To 6
            Dim position As Double = maxX * k / 6.0
            ticks.AddMajor(position, If(minsPerSample > 0, (position * minsPerSample).ToString(labelFormat), CInt(position).ToString()))
        Next

        plot.Axes.Bottom.TickGenerator = ticks

        ' Regional Stats band first so the view rectangle is drawn over it
        Chart2OverviewRegionSpan = plot.Add.HorizontalSpan(0, 1)
        Chart2OverviewRegionSpan.EnableAutoscale = False
        Chart2OverviewRegionSpan.IsVisible = False

        If light Then
            Chart2OverviewRegionSpan.FillColor = New ScottPlot.Color(Color.FromArgb(60, 0, 0, 200))
            Chart2OverviewRegionSpan.LineColor = New ScottPlot.Color(Color.FromArgb(170, 0, 0, 200))
        Else
            Chart2OverviewRegionSpan.FillColor = New ScottPlot.Color(Color.FromArgb(60, 255, 255, 0))
            Chart2OverviewRegionSpan.LineColor = New ScottPlot.Color(Color.FromArgb(170, 255, 255, 0))
        End If

        ' The view rectangle: a span (kept only to hold the X range, drawn as nothing) plus a closed box that just touches the
        ' top and bottom of the plot (a span has no top or bottom edge; a line exactly on the edge of the plot is half clipped)
        Chart2OverviewViewSpan = plot.Add.HorizontalSpan(0, 1)
        Chart2OverviewViewSpan.EnableAutoscale = False
        Chart2OverviewViewSpan.LineWidth = 0
        Chart2OverviewViewSpan.LineColor = ScottPlot.Colors.Transparent
        Chart2OverviewViewSpan.FillColor = ScottPlot.Colors.Transparent

        Dim boxInset As Double = yPad * 0.12     ' about a pixel: the 2px outline shows in full, touching the edge of the plot
        Dim boxLo As Double = yLo - yPad + boxInset
        Dim boxHi As Double = yHi + yPad - boxInset

        Chart2OverviewViewBox = plot.Add.Rectangle(0, 1, boxLo, boxHi)
        Chart2OverviewViewBox.LineWidth = 2

        If light Then
            Chart2OverviewViewBox.FillColor = ScottPlot.Colors.Transparent
            Chart2OverviewViewBox.LineColor = New ScottPlot.Color(Color.FromArgb(230, 0, 140, 0))
        Else
            Chart2OverviewViewBox.FillColor = ScottPlot.Colors.Transparent
            Chart2OverviewViewBox.LineColor = New ScottPlot.Color(Color.FromArgb(240, 50, 255, 50))
        End If

        ' The rectangle's start-end minutes, along the top of the plot; positioned by Chart2SetOverviewView
        Chart2OverviewViewText = plot.Add.Text("", 0, boxHi)
        Chart2OverviewViewText.LabelAlignment = ScottPlot.Alignment.UpperCenter
        Chart2OverviewViewText.LabelFontSize = 11
        Chart2OverviewViewText.LabelBold = True
        Chart2OverviewViewText.OffsetY = 3
        Chart2OverviewViewText.LabelFontColor = If(light, ScottPlot.Colors.Black, ScottPlot.Colors.White)
        Chart2OverviewViewText.LabelBackgroundColor = If(light, New ScottPlot.Color(Color.FromArgb(200, 255, 255, 255)),
                                                             New ScottPlot.Color(Color.FromArgb(200, 0, 0, 0)))
        Chart2OverviewViewText.IsVisible = False

        Chart2OverviewPlot.Refresh()

    End Sub

    ' Called at the end of every main chart render (RenderStarting): rebuilds the overview if the data changed, and
    ' moves the rectangle (and the Regional Stats band) to match the main chart. Does nothing unless the window is open.
    Private Sub Chart2SyncOverview()

        If Chart2OverviewForm Is Nothing OrElse Chart2OverviewForm.IsDisposed OrElse Chart2OverviewPlot Is Nothing Then Exit Sub

        Dim key As String = Chart2OverviewDataKey()

        If key <> Chart2OverviewKey Then
            Chart2OverviewKey = key
            Chart2RebuildOverview()
        End If

        If Chart2OverviewViewSpan Is Nothing Then Exit Sub

        Dim changed As Boolean = False

        ' While the mouse is dragging the rectangle follows the mouse, not the main chart (which lags a little behind)
        If Chart2OverviewDragMode = 0 Then

            Dim x1 As Double = FormsPlot2.Plot.Axes.Bottom.Min
            Dim x2 As Double = FormsPlot2.Plot.Axes.Bottom.Max

            If Not Double.IsNaN(x1) AndAlso Not Double.IsNaN(x2) AndAlso
               Not Double.IsInfinity(x1) AndAlso Not Double.IsInfinity(x2) AndAlso x2 > x1 Then
                If Chart2SetOverviewView(x1, x2) Then changed = True
            End If

        End If

        If Chart2OverviewRegionSpan IsNot Nothing Then

            Dim regionOn As Boolean = Chart2RegionSpan IsNot Nothing AndAlso Chart2RegionSpan.IsVisible
            Dim regionKey As String = If(regionOn, Chart2RegionSpan.Left.ToString("R", CultureInfo.InvariantCulture) & "|" &
                                                   Chart2RegionSpan.Right.ToString("R", CultureInfo.InvariantCulture), "off")

            If regionKey <> Chart2OverviewRegionKey Then
                Chart2OverviewRegionKey = regionKey
                Chart2OverviewRegionSpan.IsVisible = regionOn
                If regionOn Then
                    Chart2OverviewRegionSpan.X1 = Chart2RegionSpan.Left
                    Chart2OverviewRegionSpan.X2 = Chart2RegionSpan.Right
                End If
                changed = True
            End If

        End If

        If changed Then Chart2OverviewPlot.Refresh()

    End Sub

    ' Moves the rectangle; False if it is already there.
    Private Function Chart2SetOverviewView(x1 As Double, x2 As Double) As Boolean

        If Chart2OverviewViewSpan Is Nothing Then Return False
        If x1 = Chart2OverviewViewX1 AndAlso x2 = Chart2OverviewViewX2 Then Return False

        Chart2OverviewViewX1 = x1
        Chart2OverviewViewX2 = x2
        Chart2OverviewViewSpan.X1 = x1
        Chart2OverviewViewSpan.X2 = x2
        If Chart2OverviewViewBox IsNot Nothing Then
            Chart2OverviewViewBox.X1 = x1
            Chart2OverviewViewBox.X2 = x2
        End If
        Chart2UpdateOverviewViewText()

        Return True

    End Function

    ' Labels the rectangle with its start-end minutes (sample numbers if the time scale isn't known), centred on it but kept
    ' inside the plot. Hidden at the full view, where it would only repeat the axis.
    Private Sub Chart2UpdateOverviewViewText()

        If Chart2OverviewViewText Is Nothing OrElse Chart2OverviewPlot Is Nothing Then Exit Sub

        Dim maxX As Double = Chart2OverviewMaxX
        Dim x1 As Double = Math.Max(0, Chart2OverviewViewX1)
        Dim x2 As Double = Math.Min(maxX, Chart2OverviewViewX2)

        If maxX <= 0 OrElse x2 <= x1 OrElse (x1 <= 0.5 AndAlso x2 >= maxX - 0.5) Then
            Chart2OverviewViewText.IsVisible = False
            Exit Sub
        End If

        Dim minsPerSample As Double = Chart2MinsPerSample
        If minsPerSample > 0 Then
            Dim fmt As String = If((x2 - x1) * minsPerSample < 1, "0.00", "0.0")
            Chart2OverviewViewText.LabelText = (x1 * minsPerSample).ToString(fmt) & " - " & (x2 * minsPerSample).ToString(fmt) & " mins"
        Else
            Chart2OverviewViewText.LabelText = CInt(x1).ToString() & " - " & CInt(x2).ToString()
        End If

        ' Keep the whole label inside the plot near either end (width estimated from the character count)
        Dim dataWidthPx As Double = Chart2OverviewPlot.Plot.LastRender.DataRect.Width
        If dataWidthPx <= 0 Then dataWidthPx = Chart2OverviewPlot.Width - 28
        Dim halfLabel As Double = (Chart2OverviewViewText.LabelText.Length * 7.0 / 2.0 + 4) * maxX / Math.Max(1.0, dataWidthPx)
        Dim centre As Double = (x1 + x2) / 2.0
        If halfLabel * 2 < maxX Then centre = Math.Max(halfLabel, Math.Min(maxX - halfLabel, centre))

        Chart2OverviewViewText.Location = New ScottPlot.Coordinates(centre, Chart2OverviewViewText.Location.Y)
        Chart2OverviewViewText.IsVisible = True

    End Sub

    ' Gives the main chart the rectangle's X range. Unticks AutoScale, as any mouse pan or zoom does. Throttled unless forced.
    Private Sub Chart2OverviewApplyToMain(force As Boolean)

        If Double.IsNaN(Chart2OverviewViewX1) OrElse Double.IsNaN(Chart2OverviewViewX2) Then Exit Sub

        If Not force AndAlso Chart2OverviewClock.ElapsedMilliseconds - Chart2OverviewLastApplyMs < ChartOverviewApplyIntervalMs Then Exit Sub
        Chart2OverviewLastApplyMs = Chart2OverviewClock.ElapsedMilliseconds

        CheckBoxPBXYaxis.Checked = False
        FormsPlot2.Plot.Axes.SetLimitsX(Chart2OverviewViewX1, Chart2OverviewViewX2)
        FormsPlot2.Refresh()

    End Sub

    ' Double-click: the full view, the same fit as ticking AutoScale (no re-read of the file).
    Private Sub Chart2OverviewZoomAll()

        CheckBoxPBXYaxis.Checked = True
        Chart2AutoScaleXY()
        FormsPlot2.Refresh()

    End Sub

    Private Sub Chart2OverviewMouseDown(sender As Object, e As MouseEventArgs)

        If e.Button <> MouseButtons.Left Then Exit Sub
        If Chart2OverviewPlot Is Nothing OrElse Chart2OverviewViewSpan Is Nothing OrElse Chart2OverviewMaxX <= 0 Then Exit Sub
        If Double.IsNaN(Chart2OverviewViewX1) OrElse Double.IsNaN(Chart2OverviewViewX2) Then Exit Sub

        Dim thisPixel As New ScottPlot.Pixel(CSng(e.X), CSng(e.Y))

        ' Double-click (timed mouse-downs, as FormsPlot's own double-click event isn't reliable): full view
        Dim elapsedMs As Double = (DateTime.Now - Chart2OverviewLastClickTime).TotalMilliseconds
        Dim dx As Single = thisPixel.X - Chart2OverviewLastClickPixel.X
        Dim dy As Single = thisPixel.Y - Chart2OverviewLastClickPixel.Y

        If elapsedMs <= SystemInformation.DoubleClickTime AndAlso
           Math.Sqrt(dx * dx + dy * dy) <= SystemInformation.DoubleClickSize.Width Then
            Chart2OverviewLastClickTime = DateTime.MinValue
            Chart2OverviewZoomAll()
            Exit Sub
        End If

        Chart2OverviewLastClickTime = DateTime.Now
        Chart2OverviewLastClickPixel = thisPixel

        Dim plot As ScottPlot.Plot = Chart2OverviewPlot.Plot
        Dim leftPixel As Single = plot.GetPixel(New ScottPlot.Coordinates(Chart2OverviewViewX1, 0)).X
        Dim rightPixel As Single = plot.GetPixel(New ScottPlot.Coordinates(Chart2OverviewViewX2, 0)).X
        Dim mouseX As Double = plot.GetCoordinates(thisPixel).X

        ' Grab an edge only as close as a third of the rectangle's width, so a thin rectangle can still be moved
        Dim edgeReach As Single = Math.Min(ChartOverviewEdgePixels, Math.Abs(rightPixel - leftPixel) / 3.0F)

        If Math.Abs(thisPixel.X - leftPixel) <= edgeReach Then

            Chart2OverviewDragMode = 2

        ElseIf Math.Abs(thisPixel.X - rightPixel) <= edgeReach Then

            Chart2OverviewDragMode = 3

        ElseIf thisPixel.X > leftPixel AndAlso thisPixel.X < rightPixel Then

            Chart2OverviewDragMode = 1
            Chart2OverviewDragOffset = mouseX - Chart2OverviewViewX1

        Else

            ' Outside the rectangle: jump so it is centred here, and carry on dragging it from there
            Chart2OverviewDragMode = 1
            Chart2OverviewDragOffset = (Chart2OverviewViewX2 - Chart2OverviewViewX1) / 2.0
            Chart2OverviewDragTo(mouseX)

        End If

    End Sub

    Private Sub Chart2OverviewMouseMove(sender As Object, e As MouseEventArgs)

        If Chart2OverviewPlot Is Nothing OrElse Chart2OverviewViewSpan Is Nothing Then Exit Sub

        Dim plot As ScottPlot.Plot = Chart2OverviewPlot.Plot

        Chart2OverviewShowHover(e.X, e.Y, plot.GetCoordinates(New ScottPlot.Pixel(CSng(e.X), CSng(e.Y))).X)

        If Chart2OverviewDragMode <> 0 Then
            Chart2OverviewDragTo(plot.GetCoordinates(New ScottPlot.Pixel(CSng(e.X), CSng(e.Y))).X)
            Exit Sub
        End If

        If Chart2OverviewMaxX <= 0 OrElse Double.IsNaN(Chart2OverviewViewX1) OrElse Double.IsNaN(Chart2OverviewViewX2) Then
            Chart2OverviewPlot.Cursor = Cursors.Default
            Exit Sub
        End If

        ' Cursor hints: sideways arrows on an edge, four arrows inside the rectangle, a hand elsewhere (click to jump)
        Dim leftPixel As Single = plot.GetPixel(New ScottPlot.Coordinates(Chart2OverviewViewX1, 0)).X
        Dim rightPixel As Single = plot.GetPixel(New ScottPlot.Coordinates(Chart2OverviewViewX2, 0)).X
        Dim edgeReach As Single = Math.Min(ChartOverviewEdgePixels, Math.Abs(rightPixel - leftPixel) / 3.0F)

        If Math.Abs(e.X - leftPixel) <= edgeReach OrElse Math.Abs(e.X - rightPixel) <= edgeReach Then
            Chart2OverviewPlot.Cursor = Cursors.SizeWE
        ElseIf e.X > leftPixel AndAlso e.X < rightPixel Then
            Chart2OverviewPlot.Cursor = Cursors.SizeAll
        Else
            Chart2OverviewPlot.Cursor = Cursors.Hand
        End If

    End Sub

    ' The time under the pointer, in the same units as the axis (minutes; sample numbers if the time scale isn't known).
    ' Placed beside the pointer, never under it, and kept inside the plot.
    Private Sub Chart2OverviewShowHover(mouseX As Integer, mouseY As Integer, xValue As Double)

        Dim lbl As Label = Chart2OverviewHoverLabel

        If lbl Is Nothing OrElse Chart2OverviewPlot Is Nothing Then Exit Sub

        If Chart2OverviewMaxX <= 0 Then
            lbl.Visible = False
            Exit Sub
        End If

        xValue = Math.Max(0, Math.Min(Chart2OverviewMaxX, xValue))

        Dim minsPerSample As Double = Chart2MinsPerSample

        If minsPerSample > 0 Then
            lbl.Text = (xValue * minsPerSample).ToString(If(Chart2OverviewMaxX * minsPerSample >= 30, "0.0", "0.00")) & " mins"
        Else
            lbl.Text = "Sample " & CInt(xValue).ToString()
        End If

        Dim hostSize As Size = Chart2OverviewPlot.ClientSize
        Dim lx As Integer = mouseX + 14
        Dim ly As Integer = mouseY + 20

        If lx + lbl.Width > hostSize.Width Then lx = mouseX - lbl.Width - 14
        If ly + lbl.Height > hostSize.Height Then ly = mouseY - lbl.Height - 10

        lbl.Location = New Point(Math.Max(0, lx), Math.Max(0, ly))
        lbl.Visible = True
        lbl.BringToFront()

        If Chart2OverviewHoverTimer IsNot Nothing Then
            Chart2OverviewHoverTimer.Stop()
            Chart2OverviewHoverTimer.Start()        ' restart the 2 seconds
        End If

    End Sub

    Private Sub Chart2OverviewHideHover()

        If Chart2OverviewHoverTimer IsNot Nothing Then Chart2OverviewHoverTimer.Stop()
        If Chart2OverviewHoverLabel IsNot Nothing Then Chart2OverviewHoverLabel.Visible = False

    End Sub

    Private Sub Chart2OverviewMouseUp(sender As Object, e As MouseEventArgs)

        If e.Button <> MouseButtons.Left OrElse Chart2OverviewDragMode = 0 Then Exit Sub

        Chart2OverviewDragMode = 0
        Chart2OverviewApplyToMain(True)       ' final position, in case the last move was throttled

    End Sub

    ' Moves or resizes the rectangle for the mouse at X (in sample numbers), then passes it to the main chart.
    Private Sub Chart2OverviewDragTo(mouseX As Double)

        Dim maxX As Double = Chart2OverviewMaxX
        Dim minWidth As Double = Math.Max(5.0, maxX * 0.002)      ' not narrower than 5 samples (or 0.2% of the run)

        Dim x1 As Double = Chart2OverviewViewX1
        Dim x2 As Double = Chart2OverviewViewX2

        Select Case Chart2OverviewDragMode

            Case 1
                Dim width As Double = x2 - x1
                x1 = mouseX - Chart2OverviewDragOffset
                If x1 + width > maxX Then x1 = maxX - width
                If x1 < 0 Then x1 = 0
                x2 = x1 + width

            Case 2
                x1 = Math.Max(0, Math.Min(mouseX, x2 - minWidth))

            Case 3
                x2 = Math.Min(maxX, Math.Max(mouseX, x1 + minWidth))

        End Select

        If Chart2SetOverviewView(x1, x2) Then
            Chart2OverviewPlot.Refresh()
            Chart2OverviewApplyToMain(False)
        End If

    End Sub

End Class
