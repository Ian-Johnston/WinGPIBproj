Imports System.Globalization
Imports System.IO
Imports System.Text

' Playback Chart comments: right-click a trace > Add Comment... pins a small note to that point of the trace (a small box with a line
' to the point). Works on the main chart (FormsPlot2) and on the Statistics chart under it (FormsPlot3). Notes are saved beside the
' CSV in <csv name>_comments.txt (comma separated) and loaded again with the CSV.
'   - A comment is tied to a trace (main chart: Data, Mean, Short Term Mean, Temperature, Humidity or PPM Deviation of either device;
'     Statistics chart: STDEV, SEM, Max Diff. or PPM Deviation of either device), the reading's own date/time from the CSV and its
'     reading number, so it stays on the right reading whatever the zoom, Avg setting or reload.
'   - The box keeps a fixed pixel offset from its point (updated every render), and can be dragged; its text/offset are saved.
'   - CheckBoxShowComments (on the form) shows or hides them all.
Partial Public Class Chart

    Private Const PlaybackCommentsEnabled As Boolean = True        ' set False to remove the feature (menu items, drawing, file)

    Private Const ChartCommentMaxLength As Integer = 400
    Private Const ChartCommentWrapChars As Integer = 34
    Private Const ChartCommentHitPixels As Single = 30.0F          ' how close to a trace a right-click must be to add a comment
    Private Const ChartCommentFontSize As Single = 10.0F
    Private Const ChartCommentStatsPrefix As String = "Stats "     ' trace keys of the Statistics chart start with this

    Private Class ChartCommentTarget
        Public Key As String                                    ' main chart: "Data", "Mean", "Short Term Mean", "Temperature", "Humidity", "PPM Deviation"
        '                                                         Statistics chart: "Stats STDEV", "Stats SEM", "Stats Max Diff.", "Stats PPM Deviation"
        Public Slot As Integer                                  ' 1 or 2 for a device's own traces, 0 for Temperature / Humidity / PPM
        Public Host As ScottPlot.WinForms.FormsPlot
        Public Series As ScottPlot.Plottables.Scatter
        Public Data As List(Of ScottPlot.Coordinates)
        Public YAxis As ScottPlot.IYAxis
        Public Column As String                                 ' CSV column suffix (MEAN, STDEV...) when the trace skips rows without a value; "" if every row is plotted
    End Class

    Private Class ChartComment
        Public Device As String = ""                            ' device whose rows give the time of the reading
        Public TraceKey As String = ""
        Public Slot As Integer
        Public SampleIndex As Integer                           ' reading number within the device's rows (saved in the file)
        Public DataIndex As Integer                             ' position in the trace's own data (the same, unless the trace skips rows)
        Public TimeText As String = ""
        Public Text As String = ""
        Public OffsetX As Single = 30                           ' pixels from the point to the box's corner (right/down positive)
        Public OffsetY As Single = -30
        Public Matched As Boolean = True                        ' False if the reading couldn't be found in this CSV
        Public Shown As Boolean = False                         ' drawn in the last render
        Public Box As ScottPlot.Plottables.Text = Nothing
        Public Link As ScottPlot.Plottables.Scatter = Nothing
        Public LinkData As New List(Of ScottPlot.Coordinates)
        Public Dot As ScottPlot.Plottables.Marker = Nothing
    End Class

    Private Chart2Comments As New List(Of ChartComment)
    Private Chart2CommentsLoadedFor As String = ""
    Private Chart2CommentDrag As ChartComment = Nothing
    Private Chart2CommentDragHost As ScottPlot.WinForms.FormsPlot = Nothing
    Private Chart2CommentDragStart As ScottPlot.Pixel
    Private Chart2CommentDragOffsetX As Single
    Private Chart2CommentDragOffsetY As Single
    Private Chart2CommentCursorSet As Boolean = False
    Private Chart2CommentMenuAdd As ToolStripMenuItem = Nothing
    Private Chart2CommentMenuEdit As ToolStripMenuItem = Nothing
    Private Chart2CommentMenuDelete As ToolStripMenuItem = Nothing
    Private Chart2CommentMenuTarget As ChartComment = Nothing
    Private Chart3ContextMenu As ContextMenuStrip = Nothing
    Private Chart3CommentMenuAdd As ToolStripMenuItem = Nothing
    Private Chart3CommentMenuEdit As ToolStripMenuItem = Nothing
    Private Chart3CommentMenuDelete As ToolStripMenuItem = Nothing
    Private Chart3CommentMenuTarget As ChartComment = Nothing
    Private Chart3LastRightClickPixel As ScottPlot.Pixel
    Private Chart2CommentTimeCache As New Dictionary(Of String, String())
    Private Chart2CommentRowMapCache As New Dictionary(Of String, Integer())
    Private Chart2CommentTimeCacheKey As String = ""

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Traces that can carry a comment, and the time of each reading
    ' ---------------------------------------------------------------------------------------------------------------------------

    ' The traces of one chart (host = FormsPlot2 or FormsPlot3).
    Private Function Chart2CommentTargets(host As ScottPlot.WinForms.FormsPlot) As List(Of ChartCommentTarget)

        Dim targets As New List(Of ChartCommentTarget)
        If host Is Nothing Then Return targets

        Dim addTarget =
            Sub(key As String, slot As Integer, series As ScottPlot.Plottables.Scatter, data As List(Of ScottPlot.Coordinates), axis As ScottPlot.IYAxis, column As String)
                If series IsNot Nothing AndAlso data IsNot Nothing AndAlso axis IsNot Nothing Then
                    targets.Add(New ChartCommentTarget With {.Key = key, .Slot = slot, .Host = host, .Series = series, .Data = data, .YAxis = axis, .Column = column})
                End If
            End Sub

        If host Is FormsPlot2 Then

            Dim leftAxis As ScottPlot.IYAxis = FormsPlot2.Plot.Axes.Left
            addTarget("Data", 1, Chart2Dev1Series, Chart2Dev1Data, leftAxis, "")
            addTarget("Data", 2, Chart2Dev2Series, Chart2Dev2Data, leftAxis, "")
            addTarget("Mean", 1, Chart2Dev1MeanSeries, Chart2Dev1MeanData, leftAxis, "MEAN")
            addTarget("Mean", 2, Chart2Dev2MeanSeries, Chart2Dev2MeanData, leftAxis, "MEAN")
            addTarget("Short Term Mean", 1, Chart2Dev1ShortTermMeanSeries, Chart2Dev1ShortTermMeanData, leftAxis, "")
            addTarget("Short Term Mean", 2, Chart2Dev2ShortTermMeanSeries, Chart2Dev2ShortTermMeanData, leftAxis, "")
            addTarget("Temperature", 0, Chart2TempSeries, Chart2TempData, Chart2TempAxis, "")
            addTarget("Humidity", 0, Chart2HumSeries, Chart2HumData, Chart2HumAxis, "")
            addTarget("PPM Deviation", 0, Chart2PPMSeries, Chart2PPMData, Chart2PPMAxis, "")

        ElseIf host Is FormsPlot3 Then

            Dim leftAxis As ScottPlot.IYAxis = FormsPlot3.Plot.Axes.Left
            addTarget(ChartCommentStatsPrefix & "STDEV", 1, Chart3Dev1StdevSeries, Chart2Dev1StdevData, Chart3StdevAxis, "STDEV")
            addTarget(ChartCommentStatsPrefix & "STDEV", 2, Chart3Dev2StdevSeries, Chart2Dev2StdevData, Chart3StdevAxis, "STDEV")
            addTarget(ChartCommentStatsPrefix & "SEM", 1, Chart3Dev1SEMSeries, Chart2Dev1SEMData, Chart3SEMAxis, "SEM")
            addTarget(ChartCommentStatsPrefix & "SEM", 2, Chart3Dev2SEMSeries, Chart2Dev2SEMData, Chart3SEMAxis, "SEM")
            addTarget(ChartCommentStatsPrefix & "Max Diff.", 1, Chart3Dev1MaxDiffSeries, Chart2Dev1MaxDiffData, leftAxis, "MAXDIFF")
            addTarget(ChartCommentStatsPrefix & "Max Diff.", 2, Chart3Dev2MaxDiffSeries, Chart2Dev2MaxDiffData, leftAxis, "MAXDIFF")
            addTarget(ChartCommentStatsPrefix & "PPM Deviation", 1, Chart3Dev1DeviationSeries, Chart2Dev1DeviationData, Chart3PPMAxis, "DEVIATION")
            addTarget(ChartCommentStatsPrefix & "PPM Deviation", 2, Chart3Dev2DeviationSeries, Chart2Dev2DeviationData, Chart3PPMAxis, "DEVIATION")

        End If

        Return targets

    End Function

    Private Function Chart2CommentIsStats(key As String) As Boolean
        Return key.StartsWith(ChartCommentStatsPrefix, StringComparison.Ordinal)
    End Function

    ' The chart a comment lives on.
    Private Function Chart2CommentHost(key As String) As ScottPlot.WinForms.FormsPlot
        Return If(Chart2CommentIsStats(key), FormsPlot3, FormsPlot2)
    End Function

    Private Function Chart2CommentTraceTitle(key As String, slot As Integer) As String

        If Chart2CommentIsStats(key) Then Return "Dev " & slot.ToString() & " " & key.Substring(ChartCommentStatsPrefix.Length) & " (statistics chart)"
        If slot > 0 Then Return "Dev " & slot.ToString() & " " & key
        Return key

    End Function

    ' The device whose rows give a trace its time stamps (Temperature / Humidity are shared; PPM follows the PPM device).
    Private Function Chart2CommentDeviceName(key As String, slot As Integer) As String

        Select Case key
            Case "Temperature", "Humidity"
                Return If(DeviceName1.Text <> "", DeviceName1.Text, DeviceName2.Text)
            Case "PPM Deviation"
                Return If(RadioButtonDev2.Checked, DeviceName2.Text, DeviceName1.Text)
            Case Else
                Return If(slot = 1, DeviceName1.Text, DeviceName2.Text)
        End Select

    End Function

    Private Sub Chart2CommentCheckCache()

        Dim cacheKey As String = filePlayback & "|" & dataTable1.Rows.Count.ToString()
        If cacheKey <> Chart2CommentTimeCacheKey Then
            Chart2CommentTimeCache.Clear()
            Chart2CommentRowMapCache.Clear()
            Chart2CommentTimeCacheKey = cacheKey
        End If

    End Sub

    ' DATETIME text of every reading of a device, in order (cached until the loaded data changes).
    Private Function Chart2CommentTimes(deviceName As String) As String()

        If deviceName = "" Then Return Nothing

        Chart2CommentCheckCache()

        Dim times() As String = Nothing
        If Chart2CommentTimeCache.TryGetValue(deviceName, times) Then Return times

        Dim rows() As DataRow = dataTable1.Select("DEVICE ='" & deviceName.Replace("'", "''") & "'")
        times = New String(rows.Length - 1) {}
        For i As Integer = 0 To rows.Length - 1
            times(i) = Convert.ToString(rows(i)("DATETIME"))
        Next

        Chart2CommentTimeCache(deviceName) = times
        Return times

    End Function

    ' For the traces that skip rows without a value (Mean and the statistics): the reading number of each point of the trace.
    ' Nothing for a trace that plots every row (its points and readings are numbered the same).
    Private Function Chart2CommentRowMap(deviceName As String, slot As Integer, column As String) As Integer()

        If column = "" OrElse deviceName = "" Then Return Nothing

        Chart2CommentCheckCache()

        Dim mapKey As String = deviceName & "|" & column & "|" & slot.ToString()
        Dim map() As Integer = Nothing
        If Chart2CommentRowMapCache.TryGetValue(mapKey, map) Then Return map

        Dim fullName As String = "DEV" & slot.ToString() & "_" & column
        Dim rows() As DataRow = dataTable1.Select("DEVICE ='" & deviceName.Replace("'", "''") & "'")
        Dim list As New List(Of Integer)

        If dataTable1.Columns.Contains(fullName) Then
            For i As Integer = 0 To rows.Length - 1
                Dim s As String = rows(i)(fullName).ToString().Trim()
                If s <> "" AndAlso s.ToLower() <> "nil" Then list.Add(i)
            Next
        Else
            For i As Integer = 0 To rows.Length - 1
                list.Add(i)
            Next
        End If

        map = list.ToArray()
        Chart2CommentRowMapCache(mapKey) = map
        Return map

    End Function

    Private Function Chart2CommentRowFromData(deviceName As String, slot As Integer, column As String, dataIndex As Integer) As Integer

        Dim map() As Integer = Chart2CommentRowMap(deviceName, slot, column)
        If map Is Nothing Then Return dataIndex
        If dataIndex < 0 OrElse dataIndex >= map.Length Then Return -1
        Return map(dataIndex)

    End Function

    ' The point of the trace for a reading number (-1 if that reading has no point on this trace).
    Private Function Chart2CommentDataFromRow(deviceName As String, slot As Integer, column As String, rowIndex As Integer) As Integer

        Dim map() As Integer = Chart2CommentRowMap(deviceName, slot, column)
        If map Is Nothing Then Return rowIndex
        Dim position As Integer = Array.BinarySearch(map, rowIndex)
        Return If(position >= 0, position, -1)

    End Function

    Private Function Chart2CommentWhen(index As Integer, timeText As String) As String

        Dim stamp As String = timeText.Replace("_", " ")
        If Chart2MinsPerSample > 0 Then Return RpDur(index * Chart2MinsPerSample) & " into the run (" & stamp & ")"
        Return "reading " & index.ToString() & " (" & stamp & ")"

    End Function

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Finding the point a click is on
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Function Chart2CommentFindNearest(host As ScottPlot.WinForms.FormsPlot, pixel As ScottPlot.Pixel, ByRef target As ChartCommentTarget,
                                              ByRef dataIndex As Integer, ByRef coord As ScottPlot.Coordinates) As Boolean

        Dim found As Boolean = False
        Dim bestDistance As Single = Single.MaxValue
        Dim bottomAxis As ScottPlot.IXAxis = host.Plot.Axes.Bottom

        For Each t As ChartCommentTarget In Chart2CommentTargets(host)

            If Not t.Series.IsVisible OrElse t.Data.Count = 0 Then Continue For

            Dim mouseLocation As ScottPlot.Coordinates = host.Plot.GetCoordinates(pixel, bottomAxis, t.YAxis)
            Dim dataSource As ScottPlot.IDataSource = DirectCast(t.Series.Data, ScottPlot.IDataSource)
            Dim point As ScottPlot.DataPoint = ScottPlot.DataSourceUtilities.GetNearestSmart(
                dataSource, mouseLocation, host.Plot.LastRender, ChartCommentHitPixels, bottomAxis, t.YAxis)
            If Not point.IsReal Then Continue For

            Dim pointPixel As ScottPlot.Pixel = host.Plot.GetPixel(point.Coordinates, bottomAxis, t.YAxis)
            Dim distance As Single = pointPixel.DistanceFrom(pixel)

            If distance <= ChartCommentHitPixels AndAlso distance < bestDistance Then
                found = True
                bestDistance = distance
                target = t
                coord = point.Coordinates
                dataIndex = Math.Max(0, Math.Min(t.Data.Count - 1, CInt(Math.Round(point.Coordinates.X))))
            End If

        Next

        Return found

    End Function

    ' The comment (on that chart, drawn in the last render) whose box is under the pixel.
    Private Function Chart2CommentAt(host As ScottPlot.WinForms.FormsPlot, pixel As ScottPlot.Pixel) As ChartComment

        For i As Integer = Chart2Comments.Count - 1 To 0 Step -1
            Dim c As ChartComment = Chart2Comments(i)
            If Not c.Shown OrElse c.Box Is Nothing OrElse Not Object.ReferenceEquals(Chart2CommentHost(c.TraceKey), host) Then Continue For
            Dim r As ScottPlot.PixelRect = c.Box.LabelLastRenderPixelRect
            If pixel.X >= r.Left AndAlso pixel.X <= r.Right AndAlso pixel.Y >= r.Top AndAlso pixel.Y <= r.Bottom Then Return c
        Next

        Return Nothing

    End Function

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Right-click menus (main chart: items added to Chart2ContextMenu; Statistics chart: its own menu)
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Sub Chart2BuildCommentMenuItems()

        Chart2CommentMenuAdd = New ToolStripMenuItem("Add Comment...")
        Chart2CommentMenuEdit = New ToolStripMenuItem("Edit Comment...")
        Chart2CommentMenuDelete = New ToolStripMenuItem("Delete Comment")

        AddHandler Chart2CommentMenuAdd.Click, Sub(s, ev) Chart2AddComment(FormsPlot2, Chart2LastRightClickPixel)
        AddHandler Chart2CommentMenuEdit.Click, Sub(s, ev) Chart2EditComment(Chart2CommentMenuTarget)
        AddHandler Chart2CommentMenuDelete.Click, Sub(s, ev) Chart2DeleteComment(Chart2CommentMenuTarget)

        Chart2ContextMenu.Items.Add(Chart2CommentMenuAdd)
        Chart2ContextMenu.Items.Add(Chart2CommentMenuEdit)
        Chart2ContextMenu.Items.Add(Chart2CommentMenuDelete)

    End Sub

    ' Sets the Add item's text for the trace under the click (or disables it), and shows Edit / Delete over an existing comment.
    Private Sub Chart2RefreshCommentItems(host As ScottPlot.WinForms.FormsPlot, pixel As ScottPlot.Pixel, addItem As ToolStripMenuItem, editItem As ToolStripMenuItem,
                                          deleteItem As ToolStripMenuItem, ByRef hit As ChartComment)

        hit = Chart2CommentAt(host, pixel)
        editItem.Visible = hit IsNot Nothing
        deleteItem.Visible = hit IsNot Nothing

        Dim target As ChartCommentTarget = Nothing
        Dim dataIndex As Integer = 0
        Dim coord As ScottPlot.Coordinates = Nothing

        If ChartLoaded AndAlso CSVfileok AndAlso Chart2CommentFindNearest(host, pixel, target, dataIndex, coord) Then
            Dim rowIndex As Integer = Chart2CommentRowFromData(Chart2CommentDeviceName(target.Key, target.Slot), target.Slot, target.Column, dataIndex)
            Dim whenText As String = ""
            If rowIndex >= 0 AndAlso Chart2MinsPerSample > 0 Then whenText = " at " & RpDur(rowIndex * Chart2MinsPerSample)
            addItem.Text = "Add Comment on " & Chart2CommentTraceTitle(target.Key, target.Slot) & whenText & "..."
            addItem.Enabled = True
        Else
            addItem.Text = "Add Comment... (right-click on a trace)"
            addItem.Enabled = False
        End If

    End Sub

    Private Sub Chart2RefreshCommentMenu()

        If Chart2CommentMenuAdd Is Nothing Then Exit Sub
        Chart2RefreshCommentItems(FormsPlot2, Chart2LastRightClickPixel, Chart2CommentMenuAdd, Chart2CommentMenuEdit, Chart2CommentMenuDelete, Chart2CommentMenuTarget)

    End Sub

    ' The Statistics chart's own right-click menu: Save Image, plus the comment items. (Its input handling is switched off, so
    ' ScottPlot's built-in menu is removed and this one is shown from MouseUp, as on the main chart.)
    Private Sub Chart3BuildContextMenu()

        If FormsPlot3 Is Nothing Then Exit Sub

        Dim formsMenu As ScottPlot.WinForms.FormsPlotMenu = DirectCast(FormsPlot3.Menu, ScottPlot.WinForms.FormsPlotMenu)
        formsMenu.Clear()
        FormsPlot3.UserInputProcessor.RemoveAll(Of ScottPlot.Interactivity.UserActionResponses.SingleClickContextMenu)()

        Chart3ContextMenu = New ContextMenuStrip()

        Dim saveItem As New ToolStripMenuItem("Save Image")
        AddHandler saveItem.Click, Sub(s, ev) formsMenu.OpenSaveImageDialog(FormsPlot3.Plot)
        Chart3ContextMenu.Items.Add(saveItem)

        If PlaybackCommentsEnabled Then
            Chart3CommentMenuAdd = New ToolStripMenuItem("Add Comment...")
            Chart3CommentMenuEdit = New ToolStripMenuItem("Edit Comment...")
            Chart3CommentMenuDelete = New ToolStripMenuItem("Delete Comment")
            AddHandler Chart3CommentMenuAdd.Click, Sub(s, ev) Chart2AddComment(FormsPlot3, Chart3LastRightClickPixel)
            AddHandler Chart3CommentMenuEdit.Click, Sub(s, ev) Chart2EditComment(Chart3CommentMenuTarget)
            AddHandler Chart3CommentMenuDelete.Click, Sub(s, ev) Chart2DeleteComment(Chart3CommentMenuTarget)
            Chart3ContextMenu.Items.Add(Chart3CommentMenuAdd)
            Chart3ContextMenu.Items.Add(Chart3CommentMenuEdit)
            Chart3ContextMenu.Items.Add(Chart3CommentMenuDelete)

            AddHandler Chart3ContextMenu.Opening,
                Sub(s, ev) Chart2RefreshCommentItems(FormsPlot3, Chart3LastRightClickPixel, Chart3CommentMenuAdd, Chart3CommentMenuEdit, Chart3CommentMenuDelete, Chart3CommentMenuTarget)
        End If

        AddHandler FormsPlot3.MouseDown, AddressOf Chart3CommentMouseDown
        AddHandler FormsPlot3.MouseMove, AddressOf Chart3CommentMouseMove
        AddHandler FormsPlot3.MouseUp, AddressOf Chart3CommentMouseUp

        ' Comments on the Statistics chart are placed during its own renders
        FormsPlot3.Plot.RenderManager.RenderStarting = New EventHandler(Of ScottPlot.RenderPack)(AddressOf Chart3RenderStarting)

    End Sub

    Private Sub Chart3RenderStarting(sender As Object, rp As ScottPlot.RenderPack)

        Chart2UpdateCommentsOn(FormsPlot3, rp)

    End Sub

    Private Sub Chart3CommentMouseDown(sender As Object, e As MouseEventArgs)

        If e.Button = MouseButtons.Right Then
            Chart3LastRightClickPixel = New ScottPlot.Pixel(CSng(e.X), CSng(e.Y))
        ElseIf e.Button = MouseButtons.Left Then
            Chart2BeginCommentDragOn(FormsPlot3, New ScottPlot.Pixel(CSng(e.X), CSng(e.Y)))
        End If

    End Sub

    Private Sub Chart3CommentMouseMove(sender As Object, e As MouseEventArgs)

        Chart2CommentMouseMoveOn(FormsPlot3, e)

    End Sub

    Private Sub Chart3CommentMouseUp(sender As Object, e As MouseEventArgs)

        Chart2EndCommentDrag()

        If e.Button = MouseButtons.Right AndAlso Chart3ContextMenu IsNot Nothing Then
            Dim dragDistance As Single = Math.Max(Math.Abs(e.X - Chart3LastRightClickPixel.X), Math.Abs(e.Y - Chart3LastRightClickPixel.Y))
            If dragDistance < 5 Then
                Dim menuPos As New Point(e.X, e.Y)
                Me.BeginInvoke(New MethodInvoker(Sub() Chart3ContextMenu.Show(FormsPlot3, menuPos)))      ' deferred, as on the main chart
            End If
        End If

    End Sub

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Add / edit / delete
    ' ---------------------------------------------------------------------------------------------------------------------------

    ' 'detail' says which trace and time the comment is for; it goes in the window (the title bar would cut it off).
    Private Function Chart2PromptComment(title As String, detail As String, ByRef text As String) As Boolean

        Using dlg As New Form With {
            .Text = title,
            .FormBorderStyle = FormBorderStyle.FixedDialog,
            .StartPosition = FormStartPosition.CenterParent,
            .MaximizeBox = False,
            .MinimizeBox = False,
            .ShowInTaskbar = False,
            .ClientSize = New Size(480, 222)
        }

            dlg.Controls.Add(New Label With {.Text = detail, .AutoSize = False, .Location = New Point(12, 12), .Size = New Size(456, 36)})

            Dim box As New TextBox With {
                .Multiline = True,
                .AcceptsReturn = True,
                .ScrollBars = ScrollBars.Vertical,
                .MaxLength = ChartCommentMaxLength,
                .Location = New Point(12, 52),
                .Size = New Size(456, 100),
                .Text = text
            }
            dlg.Controls.Add(box)
            dlg.Controls.Add(New Label With {.Text = "Up to " & ChartCommentMaxLength.ToString() & " characters. Saved beside the CSV when you press OK.", .AutoSize = True, .Location = New Point(12, 158)})

            Dim ok As New Button With {.Text = "OK", .DialogResult = DialogResult.OK, .Size = New Size(80, 28), .Location = New Point(298, 184)}
            Dim cancel As New Button With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Size = New Size(80, 28), .Location = New Point(388, 184)}
            dlg.Controls.Add(ok)
            dlg.Controls.Add(cancel)
            dlg.CancelButton = cancel

            AddHandler dlg.Shown, Sub(s, ev)
                                      box.Focus()
                                      box.SelectionStart = box.TextLength
                                  End Sub

            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return False

            Dim entered As String = box.Text.Trim()
            If entered = "" Then Return False

            text = entered
            Return True

        End Using

    End Function

    Private Sub Chart2AddComment(host As ScottPlot.WinForms.FormsPlot, pixel As ScottPlot.Pixel)

        If Not (ChartLoaded AndAlso CSVfileok) Then Exit Sub

        Dim target As ChartCommentTarget = Nothing
        Dim dataIndex As Integer = 0
        Dim coord As ScottPlot.Coordinates = Nothing

        If Not Chart2CommentFindNearest(host, pixel, target, dataIndex, coord) Then
            MessageBox.Show("Right-click on a trace (within a few pixels of it) to add a comment there.", "Add Comment", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim deviceName As String = Chart2CommentDeviceName(target.Key, target.Slot)
        Dim times() As String = Chart2CommentTimes(deviceName)
        Dim rowIndex As Integer = Chart2CommentRowFromData(deviceName, target.Slot, target.Column, dataIndex)

        If times Is Nothing OrElse rowIndex < 0 OrElse rowIndex >= times.Length Then
            MessageBox.Show("That reading could not be found in the CSV data, so a comment cannot be attached to it.", "Add Comment", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim text As String = ""
        If Not Chart2PromptComment("Add Comment", Chart2CommentTraceTitle(target.Key, target.Slot) & ", " & Chart2CommentWhen(rowIndex, times(rowIndex)), text) Then Exit Sub

        ' Starting offset of the box from the point: up and to the right, flipped where that would run off the plot
        Dim tipPixel As ScottPlot.Pixel = host.Plot.GetPixel(coord, host.Plot.Axes.Bottom, target.YAxis)
        Dim dataRect As ScottPlot.PixelRect = host.Plot.LastRender.DataRect

        Dim c As New ChartComment With {
            .Device = deviceName,
            .TraceKey = target.Key,
            .Slot = target.Slot,
            .SampleIndex = rowIndex,
            .DataIndex = dataIndex,
            .TimeText = times(rowIndex),
            .Text = text,
            .OffsetX = If(dataRect.Right - tipPixel.X < 190, -30.0F, 30.0F),
            .OffsetY = If(tipPixel.Y - dataRect.Top < 90, 30.0F, -30.0F)
        }

        Chart2Comments.Add(c)
        Chart2EnsureCommentPlottables(c)

        If Not CheckBoxShowComments.Checked Then CheckBoxShowComments.Checked = True       ' so the new comment is seen
        Chart2SaveComments()
        Chart2RefreshCommentCharts()

    End Sub

    Private Sub Chart2EditComment(c As ChartComment)

        If c Is Nothing Then Exit Sub

        Dim text As String = c.Text
        If Not Chart2PromptComment("Edit Comment", Chart2CommentTraceTitle(c.TraceKey, c.Slot) & ", " & Chart2CommentWhen(c.SampleIndex, c.TimeText), text) Then Exit Sub

        c.Text = text
        If c.Box IsNot Nothing Then c.Box.LabelText = Chart2WrapComment(text)
        Chart2SaveComments()
        Chart2RefreshCommentCharts()

    End Sub

    Private Sub Chart2DeleteComment(c As ChartComment)

        If c Is Nothing Then Exit Sub

        If MessageBox.Show("Delete this comment?" & vbCrLf & vbCrLf & c.Text, "Delete Comment", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Exit Sub

        Chart2RemoveCommentPlottables(c)
        Chart2Comments.Remove(c)
        Chart2SaveComments()
        Chart2RefreshCommentCharts()

    End Sub

    Private Sub Chart2RefreshCommentCharts()

        If FormsPlot2 IsNot Nothing AndAlso FormsPlot2.Visible Then FormsPlot2.Refresh()
        If FormsPlot3 IsNot Nothing AndAlso FormsPlot3.Visible Then FormsPlot3.Refresh()

    End Sub

    Private Sub CheckBoxShowComments_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxShowComments.CheckedChanged

        Chart2RefreshCommentCharts()

    End Sub

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Dragging a box
    ' ---------------------------------------------------------------------------------------------------------------------------

    ' Left mouse-down on a comment box: drags the box (on the main chart so the chart does not pan). True if a drag has started.
    Private Function Chart2BeginCommentDragOn(host As ScottPlot.WinForms.FormsPlot, pixel As ScottPlot.Pixel) As Boolean

        If Not PlaybackCommentsEnabled Then Return False

        Dim c As ChartComment = Chart2CommentAt(host, pixel)
        If c Is Nothing Then Return False

        Chart2CommentDrag = c
        Chart2CommentDragHost = host
        Chart2CommentDragStart = pixel
        Chart2CommentDragOffsetX = c.OffsetX
        Chart2CommentDragOffsetY = c.OffsetY
        If host Is FormsPlot2 Then FormsPlot2.UserInputProcessor.Disable()         ' same approach as dragging the Regional Stats band
        Return True

    End Function

    Private Function Chart2BeginCommentDrag(pixel As ScottPlot.Pixel) As Boolean
        Return Chart2BeginCommentDragOn(FormsPlot2, pixel)
    End Function

    ' Mouse move: moves a box being dragged (True = handled) or shows the move cursor over a box.
    Private Function Chart2CommentMouseMoveOn(host As ScottPlot.WinForms.FormsPlot, e As MouseEventArgs) As Boolean

        If Not PlaybackCommentsEnabled OrElse Chart2Comments.Count = 0 Then Return False

        If Chart2CommentDrag IsNot Nothing Then
            If Not Object.ReferenceEquals(Chart2CommentDragHost, host) Then Return False
            Chart2CommentDrag.OffsetX = Chart2CommentDragOffsetX + (e.X - Chart2CommentDragStart.X)
            Chart2CommentDrag.OffsetY = Chart2CommentDragOffsetY + (e.Y - Chart2CommentDragStart.Y)
            host.Refresh()
            Return True
        End If

        Dim over As Boolean = Chart2CommentAt(host, New ScottPlot.Pixel(CSng(e.X), CSng(e.Y))) IsNot Nothing
        If over Then
            host.Cursor = Cursors.SizeAll
            Chart2CommentCursorSet = True
        ElseIf Chart2CommentCursorSet Then
            host.Cursor = Cursors.Default
            Chart2CommentCursorSet = False
        End If

        Return False

    End Function

    Private Function Chart2CommentMouseMove(e As MouseEventArgs) As Boolean
        Return Chart2CommentMouseMoveOn(FormsPlot2, e)
    End Function

    Private Sub Chart2EndCommentDrag()

        If Chart2CommentDrag Is Nothing Then Exit Sub

        Dim host As ScottPlot.WinForms.FormsPlot = Chart2CommentDragHost

        Chart2CommentDrag = Nothing
        Chart2CommentDragHost = Nothing
        If host Is FormsPlot2 Then FormsPlot2.UserInputProcessor.Enable()
        Chart2SaveComments()
        If host IsNot Nothing Then host.Refresh()

    End Sub

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Drawing (every render)
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Function Chart2CommentToPixel(host As ScottPlot.WinForms.FormsPlot, c As ScottPlot.Coordinates, yAxis As ScottPlot.IYAxis, dr As ScottPlot.PixelRect) As ScottPlot.Pixel

        Dim xa As ScottPlot.IXAxis = host.Plot.Axes.Bottom
        Dim xSpan As Double = xa.Max - xa.Min
        Dim ySpan As Double = yAxis.Max - yAxis.Min
        If xSpan = 0 OrElse ySpan = 0 Then Return New ScottPlot.Pixel(Single.NaN, Single.NaN)

        Dim fx As Double = (c.X - xa.Min) / xSpan
        Dim fy As Double = (c.Y - yAxis.Min) / ySpan
        Return New ScottPlot.Pixel(CSng(dr.Left + fx * dr.Width), CSng(dr.Bottom - fy * dr.Height))

    End Function

    Private Function Chart2CommentToCoordinates(host As ScottPlot.WinForms.FormsPlot, x As Single, y As Single, yAxis As ScottPlot.IYAxis, dr As ScottPlot.PixelRect) As ScottPlot.Coordinates

        Dim xa As ScottPlot.IXAxis = host.Plot.Axes.Bottom
        Dim fx As Double = (x - dr.Left) / dr.Width
        Dim fy As Double = (dr.Bottom - y) / dr.Height
        Return New ScottPlot.Coordinates(xa.Min + fx * (xa.Max - xa.Min), yAxis.Min + fy * (yAxis.Max - yAxis.Min))

    End Function

    Private Function Chart2WrapComment(text As String) As String

        Dim lines As New List(Of String)

        For Each paragraph As String In text.Replace(vbCr, "").Split(vbLf(0))
            Dim current As String = ""
            For Each word As String In paragraph.Split(" "c)
                If current = "" Then
                    current = word
                ElseIf current.Length + 1 + word.Length <= ChartCommentWrapChars Then
                    current &= " " & word
                Else
                    lines.Add(current)
                    current = word
                End If
            Next
            lines.Add(current)
        Next

        Return String.Join(vbLf, lines)

    End Function

    ' Made when a comment is added or loaded - never during a render (items added mid-render had no axes yet and the paint failed).
    Private Sub Chart2EnsureCommentPlottables(c As ChartComment)

        If c.Box IsNot Nothing Then Exit Sub

        Dim host As ScottPlot.WinForms.FormsPlot = Chart2CommentHost(c.TraceKey)
        If host Is Nothing Then Exit Sub

        Dim plot As ScottPlot.Plot = host.Plot

        c.LinkData.Clear()
        c.LinkData.Add(New ScottPlot.Coordinates(0, 0))
        c.LinkData.Add(New ScottPlot.Coordinates(0, 0))

        c.Link = plot.Add.Scatter(c.LinkData)
        c.Link.Axes.XAxis = plot.Axes.Bottom
        c.Link.Axes.YAxis = plot.Axes.Left
        c.Link.LineWidth = 1
        c.Link.MarkerStyle.IsVisible = False
        c.Link.IsVisible = False

        c.Dot = plot.Add.Marker(0, 0)
        c.Dot.Axes.XAxis = plot.Axes.Bottom
        c.Dot.Axes.YAxis = plot.Axes.Left
        c.Dot.Shape = ScottPlot.MarkerShape.FilledCircle
        c.Dot.Size = 6
        c.Dot.IsVisible = False

        c.Box = plot.Add.Text(Chart2WrapComment(c.Text), 0, 0)
        c.Box.Axes.XAxis = plot.Axes.Bottom
        c.Box.Axes.YAxis = plot.Axes.Left
        c.Box.LabelFontSize = ChartCommentFontSize
        c.Box.LabelBorderWidth = 1
        c.Box.LabelBorderRadius = 3
        c.Box.LabelPadding = 3
        c.Box.IsVisible = False

    End Sub

    Private Sub Chart2RemoveCommentPlottables(c As ChartComment)

        Dim host As ScottPlot.WinForms.FormsPlot = Chart2CommentHost(c.TraceKey)

        If host IsNot Nothing Then
            If c.Box IsNot Nothing Then host.Plot.Remove(c.Box)
            If c.Link IsNot Nothing Then host.Plot.Remove(c.Link)
            If c.Dot IsNot Nothing Then host.Plot.Remove(c.Dot)
        End If

        c.Box = Nothing
        c.Link = Nothing
        c.Dot = Nothing

    End Sub

    ' Called at the end of every render of the main chart.
    Private Sub Chart2UpdateComments(rp As ScottPlot.RenderPack)

        Chart2UpdateCommentsOn(FormsPlot2, rp)

    End Sub

    ' Places each of that chart's comment boxes at its fixed pixel offset from its point.
    Private Sub Chart2UpdateCommentsOn(host As ScottPlot.WinForms.FormsPlot, rp As ScottPlot.RenderPack)

        If Not PlaybackCommentsEnabled OrElse Chart2Comments.Count = 0 OrElse host Is Nothing Then Exit Sub

        Dim showAll As Boolean = ChartLoaded AndAlso CSVfileok AndAlso CheckBoxShowComments.Checked
        Dim light As Boolean = CheckBoxColours.Checked
        Dim dr As ScottPlot.PixelRect = rp.DataRect
        Dim targets As List(Of ChartCommentTarget) = Chart2CommentTargets(host)

        Dim textColour As ScottPlot.Color = If(light, ScottPlot.Colors.Black, ScottPlot.Colors.White)
        Dim fillColour As ScottPlot.Color = If(light, New ScottPlot.Color(Color.FromArgb(235, 255, 255, 225)), New ScottPlot.Color(Color.FromArgb(235, 40, 40, 40)))

        For Each c As ChartComment In Chart2Comments

            If Not Object.ReferenceEquals(Chart2CommentHost(c.TraceKey), host) Then Continue For

            c.Shown = False
            If c.Box Is Nothing OrElse c.Link Is Nothing OrElse c.Dot Is Nothing Then Continue For     ' made when the comment is added / loaded, never during a render

            Dim target As ChartCommentTarget = Nothing
            If showAll AndAlso c.Matched Then
                For Each t As ChartCommentTarget In targets
                    If t.Key = c.TraceKey AndAlso t.Slot = c.Slot Then
                        target = t
                        Exit For
                    End If
                Next
            End If

            Dim visible As Boolean = False
            Dim tip As ScottPlot.Coordinates = Nothing
            Dim tipPixel As ScottPlot.Pixel = Nothing

            If target IsNot Nothing AndAlso target.Series.IsVisible AndAlso dr.HasArea AndAlso c.DataIndex >= 0 AndAlso c.DataIndex < target.Data.Count Then
                tip = target.Data(c.DataIndex)
                tipPixel = Chart2CommentToPixel(host, tip, target.YAxis, dr)
                visible = Not Single.IsNaN(tipPixel.X) AndAlso tipPixel.X >= dr.Left - 1 AndAlso tipPixel.X <= dr.Right + 1 AndAlso
                          tipPixel.Y >= dr.Top - 1 AndAlso tipPixel.Y <= dr.Bottom + 1
            End If

            c.Box.IsVisible = visible
            c.Link.IsVisible = visible
            c.Dot.IsVisible = visible

            If Not visible Then Continue For

            c.Shown = True

            Dim seriesColour As ScottPlot.Color = target.Series.LineStyle.Color

            c.Box.Axes.YAxis = target.YAxis
            c.Link.Axes.YAxis = target.YAxis
            c.Dot.Axes.YAxis = target.YAxis

            ' Where the box goes: its corner nearest the point sits at the fixed offset, then the whole box is kept inside the plot
            Dim boxRect As ScottPlot.PixelRect = c.Box.LabelLastRenderPixelRect
            Dim w As Single = If(boxRect.Width > 4, boxRect.Width, 120.0F)
            Dim h As Single = If(boxRect.Height > 4, boxRect.Height, 30.0F)

            Dim boxToRight As Boolean = c.OffsetX >= 0
            Dim boxAbove As Boolean = c.OffsetY <= 0

            Dim anchorX As Single = tipPixel.X + c.OffsetX
            Dim anchorY As Single = tipPixel.Y + c.OffsetY
            Dim left As Single = If(boxToRight, anchorX, anchorX - w)
            Dim top As Single = If(boxAbove, anchorY - h, anchorY)

            left = Math.Max(dr.Left, Math.Min(left, dr.Right - w))
            top = Math.Max(dr.Top, Math.Min(top, dr.Bottom - h))

            anchorX = If(boxToRight, left, left + w)
            anchorY = If(boxAbove, top + h, top)

            Dim boxCoordinates As ScottPlot.Coordinates = Chart2CommentToCoordinates(host, anchorX, anchorY, target.YAxis, dr)

            c.Box.Location = boxCoordinates
            c.Box.LabelAlignment = If(boxAbove,
                                      If(boxToRight, ScottPlot.Alignment.LowerLeft, ScottPlot.Alignment.LowerRight),
                                      If(boxToRight, ScottPlot.Alignment.UpperLeft, ScottPlot.Alignment.UpperRight))
            c.Box.LabelFontColor = textColour
            c.Box.LabelBackgroundColor = fillColour
            c.Box.LabelBorderColor = seriesColour

            c.LinkData(0) = tip
            c.LinkData(1) = boxCoordinates
            c.Link.Color = seriesColour

            c.Dot.Location = tip
            c.Dot.MarkerStyle.FillColor = seriesColour
            c.Dot.MarkerStyle.LineColor = seriesColour

        Next

    End Sub

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  The comments file:  <csv name>_comments.txt  (comma separated)
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Const ChartCommentsHeader As String = "Device,Trace,DateTime,Sample,OffsetX,OffsetY,Comment"

    Private Function Chart2CommentsPath() As String

        If filePlayback Is Nothing OrElse filePlayback = "" Then Return ""

        Try
            Return Path.Combine(Path.GetDirectoryName(filePlayback), Path.GetFileNameWithoutExtension(filePlayback) & "_comments.txt")
        Catch
            Return ""
        End Try

    End Function

    Private Function Chart2CsvQuote(s As String) As String

        If s Is Nothing Then s = ""
        If s.IndexOfAny(New Char() {","c, """"c, vbCr(0), vbLf(0)}) >= 0 OrElse s <> s.Trim() Then
            Return """" & s.Replace("""", """""") & """"
        End If
        Return s

    End Function

    ' Splits one line of comma separated values (quoted values may contain commas and doubled quotes).
    Private Function Chart2CsvSplit(line As String) As List(Of String)

        Dim fields As New List(Of String)
        Dim current As New StringBuilder()
        Dim inQuotes As Boolean = False
        Dim i As Integer = 0

        Do While i < line.Length
            Dim ch As Char = line(i)
            If inQuotes Then
                If ch = """"c Then
                    If i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                        current.Append(""""c)
                        i += 1
                    Else
                        inQuotes = False
                    End If
                Else
                    current.Append(ch)
                End If
            ElseIf ch = """"c Then
                inQuotes = True
            ElseIf ch = ","c Then
                fields.Add(current.ToString())
                current.Clear()
            Else
                current.Append(ch)
            End If
            i += 1
        Loop

        fields.Add(current.ToString())
        Return fields

    End Function

    ' Line breaks in a comment are kept as \n and a backslash as \\ so every comment stays on one line of the file.
    Private Function Chart2EscapeComment(text As String) As String
        Return text.Replace("\", "\\").Replace(vbCr, "").Replace(vbLf, "\n")
    End Function

    Private Function Chart2UnescapeComment(text As String) As String

        Dim sb As New StringBuilder()
        Dim i As Integer = 0

        Do While i < text.Length
            Dim ch As Char = text(i)
            If ch = "\"c AndAlso i + 1 < text.Length Then
                Dim nextChar As Char = text(i + 1)
                If nextChar = "n"c Then
                    sb.Append(vbLf)
                    i += 2
                    Continue Do
                ElseIf nextChar = "\"c Then
                    sb.Append("\"c)
                    i += 2
                    Continue Do
                End If
            End If
            sb.Append(ch)
            i += 1
        Loop

        Return sb.ToString()

    End Function

    Private Sub Chart2SaveComments()

        Dim path As String = Chart2CommentsPath()
        If path = "" Then Exit Sub

        Try
            ' No comments left: remove the file rather than leave an empty one
            If Chart2Comments.Count = 0 Then
                If File.Exists(path) Then File.Delete(path)
                Exit Sub
            End If

            Dim sb As New StringBuilder()
            sb.AppendLine(ChartCommentsHeader)

            For Each c As ChartComment In Chart2Comments
                sb.AppendLine(String.Join(",", {
                    Chart2CsvQuote(c.Device),
                    Chart2CsvQuote(c.TraceKey),
                    Chart2CsvQuote(c.TimeText),
                    c.SampleIndex.ToString(CultureInfo.InvariantCulture),
                    CInt(Math.Round(c.OffsetX)).ToString(CultureInfo.InvariantCulture),
                    CInt(Math.Round(c.OffsetY)).ToString(CultureInfo.InvariantCulture),
                    Chart2CsvQuote(Chart2EscapeComment(c.Text))}))
            Next

            ' Written to a temporary file first, so a failure part-way cannot damage the existing comments file
            Dim temp As String = path & ".tmp"
            File.WriteAllText(temp, sb.ToString(), New UTF8Encoding(False))
            File.Copy(temp, path, True)
            File.Delete(temp)

        Catch ex As Exception
            MessageBox.Show("The comments could not be saved: " & ex.Message, "Comments", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

    End Sub

    Private Sub Chart2ClearComments()

        For Each c As ChartComment In Chart2Comments
            Chart2RemoveCommentPlottables(c)
        Next

        Chart2Comments.Clear()
        Chart2CommentDrag = Nothing
        Chart2CommentDragHost = Nothing
        Chart2CommentsLoadedFor = ""

    End Sub

    ' Finds the reading with this time stamp, nearest the saved reading number. -1 if there isn't one.
    Private Function Chart2ResolveCommentIndex(times() As String, timeText As String, stored As Integer) As Integer

        If times Is Nothing OrElse times.Length = 0 Then Return -1
        If stored >= 0 AndAlso stored < times.Length AndAlso times(stored) = timeText Then Return stored

        Dim best As Integer = -1
        Dim bestDistance As Integer = Integer.MaxValue

        For i As Integer = 0 To times.Length - 1
            If times(i) = timeText Then
                Dim distance As Integer = Math.Abs(i - stored)
                If distance < bestDistance Then
                    best = i
                    bestDistance = distance
                End If
            End If
        Next

        Return best

    End Function

    ' Loads <csv name>_comments.txt for the CSV that has just been loaded (replacing any comments in memory).
    Private Sub Chart2LoadComments()

        Chart2ClearComments()
        Chart2CommentsLoadedFor = filePlayback

        If Not PlaybackCommentsEnabled Then Exit Sub

        Dim path As String = Chart2CommentsPath()
        If path = "" OrElse Not File.Exists(path) Then Exit Sub

        Dim lines() As String
        Try
            lines = File.ReadAllLines(path, Encoding.UTF8)
        Catch ex As Exception
            MessageBox.Show("The comments file could not be read: " & ex.Message, "Comments", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End Try

        Dim sharedKeys As New List(Of String) From {"Temperature", "Humidity", "PPM Deviation"}
        Dim deviceKeys As New List(Of String) From {"Data", "Mean", "Short Term Mean",
                                                    ChartCommentStatsPrefix & "STDEV", ChartCommentStatsPrefix & "SEM", ChartCommentStatsPrefix & "Max Diff.", ChartCommentStatsPrefix & "PPM Deviation"}
        Dim unmatched As Integer = 0

        For Each line As String In lines

            If line.Trim() = "" OrElse line.StartsWith("Device,", StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim f As List(Of String) = Chart2CsvSplit(line)
            If f.Count < 7 Then Continue For

            Dim c As New ChartComment With {
                .Device = f(0),
                .TraceKey = f(1),
                .TimeText = f(2),
                .Text = Chart2UnescapeComment(f(6))
            }

            Dim stored As Integer = 0
            Integer.TryParse(f(3), NumberStyles.Integer, CultureInfo.InvariantCulture, stored)
            Dim ox, oy As Single
            If Single.TryParse(f(4), NumberStyles.Float, CultureInfo.InvariantCulture, ox) Then c.OffsetX = ox
            If Single.TryParse(f(5), NumberStyles.Float, CultureInfo.InvariantCulture, oy) Then c.OffsetY = oy
            c.SampleIndex = stored

            ' Which trace: device traces by the device's name, the shared ones (Temperature, Humidity, PPM) by their key
            Dim ok As Boolean = sharedKeys.Contains(c.TraceKey) OrElse deviceKeys.Contains(c.TraceKey)
            If ok AndAlso deviceKeys.Contains(c.TraceKey) Then
                If DeviceName1.Text <> "" AndAlso String.Equals(c.Device, DeviceName1.Text, StringComparison.OrdinalIgnoreCase) Then
                    c.Slot = 1
                ElseIf DeviceName2.Text <> "" AndAlso String.Equals(c.Device, DeviceName2.Text, StringComparison.OrdinalIgnoreCase) Then
                    c.Slot = 2
                Else
                    ok = False
                End If
            End If

            ' The reading, then its point on the trace (traces that skip rows have fewer points than readings)
            If ok Then
                Dim rowIndex As Integer = Chart2ResolveCommentIndex(Chart2CommentTimes(c.Device), c.TimeText, stored)
                If rowIndex >= 0 Then
                    Dim column As String = ""
                    For Each t As ChartCommentTarget In Chart2CommentTargets(Chart2CommentHost(c.TraceKey))
                        If t.Key = c.TraceKey AndAlso t.Slot = c.Slot Then
                            column = t.Column
                            Exit For
                        End If
                    Next
                    Dim dataIndex As Integer = Chart2CommentDataFromRow(c.Device, c.Slot, column, rowIndex)
                    If dataIndex >= 0 Then
                        c.SampleIndex = rowIndex
                        c.DataIndex = dataIndex
                    Else
                        ok = False
                    End If
                Else
                    ok = False
                End If
            End If

            c.Matched = ok
            If Not ok Then unmatched += 1

            Chart2Comments.Add(c)
            Chart2EnsureCommentPlottables(c)

        Next

        If unmatched > 0 Then
            MessageBox.Show(unmatched.ToString() & " of the " & Chart2Comments.Count.ToString() & " comments in " & System.IO.Path.GetFileName(path) &
                            " could not be matched to readings in this CSV (the device or time stamp isn't in it), so they are not shown. They are kept in the file.",
                            "Comments", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Chart2RefreshCommentCharts()

    End Sub

End Class
