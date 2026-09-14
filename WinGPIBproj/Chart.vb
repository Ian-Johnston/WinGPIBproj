' Playback chart control

Imports System.Runtime.InteropServices
Imports System.Windows.Forms.DataVisualization.Charting


Public Class Chart

    ' Used to hand off the Allan Deviation pop-up's resize-grip drag to
    ' Windows' own native bottom-right resize handling - same approach as
    ' LiveWatch.vb's Live Analysis chart pop-up.
    <DllImport("user32.dll")>
    Private Shared Function ReleaseCapture() As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As Integer, lParam As Integer) As Integer
    End Function

    ' My.Resources.grip is a dark icon meant for a light background (e.g.
    ' LiveWatch.vb's pop-up) - on this pop-up's solid black background it
    ' would barely be visible, so invert its colours (alpha untouched) via
    ' a ColorMatrix rather than needing a second image resource.
    Private Function InvertGripImage(source As Image) As Bitmap

        Dim inverted As New Bitmap(source.Width, source.Height)

        Dim colorMatrix As New Imaging.ColorMatrix(New Single()() {
            New Single() {-1, 0, 0, 0, 0},
            New Single() {0, -1, 0, 0, 0},
            New Single() {0, 0, -1, 0, 0},
            New Single() {0, 0, 0, 1, 0},
            New Single() {1, 1, 1, 0, 1}
        })

        Using attributes As New Imaging.ImageAttributes()
            attributes.SetColorMatrix(colorMatrix)
            Using g As Graphics = Graphics.FromImage(inverted)
                g.DrawImage(source, New Rectangle(0, 0, source.Width, source.Height),
                            0, 0, source.Width, source.Height, GraphicsUnit.Pixel, attributes)
            End Using
        End Using

        Return inverted

    End Function

    Dim gChartPlayback As Array = Array.CreateInstance(GetType(Double), 500)  ' playback chart

    Dim dataTable1 As New DataTable

    Dim CurrentPos As Integer = 0
    Dim TargetPos As Integer = 49
    Dim RangeReqd As Integer = 49
    Dim EndRange As Integer = 500
    Dim CentreRange As Integer = 0
    Dim filePlayback As String
    Dim numberlinesCSV As Integer
    Dim ChartLoaded As Boolean = False
    Dim CSVfileok As Boolean = False
    Dim PlaybackstrPath As String
    Dim CurrentPosSave As Integer = 0
    Dim TargetPosSave As Integer = 0
    Dim RangeReqdSave As Integer = 0
    Dim Ymin As Double = 0
    Dim Ymax As Double = 0
    Dim BrowseFile As Boolean = False
    'Dim fd As OpenFileDialog = New OpenFileDialog()
    Dim fd As New OpenFileDialog()
    Dim YmaxFromDT As Double = 0
    Dim YminFromDT As Double = 0

    Dim Dev1MinD As Double = 0
    Dim Dev2MinD As Double = 0
    Dim Dev1MaxD As Double = 0
    Dim Dev2MaxD As Double = 0

    Dim CSVdelimit As String = My.Settings.data29

    Dim Devname1 As String
    Dim Devname2 As String
    Dim DualDev As Boolean

    Dim ppmscalerangebit As Double

    Dim TimePoint As Double

    Dim MedianValueCSV As Double
    Dim MedianTempCSV As Double
    Dim medianvalued As Double
    Dim mediantempd As Double

    Dim currentValue As Double

    'Dim maxValue As Double = -10000000.0
    'Dim minValue As Double = 10000000.0

    Dim maxValue As Double = Double.MinValue
    Dim minValue As Double = Double.MaxValue

    Dim DevicenamefromFormTest As String

    Dim Vdiff As Double
    Dim Vchange As Double
    Dim VnomTdiff As Double
    Dim variancevalue As Double
    Dim variancetemp As Double
    Dim calcppmvalue As Double
    ' DEV1avg/DEV2avg/TEMPavg rolling-average buffers - previously all
    ' three shared one "inputvalueMeasurements" array, so whenever two of
    ' them were set to the same window size the buffer never got reset
    ' between traces and their smoothing corrupted each other. Each trace
    ' now gets its own.
    Dim Dev1AvgBuffer() As Double
    Dim Dev2AvgBuffer() As Double
    Dim TempAvgBuffer() As Double

    Dim DEV1rollingAverageValues As New List(Of Double)         ' Create a list to store the rolling average values
    Dim DEV2rollingAverageValues As New List(Of Double)         ' Create a list to store the rolling average values
    Dim TEMProllingAverageValues As New List(Of Double)         ' Create a list to store the rolling average values
    Dim tempcounter As Integer = 0
    Dim tempTEMPcounter As Integer = 0

    ' Retrospective Short-Term Mean traces (Playback top chart) - a rolling
    ' average over the last few raw VALUE readings, recomputed fresh each
    ' time the chart is (re)plotted, independent of the DEV1avg/DEV2avg/
    ' TEMPavg rolling-average feature above. Same idea as LiveWatch.vb's
    ' Short-Term Mean checkbox.
    Private Const ShortTermMeanWindow As Integer = 30

    Dim numberofmetadatalines As Integer = 0


    ' Val() only recognizes "." as a decimal separator, but Format()/.ToString()
    ' write textbox values using the CURRENT CULTURE (e.g. "," on German Windows).
    ' That mismatch let values like YaxisMaximum/YaxisMinimum and MedianValue/MedianTemp
    ' truncate or collapse on non-US locales - use this everywhere such a textbox is
    ' parsed back to a Double, paired with .ToString(..., CultureInfo.InvariantCulture)
    ' on the write side, so both sides agree on "." regardless of OS locale.
    Private Function ParseInvariantDouble(text As String) As Double
        Dim result As Double
        Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, result)
        Return result
    End Function


    Private Sub Formtest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Use Win10-style square window corners on Windows 11
        ApplySquareCorners(Me)

        ' Theme adjustment for Win11, otherwise disabled controls are hardly visible!
        If My.Settings.ThemeSet = True Then
            EnhanceTextBoxBorders(Me)
            MakeButtonsWin10ish(Me)
        End If

        ' Set Timer1 duration - Used for refresh of auto-Playback chart, i.e. auto reload of CSV
        Me.Timer1.Interval = 5000  ' 5secs
        Me.Timer1.Stop()

        ' Large tooltips - same look and feel as the main form (Formtest.vb)
        ToolTip1.OwnerDraw = True
        ToolTip1.InitialDelay = 500   ' ms before first show (default is 1000)
        ToolTip1.ReshowDelay = 5      ' delay when moving between controls
        ToolTip1.AutoPopDelay = 15000 ' how long it stays visible

        DeviceName1.BackColor = Color.Yellow
        DeviceName2.BackColor = Color.Aqua

        RadioButtonDev1.BackColor = Color.Yellow
        RadioButtonDev2.BackColor = Color.Aqua

        RadioButtonPPMDev.BackColor = Color.White
        RadioButtonPPMTempo.BackColor = Color.White

        PlaybackTemp.BackColor = Color.Red
        PlaybackHum.BackColor = Color.DodgerBlue

        ' Playback trace checkboxes - match chart trace colours
        CheckPlaybackDev1Data.BackColor = Color.Yellow
        CheckPlaybackDev1Mean.BackColor = Color.Orange
        CheckPlaybackDev1Stdev.BackColor = Color.LightGray
        CheckPlaybackDev1SEM.BackColor = Color.DeepSkyBlue

        CheckPlaybackDev2Data.BackColor = Color.Aqua
        CheckPlaybackDev2Mean.BackColor = Color.Lime
        CheckPlaybackDev2Stdev.BackColor = Color.Magenta
        CheckPlaybackDev2SEM.BackColor = Color.LimeGreen

        CheckPlaybackDev1MaxDiff.BackColor = Color.Gold
        CheckPlaybackDev1Deviation.BackColor = Color.White

        CheckPlaybackDev2MaxDiff.BackColor = Color.HotPink
        CheckPlaybackDev2Deviation.BackColor = Color.LightGray

        CheckPlaybackDev1ShortTermMean.BackColor = Color.OrangeRed
        CheckPlaybackDev2ShortTermMean.BackColor = Color.Khaki
        CheckPlaybackDev2ShortTermMean.ForeColor = Color.Black

        ' Allan Deviation pop-up chart checkboxes - light tint of each
        ' device's usual colour (Dev1=Yellow, Dev2=Aqua) so they read as
        ' related but distinct from the Data checkboxes above.
        CheckPlaybackDev1Allan.BackColor = Color.LightYellow
        CheckPlaybackDev1Allan.ForeColor = Color.Black
        CheckPlaybackDev2Allan.BackColor = Color.LightCyan
        CheckPlaybackDev2Allan.ForeColor = Color.Black

        GroupBoxMisc.Enabled = True
        GroupBoxMiscTempHum.Enabled = True
        YaxisBox1.Enabled = True

        YaxisLoad.Enabled = True

        ChartScaleMax.Text = My.Settings.data20
        ChartScaleMin.Text = My.Settings.data21
        YaxisMaximum.Text = My.Settings.data24
        YaxisMinimum.Text = My.Settings.data25
        MedianValue.Text = My.Settings.data26
        MedianTemp.Text = My.Settings.data27
        PPMscalerangeentry.Text = My.Settings.data28
        CSVdelimit = My.Settings.data29

        YaxisMaximum.ReadOnly = True
        YaxisMinimum.ReadOnly = True
        RangeRequired.ReadOnly = True

        ButtonScrollLeft.Enabled = False
        ButtonScrollRight.Enabled = False
        ButtonScrollLeftSMALL.Enabled = False
        ButtonScrollRightSMALL.Enabled = False
        ButtonZoomIn.Enabled = False
        ButtonZoomOut.Enabled = False
        ButtonYminInc.Enabled = False
        ButtonYminDec.Enabled = False
        ButtonYmaxInc.Enabled = False
        ButtonYmaxDec.Enabled = False
        ButtonDisplayAll.Enabled = False
        ButtonShiftUp.Enabled = False
        ButtonShiftDn.Enabled = False

        CSVfilenamePlayback.ReadOnly = True

        CheckBoxMedianV.Enabled = True
        CheckBoxMedianT.Enabled = False

        RadioButtonPPMDev.Checked = False
        RadioButtonPPMDev.Enabled = True
        RadioButtonPPMTempo.Enabled = True
        PPMBox1.Enabled = True
        RadioButtonDev1.Checked = True
        MedianValue.Enabled = False
        MedianTemp.Enabled = False
        RadioButtonDev1.Enabled = False
        RadioButtonDev2.Enabled = False
        MedianValueText.Enabled = False
        MedianTempText.Enabled = False
        PPMscalerangeentry.Enabled = False
        PPMscaleText.Enabled = False

        YaxisCheck1.Checked = False
        YaxisCheck2.Checked = False
        YaxisCheck3.Checked = False
        YaxisCheck4.Checked = False

        Loading.Visible = False

        PlaybackstrPath = "C:\Users\" & Environment.UserName & "\Documents\WinGPIBdata"

        CSVfileok = False

        ' Clear the screen and set the load CSV message
        ChartOffReadyForCSV()


        ' ==========================================================
        ' Chart2 initialise
        ' ==========================================================

        Chart2.Location = New Point(1, 202)
        Chart2.Size = New Size(1334, 610)

        Chart2.ChartAreas(0).AxisY.LabelStyle.Enabled = True
        Chart2.ChartAreas(0).AxisX.MajorTickMark.Enabled = True
        Chart2.ChartAreas(0).AxisX.Interval = 95

        Chart2.ChartAreas(0).AxisY.LabelStyle.Font =
        New Font("Verdana", 8)

        Chart2.ChartAreas(0).AxisY.LabelStyle.Format =
        "{000.0000000}"

        Chart2.ChartAreas(0).AxisY.MajorTickMark.Enabled = True
        Chart2.ChartAreas(0).AxisX.MinorTickMark.Enabled = False
        Chart2.ChartAreas(0).AxisY.MinorTickMark.Enabled = False

        Chart2.ChartAreas(0).AxisX.MajorGrid.Enabled = True
        Chart2.ChartAreas(0).AxisY.MajorGrid.Enabled = True
        Chart2.ChartAreas(0).AxisX.MinorGrid.Enabled = True
        Chart2.ChartAreas(0).AxisY.MinorGrid.Enabled = True

        Chart2.ChartAreas(0).AxisX.MajorGrid.LineColor =
        Color.FromArgb(255, 85, 85, 85)

        Chart2.ChartAreas(0).AxisY.MajorGrid.LineColor =
        Color.FromArgb(255, 85, 85, 85)

        Chart2.ChartAreas(0).AxisX.MinorGrid.LineColor =
        Color.FromArgb(150, 85, 85, 85)

        Chart2.ChartAreas(0).AxisY.MinorGrid.LineColor =
        Color.FromArgb(150, 85, 85, 85)

        Chart2.DataBindTable(gChartPlayback)

        Chart2.Series(0).ChartType = 2
        Chart2.Series.Clear()

        Chart2.ChartAreas(0).BorderWidth = 1


        ' ==========================================================
        ' Add Statistics ChartArea BEFORE assigning series to it
        ' ==========================================================

        Dim statsArea As New DataVisualization.Charting.ChartArea("Statistics")

        statsArea.BackColor = Color.Black
        statsArea.BorderWidth = 1

        statsArea.AxisX.LabelStyle.Enabled = False
        statsArea.AxisX.MajorTickMark.Enabled = False
        statsArea.AxisX.MinorTickMark.Enabled = False

        statsArea.AxisX.MajorGrid.Enabled = True
        statsArea.AxisX.MinorGrid.Enabled = False

        statsArea.AxisY.MajorGrid.Enabled = True
        statsArea.AxisY.MinorGrid.Enabled = False

        statsArea.AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        statsArea.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        statsArea.AxisX.MajorGrid.LineColor = Color.FromArgb(255, 85, 85, 85)

        statsArea.AxisY.MajorGrid.LineColor = Color.FromArgb(255, 85, 85, 85)

        ' Match Y-axis scale appearance to main chart
        statsArea.AxisY.LabelStyle.ForeColor = Color.Black
        statsArea.AxisY.LabelStyle.Font = New Font("Verdana", 8)

        statsArea.AxisY.IsLabelAutoFit = True
        statsArea.AxisY.LabelAutoFitStyle = DataVisualization.Charting.LabelAutoFitStyles.DecreaseFont

        'statsArea.AxisY.LabelStyle.Format = "0.0E+00"
        statsArea.AxisY.LabelStyle.Format = "0.0000000"

        statsArea.AxisX.LabelStyle.ForeColor = Color.Black

        statsArea.Position.Auto = False
        statsArea.Position = New DataVisualization.Charting.ElementPosition(7.0F, 77.0F, 91.0F, 20.0F)

        statsArea.InnerPlotPosition.Auto = False
        statsArea.InnerPlotPosition = New DataVisualization.Charting.ElementPosition(8.0F, 5.0F, 88.0F, 88.0F)

        Chart2.ChartAreas.Add(statsArea)


        ' ==========================================================
        ' Main chart area
        ' ==========================================================

        Chart2.ChartAreas(0).Position.Auto = False

        Chart2.ChartAreas(0).Position = New DataVisualization.Charting.ElementPosition(7.0F, 4.0F, 91.0F, 70.0F)

        Chart2.ChartAreas(0).InnerPlotPosition.Auto = False

        Chart2.ChartAreas(0).InnerPlotPosition = New DataVisualization.Charting.ElementPosition(8.0F, 5.0F, 88.0F, 90.0F)


        ' ==========================================================
        ' Add chart series
        ' ==========================================================

        Chart2.Series.Add("Device 1")
        Chart2.Series.Add("Device 2")
        Chart2.Series.Add("Temperature")
        Chart2.Series.Add("Humidity")
        Chart2.Series.Add("PPM Dev 1")

        Chart2.Series.Add("Dev 1 Mean")
        Chart2.Series.Add("Dev 1 STDEV")
        Chart2.Series.Add("Dev 1 SEM")

        Chart2.Series.Add("Dev 2 Mean")
        Chart2.Series.Add("Dev 2 STDEV")
        Chart2.Series.Add("Dev 2 SEM")

        Chart2.Series.Add("Dev 1 Max Diff")
        Chart2.Series.Add("Dev 1 Deviation")

        Chart2.Series.Add("Dev 2 Max Diff")
        Chart2.Series.Add("Dev 2 Deviation")

        Chart2.Series.Add("Dev 1 Short-Term Mean")
        Chart2.Series.Add("Dev 2 Short-Term Mean")


        ' ==========================================================
        ' Assign statistics series to correct ChartAreas
        ' ==========================================================

        Chart2.Series(5).ChartArea =
        Chart2.ChartAreas(0).Name

        Chart2.Series(8).ChartArea =
        Chart2.ChartAreas(0).Name

        Chart2.Series(15).ChartArea =
        Chart2.ChartAreas(0).Name

        Chart2.Series(16).ChartArea =
        Chart2.ChartAreas(0).Name

        Chart2.Series(6).ChartArea = "Statistics"
        Chart2.Series(7).ChartArea = "Statistics"
        Chart2.Series(9).ChartArea = "Statistics"
        Chart2.Series(10).ChartArea = "Statistics"

        Chart2.Series(11).ChartArea = "Statistics"
        Chart2.Series(12).ChartArea = "Statistics"
        Chart2.Series(13).ChartArea = "Statistics"
        Chart2.Series(14).ChartArea = "Statistics"


        ' ==========================================================
        ' Chart types
        ' ==========================================================

        CheckDev1Line.Checked = True
        CheckDev1Point.Checked = False
        CheckDev2Line.Checked = True
        CheckDev2Point.Checked = False

        For i As Integer = 0 To 16

            Chart2.Series(i).ChartType =
            DataVisualization.Charting.SeriesChartType.Line

        Next

        Chart2.Series(0).YValueType =
        DataVisualization.Charting.ChartValueType.Single


        ' ==========================================================
        ' Colours
        ' ==========================================================

        Chart2.Series(0).Color = Color.Yellow
        Chart2.Series(1).Color = Color.Aqua
        Chart2.Series(2).Color = Color.Red
        Chart2.Series(3).Color = Color.DodgerBlue
        Chart2.Series(4).Color = Color.White

        ' Device 1
        Chart2.Series(5).Color = Color.Orange          ' Dev 1 Mean
        Chart2.Series(6).Color = Color.LightGray       ' Dev 1 STDEV
        Chart2.Series(7).Color = Color.DeepSkyBlue     ' Dev 1 SEM
        Chart2.Series(11).Color = Color.Gold           ' Dev 1 Max Diff
        Chart2.Series(12).Color = Color.White          ' Dev 1 PPM Deviation
        Chart2.Series(15).Color = Color.OrangeRed      ' Dev 1 Short-Term Mean

        ' Device 2
        Chart2.Series(8).Color = Color.Lime            ' Dev 2 Mean
        Chart2.Series(9).Color = Color.Magenta         ' Dev 2 STDEV
        Chart2.Series(10).Color = Color.LimeGreen      ' Dev 2 SEM
        Chart2.Series(13).Color = Color.HotPink        ' Dev 2 Max Diff
        Chart2.Series(14).Color = Color.LightGray      ' Dev 2 PPM Deviation
        Chart2.Series(16).Color = Color.Khaki   ' Dev 2 Short-Term Mean


        ' ==========================================================
        ' Start statistics series hidden
        ' ==========================================================

        Chart2.Series(5).Enabled = False
        Chart2.Series(6).Enabled = False
        Chart2.Series(7).Enabled = False

        Chart2.Series(8).Enabled = False
        Chart2.Series(9).Enabled = False
        Chart2.Series(10).Enabled = False

        Chart2.Series(11).Enabled = False
        Chart2.Series(12).Enabled = False
        Chart2.Series(13).Enabled = False
        Chart2.Series(14).Enabled = False


        ' ==========================================================
        ' Existing chart settings
        ' ==========================================================

        Chart2.Legends(0).Enabled = False

        Chart2.ChartAreas(0).AxisX.IntervalAutoMode =
        DataVisualization.Charting.IntervalAutoMode.VariableCount

        Chart2.ChartAreas(0).AxisY.LabelAutoFitStyle =
        DataVisualization.Charting.LabelAutoFitStyles.DecreaseFont

        Chart2.ChartAreas(0).AxisX.IntervalOffset = 0

        Chart2.ChartAreas(0).AxisX.MajorGrid.LineDashStyle =
        DataVisualization.Charting.ChartDashStyle.Dot

        Chart2.ChartAreas(0).AxisY.MajorGrid.LineDashStyle =
        DataVisualization.Charting.ChartDashStyle.Dot

        Chart2.ChartAreas(0).AxisX.LabelStyle.Enabled = False

        Chart2.ChartAreas(0).AxisY2.MajorTickMark.Enabled = True
        Chart2.ChartAreas(0).AxisY2.MinorTickMark.Enabled = False

        Chart2.ChartAreas(0).AxisY2.LabelAutoFitStyle =
        DataVisualization.Charting.LabelAutoFitStyles.DecreaseFont

        Chart2.ChartAreas(0).AxisY2.Interval = 1

        Chart2.ChartAreas(0).AxisY2.MajorGrid.LineColor =
        Color.FromArgb(100, 85, 85, 85)

        Chart2.ChartAreas(0).AxisY2.MinorGrid.LineColor =
        Color.FromArgb(100, 85, 85, 85)


        ' ==========================================================
        ' Temperature
        ' ==========================================================

        Chart2.Series(2).YAxisType =
        DataVisualization.Charting.AxisType.Secondary

        Chart2.ChartAreas(0).AxisY2.Enabled = True
        Chart2.ChartAreas(0).AxisY2.Minimum = 15
        Chart2.ChartAreas(0).AxisY2.Maximum = 50

        Chart2.ChartAreas(0).AxisY2.Enabled =
        DataVisualization.Charting.AxisEnabled.True

        Chart2.ChartAreas(0).AxisY2.LabelStyle.Enabled = True


        ' ==========================================================
        ' Humidity
        ' ==========================================================

        Chart2.Series(3).YAxisType =
        DataVisualization.Charting.AxisType.Secondary


        ' ==========================================================
        ' CSV file format
        ' ==========================================================

        dataTable1.Columns.Add("INDEX", GetType(Integer))
        dataTable1.Columns.Add("DEVICE", GetType(String))
        dataTable1.Columns.Add("DATETIME", GetType(String))
        dataTable1.Columns.Add("VALUE", GetType(Double))
        dataTable1.Columns.Add("TEMP", GetType(Double))
        dataTable1.Columns.Add("HUM", GetType(Double))

        ' New V5 statistics columns
        dataTable1.Columns.Add("DEV1_SAMPLES", GetType(String))
        dataTable1.Columns.Add("DEV1_MEAN", GetType(String))
        dataTable1.Columns.Add("DEV1_STDEV", GetType(String))
        dataTable1.Columns.Add("DEV1_SEM", GetType(String))
        dataTable1.Columns.Add("DEV1_GAIN", GetType(String))

        dataTable1.Columns.Add("DEV2_SAMPLES", GetType(String))
        dataTable1.Columns.Add("DEV2_MEAN", GetType(String))
        dataTable1.Columns.Add("DEV2_STDEV", GetType(String))
        dataTable1.Columns.Add("DEV2_SEM", GetType(String))
        dataTable1.Columns.Add("DEV2_GAIN", GetType(String))

        ' New V6 statistics columns
        dataTable1.Columns.Add("DEV1_MAXDIFF", GetType(String))
        dataTable1.Columns.Add("DEV1_DEVIATION", GetType(String))
        dataTable1.Columns.Add("DEV2_MAXDIFF", GetType(String))
        dataTable1.Columns.Add("DEV2_DEVIATION", GetType(String))

        ' Generated during Playback
        dataTable1.Columns.Add("PPM", GetType(Double))


        ' ==========================================================
        ' Misc
        ' ==========================================================

        RMSwindow.Text = "100"

        LabelTempC.Text = My.Settings.data324
        LabelHum.Text = My.Settings.data325
        RadioButtonPPMTempo.Text =
        "PPM/" & My.Settings.data324


        ' ==========================================================
        ' Playback trace checkboxes
        ' ==========================================================

        CheckPlaybackDev1Data.Checked = True
        CheckPlaybackDev2Data.Checked = True

        CheckPlaybackDev1Mean.Checked = False
        CheckPlaybackDev1Stdev.Checked = False
        CheckPlaybackDev1SEM.Checked = False
        CheckPlaybackDev1MaxDiff.Checked = False
        CheckPlaybackDev1Deviation.Checked = False

        CheckPlaybackDev2Mean.Checked = False
        CheckPlaybackDev2Stdev.Checked = False
        CheckPlaybackDev2SEM.Checked = False
        CheckPlaybackDev2MaxDiff.Checked = False
        CheckPlaybackDev2Deviation.Checked = False

        CheckPlaybackDev1Mean.Enabled = False
        CheckPlaybackDev1Stdev.Enabled = False
        CheckPlaybackDev1SEM.Enabled = False
        CheckPlaybackDev1MaxDiff.Enabled = False
        CheckPlaybackDev1Deviation.Enabled = False

        CheckPlaybackDev2Mean.Enabled = False
        CheckPlaybackDev2Stdev.Enabled = False
        CheckPlaybackDev2SEM.Enabled = False
        CheckPlaybackDev2MaxDiff.Enabled = False
        CheckPlaybackDev2Deviation.Enabled = False
        CheckPlaybackDev2ShortTermMean.Enabled = False
        CheckPlaybackDev2Allan.Enabled = False

    End Sub


    ' Large tooltips - same look and feel as the main form (Formtest.vb)
    Private Sub ToolTip1_Draw(sender As Object, e As DrawToolTipEventArgs) _
    Handles ToolTip1.Draw

        Using f As New Font("Segoe UI", 12.0F)
            e.Graphics.FillRectangle(SystemBrushes.Info, e.Bounds)

            Dim rc As Rectangle = New Rectangle(
            e.Bounds.X + 6,
            e.Bounds.Y + 4,
            e.Bounds.Width - 12,
            e.Bounds.Height - 8
        )

            TextRenderer.DrawText(
            e.Graphics,
            e.ToolTipText,
            f,
            rc,
            Color.Black,
            TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix Or TextFormatFlags.NoClipping
        )
        End Using
    End Sub


    ' Large tooltips - same look and feel as the main form (Formtest.vb)
    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) _
    Handles ToolTip1.Popup

        Dim tt As ToolTip = CType(sender, ToolTip)
        Dim text As String = tt.GetToolTip(e.AssociatedControl)

        Using f As New Font("Segoe UI", 12.0F)
            Dim sz = TextRenderer.MeasureText(
            text,
            f,
            New Size(1200, Integer.MaxValue),
            TextFormatFlags.WordBreak
        )

            e.ToolTipSize = New Size(sz.Width + 14, sz.Height + 8)
        End Using
    End Sub


    Private Sub BrowseToFile_Click(sender As Object, e As EventArgs) Handles BrowseToFile.Click

        ' User Browse to File button, check and load the CSV file into the datatable.
        CurrentPos = 0
        TargetPos = 49
        RangeReqd = 49
        EndRange = 500
        CentreRange = 0

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = PlaybackstrPath
        fd.Filter = "All files (*.csv)|*.csv|All files (*.csv)|*.csv"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then

            PleaseLoadCSV.Visible = False
            filePlayback = fd.FileName
            CSVfilenamePlayback.Text = filePlayback
            BrowseFile = True

            Loading.Visible = True
            Refresh()

        Else

            PleaseLoadCSV.Visible = True
            Return

        End If


        ' ==========================================================
        ' Reset table / chart
        ' ==========================================================

        dataTable1.Clear()

        For i As Integer = 0 To 4
            Chart2.Series(i).Points.Clear()
        Next


        ' A new CSV invalidates whatever the Allan Deviation pop-up (if
        ' open) was showing - it doesn't refresh itself on a new load, so
        ' close it and let the user re-check a box once the new file is
        ' in, rather than leaving it showing stale data from the old CSV.
        CheckPlaybackDev1Allan.Checked = False
        CheckPlaybackDev2Allan.Checked = False

        ' Reset statistics controls until file format is known.
        CheckPlaybackDev1Mean.Checked = False
        CheckPlaybackDev1Stdev.Checked = False
        CheckPlaybackDev1SEM.Checked = False
        CheckPlaybackDev1MaxDiff.Checked = False
        CheckPlaybackDev1Deviation.Checked = False

        CheckPlaybackDev2Mean.Checked = False
        CheckPlaybackDev2Stdev.Checked = False
        CheckPlaybackDev2SEM.Checked = False
        CheckPlaybackDev2MaxDiff.Checked = False
        CheckPlaybackDev2Deviation.Checked = False

        CheckPlaybackDev1Mean.Enabled = False
        CheckPlaybackDev1Stdev.Enabled = False
        CheckPlaybackDev1SEM.Enabled = False
        CheckPlaybackDev1MaxDiff.Enabled = False
        CheckPlaybackDev1Deviation.Enabled = False

        CheckPlaybackDev2Mean.Enabled = False
        CheckPlaybackDev2Stdev.Enabled = False
        CheckPlaybackDev2SEM.Enabled = False
        CheckPlaybackDev2MaxDiff.Enabled = False
        CheckPlaybackDev2Deviation.Enabled = False
        CheckPlaybackDev2ShortTermMean.Enabled = False
        CheckPlaybackDev2Allan.Enabled = False


        ' ==========================================================
        ' Read CSV file
        ' ==========================================================

        Dim lines As List(Of String) =
        IO.File.ReadAllLines(filePlayback).ToList()


        ' Initialize variables.
        CSVdelimit = ""
        MetadataChart.Text = ""

        Dim separator As String = "------------------------------"
        Dim isFirstGroup As Boolean = True
        Dim previousLineIsMetadata As Boolean = False

        numberlinesCSV = 0

        Dim metadataBuilder As New System.Text.StringBuilder()
        Dim rowsToAdd As New List(Of DataRow)()

        Dim statsColumnsDetected As Boolean = False
        Dim v6ColumnsDetected As Boolean = False


        ' ==========================================================
        ' Single pass to analyze, load data and collect metadata
        ' ==========================================================

        For Each line As String In lines

            If line.TrimStart().StartsWith("//") Then

                ' Process metadata.
                If Not previousLineIsMetadata AndAlso
               Not isFirstGroup Then

                    metadataBuilder.AppendLine(separator)

                End If

                metadataBuilder.AppendLine(line.Substring(2))

                previousLineIsMetadata = True

            Else

                previousLineIsMetadata = False
                isFirstGroup = False


                ' Skip blank lines.
                If String.IsNullOrWhiteSpace(line) Then
                    Continue For
                End If


                ' ----------------------------------------------------------
                ' Determine delimiter
                ' ----------------------------------------------------------

                If CSVdelimit = "" Then

                    Dim commaCount As Integer =
                    line.Split(","c).Length - 1

                    Dim semicolonCount As Integer =
                    line.Split(";"c).Length - 1

                    If commaCount >= 5 Then

                        CSVdelimit = ","

                    ElseIf semicolonCount >= 5 Then

                        CSVdelimit = ";"

                    End If

                End If


                If String.IsNullOrEmpty(CSVdelimit) Then
                    Continue For
                End If


                ' ----------------------------------------------------------
                ' Split data line
                ' ----------------------------------------------------------

                Dim values As String() =
                line.Split(New String() {CSVdelimit},
                           StringSplitOptions.None)


                ' Minimum valid old WinGPIB CSV = 6 fields.
                If values.Length < 6 Then

                    Dialog2.Warning1 =
                    "Inconsistent CSV - Invalid data format detected"

                    Dialog2.Warning2 =
                    "Each data line must contain at least 6 fields"

                    Dialog2.Warning3 =
                    "Please fix and try again."

                    Dialog2.ShowDialog(Me)

                    ChartOffReadyForCSV()
                    Return

                End If


                ' First field must be numeric INDEX.
                If Not IsNumeric(values(0)) Then

                    Dialog2.Warning1 =
                    "Inconsistent CSV - Invalid data format detected"

                    Dialog2.Warning2 =
                    "Each data line in the CSV should start with a number"

                    Dialog2.Warning3 =
                    "Please fix and try again."

                    Dialog2.ShowDialog(Me)

                    ChartOffReadyForCSV()
                    Return

                End If


                ' ==========================================================
                ' Detect CSV format
                '
                ' Old CSV:
                '   0 INDEX
                '   1 DEVICE
                '   2 DATETIME
                '   3 VALUE
                '   4 TEMP
                '   5 HUM
                '
                ' New V5 CSV:
                '   + DEV1 Samples
                '   + DEV1 Mean
                '   + DEV1 STDEV
                '   + DEV1 SEM
                '   + DEV1 Gain
                '   + DEV2 Samples
                '   + DEV2 Mean
                '   + DEV2 STDEV
                '   + DEV2 SEM
                '   + DEV2 Gain
                ' ==========================================================

                If values.Length >= 16 Then
                    statsColumnsDetected = True
                End If

                If values.Length >= 20 Then
                    v6ColumnsDetected = True
                End If


                ' ==========================================================
                ' Add data to DataTable
                ' ==========================================================

                Dim row As DataRow = dataTable1.NewRow()


                ' Standard fields - present in old and new CSV.
                row("INDEX") = CInt(Val(values(0)))
                row("DEVICE") = values(1)
                row("DATETIME") = values(2)

                row("VALUE") =
                CDbl(Val(values(3)))

                row("TEMP") =
                CDbl(Val(values(4)))

                row("HUM") =
                CDbl(Val(values(5)))


                ' ----------------------------------------------------------
                ' V5 statistics fields
                ' ----------------------------------------------------------

                If values.Length >= 16 Then

                    row("DEV1_SAMPLES") = values(6)
                    row("DEV1_MEAN") = values(7)
                    row("DEV1_STDEV") = values(8)
                    row("DEV1_SEM") = values(9)
                    row("DEV1_GAIN") = values(10)

                    row("DEV2_SAMPLES") = values(11)
                    row("DEV2_MEAN") = values(12)
                    row("DEV2_STDEV") = values(13)
                    row("DEV2_SEM") = values(14)
                    row("DEV2_GAIN") = values(15)

                Else

                    ' Old CSV - statistics do not exist.
                    row("DEV1_SAMPLES") = ""
                    row("DEV1_MEAN") = ""
                    row("DEV1_STDEV") = ""
                    row("DEV1_SEM") = ""
                    row("DEV1_GAIN") = ""

                    row("DEV2_SAMPLES") = ""
                    row("DEV2_MEAN") = ""
                    row("DEV2_STDEV") = ""
                    row("DEV2_SEM") = ""
                    row("DEV2_GAIN") = ""

                End If


                ' ----------------------------------------------------------
                ' V6 statistics fields (Max Diff / Deviation) - appended
                ' after the original V5 block, so V5 CSVs (exactly 16
                ' fields) still load correctly with these left blank.
                ' ----------------------------------------------------------

                If values.Length >= 20 Then

                    row("DEV1_MAXDIFF") = values(16)
                    row("DEV1_DEVIATION") = values(17)
                    row("DEV2_MAXDIFF") = values(18)
                    row("DEV2_DEVIATION") = values(19)

                Else

                    row("DEV1_MAXDIFF") = ""
                    row("DEV1_DEVIATION") = ""
                    row("DEV2_MAXDIFF") = ""
                    row("DEV2_DEVIATION") = ""

                End If


                ' PPM is generated later by Playback.
                row("PPM") = 0.0


                rowsToAdd.Add(row)

                ' Count valid data lines.
                numberlinesCSV += 1

            End If

        Next


        ' ==========================================================
        ' Update metadata
        ' ==========================================================

        MetadataChart.Text =
        metadataBuilder.ToString()


        numberofmetadatalines =
        lines.Count - numberlinesCSV

        numberlinesCSV =
        lines.Count


        ' ==========================================================
        ' Check if CSV has enough lines
        ' ==========================================================

        If numberlinesCSV < 40 Then

            Loading.Visible = False
            Chart2.Visible = False

            Dialog2.Warning1 = "CSV file empty or too small!"

            Dialog2.Warning2 = "40 lines minimum, your CSV has " & numberlinesCSV & " lines"

            Dialog2.Warning3 = "( X-Axis labels may not display properly below 100 lines )"

            Dialog2.ShowDialog(Me)

            ChartOffReadyForCSV()
            PleaseLoadCSV.Visible = True

            Return

        End If


        ' ==========================================================
        ' Check for missing delimiters
        ' ==========================================================

        If String.IsNullOrEmpty(CSVdelimit) Then

            Loading.Visible = False
            Chart2.Visible = False

            Dialog2.Warning1 = "Inconsistent CSV - Delimiters missing"

            Dialog2.Warning2 = "Each line contains multiple data separated by , or ;"

            Dialog2.Warning3 = "Please fix and try again."

            Dialog2.ShowDialog(Me)

            ChartOffReadyForCSV()
            PleaseLoadCSV.Visible = True

            Return

        End If


        ' ==========================================================
        ' Add rows to dataTable1 in bulk
        ' ==========================================================

        dataTable1.BeginLoadData()

        For Each row As DataRow In rowsToAdd
            dataTable1.Rows.Add(row)
        Next

        dataTable1.EndLoadData()


        ' ==========================================================
        ' Enable statistics controls for V5 CSV
        ' ==========================================================

        If statsColumnsDetected = True Then

            CheckPlaybackDev1Mean.Enabled = True
            CheckPlaybackDev1Stdev.Enabled = True
            CheckPlaybackDev1SEM.Enabled = True

            CheckPlaybackDev2Mean.Enabled = True
            CheckPlaybackDev2Stdev.Enabled = True
            CheckPlaybackDev2SEM.Enabled = True

        Else

            CheckPlaybackDev1Mean.Enabled = False
            CheckPlaybackDev1Stdev.Enabled = False
            CheckPlaybackDev1SEM.Enabled = False

            CheckPlaybackDev2Mean.Enabled = False
            CheckPlaybackDev2Stdev.Enabled = False
            CheckPlaybackDev2SEM.Enabled = False

        End If


        ' ==========================================================
        ' Enable Max Diff / Deviation controls for V6 CSV only
        ' ==========================================================

        If v6ColumnsDetected = True Then

            CheckPlaybackDev1MaxDiff.Enabled = True
            CheckPlaybackDev1Deviation.Enabled = True

            CheckPlaybackDev2MaxDiff.Enabled = True
            CheckPlaybackDev2Deviation.Enabled = True

        Else

            CheckPlaybackDev1MaxDiff.Enabled = False
            CheckPlaybackDev1Deviation.Enabled = False

            CheckPlaybackDev2MaxDiff.Enabled = False
            CheckPlaybackDev2Deviation.Enabled = False

        End If


        ' ==========================================================
        ' Check if dual devices exist and update UI
        ' ==========================================================

        Devname1 =
        dataTable1.Rows(0).ItemArray(1).ToString()

        Devname2 =
        dataTable1.Rows(1).ItemArray(1).ToString()


        If Devname1 = Devname2 Then

            DualDev = False

            DeviceName1.Text = Devname1
            DeviceName2.Text = ""

            DisableDualDeviceControls()

        Else

            DualDev = True

            ' Which of these two device names is "Dev 1" vs "Dev 2" must not
            ' be decided by which one merely happens to log first in the
            ' file - that's arbitrary per run and can differ CSV to CSV (or
            ' even by deleting the first record). DEV1_MEAN/DEV2_MEAN are
            ' fixed to the actual acquisition-time hardware slots, not to
            ' row order, so if row 0's own VALUE tracks DEV2_MEAN more
            ' closely than DEV1_MEAN, row 0's device is really physical
            ' Device 2 - swap the labels so everything downstream (colours,
            ' checkboxes, Mean/STDEV/SEM/MaxDiff/Deviation traces) lines up
            ' with the correct device.
            Dim row0 As DataRow = dataTable1.Rows(0)
            Dim dev1MeanText As String = row0("DEV1_MEAN").ToString().Trim()
            Dim dev2MeanText As String = row0("DEV2_MEAN").ToString().Trim()

            If dev1MeanText <> "" AndAlso dev1MeanText.ToLower() <> "nil" AndAlso
               dev2MeanText <> "" AndAlso dev2MeanText.ToLower() <> "nil" Then

                Dim rowValue As Double = Convert.ToDouble(row0("VALUE"))
                Dim dev1Mean As Double = ParseInvariantDouble(dev1MeanText)
                Dim dev2Mean As Double = ParseInvariantDouble(dev2MeanText)

                If Math.Abs(rowValue - dev2Mean) < Math.Abs(rowValue - dev1Mean) Then
                    Dim swapName As String = Devname1
                    Devname1 = Devname2
                    Devname2 = swapName
                End If

            End If
            ' Else: no usable stats columns (old CSV) - nothing to check
            ' against, so fall back to file row order as before.

            DeviceName1.Text = Devname1
            DeviceName2.Text = Devname2

            EnableDualDeviceControls()

        End If


        ' ==========================================================
        ' Median starting values
        ' ==========================================================

        If Not DualDev Then

            MedianValueCSV =
            dataTable1.Rows(0).ItemArray(3).ToString()

            MedianTempCSV =
            dataTable1.Rows(0).ItemArray(4).ToString()

        ElseIf RadioButtonDev1.Checked Then

            MedianValueCSV =
            dataTable1.Rows(0).ItemArray(3).ToString()

            MedianTempCSV =
            dataTable1.Rows(0).ItemArray(4).ToString()

        ElseIf RadioButtonDev2.Checked Then

            MedianValueCSV =
            dataTable1.Rows(1).ItemArray(3).ToString()

            MedianTempCSV =
            dataTable1.Rows(1).ItemArray(4).ToString()

        End If


        ' ==========================================================
        ' Sample rate
        ' ==========================================================

        If DualDev = False Then

            Dim formatdata As String =
            "yyyy-MM-dd_HH:mm:ss"

            Dim DateTime8th As DateTime =
            DateTime.ParseExact(
                dataTable1.Rows(8)("DATETIME").ToString(),
                formatdata,
                System.Globalization.CultureInfo.InvariantCulture)

            Dim DateTime9th As DateTime =
            DateTime.ParseExact(
                dataTable1.Rows(9)("DATETIME").ToString(),
                formatdata,
                System.Globalization.CultureInfo.InvariantCulture)

            Dim timeDifference As TimeSpan =
            DateTime9th.Subtract(DateTime8th)

            SampleRateSecs.Text =
            timeDifference.TotalSeconds

        Else

            Dim formatdata As String =
            "yyyy-MM-dd_HH:mm:ss"

            Dim DateTime8th As DateTime =
            DateTime.ParseExact(
                dataTable1.Rows(8)("DATETIME").ToString(),
                formatdata,
                System.Globalization.CultureInfo.InvariantCulture)

            Dim DateTime10th As DateTime =
            DateTime.ParseExact(
                dataTable1.Rows(10)("DATETIME").ToString(),
                formatdata,
                System.Globalization.CultureInfo.InvariantCulture)

            Dim timeDifference As TimeSpan =
            DateTime10th.Subtract(DateTime8th)

            SampleRateSecs.Text =
            timeDifference.TotalSeconds

        End If


        ' ==========================================================
        ' Finalize chart
        ' ==========================================================

        CheckPathCSVfile()
        PrintXscale()
        GetMinMaxScales()
        FixTicks()
        AverageNoise()

        Yscaletidy()

        Loading.Visible = False
        PleaseLoadCSV.Visible = False

    End Sub

    Private Sub DisableDualDeviceControls()
        DEV2avg.Enabled = False
        CheckDev2Line.Enabled = False
        CheckDev2Point.Enabled = False
        Dev2MaxMin.Enabled = False
        RMSaverageDev2.Enabled = False
        Label22.Enabled = False
        Label17.Enabled = False
        Label9.Enabled = False

        ' No Device 2 data at all in a single-device CSV, so force
        ' every Dev.2 Playback checkbox off regardless of what the
        ' V5/V6 column-detection block above set them to.
        CheckPlaybackDev2Data.Checked = False
        CheckPlaybackDev2Data.Enabled = False

        CheckPlaybackDev2Mean.Checked = False
        CheckPlaybackDev2Mean.Enabled = False
        CheckPlaybackDev2Stdev.Checked = False
        CheckPlaybackDev2Stdev.Enabled = False
        CheckPlaybackDev2SEM.Checked = False
        CheckPlaybackDev2SEM.Enabled = False

        CheckPlaybackDev2MaxDiff.Checked = False
        CheckPlaybackDev2MaxDiff.Enabled = False
        CheckPlaybackDev2Deviation.Checked = False
        CheckPlaybackDev2Deviation.Enabled = False

        CheckPlaybackDev2ShortTermMean.Checked = False
        CheckPlaybackDev2ShortTermMean.Enabled = False

        CheckPlaybackDev2Allan.Checked = False
        CheckPlaybackDev2Allan.Enabled = False
    End Sub

    Private Sub EnableDualDeviceControls()
        DEV1avg.Enabled = True
        CheckDev1Line.Enabled = True
        CheckDev1Point.Enabled = True
        Dev1MaxMin.Enabled = True
        RMSaverageDev1.Enabled = True
        Label12.Enabled = True
        Label10.Enabled = True
        Label8.Enabled = True

        DEV2avg.Enabled = True
        CheckDev2Line.Enabled = True
        CheckDev2Point.Enabled = True
        Dev2MaxMin.Enabled = True
        RMSaverageDev2.Enabled = True
        Label22.Enabled = True
        Label17.Enabled = True
        Label9.Enabled = True

        ' Dev.2 stats checkboxes (Mean/Stdev/SEM/MaxDiff/Deviation)
        ' are left alone here - their Enabled state is already set
        ' correctly by the V5/V6 column-detection block based on
        ' what's actually in the CSV. Only the raw Dev.2 data
        ' checkbox isn't covered by that block, so re-enable it here.
        CheckPlaybackDev2Data.Enabled = True

        ' Short-Term Mean isn't a recorded CSV column either (it's
        ' recomputed from raw VALUE), so it isn't covered by the V5/V6
        ' block - re-enable it here for the same reason as Dev.2 Data.
        CheckPlaybackDev2ShortTermMean.Enabled = True

        ' Allan Deviation checkbox is likewise not a recorded CSV column.
        CheckPlaybackDev2Allan.Enabled = True
    End Sub


    Private Sub RefreshFile_Click(sender As Object, e As EventArgs)

        RefreshPlaybackCSVFile()
        'ShowAll()
        PleaseLoadCSV.Visible = False

    End Sub


    Private Sub RefreshPlaybackCSVFile()

        ' A slimmed down version of loading the CSV file again as it is written to externally.

        If (CSVfilenamePlayback.Text <> "" And ChartLoaded = True And CSVfileok = True) Then

            ' NOTE: CurrentPos/TargetPos/RangeReqd/EndRange/CentreRange are
            ' deliberately NOT reset here. This function doesn't actually
            ' re-read a fresh window of the file (that loop below is
            ' commented out) - it just refreshes settings/traces against
            ' whatever's already in dataTable1. Resetting the zoom/scroll
            ' bookkeeping here used to silently discard the user's current
            ' zoom level (e.g. Zoom In several times, then check a
            ' checkbox that routes through here) without ever restoring
            ' the actual zoomed view, so the next Scroll/Shift button
            ' would jump back out to whatever this reset left behind.

            'Loading.Visible = True
            Me.Refresh()

            BrowseFile = True
            filePlayback = CSVfilenamePlayback.Text     ' set filepath and file to same as existing

            'dataTable1.Clear()
            'Chart2.Series(0).Points.Clear()
            'Chart2.Series(1).Points.Clear()
            'Chart2.Series(2).Points.Clear()
            'Chart2.Series(3).Points.Clear()
            'Chart2.Series(4).Points.Clear()

            ' open CSV file and check for device
            ' Add entire CSV into datatable
            'numberlinesCSV = (System.IO.File.ReadAllLines(filePlayback).Length)

            ' load CSV to datatable, check each line is not blank
            'Dim linecount As Integer = 1
            'For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(numberlinesCSV)    '.First()           ' .Skip(CurrentPos).Take(TargetPos - CurrentPos)  

            'If line.Contains("//") Then
            ' Skip lines containing "//"
            'Continue For
            'End If

            'linecount = linecount + 1
            'dataTable1.Rows.Add(line.Split(CSVdelimit))
            'Next

            'Devname1 = dataTable1.Rows(0).ItemArray(1).ToString()
            'Devname2 = dataTable1.Rows(1).ItemArray(1).ToString()

            If DualDev = False Then
                ' Get first VALUE & TEMP from CSV read to use for MEDIAN VALUE & MEDIAN TEMP if required later
                MedianValueCSV = dataTable1.Rows(0).ItemArray(3).ToString()
                MedianTempCSV = dataTable1.Rows(0).ItemArray(4).ToString()
            Else
                If RadioButtonDev1.Checked = True Then
                    MedianValueCSV = dataTable1.Rows(0).ItemArray(3).ToString()
                    MedianTempCSV = dataTable1.Rows(0).ItemArray(4).ToString()
                End If
                If RadioButtonDev2.Checked = True Then
                    MedianValueCSV = dataTable1.Rows(1).ItemArray(3).ToString()
                    MedianTempCSV = dataTable1.Rows(1).ItemArray(4).ToString()
                End If
            End If

            CheckPathCSVfile(recalculateYAxis:=False)
            PrintXscale()
            GetMinMaxScales()
            FixTicks()
            AverageNoise()

            Loading.Visible = False

        End If

    End Sub


    Sub ShowAll()

        ' This sub is only used by the ZOOM ALL button....really need to fix RefreshPlaybackCSVFile() so that it covers what this button does also.
        ' Also need to have the ZOOM buttons not read the CSV again, just work with the datatable.

        If (CSVfileok = True) Then

            CurrentPos = 0
            TargetPos = numberlinesCSV
            RangeReqd = numberlinesCSV
            RangeRequired.Text = numberlinesCSV

            CurrentPosition.Text = CurrentPos
            TargetPosition.Text = TargetPos
            RangeRequired.Text = RangeReqd

            dataTable1.Clear()
            Chart2.Series(0).Points.Clear()
            Chart2.Series(1).Points.Clear()
            Chart2.Series(2).Points.Clear()
            Chart2.Series(3).Points.Clear()
            'Chart2.Series(4).Points.Clear()

            ' Pull in batch of RangeReqd lines
            For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
                If line.Contains("//") Then
                    ' Skip lines containing "//"
                    Continue For
                End If
                AddPlaybackCSVRow(line)
            Next

            FilterDeviceName1()
            FilterDeviceName2()
            FilterShortTermMeanDevice1()
            FilterShortTermMeanDevice2()
            FilterTempDevice1()
            FilterHumDevice1()
            GeneratePPMColumn()
            FilterGenPPMDevice1()
            FilterGenPPMDevice2()
            UpdatePlaybackStatsSeries()

            ' ZOOM ALL just reloaded the entire file into dataTable1, so
            ' this is the one place Auto Min/Max should actually re-fit to
            ' everything - refresh YmaxFromDT/YminFromDT here (otherwise
            ' GetMinMaxScales() reuses whatever was left over from the
            ' initial load's smaller window, and traces outside that stale
            ' range stay clipped even after "showing all"). X-axis-only
            ' navigation (Zoom In/Out, Scroll, Shift) and other checkboxes
            ' deliberately do NOT do this - the Y-axis should only move
            ' when the user asks it to via Zoom All or the Y-axis controls.
            If CheckBoxMaxMin.Checked Then

                Dim scanMax As Double = Double.MinValue
                Dim scanMin As Double = Double.MaxValue

                For Each row As DataRow In dataTable1.Rows
                    Dim currentValue As Double = CDbl(row("VALUE"))
                    If currentValue > scanMax Then scanMax = currentValue
                    If currentValue < scanMin Then scanMin = currentValue
                Next

                YmaxFromDT = scanMax
                YminFromDT = scanMin

                If YmaxFromDT - YminFromDT = 0 Then
                    YmaxFromDT += YmaxFromDT / 1000
                    YminFromDT -= YmaxFromDT / 1000
                End If

            End If

            'Get max and min values of Dev1 & Dev2, keep whichever is max/min value and use for setting scale
            GetMinMaxScales()

            DevicesMinMax()

            FixTicks()

        End If

    End Sub


    Private Sub FixTicks()

        ' Clear any existing custom labels.
        Chart2.ChartAreas(0).AxisX.CustomLabels.Clear()

        ' Rotate the labels on the secondary X-axis to display vertically
        'Chart2.ChartAreas(0).AxisX.LabelStyle.Angle = -90


        ' Single device CSV
        If DualDev = False Then
            ' Assuming ScaleX1.Text and ScaleX28.Text contain the values of the first and last data points
            Dim minX As Double
            minX = Double.Parse(ScaleX1.Text)
            'minX = Val(ScaleX1.Text)

            Dim maxX As Double
            maxX = Double.Parse(ScaleX28.Text)
            'maxX = Val(ScaleX28.Text)

            ' min max counts
            Dim minXc As Double
            minXc = Double.Parse(CurrentPosition.Text)
            'minXc = Val(CurrentPosition.Text)

            Dim maxXc As Double
            maxXc = Double.Parse(TargetPosition.Text)
            'maxXc = Val(TargetPosition.Text)

            ' Calculate the total range and the desired number of ticks (28 in this case)
            Dim totalRange As Double = maxX - minX
            Dim numberOfTicks As Integer = 28

            ' Calculate the interval to evenly space the ticks
            Dim interval As Double = totalRange / (numberOfTicks - 1)           ' mins
            Dim intervalc As Double = (maxXc - minXc) / (numberOfTicks - 1)     ' counts

            ' Ratio of counts to mins. Xscaletotal can be "0" for a CSV
            ' whose logged duration rounds to 0 minutes (e.g. a very short
            ' file, or - as with a repeated-timestamp test file - one where
            ' every row shares the same DATETIME so no elapsed time can be
            ' derived at all); dividing by that zero produces Infinity,
            ' which throws OverflowException when narrowed to Integer.
            ' Fall back to a ratio of 1 rather than crash.
            Dim ticklabelgridratioRaw As Double = (maxXc - minXc) / Val(Xscaletotal.Text)
            Dim ticklabelgridratio As Integer = 1
            If Not Double.IsNaN(ticklabelgridratioRaw) AndAlso
               Math.Abs(ticklabelgridratioRaw) <= Integer.MaxValue Then
                ticklabelgridratio = CInt(ticklabelgridratioRaw)
                If ticklabelgridratio = 0 Then ticklabelgridratio = 1
            End If

            With Chart2.ChartAreas(0).AxisX
                '.Minimum = minX
                '.Maximum = maxX

                .MinorGrid.Enabled = False
                .MinorTickMark.Enabled = False

                .MajorGrid.Interval = intervalc
                .MajorTickMark.Enabled = True
                .MajorTickMark.Interval = intervalc
                .LabelStyle.Enabled = True
                .LabelStyle.Interval = intervalc
                .LabelStyle.Font = New Font("Arial", 8)

                ' Calculate the x-axis labels
                For i As Double = 1 To (Val(CSVfileLines.Text) + 2) Step intervalc              ' + 2 at the end seems to help fill in the far right X-Scale label that is sometimes missing!
                    Dim ii As Double = i / ticklabelgridratio                                   ' Value to be displayed, i.e. 0 / 69.5 = 0, or, 1877 / 12 = 15
                    .CustomLabels.Add(i - intervalc / 2, i + intervalc / 2, (ii + Val(CurrentPosition.Text) / ticklabelgridratio).ToString("0.00"))   ' Format to X.XX
                Next

            End With

        End If


        ' Dual device CSV
        If DualDev = True Then
            ' Assuming ScaleX1.Text and ScaleX28.Text contain the values of the first and last data points
            Dim minX As Double
            minX = Double.Parse(ScaleX1.Text)

            Dim maxX As Double
            maxX = Double.Parse(ScaleX28.Text) / 2

            ' min max counts
            Dim minXc As Double
            minXc = Double.Parse(Val(CurrentPosition.Text) / 2)

            Dim maxXc As Double
            maxXc = Double.Parse(Val(TargetPosition.Text) / 2)

            ' Calculate the total range and the desired number of ticks (28 in this case)
            Dim totalRange As Double = maxX - minX
            Dim numberOfTicks As Integer = 28

            ' Calculate the interval to evenly space the ticks
            Dim interval As Double = totalRange / (numberOfTicks - 1)           ' mins
            Dim intervalc As Double = (maxXc - minXc) / (numberOfTicks - 1)     ' counts

            ' Ratio of counts to mins - see the single-device branch above
            ' for why this needs to be guarded against Xscaletotal = "0".
            Dim ticklabelgridratioRaw As Double = (maxXc - minXc) / Val(Xscaletotal.Text)
            Dim ticklabelgridratio As Integer = 1
            If Not Double.IsNaN(ticklabelgridratioRaw) AndAlso
               Math.Abs(ticklabelgridratioRaw) <= Integer.MaxValue Then
                ticklabelgridratio = CInt(ticklabelgridratioRaw)
                If ticklabelgridratio = 0 Then ticklabelgridratio = 1
            End If

            With Chart2.ChartAreas(0).AxisX

                .MinorGrid.Enabled = False
                .MinorTickMark.Enabled = False

                .MajorGrid.Interval = intervalc
                .MajorTickMark.Enabled = True
                .MajorTickMark.Interval = intervalc
                .LabelStyle.Enabled = True
                .LabelStyle.Interval = intervalc
                .LabelStyle.Font = New Font("Arial", 8)

                For i As Double = 1 To (Val(CSVfileLines.Text) + 2) Step intervalc              ' + 2 at the end seems to help fill in the far right X-Scale label that is sometimes missing!
                    Dim ii As Integer = i / ticklabelgridratio                                   ' Value to be displayed, i.e. 0 / 69.5 = 0, or, 1877 / 12 = 15
                    .CustomLabels.Add(i - intervalc / 2, i + intervalc / 2, (ii + (Val(CurrentPosition.Text) / 2) / ticklabelgridratio).ToString("0.00"))   ' Format to X.XX
                Next

            End With
        End If

    End Sub

    Sub GetSeconds()

        ' Work out how many seconds have elapsed in the CSV for use with the x-axis labels in the chart, CSV mins total, start & end etc.

        Dim DateStart As Date
        Dim TimeStart As Date
        Dim DateStop As Date
        Dim TimeStop As Date
        Dim split As String()
        Dim DateTimeSplit1 As DateTime
        Dim DateTimeSplit2 As DateTime

        ' Find the first non-metadata line
        Dim startRowIndex As Integer = 0
        For i As Integer = 0 To dataTable1.Rows.Count - 1
            If Not dataTable1.Rows(i).ItemArray(2).ToString().StartsWith("//") Then
                startRowIndex = i
                Exit For
            End If
        Next

        ' Extract start timestamp from CSV
        Dim StartTimestamp As String = dataTable1.Rows(startRowIndex).ItemArray(2).ToString() ' start at non-metadata row
        split = StartTimestamp.Split("_"c)
        DateStart = Date.Parse(split(0).Trim())
        TimeStart = Date.Parse(split(1).Trim())

        ' Find the last non-empty line
        Dim endRowIndex As Integer = dataTable1.Rows.Count - 1
        For i As Integer = dataTable1.Rows.Count - 1 To startRowIndex Step -1
            If Not String.IsNullOrEmpty(dataTable1.Rows(i).ItemArray(2).ToString().Trim()) Then
                endRowIndex = i
                Exit For
            End If
        Next

        ' Extract stop timestamp from CSV
        Dim StopTimestamp As String = dataTable1.Rows(endRowIndex).ItemArray(2).ToString() ' end at last non-empty row
        split = StopTimestamp.Split("_"c)
        DateStop = Date.Parse(split(0).Trim())
        TimeStop = Date.Parse(split(1).Trim())

        ' Calculate elapsed time
        If (DualDev = True) Then
            DateTimeSplit1 = DateStart.Add(TimeStart.TimeOfDay)
            DateTimeSplit2 = DateStop.Add(TimeStop.TimeOfDay)
            TimePoint = ((DateTimeSplit2 - DateTimeSplit1).TotalSeconds) / ((endRowIndex - startRowIndex) / 2) ' /2 due to dual device CSV so half numbers of entries each
        Else
            DateTimeSplit1 = DateStart.Add(TimeStart.TimeOfDay)
            DateTimeSplit2 = DateStop.Add(TimeStop.TimeOfDay)
            TimePoint = ((DateTimeSplit2 - DateTimeSplit1).TotalSeconds) / (endRowIndex - startRowIndex)
        End If

        ' A CSV where every row shares the same DATETIME (or logs faster
        ' than the timestamp's 1-second resolution) computes an elapsed
        ' time of 0, so TimePoint ends up 0 - which then zeroes out
        ' MinsTotal/Xscaletotal downstream and can divide by zero in
        ' FixTicks. Fall back to 1 second/sample so the chart still shows
        ' sensible (if approximate) time labels instead of all zeros.
        If TimePoint <= 0 Then TimePoint = 1

    End Sub


    Sub PrintXscale()

        ' TimePoint = time per point in secs

        Dim TPsecs As Long

        TPsecs = (TargetPos * TimePoint) / 60     ' IE. 905 entries @ 1sec sample rate = 15.08

        Dim CPsecs As Long = (CurrentPos * TimePoint) / 60    ' Start       IE. 0 * 1 / 600 = n
        Dim DispSecs As Long = TPsecs - CPsecs
        Dim DispSecs28 As Long = (DispSecs / 27)                   ' per div

        ' Each div point therefore = (DispSecs28 * number) + start  IE. 905 entries @ 1sec sample rate = 15mins, so (15/27)*0 = 0 thro to (15/27)*27 = 15
        Xscaletotal.Text = (Val(RangeRequired.Text) * Val(MinsTotal.Text)) / numberlinesCSV ' length in mins of x-scale at current zoom/position

        Dim XscaletotalInt As Double = Xscaletotal.Text
        'Xscaletotal.Text = "" & Format(Math.Round(XscaletotalInt, 3), "###.0")
        Xscaletotal.Text = "" & Format(Math.Round(XscaletotalInt, 3), "0.0")

        ' Only display X scale divisions if more than 60mins log data
        If (TPsecs > 1) Then       ' was TPsecs 27............bug here so disabled for now.
            ' Make Xscale
            'ScaleX1.Text = "" & Format(Math.Round(((XscaletotalInt / 27) * 0) + CPsecs, 2), "#0.0")           ' Format(Ymin, "#0.00000000")
            'ScaleX28.Text = "" & Format(Math.Round(((XscaletotalInt / 27) * 27) + CPsecs, 2), "#0.0")
            ScaleX1.Text = Format(Math.Round(((XscaletotalInt / 27) * 0) + CPsecs, 2))           ' Format(Ymin, "#0.00000000")
            ScaleX28.Text = Format(Math.Round(((XscaletotalInt / 27) * 27) + CPsecs, 2))
        End If

        If DualDev = False Then
            MinsTotal.Text = Format(Math.Round((numberlinesCSV * TimePoint) / 60, 2), "#0")
        Else
            MinsTotal.Text = Format(Math.Round(((numberlinesCSV / 2) * TimePoint) / 60, 2), "#0")
        End If

        ' Dynamically set grid lines major & minor
        ' X-axis
        Chart2.ChartAreas(0).AxisX.MajorGrid.Interval = Xscaletotal.Text * 1.024 * 10
        Chart2.ChartAreas(0).AxisX.MinorGrid.Interval = Xscaletotal.Text * 0.513 * 10
        ' Y-Axis
        Chart2.ChartAreas(0).AxisY.MajorGrid.Interval = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 8
        Chart2.ChartAreas(0).AxisY.MinorGrid.Interval = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
        'Chart2.Refresh()   ' This glitches the chart :-(


    End Sub


    Sub PrintYscale()

        ' Make scale
        Scale1.Text = Format(Math.Round(ppmscalerangebit * 12, 2), "#0.00")           ' Format(Ymin, "#0.00000000")          Was using "--   " & & Format(Math.Round(ppmscalerangebit * 12, 2), "#0.00")
        Scale2.Text = Format(Math.Round(ppmscalerangebit * 11, 2), "#0.00")
        Scale3.Text = Format(Math.Round(ppmscalerangebit * 10, 2), "#0.00")
        Scale4.Text = Format(Math.Round(ppmscalerangebit * 9, 2), "#0.00")
        Scale5.Text = Format(Math.Round(ppmscalerangebit * 8, 2), "#0.00")
        Scale6.Text = Format(Math.Round(ppmscalerangebit * 7, 2), "#0.00")
        Scale7.Text = Format(Math.Round(ppmscalerangebit * 6, 2), "#0.00")
        Scale8.Text = Format(Math.Round(ppmscalerangebit * 5, 2), "#0.00")
        Scale9.Text = Format(Math.Round(ppmscalerangebit * 4, 2), "#0.00")
        Scale10.Text = Format(Math.Round(ppmscalerangebit * 3, 2), "#0.00")
        Scale11.Text = Format(Math.Round(ppmscalerangebit * 2, 2), "#0.00")
        Scale12.Text = Format(Math.Round(ppmscalerangebit, 2), "#0.00")
        Scale13.Text = Format(Math.Round(0, 2), "#0.00")         '"0.0"
        Scale14.Text = Format(-Math.Round(ppmscalerangebit, 2), "#0.00")
        Scale15.Text = Format(-Math.Round(ppmscalerangebit * 2, 2), "#0.00")
        Scale16.Text = Format(-Math.Round(ppmscalerangebit * 3, 2), "#0.00")
        Scale17.Text = Format(-Math.Round(ppmscalerangebit * 4, 2), "#0.00")
        Scale18.Text = Format(-Math.Round(ppmscalerangebit * 5, 2), "#0.00")
        Scale19.Text = Format(-Math.Round(ppmscalerangebit * 6, 2), "#0.00")
        Scale20.Text = Format(-Math.Round(ppmscalerangebit * 7, 2), "#0.00")
        Scale21.Text = Format(-Math.Round(ppmscalerangebit * 8, 2), "#0.00")
        Scale22.Text = Format(-Math.Round(ppmscalerangebit * 9, 2), "#0.00")
        Scale23.Text = Format(-Math.Round(ppmscalerangebit * 10, 2), "#0.00")
        Scale24.Text = Format(-Math.Round(ppmscalerangebit * 11, 2), "#0.00")
        Scale25.Text = Format(-Math.Round(ppmscalerangebit * 12, 2), "#0.00")

    End Sub










    Private Sub CheckPathCSVfile(Optional recalculateYAxis As Boolean = True)

        ' With the data now in the datatable now process it.
        '
        ' recalculateYAxis gates two things further down that should only
        ' happen on a genuine fresh load: the Auto Min/Max Y-axis recompute,
        ' and resetting CurrentPos/TargetPos/RangeReqd back to the full
        ' file. It defaults True so BrowseToFile_Click (a real fresh load)
        ' behaves exactly as before. RefreshPlaybackCSVFile() passes False,
        ' because it calls this without having re-read the file first -
        ' dataTable1 still only holds whatever the last zoom/scroll left in
        ' it, so doing either of those here would silently rescale the
        ' Y-axis and/or discard the user's current zoom level.

        ' Flag is true if user browsed for file, false if using text boxes
        If (BrowseFile = False) Then

            ' Browsefile is false so user aborted
            CSVfileok = False
            Dialog2.Warning1 = "A CSV file must be selected"
            Dialog2.Warning2 = ""
            Dialog2.Warning3 = ""
            'Dialog2.Show() ' this method positions anywhere!
            Dialog2.ShowDialog(Me)  ' this method positions centre of parent form, and requires to hit OK to return back to parent

        Else

            ' Browsefile is true so user successfuly browsed to file
            BrowseFile = False  ' reset flag
            CSVfileok = True

            ' If displaying dual log then number of lines will be exactly half, so divide by 2, else assumed all are for 1 device.
            If (DualDev = True) Then
                ' Dual device
                'numberlinesCSV = (System.IO.File.ReadAllLines(filePlayback).Length / 2) + 1     ' need to add 1 here otherwise Device 2 will be missing 1 data point at end......something broke here hmmmmm!
                numberlinesCSV = System.IO.File.ReadAllLines(filePlayback).Length
                'CSVfileLines.Text = (numberlinesCSV / 2) - numberofmetadatalines
                CSVfileLines.Text = CInt(Math.Ceiling((numberlinesCSV / 2) - numberofmetadatalines)) + 1
            Else
                ' CSV file is a single device only
                numberlinesCSV = System.IO.File.ReadAllLines(filePlayback).Length       ' don't need to half the value now....?
                CSVfileLines.Text = numberlinesCSV - numberofmetadatalines
            End If

            ' CSVfileLines.Text = numberlinesCSV - numberofmetadatalines
            EndRange = numberlinesCSV

            YaxisMaximum.ReadOnly = False
            YaxisMinimum.ReadOnly = False
            RangeRequired.ReadOnly = False
            ButtonScrollLeft.Enabled = True
            ButtonScrollRight.Enabled = True
            ButtonScrollLeftSMALL.Enabled = True
            ButtonScrollRightSMALL.Enabled = True
            ButtonZoomIn.Enabled = True
            ButtonZoomOut.Enabled = True
            ButtonYminInc.Enabled = True
            ButtonYminDec.Enabled = True
            ButtonYmaxInc.Enabled = True
            ButtonYmaxDec.Enabled = True
            ButtonDisplayAll.Enabled = True
            ButtonShiftUp.Enabled = True
            ButtonShiftDn.Enabled = True

            ' Temp override the above for testing
            'RangeReqd = numberlinesCSV

            ' Only reset the zoom/scroll window to the full file on a
            ' genuine fresh load. RefreshPlaybackCSVFile() calls this with
            ' recalculateYAxis:=False specifically because it's refreshing
            ' settings/traces against whatever the user has already
            ' zoomed/scrolled to - forcing CurrentPos/TargetPos/RangeReqd
            ' back to the full range here discarded that zoom level (the
            ' next Scroll/Zoom button would then jump from wherever this
            ' left things, not from where the user actually was).
            If recalculateYAxis Then

                'If (DualDev = True) Then
                'RangeRequired.Text = (numberlinesCSV / 2) - numberofmetadatalines
                'TargetPosition.Text = (numberlinesCSV / 2) - numberofmetadatalines
                'RangeReqd = numberlinesCSV / 2
                'Else
                RangeRequired.Text = numberlinesCSV - numberofmetadatalines
                TargetPosition.Text = numberlinesCSV - numberofmetadatalines
                RangeReqd = numberlinesCSV - numberofmetadatalines
                'End If

                CurrentPosition.Text = CurrentPos
                CurrentPos = 0
                TargetPos = RangeReqd

            End If

            ' Print Yscale to chart
            GetSeconds()
            PrintXscale()

            Chart2.Visible = True

            Scale1.Visible = True
            Scale2.Visible = True
            Scale3.Visible = True
            Scale4.Visible = True
            Scale5.Visible = True
            Scale6.Visible = True
            Scale7.Visible = True
            Scale8.Visible = True
            Scale9.Visible = True
            Scale10.Visible = True
            Scale11.Visible = True
            Scale12.Visible = True
            Scale13.Visible = True
            Scale14.Visible = True
            Scale15.Visible = True
            Scale16.Visible = True
            Scale17.Visible = True
            Scale18.Visible = True
            Scale19.Visible = True
            Scale20.Visible = True
            Scale21.Visible = True
            Scale22.Visible = True
            Scale23.Visible = True
            Scale24.Visible = True
            Scale25.Visible = True
            ButtonShiftUp.Visible = True
            ButtonShiftDn.Visible = True
            Xscale.Visible = True
            Xscaletotal.Visible = True
            LabelTempC.Visible = True
            LabelHum.Visible = True
            LabelPPMtop.Visible = True
            LabelPPMdegctop.Visible = True
            Loading.Visible = True

            'dataTable1.Clear()
            Chart2.Series(0).Points.Clear()
            Chart2.Series(1).Points.Clear()
            Chart2.Series(2).Points.Clear()
            Chart2.Series(3).Points.Clear()
            Chart2.Series(4).Points.Clear()


            ' dataTable1 format:
            ' "INDEX" Integer
            ' "DEVICE" String
            ' "DATETIME" String
            ' "VALUE" Double
            ' "TEMP" Double
            ' "HUM" Double

            ' Add entire CSV into datatable, whether dual or single device
            'For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
            'If line.Contains("//") Then
            ' Skip lines containing "//"
            'Continue For
            'End If
            'dataTable1.Rows.Add(line.Split(CSVdelimit))
            'Next

            FilterDeviceName1()
            FilterDeviceName2()
            FilterShortTermMeanDevice1()
            FilterShortTermMeanDevice2()
            FilterTempDevice1()
            FilterHumDevice1()

            DevicesMinMax()         ' get device min & max values



            If CheckBoxMaxMin.Checked AndAlso recalculateYAxis Then

                ' Reset min and max before scanning the current data
                maxValue = Double.MinValue
                minValue = Double.MaxValue

                ' Single loop to compute min and max values
                For Each row As DataRow In dataTable1.Rows
                    Dim currentValue As Double = CDbl(row("VALUE"))
                    If currentValue > maxValue Then maxValue = currentValue
                    If currentValue < minValue Then minValue = currentValue
                Next

                ' Assign to variables
                YmaxFromDT = maxValue
                YminFromDT = minValue

                ' Adjust if min and max are the same
                If YmaxFromDT - YminFromDT = 0 Then
                    YmaxFromDT += YmaxFromDT / 1000
                    YminFromDT -= YmaxFromDT / 1000
                End If


                Chart2.ChartAreas(0).AxisY.IsLogarithmic = False
                ButtonShiftUp.Enabled = True
                ButtonShiftDn.Enabled = True
                Chart2.ChartAreas(0).AxisY.Maximum = Math.Round(YmaxFromDT, 7)
                Chart2.ChartAreas(0).AxisY.Minimum = Math.Round(YminFromDT, 7)
                YaxisMaximum.Text = YmaxFromDT.ToString(Globalization.CultureInfo.InvariantCulture)
                YaxisMinimum.Text = YminFromDT.ToString(Globalization.CultureInfo.InvariantCulture)

                Dim result As Double = (YmaxFromDT - YminFromDT) / 32
                YaxisPerDiv.Text = result.ToString("#0.000000000")

                ' Set axis interval
                Chart2.ChartAreas(0).AxisY.Interval = (YmaxFromDT - YminFromDT) / 32

            End If







            'Temp/Hum scale setting
            Chart2.ChartAreas(0).AxisY2.Minimum = ParseInvariantDouble(ChartScaleMin.Text)
            Chart2.ChartAreas(0).AxisY2.Maximum = ParseInvariantDouble(ChartScaleMax.Text)

            Chart2.ChartAreas(0).AxisY2.Interval = (ParseInvariantDouble(ChartScaleMax.Text) - ParseInvariantDouble(ChartScaleMin.Text)) / 32
            Chart2.ChartAreas(0).AxisY2.LabelStyle.Format = "00.0"



            ' Generate PPM column in table from data, then plot it.
            ' Extracted into GeneratePPMColumn() so zoom/scroll/shift
            ' can also regenerate PPM values for whatever subset of
            ' rows they just reloaded, instead of only ever running
            ' once against the initial full load.
            If (CheckBoxPPMenable.Checked = True) Then

                GeneratePPMColumn()

                FilterGenPPMDevice1()
                FilterGenPPMDevice2()

                PrintYscale()
                Yscaletidy()      ' Tidy up X-scale annotations on graph in order to keep length same irrespective of numerical data and No. DP's

                LabelPPMtop.Visible = True

                If RadioButtonPPMTempo.Checked = True Then
                    LabelPPMdegctop.Visible = True
                Else
                    LabelPPMdegctop.Visible = False
                End If

            Else

                ' Erase scale
                Scale1.Text = ""
                Scale2.Text = ""
                Scale3.Text = ""
                Scale4.Text = ""
                Scale5.Text = ""
                Scale6.Text = ""
                Scale7.Text = ""
                Scale8.Text = ""
                Scale9.Text = ""
                Scale10.Text = ""
                Scale11.Text = ""
                Scale12.Text = ""
                Scale13.Text = ""
                Scale14.Text = ""
                Scale15.Text = ""
                Scale16.Text = ""
                Scale17.Text = ""
                Scale18.Text = ""
                Scale19.Text = ""
                Scale20.Text = ""
                Scale21.Text = ""
                Scale22.Text = ""
                Scale23.Text = ""
                Scale24.Text = ""
                Scale25.Text = ""

                LabelPPMtop.Visible = False
                LabelPPMdegctop.Visible = False

            End If


            ChartLoaded = True      ' set flag allowing chart controls to work

            'DevicesMinMax()         ' get device min & max values

        End If

        'CheckBoxMaxMin.Checked = False

    End Sub










    Function CalculateRollingAverage(ByVal inputvalue As Double, numDataPoints As Integer, ByRef buffer() As Double) As Double                     ' rolling average

        If numDataPoints < 1 Then
            numDataPoints = 1
        End If

        If numDataPoints > 500 Then
            numDataPoints = 500
        End If

        ' Initialize this trace's own buffer if not already done
        If buffer Is Nothing OrElse buffer.Length <> numDataPoints Then
            ReDim buffer(numDataPoints - 1)
        End If

        ' Update the array with the latest values
        For i As Integer = 0 To numDataPoints - 2
            buffer(i) = buffer(i + 1)
        Next
        buffer(numDataPoints - 1) = inputvalue

        ' Calculate the rolling average of the values
        Dim sumValue As Double = 0.0
        Dim numValidPoints As Integer = 0 ' To track the number of valid data points (non-zero)
        For i As Integer = 0 To numDataPoints - 1
            If buffer(i) <> 0 Then
                sumValue += buffer(i)
                numValidPoints += 1
            End If
        Next

        ' Avoid division by zero
        If numValidPoints > 0 Then
            Dim rollingAveragevalueoutput As Double = sumValue / numValidPoints

            ' Return the rolling average value
            Return rollingAveragevalueoutput
        End If

        ' Return 0 if there are no valid data points
        Return 0.0
    End Function


    Private Sub ButtonZoomIn_Click(sender As Object, e As EventArgs) Handles ButtonZoomIn.Click

        If ChartLoaded = True Then

            ' CheckPathCSVfile()

            If (CSVfileok = True) Then

                ' Save off settings and calculate new zoom out settings
                CurrentPosSave = CurrentPos
                TargetPosSave = TargetPos
                RangeReqdSave = RangeReqd
                CentreRange = CurrentPos + (RangeReqd / 2)  ' current centre position
                CurrentPos = CentreRange - (RangeReqd / 4)
                TargetPos = CentreRange + (RangeReqd / 4)
                'RangeReqd = RangeReqd / 2
                RangeReqd /= 2
                RangeRequired.Text = RangeReqd / 2

                ' For dual-device CSVs, each sample is TWO consecutive
                ' lines (Dev1 then Dev2). Keep the window aligned to
                ' whole pairs so zooming never splits a pair - otherwise
                ' the two devices end up with mismatched sample counts
                ' and drift out of alignment with each other.
                If DualDev = True Then
                    If CurrentPos Mod 2 <> 0 Then CurrentPos -= 1
                    Dim windowLen As Integer = TargetPos - CurrentPos
                    If windowLen Mod 2 <> 0 Then windowLen += 1
                    TargetPos = CurrentPos + windowLen
                    RangeReqd = windowLen
                End If

                ' check new settings and if any out of range then put them back
                If (CurrentPos < 1 Or TargetPos > EndRange Or RangeReqd < 50) Then
                    CurrentPos = CurrentPosSave
                    TargetPos = TargetPosSave
                    RangeReqd = RangeReqdSave
                    RangeRequired.Text = RangeReqdSave
                End If

                dataTable1.Clear()
                Chart2.Series(0).Points.Clear()
                Chart2.Series(1).Points.Clear()
                Chart2.Series(2).Points.Clear()
                Chart2.Series(3).Points.Clear()
                'Chart2.Series(4).Points.Clear()

                CurrentPosition.Text = CurrentPos
                TargetPosition.Text = TargetPos
                RangeRequired.Text = RangeReqd

                ' Print Xscale to chart
                PrintXscale()

                ' Pull in batch of RangeReqd lines
                For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
                    If line.Contains("//") Then
                        ' Skip lines containing "//"
                        Continue For
                    End If
                    AddPlaybackCSVRow(line)
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterShortTermMeanDevice1()
                FilterShortTermMeanDevice2()
                FilterTempDevice1()
                FilterHumDevice1()
                GeneratePPMColumn()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()
                UpdatePlaybackStatsSeries()

                'Get max and min values of Dev1 & Dev2, keep whichever is max/min value and use for setting scale
                GetMinMaxScales()

                DevicesMinMax()

                FixTicks()

                AverageNoise()

            End If
        End If

    End Sub


    Private Sub ButtonZoomOut_Click(sender As Object, e As EventArgs) Handles ButtonZoomOut.Click

        If ChartLoaded = True Then

            If (CSVfileok = True) Then

                ' Save off settings and calculate new zoom out settings
                CurrentPosSave = CurrentPos
                TargetPosSave = TargetPos
                RangeReqdSave = RangeReqd

                CentreRange = CurrentPos + (RangeReqd / 2)  ' current centre position
                CurrentPos = CentreRange - RangeReqd
                TargetPos = CentreRange + RangeReqd

                'RangeReqd = RangeReqd * 2
                RangeReqd *= 2
                RangeRequired.Text = RangeReqd * 2

                ' For dual-device CSVs, each sample is TWO consecutive
                ' lines (Dev1 then Dev2). Keep the window aligned to
                ' whole pairs so zooming never splits a pair - otherwise
                ' the two devices end up with mismatched sample counts
                ' and drift out of alignment with each other.
                If DualDev = True Then
                    If CurrentPos Mod 2 <> 0 Then CurrentPos -= 1
                    Dim windowLen As Integer = TargetPos - CurrentPos
                    If windowLen Mod 2 <> 0 Then windowLen += 1
                    TargetPos = CurrentPos + windowLen
                    RangeReqd = windowLen
                End If

                ' check new settings and if any out of range then put them back
                'If (CurrentPos < 1 Or TargetPos > EndRange Or RangeReqd > EndRange) Then
                'CurrentPos = CurrentPosSave
                'TargetPos = TargetPosSave
                'RangeReqd = RangeReqdSave
                'RangeRequired.Text = RangeReqdSave
                'End If

                ' check new settings and if any out of range
                If (CurrentPos < 1 Or TargetPos > EndRange Or RangeReqd > EndRange) Then

                    'RefreshPlaybackCSVFile()        ' Refresh entire chart since we've zooomed out fully
                    'Exit Sub

                    CurrentPos = 0
                    TargetPos = numberlinesCSV
                    RangeReqd = numberlinesCSV
                    RangeRequired.Text = numberlinesCSV

                    CurrentPosition.Text = CurrentPos
                    TargetPosition.Text = TargetPos
                    RangeRequired.Text = RangeReqd
                End If


                dataTable1.Clear()
                Chart2.Series(0).Points.Clear()
                Chart2.Series(1).Points.Clear()
                Chart2.Series(2).Points.Clear()
                Chart2.Series(3).Points.Clear()
                'Chart2.Series(4).Points.Clear()

                CurrentPosition.Text = CurrentPos
                TargetPosition.Text = TargetPos
                RangeRequired.Text = RangeReqd

                ' Print Xscale to chart
                'PrintXscale()

                ' Pull in batch of RangeReqd lines
                'For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
                'dataTable1.Rows.Add(line.Split(CSVdelimit))
                'Next
                For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
                    If line.Contains("//") Then
                        ' Skip lines containing "//"
                        Continue For
                    End If
                    AddPlaybackCSVRow(line)
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterShortTermMeanDevice1()
                FilterShortTermMeanDevice2()
                FilterTempDevice1()
                FilterHumDevice1()
                GeneratePPMColumn()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()
                UpdatePlaybackStatsSeries()

                PrintXscale()
                GetMinMaxScales()                  'Get max and min values of Dev1 & Dev2, keep whichever is max/min value and use for setting scale
                DevicesMinMax()
                FixTicks()
                AverageNoise()

            End If
        End If


    End Sub

    Private Sub ButtonShowAll_Click(sender As Object, e As EventArgs) Handles ButtonDisplayAll.Click

        ShowAll()
        'RefreshPlaybackCSVFile()

    End Sub

    Private Sub ButtonScrollRight_Click(sender As Object, e As EventArgs) Handles ButtonScrollRight.Click

        If ChartLoaded = True Then

            If (CSVfileok = True And RangeRequired.Text <> numberlinesCSV) Then      ' CSV ok and also only if graph is not full screen

                If RangeRequired.Text > numberlinesCSV / 2 Then      ' must be less than half the entire data range
                    RangeRequired.Text = numberlinesCSV / 2
                    RangeReqd = numberlinesCSV / 2
                Else
                    RangeReqd = RangeRequired.Text
                End If

                ' Temp override the above for testing
                RangeReqd = RangeRequired.Text
                If (RangeReqd > numberlinesCSV) Then
                    RangeReqd = numberlinesCSV
                    RangeRequired.Text = numberlinesCSV
                Else
                    RangeRequired.Text = RangeReqd
                End If

                dataTable1.Clear()
                Chart2.Series(0).Points.Clear()
                Chart2.Series(1).Points.Clear()
                Chart2.Series(2).Points.Clear()
                Chart2.Series(3).Points.Clear()
                'Chart2.Series(4).Points.Clear()

                ' Update new start position var
                'CurrentPos = CurrentPos + RangeReqd
                CurrentPos += RangeReqd

                ' Check that lines are available in current CSV file
                If CurrentPos < EndRange Then
                    TargetPos = CurrentPos + RangeReqd
                    If TargetPos > EndRange Then
                        TargetPos = EndRange
                        CurrentPos = EndRange - RangeReqd
                    End If
                Else
                    ' just make the range whats left to display however small
                    'CurrentPos = CurrentPos - RangeReqd
                    CurrentPos -= RangeReqd
                    TargetPos = EndRange
                End If

                ' Print Xscale to chart
                PrintXscale()

                CurrentPosition.Text = CurrentPos
                TargetPosition.Text = TargetPos
                RangeRequired.Text = RangeReqd

                ' Pull in batch of RangeReqd lines
                For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
                    If line.Contains("//") Then
                        ' Skip lines containing "//"
                        Continue For
                    End If
                    AddPlaybackCSVRow(line)
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterShortTermMeanDevice1()
                FilterShortTermMeanDevice2()
                FilterTempDevice1()
                FilterHumDevice1()
                GeneratePPMColumn()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()
                UpdatePlaybackStatsSeries()

                'Get max and min values of Dev1 & Dev2, keep whichever is max/min value and use for setting scale
                GetMinMaxScales()

                DevicesMinMax()

                FixTicks()

                AverageNoise()

            End If
        End If

    End Sub

    Private Sub ButtonScrollRightSMALL_Click(sender As Object, e As EventArgs) Handles ButtonScrollRightSMALL.Click

        If ChartLoaded = True Then

            If (CSVfileok = True And RangeRequired.Text <> numberlinesCSV) Then      ' CSV ok and also only if graph is not full screen

                If RangeRequired.Text > numberlinesCSV / 2 Then      ' must be less than half the entire data range
                    RangeRequired.Text = numberlinesCSV / 2
                    RangeReqd = numberlinesCSV / 2
                Else
                    RangeReqd = RangeRequired.Text
                End If

                ' Temp override the above for testing
                RangeReqd = RangeRequired.Text
                If (RangeReqd > numberlinesCSV) Then
                    RangeReqd = numberlinesCSV
                    RangeRequired.Text = numberlinesCSV
                Else
                    RangeRequired.Text = RangeReqd
                End If

                dataTable1.Clear()
                Chart2.Series(0).Points.Clear()
                Chart2.Series(1).Points.Clear()
                Chart2.Series(2).Points.Clear()
                Chart2.Series(3).Points.Clear()
                'Chart2.Series(4).Points.Clear()

                ' Update new start & End positions vars
                'CurrentPos = CurrentPos + (RangeReqd / 10)
                CurrentPos += (RangeReqd / 10)
                TargetPos = CurrentPos + (RangeReqd / 10)

                ' Check that lines are available in current CSV file
                If CurrentPos < EndRange Then
                    TargetPos = CurrentPos + RangeReqd
                    If TargetPos > EndRange Then
                        TargetPos = EndRange
                        CurrentPos = EndRange - RangeReqd
                    End If
                Else
                    ' just make the range whats left to display however small
                    'CurrentPos = CurrentPos - RangeReqd
                    CurrentPos -= RangeReqd
                    TargetPos = EndRange
                End If

                ' Print Xscale to chart
                PrintXscale()

                CurrentPosition.Text = CurrentPos
                TargetPosition.Text = TargetPos
                RangeRequired.Text = RangeReqd

                ' Pull in batch of RangeReqd lines
                For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
                    If line.Contains("//") Then
                        ' Skip lines containing "//"
                        Continue For
                    End If
                    AddPlaybackCSVRow(line)
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterShortTermMeanDevice1()
                FilterShortTermMeanDevice2()
                FilterTempDevice1()
                FilterHumDevice1()
                GeneratePPMColumn()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()
                UpdatePlaybackStatsSeries()

                'Get max and min values of Dev1 & Dev2, keep whichever is max/min value and use for setting scale
                GetMinMaxScales()

                DevicesMinMax()

                FixTicks()

                AverageNoise()

            End If
        End If

    End Sub


    Private Sub ButtonScrollLeft_Click(sender As Object, e As EventArgs) Handles ButtonScrollLeft.Click

        If ChartLoaded = True Then

            If (CSVfileok = True And RangeRequired.Text <> numberlinesCSV) Then     ' CSV ok and also only if graph is not full screen

                If RangeRequired.Text > numberlinesCSV / 2 Then      ' must be less than half the entire data range
                    RangeRequired.Text = numberlinesCSV / 2
                    RangeReqd = numberlinesCSV / 2
                Else
                    RangeReqd = RangeRequired.Text
                End If

                ' Temp override the above for testing
                RangeReqd = RangeRequired.Text
                If (RangeReqd > numberlinesCSV) Then
                    RangeReqd = numberlinesCSV
                    RangeRequired.Text = numberlinesCSV
                Else
                    RangeRequired.Text = RangeReqd
                End If

                dataTable1.Clear()
                Chart2.Series(0).Points.Clear()
                Chart2.Series(1).Points.Clear()
                Chart2.Series(2).Points.Clear()
                Chart2.Series(3).Points.Clear()
                'Chart2.Series(4).Points.Clear()

                ' Check that lines are available in current CSV
                TargetPos = CurrentPos
                CurrentPos = TargetPos - RangeReqd
                If CurrentPos < 0 Then
                    CurrentPos = 0
                    TargetPos = CurrentPos + RangeReqd
                End If

                ' Print Xscale to chart
                PrintXscale()

                CurrentPosition.Text = CurrentPos
                TargetPosition.Text = TargetPos
                RangeRequired.Text = RangeReqd

                ' Pull in batch of RangeReqd lines
                For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
                    If line.Contains("//") Then
                        ' Skip lines containing "//"
                        Continue For
                    End If
                    AddPlaybackCSVRow(line)
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterShortTermMeanDevice1()
                FilterShortTermMeanDevice2()
                FilterTempDevice1()
                FilterHumDevice1()
                GeneratePPMColumn()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()
                UpdatePlaybackStatsSeries()

                'Get max and min values of Dev1 & Dev2, keep whichever is max/min value and use for setting scale
                GetMinMaxScales()

                DevicesMinMax()

                FixTicks()

                AverageNoise()

            End If

        End If

    End Sub

    Private Sub ButtonScrollLeftSMALL_Click(sender As Object, e As EventArgs) Handles ButtonScrollLeftSMALL.Click

        If ChartLoaded = True Then

            If (CSVfileok = True And RangeRequired.Text <> numberlinesCSV) Then     ' CSV ok and also only if graph is not full screen

                If RangeRequired.Text > numberlinesCSV / 2 Then      ' must be less than half the entire data range
                    RangeRequired.Text = numberlinesCSV / 2
                    RangeReqd = numberlinesCSV / 2
                Else
                    RangeReqd = RangeRequired.Text
                End If

                ' Temp override the above for testing
                RangeReqd = RangeRequired.Text
                If (RangeReqd > numberlinesCSV) Then
                    RangeReqd = numberlinesCSV
                    RangeRequired.Text = numberlinesCSV
                Else
                    RangeRequired.Text = RangeReqd
                End If

                dataTable1.Clear()
                Chart2.Series(0).Points.Clear()
                Chart2.Series(1).Points.Clear()
                Chart2.Series(2).Points.Clear()
                Chart2.Series(3).Points.Clear()
                'Chart2.Series(4).Points.Clear()

                ' Update new start & End positions vars
                'CurrentPos = CurrentPos - (RangeReqd / 10)
                CurrentPos -= (RangeReqd / 10)
                'TargetPos = TargetPos - (RangeReqd / 10)
                TargetPos -= (RangeReqd / 10)

                ' Check that lines are available in current CSV
                'TargetPos = CurrentPos
                'CurrentPos = TargetPos - RangeReqd
                If CurrentPos < 0 Then
                    CurrentPos = 0
                    TargetPos = CurrentPos + RangeReqd
                End If

                ' Print Xscale to chart
                PrintXscale()

                CurrentPosition.Text = CurrentPos
                TargetPosition.Text = TargetPos
                RangeRequired.Text = RangeReqd

                ' Pull in batch of RangeReqd lines
                For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
                    If line.Contains("//") Then
                        ' Skip lines containing "//"
                        Continue For
                    End If
                    AddPlaybackCSVRow(line)
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterShortTermMeanDevice1()
                FilterShortTermMeanDevice2()
                FilterTempDevice1()
                FilterHumDevice1()
                GeneratePPMColumn()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()
                UpdatePlaybackStatsSeries()

                'Get max and min values of Dev1 & Dev2, keep whichever is max/min value and use for setting scale
                GetMinMaxScales()

                DevicesMinMax()

                FixTicks()

                AverageNoise()

            End If

        End If

    End Sub

    Private Sub ButtonShiftUp_Click(sender As Object, e As EventArgs) Handles ButtonShiftUp.Click

        If ChartLoaded = True Then

            If (CSVfileok = True) Then

                Dim Playbacknewmax As Double = ParseInvariantDouble(YaxisMaximum.Text)
                Dim Playbacknewmin As Double = ParseInvariantDouble(YaxisMinimum.Text)

                'Playbacknewmax = Playbacknewmax + ((Playbacknewmax - Playbacknewmin) / 20)
                Playbacknewmax += (Playbacknewmax - Playbacknewmin) / 20
                YaxisMaximum.Text = Playbacknewmax.ToString(Globalization.CultureInfo.InvariantCulture)
                YaxisMaximum.Text = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7).ToString(Globalization.CultureInfo.InvariantCulture)

                'Playbacknewmin = Playbacknewmin + ((Playbacknewmax - Playbacknewmin) / 20)
                Playbacknewmin += (Playbacknewmax - Playbacknewmin) / 20
                YaxisMinimum.Text = Playbacknewmin.ToString(Globalization.CultureInfo.InvariantCulture)
                YaxisMinimum.Text = Math.Round((ParseInvariantDouble(YaxisMinimum.Text)), 7).ToString(Globalization.CultureInfo.InvariantCulture)

                Dim result As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
                YaxisPerDiv.Text = result.ToString("#0.000000000")


                dataTable1.Clear()
                Chart2.Series(0).Points.Clear()
                Chart2.Series(1).Points.Clear()
                Chart2.Series(2).Points.Clear()
                Chart2.Series(3).Points.Clear()
                'Chart2.Series(4).Points.Clear()

                ' Pull in batch of RangeReqd lines
                For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
                    If line.Contains("//") Then
                        ' Skip lines containing "//"
                        Continue For
                    End If
                    AddPlaybackCSVRow(line)
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterShortTermMeanDevice1()
                FilterShortTermMeanDevice2()
                FilterTempDevice1()
                FilterHumDevice1()
                GeneratePPMColumn()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()
                UpdatePlaybackStatsSeries()

                'Manual device scale setting
                Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((ParseInvariantDouble(YaxisMinimum.Text)), 7)
                Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)  ' was 7
                Dim intervalY As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 20
                If intervalY <= 0 Then intervalY = 0.0000001   ' avoid MSChart crash when Y-max = Y-min (flat/no-variance data)
                Chart2.ChartAreas(0).AxisY.Interval = intervalY

                YaxisCheck1.Checked = False
                YaxisCheck2.Checked = False
                YaxisCheck3.Checked = False
                YaxisCheck4.Checked = False

            End If
        End If

    End Sub

    Private Sub ButtonShiftDn_Click(sender As Object, e As EventArgs) Handles ButtonShiftDn.Click

        If ChartLoaded = True Then

            If (CSVfileok = True) Then

                Dim Playbacknewmax As Double = ParseInvariantDouble(YaxisMaximum.Text)
                Dim Playbacknewmin As Double = ParseInvariantDouble(YaxisMinimum.Text)

                'Playbacknewmax = Playbacknewmax - ((Playbacknewmax - Playbacknewmin) / 20)
                Playbacknewmax -= (Playbacknewmax - Playbacknewmin) / 20
                YaxisMaximum.Text = Playbacknewmax.ToString(Globalization.CultureInfo.InvariantCulture)
                YaxisMaximum.Text = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7).ToString(Globalization.CultureInfo.InvariantCulture)

                'Playbacknewmin = Playbacknewmin - ((Playbacknewmax - Playbacknewmin) / 20)
                Playbacknewmin -= (Playbacknewmax - Playbacknewmin) / 20
                'If (Playbacknewmin < 0) Then
                'Playbacknewmin = 0
                'End If
                YaxisMinimum.Text = Playbacknewmin.ToString(Globalization.CultureInfo.InvariantCulture)
                YaxisMinimum.Text = Math.Round((ParseInvariantDouble(YaxisMinimum.Text)), 7).ToString(Globalization.CultureInfo.InvariantCulture)

                Dim result As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
                YaxisPerDiv.Text = result.ToString("#0.000000000")


                dataTable1.Clear()
                Chart2.Series(0).Points.Clear()
                Chart2.Series(1).Points.Clear()
                Chart2.Series(2).Points.Clear()
                Chart2.Series(3).Points.Clear()
                'Chart2.Series(4).Points.Clear()

                ' Pull in batch of RangeReqd lines
                For Each line As String In System.IO.File.ReadLines(filePlayback).Skip(CurrentPos).Take(TargetPos - CurrentPos)
                    If line.Contains("//") Then
                        ' Skip lines containing "//"
                        Continue For
                    End If
                    AddPlaybackCSVRow(line)
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterShortTermMeanDevice1()
                FilterShortTermMeanDevice2()
                FilterTempDevice1()
                FilterHumDevice1()
                GeneratePPMColumn()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()
                UpdatePlaybackStatsSeries()

                'Manual device scale setting
                Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((ParseInvariantDouble(YaxisMinimum.Text)), 7)
                Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)   ' was 7
                Dim intervalY As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 20
                If intervalY <= 0 Then intervalY = 0.0000001   ' avoid MSChart crash when Y-max = Y-min (flat/no-variance data)
                Chart2.ChartAreas(0).AxisY.Interval = intervalY

                YaxisCheck1.Checked = False
                YaxisCheck2.Checked = False
                YaxisCheck3.Checked = False
                YaxisCheck4.Checked = False

            End If

        End If

    End Sub

    Private Sub Yscaletidy()

        ' Tidy up Y-scale annotations on graph in order to keep length same irrespective of numerical data and No. DP's
        If CheckBoxYscaletidy.Checked = True Then
            Dim x1 As String = CStr(YaxisMaximum.Text)   ' 0.9999995 or 999.0000000 etc
            Dim x2 As String = CStr(YaxisMinimum.Text)   ' 0.9999995 or 999.0000000 etc

            Dim CountMaxAfter = x1.Length - InStr(x1, ".")    ' after DP     5.0000165 would give 7, 999.95606 would give 5
            Dim CountMinAfter = x2.Length - InStr(x2, ".")    ' after DP

            Dim CountMaxBefore = YaxisMaximum.TextLength - CountMaxAfter - 1 ' before DP     5.0000165 would give 1, 999.95606 would give 3
            Dim CountMinBefore = YaxisMinimum.TextLength - CountMinAfter - 1 ' before DP

            ' test because CountMinBefore was coming in as 0, CountMaxBefore as 2......not sure why but only when setting a manual Y-axis scale!
            If (CountMinBefore = 0) And CountMaxBefore = 1 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{0.0000000000}"  ' 0.999999599
            End If
            If (CountMinBefore = 0) And CountMaxBefore = 2 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{00.000000000}"  ' 00.999999599
            End If
            If (CountMinBefore = 0) And CountMaxBefore = 3 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{000.00000000}"  ' 000.99999999
            End If
            If (CountMinBefore = 0) And CountMaxBefore = 4 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{0000.0000000}"  ' 0000.9999999
            End If
            If (CountMinBefore = 0) And CountMaxBefore = 5 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{00000.000000}"  ' 00000.999999
            End If
            If (CountMinBefore = 0) And CountMaxBefore = 6 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{000000.00000}"  ' 000000.99999
            End If
            If (CountMinBefore = 0) And CountMaxBefore = 7 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{0000000.0000}"  ' 0000000.9999
            End If
            If (CountMinBefore = 0) And CountMaxBefore = 9 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{00000000.000}"  ' 00000000.999
            End If
            If (CountMinBefore = 0) And CountMaxBefore = 10 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{000000000.00}"  ' 000000000.99
            End If


            ' original
            If CountMinBefore = 1 And CountMaxBefore = 1 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{0.0000000000}"  ' 0.999999599
            End If

            If (CountMinBefore = 1 Or CountMinBefore = 2) And CountMaxBefore = 2 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{00.000000000}"  ' 00.999999599
            End If

            If (CountMinBefore = 2 Or CountMinBefore = 3) And CountMaxBefore = 3 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{000.00000000}"  ' 000.99999999
            End If

            If (CountMinBefore = 3 Or CountMinBefore = 4) And CountMaxBefore = 4 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{0000.0000000}"  ' 0000.9999999
            End If

            If (CountMinBefore = 4 Or CountMinBefore = 5) And CountMaxBefore = 5 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{00000.000000}"  ' 00000.999999
            End If

            If (CountMinBefore = 5 Or CountMinBefore = 6) And CountMaxBefore = 6 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{000000.00000}"  ' 000000.99999
            End If

            If (CountMinBefore = 6 Or CountMinBefore = 7) And CountMaxBefore = 7 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{0000000.0000}"  ' 0000000.9999
            End If

            If (CountMinBefore = 7 Or CountMinBefore = 8) And CountMaxBefore = 8 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{00000000.000}"  ' 00000000.999
            End If

            If (CountMinBefore = 8 Or CountMinBefore = 9) And CountMaxBefore = 9 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{000000000.00}"  ' 000000000.99
            End If

            If (CountMinBefore = 9 Or CountMinBefore = 10) And CountMaxBefore = 10 Then
                Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{0000000000.0}"  ' 0000000000.
            End If

        Else
            Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{000.0000000}"     ' default as set up top by default
        End If

        'Chart2.Width = 1481


    End Sub

    ' Manually adjust Y value - max down
    Private Sub ButtonYmaxDec_Click(sender As Object, e As EventArgs) Handles ButtonYmaxDec.Click
        Ymin = ParseInvariantDouble(YaxisMinimum.Text)
        Ymin += 0.00000000001
        Ymax = ParseInvariantDouble(YaxisMaximum.Text)
        Ymax -= (Ymax - Ymin) / 10  ' shift by a tenth
        'If (Ymin < 0) Then
        'Ymin = 0.0000001
        'End If
        'If (Ymin < Ymax And Ymin >= 0) Then
        If (Ymin < Ymax) Then
            YaxisMinimum.Text = Ymin.ToString("#0.00000000", Globalization.CultureInfo.InvariantCulture)
            YaxisMaximum.Text = Ymax.ToString("#0.00000000", Globalization.CultureInfo.InvariantCulture)

            Dim result As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")

            Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((ParseInvariantDouble(YaxisMinimum.Text)), 7)
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)   ' was 7
            Dim intervalY As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 20
            If intervalY <= 0 Then intervalY = 0.0000001   ' avoid MSChart crash when Y-max = Y-min (flat/no-variance data)
            Chart2.ChartAreas(0).AxisY.Interval = intervalY
        End If

        YaxisCheck1.Checked = False
        YaxisCheck2.Checked = False
        YaxisCheck3.Checked = False
        YaxisCheck4.Checked = False
    End Sub


    ' Manually adjust Y value - max up
    Private Sub ButtonYmaxInc_Click(sender As Object, e As EventArgs) Handles ButtonYmaxInc.Click
        Ymin = ParseInvariantDouble(YaxisMinimum.Text)
        Ymin += 0.00000000001
        Ymax = ParseInvariantDouble(YaxisMaximum.Text)
        Ymax += (Ymax - Ymin) / 10  ' shift by a tenth

        '        If (Ymin < 0) Then
        '        Ymin = 0.0000001
        '        End If
        'If (Ymin < Ymax And Ymin >= 0) Then
        If (Ymin < Ymax) Then
            YaxisMinimum.Text = Ymin.ToString("#0.00000000", Globalization.CultureInfo.InvariantCulture)
            YaxisMaximum.Text = Ymax.ToString("#0.00000000", Globalization.CultureInfo.InvariantCulture)

            Dim result As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")

            Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((ParseInvariantDouble(YaxisMinimum.Text)), 7)
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)   ' was 7
            Dim intervalY As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 20
            If intervalY <= 0 Then intervalY = 0.0000001   ' avoid MSChart crash when Y-max = Y-min (flat/no-variance data)
            Chart2.ChartAreas(0).AxisY.Interval = intervalY
        End If

        YaxisCheck1.Checked = False
        YaxisCheck2.Checked = False
        YaxisCheck3.Checked = False
        YaxisCheck4.Checked = False
    End Sub


    ' Manually adjust Y value - min down
    Private Sub ButtonYminDec_Click(sender As Object, e As EventArgs) Handles ButtonYminDec.Click
        Ymin = ParseInvariantDouble(YaxisMinimum.Text)
        Ymax = ParseInvariantDouble(YaxisMaximum.Text)
        Ymin -= (Ymax - Ymin) / 10  ' shift by a tenth

        '       If (Ymin < 0) Then
        '       Ymin = 0.0000001
        '       End If
        'If (Ymin < Ymax And Ymin >= 0) Then
        If (Ymin < Ymax) Then
            YaxisMinimum.Text = Ymin.ToString("#0.00000000", Globalization.CultureInfo.InvariantCulture)
            YaxisMaximum.Text = Ymax.ToString("#0.00000000", Globalization.CultureInfo.InvariantCulture)

            Dim result As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000") '

            Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((ParseInvariantDouble(YaxisMinimum.Text)), 7)
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)   ' was 7
            Dim intervalY As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 20
            If intervalY <= 0 Then intervalY = 0.0000001   ' avoid MSChart crash when Y-max = Y-min (flat/no-variance data)
            Chart2.ChartAreas(0).AxisY.Interval = intervalY
        End If

        YaxisCheck1.Checked = False
        YaxisCheck2.Checked = False
        YaxisCheck3.Checked = False
        YaxisCheck4.Checked = False
    End Sub


    ' Manually adjust Y value - max up
    Private Sub ButtonYminInc_Click(sender As Object, e As EventArgs) Handles ButtonYminInc.Click
        Ymin = ParseInvariantDouble(YaxisMinimum.Text)
        Ymax = ParseInvariantDouble(YaxisMaximum.Text)
        Ymin += (Ymax - Ymin) / 10  ' shift by a tenth
        '        If (Ymin < 0) Then
        '        Ymin = 0.0000001
        '        End If
        'If (Ymin < Ymax And Ymin >= 0) Then
        If (Ymin < Ymax) Then
            YaxisMinimum.Text = Ymin.ToString("#0.0000000000", Globalization.CultureInfo.InvariantCulture)
            YaxisMaximum.Text = Ymax.ToString("#0.0000000000", Globalization.CultureInfo.InvariantCulture)

            Dim result As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")

            Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((ParseInvariantDouble(YaxisMinimum.Text)), 7)
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)   ' was 7
            Dim intervalY As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 20
            If intervalY <= 0 Then intervalY = 0.0000001   ' avoid MSChart crash when Y-max = Y-min (flat/no-variance data)
            Chart2.ChartAreas(0).AxisY.Interval = intervalY
        End If

        YaxisCheck1.Checked = False
        YaxisCheck2.Checked = False
        YaxisCheck3.Checked = False
        YaxisCheck4.Checked = False
    End Sub


    Private Sub ButtonSaveSettings_Click(sender As Object, e As EventArgs) Handles ButtonSaveSettings.Click

        My.Settings.data20 = ChartScaleMax.Text
        My.Settings.data21 = ChartScaleMin.Text
        My.Settings.data24 = YaxisMaximum.Text
        My.Settings.data25 = YaxisMinimum.Text
        My.Settings.data26 = MedianValue.Text
        My.Settings.data27 = MedianTemp.Text
        My.Settings.data28 = PPMscalerangeentry.Text

    End Sub


    ' Show value of point on graph by mouse hover
    Private Sub Chart2_GetToolTipText(sender As Object, e As ToolTipEventArgs) Handles Chart2.GetToolTipText

        If (CheckBoxToolTips.Checked = True) Then
            Chart2.Series(0).ToolTip = "#VAL{0.00000000}"   ' dev 1
            Chart2.Series(1).ToolTip = "#VAL{0.00000000}"   ' dev 2
            Chart2.Series(2).ToolTip = "#VAL{0.0}"          ' temperature
            Chart2.Series(3).ToolTip = "#VAL{0.0}"          ' humidity
            'Chart2.Series(4).ToolTip = "#VAL{0.00}"         ' PPM
        Else
            Chart2.Series(0).ToolTip = ""   ' dev 1
            Chart2.Series(1).ToolTip = ""   ' dev 2
            Chart2.Series(2).ToolTip = ""   ' temperature
            Chart2.Series(3).ToolTip = ""   ' humidity
            'Chart2.Series(4).ToolTip = ""   ' PPM
        End If

    End Sub


    Private Sub CheckBoxPPMenable_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxPPMenable.CheckedChanged

        If (RadioButtonPPMTempo.Checked = True) Then
            MedianTemp.Enabled = True
            MedianTempText.Enabled = True
            CheckBoxMedianT.Enabled = True
        Else
            MedianTemp.Enabled = False
            MedianTempText.Enabled = False
            CheckBoxMedianT.Enabled = False
        End If

        If CheckBoxPPMenable.Checked = True Then
            'RefreshChart.Enabled = True
            RadioButtonPPMDev.Enabled = True
            RadioButtonPPMTempo.Enabled = True
            MedianValue.Enabled = True
            RadioButtonDev1.Enabled = True
            RadioButtonDev2.Enabled = True
            MedianValueText.Enabled = True
            PPMscalerangeentry.Enabled = True
            PPMscaleText.Enabled = True
            CheckBoxMedianV.Enabled = True
            ' Scroll/Zoom/Y-adjust/Shift used to be disabled here
            ' because enabling PPM mode never regenerated PPM values
            ' for a reloaded subset (see GeneratePPMColumn()) - now
            ' that every reload regenerates PPM correctly, these stay
            ' enabled while PPM is on.
        Else
            CheckBoxMedianT.Enabled = False
            MedianTempText.Enabled = False
            MedianTemp.Enabled = False
            'RefreshChart.Enabled = False
            RadioButtonPPMDev.Enabled = False
            RadioButtonPPMTempo.Enabled = False
            MedianValue.Enabled = False
            RadioButtonDev1.Enabled = False
            RadioButtonDev2.Enabled = False
            MedianValueText.Enabled = False
            PPMscalerangeentry.Enabled = False
            PPMscaleText.Enabled = False
            CheckBoxMedianV.Enabled = False
            ButtonScrollLeft.Enabled = True
            ButtonScrollRight.Enabled = True
            ButtonScrollLeftSMALL.Enabled = True
            ButtonScrollRightSMALL.Enabled = True
            ButtonZoomIn.Enabled = True
            ButtonZoomOut.Enabled = True
            ButtonYminInc.Enabled = True
            ButtonYminDec.Enabled = True
            ButtonYmaxInc.Enabled = True
            ButtonYmaxDec.Enabled = True
            ButtonDisplayAll.Enabled = True
            ButtonShiftUp.Enabled = True
            ButtonShiftDn.Enabled = True

            ' Erase scale
            Scale1.Text = ""
            Scale2.Text = ""
            Scale3.Text = ""
            Scale4.Text = ""
            Scale5.Text = ""
            Scale6.Text = ""
            Scale7.Text = ""
            Scale8.Text = ""
            Scale9.Text = ""
            Scale10.Text = ""
            Scale11.Text = ""
            Scale12.Text = ""
            Scale13.Text = ""
            Scale14.Text = ""
            Scale15.Text = ""
            Scale16.Text = ""
            Scale17.Text = ""
            Scale18.Text = ""
            Scale19.Text = ""
            Scale20.Text = ""
            Scale21.Text = ""
            Scale22.Text = ""
            Scale23.Text = ""
            Scale24.Text = ""
            Scale25.Text = ""

            LabelPPMtop.Visible = False
            LabelPPMdegctop.Visible = False

        End If

        'RefreshPlaybackCSVFile()

        If CheckBoxPPMenable.Checked = False Then
            Chart2.Series(4).Enabled = False
        Else
            Chart2.Series(4).Enabled = True
            RefreshPlaybackCSVFile()
        End If

    End Sub


    Private Sub RadioButtonPPMDev_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonPPMDev.CheckedChanged
        MedianTemp.Enabled = False
        MedianTempText.Enabled = False
        CheckBoxMedianT.Enabled = False
        RefreshPlaybackCSVFile()
    End Sub


    Private Sub RadioButtonPPMTempo_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonPPMTempo.CheckedChanged
        MedianTemp.Enabled = True
        MedianTempText.Enabled = True
        CheckBoxMedianT.Enabled = True
        RefreshPlaybackCSVFile()
    End Sub


    Private Sub FilterDeviceName1()

        ' Device 1
        If DEV1avg.Text = "0" Then

            ' Filter DeviceName1
            If (DeviceName1.Text <> "") Then
                Dim selectedRows() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")
                ''Add filtered data to series
                For Each dr As DataRow In selectedRows
                    Chart2.Series(0).Points.AddXY(dr("DEVICE"), dr("VALUE"))
                Next
            End If

        Else

            If (DeviceName1.Text <> "") Then

                DEV1rollingAverageValues.Clear()

                Dim selectedRows() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")

                ' Create a list to store the rolling average values
                'Dim Dev1rollingAverageValues As New List(Of Double)

                ' Iterate through the filtered data and calculate the rolling average for each data point
                For Each dr As DataRow In selectedRows

                    Dim variancevalue As Double = Convert.ToDouble(dr("VALUE"))

                    ' Call the CalculateRollingAverage function to get the rolling average
                    Dim Dev1rollingAverageValue As Double = CalculateRollingAverage(variancevalue, Val(DEV1avg.Text), Dev1AvgBuffer)

                    ' Store the rolling average value in the list
                    DEV1rollingAverageValues.Add(Dev1rollingAverageValue)

                    ' Add the data point with the rolling average value to the chart's series
                    Chart2.Series(0).Points.AddXY(dr("DEVICE"), Dev1rollingAverageValue)

                    'Console.WriteLine("Rolling Average Value for Device 1 " & Dev1rollingAverageValue)

                Next
            End If

        End If

    End Sub


    Private Sub FilterDeviceName2()

        If DEV2avg.Text = "0" Then

            ' Filter DeviceName2
            If (DeviceName2.Text <> "") Then
                Dim selectedRows2() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")
                'Add filtered data to series
                For Each dr As DataRow In selectedRows2
                    Chart2.Series(1).Points.AddXY(dr("DEVICE"), dr("VALUE"))
                Next
            End If

        Else

            If (DeviceName2.Text <> "") Then

                DEV2rollingAverageValues.Clear()

                Dim selectedRows2() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")

                ' Create a list to store the rolling average values
                'Dim Dev1rollingAverageValues As New List(Of Double)

                ' Iterate through the filtered data and calculate the rolling average for each data point
                For Each dr As DataRow In selectedRows2

                    Dim variancevalue As Double = Convert.ToDouble(dr("VALUE"))

                    ' Call the CalculateRollingAverage function to get the rolling average
                    Dim Dev2rollingAverageValue As Double = CalculateRollingAverage(variancevalue, Val(DEV2avg.Text), Dev2AvgBuffer)

                    ' Store the rolling average value in the list
                    DEV2rollingAverageValues.Add(Dev2rollingAverageValue)

                    ' Add the data point with the rolling average value to the chart's series
                    Chart2.Series(1).Points.AddXY(dr("DEVICE"), Dev2rollingAverageValue)

                    'Console.WriteLine("Rolling Average Value for Device 1 " & Dev1rollingAverageValue)

                Next
            End If

        End If


    End Sub


    Private Sub FilterShortTermMeanDevice1()

        Chart2.Series("Dev 1 Short-Term Mean").Points.Clear()

        If Not CheckPlaybackDev1ShortTermMean.Checked Then Exit Sub
        If (DeviceName1.Text = "") Then Exit Sub

        Dim selectedRows() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")

        Dim window As New Queue(Of Double)
        Dim windowSum As Double = 0.0

        For Each dr As DataRow In selectedRows

            Dim v As Double = Convert.ToDouble(dr("VALUE"))
            window.Enqueue(v) : windowSum += v
            If window.Count > ShortTermMeanWindow Then windowSum -= window.Dequeue()

            Chart2.Series("Dev 1 Short-Term Mean").Points.AddXY(dr("DEVICE"), windowSum / window.Count)

        Next

    End Sub


    Private Sub FilterShortTermMeanDevice2()

        Chart2.Series("Dev 2 Short-Term Mean").Points.Clear()

        If Not CheckPlaybackDev2ShortTermMean.Checked Then Exit Sub
        If (DeviceName2.Text = "") Then Exit Sub

        Dim selectedRows2() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")

        Dim window As New Queue(Of Double)
        Dim windowSum As Double = 0.0

        For Each dr As DataRow In selectedRows2

            Dim v As Double = Convert.ToDouble(dr("VALUE"))
            window.Enqueue(v) : windowSum += v
            If window.Count > ShortTermMeanWindow Then windowSum -= window.Dequeue()

            Chart2.Series("Dev 2 Short-Term Mean").Points.AddXY(dr("DEVICE"), windowSum / window.Count)

        Next

    End Sub


    Private Sub FilterTempDevice1()

        ' Unlike its sibling Filter*() functions, this one never cleared the
        ' series before repopulating - callers were relying on having
        ' already cleared Chart2.Series(2) themselves beforehand (e.g.
        ' ShowAll() does this explicitly). Any caller that doesn't do that
        ' (like TEMPavg's own TextChanged handler) ends up appending a full
        ' duplicate copy of the Temp trace on top of the existing points
        ' every time it runs, which is what was doubling the chart.
        Chart2.Series(2).Points.Clear()

        If PlaybackTemp.Checked = True Then

            If TEMPavg.Text = "0" Then

                ' Filter Temperature from Device 1
                If (PlaybackTemp.Checked = True) Then
                    Dim selectedRows3() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")
                    'Add filtered data to series

                    For Each dr As DataRow In selectedRows3
                        Chart2.Series(2).Points.AddXY(dr("DEVICE"), dr("TEMP"))
                    Next

                End If

            Else

                TEMProllingAverageValues.Clear()

                Dim selectedRows3() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")

                ' Iterate through the filtered data and calculate the rolling average for each data point
                For Each dr As DataRow In selectedRows3

                    Dim variancevalue As Double = Convert.ToDouble(dr("TEMP"))

                    ' Call the CalculateRollingAverage function to get the rolling average
                    Dim TemprollingAverageValue As Double = CalculateRollingAverage(variancevalue, Val(TEMPavg.Text), TempAvgBuffer)

                    ' Store the rolling average value in the list
                    TEMProllingAverageValues.Add(TemprollingAverageValue)

                    ' Add the data point with the rolling average value to the chart's series
                    Chart2.Series(2).Points.AddXY(dr("DEVICE"), TemprollingAverageValue)

                    'Console.WriteLine("Rolling Average Value for Temperature " & TemprollingAverageValue)

                Next

            End If

        End If

    End Sub


    Private Sub FilterHumDevice1()

        ' Same missing-clear bug FilterTempDevice1() had - callers were
        ' relying on having already cleared Chart2.Series(3) themselves
        ' beforehand. Clearing here makes this function self-contained and
        ' safe to call more than once per load, same as its siblings.
        Chart2.Series(3).Points.Clear()

        ' Filter Humidity from Device 1
        If (PlaybackHum.Checked = True) Then
            Dim selectedRows4() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")
            'Add filtered data to series
            For Each dr As DataRow In selectedRows4
                Chart2.Series(3).Points.AddXY(dr("DEVICE"), dr("HUM"))
            Next
        End If

    End Sub


    Private Sub GeneratePPMColumn()

        ' Computes the PPM column for whatever rows currently sit in
        ' dataTable1 - the initial full load, or a zoomed/scrolled/
        ' shifted subset. Callers are responsible for having already
        ' reloaded dataTable1 (and rebuilt the rolling-average lists
        ' via FilterDeviceName1/2 and FilterTempDevice1, which happens
        ' automatically since those always run before this is called)
        ' for whatever range is currently in view.

        If (CheckBoxPPMenable.Checked = False) Then Exit Sub

        ' Set PPM scale vars for calc
        Dim ppmscalerange As Double

        Dim ppmText As String =
PPMscalerangeentry.Text.Replace(vbCr, "").Replace(vbLf, "").Trim()

        If Not Double.TryParse(ppmText, ppmscalerange) Then

            ppmscalerange = 40
            PPMscalerangeentry.Text = "40"

        End If

        If ppmscalerange > 198 Then

            ppmscalerange = 198
            PPMscalerangeentry.Text = "198"

        ElseIf ppmscalerange < 0.2 Then

            ppmscalerange = 0.2
            PPMscalerangeentry.Text = "0.2"

        End If

        ppmscalerangebit = (ppmscalerange / 2) / 12


        ' Calculate PPM for given row
        ' Get initial value/Temp from CSV if selected
        If (CheckBoxMedianV.Checked = False) Then
            medianvalued = ParseInvariantDouble(MedianValue.Text)
        Else
            medianvalued = MedianValueCSV
            MedianValue.Text = MedianValueCSV.ToString(Globalization.CultureInfo.InvariantCulture)
        End If

        If (CheckBoxMedianT.Checked = False) Then
            mediantempd = ParseInvariantDouble(MedianTemp.Text)
        Else
            mediantempd = MedianTempCSV
            MedianTemp.Text = MedianTempCSV.ToString(Globalization.CultureInfo.InvariantCulture)
        End If

        Dim PPMdevice As String = ""
        Dim YaxisMaximumVal As Double = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)
        Dim YaxisMinimumVal As Double = Math.Round((ParseInvariantDouble(YaxisMinimum.Text)), 7)
        'Dim PPMscale As Double = ppmscalerange


        ' Get device from radio buttons
        If (RadioButtonDev1.Checked = True) Then
            PPMdevice = DeviceName1.Text
        End If
        If (RadioButtonDev2.Checked = True) Then
            PPMdevice = DeviceName2.Text
        End If




        ' Add PPM data to datatable - PPM Tempco calculation
        ' See https://www.allaboutcircuits.com/technical-articles/understanding-the-temperature-coefficient-of-a-voltage-reference/
        If (RadioButtonPPMTempo.Checked = True) Then


            If PPMdevice = DeviceName1.Text Then
                tempcounter = 0
                tempTEMPcounter = 0
            End If

            If PPMdevice = DeviceName2.Text Then
                tempcounter = 1
                tempTEMPcounter = 1
            End If


            ' loop
            For i = 0 To dataTable1.Rows.Count - 1
                If (dataTable1.Rows(i)("DEVICE")) = PPMdevice Then

                    Dim PPMdegCrollingAverageValue As Double
                    Dim TEMProllingAverageValue As Double

                    If PPMdevice = DeviceName1.Text Then
                        If DEV1avg.Text = "0" Then
                            ' Vars from CSV
                            variancevalue = Val((dataTable1.Rows(i)("VALUE")))

                            ' Use the corresponding rolling average value from the rollingAverageValues list
                            'variancevalue = DEV1rollingAverageValues(i)
                        Else
                            ' Use the rolling average value from the Dev1rollingAverageValue list
                            PPMdegCrollingAverageValue = DEV1rollingAverageValues(tempcounter)
                            'tempcounter = tempcounter + 1
                            tempcounter += 1
                            If tempcounter = DEV1rollingAverageValues.Count Then         ' protect counter overruning past last entry
                                'tempcounter = tempcounter - 1
                                tempcounter -= 1
                            End If
                        End If
                    End If


                    If PPMdevice = DeviceName2.Text Then
                        If DEV2avg.Text = "0" Then
                            ' Vars from CSV
                            variancevalue = Val((dataTable1.Rows(i)("VALUE")))

                            ' Use the corresponding rolling average value from the rollingAverageValues list
                            'variancevalue = DEV1rollingAverageValues(i)
                        Else
                            ' Use the rolling average value from the Dev1rollingAverageValue list
                            PPMdegCrollingAverageValue = DEV2rollingAverageValues(tempcounter)
                            'tempcounter = tempcounter + 1
                            tempcounter += 1
                            If tempcounter = DEV2rollingAverageValues.Count Then         ' protect counter overruning past last entry
                                'tempcounter = tempcounter - 1
                                tempcounter -= 1
                            End If
                        End If
                    End If


                    If PlaybackTemp.Checked = True Then

                        ' get Temp value either from csv data or from AVG list
                        If TEMPavg.Text = "0" Then
                            variancetemp = Val((dataTable1.Rows(i)("TEMP")))
                            TEMProllingAverageValue = variancetemp      ' this is the value that is used later
                        Else
                            ' Use the rolling average value from the TemprollingAverageValue list
                            TEMProllingAverageValue = TEMProllingAverageValues(tempTEMPcounter)
                            'variancetemp = TEMProllingAverageValues(tempTEMPcounter)
                            'tempTEMPcounter = tempTEMPcounter + 1
                            tempTEMPcounter += 1
                            If tempTEMPcounter = TEMProllingAverageValues.Count Then         ' protect counter overruning past last entry
                                'tempTEMPcounter = tempTEMPcounter - 1
                                tempTEMPcounter -= 1
                            End If
                        End If

                    End If


                    ' hack to compensate for DIV/0 problem. Slightly adjust the temperature!
                    If TEMProllingAverageValue = mediantempd Then       ' avoid DIV/0
                        'TEMProllingAverageValue = TEMProllingAverageValue + 0.0000001
                        TEMProllingAverageValue += 0.0000001
                    End If


                    If (TEMProllingAverageValue - mediantempd) <> 0 Then       ' avoid DIV/0

                        ' Tempco calc

                        If PPMdevice = DeviceName1.Text Then
                            If DEV1avg.Text = "0" Then
                                Vdiff = (variancevalue - medianvalued)
                            Else
                                Vdiff = (PPMdegCrollingAverageValue - medianvalued)
                            End If
                        End If

                        If PPMdevice = DeviceName2.Text Then
                            If DEV2avg.Text = "0" Then
                                Vdiff = (variancevalue - medianvalued)
                            Else
                                Vdiff = (PPMdegCrollingAverageValue - medianvalued)
                            End If
                        End If

                        ' so either Vdiff = (variancevalue - medianvalued)
                        ' or it's   Vdiff = (PPMdegCrollingAverageValue - medianvalued)
                        VnomTdiff = medianvalued * (TEMProllingAverageValue - mediantempd)
                        calcppmvalue = (Vdiff / VnomTdiff) * 1000000


                        ' limits of PPM scale - entered value = 40
                        If calcppmvalue > 99 Then
                            calcppmvalue = 99
                        End If
                        If calcppmvalue < -99 Then
                            calcppmvalue = -99
                        End If

                        ' Adjust position of PPM graph on chart to suit right hand PPM scale - a hack!
                        Dim offsetfactor As Double = ((YaxisMaximumVal - YaxisMinimumVal) / 2) + YaxisMinimumVal
                        Dim scalefactor As Double = ((YaxisMaximumVal - YaxisMinimumVal) / ppmscalerange)
                        'calcppmvalue = calcppmvalue * scalefactor
                        calcppmvalue *= scalefactor
                        'calcppmvalue = calcppmvalue + offsetfactor
                        calcppmvalue += offsetfactor

                        dataTable1.Rows(i)("PPM") = calcppmvalue

                    Else        ' force PPM/DegC to 0.0 if variance and median values are exactly the same

                        calcppmvalue = 0.00000001     ' protecting against DIV/0

                        ' Adjust position of PPM graph on chart to suit right hand PPM scale - a hack!
                        Dim offsetfactor As Double = ((YaxisMaximumVal - YaxisMinimumVal) / 2) + YaxisMinimumVal
                        Dim scalefactor As Double = ((YaxisMaximumVal - YaxisMinimumVal) / ppmscalerange)
                        'calcppmvalue = calcppmvalue * scalefactor
                        calcppmvalue *= scalefactor
                        'calcppmvalue = calcppmvalue + offsetfactor
                        calcppmvalue += offsetfactor

                        dataTable1.Rows(i)("PPM") = calcppmvalue

                    End If
                End If
            Next
        End If


        ' Add PPM data to datatable - PPM Deviation
        If (RadioButtonPPMDev.Checked = True) Then

            If PPMdevice = DeviceName1.Text Then
                tempcounter = 0
            End If

            If PPMdevice = DeviceName2.Text Then
                tempcounter = 1
            End If


            For i = 0 To dataTable1.Rows.Count - 1
                If (dataTable1.Rows(i)("DEVICE")) = PPMdevice Then

                    Dim DeviationrollingAverageValue As Double


                    If PPMdevice = DeviceName1.Text Then

                        If DEV1avg.Text = "0" Then
                            ' Vars from CSV
                            variancevalue = Val((dataTable1.Rows(i)("VALUE")))

                            ' PPM Deviation calc
                            Vchange = (variancevalue - medianvalued)
                            'calcppmvalue = Vchange * 1000000
                        Else
                            ' Use the rolling average value from the Dev1rollingAverageValue list
                            DeviationrollingAverageValue = DEV1rollingAverageValues(tempcounter)
                            'tempcounter = tempcounter + 1
                            tempcounter += 1
                            If tempcounter = DEV1rollingAverageValues.Count Then         ' protect counter overruning past last entry
                                'tempcounter = tempcounter - 1
                                tempcounter -= 1
                            End If

                        End If

                    End If


                    If PPMdevice = DeviceName2.Text Then

                        If DEV2avg.Text = "0" Then
                            ' Vars from CSV
                            variancevalue = Val((dataTable1.Rows(i)("VALUE")))

                        Else
                            ' Use the rolling average value from the Dev1rollingAverageValue list
                            DeviationrollingAverageValue = DEV2rollingAverageValues(tempcounter)
                            'tempcounter = tempcounter + 1
                            tempcounter += 1
                            If tempcounter = DEV2rollingAverageValues.Count Then         ' protect counter overruning past last entry
                                'tempcounter = tempcounter - 1
                                tempcounter -= 1
                            End If

                        End If

                    End If


                    If PPMdevice = DeviceName1.Text Then
                        If DEV1avg.Text = "0" Then
                            Vchange = (variancevalue - medianvalued)
                        Else
                            Vchange = (DeviationrollingAverageValue - medianvalued)
                        End If
                    End If


                    If PPMdevice = DeviceName2.Text Then
                        If DEV2avg.Text = "0" Then
                            Vchange = (variancevalue - medianvalued)
                        Else
                            Vchange = (DeviationrollingAverageValue - medianvalued)
                        End If
                    End If


                    ' PPM Deviation calc - Finish
                    'Vchange = (variancevalue - medianvalued)
                    'calcppmvalue = Vchange * 1000000
                    calcppmvalue = If(medianvalued <> 0, (Vchange / medianvalued) * 1000000, 0)


                    ' limits of PPM scale - entered value = 40
                    If calcppmvalue > 99 Then
                        calcppmvalue = 99
                    End If
                    If calcppmvalue < -99 Then
                        calcppmvalue = -99
                    End If

                    ' Adjust position of PPM graph on chart to suit right hand PPM scale - a hack!
                    Dim offsetfactor As Double = ((YaxisMaximumVal - YaxisMinimumVal) / 2) + YaxisMinimumVal
                    Dim scalefactor As Double = ((YaxisMaximumVal - YaxisMinimumVal) / ppmscalerange)
                    'calcppmvalue = calcppmvalue * scalefactor
                    calcppmvalue *= scalefactor
                    'calcppmvalue = calcppmvalue + offsetfactor
                    calcppmvalue += offsetfactor


                    dataTable1.Rows(i)("PPM") = calcppmvalue

                End If
            Next
        End If

    End Sub


    Private Sub FilterGenPPMDevice1()

        ' Filter Dev 1 generated PPM 
        If (CheckBoxPPMenable.Checked = True And RadioButtonDev1.Checked = True) Then
            Chart2.Series(4).Points.Clear()
            Dim selectedRows5() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")
            'Add filtered data to series
            For Each dr As DataRow In selectedRows5
                Chart2.Series(4).Points.AddXY(dr("DEVICE"), dr("PPM"))
            Next
        End If

    End Sub


    Private Sub FilterGenPPMDevice2()

        ' Filter Dev 2 generated PPM 
        If (CheckBoxPPMenable.Checked = True And RadioButtonDev2.Checked = True) Then
            Chart2.Series(4).Points.Clear()
            Dim selectedRows5() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")
            'Add filtered data to series
            For Each dr As DataRow In selectedRows5
                Chart2.Series(4).Points.AddXY(dr("DEVICE"), dr("PPM"))
            Next
        End If

    End Sub

    Sub DevicesMinMax()

        'Get max and min values of Dev1 & Dev2 for data range of each device
        Dev1MaxMin.Text = ""
        Dev2MaxMin.Text = ""

        maxValue = -10000000.0
        minValue = 10000000.0

        If (DeviceName1.Text <> "") Then
            Dim selectedRows() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")
            For Each dr As DataRow In selectedRows
                currentValue = dr("VALUE")
                If currentValue > maxValue Then maxValue = currentValue
            Next

            Dev1MaxD = maxValue

            For Each dr As DataRow In selectedRows
                currentValue = dr("VALUE")
                If currentValue < minValue Then minValue = currentValue
            Next

            Dev1MinD = minValue

            If (CheckX1000000.Checked = True Or CheckX1000.Checked = True) Then

                If (CheckX1000000.Checked = True) Then
                    Dev1MaxMin.Text = ((Dev1MaxD - Dev1MinD) * 1000000).ToString("0.#########")
                End If
                If (CheckX1000.Checked = True) Then
                    Dev1MaxMin.Text = ((Dev1MaxD - Dev1MinD) * 1000).ToString("0.#########")
                End If
            Else
                Dev1MaxMin.Text = CDec(Dev1MaxD - Dev1MinD).ToString("0.#########")     ' e-notation to decimal, max 9 DP
            End If
        End If

        maxValue = -10000000.0
        minValue = 10000000.0

        If (DeviceName2.Text <> "") Then
            Dim selectedRows() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")
            For Each dr As DataRow In selectedRows
                currentValue = dr("VALUE")
                If currentValue > maxValue Then maxValue = currentValue
            Next

            Dev2MaxD = maxValue

            For Each dr As DataRow In selectedRows
                currentValue = dr("VALUE")
                If currentValue < minValue Then minValue = currentValue
            Next

            Dev2MinD = minValue

            If (CheckX1000000.Checked = True Or CheckX1000.Checked = True) Then

                If (CheckX1000000.Checked = True) Then
                    Dev2MaxMin.Text = ((Dev2MaxD - Dev2MinD) * 1000000).ToString("0.#########")
                End If

                If (CheckX1000.Checked = True) Then
                    Dev2MaxMin.Text = ((Dev2MaxD - Dev2MinD) * 1000).ToString("0.#########")
                End If

            Else

                Dev2MaxMin.Text = CDec(Dev2MaxD - Dev2MinD).ToString("0.#########")     ' e-notation to decimal, max 9 DP

            End If

        End If

    End Sub












    Sub GetMinMaxScales()
        Dim axisYMin, axisYMax As Double
        Dim range As Double
        Dim interval As Double

        ' Calculate the axis Y minimum and maximum based on conditions
        If CheckBoxMaxMin.Checked Then
            axisYMax = Math.Round(CDbl(YmaxFromDT), 7)
            axisYMin = Math.Round(CDbl(YminFromDT), 7)

            Chart2.ChartAreas(0).AxisY.Maximum = axisYMax
            Chart2.ChartAreas(0).AxisY.Minimum = axisYMin
            YaxisMaximum.Text = axisYMax.ToString(Globalization.CultureInfo.InvariantCulture)
            YaxisMinimum.Text = axisYMin.ToString(Globalization.CultureInfo.InvariantCulture)

        Else
            axisYMax = Math.Round(ParseInvariantDouble(YaxisMaximum.Text), 7)
            axisYMin = Math.Round(ParseInvariantDouble(YaxisMinimum.Text), 7)
            range = axisYMax - axisYMin
            interval = range / 20

            Chart2.ChartAreas(0).AxisY.Maximum = axisYMax
            Chart2.ChartAreas(0).AxisY.Minimum = axisYMin
            Chart2.ChartAreas(0).AxisY.Interval = interval
            YaxisMaximum.Text = axisYMax.ToString(Globalization.CultureInfo.InvariantCulture)
            YaxisMinimum.Text = axisYMin.ToString(Globalization.CultureInfo.InvariantCulture)
        End If

        ' Update Y axis per division
        range = axisYMax - axisYMin
        YaxisPerDiv.Text = (range / 32).ToString("#0.000000000")

        ' Temp/Hum scale setting
        Dim tempHumMin, tempHumMax As Double
        tempHumMin = ParseInvariantDouble(ChartScaleMin.Text)
        tempHumMax = ParseInvariantDouble(ChartScaleMax.Text)
        interval = (tempHumMax - tempHumMin) / 32

        Chart2.ChartAreas(0).AxisY2.Minimum = tempHumMin
        Chart2.ChartAreas(0).AxisY2.Maximum = tempHumMax
        Chart2.ChartAreas(0).AxisY2.Interval = interval
        Chart2.ChartAreas(0).AxisY2.LabelStyle.Format = "00.0"
    End Sub









    Private Sub PlaybackTemp_CheckedChanged(sender As Object, e As EventArgs) Handles PlaybackTemp.CheckedChanged

        'RefreshPlaybackCSVFile()

        If PlaybackTemp.Checked = False Then
            Chart2.Series(2).Enabled = False
        Else
            Chart2.Series(2).Enabled = True
            RefreshPlaybackCSVFile()
        End If

    End Sub

    Private Sub PlaybackHum_CheckedChanged(sender As Object, e As EventArgs) Handles PlaybackHum.CheckedChanged

        'RefreshPlaybackCSVFile()

        If PlaybackHum.Checked = False Then
            Chart2.Series(3).Enabled = False
        Else
            Chart2.Series(3).Enabled = True
            RefreshPlaybackCSVFile()
        End If

    End Sub

    Private Sub LogYaxis_CheckedChanged(sender As Object, e As EventArgs)

        RefreshPlaybackCSVFile()

    End Sub

    Private Sub CheckBoxYscaletidy_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxYscaletidy.CheckedChanged

        ' Tidy up Y-scale annotations on graph in order to keep length same irrespective of numerical data and No. DP's
        Yscaletidy()

    End Sub



    Private Sub CheckDev1Line_CheckedChanged(sender As Object, e As EventArgs) Handles CheckDev1Line.CheckedChanged

        If CheckDev1Line.Checked = True Then
            CheckDev1Point.Checked = False
            System.Threading.Thread.Sleep(50)
            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Else
            CheckDev1Point.Checked = True
            System.Threading.Thread.Sleep(50)
            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
            Chart2.Series(0).MarkerStep = 1
        End If

    End Sub

    Private Sub CheckDev1Point_CheckedChanged(sender As Object, e As EventArgs) Handles CheckDev1Point.CheckedChanged

        If CheckDev1Point.Checked = True Then
            CheckDev1Line.Checked = False
            System.Threading.Thread.Sleep(50)
            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
            Chart2.Series(0).MarkerStep = 1
            Chart2.Series(0).MarkerSize = 2
        Else
            CheckDev1Line.Checked = True
            System.Threading.Thread.Sleep(50)
            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
        End If

    End Sub

    Private Sub CheckDev2Line_CheckedChanged(sender As Object, e As EventArgs) Handles CheckDev2Line.CheckedChanged

        If CheckDev2Line.Checked = True Then
            CheckDev2Point.Checked = False
            System.Threading.Thread.Sleep(50)
            Chart2.Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Else
            CheckDev2Point.Checked = True
            System.Threading.Thread.Sleep(50)
            Chart2.Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Point
            Chart2.Series(1).MarkerStep = 1
            Chart2.Series(1).MarkerSize = 2
        End If

    End Sub

    Private Sub CheckDev2Point_CheckedChanged(sender As Object, e As EventArgs) Handles CheckDev2Point.CheckedChanged

        If CheckDev2Point.Checked = True Then
            CheckDev2Line.Checked = False
            System.Threading.Thread.Sleep(50)
            Chart2.Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Point
            Chart2.Series(1).MarkerStep = 1
            Chart2.Series(1).MarkerSize = 2
        Else
            CheckDev2Line.Checked = True
            System.Threading.Thread.Sleep(50)
            Chart2.Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Line
        End If

    End Sub

    Private Sub ShowFiles2_Click(sender As Object, e As EventArgs) Handles ShowFiles2.Click
        Process.Start("explorer.exe", String.Format("/n, /e, {0}", PlaybackstrPath))
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        ' 5sec timer for automatic refresh of Playback chart
        RefreshPlaybackCSVFile()

    End Sub

    Private Sub YaxisSave_Click(sender As Object, e As EventArgs) Handles YaxisSave.Click

        If (YaxisCheck1.Checked = True) Then
            My.Settings.data187 = ParseInvariantDouble(YaxisMaximum.Text)
            My.Settings.data188 = ParseInvariantDouble(YaxisMinimum.Text)
        End If

        If (YaxisCheck2.Checked = True) Then
            My.Settings.data189 = ParseInvariantDouble(YaxisMaximum.Text)
            My.Settings.data190 = ParseInvariantDouble(YaxisMinimum.Text)
        End If

        If (YaxisCheck3.Checked = True) Then
            My.Settings.data191 = ParseInvariantDouble(YaxisMaximum.Text)
            My.Settings.data192 = ParseInvariantDouble(YaxisMinimum.Text)
        End If

        If (YaxisCheck4.Checked = True) Then
            My.Settings.data193 = ParseInvariantDouble(YaxisMaximum.Text)
            My.Settings.data194 = ParseInvariantDouble(YaxisMinimum.Text)
        End If

    End Sub

    Private Sub YaxisLoad_Click(sender As Object, e As EventArgs) Handles YaxisLoad.Click

        If (YaxisCheck1.Checked = True) Then
            YaxisMaximum.Text = My.Settings.data187.ToString(Globalization.CultureInfo.InvariantCulture)
            YaxisMinimum.Text = My.Settings.data188.ToString(Globalization.CultureInfo.InvariantCulture)

            Dim result As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")


            RefreshPlaybackCSVFile()
            Dim intervalY As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 20
            If intervalY <= 0 Then intervalY = 0.0000001   ' avoid MSChart crash when Y-max = Y-min (flat/no-variance data)
            Chart2.ChartAreas(0).AxisY.Interval = intervalY
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)
        End If

        If (YaxisCheck2.Checked = True) Then
            YaxisMaximum.Text = My.Settings.data189.ToString(Globalization.CultureInfo.InvariantCulture)
            YaxisMinimum.Text = My.Settings.data190.ToString(Globalization.CultureInfo.InvariantCulture)

            Dim result As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")


            RefreshPlaybackCSVFile()
            Dim intervalY As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 20
            If intervalY <= 0 Then intervalY = 0.0000001   ' avoid MSChart crash when Y-max = Y-min (flat/no-variance data)
            Chart2.ChartAreas(0).AxisY.Interval = intervalY
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)
        End If

        If (YaxisCheck3.Checked = True) Then
            YaxisMaximum.Text = My.Settings.data191.ToString(Globalization.CultureInfo.InvariantCulture)
            YaxisMinimum.Text = My.Settings.data192.ToString(Globalization.CultureInfo.InvariantCulture)

            Dim result As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")


            RefreshPlaybackCSVFile()
            Dim intervalY As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 20
            If intervalY <= 0 Then intervalY = 0.0000001   ' avoid MSChart crash when Y-max = Y-min (flat/no-variance data)
            Chart2.ChartAreas(0).AxisY.Interval = intervalY
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)
        End If

        If (YaxisCheck4.Checked = True) Then
            YaxisMaximum.Text = My.Settings.data193.ToString(Globalization.CultureInfo.InvariantCulture)
            YaxisMinimum.Text = My.Settings.data194.ToString(Globalization.CultureInfo.InvariantCulture)

            Dim result As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")


            RefreshPlaybackCSVFile()
            Dim intervalY As Double = (ParseInvariantDouble(YaxisMaximum.Text) - ParseInvariantDouble(YaxisMinimum.Text)) / 20
            If intervalY <= 0 Then intervalY = 0.0000001   ' avoid MSChart crash when Y-max = Y-min (flat/no-variance data)
            Chart2.ChartAreas(0).AxisY.Interval = intervalY
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 7)
        End If





        'Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((ParseInvariantDouble(YaxisMinimum.Text)), 7)
        'Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((ParseInvariantDouble(YaxisMaximum.Text)), 1)
        'Chart2.ChartAreas(0).AxisY.Minimum = YaxisMinimum.Text
        'Chart2.ChartAreas(0).AxisY.Maximum = YaxisMaximum.Text



    End Sub

    Private Sub YaxisCheck1_CheckedChanged(sender As Object, e As EventArgs) Handles YaxisCheck1.CheckedChanged

        If (YaxisCheck1.Checked = True) Then

            CheckBoxMaxMin.Checked = False

            YaxisCheck2.Checked = False
            YaxisCheck3.Checked = False
            YaxisCheck4.Checked = False
        End If

    End Sub

    Private Sub YaxisCheck2_CheckedChanged(sender As Object, e As EventArgs) Handles YaxisCheck2.CheckedChanged

        If (YaxisCheck2.Checked = True) Then

            CheckBoxMaxMin.Checked = False

            YaxisCheck1.Checked = False
            YaxisCheck3.Checked = False
            YaxisCheck4.Checked = False
        End If

    End Sub

    Private Sub YaxisCheck3_CheckedChanged(sender As Object, e As EventArgs) Handles YaxisCheck3.CheckedChanged

        If (YaxisCheck3.Checked = True) Then

            CheckBoxMaxMin.Checked = False

            YaxisCheck1.Checked = False
            YaxisCheck2.Checked = False
            YaxisCheck4.Checked = False
        End If

    End Sub

    Private Sub YaxisCheck4_CheckedChanged(sender As Object, e As EventArgs) Handles YaxisCheck4.CheckedChanged

        If (YaxisCheck4.Checked = True) Then

            CheckBoxMaxMin.Checked = False

            YaxisCheck1.Checked = False
            YaxisCheck2.Checked = False
            YaxisCheck3.Checked = False
        End If

    End Sub

    Private Sub CheckBoxMaxMin_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxMaxMin.CheckedChanged

        If (CheckBoxMaxMin.Checked = True) Then

            YaxisCheck1.Checked = False
            YaxisCheck2.Checked = False
            YaxisCheck3.Checked = False
            YaxisCheck4.Checked = False

        End If

    End Sub

    Private Sub CheckX10000_CheckedChanged(sender As Object, e As EventArgs) Handles CheckX1000000.CheckedChanged

        If (CheckX1000000.Checked = True) Then
            CheckX1000.Checked = False
        End If

        DevicesMinMax()
        AverageNoise()

    End Sub

    Private Sub CheckX1000_CheckedChanged(sender As Object, e As EventArgs) Handles CheckX1000.CheckedChanged

        If (CheckX1000.Checked = True) Then
            CheckX1000000.Checked = False
        End If

        DevicesMinMax()
        AverageNoise()

    End Sub

    Private Sub DEV1avg_TextChanged(sender As Object, e As EventArgs) Handles DEV1avg.TextChanged

        Dim userInput As String = DEV1avg.Text.Trim()
        Dim isNumeric As Boolean = Integer.TryParse(userInput, Nothing)

        If Not String.IsNullOrEmpty(userInput) AndAlso isNumeric Then
            ' limits of Dev 1 averaging
            If DEV1avg.Text > 100 Then
                DEV1avg.Text = 100
            End If
            RefreshPlaybackCSVFile()

            ' RefreshPlaybackCSVFile() only recalculates scales/ticks - it
            ' never re-plots the Data trace itself, so changing the
            ' averaging window here had no visible effect until something
            ' else (e.g. Zoom All) happened to trigger a full replot.
            FilterDeviceName1()
        End If

    End Sub

    Private Sub DEV2avg_TextChanged(sender As Object, e As EventArgs) Handles DEV2avg.TextChanged

        Dim userInput As String = DEV2avg.Text.Trim()
        Dim isNumeric As Boolean = Integer.TryParse(userInput, Nothing)

        If Not String.IsNullOrEmpty(userInput) AndAlso isNumeric Then
            ' limits of Dev 2 averaging
            If DEV2avg.Text > 100 Then
                DEV2avg.Text = 100
            End If
            RefreshPlaybackCSVFile()

            ' See DEV1avg_TextChanged - RefreshPlaybackCSVFile() alone
            ' doesn't re-plot the Data trace.
            FilterDeviceName2()
        End If

    End Sub

    Private Sub MedianTemp_TextChanged(sender As Object, e As EventArgs) Handles MedianTemp.TextChanged

        'CheckBoxMedianT.Checked = False

    End Sub

    Private Sub MedianValueText_Click(sender As Object, e As EventArgs) Handles MedianValueText.Click

        'CheckBoxMedianV.Checked = False

    End Sub

    Private Sub RadioButtonDev1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonDev1.CheckedChanged

        RefreshPlaybackCSVFile()

    End Sub

    Private Sub RadioButtonDev2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonDev2.CheckedChanged

        RefreshPlaybackCSVFile()

    End Sub

    Private Sub CheckBoxMedianT_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxMedianT.CheckedChanged

        RefreshPlaybackCSVFile()

    End Sub

    Private Sub CheckBoxMedianV_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxMedianV.CheckedChanged

        RefreshPlaybackCSVFile()

    End Sub

    Private Sub PPMscalerangeentry_TextChanged(sender As Object, e As EventArgs) Handles PPMscalerangeentry.TextChanged

        Dim userInput As String =
        PPMscalerangeentry.Text.Replace(vbCr, "").Replace(vbLf, "").Trim()

        Dim value As Double

        If Double.TryParse(userInput, value) Then

            If value > 198 Then value = 198
            If value < 0.2 Then value = 0.2

            Dim correctedText As String = value.ToString("0.###")

            If PPMscalerangeentry.Text <> correctedText Then

                PPMscalerangeentry.Text = correctedText
                PPMscalerangeentry.SelectionStart =
                PPMscalerangeentry.Text.Length

                Exit Sub

            End If

            RefreshPlaybackCSVFile()

        End If

    End Sub

    Private Sub ChartScaleMax_TextChanged(sender As Object, e As EventArgs) Handles ChartScaleMax.TextChanged

        Dim userInput As String = ChartScaleMax.Text.Trim()
        Dim isNumeric As Boolean = Integer.TryParse(userInput, Nothing)

        If Not String.IsNullOrEmpty(userInput) AndAlso isNumeric AndAlso ParseInvariantDouble(ChartScaleMax.Text) > ParseInvariantDouble(ChartScaleMin.Text) Then
            ' limits of Dev 2 averaging
            If ChartScaleMax.Text > 200 Then
                ChartScaleMax.Text = 200
            End If
            RefreshPlaybackCSVFile()
        End If

    End Sub

    Private Sub TEMPavg_TextChanged(sender As Object, e As EventArgs) Handles TEMPavg.TextChanged

        If PlaybackTemp.Checked = True Then

            Dim userInput As String = TEMPavg.Text.Trim()
            Dim isNumeric As Boolean = Integer.TryParse(userInput, Nothing)

            If Not String.IsNullOrEmpty(userInput) AndAlso isNumeric Then
                ' limits of Dev 1 averaging
                If TEMPavg.Text > 100 Then
                    TEMPavg.Text = 100
                End If
                RefreshPlaybackCSVFile()

                ' See DEV1avg_TextChanged - RefreshPlaybackCSVFile() alone
                ' doesn't re-plot the Temp trace.
                FilterTempDevice1()
            End If

        End If

    End Sub

    Private Sub ChartScaleMin_TextChanged(sender As Object, e As EventArgs) Handles ChartScaleMin.TextChanged

        Dim userInput As String = ChartScaleMin.Text.Trim()
        Dim isNumeric As Boolean = Integer.TryParse(userInput, Nothing)

        If Not String.IsNullOrEmpty(userInput) AndAlso isNumeric AndAlso ParseInvariantDouble(ChartScaleMax.Text) > ParseInvariantDouble(ChartScaleMin.Text) Then
            ' limits of Dev 2 averaging
            If ChartScaleMin.Text > 200 Then
                ChartScaleMin.Text = 200
            End If
            RefreshPlaybackCSVFile()
        End If

    End Sub

    Private Sub ChartOffReadyForCSV()

        Chart2.Visible = False      ' invisible until CSV loading

        Scale1.Visible = False
        Scale2.Visible = False
        Scale3.Visible = False
        Scale4.Visible = False
        Scale5.Visible = False
        Scale6.Visible = False
        Scale7.Visible = False
        Scale8.Visible = False
        Scale9.Visible = False
        Scale10.Visible = False
        Scale11.Visible = False
        Scale12.Visible = False
        Scale13.Visible = False
        Scale14.Visible = False
        Scale15.Visible = False
        Scale16.Visible = False
        Scale17.Visible = False
        Scale18.Visible = False
        Scale19.Visible = False
        Scale20.Visible = False
        Scale21.Visible = False
        Scale22.Visible = False
        Scale23.Visible = False
        Scale24.Visible = False
        Scale25.Visible = False
        ButtonShiftUp.Visible = False
        ButtonShiftDn.Visible = False
        Xscale.Visible = False
        Xscaletotal.Visible = False
        LabelTempC.Visible = False
        LabelHum.Visible = False
        LabelPPMtop.Visible = False
        LabelPPMdegctop.Visible = False
        Loading.Visible = False
        PleaseLoadCSV.Visible = True

    End Sub

    Private Sub CheckBoxColours_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxColours.CheckedChanged

        If CheckBoxColours.Checked = False Then
            ' normal mode
            Chart2.ChartAreas(0).AxisX.MajorGrid.LineColor = Color.FromArgb(255, 85, 85, 85)
            Chart2.ChartAreas(0).AxisY.MajorGrid.LineColor = Color.FromArgb(255, 85, 85, 85)
            Chart2.ChartAreas(0).AxisX.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
            Chart2.ChartAreas(0).AxisY.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
            Chart2.ChartAreas(0).AxisY2.MajorGrid.LineColor = Color.FromArgb(100, 85, 85, 85)
            Chart2.ChartAreas(0).AxisY2.MinorGrid.LineColor = Color.FromArgb(100, 85, 85, 85)
            Chart2.Series(0).Color = Color.GreenYellow
            Chart2.Series(1).Color = Color.Violet
            Chart2.Series(2).Color = Color.Red
            Chart2.Series(3).Color = Color.DodgerBlue
            Chart2.Series(4).Color = Color.White
            Chart2.ChartAreas(0).BackColor = Color.Black

            ' label colours
            Xscale.BackColor = Color.Black
            Xscale.ForeColor = Color.Orange
            LabelPPMtop.BackColor = Color.Black
            LabelPPMtop.ForeColor = Color.White
            LabelPPMdegctop.BackColor = Color.Black
            LabelPPMdegctop.ForeColor = Color.White
            LabelTempC.BackColor = Color.Black
            LabelTempC.ForeColor = Color.Red
            LabelHum.BackColor = Color.Black
            LabelHum.ForeColor = Color.DodgerBlue

            ' Set background colours to normal
            Chart2.BackColor = SystemColors.Control
            Me.BackColor = SystemColors.Control
        Else
            ' light mode
            Chart2.ChartAreas(0).AxisX.MajorGrid.LineColor = Color.FromArgb(155, 185, 185, 185)
            Chart2.ChartAreas(0).AxisY.MajorGrid.LineColor = Color.FromArgb(155, 185, 185, 185)
            Chart2.ChartAreas(0).AxisX.MinorGrid.LineColor = Color.FromArgb(155, 185, 185, 185)
            Chart2.ChartAreas(0).AxisY.MinorGrid.LineColor = Color.FromArgb(155, 185, 185, 185)
            Chart2.ChartAreas(0).AxisY2.MajorGrid.LineColor = Color.FromArgb(50, 185, 185, 185)
            Chart2.ChartAreas(0).AxisY2.MinorGrid.LineColor = Color.FromArgb(50, 185, 185, 185)
            Chart2.Series(0).Color = Color.DarkGreen
            Chart2.Series(1).Color = Color.DarkViolet
            Chart2.Series(2).Color = Color.Red
            Chart2.Series(3).Color = Color.DodgerBlue
            Chart2.Series(4).Color = Color.Gray
            Chart2.ChartAreas(0).BackColor = Color.White

            ' label colours
            Xscale.BackColor = Color.White
            Xscale.ForeColor = Color.Black
            LabelPPMtop.BackColor = Color.White
            LabelPPMtop.ForeColor = Color.Black
            LabelPPMdegctop.BackColor = Color.White
            LabelPPMdegctop.ForeColor = Color.Black
            LabelTempC.BackColor = Color.White
            LabelTempC.ForeColor = Color.Black
            LabelHum.BackColor = Color.White
            LabelHum.ForeColor = Color.Black

            ' Set background colours to white
            Chart2.BackColor = Color.White
            Me.BackColor = Color.White
        End If


    End Sub


    Private Sub AverageNoise()

        RMSaverageDev1.Text = ""
        RMSaverageDev2.Text = ""

        ' Dev 1 rolling average with compensation for drift
        If (DeviceName1.Text <> "") Then

            ' Parameters for moving average filter
            Dim columnName As String = "VALUE"
            'Dim windowSize As Integer = CSVfileLines.Text
            Dim windowSize As Integer = 100
            windowSize = Val(RMSwindow.Text)
            If windowSize <= 0 Then
                windowSize = 10
                RMSwindow.Text = "10"
            End If
            If windowSize > Val(CSVfileLines.Text) Then
                windowSize = Val(CSVfileLines.Text)
                RMSwindow.Text = Val(CSVfileLines.Text)
            End If

            ' Lists to store data
            Dim voltageData As New List(Of Double)
            Dim baseline As New List(Of Double)

            Dim selectedRows() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")

            For Each dr As DataRow In selectedRows
                If Not dr.IsNull(columnName) Then
                    voltageData.Add(Convert.ToDouble(dr(columnName)))
                End If
            Next

            ' Calculate moving average (baseline)
            For i As Integer = 0 To voltageData.Count - 1
                Dim startIndex As Integer = Math.Max(0, i - windowSize + 1)
                Dim endIndex As Integer = i
                Dim window As List(Of Double) = voltageData.GetRange(startIndex, endIndex - startIndex + 1)
                Dim average As Double = window.Average()
                baseline.Add(average)
            Next

            ' Calculate noise (RMS) after removing baseline (drift)
            Dim noiseData As List(Of Double) = voltageData.Zip(baseline, Function(voltage, baselineValue) voltage - baselineValue).ToList()
            Dim sumOfSquares As Double = noiseData.Sum(Function(value) value * value)
            Dim meanSquare As Double = sumOfSquares / noiseData.Count
            Dim rmsNoise As Double = Math.Sqrt(meanSquare)

            'Console.WriteLine($"RMS Noise: {rmsNoise}")
            'Console.WriteLine($"RMS Noise: {rmsNoise:F10}")

            If (CheckX1000000.Checked = True Or CheckX1000.Checked = True) Then
                If (CheckX1000000.Checked = True) Then
                    RMSaverageDev1.Text = (rmsNoise * 1000000).ToString("0.#########")
                End If
                If (CheckX1000.Checked = True) Then
                    RMSaverageDev1.Text = (rmsNoise * 1000).ToString("0.#########")
                End If
            Else
                RMSaverageDev1.Text = rmsNoise.ToString("0.#########")
            End If

        End If


        ' Dev 2 rolling average with compensation for drift
        If (DeviceName2.Text <> "") Then

            ' Parameters for moving average filter
            Dim columnName As String = "VALUE"
            'Dim windowSize As Integer = CSVfileLines.Text
            Dim windowSize As Integer = 100
            windowSize = Val(RMSwindow.Text)
            If windowSize <= 0 Then
                windowSize = 10
                RMSwindow.Text = "10"
            End If
            If windowSize > Val(CSVfileLines.Text) Then
                windowSize = Val(CSVfileLines.Text)
                RMSwindow.Text = Val(CSVfileLines.Text)
            End If

            ' Lists to store data
            Dim voltageData As New List(Of Double)
            Dim baseline As New List(Of Double)

            Dim selectedRows() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")

            For Each dr As DataRow In selectedRows
                If Not dr.IsNull(columnName) Then
                    voltageData.Add(Convert.ToDouble(dr(columnName)))
                End If
            Next

            ' Calculate moving average (baseline)
            For i As Integer = 0 To voltageData.Count - 1
                Dim startIndex As Integer = Math.Max(0, i - windowSize + 1)
                Dim endIndex As Integer = i
                Dim window As List(Of Double) = voltageData.GetRange(startIndex, endIndex - startIndex + 1)
                Dim average As Double = window.Average()
                baseline.Add(average)
            Next

            ' Calculate noise (RMS) after removing baseline (drift)
            Dim noiseData As List(Of Double) = voltageData.Zip(baseline, Function(voltage, baselineValue) voltage - baselineValue).ToList()
            Dim sumOfSquares As Double = noiseData.Sum(Function(value) value * value)
            Dim meanSquare As Double = sumOfSquares / noiseData.Count
            Dim rmsNoise As Double = Math.Sqrt(meanSquare)

            'Console.WriteLine($"RMS Noise: {rmsNoise}")
            'Console.WriteLine($"RMS Noise: {rmsNoise:F10}")

            If (CheckX1000000.Checked = True Or CheckX1000.Checked = True) Then

                If (CheckX1000000.Checked = True) Then
                    RMSaverageDev2.Text = (rmsNoise * 1000000).ToString("0.#########")
                End If

                If (CheckX1000.Checked = True) Then
                    RMSaverageDev2.Text = (rmsNoise * 1000).ToString("0.#########")
                End If

            Else

                RMSaverageDev2.Text = rmsNoise.ToString("0.#########")

            End If

        End If

    End Sub

    Private Sub RMSwindow_TextChanged(sender As Object, e As EventArgs) Handles RMSwindow.TextChanged

        AverageNoise()

    End Sub






    Private Sub EnhanceTextBoxBorders(root As Control)
        For Each c As Control In AllControls(root)

            If TypeOf c Is TextBox Then
                Dim tb = DirectCast(c, TextBox)

                ' Skip only if already wrapped by OUR border panel
                If TypeOf tb.Parent Is Panel AndAlso Equals(tb.Parent.Tag, "TB_BORDER") Then Continue For

                Dim parent = tb.Parent

                ' Remember Z-order position before re-parenting
                Dim z = parent.Controls.GetChildIndex(tb)

                ' Outer border panel (grey)
                Dim border As New Panel With {
                .Tag = "TB_BORDER",
                .BackColor = Color.FromArgb(160, 160, 160),
                .Location = tb.Location,
                .Size = tb.Size,
                .Anchor = tb.Anchor,
                .Margin = tb.Margin,
                .Padding = New Padding(1)
            }

                ' Inner panel (white) provides the padding/vertical offset
                Dim inner As New Panel With {
                .BackColor = Color.White,
                .Dock = DockStyle.Fill,
                .Padding = New Padding(0, 2, 0, 0)
            }

                ' TextBox inside
                tb.BorderStyle = BorderStyle.None
                'tb.Multiline = True
                tb.Dock = DockStyle.Fill
                tb.Margin = New Padding(0)

                ' Re-parent keeping original Z order
                parent.Controls.Add(border)
                parent.Controls.SetChildIndex(border, z)
                border.Controls.Add(inner)
                inner.Controls.Add(tb)
            End If

        Next
    End Sub




    Private Sub MakeButtonsWin10ish(root As Control)
        For Each c As Control In AllControls(root)
            If TypeOf c Is Button Then
                Dim b = DirectCast(c, Button)

                b.FlatStyle = FlatStyle.Flat
                b.UseVisualStyleBackColor = False

                b.FlatAppearance.BorderSize = 1
                b.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200)

                If b.Enabled Then
                    b.BackColor = Color.White
                    b.ForeColor = Color.Black
                Else
                    b.BackColor = Color.FromArgb(245, 245, 245)
                    b.ForeColor = Color.FromArgb(80, 80, 80)
                End If
            End If
        Next
    End Sub


    Private Iterator Function AllControls(root As Control) As IEnumerable(Of Control)
        Dim stack As New Stack(Of Control)
        stack.Push(root)
        While stack.Count > 0
            Dim parent = stack.Pop()
            For Each child As Control In parent.Controls
                Yield child
                If child.HasChildren Then stack.Push(child)
            Next
        End While
    End Function


    Private Sub AddPlaybackCSVRow(line As String)

        If String.IsNullOrWhiteSpace(line) Then Exit Sub
        If line.TrimStart().StartsWith("//") Then Exit Sub
        If String.IsNullOrEmpty(CSVdelimit) Then Exit Sub

        Dim values As String() =
            line.Split(New String() {CSVdelimit},
                       StringSplitOptions.None)

        ' Old WinGPIB CSV requires at least the original 6 fields.
        If values.Length < 6 Then Exit Sub

        If Not IsNumeric(values(0)) Then Exit Sub

        Dim row As DataRow = dataTable1.NewRow()

        ' ==========================================================
        ' Standard fields - old and new CSV
        ' ==========================================================

        row("INDEX") = CInt(Val(values(0)))
        row("DEVICE") = values(1)
        row("DATETIME") = values(2)

        row("VALUE") = CDbl(Val(values(3)))
        row("TEMP") = CDbl(Val(values(4)))
        row("HUM") = CDbl(Val(values(5)))


        ' ==========================================================
        ' V5 statistics fields
        ' ==========================================================

        If values.Length >= 16 Then

            row("DEV1_SAMPLES") = values(6)
            row("DEV1_MEAN") = values(7)
            row("DEV1_STDEV") = values(8)
            row("DEV1_SEM") = values(9)
            row("DEV1_GAIN") = values(10)

            row("DEV2_SAMPLES") = values(11)
            row("DEV2_MEAN") = values(12)
            row("DEV2_STDEV") = values(13)
            row("DEV2_SEM") = values(14)
            row("DEV2_GAIN") = values(15)

        Else

            ' Old CSV - statistics do not exist.
            row("DEV1_SAMPLES") = ""
            row("DEV1_MEAN") = ""
            row("DEV1_STDEV") = ""
            row("DEV1_SEM") = ""
            row("DEV1_GAIN") = ""

            row("DEV2_SAMPLES") = ""
            row("DEV2_MEAN") = ""
            row("DEV2_STDEV") = ""
            row("DEV2_SEM") = ""
            row("DEV2_GAIN") = ""

        End If


        ' ==========================================================
        ' V6 statistics fields (Max Diff / Deviation) - appended
        ' after the original V5 block, so V5 CSVs (exactly 16
        ' fields) still load correctly with these left blank.
        ' ==========================================================

        If values.Length >= 20 Then

            row("DEV1_MAXDIFF") = values(16)
            row("DEV1_DEVIATION") = values(17)
            row("DEV2_MAXDIFF") = values(18)
            row("DEV2_DEVIATION") = values(19)

        Else

            row("DEV1_MAXDIFF") = ""
            row("DEV1_DEVIATION") = ""
            row("DEV2_MAXDIFF") = ""
            row("DEV2_DEVIATION") = ""

        End If


        ' PPM is calculated by Playback.
        row("PPM") = 0.0

        dataTable1.Rows.Add(row)

    End Sub


    Private Sub FilterDev1Mean()

        Chart2.Series(5).Points.Clear()

        If CheckPlaybackDev1Mean.Checked = False Then Exit Sub
        If DeviceName1.Text = "" Then Exit Sub

        Dim selectedRows() As DataRow =
        dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")

        For Each dr As DataRow In selectedRows

            Dim s As String = dr("DEV1_MEAN").ToString().Trim()

            If s <> "" AndAlso s.ToLower() <> "nil" Then
                Chart2.Series(5).Points.AddXY(
                dr("DEVICE"),
                CDbl(Val(s)))
            End If

        Next

    End Sub


    Private Sub FilterDev1Stdev()

        Chart2.Series(6).Points.Clear()

        If CheckPlaybackDev1Stdev.Checked = False Then Exit Sub
        If DeviceName1.Text = "" Then Exit Sub

        Dim selectedRows() As DataRow =
        dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")

        For Each dr As DataRow In selectedRows

            Dim s As String = dr("DEV1_STDEV").ToString().Trim()

            If s <> "" AndAlso s.ToLower() <> "nil" Then
                Chart2.Series(6).Points.AddXY(
                dr("DEVICE"),
                CDbl(Val(s)))
            End If

        Next

    End Sub


    Private Sub FilterDev1SEM()

        Chart2.Series(7).Points.Clear()

        If CheckPlaybackDev1SEM.Checked = False Then Exit Sub
        If DeviceName1.Text = "" Then Exit Sub

        Dim selectedRows() As DataRow =
        dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")

        For Each dr As DataRow In selectedRows

            Dim s As String = dr("DEV1_SEM").ToString().Trim()

            If s <> "" AndAlso s.ToLower() <> "nil" Then
                Chart2.Series(7).Points.AddXY(
                dr("DEVICE"),
                CDbl(Val(s)))
            End If

        Next

    End Sub


    Private Sub FilterDev2Mean()

        Chart2.Series(8).Points.Clear()

        If CheckPlaybackDev2Mean.Checked = False Then Exit Sub
        If DeviceName2.Text = "" Then Exit Sub

        Dim selectedRows() As DataRow =
        dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")

        For Each dr As DataRow In selectedRows

            Dim s As String = dr("DEV2_MEAN").ToString().Trim()

            If s <> "" AndAlso s.ToLower() <> "nil" Then
                Chart2.Series(8).Points.AddXY(
                dr("DEVICE"),
                CDbl(Val(s)))
            End If

        Next

    End Sub


    Private Sub FilterDev2Stdev()

        Chart2.Series(9).Points.Clear()

        If CheckPlaybackDev2Stdev.Checked = False Then Exit Sub
        If DeviceName2.Text = "" Then Exit Sub

        Dim selectedRows() As DataRow =
        dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")

        For Each dr As DataRow In selectedRows

            Dim s As String = dr("DEV2_STDEV").ToString().Trim()

            If s <> "" AndAlso s.ToLower() <> "nil" Then
                Chart2.Series(9).Points.AddXY(
                dr("DEVICE"),
                CDbl(Val(s)))
            End If

        Next

    End Sub


    Private Sub FilterDev2SEM()

        Chart2.Series(10).Points.Clear()

        If CheckPlaybackDev2SEM.Checked = False Then Exit Sub
        If DeviceName2.Text = "" Then Exit Sub

        Dim selectedRows() As DataRow =
        dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")

        For Each dr As DataRow In selectedRows

            Dim s As String = dr("DEV2_SEM").ToString().Trim()

            If s <> "" AndAlso s.ToLower() <> "nil" Then
                Chart2.Series(10).Points.AddXY(
                dr("DEVICE"),
                CDbl(Val(s)))
            End If

        Next

    End Sub


    Private Sub FilterDev1MaxDiff()

        Chart2.Series(11).Points.Clear()

        If CheckPlaybackDev1MaxDiff.Checked = False Then Exit Sub
        If DeviceName1.Text = "" Then Exit Sub

        Dim selectedRows() As DataRow =
        dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")

        For Each dr As DataRow In selectedRows

            Dim s As String = dr("DEV1_MAXDIFF").ToString().Trim()

            If s <> "" AndAlso s.ToLower() <> "nil" Then
                Chart2.Series(11).Points.AddXY(
                dr("DEVICE"),
                CDbl(Val(s)))
            End If

        Next

    End Sub


    Private Sub FilterDev1Deviation()

        Chart2.Series(12).Points.Clear()

        If CheckPlaybackDev1Deviation.Checked = False Then Exit Sub
        If DeviceName1.Text = "" Then Exit Sub

        Dim selectedRows() As DataRow =
        dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")

        For Each dr As DataRow In selectedRows

            Dim s As String = dr("DEV1_DEVIATION").ToString().Trim()

            If s <> "" AndAlso s.ToLower() <> "nil" Then
                Chart2.Series(12).Points.AddXY(
                dr("DEVICE"),
                CDbl(Val(s)))
            End If

        Next

    End Sub


    Private Sub FilterDev2MaxDiff()

        Chart2.Series(13).Points.Clear()

        If CheckPlaybackDev2MaxDiff.Checked = False Then Exit Sub
        If DeviceName2.Text = "" Then Exit Sub

        Dim selectedRows() As DataRow =
        dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")

        For Each dr As DataRow In selectedRows

            Dim s As String = dr("DEV2_MAXDIFF").ToString().Trim()

            If s <> "" AndAlso s.ToLower() <> "nil" Then
                Chart2.Series(13).Points.AddXY(
                dr("DEVICE"),
                CDbl(Val(s)))
            End If

        Next

    End Sub


    Private Sub FilterDev2Deviation()

        Chart2.Series(14).Points.Clear()

        If CheckPlaybackDev2Deviation.Checked = False Then Exit Sub
        If DeviceName2.Text = "" Then Exit Sub

        Dim selectedRows() As DataRow =
        dataTable1.Select("DEVICE ='" & DeviceName2.Text & "'")

        For Each dr As DataRow In selectedRows

            Dim s As String = dr("DEV2_DEVIATION").ToString().Trim()

            If s <> "" AndAlso s.ToLower() <> "nil" Then
                Chart2.Series(14).Points.AddXY(
                dr("DEVICE"),
                CDbl(Val(s)))
            End If

        Next

    End Sub


    Private Sub UpdatePlaybackStatsSeries()

        FilterDev1Mean()
        FilterDev1Stdev()
        FilterDev1SEM()
        FilterDev1MaxDiff()
        FilterDev1Deviation()

        FilterDev2Mean()
        FilterDev2Stdev()
        FilterDev2SEM()
        FilterDev2MaxDiff()
        FilterDev2Deviation()

        If Chart2.ChartAreas.IndexOf("Statistics") >= 0 Then

            ' Reset to auto (NaN) before recalculating - once
            ' RecalculateAxesScale() runs, it assigns concrete
            ' numbers to Minimum/Maximum rather than leaving the
            ' axis in auto mode, so without this reset every call
            ' after the first just reuses the original range instead
            ' of rescaling to the newly zoomed/scrolled data.
            With Chart2.ChartAreas("Statistics").AxisY
                .Minimum = Double.NaN
                .Maximum = Double.NaN
                .Interval = Double.NaN
            End With

            Chart2.ChartAreas("Statistics").RecalculateAxesScale()

        End If

    End Sub


    Private Sub PlaybackTrace_CheckedChanged(sender As Object, e As EventArgs) _
    Handles CheckPlaybackDev1Data.CheckedChanged,
            CheckPlaybackDev1Mean.CheckedChanged,
            CheckPlaybackDev1Stdev.CheckedChanged,
            CheckPlaybackDev1SEM.CheckedChanged,
            CheckPlaybackDev1MaxDiff.CheckedChanged,
            CheckPlaybackDev1Deviation.CheckedChanged,
            CheckPlaybackDev1ShortTermMean.CheckedChanged,
            CheckPlaybackDev2Data.CheckedChanged,
            CheckPlaybackDev2Mean.CheckedChanged,
            CheckPlaybackDev2Stdev.CheckedChanged,
            CheckPlaybackDev2SEM.CheckedChanged,
            CheckPlaybackDev2MaxDiff.CheckedChanged,
            CheckPlaybackDev2Deviation.CheckedChanged,
            CheckPlaybackDev2ShortTermMean.CheckedChanged

        If Chart2.Series.Count < 17 Then Exit Sub

        Chart2.Series(0).Enabled = CheckPlaybackDev1Data.Checked
        Chart2.Series(1).Enabled = CheckPlaybackDev2Data.Checked

        Chart2.Series(5).Enabled = CheckPlaybackDev1Mean.Checked
        Chart2.Series(6).Enabled = CheckPlaybackDev1Stdev.Checked
        Chart2.Series(7).Enabled = CheckPlaybackDev1SEM.Checked
        Chart2.Series(11).Enabled = CheckPlaybackDev1MaxDiff.Checked
        Chart2.Series(12).Enabled = CheckPlaybackDev1Deviation.Checked
        Chart2.Series(15).Enabled = CheckPlaybackDev1ShortTermMean.Checked

        Chart2.Series(8).Enabled = CheckPlaybackDev2Mean.Checked
        Chart2.Series(9).Enabled = CheckPlaybackDev2Stdev.Checked
        Chart2.Series(10).Enabled = CheckPlaybackDev2SEM.Checked
        Chart2.Series(13).Enabled = CheckPlaybackDev2MaxDiff.Checked
        Chart2.Series(14).Enabled = CheckPlaybackDev2Deviation.Checked
        Chart2.Series(16).Enabled = CheckPlaybackDev2ShortTermMean.Checked

        ' Short-Term Mean isn't populated by UpdatePlaybackStatsSeries() (it's
        ' not a recorded CSV column, it's recomputed from raw VALUE) - refresh
        ' it directly so toggling the checkbox actually shows/hides real points.
        FilterShortTermMeanDevice1()
        FilterShortTermMeanDevice2()

        If ChartLoaded = True AndAlso CSVfileok = True Then
            UpdatePlaybackStatsSeries()
        End If

    End Sub


    ' ==============================================================
    ' Allan Deviation pop-up chart (Dev 1 / Dev 2)
    '
    ' A retrospective stability plot computed from the raw VALUE
    ' column for whichever device(s) are checked - independent of
    ' Chart2 and everything else on the Playback chart. Log-log axes:
    ' averaging time (tau, in samples) on X, Allan deviation (ppm of
    ' the device's overall mean) on Y. Opens on first checkbox tick,
    ' closes when both are unchecked or the user closes it directly.
    ' ==============================================================

    Private AllanPopupForm As Form = Nothing
    Private AllanPopupChart As DataVisualization.Charting.Chart = Nothing
    Private AllanToolTip As ToolTip = Nothing

    ' False = non-overlapping (disjoint tau-length blocks, fewer pairs at
    ' large tau, noisier tail). True = overlapping (sliding window, reuses
    ' every sample many times over, much smoother tail from the same data).
    Private AllanUseOverlapping As Boolean = False

    ' Adds a Modified Allan Deviation (MDEV) curve alongside each shown
    ' device's regular ADEV curve. MDEV applies an extra averaging stage
    ' that makes it react differently to phase noise than ADEV does, so a
    ' visibly steeper MDEV-vs-ADEV gap at short tau indicates phase/timing
    ' noise the regular ADEV curve can't distinguish on its own. Always
    ' uses the standard (overlapping) MDEV estimator regardless of the
    ' Overlapping checkbox above, since that's the only form MDEV is
    ' normally computed in.
    Private AllanShowMDEV As Boolean = False

    Private Sub AllanCheckbox_CheckedChanged(sender As Object, e As EventArgs) _
    Handles CheckPlaybackDev1Allan.CheckedChanged, CheckPlaybackDev2Allan.CheckedChanged

        RefreshAllanChart()

    End Sub

    Private Sub RefreshAllanChart()

        ' No CSV loaded means DeviceName1/2.Text are both blank, so neither
        ' device would have anything to plot - opening the pop-up anyway
        ' would leave its logarithmic axes with zero data points to
        ' auto-range from, which crashes MSChart on the next repaint.
        If Not (ChartLoaded AndAlso CSVfileok) Then
            CheckPlaybackDev1Allan.Checked = False
            CheckPlaybackDev2Allan.Checked = False
            If AllanPopupForm IsNot Nothing Then AllanPopupForm.Close()
            Exit Sub
        End If

        Dim showDev1 As Boolean = CheckPlaybackDev1Allan.Checked
        Dim showDev2 As Boolean = CheckPlaybackDev2Allan.Checked

        If Not showDev1 AndAlso Not showDev2 Then
            If AllanPopupForm IsNot Nothing Then AllanPopupForm.Close()
            Exit Sub
        End If

        EnsureAllanPopupOpen()

        UpdateAllanSeries("Dev 1 Allan Deviation", DeviceName1.Text, showDev1, Color.Yellow)
        UpdateAllanSeries("Dev 2 Allan Deviation", DeviceName2.Text, showDev2, Color.Aqua)

        RescaleAllanAxes()

    End Sub

    ' MSChart's logarithmic axis only labels whole decades (1, 10, 100, ...),
    ' which can leave very few gridlines when the data spans less than a
    ' couple of decades - exactly the "not many points" look. Pin the axis
    ' range to whole decades from the actual data, then add unlabeled minor
    ' gridlines at 2x-9x within each decade (the standard look for a log-log
    ' plot) so there's always a useful density of reference lines.
    Private Sub RescaleAllanAxes()

        If AllanPopupChart Is Nothing Then Exit Sub
        If AllanPopupChart.Series.Count = 0 Then Exit Sub

        Dim xMin As Double = Double.MaxValue, xMax As Double = Double.MinValue
        Dim yMin As Double = Double.MaxValue, yMax As Double = Double.MinValue

        For Each s As Series In AllanPopupChart.Series
            For Each pt As DataPoint In s.Points
                xMin = Math.Min(xMin, pt.XValue)
                xMax = Math.Max(xMax, pt.XValue)
                yMin = Math.Min(yMin, pt.YValues(0))
                yMax = Math.Max(yMax, pt.YValues(0))
            Next
        Next

        If xMin = Double.MaxValue Then Exit Sub   ' no points plotted yet

        Dim axisXMin As Double = Math.Pow(10.0, Math.Floor(Math.Log10(xMin)))
        Dim axisXMax As Double = Math.Pow(10.0, Math.Ceiling(Math.Log10(xMax)))
        Dim axisYMin As Double = Math.Pow(10.0, Math.Floor(Math.Log10(yMin)))
        Dim axisYMax As Double = Math.Pow(10.0, Math.Ceiling(Math.Log10(yMax)))

        Dim ca As ChartArea = AllanPopupChart.ChartAreas("Main")
        ca.AxisX.Minimum = axisXMin
        ca.AxisX.Maximum = axisXMax
        ca.AxisY.Minimum = axisYMin
        ca.AxisY.Maximum = axisYMax

        ' For a logarithmic axis, Interval is a power-of-ten step rather than
        ' a data value - 1 (the default) only ticks whole decades. 0.5 adds a
        ' gridline/label at the half-decade point too (e.g. 0.01, 0.0316, 0.1,
        ' 0.316, 1 instead of just 0.01, 0.1, 1), which is the same mechanism
        ' already drawing the decade lines, just ticking twice as often.
        ca.AxisX.Interval = 0.5
        ca.AxisY.Interval = 0.25
        ' X (tau, always a whole sample count) rounded to no decimals for
        ' display; Y kept at 3 decimals since Allan deviation values are
        ' usually well under 1 and would mostly round to "0".
        ca.AxisX.LabelStyle.Format = "0"
        ca.AxisY.LabelStyle.Format = "0.###"

    End Sub

    Private Sub EnsureAllanPopupOpen()

        If AllanPopupForm IsNot Nothing AndAlso Not AllanPopupForm.IsDisposed Then
            AllanPopupForm.BringToFront()
            Exit Sub
        End If

        AllanPopupForm = New Form With {
            .Text = "Allan Deviation",
            .Width = 780,
            .Height = 540,
            .MinimumSize = New Size(780, 540),
            .StartPosition = FormStartPosition.CenterParent,
            .ShowIcon = False
        }

        AllanPopupChart = New DataVisualization.Charting.Chart With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.Black
        }

        Dim ca As New ChartArea("Main")
        ca.BackColor = Color.Black

        ' Explicit fallback Minimum/Maximum on both axes - a logarithmic
        ' axis that's left on Auto with zero series/points to range from
        ' (e.g. the pop-up ends up empty for any reason) throws an
        ' InvalidOperationException from MSChart on the next repaint.
        ' RescaleAllanAxes() overwrites these with real values as soon as
        ' there's actual data.
        ca.AxisX.IsLogarithmic = True
        ca.AxisX.Minimum = 1
        ca.AxisX.Maximum = 10
        ca.AxisX.Title = "Averaging Time - tau (samples)"
        ca.AxisX.TitleForeColor = Color.White
        ca.AxisX.LabelStyle.ForeColor = Color.White
        ca.AxisX.LineColor = Color.Gray
        ca.AxisX.MajorGrid.LineColor = Color.FromArgb(45, 45, 45)

        ca.AxisY.IsLogarithmic = True
        ca.AxisY.Minimum = 0.001
        ca.AxisY.Maximum = 1
        ca.AxisY.Title = "Allan Deviation (ppm)"
        ca.AxisY.TitleForeColor = Color.White
        ca.AxisY.LabelStyle.ForeColor = Color.White
        ca.AxisY.LineColor = Color.Gray
        ca.AxisY.MajorGrid.LineColor = Color.FromArgb(45, 45, 45)

        AllanPopupChart.ChartAreas.Add(ca)

        Dim lg As New Legend("Main")
        lg.ForeColor = Color.White
        lg.BackColor = Color.Black
        AllanPopupChart.Legends.Add(lg)

        ' Overlapping vs non-overlapping Allan deviation. Overlapping
        ' reuses every sample in many sliding windows instead of chopping
        ' the data into disjoint blocks, giving a much smoother curve at
        ' large tau from the same file - at the cost of the points no
        ' longer being statistically independent of each other.
        Dim overlapCheck As New CheckBox With {
            .Location = New Point(540, 115),
            .Size = New Size(230, 24),
            .Text = "Overlapping (smoother tail)",
            .ForeColor = Color.White,
            .BackColor = Color.Black,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Checked = AllanUseOverlapping
        }
        AddHandler overlapCheck.CheckedChanged,
            Sub()
                AllanUseOverlapping = overlapCheck.Checked
                RefreshAllanChart()
            End Sub

        ' Adds a dotted MDEV curve alongside each shown device's ADEV
        ' curve, in that device's own colour - see AllanShowMDEV.
        Dim mdevCheck As New CheckBox With {
            .Location = New Point(540, 142),
            .Size = New Size(230, 24),
            .Text = "Show MDEV",
            .ForeColor = Color.White,
            .BackColor = Color.Black,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Checked = AllanShowMDEV
        }
        AddHandler mdevCheck.CheckedChanged,
            Sub()
                AllanShowMDEV = mdevCheck.Checked
                RefreshAllanChart()
            End Sub

        AllanToolTip = New ToolTip()
        AllanToolTip.SetToolTip(overlapCheck, "Smooths the tail by reusing every sample in sliding windows instead of separate blocks.")
        AllanToolTip.SetToolTip(mdevCheck, "Adds a dotted curve that reveals phase/timing noise regular ADEV can't show on its own.")

        ' Bottom-right resize grip - purely a visual cue that the window can
        ' be resized. Dragging it hands off to Windows' own native resize
        ' (WM_NCLBUTTONDOWN / HTBOTTOMRIGHT) rather than us tracking the
        ' drag - same approach as LiveWatch.vb's Live Analysis chart pop-up.
        Dim allanGrip As New PictureBox With {
            .Image = InvertGripImage(My.Resources.grip),
            .SizeMode = PictureBoxSizeMode.StretchImage,
            .Size = New Size(36, 36),
            .BackColor = Color.Transparent,
            .Cursor = Cursors.SizeNWSE,
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        }
        allanGrip.Location = New Point(
            AllanPopupForm.ClientSize.Width - allanGrip.Width,
            AllanPopupForm.ClientSize.Height - allanGrip.Height)

        AddHandler allanGrip.MouseDown,
        Sub(gripSender As Object, gripArgs As MouseEventArgs)
            If gripArgs.Button = MouseButtons.Left Then
                ReleaseCapture()
                SendMessage(AllanPopupForm.Handle, &HA1, 17, 0)   ' WM_NCLBUTTONDOWN, HTBOTTOMRIGHT
            End If
        End Sub

        AllanPopupForm.Controls.Add(AllanPopupChart)
        AllanPopupForm.Controls.Add(overlapCheck)
        AllanPopupForm.Controls.Add(mdevCheck)
        AllanPopupForm.Controls.Add(allanGrip)
        overlapCheck.BringToFront()
        mdevCheck.BringToFront()
        allanGrip.BringToFront()

        AddHandler AllanPopupForm.FormClosed, AddressOf AllanPopupForm_FormClosed

        ApplySquareCorners(AllanPopupForm)
        AllanPopupForm.Show()

    End Sub

    Private Sub AllanPopupForm_FormClosed(sender As Object, e As FormClosedEventArgs)

        ' Keep the checkboxes in sync if the user closes the pop-up directly
        ' (via its own close button) instead of unchecking both boxes first.
        ' Guarded against this Chart form already being closed/disposed
        ' (e.g. when Chart_FormClosing below is what triggered this Close)
        ' - its own checkboxes would already be gone, so touching them
        ' would throw.
        If Not Me.IsDisposed Then
            CheckPlaybackDev1Allan.Checked = False
            CheckPlaybackDev2Allan.Checked = False
        End If

        AllanPopupForm = Nothing
        AllanPopupChart = Nothing

        If AllanToolTip IsNot Nothing Then
            AllanToolTip.Dispose()
            AllanToolTip = Nothing
        End If

    End Sub

    Private Sub Chart_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        ' The Allan Deviation pop-up depends entirely on this form's own
        ' live state (dataTable1, DeviceName1/2, the Allan checkboxes) - it
        ' must not be left running once this form is gone. Interacting
        ' with it afterwards (e.g. the Overlapping checkbox, which
        ' recomputes and repaints the chart) could crash with an MSChart
        ' "logarithmic scale" exception once that data is no longer valid.
        ' Closing it here, while this form's own controls are still alive,
        ' also lets AllanPopupForm_FormClosed's checkbox sync run safely.
        If AllanPopupForm IsNot Nothing AndAlso Not AllanPopupForm.IsDisposed Then
            AllanPopupForm.Close()
        End If

    End Sub

    Private Sub UpdateAllanSeries(seriesName As String, deviceName As String, show As Boolean, seriesColor As Color)

        Dim idealSeriesName As String = seriesName & " (Ideal)"
        Dim mdevSeriesName As String = seriesName & " (MDEV)"

        If AllanPopupChart Is Nothing Then Exit Sub

        If AllanPopupChart.Series.IndexOf(seriesName) >= 0 Then
            AllanPopupChart.Series.Remove(AllanPopupChart.Series(seriesName))
        End If
        If AllanPopupChart.Series.IndexOf(idealSeriesName) >= 0 Then
            AllanPopupChart.Series.Remove(AllanPopupChart.Series(idealSeriesName))
        End If
        If AllanPopupChart.Series.IndexOf(mdevSeriesName) >= 0 Then
            AllanPopupChart.Series.Remove(AllanPopupChart.Series(mdevSeriesName))
        End If

        If Not show Then Exit Sub
        If deviceName = "" Then Exit Sub

        Dim selectedRows() As DataRow = dataTable1.Select("DEVICE ='" & deviceName & "'")
        If selectedRows.Length < 4 Then Exit Sub    ' not enough samples for any tau

        Dim rawValues As New List(Of Double)
        For Each dr As DataRow In selectedRows
            rawValues.Add(Convert.ToDouble(dr("VALUE")))
        Next

        Dim overallMean As Double = rawValues.Average()
        If overallMean = 0.0 Then Exit Sub   ' avoid a divide-by-zero in the ppm conversion

        Dim newSeries As New Series(seriesName) With {
            .ChartType = SeriesChartType.Line,
            .ChartArea = "Main",
            .Legend = "Main",
            .Color = seriesColor,
            .BorderWidth = 2,
            .MarkerStyle = MarkerStyle.Circle,
            .MarkerSize = 6,
            .MarkerColor = seriesColor
        }

        Dim firstTau As Integer = 0
        Dim firstSigmaPpm As Double = 0.0
        Dim haveFirst As Boolean = False

        For Each point As KeyValuePair(Of Integer, Double) In ComputeAllanDeviation(rawValues, AllanUseOverlapping)
            Dim sigmaPpm As Double = (point.Value / overallMean) * 1000000.0
            If sigmaPpm > 0.0 Then
                newSeries.Points.AddXY(point.Key, sigmaPpm)
                If Not haveFirst Then
                    firstTau = point.Key
                    firstSigmaPpm = sigmaPpm
                    haveFirst = True
                End If
            End If
        Next

        AllanPopupChart.Series.Add(newSeries)

        ' "Ideal" (white noise) reference line: a straight slope -1/2 on
        ' these log-log axes, anchored to this device's own first plotted
        ' point - i.e. what the curve would look like if averaging longer
        ' kept reducing noise indefinitely with no floor or drift. Only
        ' needs two points since it's a straight line on a log-log plot.
        ' Wherever the real curve departs upward from this dashed line,
        ' something other than plain white noise has taken over (a
        ' flicker floor, or long-term drift) and further averaging isn't
        ' buying you anything.
        If haveFirst AndAlso newSeries.Points.Count >= 2 Then

            Dim lastTau As Double = newSeries.Points(newSeries.Points.Count - 1).XValue

            Dim idealSeries As New Series(idealSeriesName) With {
                .ChartType = SeriesChartType.Line,
                .ChartArea = "Main",
                .Legend = "Main",
                .Color = Color.Gray,
                .BorderWidth = 2,
                .BorderDashStyle = ChartDashStyle.Dot
            }

            idealSeries.Points.AddXY(firstTau, firstSigmaPpm)
            Dim idealEndSigma As Double = firstSigmaPpm * Math.Sqrt(firstTau / lastTau)
            idealSeries.Points.AddXY(lastTau, idealEndSigma)

            AllanPopupChart.Series.Add(idealSeries)

        End If

        ' Modified Allan Deviation (MDEV) - an additional curve in this
        ' device's own colour, dotted so it reads as a companion to the
        ' solid ADEV line above rather than a competing trace. See
        ' AllanShowMDEV for why this exists.
        If AllanShowMDEV Then

            Dim mdevSeries As New Series(mdevSeriesName) With {
                .ChartType = SeriesChartType.Line,
                .ChartArea = "Main",
                .Legend = "Main",
                .Color = seriesColor,
                .BorderWidth = 2,
                .BorderDashStyle = ChartDashStyle.Dot
            }

            For Each point As KeyValuePair(Of Integer, Double) In ComputeModifiedAllanDeviation(rawValues)
                Dim sigmaPpm As Double = (point.Value / overallMean) * 1000000.0
                If sigmaPpm > 0.0 Then mdevSeries.Points.AddXY(point.Key, sigmaPpm)
            Next

            If mdevSeries.Points.Count > 0 Then AllanPopupChart.Series.Add(mdevSeries)

        End If

    End Sub

    ' Non-overlapping Allan deviation: bin the raw samples into
    ' consecutive, disjoint windows of length tau, average each bin, then
    ' take the RMS of the differences between consecutive bin averages.
    ' Overlapping Allan deviation: slide a length-tau window forward one
    ' sample at a time instead of jumping by tau, reusing every sample in
    ' many windows, then RMS the differences between window averages that
    ' are tau apart. Same underlying shape either way, but overlapping
    ' gives far more pairs to average at large tau (where non-overlapping
    ' only has a couple of disjoint blocks left), so the tail comes out
    ' much smoother - at the cost of those pairs no longer being fully
    ' statistically independent of each other.
    ' Both use a log-spaced set of tau values from 1 sample up to half the
    ' total sample count (the minimum needed to form one difference).
    Private Function ComputeAllanDeviation(rawValues As List(Of Double), overlapping As Boolean) As List(Of KeyValuePair(Of Integer, Double))

        Dim results As New List(Of KeyValuePair(Of Integer, Double))

        Dim n As Integer = rawValues.Count
        Dim maxTau As Integer = n \ 2
        If maxTau < 1 Then Return results

        ' Running sum so any window's average is an O(1) lookup, however
        ' many overlapping windows a given tau ends up needing.
        Dim prefixSum(n) As Double
        For k As Integer = 0 To n - 1
            prefixSum(k + 1) = prefixSum(k) + rawValues(k)
        Next
        Dim windowMean = Function(startIndex As Integer, tau As Integer) As Double
                             Return (prefixSum(startIndex + tau) - prefixSum(startIndex)) / tau
                         End Function

        Const stepsPerDecade As Integer = 8
        Dim taus As New SortedSet(Of Integer)
        Dim i As Integer = 0
        Do
            Dim tau As Integer = CInt(Math.Round(Math.Pow(10.0, i / stepsPerDecade)))
            If tau >= 1 AndAlso tau <= maxTau Then taus.Add(tau)
            i += 1
        Loop While Math.Pow(10.0, i / stepsPerDecade) <= maxTau

        For Each tau As Integer In taus

            Dim sumSqDiff As Double = 0.0
            Dim pairCount As Integer = 0

            If overlapping Then

                ' Windows starting at every sample position, compared to
                ' the window starting tau samples later.
                pairCount = n - 2 * tau + 1
                If pairCount < 1 Then Continue For

                For startIndex As Integer = 0 To pairCount - 1
                    Dim diff As Double = windowMean(startIndex + tau, tau) - windowMean(startIndex, tau)
                    sumSqDiff += diff * diff
                Next

            Else

                Dim binCount As Integer = n \ tau
                If binCount < 2 Then Continue For
                pairCount = binCount - 1

                Dim binMeans As New List(Of Double)
                For b As Integer = 0 To binCount - 1
                    binMeans.Add(windowMean(b * tau, tau))
                Next

                For b As Integer = 0 To binMeans.Count - 2
                    Dim diff As Double = binMeans(b + 1) - binMeans(b)
                    sumSqDiff += diff * diff
                Next

            End If

            Dim sigma As Double = Math.Sqrt(0.5 * sumSqDiff / pairCount)
            results.Add(New KeyValuePair(Of Integer, Double)(tau, sigma))

        Next

        Return results

    End Function

    ' Modified Allan Deviation (MDEV) - the standard IEEE-1139 estimator,
    ' always computed the "overlapping" way since that's the only form
    ' MDEV is normally used in (there's no meaningful non-overlapping
    ' variant). Defined on phase data, so the raw frequency-like VALUE
    ' readings are integrated (cumulative sum) first. The extra averaging
    ' stage this adds is what makes MDEV react differently than ADEV to
    ' phase noise - that's its whole purpose.
    Private Function ComputeModifiedAllanDeviation(rawValues As List(Of Double)) As List(Of KeyValuePair(Of Integer, Double))

        Dim results As New List(Of KeyValuePair(Of Integer, Double))

        Dim n As Integer = rawValues.Count
        If n < 6 Then Return results

        Dim x(n) As Double
        x(0) = 0.0
        For k As Integer = 0 To n - 1
            x(k + 1) = x(k) + rawValues(k)
        Next

        Dim maxTau As Integer = n \ 3
        If maxTau < 1 Then Return results

        Const stepsPerDecade As Integer = 8
        Dim taus As New SortedSet(Of Integer)
        Dim i As Integer = 0
        Do
            Dim tau As Integer = CInt(Math.Round(Math.Pow(10.0, i / stepsPerDecade)))
            If tau >= 1 AndAlso tau <= maxTau Then taus.Add(tau)
            i += 1
        Loop While Math.Pow(10.0, i / stepsPerDecade) <= maxTau

        For Each tau As Integer In taus

            Dim m As Integer = tau
            Dim count As Integer = n - 3 * m + 1
            If count < 1 Then Continue For

            ' Running sum of the second difference of phase (three points m
            ' apart) over an m-wide inner window, slid one sample at a time -
            ' the standard efficient way to compute MDEV without recomputing
            ' the inner sum from scratch at every step.
            Dim innerSum As Double = 0.0
            For k As Integer = 0 To m - 1
                innerSum += x(k + 2 * m) - 2.0 * x(k + m) + x(k)
            Next

            Dim sumSq As Double = innerSum * innerSum

            For j As Integer = 1 To count - 1
                Dim dropIndex As Integer = j - 1
                Dim addIndex As Integer = dropIndex + m
                Dim dropVal As Double = x(dropIndex + 2 * m) - 2.0 * x(dropIndex + m) + x(dropIndex)
                Dim addVal As Double = x(addIndex + 2 * m) - 2.0 * x(addIndex + m) + x(addIndex)
                innerSum += addVal - dropVal
                sumSq += innerSum * innerSum
            Next

            Dim sigma As Double = Math.Sqrt(sumSq / (2.0 * tau * tau * m * m * count))
            results.Add(New KeyValuePair(Of Integer, Double)(tau, sigma))

        Next

        Return results

    End Function


    ' ==============================================================
    ' Playback Chart help
    ' ==============================================================

    Private Sub ButtonPlaybackHelp_Click(sender As Object, e As EventArgs) Handles ButtonPlaybackHelp.Click

        Dim frm As New Form With {
            .Text = "Playback Chart Help / Info",
            .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.FixedDialog,
            .ShowIcon = False,
            .ShowInTaskbar = False,
            .Width = 760,
            .Height = 680,
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
"PLAYBACK CHART" & vbLf &
"The Playback Chart loads a previously saved CSV log file and lets you review, zoom and analyse it after the fact - independent of the Live Chart, which only shows data while a device is actively running." & vbLf & vbLf &
"LOADING A CSV" & vbLf &
"LOAD .CSV FILE opens a saved log file from disk. If the CSV only contains data for one device, every Dev.2 checkbox and control is automatically greyed out and unchecked - there is nothing to plot for a device that isn't in the file." & vbLf & vbLf &
"Save Settings stores the current chart control settings (scale, checkboxes, etc.) so they're restored next time." & vbLf & vbLf &
"DEVICES" & vbLf &
"The Dev 1 / Dev 2 radio buttons choose which device's readings feed the PPM Deviation/Tempco calculation and the Y-axis Min/Max reference - they don't hide or show any traces themselves." & vbLf & vbLf &
"X-AXIS SCALE" & vbLf &
"Sets the chart's time axis in minutes and controls how much of the log is visible at once." & vbLf & vbLf &
"Y-AXIS SCALE" & vbLf &
"ZOOM IN / ZOOM OUT - zoom the Y-axis in or out around the centre line." & vbLf & vbLf &
"SHIFT UP / SHIFT DOWN - move the current Y-axis max/min window up or down by 20%, for panning through a large range without changing the zoom level." & vbLf & vbLf &
"ZOOM ALL - resets the Y-axis to show the entire chart." & vbLf & vbLf &
"Auto Min/Max - automatically sets the Y-axis range from the data instead of a fixed range." & vbLf & vbLf &
"Tidy Scale - rounds the Y-axis labels to tidier numbers instead of raw calculated values." & vbLf & vbLf &
"SAVE / LOAD - stores or recalls the current Y-axis Min/Max into one of four saved slots, for quickly switching between preferred view ranges." & vbLf & vbLf &
"x1k / x1000k - rescales the displayed values by 1,000 or 1,000,000 (e.g. VDC to mVDC or " & Global.Microsoft.VisualBasic.ChrW(181) & "VDC) without altering the underlying data." & vbLf & vbLf &
"NAVIGATION" & vbLf &
"Scroll and zoom controls let you move through the loaded file and adjust how much time is shown at once, in both large and small steps." & vbLf & vbLf &
"DEV 1 TRACES / DEV 2 TRACES" & vbLf &
"Each checkbox shows or hides one trace on the top chart, all calculated from the loaded CSV:" & vbLf & vbLf &
"Data - the raw VALUE reading logged for every sample." & vbLf & vbLf &
"Mean - the cumulative Mean recorded in the CSV statistics for that device, running from whenever stats were last reset during acquisition." & vbLf & vbLf &
"STDEV - the recorded Standard Deviation for that device." & vbLf &
"Formula: sqrt( sum( (Xi - Mean)^2 ) / (N - 1) )" & vbLf & vbLf &
"SEM - the recorded Standard Error of the Mean for that device." & vbLf &
"Formula: STDEV / sqrt(N)" & vbLf & vbLf &
"Max Diff. - the recorded Maximum-Minimum spread for that device." & vbLf &
"Formula: Max - Min" & vbLf & vbLf &
"PPM Deviation - the recorded PPM deviation statistic for that device (this is the value saved to the CSV during acquisition - see the PPM Deviation / Tempco section below for the separate, recalculated-on-the-fly PPM trace)." & vbLf &
"Formula: (Value - First Value) / First Value x 1,000,000" & vbLf & vbLf &
"Short Term Mean - see below." & vbLf & vbLf &
"Allan Deviation - see below." & vbLf & vbLf &
$"SHORT TERM MEAN" & vbLf &
$"Plots a rolling average of only the last {ShortTermMeanWindow} raw readings, recomputed directly from the CSV's VALUE column." & vbLf &
"Formula: mean of readings i-N+1 through i (a simple sliding-window average)." & vbLf & vbLf &
"Same concept as the Short-Term Mean on the Live Analysis chart, but calculated retrospectively from the file rather than live." & vbLf & vbLf &
"Responds faster to recent changes than the recorded Mean trace, at the cost of being noisier." & vbLf & vbLf &
"Purely a display trace - it doesn't affect the recorded Mean/STDEV/SEM or anything written back to the CSV." & vbLf & vbLf &
"ALLAN DEVIATION" & vbLf &
"Checking Dev 1 or Dev 2 Allan Deviation opens a separate pop-up chart plotting that device's Allan Deviation (ADEV)." & vbLf & vbLf &
"ADEV is a stability metric showing how much the average reading wanders as you change the averaging time (tau), rather than a single STDEV number for the whole file." & vbLf & vbLf &
"The pop-up's X-axis is averaging time (tau, in samples); the Y-axis is deviation in ppm of that device's overall mean." & vbLf & vbLf &
"Both axes are log-log, rounded outward to whole decades (1, 10, 100...) rather than tightly fitted to the data." & vbLf & vbLf &
"Non-overlapping ADEV formula:" & vbLf &
"sigma(tau) = sqrt( sum( (Ybar[k+1] - Ybar[k])^2 ) / (2 x (M-1)) ), where Ybar[k] is the average of block k (blocks of length tau, M = N/tau blocks)." & vbLf & vbLf &
"Overlapping ADEV formula:" & vbLf &
"sigma(tau) = sqrt( sum( (Ybar[i+tau] - Ybar[i])^2 ) / (2 x (N-2tau+1)) ), using a sliding window instead of separate blocks." & vbLf & vbLf &
"Starting at the top-left (tau=1) and reading rightward: the closer the curve hugs its dashed 'Ideal' line, the more that stretch behaves like pure random noise - each step right is genuinely buying more stability." & vbLf & vbLf &
"Where the curve pulls away and rises above the dashed line, averaging longer has stopped helping - a flat stretch is a noise floor, a rising stretch is long-term drift making things worse." & vbLf & vbLf &
"Each device gets its own grey dotted 'Ideal' reference line, anchored to that device's own first plotted point, showing pure white-noise behaviour (a straight slope of -1/2 on the log-log axes)." & vbLf & vbLf &
"It's a reference, not a hard boundary - the real curve can dip below it too, which is just statistical scatter in the estimate, especially on the right where only a few independent samples remain to compare." & vbLf & vbLf &
"The 'Overlapping (smoother tail)' checkbox switches between non-overlapping ADEV (disjoint tau-length blocks - fewer pairs at large tau, so the tail can look noisy/jagged) and overlapping ADEV (a sliding window that reuses every sample many times over)." & vbLf & vbLf &
"Overlapping gives a much smoother tail from the same data, at the cost of the points no longer being fully statistically independent." & vbLf & vbLf &
"The 'Show MDEV' checkbox adds a second, dotted curve per device - Modified Allan Deviation (MDEV)." & vbLf & vbLf &
"MDEV applies an extra averaging stage that makes it react differently than ADEV to phase/timing noise specifically." & vbLf & vbLf &
"On its own, ADEV can't tell ordinary amplitude noise apart from phase noise - both just look like a similar falling slope." & vbLf & vbLf &
"If MDEV runs visibly steeper than its device's ADEV curve at short tau, that's a sign of phase noise ADEV alone wouldn't show." & vbLf & vbLf &
"MDEV formula:" & vbLf &
"Mod sigma(tau) = sqrt( sum( (second-difference sum over an m-sample window)^2 ) / (2 x tau^2 x m^2 x (N-3m+1)) ), where m = tau and the second difference is taken on x, the cumulative sum (integration) of the raw readings." & vbLf & vbLf &
"COMMON QUESTIONS" & vbLf &
"Why doesn't Allan Deviation match the recorded STDEV?" & vbLf & vbLf &
"That's expected, not a bug - STDEV and Allan Deviation at tau=1 are answering two different questions." & vbLf & vbLf &
"STDEV (the recorded/Data tab/Live Analysis figure) measures how far every individual reading sits from the overall mean of the whole run: sqrt(sum((Xi - Mean)^2) / (N-1)). If the reading drifts slowly over the logging session (thermal settling, reference aging, environmental changes), that drift adds to the spread away from the overall mean, and STDEV counts all of that as deviation, whether it's random noise or systematic drift." & vbLf & vbLf &
"Allan Deviation at tau=1 measures something narrower: the RMS of the difference between consecutive readings. Two back-to-back samples are barely affected by slow drift, so ADEV(tau=1) picks up almost purely the short-term, sample-to-sample (white) noise floor, filtered clean of slow drift." & vbLf & vbLf &
"Example: if ADEV(tau=1) reads ~0.2 ppm while recorded STDEV reads ~0.4 ppm and never drops below ~0.3 ppm, that gap is informative - it means the true random noise floor is around 0.2 ppm, and roughly half of what STDEV reports as variation is actually systematic drift, not noise." & vbLf & vbLf &
"STDEV alone can't separate genuinely noisy from drifting, and will always read equal to or higher than the ADEV noise floor whenever any drift is present. Check whether the Allan Deviation curve rises again at larger tau - that's the classic drift signature, and it's where the variability STDEV was counting shows up." & vbLf & vbLf &
"AVERAGING / NOISE / RANGE (per device)" & vbLf &
"The numeric box next to '- Avg.' sets how many points the raw Data trace itself is rolling-averaged over before being plotted (0 disables it, range 0-100). This smooths the Data trace directly, unlike Short Term Mean, which is a separate overlay trace and never alters Data itself." & vbLf & vbLf &
"'- RMS Noise' and '- Max-Min' are read-only figures calculated for whatever portion of the chart is currently visible/zoomed: RMS Noise is a noise calculation that accounts for drift over time, and Max-Min is the peak-to-peak spread of the visible data." & vbLf & vbLf &
"Line / Point switch that device's Data trace between a connected line and individual points." & vbLf & vbLf &
"PPM DEVIATION / TEMPCO" & vbLf &
"Enable PPM turns on a separate, live-recalculated PPM trace (distinct from the recorded 'PPM Deviation' checkbox trace above) for whichever device is selected by the Dev 1/Dev 2 radio buttons in the DEVICES panel." & vbLf & vbLf &
"PPM Deviation formula: (Value - Baseline Value) / Baseline Value x 1,000,000" & vbLf & vbLf &
"PPM/DegC (temperature coefficient) formula: PPM Deviation / (Temp - Baseline Temp)" & vbLf & vbLf &
"TEMP/HUM" & vbLf &
"Temp and Hum. show or hide the logged temperature and humidity traces. Temp/Hum Max. and Min. and Temp Avg. summarise the recorded values." & vbLf & vbLf &
"MISC." & vbLf &
"ToolTip Values - shows a tooltip with the exact value when hovering over a point on the chart." & vbLf & vbLf &
"Light Mode - switches the chart to a white background, better suited to printing than the default dark theme." & vbLf & vbLf &
"IMPORTANT" & vbLf &
"- All Dev.2 controls are automatically disabled for a single-device CSV - there's no need to manually hide them." & vbLf & vbLf &
"- Short Term Mean and Allan Deviation are both computed fresh from the raw VALUE column every time - they are not values that were written to the CSV during acquisition, and toggling them never changes the underlying log file." & vbLf & vbLf &
"- The recorded Mean/STDEV/SEM/Max Diff./PPM Deviation traces reflect whatever statistics were being calculated live at acquisition time, and depend on when Reset Stats was last pressed during logging."
    }

        ' Make headings bold.
        Dim headings() As String = {
        "PLAYBACK CHART",
        "LOADING A CSV",
        "DEVICES",
        "X-AXIS SCALE",
        "Y-AXIS SCALE",
        "NAVIGATION",
        "DEV 1 TRACES / DEV 2 TRACES",
        "SHORT TERM MEAN",
        "ALLAN DEVIATION",
        "COMMON QUESTIONS",
        "AVERAGING / NOISE / RANGE (per device)",
        "PPM DEVIATION / TEMPCO",
        "TEMP/HUM",
        "MISC.",
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

        ' Make just the equation part of each formula line bold - not the
        ' "Formula:" label and not any trailing explanatory clause - so it
        ' stands out from the surrounding explanatory prose without the
        ' whole sentence turning bold.
        Dim formulaLines() As String = {
        "sqrt( sum( (Xi - Mean)^2 ) / (N - 1) )",
        "STDEV / sqrt(N)",
        "Max - Min",
        "(Value - First Value) / First Value x 1,000,000",
        "mean of readings i-N+1 through i",
        "sigma(tau) = sqrt( sum( (Ybar[k+1] - Ybar[k])^2 ) / (2 x (M-1)) )",
        "sigma(tau) = sqrt( sum( (Ybar[i+tau] - Ybar[i])^2 ) / (2 x (N-2tau+1)) )",
        "Mod sigma(tau) = sqrt( sum( (second-difference sum over an m-sample window)^2 ) / (2 x tau^2 x m^2 x (N-3m+1)) )",
        "(Value - Baseline Value) / Baseline Value x 1,000,000",
        "PPM Deviation / (Temp - Baseline Temp)"
    }

        For Each formulaLine As String In formulaLines

            Dim start As Integer =
            txt.Text.IndexOf(formulaLine, StringComparison.Ordinal)

            If start >= 0 Then
                txt.Select(start, formulaLine.Length)
                txt.SelectionFont = New Font(txt.Font, FontStyle.Bold)
            End If

        Next

        ' Indent each section's body text (everything between one heading
        ' and the next) so it reads as clearly belonging under its
        ' heading. Uses SelectionIndent (a paragraph-level left margin)
        ' rather than literal leading spaces - spaces would only indent
        ' the first visual line of a wrapped paragraph, leaving wrapped
        ' continuation lines flush left and ragged.
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

End Class


