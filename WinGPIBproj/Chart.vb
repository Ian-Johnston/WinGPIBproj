' Playback chart control

Imports System.Windows.Forms.DataVisualization.Charting


Public Class Chart

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
    Dim inputvalueMeasurements() As Double

    Dim DEV1rollingAverageValues As New List(Of Double)         ' Create a list to store the rolling average values
    Dim DEV2rollingAverageValues As New List(Of Double)         ' Create a list to store the rolling average values
    Dim TEMProllingAverageValues As New List(Of Double)         ' Create a list to store the rolling average values
    Dim tempcounter As Integer = 0
    Dim tempTEMPcounter As Integer = 0

    Dim numberofmetadatalines As Integer = 0


    Private Sub Formtest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        ' Set Timer1 duration - Used for refresh of auto-Playback chart, i.e. auto reload of CSV
        Me.Timer1.Interval = 5000  ' 5secs
        Me.Timer1.Stop()

        DeviceName1.BackColor = Color.GreenYellow
        DeviceName2.BackColor = Color.Violet
        RadioButtonDev1.BackColor = Color.GreenYellow
        RadioButtonDev2.BackColor = Color.Violet
        RadioButtonPPMDev.BackColor = Color.White
        RadioButtonPPMTempo.BackColor = Color.White
        PlaybackTemp.BackColor = Color.Red
        PlaybackHum.BackColor = Color.DodgerBlue

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
        CSVdelimit = My.Settings.data29                 ' comma or semi-colon

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

        CheckBoxMedianV.Enabled = False
        CheckBoxMedianT.Enabled = False

        RadioButtonPPMDev.Checked = True
        RadioButtonPPMDev.Enabled = False
        RadioButtonPPMTempo.Enabled = False
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

        ' Os Version
        'PlaybackstrPath = String.Format("{0}", Environment.CurrentDirectory)    ' folder where app is running
        PlaybackstrPath = "C:\Users\" & String.Format("{0}", Environment.UserName) & "\Documents\WinGPIBdata"    ' users data folder

        ' Check Os and set data folder path accordingly (change the same in Formtest.vb)
        ' https://docs.microsoft.com/en-us/windows/win32/sysinfo/operating-system-version
        'Dim osVer As Version = Environment.OSVersion.Version    ' Operating system
        'If osVer.Major = 6 And osVer.Minor = 1 Then     ' 6.1 = Win7
        'PlaybackstrPath = "C:\Users\" & String.Format("{0}", Environment.UserName) & "\Documents\WinGPIBdata"
        'End If
        'If osVer.Major = 6 And osVer.Minor = 2 Then     ' 6.3 = Win8
        'PlaybackstrPath = "C:\Users\" & String.Format("{0}", Environment.UserName) & "\Documents\WinGPIBdata"
        'End If
        'If osVer.Major = 6 And osVer.Minor = 3 Then     ' 6.3 = Win8.1
        'PlaybackstrPath = "C:\Users\" & String.Format("{0}", Environment.UserName) & "\Documents\WinGPIBdata"
        'End If
        'If osVer.Major = 10 And osVer.Minor = 0 Then     ' 10.0 = Win10/11
        'PlaybackstrPath = "C:\Users\" & String.Format("{0}", Environment.UserName) & "\Documents\WinGPIBdata"
        'End If

        CSVfileok = False

        ' Clear the screen and set the load CSV message
        ChartOffReadyForCSV()

        ' Chart2 initialize
        Chart2.ChartAreas(0).AxisY.LabelStyle.Enabled = True
        Chart2.ChartAreas(0).AxisX.MajorTickMark.Enabled = True
        Chart2.ChartAreas(0).AxisX.Interval = 95
        Chart2.ChartAreas(0).AxisY.LabelStyle.Font = New Font("Verdana", 8)  ' change x-axis label font style
        Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "{000.0000000}"
        Chart2.ChartAreas(0).AxisY.MajorTickMark.Enabled = True
        Chart2.ChartAreas(0).AxisX.MinorTickMark.Enabled = False
        Chart2.ChartAreas(0).AxisY.MinorTickMark.Enabled = False
        Chart2.ChartAreas(0).AxisX.MajorGrid.Enabled = True
        Chart2.ChartAreas(0).AxisY.MajorGrid.Enabled = True
        Chart2.ChartAreas(0).AxisX.MinorGrid.Enabled = True
        Chart2.ChartAreas(0).AxisY.MinorGrid.Enabled = True
        Chart2.ChartAreas(0).AxisX.MajorGrid.LineColor = Color.FromArgb(255, 85, 85, 85)
        Chart2.ChartAreas(0).AxisY.MajorGrid.LineColor = Color.FromArgb(255, 85, 85, 85)
        Chart2.ChartAreas(0).AxisX.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
        Chart2.ChartAreas(0).AxisY.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
        Chart2.DataBindTable(gChartPlayback)
        Chart2.Series(0).ChartType = 2
        Chart2.Series.Clear()
        Chart2.ChartAreas(0).BorderWidth = 1
        Chart2.Series.Add("Device 1")
        Chart2.Series.Add("Device 2")
        Chart2.Series.Add("Temperature")
        Chart2.Series.Add("Humidity")
        Chart2.Series.Add("PPM Dev 1")

        CheckDev1Line.Checked = True
        CheckDev1Point.Checked = False
        CheckDev2Line.Checked = True
        CheckDev2Point.Checked = False

        Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart2.Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart2.Series(2).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart2.Series(3).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart2.Series(4).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart2.Series(0).YValueType = DataVisualization.Charting.ChartValueType.Single
        Chart2.Legends(0).Enabled = False
        Chart2.ChartAreas(0).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        'Chart2.ChartAreas(0).AxisY.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount    ' this is set dynamically
        Chart2.ChartAreas(0).AxisY.LabelAutoFitStyle = DataVisualization.Charting.LabelAutoFitStyles.DecreaseFont 'default is staggered
        Chart2.ChartAreas(0).AxisX.IntervalOffset = 0

        ' Straight colours
        Chart2.Series(0).Color = Color.GreenYellow
        Chart2.Series(1).Color = Color.Violet
        Chart2.Series(2).Color = Color.Red
        Chart2.Series(3).Color = Color.DodgerBlue
        Chart2.Series(4).Color = Color.White

        Chart2.ChartAreas(0).AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot
        Chart2.ChartAreas(0).AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot
        Chart2.ChartAreas(0).AxisX.LabelStyle.Enabled = False   'disable X-axis scale
        Chart2.ChartAreas(0).AxisY2.MajorTickMark.Enabled = True
        Chart2.ChartAreas(0).AxisY2.MinorTickMark.Enabled = False
        Chart2.ChartAreas(0).AxisY2.LabelAutoFitStyle = DataVisualization.Charting.LabelAutoFitStyles.DecreaseFont 'default is staggered
        Chart2.ChartAreas(0).AxisY2.Interval = 1

        Chart2.ChartAreas(0).AxisY2.MajorGrid.LineColor = Color.FromArgb(100, 85, 85, 85)
        Chart2.ChartAreas(0).AxisY2.MinorGrid.LineColor = Color.FromArgb(100, 85, 85, 85)

        ' Temperature
        Chart2.Series(2).YAxisType = DataVisualization.Charting.AxisType.Secondary
        Chart2.ChartAreas(0).AxisY2.Enabled = True
        Chart2.ChartAreas(0).AxisY2.Minimum = 15
        Chart2.ChartAreas(0).AxisY2.Maximum = 50
        Chart2.ChartAreas(0).AxisY2.Enabled = DataVisualization.Charting.AxisEnabled.True
        Chart2.ChartAreas(0).AxisY2.LabelStyle.Enabled = True

        ' Humidity
        Chart2.Series(3).YAxisType = DataVisualization.Charting.AxisType.Secondary

        ' CSV file format
        dataTable1.Columns.Add("INDEX", GetType(Integer))
        dataTable1.Columns.Add("DEVICE", GetType(String))
        dataTable1.Columns.Add("DATETIME", GetType(String))
        dataTable1.Columns.Add("VALUE", GetType(Double))
        dataTable1.Columns.Add("TEMP", GetType(Double))
        dataTable1.Columns.Add("HUM", GetType(Double))

        ' Additional columns add to datatable
        dataTable1.Columns.Add("PPM", GetType(Double))

        'Dim windowSize As Integer = 100
        RMSwindow.Text = "100"

        LabelTempC.Text = My.Settings.data324
        LabelHum.Text = My.Settings.data325
        RadioButtonPPMTempo.Text = "PPM/" & My.Settings.data324

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

        dataTable1.Clear()
        For i As Integer = 0 To 4
            Chart2.Series(i).Points.Clear()
        Next

        ' Read CSV file and process in one pass
        Dim lines As List(Of String) = IO.File.ReadAllLines(filePlayback).ToList()

        ' Initialize variables
        CSVdelimit = ""
        MetadataChart.Text = ""
        Dim separator As String = "------------------------------"
        Dim isFirstGroup As Boolean = True
        Dim previousLineIsMetadata As Boolean = False
        numberlinesCSV = 0

        Dim metadataBuilder As New System.Text.StringBuilder()
        Dim rowsToAdd As New List(Of DataRow)()

        ' Single pass to analyze, load data, and collect metadata
        For Each line As String In lines
            If line.TrimStart().StartsWith("//") Then
                ' Process metadata
                If Not previousLineIsMetadata AndAlso Not isFirstGroup Then
                    metadataBuilder.AppendLine(separator)
                End If
                metadataBuilder.AppendLine(line.Substring(2))
                previousLineIsMetadata = True
            Else
                previousLineIsMetadata = False
                isFirstGroup = False

                ' Determine delimiter (comma or semicolon) if not already set
                If CSVdelimit = "" Then
                    Dim commaCount As Integer = line.Split(","c).Length - 1
                    Dim semicolonCount As Integer = line.Split(";"c).Length - 1
                    If commaCount >= 5 Then
                        CSVdelimit = ","
                    ElseIf semicolonCount >= 5 Then
                        CSVdelimit = ";"
                    End If
                End If

                ' Skip blank lines
                If String.IsNullOrWhiteSpace(line) Then Continue For

                ' Process the data line
                Dim values As String() = line.Split(CSVdelimit)
                If Not IsNumeric(values(0)) Then
                    Dialog2.Warning1 = "Inconsistent CSV - Invalid data format detected"
                    Dialog2.Warning2 = "Each data line in the CSV should start with a number"
                    Dialog2.Warning3 = "Please fix and try again."
                    Dialog2.ShowDialog(Me)
                    ChartOffReadyForCSV()
                    Return
                End If

                ' Add data to dataTable1
                Dim row As DataRow = dataTable1.NewRow()
                row.ItemArray = values
                rowsToAdd.Add(row)

                ' Count valid data lines
                numberlinesCSV += 1
            End If
        Next

        ' Update MetadataChart once
        MetadataChart.Text = metadataBuilder.ToString()

        numberofmetadatalines = lines.Count - numberlinesCSV
        numberlinesCSV = lines.Count

        ' Check if CSV has enough lines
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

        ' Check for missing delimiters
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

        ' Add rows to dataTable1 in bulk
        dataTable1.BeginLoadData()
        For Each row As DataRow In rowsToAdd
            dataTable1.Rows.Add(row)
        Next
        dataTable1.EndLoadData()

        ' Check if dual devices exist and update UI
        Devname1 = dataTable1.Rows(0).ItemArray(1).ToString()
        Devname2 = dataTable1.Rows(1).ItemArray(1).ToString()

        If Devname1 = Devname2 Then
            DualDev = False
            DeviceName1.Text = Devname1
            DeviceName2.Text = ""
            DisableDualDeviceControls()
        Else
            DualDev = True
            DeviceName1.Text = Devname1
            DeviceName2.Text = Devname2
            EnableDualDeviceControls()
        End If

        If Not DualDev Then
            MedianValueCSV = dataTable1.Rows(0).ItemArray(3).ToString()
            MedianTempCSV = dataTable1.Rows(0).ItemArray(4).ToString()
        ElseIf RadioButtonDev1.Checked Then
            MedianValueCSV = dataTable1.Rows(0).ItemArray(3).ToString()
            MedianTempCSV = dataTable1.Rows(0).ItemArray(4).ToString()
        ElseIf RadioButtonDev2.Checked Then
            MedianValueCSV = dataTable1.Rows(1).ItemArray(3).ToString()
            MedianTempCSV = dataTable1.Rows(1).ItemArray(4).ToString()
        End If

        If DualDev = False Then
            Dim formatdata As String = "yyyy-MM-dd_HH:mm:ss"
            ' Convert the first and third DATETIME entries to DateTime objects using the custom format
            Dim DateTime8th As DateTime = DateTime.ParseExact(dataTable1.Rows(8)("DATETIME").ToString(), formatdata, System.Globalization.CultureInfo.InvariantCulture)
            Dim DateTime9th As DateTime = DateTime.ParseExact(dataTable1.Rows(9)("DATETIME").ToString(), formatdata, System.Globalization.CultureInfo.InvariantCulture)
            Dim timeDifference As TimeSpan = DateTime9th.Subtract(DateTime8th)            ' Calculate the difference as a TimeSpan
            SampleRateSecs.Text = timeDifference.TotalSeconds            ' Get the difference in seconds
        Else
            Dim formatdata As String = "yyyy-MM-dd_HH:mm:ss"
            ' Convert the first and third DATETIME entries to DateTime objects using the custom format
            Dim DateTime8th As DateTime = DateTime.ParseExact(dataTable1.Rows(8)("DATETIME").ToString(), formatdata, System.Globalization.CultureInfo.InvariantCulture)
            Dim DateTime10th As DateTime = DateTime.ParseExact(dataTable1.Rows(10)("DATETIME").ToString(), formatdata, System.Globalization.CultureInfo.InvariantCulture)
            Dim timeDifference As TimeSpan = DateTime10th.Subtract(DateTime8th)            ' Calculate the difference as a TimeSpan
            SampleRateSecs.Text = timeDifference.TotalSeconds            ' Get the difference in seconds
        End If

        ' Call methods to finalize the chart
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
    End Sub


    Private Sub RefreshFile_Click(sender As Object, e As EventArgs) Handles RefreshFile.Click

        RefreshPlaybackCSVFile()
        'ShowAll()
        PleaseLoadCSV.Visible = False

    End Sub


    Private Sub RefreshPlaybackCSVFile()

        ' A slimmed down version of loading the CSV file again as it is written to externally.

        If (CSVfilenamePlayback.Text <> "" And ChartLoaded = True And CSVfileok = True) Then

            ' Reset various vars
            CurrentPos = 0
            TargetPos = 49
            RangeReqd = 49
            EndRange = 500
            CentreRange = 0

            Loading.Visible = True
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

            CheckPathCSVfile()
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
                dataTable1.Rows.Add(line.Split(CSVdelimit))
            Next

            FilterDeviceName1()
            FilterDeviceName2()
            FilterTempDevice1()
            FilterHumDevice1()
            FilterGenPPMDevice1()
            FilterGenPPMDevice2()

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

            ' Ratio of counts to mins
            Dim ticklabelgridratio As Integer = ((maxXc - minXc) / Xscaletotal.Text)

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

            ' Ratio of counts to mins
            Dim ticklabelgridratio As Integer = ((maxXc - minXc) / Xscaletotal.Text)

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
        Chart2.ChartAreas(0).AxisY.MajorGrid.Interval = (YaxisMaximum.Text - YaxisMinimum.Text) / 8
        Chart2.ChartAreas(0).AxisY.MinorGrid.Interval = (YaxisMaximum.Text - YaxisMinimum.Text) / 32
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










    Private Sub CheckPathCSVfile()

        ' With the data now in the datatable now process it.

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
            FilterTempDevice1()
            FilterHumDevice1()

            DevicesMinMax()         ' get device min & max values



            If CheckBoxMaxMin.Checked Then

                ' Initialize min and max values
                'Dim maxValue As Double = Double.MinValue
                'Dim minValue As Double = Double.MaxValue

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

                ' Configure Y-axis based on LogYaxis checkbox
                If LogYaxis.Checked Then
                    Chart2.ChartAreas(0).AxisY.IsLogarithmic = True
                    ButtonShiftUp.Enabled = False
                    ButtonShiftDn.Enabled = False
                    Chart2.ChartAreas(0).AxisY.Minimum = Math.Round(YminFromDT, 7)
                    YaxisMinimum.Text = YminFromDT.ToString()
                Else
                    Chart2.ChartAreas(0).AxisY.IsLogarithmic = False
                    ButtonShiftUp.Enabled = True
                    ButtonShiftDn.Enabled = True
                    Chart2.ChartAreas(0).AxisY.Maximum = Math.Round(YmaxFromDT, 7)
                    Chart2.ChartAreas(0).AxisY.Minimum = Math.Round(YminFromDT, 7)
                    YaxisMaximum.Text = YmaxFromDT.ToString()
                    YaxisMinimum.Text = YminFromDT.ToString()

                    Dim result As Double = (YmaxFromDT - YminFromDT) / 32
                    YaxisPerDiv.Text = result.ToString("#0.000000000")
                End If

                ' Set axis interval
                Chart2.ChartAreas(0).AxisY.Interval = (YmaxFromDT - YminFromDT) / 32

            End If







            'Temp/Hum scale setting
            Chart2.ChartAreas(0).AxisY2.Minimum = CDbl(Val(ChartScaleMin.Text))
            Chart2.ChartAreas(0).AxisY2.Maximum = CDbl(Val(ChartScaleMax.Text))

            Chart2.ChartAreas(0).AxisY2.Interval = (Val(ChartScaleMax.Text) - Val(ChartScaleMin.Text)) / 32
            Chart2.ChartAreas(0).AxisY2.LabelStyle.Format = "00.0"



            ' Generate PPM column in table from data
            If (CheckBoxPPMenable.Checked = True) Then

                ' Set PPM scale vars for calc
                Dim ppmscalerange As Double = PPMscalerangeentry.Text

                If ppmscalerange > 198 Then
                    ppmscalerange = 198
                    PPMscalerangeentry.Text = "198"
                End If
                If ppmscalerange < 0.2 Then
                    ppmscalerange = 0.2
                    PPMscalerangeentry.Text = "0.2"
                End If

                ppmscalerangebit = (ppmscalerange / 2) / 12


                ' Calculate PPM for given row
                ' Get initial value/Temp from CSV if selected
                If (CheckBoxMedianV.Checked = False) Then
                    medianvalued = CDbl(Val(MedianValue.Text))
                Else
                    medianvalued = MedianValueCSV
                    MedianValue.Text = MedianValueCSV
                End If

                If (CheckBoxMedianT.Checked = False) Then
                    mediantempd = CDbl(Val(MedianTemp.Text))
                Else
                    mediantempd = MedianTempCSV
                    MedianTemp.Text = MedianTempCSV
                End If

                Dim PPMdevice As String = ""
                Dim YaxisMaximumVal As Double = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)
                Dim YaxisMinimumVal As Double = Math.Round((CDbl(Val(YaxisMinimum.Text))), 7)
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
                            calcppmvalue = Vchange * 1000000


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


                ' Add PPM to chart for device specified only
                Dim selectedRows5() As DataRow = dataTable1.Select("DEVICE ='" & PPMdevice & "'")
                'Add filtered data to series
                For Each dr As DataRow In selectedRows5
                    Chart2.Series(4).Points.AddXY(dr("DEVICE"), dr("PPM"))
                Next

                PrintYscale()
                Yscaletidy()      ' Tidy up X-scale annotations on graph in order to keep length same irrespective of numerical data and No. DP's

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










    Function CalculateRollingAverage(ByVal inputvalue As Double, numDataPoints As Integer) As Double                     ' rolling average

        If numDataPoints < 1 Then
            numDataPoints = 1
        End If

        If numDataPoints > 500 Then
            numDataPoints = 500
        End If

        ' Initialize the inputvalueMeasurements array if not already done
        If inputvalueMeasurements Is Nothing OrElse inputvalueMeasurements.Length <> numDataPoints Then
            ReDim inputvalueMeasurements(numDataPoints - 1)
        End If

        ' Update the array with the latest values
        For i As Integer = 0 To numDataPoints - 2
            inputvalueMeasurements(i) = inputvalueMeasurements(i + 1)
        Next
        inputvalueMeasurements(numDataPoints - 1) = inputvalue

        ' Calculate the rolling average of the values
        Dim sumValue As Double = 0.0
        Dim numValidPoints As Integer = 0 ' To track the number of valid data points (non-zero)
        For i As Integer = 0 To numDataPoints - 1
            If inputvalueMeasurements(i) <> 0 Then
                sumValue += inputvalueMeasurements(i)
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
                    dataTable1.Rows.Add(line.Split(CSVdelimit))
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterTempDevice1()
                FilterHumDevice1()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()

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
                    dataTable1.Rows.Add(line.Split(CSVdelimit))
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterTempDevice1()
                FilterHumDevice1()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()

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
                    dataTable1.Rows.Add(line.Split(CSVdelimit))
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterTempDevice1()
                FilterHumDevice1()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()

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
                    dataTable1.Rows.Add(line.Split(CSVdelimit))
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterTempDevice1()
                FilterHumDevice1()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()

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
                    dataTable1.Rows.Add(line.Split(CSVdelimit))
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterTempDevice1()
                FilterHumDevice1()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()

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
                    dataTable1.Rows.Add(line.Split(CSVdelimit))
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterTempDevice1()
                FilterHumDevice1()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()

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

                Dim Playbacknewmax As Double = CDbl(Val(YaxisMaximum.Text))
                Dim Playbacknewmin As Double = CDbl(Val(YaxisMinimum.Text))

                'Playbacknewmax = Playbacknewmax + ((Playbacknewmax - Playbacknewmin) / 20)
                Playbacknewmax += (Playbacknewmax - Playbacknewmin) / 20
                YaxisMaximum.Text = Playbacknewmax
                YaxisMaximum.Text = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)

                'Playbacknewmin = Playbacknewmin + ((Playbacknewmax - Playbacknewmin) / 20)
                Playbacknewmin += (Playbacknewmax - Playbacknewmin) / 20
                YaxisMinimum.Text = Playbacknewmin
                YaxisMinimum.Text = Math.Round((CDbl(Val(YaxisMinimum.Text))), 7)

                Dim result As Double = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 32
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
                    dataTable1.Rows.Add(line.Split(CSVdelimit))
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterTempDevice1()
                FilterHumDevice1()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()

                'Manual device scale setting
                Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((CDbl(Val(YaxisMinimum.Text))), 7)
                Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)  ' was 7
                Chart2.ChartAreas(0).AxisY.Interval = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 20

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

                Dim Playbacknewmax As Double = CDbl(Val(YaxisMaximum.Text))
                Dim Playbacknewmin As Double = CDbl(Val(YaxisMinimum.Text))

                'Playbacknewmax = Playbacknewmax - ((Playbacknewmax - Playbacknewmin) / 20)
                Playbacknewmax -= (Playbacknewmax - Playbacknewmin) / 20
                YaxisMaximum.Text = Playbacknewmax
                YaxisMaximum.Text = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)

                'Playbacknewmin = Playbacknewmin - ((Playbacknewmax - Playbacknewmin) / 20)
                Playbacknewmin -= (Playbacknewmax - Playbacknewmin) / 20
                'If (Playbacknewmin < 0) Then
                'Playbacknewmin = 0
                'End If
                YaxisMinimum.Text = Playbacknewmin
                YaxisMinimum.Text = Math.Round((CDbl(Val(YaxisMinimum.Text))), 7)

                Dim result As Double = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 32
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
                    dataTable1.Rows.Add(line.Split(CSVdelimit))
                Next

                FilterDeviceName1()
                FilterDeviceName2()
                FilterTempDevice1()
                FilterHumDevice1()
                FilterGenPPMDevice1()
                FilterGenPPMDevice2()

                'Manual device scale setting
                Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((CDbl(Val(YaxisMinimum.Text))), 7)
                Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)   ' was 7
                Chart2.ChartAreas(0).AxisY.Interval = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 20

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
            Dim x2 As String = CStr(YaxisMaximum.Text)   ' 0.9999995 or 999.0000000 etc

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
        Ymin = YaxisMinimum.Text
        Ymin += 0.00000000001
        Ymax = YaxisMaximum.Text
        Ymax -= (Ymax - Ymin) / 10  ' shift by a tenth
        'If (Ymin < 0) Then
        'Ymin = 0.0000001
        'End If
        'If (Ymin < Ymax And Ymin >= 0) Then
        If (Ymin < Ymax) Then
            YaxisMinimum.Text = Format(Ymin, "#0.00000000")
            YaxisMaximum.Text = Format(Ymax, "#0.00000000")

            Dim result As Double = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")

            Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((CDbl(Val(YaxisMinimum.Text))), 7)
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)   ' was 7
            Chart2.ChartAreas(0).AxisY.Interval = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 20
        End If

        YaxisCheck1.Checked = False
        YaxisCheck2.Checked = False
        YaxisCheck3.Checked = False
        YaxisCheck4.Checked = False
    End Sub


    ' Manually adjust Y value - max up
    Private Sub ButtonYmaxInc_Click(sender As Object, e As EventArgs) Handles ButtonYmaxInc.Click
        Ymin = YaxisMinimum.Text
        Ymin += 0.00000000001
        Ymax = YaxisMaximum.Text
        Ymax += (Ymax - Ymin) / 10  ' shift by a tenth

        '        If (Ymin < 0) Then
        '        Ymin = 0.0000001
        '        End If
        'If (Ymin < Ymax And Ymin >= 0) Then
        If (Ymin < Ymax) Then
            YaxisMinimum.Text = Format(Ymin, "#0.00000000")
            YaxisMaximum.Text = Format(Ymax, "#0.00000000")

            Dim result As Double = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")

            Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((CDbl(Val(YaxisMinimum.Text))), 7)
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)   ' was 7
            Chart2.ChartAreas(0).AxisY.Interval = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 20
        End If

        YaxisCheck1.Checked = False
        YaxisCheck2.Checked = False
        YaxisCheck3.Checked = False
        YaxisCheck4.Checked = False
    End Sub


    ' Manually adjust Y value - min down
    Private Sub ButtonYminDec_Click(sender As Object, e As EventArgs) Handles ButtonYminDec.Click
        Ymin = YaxisMinimum.Text
        Ymax = YaxisMaximum.Text
        Ymin -= (Ymax - Ymin) / 10  ' shift by a tenth

        '       If (Ymin < 0) Then
        '       Ymin = 0.0000001
        '       End If
        'If (Ymin < Ymax And Ymin >= 0) Then
        If (Ymin < Ymax) Then
            YaxisMinimum.Text = Format(Ymin, "#0.00000000")
            YaxisMaximum.Text = Format(Ymax, "#0.00000000")

            Dim result As Double = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000") '

            Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((CDbl(Val(YaxisMinimum.Text))), 7)
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)   ' was 7
            Chart2.ChartAreas(0).AxisY.Interval = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 20
        End If

        YaxisCheck1.Checked = False
        YaxisCheck2.Checked = False
        YaxisCheck3.Checked = False
        YaxisCheck4.Checked = False
    End Sub


    ' Manually adjust Y value - max up
    Private Sub ButtonYminInc_Click(sender As Object, e As EventArgs) Handles ButtonYminInc.Click
        Ymin = YaxisMinimum.Text
        Ymax = YaxisMaximum.Text
        Ymin += (Ymax - Ymin) / 10  ' shift by a tenth
        '        If (Ymin < 0) Then
        '        Ymin = 0.0000001
        '        End If
        'If (Ymin < Ymax And Ymin >= 0) Then
        If (Ymin < Ymax) Then
            YaxisMinimum.Text = Format(Ymin, "#0.0000000000")
            YaxisMaximum.Text = Format(Ymax, "#0.0000000000")

            Dim result As Double = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")

            Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((CDbl(Val(YaxisMinimum.Text))), 7)
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)   ' was 7
            Chart2.ChartAreas(0).AxisY.Interval = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 20
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
                    Dim Dev1rollingAverageValue As Double = CalculateRollingAverage(variancevalue, Val(DEV1avg.Text))

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
                    Dim Dev2rollingAverageValue As Double = CalculateRollingAverage(variancevalue, Val(DEV2avg.Text))

                    ' Store the rolling average value in the list
                    DEV2rollingAverageValues.Add(Dev2rollingAverageValue)

                    ' Add the data point with the rolling average value to the chart's series
                    Chart2.Series(1).Points.AddXY(dr("DEVICE"), Dev2rollingAverageValue)

                    'Console.WriteLine("Rolling Average Value for Device 1 " & Dev1rollingAverageValue)

                Next
            End If

        End If


    End Sub


    Private Sub FilterTempDevice1()

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
                    Dim TemprollingAverageValue As Double = CalculateRollingAverage(variancevalue, Val(TEMPavg.Text))

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

        ' Filter Humidity from Device 1
        If (PlaybackHum.Checked = True) Then
            Dim selectedRows4() As DataRow = dataTable1.Select("DEVICE ='" & DeviceName1.Text & "'")
            'Add filtered data to series
            For Each dr As DataRow In selectedRows4
                Chart2.Series(3).Points.AddXY(dr("DEVICE"), dr("HUM"))
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
                    Dev1MaxMin.Text = (Dev1MaxD - Dev1MinD) * 1000000
                End If
                If (CheckX1000.Checked = True) Then
                    Dev1MaxMin.Text = (Dev1MaxD - Dev1MinD) * 1000
                End If
            Else
                Dev1MaxMin.Text = CDec(Dev1MaxD - Dev1MinD)     ' e-notation to decimal
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
                    Dev2MaxMin.Text = (Dev2MaxD - Dev2MinD) * 1000000
                End If
                If (CheckX1000.Checked = True) Then
                    Dev2MaxMin.Text = (Dev2MaxD - Dev2MinD) * 1000
                End If
            Else
                Dev2MaxMin.Text = CDec(Dev2MaxD - Dev2MinD)     ' e-notation to decimal
            End If

        End If

    End Sub




    Private Sub Screenshot_Click(sender As Object, e As EventArgs) Handles Screenshot.Click

        FormBorderStyle = BorderStyle.None ' Set control's border style none

        ' Force the form to repaint before taking the screenshot
        Update()

        ' Capture entire form or graph only
        If CheckBoxGraphOnly.Checked = True Then

            ' Define the rectangle representing the region to capture
            Dim captureRect As New Rectangle(60, 190, 1350, 695) ' Adjust these values as needed

            ' Create a bitmap to store the screenshot
            Using bmp As New Bitmap(captureRect.Width, captureRect.Height)
                Using g As Graphics = Graphics.FromImage(bmp)
                    ' Capture the specified region of the form
                    g.CopyFromScreen(Me.PointToScreen(captureRect.Location), Point.Empty, captureRect.Size)
                End Using

                ' Save the screenshot
                bmp.Save(PlaybackstrPath & "\" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & "_Log.png", Drawing.Imaging.ImageFormat.Png)
            End Using

        Else

            ' Capture the entire form as an image
            Dim bmp As New Bitmap(Me.Width, Me.Height)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.CopyFromScreen(Me.PointToScreen(New Point(0, 0)), New Point(0, 0), Me.Size)
            End Using

            ' Save the screenshot
            bmp.Save(PlaybackstrPath & "\" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & "_Log.png", Drawing.Imaging.ImageFormat.Png)
        End If



        FormBorderStyle = BorderStyle.FixedSingle ' Set control's border style none

        Dialog2.Warning1 = "Screenshot saved as:"
        Dialog2.Warning2 = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & "_Log.png"
        Dialog2.Warning3 = ""
        'Dialog2.Show() ' this method positions anywhere!
        Dialog2.ShowDialog(Me)  ' this method positions center of the parent form and requires hitting OK to return back to the parent

    End Sub







    Sub GetMinMaxScales()
        Dim axisYMin, axisYMax As Double
        Dim range As Double
        Dim interval As Double

        ' Calculate the axis Y minimum and maximum based on conditions
        If CheckBoxMaxMin.Checked Then
            axisYMax = Math.Round(CDbl(YmaxFromDT), 7)
            axisYMin = Math.Round(CDbl(YminFromDT), 7)

            If LogYaxis.Checked AndAlso (axisYMin - ((YmaxFromDT - YminFromDT) / 10) < 0) Then
                axisYMin = Math.Round(YminFromDT, 7)
            End If

            Chart2.ChartAreas(0).AxisY.Maximum = axisYMax
            Chart2.ChartAreas(0).AxisY.Minimum = axisYMin
            YaxisMaximum.Text = axisYMax.ToString()
            YaxisMinimum.Text = axisYMin.ToString()

        Else
            axisYMax = Math.Round(CDbl(YaxisMaximum.Text), 7)
            axisYMin = Math.Round(CDbl(YaxisMinimum.Text), 7)
            range = axisYMax - axisYMin
            interval = range / 20

            Chart2.ChartAreas(0).AxisY.Maximum = axisYMax
            Chart2.ChartAreas(0).AxisY.Minimum = axisYMin
            Chart2.ChartAreas(0).AxisY.Interval = interval
            YaxisMaximum.Text = axisYMax.ToString()
            YaxisMinimum.Text = axisYMin.ToString()
        End If

        ' Update Y axis per division
        range = axisYMax - axisYMin
        YaxisPerDiv.Text = (range / 32).ToString("#0.000000000")

        ' Temp/Hum scale setting
        Dim tempHumMin, tempHumMax As Double
        tempHumMin = CDbl(ChartScaleMin.Text)
        tempHumMax = CDbl(ChartScaleMax.Text)
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

    Private Sub LogYaxis_CheckedChanged(sender As Object, e As EventArgs) Handles LogYaxis.CheckedChanged

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
            My.Settings.data187 = YaxisMaximum.Text
            My.Settings.data188 = YaxisMinimum.Text
        End If

        If (YaxisCheck2.Checked = True) Then
            My.Settings.data189 = YaxisMaximum.Text
            My.Settings.data190 = YaxisMinimum.Text
        End If

        If (YaxisCheck3.Checked = True) Then
            My.Settings.data191 = YaxisMaximum.Text
            My.Settings.data192 = YaxisMinimum.Text
        End If

        If (YaxisCheck4.Checked = True) Then
            My.Settings.data193 = YaxisMaximum.Text
            My.Settings.data194 = YaxisMinimum.Text
        End If

    End Sub

    Private Sub YaxisLoad_Click(sender As Object, e As EventArgs) Handles YaxisLoad.Click

        If (YaxisCheck1.Checked = True) Then
            YaxisMaximum.Text = My.Settings.data187
            YaxisMinimum.Text = My.Settings.data188

            Dim result As Double = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")


            RefreshPlaybackCSVFile()
            Chart2.ChartAreas(0).AxisY.Interval = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 20
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)
        End If

        If (YaxisCheck2.Checked = True) Then
            YaxisMaximum.Text = My.Settings.data189
            YaxisMinimum.Text = My.Settings.data190

            Dim result As Double = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")


            RefreshPlaybackCSVFile()
            Chart2.ChartAreas(0).AxisY.Interval = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 20
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)
        End If

        If (YaxisCheck3.Checked = True) Then
            YaxisMaximum.Text = My.Settings.data191
            YaxisMinimum.Text = My.Settings.data192

            Dim result As Double = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")


            RefreshPlaybackCSVFile()
            Chart2.ChartAreas(0).AxisY.Interval = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 20
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)
        End If

        If (YaxisCheck4.Checked = True) Then
            YaxisMaximum.Text = My.Settings.data193
            YaxisMinimum.Text = My.Settings.data194

            Dim result As Double = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 32
            YaxisPerDiv.Text = result.ToString("#0.000000000")


            RefreshPlaybackCSVFile()
            Chart2.ChartAreas(0).AxisY.Interval = (Val(YaxisMaximum.Text) - Val(YaxisMinimum.Text)) / 20
            Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 7)
        End If





        'Chart2.ChartAreas(0).AxisY.Minimum = Math.Round((CDbl(Val(YaxisMinimum.Text))), 7)
        'Chart2.ChartAreas(0).AxisY.Maximum = Math.Round((CDbl(Val(YaxisMaximum.Text))), 1)
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

        'If Not String.IsNullOrEmpty(PPMscalerangeentry.Text) Then
        'RefreshPlaybackCSVFile()
        'End If

        ' checks for numbers and that it's not empty before applying

        Dim userInput As String = PPMscalerangeentry.Text.Trim()
        Dim isNumeric As Boolean = Integer.TryParse(userInput, Nothing)

        If Not String.IsNullOrEmpty(userInput) AndAlso isNumeric Then
            RefreshPlaybackCSVFile()
        End If

    End Sub

    Private Sub ChartScaleMax_TextChanged(sender As Object, e As EventArgs) Handles ChartScaleMax.TextChanged

        Dim userInput As String = ChartScaleMax.Text.Trim()
        Dim isNumeric As Boolean = Integer.TryParse(userInput, Nothing)

        If Not String.IsNullOrEmpty(userInput) AndAlso isNumeric AndAlso Val(ChartScaleMax.Text) > Val(ChartScaleMin.Text) Then
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
            End If

        End If

    End Sub

    Private Sub ChartScaleMin_TextChanged(sender As Object, e As EventArgs) Handles ChartScaleMin.TextChanged

        Dim userInput As String = ChartScaleMin.Text.Trim()
        Dim isNumeric As Boolean = Integer.TryParse(userInput, Nothing)

        If Not String.IsNullOrEmpty(userInput) AndAlso isNumeric AndAlso Val(ChartScaleMax.Text) > Val(ChartScaleMin.Text) Then
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
                    RMSaverageDev1.Text = $"{rmsNoise:F10}" * 1000000
                End If
                If (CheckX1000.Checked = True) Then
                    RMSaverageDev1.Text = $"{rmsNoise:F10}" * 1000
                End If
            Else
                RMSaverageDev1.Text = $"{rmsNoise:F10}"
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
                    RMSaverageDev2.Text = $"{rmsNoise:F10}" * 1000000
                End If
                If (CheckX1000.Checked = True) Then
                    RMSaverageDev2.Text = $"{rmsNoise:F10}" * 1000
                End If
            Else
                RMSaverageDev2.Text = $"{rmsNoise:F10}"
            End If

        End If

    End Sub

    Private Sub RMSwindow_TextChanged(sender As Object, e As EventArgs) Handles RMSwindow.TextChanged

        AverageNoise()

    End Sub


End Class


