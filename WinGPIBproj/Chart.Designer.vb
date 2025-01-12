<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Chart
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Chart))
        Me.Chart2 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.ButtonScrollLeft = New System.Windows.Forms.Button()
        Me.ButtonScrollRight = New System.Windows.Forms.Button()
        Me.CurrentPosition = New System.Windows.Forms.TextBox()
        Me.TargetPosition = New System.Windows.Forms.TextBox()
        Me.RangeRequired = New System.Windows.Forms.TextBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.CSVfilenamePlayback = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CSVfileLines = New System.Windows.Forms.TextBox()
        Me.YaxisMinimum = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.YaxisMax = New System.Windows.Forms.Label()
        Me.YaxisMaximum = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ButtonZoomOut = New System.Windows.Forms.Button()
        Me.ButtonZoomIn = New System.Windows.Forms.Button()
        Me.ButtonYmaxInc = New System.Windows.Forms.Button()
        Me.ButtonYmaxDec = New System.Windows.Forms.Button()
        Me.ButtonYminDec = New System.Windows.Forms.Button()
        Me.ButtonYminInc = New System.Windows.Forms.Button()
        Me.BrowseToFile = New System.Windows.Forms.Button()
        Me.ButtonDisplayAll = New System.Windows.Forms.Button()
        Me.DeviceName1 = New System.Windows.Forms.TextBox()
        Me.DeviceName2 = New System.Windows.Forms.TextBox()
        Me.PlaybackTemp = New System.Windows.Forms.CheckBox()
        Me.PlaybackHum = New System.Windows.Forms.CheckBox()
        Me.ChartScaleMax = New System.Windows.Forms.TextBox()
        Me.ChartScaleMin = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.ButtonSaveSettings = New System.Windows.Forms.Button()
        Me.LabelHum = New System.Windows.Forms.Label()
        Me.LabelTempC = New System.Windows.Forms.Label()
        Me.ButtonShiftUp = New System.Windows.Forms.Button()
        Me.ButtonShiftDn = New System.Windows.Forms.Button()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.CheckBoxToolTips = New System.Windows.Forms.CheckBox()
        Me.CheckBoxMaxMin = New System.Windows.Forms.CheckBox()
        Me.Scale1 = New System.Windows.Forms.Label()
        Me.Scale2 = New System.Windows.Forms.Label()
        Me.Scale3 = New System.Windows.Forms.Label()
        Me.Scale4 = New System.Windows.Forms.Label()
        Me.Scale5 = New System.Windows.Forms.Label()
        Me.Scale6 = New System.Windows.Forms.Label()
        Me.Scale7 = New System.Windows.Forms.Label()
        Me.Scale8 = New System.Windows.Forms.Label()
        Me.Scale9 = New System.Windows.Forms.Label()
        Me.Scale10 = New System.Windows.Forms.Label()
        Me.Scale11 = New System.Windows.Forms.Label()
        Me.Scale12 = New System.Windows.Forms.Label()
        Me.Scale13 = New System.Windows.Forms.Label()
        Me.Scale14 = New System.Windows.Forms.Label()
        Me.Scale15 = New System.Windows.Forms.Label()
        Me.Scale16 = New System.Windows.Forms.Label()
        Me.Scale25 = New System.Windows.Forms.Label()
        Me.Scale17 = New System.Windows.Forms.Label()
        Me.Scale24 = New System.Windows.Forms.Label()
        Me.Scale18 = New System.Windows.Forms.Label()
        Me.Scale23 = New System.Windows.Forms.Label()
        Me.Scale19 = New System.Windows.Forms.Label()
        Me.Scale20 = New System.Windows.Forms.Label()
        Me.Scale21 = New System.Windows.Forms.Label()
        Me.Scale22 = New System.Windows.Forms.Label()
        Me.MedianValue = New System.Windows.Forms.TextBox()
        Me.MedianValueText = New System.Windows.Forms.Label()
        Me.CheckBoxPPMenable = New System.Windows.Forms.CheckBox()
        Me.RadioButtonDev1 = New System.Windows.Forms.RadioButton()
        Me.RadioButtonDev2 = New System.Windows.Forms.RadioButton()
        Me.PPMBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxMedianT = New System.Windows.Forms.CheckBox()
        Me.CheckBoxMedianV = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.RadioButtonPPMTempo = New System.Windows.Forms.RadioButton()
        Me.RadioButtonPPMDev = New System.Windows.Forms.RadioButton()
        Me.PPMscaleText = New System.Windows.Forms.Label()
        Me.PPMscalerangeentry = New System.Windows.Forms.TextBox()
        Me.MedianTempText = New System.Windows.Forms.Label()
        Me.MedianTemp = New System.Windows.Forms.TextBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.LabelPPMtop = New System.Windows.Forms.Label()
        Me.GroupBoxMisc = New System.Windows.Forms.GroupBox()
        Me.CheckBoxColours = New System.Windows.Forms.CheckBox()
        Me.CheckBoxGraphOnly = New System.Windows.Forms.CheckBox()
        Me.Screenshot = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Xscale = New System.Windows.Forms.Label()
        Me.MinsTotal = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ButtonScrollLeftSMALL = New System.Windows.Forms.Button()
        Me.ButtonScrollRightSMALL = New System.Windows.Forms.Button()
        Me.CheckBoxYscaletidy = New System.Windows.Forms.CheckBox()
        Me.RefreshFile = New System.Windows.Forms.Button()
        Me.ShowFiles2 = New System.Windows.Forms.Button()
        Me.YaxisSave = New System.Windows.Forms.Button()
        Me.YaxisLoad = New System.Windows.Forms.Button()
        Me.CheckX1000 = New System.Windows.Forms.CheckBox()
        Me.CheckX1000000 = New System.Windows.Forms.CheckBox()
        Me.DEV2avg = New System.Windows.Forms.TextBox()
        Me.DEV1avg = New System.Windows.Forms.TextBox()
        Me.TEMPavg = New System.Windows.Forms.TextBox()
        Me.Dev2MaxMin = New System.Windows.Forms.TextBox()
        Me.RMSaverageDev2 = New System.Windows.Forms.TextBox()
        Me.RMSaverageDev1 = New System.Windows.Forms.TextBox()
        Me.Dev1MaxMin = New System.Windows.Forms.TextBox()
        Me.RMSwindow = New System.Windows.Forms.TextBox()
        Me.LogYaxis = New System.Windows.Forms.CheckBox()
        Me.Xscaletotal = New System.Windows.Forms.Label()
        Me.Loading = New System.Windows.Forms.Label()
        Me.CheckDev1Point = New System.Windows.Forms.CheckBox()
        Me.CheckDev1Line = New System.Windows.Forms.CheckBox()
        Me.CheckDev2Line = New System.Windows.Forms.CheckBox()
        Me.CheckDev2Point = New System.Windows.Forms.CheckBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.YaxisCheck1 = New System.Windows.Forms.CheckBox()
        Me.YaxisCheck2 = New System.Windows.Forms.CheckBox()
        Me.YaxisCheck3 = New System.Windows.Forms.CheckBox()
        Me.YaxisCheck4 = New System.Windows.Forms.CheckBox()
        Me.YaxisBox1 = New System.Windows.Forms.GroupBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.YaxisPerDiv = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.SampleRateSecs = New System.Windows.Forms.TextBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.LabelPPMdegctop = New System.Windows.Forms.Label()
        Me.GroupBoxMiscTempHum = New System.Windows.Forms.GroupBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.PleaseLoadCSV = New System.Windows.Forms.Label()
        Me.MetadataChart = New System.Windows.Forms.TextBox()
        Me.ScaleX28 = New System.Windows.Forms.Label()
        Me.ScaleX1 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        CType(Me.Chart2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PPMBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBoxMisc.SuspendLayout()
        Me.YaxisBox1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBoxMiscTempHum.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Chart2
        '
        Me.Chart2.BackColor = System.Drawing.SystemColors.Control
        ChartArea1.BackColor = System.Drawing.Color.Black
        ChartArea1.BorderColor = System.Drawing.Color.White
        ChartArea1.BorderWidth = 2
        ChartArea1.Name = "ChartArea1"
        Me.Chart2.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.Chart2.Legends.Add(Legend1)
        Me.Chart2.Location = New System.Drawing.Point(1, 191)
        Me.Chart2.Name = "Chart2"
        Series1.ChartArea = "ChartArea1"
        Series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        Series1.Color = System.Drawing.Color.Yellow
        Series1.Enabled = False
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Me.Chart2.Series.Add(Series1)
        Me.Chart2.Size = New System.Drawing.Size(1334, 610)
        Me.Chart2.TabIndex = 52
        Me.Chart2.Text = "Chart2"
        '
        'ButtonScrollLeft
        '
        Me.ButtonScrollLeft.Location = New System.Drawing.Point(273, 112)
        Me.ButtonScrollLeft.Name = "ButtonScrollLeft"
        Me.ButtonScrollLeft.Size = New System.Drawing.Size(73, 30)
        Me.ButtonScrollLeft.TabIndex = 54
        Me.ButtonScrollLeft.Text = "<< LEFT"
        Me.ToolTip1.SetToolTip(Me.ButtonScrollLeft, "Scroll left through chart")
        Me.ButtonScrollLeft.UseVisualStyleBackColor = True
        '
        'ButtonScrollRight
        '
        Me.ButtonScrollRight.Location = New System.Drawing.Point(421, 112)
        Me.ButtonScrollRight.Name = "ButtonScrollRight"
        Me.ButtonScrollRight.Size = New System.Drawing.Size(73, 30)
        Me.ButtonScrollRight.TabIndex = 55
        Me.ButtonScrollRight.Text = "RIGHT >>"
        Me.ToolTip1.SetToolTip(Me.ButtonScrollRight, "Scroll right through chart")
        Me.ButtonScrollRight.UseVisualStyleBackColor = True
        '
        'CurrentPosition
        '
        Me.CurrentPosition.Location = New System.Drawing.Point(6, 63)
        Me.CurrentPosition.Name = "CurrentPosition"
        Me.CurrentPosition.ReadOnly = True
        Me.CurrentPosition.Size = New System.Drawing.Size(46, 20)
        Me.CurrentPosition.TabIndex = 56
        Me.CurrentPosition.Text = "0"
        '
        'TargetPosition
        '
        Me.TargetPosition.Location = New System.Drawing.Point(6, 82)
        Me.TargetPosition.Name = "TargetPosition"
        Me.TargetPosition.ReadOnly = True
        Me.TargetPosition.Size = New System.Drawing.Size(46, 20)
        Me.TargetPosition.TabIndex = 57
        Me.TargetPosition.Text = "0"
        '
        'RangeRequired
        '
        Me.RangeRequired.Location = New System.Drawing.Point(642, 29)
        Me.RangeRequired.Name = "RangeRequired"
        Me.RangeRequired.Size = New System.Drawing.Size(46, 20)
        Me.RangeRequired.TabIndex = 58
        Me.RangeRequired.Text = "50"
        Me.RangeRequired.Visible = False
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.Location = New System.Drawing.Point(688, 32)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(101, 13)
        Me.Label40.TabIndex = 61
        Me.Label40.Text = "- Visible Data Points"
        Me.Label40.Visible = False
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(18, 65)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(111, 13)
        Me.Label27.TabIndex = 65
        Me.Label27.Text = "CSV File Path/Name -"
        '
        'CSVfilenamePlayback
        '
        Me.CSVfilenamePlayback.Location = New System.Drawing.Point(130, 62)
        Me.CSVfilenamePlayback.Name = "CSVfilenamePlayback"
        Me.CSVfilenamePlayback.Size = New System.Drawing.Size(446, 20)
        Me.CSVfilenamePlayback.TabIndex = 64
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(50, 67)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 13)
        Me.Label1.TabIndex = 66
        Me.Label1.Text = "- Start Data"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(50, 86)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 67
        Me.Label2.Text = "- End Data"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(440, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 13)
        Me.Label3.TabIndex = 69
        Me.Label3.Text = "- Points (Dev)"
        '
        'CSVfileLines
        '
        Me.CSVfileLines.Location = New System.Drawing.Point(405, 23)
        Me.CSVfileLines.Name = "CSVfileLines"
        Me.CSVfileLines.ReadOnly = True
        Me.CSVfileLines.Size = New System.Drawing.Size(35, 20)
        Me.CSVfileLines.TabIndex = 70
        Me.CSVfileLines.Text = "0"
        '
        'YaxisMinimum
        '
        Me.YaxisMinimum.Location = New System.Drawing.Point(13, 146)
        Me.YaxisMinimum.Name = "YaxisMinimum"
        Me.YaxisMinimum.Size = New System.Drawing.Size(83, 20)
        Me.YaxisMinimum.TabIndex = 71
        Me.YaxisMinimum.Text = "0.000000"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(97, 130)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(36, 13)
        Me.Label4.TabIndex = 72
        Me.Label4.Text = "- Max."
        '
        'YaxisMax
        '
        Me.YaxisMax.AutoSize = True
        Me.YaxisMax.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.YaxisMax.Location = New System.Drawing.Point(97, 149)
        Me.YaxisMax.Name = "YaxisMax"
        Me.YaxisMax.Size = New System.Drawing.Size(33, 13)
        Me.YaxisMax.TabIndex = 74
        Me.YaxisMax.Text = "- Min."
        '
        'YaxisMaximum
        '
        Me.YaxisMaximum.Location = New System.Drawing.Point(13, 127)
        Me.YaxisMaximum.Name = "YaxisMaximum"
        Me.YaxisMaximum.Size = New System.Drawing.Size(83, 20)
        Me.YaxisMaximum.TabIndex = 73
        Me.YaxisMaximum.Text = "15.000000"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(584, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(216, 25)
        Me.Label5.TabIndex = 75
        Me.Label5.Text = "PLAYBACK CHART"
        '
        'ButtonZoomOut
        '
        Me.ButtonZoomOut.Location = New System.Drawing.Point(347, 153)
        Me.ButtonZoomOut.Name = "ButtonZoomOut"
        Me.ButtonZoomOut.Size = New System.Drawing.Size(73, 30)
        Me.ButtonZoomOut.TabIndex = 78
        Me.ButtonZoomOut.Text = "ZOOM OUT"
        Me.ToolTip1.SetToolTip(Me.ButtonZoomOut, "Zoom out Y-axis on centre line")
        Me.ButtonZoomOut.UseVisualStyleBackColor = True
        '
        'ButtonZoomIn
        '
        Me.ButtonZoomIn.Location = New System.Drawing.Point(273, 153)
        Me.ButtonZoomIn.Name = "ButtonZoomIn"
        Me.ButtonZoomIn.Size = New System.Drawing.Size(73, 30)
        Me.ButtonZoomIn.TabIndex = 79
        Me.ButtonZoomIn.Text = "ZOOM IN"
        Me.ToolTip1.SetToolTip(Me.ButtonZoomIn, "Zoom in Y-axis on centre line")
        Me.ButtonZoomIn.UseVisualStyleBackColor = True
        '
        'ButtonYmaxInc
        '
        Me.ButtonYmaxInc.Location = New System.Drawing.Point(212, 114)
        Me.ButtonYmaxInc.Name = "ButtonYmaxInc"
        Me.ButtonYmaxInc.Size = New System.Drawing.Size(30, 20)
        Me.ButtonYmaxInc.TabIndex = 80
        Me.ButtonYmaxInc.Text = ">"
        Me.ToolTip1.SetToolTip(Me.ButtonYmaxInc, "Increase Y-axis maximum value")
        Me.ButtonYmaxInc.UseVisualStyleBackColor = True
        '
        'ButtonYmaxDec
        '
        Me.ButtonYmaxDec.Location = New System.Drawing.Point(179, 114)
        Me.ButtonYmaxDec.Name = "ButtonYmaxDec"
        Me.ButtonYmaxDec.Size = New System.Drawing.Size(30, 20)
        Me.ButtonYmaxDec.TabIndex = 81
        Me.ButtonYmaxDec.Text = "<"
        Me.ToolTip1.SetToolTip(Me.ButtonYmaxDec, "Decrease Y-axis maximum value")
        Me.ButtonYmaxDec.UseVisualStyleBackColor = True
        '
        'ButtonYminDec
        '
        Me.ButtonYminDec.Location = New System.Drawing.Point(179, 135)
        Me.ButtonYminDec.Name = "ButtonYminDec"
        Me.ButtonYminDec.Size = New System.Drawing.Size(30, 20)
        Me.ButtonYminDec.TabIndex = 83
        Me.ButtonYminDec.Text = "<"
        Me.ToolTip1.SetToolTip(Me.ButtonYminDec, "Decrease Y-axis minimum value")
        Me.ButtonYminDec.UseVisualStyleBackColor = True
        '
        'ButtonYminInc
        '
        Me.ButtonYminInc.Location = New System.Drawing.Point(212, 135)
        Me.ButtonYminInc.Name = "ButtonYminInc"
        Me.ButtonYminInc.Size = New System.Drawing.Size(30, 20)
        Me.ButtonYminInc.TabIndex = 82
        Me.ButtonYminInc.Text = ">"
        Me.ToolTip1.SetToolTip(Me.ButtonYminInc, "Increase Y-axis minimum value")
        Me.ButtonYminInc.UseVisualStyleBackColor = True
        '
        'BrowseToFile
        '
        Me.BrowseToFile.Location = New System.Drawing.Point(15, 9)
        Me.BrowseToFile.Name = "BrowseToFile"
        Me.BrowseToFile.Size = New System.Drawing.Size(101, 45)
        Me.BrowseToFile.TabIndex = 84
        Me.BrowseToFile.Text = "LOAD .CSV FILE"
        Me.ToolTip1.SetToolTip(Me.BrowseToFile, "Load CSV from disk.")
        Me.BrowseToFile.UseVisualStyleBackColor = True
        '
        'ButtonDisplayAll
        '
        Me.ButtonDisplayAll.Location = New System.Drawing.Point(421, 153)
        Me.ButtonDisplayAll.Name = "ButtonDisplayAll"
        Me.ButtonDisplayAll.Size = New System.Drawing.Size(73, 30)
        Me.ButtonDisplayAll.TabIndex = 85
        Me.ButtonDisplayAll.Text = "ZOOM ALL"
        Me.ToolTip1.SetToolTip(Me.ButtonDisplayAll, "Display all of chart")
        Me.ButtonDisplayAll.UseVisualStyleBackColor = True
        '
        'DeviceName1
        '
        Me.DeviceName1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.DeviceName1.Location = New System.Drawing.Point(665, 127)
        Me.DeviceName1.Name = "DeviceName1"
        Me.DeviceName1.ReadOnly = True
        Me.DeviceName1.Size = New System.Drawing.Size(121, 20)
        Me.DeviceName1.TabIndex = 86
        '
        'DeviceName2
        '
        Me.DeviceName2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.DeviceName2.Location = New System.Drawing.Point(205, 23)
        Me.DeviceName2.Name = "DeviceName2"
        Me.DeviceName2.ReadOnly = True
        Me.DeviceName2.Size = New System.Drawing.Size(123, 20)
        Me.DeviceName2.TabIndex = 91
        '
        'PlaybackTemp
        '
        Me.PlaybackTemp.AutoSize = True
        Me.PlaybackTemp.Location = New System.Drawing.Point(116, 24)
        Me.PlaybackTemp.Name = "PlaybackTemp"
        Me.PlaybackTemp.Size = New System.Drawing.Size(53, 17)
        Me.PlaybackTemp.TabIndex = 92
        Me.PlaybackTemp.Text = "Temp"
        Me.ToolTip1.SetToolTip(Me.PlaybackTemp, "Enable Temperature")
        Me.PlaybackTemp.UseVisualStyleBackColor = True
        '
        'PlaybackHum
        '
        Me.PlaybackHum.AutoSize = True
        Me.PlaybackHum.Location = New System.Drawing.Point(116, 44)
        Me.PlaybackHum.Name = "PlaybackHum"
        Me.PlaybackHum.Size = New System.Drawing.Size(51, 17)
        Me.PlaybackHum.TabIndex = 93
        Me.PlaybackHum.Text = "Hum."
        Me.ToolTip1.SetToolTip(Me.PlaybackHum, "Enable Humidity")
        Me.PlaybackHum.UseVisualStyleBackColor = True
        '
        'ChartScaleMax
        '
        Me.ChartScaleMax.Location = New System.Drawing.Point(6, 44)
        Me.ChartScaleMax.Name = "ChartScaleMax"
        Me.ChartScaleMax.Size = New System.Drawing.Size(26, 20)
        Me.ChartScaleMax.TabIndex = 94
        Me.ChartScaleMax.Text = "50"
        '
        'ChartScaleMin
        '
        Me.ChartScaleMin.Location = New System.Drawing.Point(6, 63)
        Me.ChartScaleMin.Name = "ChartScaleMin"
        Me.ChartScaleMin.Size = New System.Drawing.Size(26, 20)
        Me.ChartScaleMin.TabIndex = 95
        Me.ChartScaleMin.Text = "15"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(32, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(87, 13)
        Me.Label6.TabIndex = 96
        Me.Label6.Text = "Temp/Hum Max."
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(32, 67)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(84, 13)
        Me.Label7.TabIndex = 97
        Me.Label7.Text = "Temp/Hum Min."
        '
        'ButtonSaveSettings
        '
        Me.ButtonSaveSettings.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ButtonSaveSettings.Location = New System.Drawing.Point(448, 35)
        Me.ButtonSaveSettings.Name = "ButtonSaveSettings"
        Me.ButtonSaveSettings.Size = New System.Drawing.Size(129, 22)
        Me.ButtonSaveSettings.TabIndex = 98
        Me.ButtonSaveSettings.Text = "Save Playback Settings"
        Me.ToolTip1.SetToolTip(Me.ButtonSaveSettings, "Save settings for most of the user data on this form")
        Me.ButtonSaveSettings.UseVisualStyleBackColor = True
        '
        'LabelHum
        '
        Me.LabelHum.AutoSize = True
        Me.LabelHum.BackColor = System.Drawing.Color.Black
        Me.LabelHum.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelHum.ForeColor = System.Drawing.Color.DodgerBlue
        Me.LabelHum.Location = New System.Drawing.Point(1265, 190)
        Me.LabelHum.Name = "LabelHum"
        Me.LabelHum.Size = New System.Drawing.Size(36, 15)
        Me.LabelHum.TabIndex = 102
        Me.LabelHum.Text = "%RH"
        '
        'LabelTempC
        '
        Me.LabelTempC.AutoSize = True
        Me.LabelTempC.BackColor = System.Drawing.Color.Black
        Me.LabelTempC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelTempC.ForeColor = System.Drawing.Color.Red
        Me.LabelTempC.Location = New System.Drawing.Point(1229, 190)
        Me.LabelTempC.Name = "LabelTempC"
        Me.LabelTempC.Size = New System.Drawing.Size(38, 15)
        Me.LabelTempC.TabIndex = 101
        Me.LabelTempC.Text = "DegC"
        '
        'ButtonShiftUp
        '
        Me.ButtonShiftUp.Location = New System.Drawing.Point(11, 404)
        Me.ButtonShiftUp.Name = "ButtonShiftUp"
        Me.ButtonShiftUp.Size = New System.Drawing.Size(50, 72)
        Me.ButtonShiftUp.TabIndex = 104
        Me.ButtonShiftUp.Text = "SHIFT" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "UP"
        Me.ToolTip1.SetToolTip(Me.ButtonShiftUp, "Move graph max/min limits up by 20%")
        Me.ButtonShiftUp.UseVisualStyleBackColor = True
        '
        'ButtonShiftDn
        '
        Me.ButtonShiftDn.Location = New System.Drawing.Point(11, 484)
        Me.ButtonShiftDn.Name = "ButtonShiftDn"
        Me.ButtonShiftDn.Size = New System.Drawing.Size(50, 72)
        Me.ButtonShiftDn.TabIndex = 105
        Me.ButtonShiftDn.Text = "SHIFT" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "DOWN"
        Me.ToolTip1.SetToolTip(Me.ButtonShiftDn, "Move graph max/min limits down by 20%")
        Me.ButtonShiftDn.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(501, 785)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(331, 13)
        Me.Label14.TabIndex = 113
        Me.Label14.Text = "Hover mouse of data points on chart to see value (ToolTips enabled)"
        '
        'CheckBoxToolTips
        '
        Me.CheckBoxToolTips.AutoSize = True
        Me.CheckBoxToolTips.Location = New System.Drawing.Point(8, 27)
        Me.CheckBoxToolTips.Name = "CheckBoxToolTips"
        Me.CheckBoxToolTips.Size = New System.Drawing.Size(97, 17)
        Me.CheckBoxToolTips.TabIndex = 114
        Me.CheckBoxToolTips.Text = "ToolTip Values"
        Me.ToolTip1.SetToolTip(Me.CheckBoxToolTips, "Enable tootip values on graphs")
        Me.CheckBoxToolTips.UseVisualStyleBackColor = True
        '
        'CheckBoxMaxMin
        '
        Me.CheckBoxMaxMin.AutoSize = True
        Me.CheckBoxMaxMin.Checked = True
        Me.CheckBoxMaxMin.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxMaxMin.Location = New System.Drawing.Point(156, 90)
        Me.CheckBoxMaxMin.Name = "CheckBoxMaxMin"
        Me.CheckBoxMaxMin.Size = New System.Drawing.Size(93, 17)
        Me.CheckBoxMaxMin.TabIndex = 115
        Me.CheckBoxMaxMin.Text = "Auto Min/Max"
        Me.CheckBoxMaxMin.UseVisualStyleBackColor = True
        '
        'Scale1
        '
        Me.Scale1.AutoSize = True
        Me.Scale1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale1.Location = New System.Drawing.Point(1311, 212)
        Me.Scale1.Name = "Scale1"
        Me.Scale1.Size = New System.Drawing.Size(29, 13)
        Me.Scale1.TabIndex = 116
        Me.Scale1.Text = "#.#"
        '
        'Scale2
        '
        Me.Scale2.AutoSize = True
        Me.Scale2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale2.Location = New System.Drawing.Point(1311, 234)
        Me.Scale2.Name = "Scale2"
        Me.Scale2.Size = New System.Drawing.Size(29, 13)
        Me.Scale2.TabIndex = 118
        Me.Scale2.Text = "#.#"
        '
        'Scale3
        '
        Me.Scale3.AutoSize = True
        Me.Scale3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale3.Location = New System.Drawing.Point(1311, 256)
        Me.Scale3.Name = "Scale3"
        Me.Scale3.Size = New System.Drawing.Size(29, 13)
        Me.Scale3.TabIndex = 119
        Me.Scale3.Text = "#.#"
        '
        'Scale4
        '
        Me.Scale4.AutoSize = True
        Me.Scale4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale4.Location = New System.Drawing.Point(1311, 279)
        Me.Scale4.Name = "Scale4"
        Me.Scale4.Size = New System.Drawing.Size(29, 13)
        Me.Scale4.TabIndex = 120
        Me.Scale4.Text = "#.#"
        '
        'Scale5
        '
        Me.Scale5.AutoSize = True
        Me.Scale5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale5.Location = New System.Drawing.Point(1311, 301)
        Me.Scale5.Name = "Scale5"
        Me.Scale5.Size = New System.Drawing.Size(29, 13)
        Me.Scale5.TabIndex = 121
        Me.Scale5.Text = "#.#"
        '
        'Scale6
        '
        Me.Scale6.AutoSize = True
        Me.Scale6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale6.Location = New System.Drawing.Point(1311, 324)
        Me.Scale6.Name = "Scale6"
        Me.Scale6.Size = New System.Drawing.Size(29, 13)
        Me.Scale6.TabIndex = 122
        Me.Scale6.Text = "#.#"
        '
        'Scale7
        '
        Me.Scale7.AutoSize = True
        Me.Scale7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale7.Location = New System.Drawing.Point(1311, 345)
        Me.Scale7.Name = "Scale7"
        Me.Scale7.Size = New System.Drawing.Size(29, 13)
        Me.Scale7.TabIndex = 123
        Me.Scale7.Text = "#.#"
        '
        'Scale8
        '
        Me.Scale8.AutoSize = True
        Me.Scale8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale8.Location = New System.Drawing.Point(1311, 368)
        Me.Scale8.Name = "Scale8"
        Me.Scale8.Size = New System.Drawing.Size(29, 13)
        Me.Scale8.TabIndex = 124
        Me.Scale8.Text = "#.#"
        '
        'Scale9
        '
        Me.Scale9.AutoSize = True
        Me.Scale9.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale9.Location = New System.Drawing.Point(1311, 390)
        Me.Scale9.Name = "Scale9"
        Me.Scale9.Size = New System.Drawing.Size(29, 13)
        Me.Scale9.TabIndex = 125
        Me.Scale9.Text = "#.#"
        '
        'Scale10
        '
        Me.Scale10.AutoSize = True
        Me.Scale10.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale10.Location = New System.Drawing.Point(1311, 412)
        Me.Scale10.Name = "Scale10"
        Me.Scale10.Size = New System.Drawing.Size(29, 13)
        Me.Scale10.TabIndex = 126
        Me.Scale10.Text = "#.#"
        '
        'Scale11
        '
        Me.Scale11.AutoSize = True
        Me.Scale11.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale11.Location = New System.Drawing.Point(1311, 434)
        Me.Scale11.Name = "Scale11"
        Me.Scale11.Size = New System.Drawing.Size(29, 13)
        Me.Scale11.TabIndex = 127
        Me.Scale11.Text = "#.#"
        '
        'Scale12
        '
        Me.Scale12.AutoSize = True
        Me.Scale12.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale12.Location = New System.Drawing.Point(1311, 457)
        Me.Scale12.Name = "Scale12"
        Me.Scale12.Size = New System.Drawing.Size(29, 13)
        Me.Scale12.TabIndex = 129
        Me.Scale12.Text = "#.#"
        '
        'Scale13
        '
        Me.Scale13.AutoSize = True
        Me.Scale13.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale13.Location = New System.Drawing.Point(1311, 479)
        Me.Scale13.Name = "Scale13"
        Me.Scale13.Size = New System.Drawing.Size(29, 13)
        Me.Scale13.TabIndex = 130
        Me.Scale13.Text = "#.#"
        '
        'Scale14
        '
        Me.Scale14.AutoSize = True
        Me.Scale14.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale14.Location = New System.Drawing.Point(1311, 502)
        Me.Scale14.Name = "Scale14"
        Me.Scale14.Size = New System.Drawing.Size(29, 13)
        Me.Scale14.TabIndex = 131
        Me.Scale14.Text = "#.#"
        '
        'Scale15
        '
        Me.Scale15.AutoSize = True
        Me.Scale15.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale15.Location = New System.Drawing.Point(1311, 524)
        Me.Scale15.Name = "Scale15"
        Me.Scale15.Size = New System.Drawing.Size(29, 13)
        Me.Scale15.TabIndex = 132
        Me.Scale15.Text = "#.#"
        '
        'Scale16
        '
        Me.Scale16.AutoSize = True
        Me.Scale16.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale16.Location = New System.Drawing.Point(1311, 547)
        Me.Scale16.Name = "Scale16"
        Me.Scale16.Size = New System.Drawing.Size(29, 13)
        Me.Scale16.TabIndex = 134
        Me.Scale16.Text = "#.#"
        '
        'Scale25
        '
        Me.Scale25.AutoSize = True
        Me.Scale25.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale25.Location = New System.Drawing.Point(1311, 744)
        Me.Scale25.Name = "Scale25"
        Me.Scale25.Size = New System.Drawing.Size(29, 13)
        Me.Scale25.TabIndex = 135
        Me.Scale25.Text = "#.#"
        '
        'Scale17
        '
        Me.Scale17.AutoSize = True
        Me.Scale17.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale17.Location = New System.Drawing.Point(1311, 569)
        Me.Scale17.Name = "Scale17"
        Me.Scale17.Size = New System.Drawing.Size(29, 13)
        Me.Scale17.TabIndex = 136
        Me.Scale17.Text = "#.#"
        '
        'Scale24
        '
        Me.Scale24.AutoSize = True
        Me.Scale24.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale24.Location = New System.Drawing.Point(1311, 723)
        Me.Scale24.Name = "Scale24"
        Me.Scale24.Size = New System.Drawing.Size(29, 13)
        Me.Scale24.TabIndex = 137
        Me.Scale24.Text = "#.#"
        '
        'Scale18
        '
        Me.Scale18.AutoSize = True
        Me.Scale18.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale18.Location = New System.Drawing.Point(1311, 591)
        Me.Scale18.Name = "Scale18"
        Me.Scale18.Size = New System.Drawing.Size(29, 13)
        Me.Scale18.TabIndex = 138
        Me.Scale18.Text = "#.#"
        '
        'Scale23
        '
        Me.Scale23.AutoSize = True
        Me.Scale23.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale23.Location = New System.Drawing.Point(1311, 700)
        Me.Scale23.Name = "Scale23"
        Me.Scale23.Size = New System.Drawing.Size(29, 13)
        Me.Scale23.TabIndex = 139
        Me.Scale23.Text = "#.#"
        '
        'Scale19
        '
        Me.Scale19.AutoSize = True
        Me.Scale19.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale19.Location = New System.Drawing.Point(1311, 613)
        Me.Scale19.Name = "Scale19"
        Me.Scale19.Size = New System.Drawing.Size(29, 13)
        Me.Scale19.TabIndex = 140
        Me.Scale19.Text = "#.#"
        '
        'Scale20
        '
        Me.Scale20.AutoSize = True
        Me.Scale20.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale20.Location = New System.Drawing.Point(1311, 635)
        Me.Scale20.Name = "Scale20"
        Me.Scale20.Size = New System.Drawing.Size(29, 13)
        Me.Scale20.TabIndex = 141
        Me.Scale20.Text = "#.#"
        '
        'Scale21
        '
        Me.Scale21.AutoSize = True
        Me.Scale21.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale21.Location = New System.Drawing.Point(1311, 657)
        Me.Scale21.Name = "Scale21"
        Me.Scale21.Size = New System.Drawing.Size(29, 13)
        Me.Scale21.TabIndex = 142
        Me.Scale21.Text = "#.#"
        '
        'Scale22
        '
        Me.Scale22.AutoSize = True
        Me.Scale22.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Scale22.Location = New System.Drawing.Point(1311, 678)
        Me.Scale22.Name = "Scale22"
        Me.Scale22.Size = New System.Drawing.Size(29, 13)
        Me.Scale22.TabIndex = 143
        Me.Scale22.Text = "#.#"
        '
        'MedianValue
        '
        Me.MedianValue.Location = New System.Drawing.Point(817, 59)
        Me.MedianValue.Name = "MedianValue"
        Me.MedianValue.Size = New System.Drawing.Size(75, 20)
        Me.MedianValue.TabIndex = 144
        Me.MedianValue.Text = "10.0000000"
        '
        'MedianValueText
        '
        Me.MedianValueText.AutoSize = True
        Me.MedianValueText.Location = New System.Drawing.Point(892, 62)
        Me.MedianValueText.Name = "MedianValueText"
        Me.MedianValueText.Size = New System.Drawing.Size(67, 13)
        Me.MedianValueText.TabIndex = 145
        Me.MedianValueText.Text = "- Initial Value"
        '
        'CheckBoxPPMenable
        '
        Me.CheckBoxPPMenable.AutoSize = True
        Me.CheckBoxPPMenable.Location = New System.Drawing.Point(817, 32)
        Me.CheckBoxPPMenable.Name = "CheckBoxPPMenable"
        Me.CheckBoxPPMenable.Size = New System.Drawing.Size(85, 17)
        Me.CheckBoxPPMenable.TabIndex = 146
        Me.CheckBoxPPMenable.Text = "Enable PPM"
        Me.ToolTip1.SetToolTip(Me.CheckBoxPPMenable, "Enable PPM")
        Me.CheckBoxPPMenable.UseVisualStyleBackColor = True
        '
        'RadioButtonDev1
        '
        Me.RadioButtonDev1.AutoSize = True
        Me.RadioButtonDev1.Location = New System.Drawing.Point(110, 27)
        Me.RadioButtonDev1.Name = "RadioButtonDev1"
        Me.RadioButtonDev1.Size = New System.Drawing.Size(59, 17)
        Me.RadioButtonDev1.TabIndex = 147
        Me.RadioButtonDev1.TabStop = True
        Me.RadioButtonDev1.Text = "Device"
        Me.RadioButtonDev1.UseVisualStyleBackColor = True
        '
        'RadioButtonDev2
        '
        Me.RadioButtonDev2.AutoSize = True
        Me.RadioButtonDev2.Location = New System.Drawing.Point(184, 27)
        Me.RadioButtonDev2.Name = "RadioButtonDev2"
        Me.RadioButtonDev2.Size = New System.Drawing.Size(59, 17)
        Me.RadioButtonDev2.TabIndex = 148
        Me.RadioButtonDev2.TabStop = True
        Me.RadioButtonDev2.Text = "Device"
        Me.RadioButtonDev2.UseVisualStyleBackColor = True
        '
        'PPMBox1
        '
        Me.PPMBox1.Controls.Add(Me.CheckBoxMedianT)
        Me.PPMBox1.Controls.Add(Me.CheckBoxMedianV)
        Me.PPMBox1.Controls.Add(Me.Panel1)
        Me.PPMBox1.Controls.Add(Me.PPMscaleText)
        Me.PPMBox1.Controls.Add(Me.PPMscalerangeentry)
        Me.PPMBox1.Controls.Add(Me.MedianTempText)
        Me.PPMBox1.Controls.Add(Me.RadioButtonDev2)
        Me.PPMBox1.Controls.Add(Me.MedianTemp)
        Me.PPMBox1.Controls.Add(Me.RadioButtonDev1)
        Me.PPMBox1.Controls.Add(Me.Label38)
        Me.PPMBox1.Enabled = False
        Me.PPMBox1.Location = New System.Drawing.Point(807, 4)
        Me.PPMBox1.Name = "PPMBox1"
        Me.PPMBox1.Size = New System.Drawing.Size(371, 100)
        Me.PPMBox1.TabIndex = 149
        Me.PPMBox1.TabStop = False
        '
        'CheckBoxMedianT
        '
        Me.CheckBoxMedianT.AutoSize = True
        Me.CheckBoxMedianT.Location = New System.Drawing.Point(167, 77)
        Me.CheckBoxMedianT.Name = "CheckBoxMedianT"
        Me.CheckBoxMedianT.Size = New System.Drawing.Size(79, 17)
        Me.CheckBoxMedianT.TabIndex = 196
        Me.CheckBoxMedianT.Text = "- From CSV"
        Me.CheckBoxMedianT.UseVisualStyleBackColor = True
        '
        'CheckBoxMedianV
        '
        Me.CheckBoxMedianV.AutoSize = True
        Me.CheckBoxMedianV.Location = New System.Drawing.Point(167, 58)
        Me.CheckBoxMedianV.Name = "CheckBoxMedianV"
        Me.CheckBoxMedianV.Size = New System.Drawing.Size(79, 17)
        Me.CheckBoxMedianV.TabIndex = 195
        Me.CheckBoxMedianV.Text = "- From CSV"
        Me.CheckBoxMedianV.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.RadioButtonPPMTempo)
        Me.Panel1.Controls.Add(Me.RadioButtonPPMDev)
        Me.Panel1.Location = New System.Drawing.Point(258, 22)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(102, 45)
        Me.Panel1.TabIndex = 157
        '
        'RadioButtonPPMTempo
        '
        Me.RadioButtonPPMTempo.AutoSize = True
        Me.RadioButtonPPMTempo.Location = New System.Drawing.Point(3, 25)
        Me.RadioButtonPPMTempo.Name = "RadioButtonPPMTempo"
        Me.RadioButtonPPMTempo.Size = New System.Drawing.Size(80, 17)
        Me.RadioButtonPPMTempo.TabIndex = 156
        Me.RadioButtonPPMTempo.TabStop = True
        Me.RadioButtonPPMTempo.Text = "PPM/DegC"
        Me.RadioButtonPPMTempo.UseVisualStyleBackColor = True
        '
        'RadioButtonPPMDev
        '
        Me.RadioButtonPPMDev.AutoSize = True
        Me.RadioButtonPPMDev.Location = New System.Drawing.Point(3, 5)
        Me.RadioButtonPPMDev.Name = "RadioButtonPPMDev"
        Me.RadioButtonPPMDev.Size = New System.Drawing.Size(96, 17)
        Me.RadioButtonPPMDev.TabIndex = 155
        Me.RadioButtonPPMDev.TabStop = True
        Me.RadioButtonPPMDev.Text = "PPM Deviation"
        Me.RadioButtonPPMDev.UseVisualStyleBackColor = True
        '
        'PPMscaleText
        '
        Me.PPMscaleText.AutoSize = True
        Me.PPMscaleText.Location = New System.Drawing.Point(293, 76)
        Me.PPMscaleText.Name = "PPMscaleText"
        Me.PPMscaleText.Size = New System.Drawing.Size(66, 13)
        Me.PPMscaleText.TabIndex = 154
        Me.PPMscaleText.Text = "- PPM Scale"
        '
        'PPMscalerangeentry
        '
        Me.PPMscalerangeentry.Location = New System.Drawing.Point(261, 73)
        Me.PPMscalerangeentry.Name = "PPMscalerangeentry"
        Me.PPMscalerangeentry.Size = New System.Drawing.Size(26, 20)
        Me.PPMscalerangeentry.TabIndex = 151
        Me.PPMscalerangeentry.Text = "6"
        '
        'MedianTempText
        '
        Me.MedianTempText.AutoSize = True
        Me.MedianTempText.Location = New System.Drawing.Point(85, 77)
        Me.MedianTempText.Name = "MedianTempText"
        Me.MedianTempText.Size = New System.Drawing.Size(67, 13)
        Me.MedianTempText.TabIndex = 150
        Me.MedianTempText.Text = "- Initial Temp"
        '
        'MedianTemp
        '
        Me.MedianTemp.Location = New System.Drawing.Point(10, 74)
        Me.MedianTemp.Name = "MedianTemp"
        Me.MedianTemp.Size = New System.Drawing.Size(75, 20)
        Me.MedianTemp.TabIndex = 150
        Me.MedianTemp.Text = "23.4"
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.Location = New System.Drawing.Point(3, 8)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(173, 13)
        Me.Label38.TabIndex = 48
        Me.Label38.Text = "PPM. DEVIATION / TEMPCO"
        '
        'LabelPPMtop
        '
        Me.LabelPPMtop.AutoSize = True
        Me.LabelPPMtop.BackColor = System.Drawing.Color.Black
        Me.LabelPPMtop.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelPPMtop.ForeColor = System.Drawing.Color.White
        Me.LabelPPMtop.Location = New System.Drawing.Point(1312, 190)
        Me.LabelPPMtop.Name = "LabelPPMtop"
        Me.LabelPPMtop.Size = New System.Drawing.Size(34, 15)
        Me.LabelPPMtop.TabIndex = 152
        Me.LabelPPMtop.Text = "PPM"
        '
        'GroupBoxMisc
        '
        Me.GroupBoxMisc.Controls.Add(Me.CheckBoxColours)
        Me.GroupBoxMisc.Controls.Add(Me.CheckBoxGraphOnly)
        Me.GroupBoxMisc.Controls.Add(Me.Screenshot)
        Me.GroupBoxMisc.Controls.Add(Me.Label13)
        Me.GroupBoxMisc.Controls.Add(Me.CheckBoxToolTips)
        Me.GroupBoxMisc.Enabled = False
        Me.GroupBoxMisc.Location = New System.Drawing.Point(1183, 4)
        Me.GroupBoxMisc.Name = "GroupBoxMisc"
        Me.GroupBoxMisc.Size = New System.Drawing.Size(170, 100)
        Me.GroupBoxMisc.TabIndex = 155
        Me.GroupBoxMisc.TabStop = False
        '
        'CheckBoxColours
        '
        Me.CheckBoxColours.AutoSize = True
        Me.CheckBoxColours.Location = New System.Drawing.Point(8, 48)
        Me.CheckBoxColours.Name = "CheckBoxColours"
        Me.CheckBoxColours.Size = New System.Drawing.Size(79, 17)
        Me.CheckBoxColours.TabIndex = 117
        Me.CheckBoxColours.Text = "Light Mode"
        Me.ToolTip1.SetToolTip(Me.CheckBoxColours, "Set form to light mode, better for printing.")
        Me.CheckBoxColours.UseVisualStyleBackColor = True
        '
        'CheckBoxGraphOnly
        '
        Me.CheckBoxGraphOnly.AutoSize = True
        Me.CheckBoxGraphOnly.Location = New System.Drawing.Point(89, 76)
        Me.CheckBoxGraphOnly.Name = "CheckBoxGraphOnly"
        Me.CheckBoxGraphOnly.Size = New System.Drawing.Size(79, 17)
        Me.CheckBoxGraphOnly.TabIndex = 116
        Me.CheckBoxGraphOnly.Text = "Graph Only"
        Me.ToolTip1.SetToolTip(Me.CheckBoxGraphOnly, "Screenshot the graph area only.")
        Me.CheckBoxGraphOnly.UseVisualStyleBackColor = True
        '
        'Screenshot
        '
        Me.Screenshot.Location = New System.Drawing.Point(7, 71)
        Me.Screenshot.Name = "Screenshot"
        Me.Screenshot.Size = New System.Drawing.Size(78, 24)
        Me.Screenshot.TabIndex = 115
        Me.Screenshot.Text = "Screenshot"
        Me.ToolTip1.SetToolTip(Me.Screenshot, "Save screenshot (PNG) to current folder," & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "uses log file name as filename" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        Me.Screenshot.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(4, 8)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 48
        Me.Label13.Text = "MISC."
        '
        'Xscale
        '
        Me.Xscale.AutoSize = True
        Me.Xscale.BackColor = System.Drawing.Color.Black
        Me.Xscale.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Xscale.ForeColor = System.Drawing.Color.Orange
        Me.Xscale.Location = New System.Drawing.Point(646, 198)
        Me.Xscale.Name = "Xscale"
        Me.Xscale.Size = New System.Drawing.Size(80, 16)
        Me.Xscale.TabIndex = 192
        Me.Xscale.Text = "Time (mins):"
        '
        'MinsTotal
        '
        Me.MinsTotal.Location = New System.Drawing.Point(6, 44)
        Me.MinsTotal.Name = "MinsTotal"
        Me.MinsTotal.ReadOnly = True
        Me.MinsTotal.Size = New System.Drawing.Size(46, 20)
        Me.MinsTotal.TabIndex = 193
        Me.MinsTotal.Text = "0"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(50, 48)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(62, 13)
        Me.Label16.TabIndex = 194
        Me.Label16.Text = "- Mins Total"
        '
        'ButtonScrollLeftSMALL
        '
        Me.ButtonScrollLeftSMALL.Location = New System.Drawing.Point(347, 112)
        Me.ButtonScrollLeftSMALL.Name = "ButtonScrollLeftSMALL"
        Me.ButtonScrollLeftSMALL.Size = New System.Drawing.Size(30, 30)
        Me.ButtonScrollLeftSMALL.TabIndex = 198
        Me.ButtonScrollLeftSMALL.Text = "<"
        Me.ToolTip1.SetToolTip(Me.ButtonScrollLeftSMALL, "Scroll left through chart")
        Me.ButtonScrollLeftSMALL.UseVisualStyleBackColor = True
        '
        'ButtonScrollRightSMALL
        '
        Me.ButtonScrollRightSMALL.Location = New System.Drawing.Point(390, 112)
        Me.ButtonScrollRightSMALL.Name = "ButtonScrollRightSMALL"
        Me.ButtonScrollRightSMALL.Size = New System.Drawing.Size(30, 30)
        Me.ButtonScrollRightSMALL.TabIndex = 199
        Me.ButtonScrollRightSMALL.Text = ">"
        Me.ToolTip1.SetToolTip(Me.ButtonScrollRightSMALL, "Scroll left through chart")
        Me.ButtonScrollRightSMALL.UseVisualStyleBackColor = True
        '
        'CheckBoxYscaletidy
        '
        Me.CheckBoxYscaletidy.AutoSize = True
        Me.CheckBoxYscaletidy.Checked = True
        Me.CheckBoxYscaletidy.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxYscaletidy.Location = New System.Drawing.Point(156, 73)
        Me.CheckBoxYscaletidy.Name = "CheckBoxYscaletidy"
        Me.CheckBoxYscaletidy.Size = New System.Drawing.Size(76, 17)
        Me.CheckBoxYscaletidy.TabIndex = 200
        Me.CheckBoxYscaletidy.Text = "Tidy Scale"
        Me.ToolTip1.SetToolTip(Me.CheckBoxYscaletidy, "Tidy-up Y-Scale annotations")
        Me.CheckBoxYscaletidy.UseVisualStyleBackColor = True
        '
        'RefreshFile
        '
        Me.RefreshFile.Location = New System.Drawing.Point(626, 51)
        Me.RefreshFile.Name = "RefreshFile"
        Me.RefreshFile.Size = New System.Drawing.Size(146, 27)
        Me.RefreshFile.TabIndex = 201
        Me.RefreshFile.Text = "MANUAL REFRESH .CSV"
        Me.ToolTip1.SetToolTip(Me.RefreshFile, "Manually refresh the CSV file currently loaded.")
        Me.RefreshFile.UseVisualStyleBackColor = True
        Me.RefreshFile.Visible = False
        '
        'ShowFiles2
        '
        Me.ShowFiles2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ShowFiles2.Location = New System.Drawing.Point(448, 10)
        Me.ShowFiles2.Name = "ShowFiles2"
        Me.ShowFiles2.Size = New System.Drawing.Size(85, 22)
        Me.ShowFiles2.TabIndex = 558
        Me.ShowFiles2.Text = "\WinGPIBdata"
        Me.ToolTip1.SetToolTip(Me.ShowFiles2, "Launch Windows File Explorer")
        Me.ShowFiles2.UseVisualStyleBackColor = True
        '
        'YaxisSave
        '
        Me.YaxisSave.Location = New System.Drawing.Point(212, 92)
        Me.YaxisSave.Name = "YaxisSave"
        Me.YaxisSave.Size = New System.Drawing.Size(48, 20)
        Me.YaxisSave.TabIndex = 565
        Me.YaxisSave.Text = "SAVE"
        Me.ToolTip1.SetToolTip(Me.YaxisSave, "Save Min/Max to 1, 2, 3 or 4")
        Me.YaxisSave.UseVisualStyleBackColor = True
        '
        'YaxisLoad
        '
        Me.YaxisLoad.Location = New System.Drawing.Point(155, 9)
        Me.YaxisLoad.Name = "YaxisLoad"
        Me.YaxisLoad.Size = New System.Drawing.Size(48, 20)
        Me.YaxisLoad.TabIndex = 567
        Me.YaxisLoad.Text = "LOAD"
        Me.ToolTip1.SetToolTip(Me.YaxisLoad, "Load Y-Axis Min/Max to the graph")
        Me.YaxisLoad.UseVisualStyleBackColor = True
        '
        'CheckX1000
        '
        Me.CheckX1000.AutoSize = True
        Me.CheckX1000.Location = New System.Drawing.Point(406, 67)
        Me.CheckX1000.Name = "CheckX1000"
        Me.CheckX1000.Size = New System.Drawing.Size(43, 17)
        Me.CheckX1000.TabIndex = 578
        Me.CheckX1000.Text = "x1k"
        Me.ToolTip1.SetToolTip(Me.CheckX1000, "I.E. Vdc to mVdc")
        Me.CheckX1000.UseVisualStyleBackColor = True
        '
        'CheckX1000000
        '
        Me.CheckX1000000.AutoSize = True
        Me.CheckX1000000.Location = New System.Drawing.Point(449, 67)
        Me.CheckX1000000.Name = "CheckX1000000"
        Me.CheckX1000000.Size = New System.Drawing.Size(61, 17)
        Me.CheckX1000000.TabIndex = 577
        Me.CheckX1000000.Text = "x1000k"
        Me.ToolTip1.SetToolTip(Me.CheckX1000000, "I.E. Vdc to uVdc")
        Me.CheckX1000000.UseVisualStyleBackColor = True
        '
        'DEV2avg
        '
        Me.DEV2avg.Location = New System.Drawing.Point(332, 23)
        Me.DEV2avg.Name = "DEV2avg"
        Me.DEV2avg.Size = New System.Drawing.Size(26, 20)
        Me.DEV2avg.TabIndex = 579
        Me.DEV2avg.Text = "0"
        Me.ToolTip1.SetToolTip(Me.DEV2avg, "Set to '0' to disable averaging" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Range = 0 to 100.")
        '
        'DEV1avg
        '
        Me.DEV1avg.Location = New System.Drawing.Point(131, 24)
        Me.DEV1avg.Name = "DEV1avg"
        Me.DEV1avg.Size = New System.Drawing.Size(26, 20)
        Me.DEV1avg.TabIndex = 199
        Me.DEV1avg.Text = "0"
        Me.ToolTip1.SetToolTip(Me.DEV1avg, "Set to '0' to disable averaging" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Range = 0 to 100.")
        '
        'TEMPavg
        '
        Me.TEMPavg.Location = New System.Drawing.Point(6, 25)
        Me.TEMPavg.Name = "TEMPavg"
        Me.TEMPavg.Size = New System.Drawing.Size(26, 20)
        Me.TEMPavg.TabIndex = 582
        Me.TEMPavg.Text = "0"
        Me.ToolTip1.SetToolTip(Me.TEMPavg, "Set to '0' to disable averaging" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Range = 0 to 100.")
        '
        'Dev2MaxMin
        '
        Me.Dev2MaxMin.Location = New System.Drawing.Point(205, 42)
        Me.Dev2MaxMin.Name = "Dev2MaxMin"
        Me.Dev2MaxMin.ReadOnly = True
        Me.Dev2MaxMin.Size = New System.Drawing.Size(75, 20)
        Me.Dev2MaxMin.TabIndex = 570
        Me.ToolTip1.SetToolTip(Me.Dev2MaxMin, "Maximum - Minimum for the chart")
        '
        'RMSaverageDev2
        '
        Me.RMSaverageDev2.Location = New System.Drawing.Point(205, 61)
        Me.RMSaverageDev2.Name = "RMSaverageDev2"
        Me.RMSaverageDev2.ReadOnly = True
        Me.RMSaverageDev2.Size = New System.Drawing.Size(75, 20)
        Me.RMSaverageDev2.TabIndex = 583
        Me.ToolTip1.SetToolTip(Me.RMSaverageDev2, "Noise calculation whilst taking into consideration drift over time.")
        '
        'RMSaverageDev1
        '
        Me.RMSaverageDev1.Location = New System.Drawing.Point(6, 62)
        Me.RMSaverageDev1.Name = "RMSaverageDev1"
        Me.RMSaverageDev1.ReadOnly = True
        Me.RMSaverageDev1.Size = New System.Drawing.Size(75, 20)
        Me.RMSaverageDev1.TabIndex = 581
        Me.ToolTip1.SetToolTip(Me.RMSaverageDev1, "Noise calculation whilst taking into consideration drift over time.")
        '
        'Dev1MaxMin
        '
        Me.Dev1MaxMin.Location = New System.Drawing.Point(6, 43)
        Me.Dev1MaxMin.Name = "Dev1MaxMin"
        Me.Dev1MaxMin.ReadOnly = True
        Me.Dev1MaxMin.Size = New System.Drawing.Size(75, 20)
        Me.Dev1MaxMin.TabIndex = 575
        Me.ToolTip1.SetToolTip(Me.Dev1MaxMin, "Maximum - Minimum for the chart")
        '
        'RMSwindow
        '
        Me.RMSwindow.Location = New System.Drawing.Point(405, 42)
        Me.RMSwindow.Name = "RMSwindow"
        Me.RMSwindow.Size = New System.Drawing.Size(35, 20)
        Me.RMSwindow.TabIndex = 586
        Me.ToolTip1.SetToolTip(Me.RMSwindow, resources.GetString("RMSwindow.ToolTip"))
        '
        'LogYaxis
        '
        Me.LogYaxis.AutoSize = True
        Me.LogYaxis.Location = New System.Drawing.Point(685, 87)
        Me.LogYaxis.Name = "LogYaxis"
        Me.LogYaxis.Size = New System.Drawing.Size(75, 17)
        Me.LogYaxis.TabIndex = 195
        Me.LogYaxis.Text = "Log Y-axis"
        Me.LogYaxis.UseVisualStyleBackColor = True
        Me.LogYaxis.Visible = False
        '
        'Xscaletotal
        '
        Me.Xscaletotal.AutoSize = True
        Me.Xscaletotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Xscaletotal.Location = New System.Drawing.Point(729, 198)
        Me.Xscaletotal.Name = "Xscaletotal"
        Me.Xscaletotal.Size = New System.Drawing.Size(31, 16)
        Me.Xscaletotal.TabIndex = 196
        Me.Xscaletotal.Text = "###"
        '
        'Loading
        '
        Me.Loading.AutoSize = True
        Me.Loading.BackColor = System.Drawing.SystemColors.Control
        Me.Loading.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Loading.ForeColor = System.Drawing.Color.Black
        Me.Loading.Location = New System.Drawing.Point(436, 464)
        Me.Loading.Name = "Loading"
        Me.Loading.Size = New System.Drawing.Size(440, 37)
        Me.Loading.TabIndex = 197
        Me.Loading.Text = "Loading CSV, Please Wait....."
        '
        'CheckDev1Point
        '
        Me.CheckDev1Point.AutoSize = True
        Me.CheckDev1Point.Location = New System.Drawing.Point(151, 68)
        Me.CheckDev1Point.Name = "CheckDev1Point"
        Me.CheckDev1Point.Size = New System.Drawing.Size(50, 17)
        Me.CheckDev1Point.TabIndex = 202
        Me.CheckDev1Point.Text = "Point"
        Me.CheckDev1Point.UseVisualStyleBackColor = True
        '
        'CheckDev1Line
        '
        Me.CheckDev1Line.AutoSize = True
        Me.CheckDev1Line.Location = New System.Drawing.Point(151, 49)
        Me.CheckDev1Line.Name = "CheckDev1Line"
        Me.CheckDev1Line.Size = New System.Drawing.Size(46, 17)
        Me.CheckDev1Line.TabIndex = 203
        Me.CheckDev1Line.Text = "Line"
        Me.CheckDev1Line.UseVisualStyleBackColor = True
        '
        'CheckDev2Line
        '
        Me.CheckDev2Line.AutoSize = True
        Me.CheckDev2Line.Location = New System.Drawing.Point(351, 48)
        Me.CheckDev2Line.Name = "CheckDev2Line"
        Me.CheckDev2Line.Size = New System.Drawing.Size(46, 17)
        Me.CheckDev2Line.TabIndex = 206
        Me.CheckDev2Line.Text = "Line"
        Me.CheckDev2Line.UseVisualStyleBackColor = True
        '
        'CheckDev2Point
        '
        Me.CheckDev2Point.AutoSize = True
        Me.CheckDev2Point.Location = New System.Drawing.Point(351, 67)
        Me.CheckDev2Point.Name = "CheckDev2Point"
        Me.CheckDev2Point.Size = New System.Drawing.Size(50, 17)
        Me.CheckDev2Point.TabIndex = 205
        Me.CheckDev2Point.Text = "Point"
        Me.CheckDev2Point.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 10000
        '
        'YaxisCheck1
        '
        Me.YaxisCheck1.AutoSize = True
        Me.YaxisCheck1.Location = New System.Drawing.Point(13, 106)
        Me.YaxisCheck1.Name = "YaxisCheck1"
        Me.YaxisCheck1.Size = New System.Drawing.Size(32, 17)
        Me.YaxisCheck1.TabIndex = 560
        Me.YaxisCheck1.Text = "1"
        Me.YaxisCheck1.UseVisualStyleBackColor = True
        '
        'YaxisCheck2
        '
        Me.YaxisCheck2.AutoSize = True
        Me.YaxisCheck2.Location = New System.Drawing.Point(50, 106)
        Me.YaxisCheck2.Name = "YaxisCheck2"
        Me.YaxisCheck2.Size = New System.Drawing.Size(32, 17)
        Me.YaxisCheck2.TabIndex = 562
        Me.YaxisCheck2.Text = "2"
        Me.YaxisCheck2.UseVisualStyleBackColor = True
        '
        'YaxisCheck3
        '
        Me.YaxisCheck3.AutoSize = True
        Me.YaxisCheck3.Location = New System.Drawing.Point(87, 106)
        Me.YaxisCheck3.Name = "YaxisCheck3"
        Me.YaxisCheck3.Size = New System.Drawing.Size(32, 17)
        Me.YaxisCheck3.TabIndex = 563
        Me.YaxisCheck3.Text = "3"
        Me.YaxisCheck3.UseVisualStyleBackColor = True
        '
        'YaxisCheck4
        '
        Me.YaxisCheck4.AutoSize = True
        Me.YaxisCheck4.Location = New System.Drawing.Point(124, 106)
        Me.YaxisCheck4.Name = "YaxisCheck4"
        Me.YaxisCheck4.Size = New System.Drawing.Size(32, 17)
        Me.YaxisCheck4.TabIndex = 564
        Me.YaxisCheck4.Text = "4"
        Me.YaxisCheck4.UseVisualStyleBackColor = True
        '
        'YaxisBox1
        '
        Me.YaxisBox1.Controls.Add(Me.Label23)
        Me.YaxisBox1.Controls.Add(Me.YaxisLoad)
        Me.YaxisBox1.Controls.Add(Me.YaxisPerDiv)
        Me.YaxisBox1.Controls.Add(Me.Label19)
        Me.YaxisBox1.Controls.Add(Me.CheckBoxMaxMin)
        Me.YaxisBox1.Controls.Add(Me.CheckBoxYscaletidy)
        Me.YaxisBox1.Enabled = False
        Me.YaxisBox1.Location = New System.Drawing.Point(6, 83)
        Me.YaxisBox1.Name = "YaxisBox1"
        Me.YaxisBox1.Size = New System.Drawing.Size(258, 109)
        Me.YaxisBox1.TabIndex = 566
        Me.YaxisBox1.TabStop = False
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.Location = New System.Drawing.Point(91, 85)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(51, 13)
        Me.Label23.TabIndex = 577
        Me.Label23.Text = "- Per Div."
        '
        'YaxisPerDiv
        '
        Me.YaxisPerDiv.Location = New System.Drawing.Point(7, 82)
        Me.YaxisPerDiv.Name = "YaxisPerDiv"
        Me.YaxisPerDiv.ReadOnly = True
        Me.YaxisPerDiv.Size = New System.Drawing.Size(83, 20)
        Me.YaxisPerDiv.TabIndex = 576
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(3, 8)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(90, 13)
        Me.Label19.TabIndex = 116
        Me.Label19.Text = "Y-AXIS SCALE"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Location = New System.Drawing.Point(269, 83)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(229, 109)
        Me.GroupBox1.TabIndex = 567
        Me.GroupBox1.TabStop = False
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(3, 8)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(90, 13)
        Me.Label18.TabIndex = 568
        Me.Label18.Text = "X-AXIS SCALE"
        '
        'SampleRateSecs
        '
        Me.SampleRateSecs.Location = New System.Drawing.Point(6, 25)
        Me.SampleRateSecs.Name = "SampleRateSecs"
        Me.SampleRateSecs.ReadOnly = True
        Me.SampleRateSecs.Size = New System.Drawing.Size(46, 20)
        Me.SampleRateSecs.TabIndex = 569
        Me.SampleRateSecs.Text = "0"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(50, 29)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(101, 13)
        Me.Label25.TabIndex = 570
        Me.Label25.Text = "- Sample Rate Secs"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(281, 48)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(53, 13)
        Me.Label22.TabIndex = 571
        Me.Label22.Text = "- Max-Min"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RMSwindow)
        Me.GroupBox2.Controls.Add(Me.Label21)
        Me.GroupBox2.Controls.Add(Me.CheckX1000000)
        Me.GroupBox2.Controls.Add(Me.RMSaverageDev2)
        Me.GroupBox2.Controls.Add(Me.Label17)
        Me.GroupBox2.Controls.Add(Me.RMSaverageDev1)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.CheckDev2Point)
        Me.GroupBox2.Controls.Add(Me.CheckDev1Point)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.DEV2avg)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label22)
        Me.GroupBox2.Controls.Add(Me.CheckX1000)
        Me.GroupBox2.Controls.Add(Me.DEV1avg)
        Me.GroupBox2.Controls.Add(Me.Dev2MaxMin)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.Dev1MaxMin)
        Me.GroupBox2.Controls.Add(Me.CheckDev2Line)
        Me.GroupBox2.Controls.Add(Me.Label20)
        Me.GroupBox2.Controls.Add(Me.CheckDev1Line)
        Me.GroupBox2.Controls.Add(Me.CSVfileLines)
        Me.GroupBox2.Controls.Add(Me.DeviceName2)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox2.Location = New System.Drawing.Point(659, 103)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(520, 89)
        Me.GroupBox2.TabIndex = 572
        Me.GroupBox2.TabStop = False
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(440, 46)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(76, 13)
        Me.Label21.TabIndex = 585
        Me.Label21.Text = "- RMS window"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(281, 66)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(67, 13)
        Me.Label17.TabIndex = 584
        Me.Label17.Text = "- RMS Noise"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(81, 66)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(67, 13)
        Me.Label10.TabIndex = 582
        Me.Label10.Text = "- RMS Noise"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(359, 27)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(35, 13)
        Me.Label9.TabIndex = 580
        Me.Label9.Text = "- Avg."
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(158, 26)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(35, 13)
        Me.Label8.TabIndex = 200
        Me.Label8.Text = "- Avg."
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(81, 49)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(53, 13)
        Me.Label12.TabIndex = 576
        Me.Label12.Text = "- Max-Min"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(3, 8)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(60, 13)
        Me.Label20.TabIndex = 197
        Me.Label20.Text = "DEVICES"
        '
        'LabelPPMdegctop
        '
        Me.LabelPPMdegctop.AutoSize = True
        Me.LabelPPMdegctop.BackColor = System.Drawing.Color.Black
        Me.LabelPPMdegctop.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelPPMdegctop.ForeColor = System.Drawing.Color.White
        Me.LabelPPMdegctop.Location = New System.Drawing.Point(1312, 202)
        Me.LabelPPMdegctop.Name = "LabelPPMdegctop"
        Me.LabelPPMdegctop.Size = New System.Drawing.Size(34, 12)
        Me.LabelPPMdegctop.TabIndex = 573
        Me.LabelPPMdegctop.Text = " /DegC"
        '
        'GroupBoxMiscTempHum
        '
        Me.GroupBoxMiscTempHum.Controls.Add(Me.Label15)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.Label11)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.TEMPavg)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.PlaybackHum)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.ChartScaleMax)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.ChartScaleMin)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.Label6)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.Label7)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.PlaybackTemp)
        Me.GroupBoxMiscTempHum.Enabled = False
        Me.GroupBoxMiscTempHum.Location = New System.Drawing.Point(1183, 103)
        Me.GroupBoxMiscTempHum.Name = "GroupBoxMiscTempHum"
        Me.GroupBoxMiscTempHum.Size = New System.Drawing.Size(170, 89)
        Me.GroupBoxMiscTempHum.TabIndex = 156
        Me.GroupBoxMiscTempHum.TabStop = False
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(4, 8)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(69, 13)
        Me.Label15.TabIndex = 582
        Me.Label15.Text = "Temp/Hum"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(32, 29)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(59, 13)
        Me.Label11.TabIndex = 582
        Me.Label11.Text = "Temp Avg."
        '
        'PleaseLoadCSV
        '
        Me.PleaseLoadCSV.AutoSize = True
        Me.PleaseLoadCSV.BackColor = System.Drawing.SystemColors.Control
        Me.PleaseLoadCSV.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PleaseLoadCSV.ForeColor = System.Drawing.Color.Black
        Me.PleaseLoadCSV.Location = New System.Drawing.Point(484, 427)
        Me.PleaseLoadCSV.Name = "PleaseLoadCSV"
        Me.PleaseLoadCSV.Size = New System.Drawing.Size(343, 37)
        Me.PleaseLoadCSV.TabIndex = 574
        Me.PleaseLoadCSV.Text = "Please load a CSV file!"
        '
        'MetadataChart
        '
        Me.MetadataChart.Location = New System.Drawing.Point(130, 10)
        Me.MetadataChart.Multiline = True
        Me.MetadataChart.Name = "MetadataChart"
        Me.MetadataChart.ReadOnly = True
        Me.MetadataChart.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.MetadataChart.Size = New System.Drawing.Size(247, 45)
        Me.MetadataChart.TabIndex = 575
        '
        'ScaleX28
        '
        Me.ScaleX28.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ScaleX28.Location = New System.Drawing.Point(192, 193)
        Me.ScaleX28.Name = "ScaleX28"
        Me.ScaleX28.Size = New System.Drawing.Size(38, 13)
        Me.ScaleX28.TabIndex = 187
        Me.ScaleX28.Text = "0"
        Me.ScaleX28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ScaleX28.Visible = False
        '
        'ScaleX1
        '
        Me.ScaleX1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ScaleX1.Location = New System.Drawing.Point(127, 193)
        Me.ScaleX1.Name = "ScaleX1"
        Me.ScaleX1.Size = New System.Drawing.Size(38, 13)
        Me.ScaleX1.TabIndex = 160
        Me.ScaleX1.Text = "0"
        Me.ScaleX1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ScaleX1.Visible = False
        '
        'Label24
        '
        Me.Label24.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.Location = New System.Drawing.Point(328, 193)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(38, 13)
        Me.Label24.TabIndex = 576
        Me.Label24.Text = "#"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label24.Visible = False
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.Location = New System.Drawing.Point(4, 9)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(31, 13)
        Me.Label26.TabIndex = 571
        Me.Label26.Text = "CSV"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.CurrentPosition)
        Me.GroupBox3.Controls.Add(Me.Label26)
        Me.GroupBox3.Controls.Add(Me.MinsTotal)
        Me.GroupBox3.Controls.Add(Me.SampleRateSecs)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.Label25)
        Me.GroupBox3.Controls.Add(Me.Label16)
        Me.GroupBox3.Controls.Add(Me.TargetPosition)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Location = New System.Drawing.Point(502, 83)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(153, 109)
        Me.GroupBox3.TabIndex = 577
        Me.GroupBox3.TabStop = False
        '
        'Chart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1359, 801)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.MetadataChart)
        Me.Controls.Add(Me.PleaseLoadCSV)
        Me.Controls.Add(Me.GroupBoxMiscTempHum)
        Me.Controls.Add(Me.LabelPPMtop)
        Me.Controls.Add(Me.LabelPPMdegctop)
        Me.Controls.Add(Me.RangeRequired)
        Me.Controls.Add(Me.Label40)
        Me.Controls.Add(Me.ButtonYminInc)
        Me.Controls.Add(Me.ButtonYminDec)
        Me.Controls.Add(Me.ButtonYmaxInc)
        Me.Controls.Add(Me.ButtonSaveSettings)
        Me.Controls.Add(Me.ButtonYmaxDec)
        Me.Controls.Add(Me.YaxisSave)
        Me.Controls.Add(Me.YaxisCheck4)
        Me.Controls.Add(Me.YaxisCheck3)
        Me.Controls.Add(Me.YaxisCheck2)
        Me.Controls.Add(Me.YaxisCheck1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.ShowFiles2)
        Me.Controls.Add(Me.RefreshFile)
        Me.Controls.Add(Me.ButtonScrollRightSMALL)
        Me.Controls.Add(Me.ButtonScrollLeftSMALL)
        Me.Controls.Add(Me.Loading)
        Me.Controls.Add(Me.Xscaletotal)
        Me.Controls.Add(Me.LogYaxis)
        Me.Controls.Add(Me.Xscale)
        Me.Controls.Add(Me.ScaleX28)
        Me.Controls.Add(Me.ScaleX1)
        Me.Controls.Add(Me.CheckBoxPPMenable)
        Me.Controls.Add(Me.MedianValueText)
        Me.Controls.Add(Me.MedianValue)
        Me.Controls.Add(Me.Scale22)
        Me.Controls.Add(Me.Scale21)
        Me.Controls.Add(Me.Scale20)
        Me.Controls.Add(Me.Scale19)
        Me.Controls.Add(Me.Scale23)
        Me.Controls.Add(Me.Scale18)
        Me.Controls.Add(Me.Scale24)
        Me.Controls.Add(Me.Scale17)
        Me.Controls.Add(Me.Scale25)
        Me.Controls.Add(Me.Scale16)
        Me.Controls.Add(Me.Scale15)
        Me.Controls.Add(Me.Scale14)
        Me.Controls.Add(Me.Scale13)
        Me.Controls.Add(Me.Scale12)
        Me.Controls.Add(Me.Scale11)
        Me.Controls.Add(Me.Scale10)
        Me.Controls.Add(Me.Scale9)
        Me.Controls.Add(Me.Scale8)
        Me.Controls.Add(Me.Scale7)
        Me.Controls.Add(Me.Scale6)
        Me.Controls.Add(Me.Scale5)
        Me.Controls.Add(Me.Scale4)
        Me.Controls.Add(Me.Scale3)
        Me.Controls.Add(Me.Scale2)
        Me.Controls.Add(Me.Scale1)
        Me.Controls.Add(Me.LabelTempC)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.ButtonShiftDn)
        Me.Controls.Add(Me.ButtonShiftUp)
        Me.Controls.Add(Me.LabelHum)
        Me.Controls.Add(Me.DeviceName1)
        Me.Controls.Add(Me.ButtonDisplayAll)
        Me.Controls.Add(Me.BrowseToFile)
        Me.Controls.Add(Me.ButtonZoomIn)
        Me.Controls.Add(Me.ButtonZoomOut)
        Me.Controls.Add(Me.YaxisMax)
        Me.Controls.Add(Me.YaxisMaximum)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.YaxisMinimum)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.CSVfilenamePlayback)
        Me.Controls.Add(Me.ButtonScrollRight)
        Me.Controls.Add(Me.ButtonScrollLeft)
        Me.Controls.Add(Me.Chart2)
        Me.Controls.Add(Me.PPMBox1)
        Me.Controls.Add(Me.GroupBoxMisc)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.YaxisBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "Chart"
        Me.Text = resources.GetString("$this.Text")
        CType(Me.Chart2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PPMBox1.ResumeLayout(False)
        Me.PPMBox1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBoxMisc.ResumeLayout(False)
        Me.GroupBoxMisc.PerformLayout()
        Me.YaxisBox1.ResumeLayout(False)
        Me.YaxisBox1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBoxMiscTempHum.ResumeLayout(False)
        Me.GroupBoxMiscTempHum.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Chart2 As DataVisualization.Charting.Chart
    Friend WithEvents ButtonScrollLeft As Button
    Friend WithEvents ButtonScrollRight As Button
    Friend WithEvents CurrentPosition As TextBox
    Friend WithEvents TargetPosition As TextBox
    Friend WithEvents RangeRequired As TextBox
    Friend WithEvents Label40 As Label
    Friend WithEvents Label27 As Label
    Friend WithEvents CSVfilenamePlayback As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents CSVfileLines As TextBox
    Friend WithEvents YaxisMinimum As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents YaxisMax As Label
    Friend WithEvents YaxisMaximum As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents ButtonZoomOut As Button
    Friend WithEvents ButtonZoomIn As Button
    Friend WithEvents ButtonYmaxInc As Button
    Friend WithEvents ButtonYmaxDec As Button
    Friend WithEvents ButtonYminDec As Button
    Friend WithEvents ButtonYminInc As Button
    Friend WithEvents BrowseToFile As Button
    Friend WithEvents ButtonDisplayAll As Button
    Friend WithEvents DeviceName1 As TextBox
    Friend WithEvents DeviceName2 As TextBox
    Friend WithEvents PlaybackTemp As CheckBox
    Friend WithEvents PlaybackHum As CheckBox
    Friend WithEvents ChartScaleMax As TextBox
    Friend WithEvents ChartScaleMin As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents ButtonSaveSettings As Button
    'Friend WithEvents RectangleShape1 As PowerPacks.RectangleShape
    Friend WithEvents LabelHum As Label
    Friend WithEvents LabelTempC As Label
    Friend WithEvents ButtonShiftUp As Button
    Friend WithEvents ButtonShiftDn As Button
    Friend WithEvents Label14 As Label
    Friend WithEvents CheckBoxToolTips As CheckBox
    Friend WithEvents CheckBoxMaxMin As CheckBox
    Friend WithEvents Scale1 As Label
    Friend WithEvents Scale2 As Label
    Friend WithEvents Scale3 As Label
    Friend WithEvents Scale4 As Label
    Friend WithEvents Scale5 As Label
    Friend WithEvents Scale6 As Label
    Friend WithEvents Scale7 As Label
    Friend WithEvents Scale8 As Label
    Friend WithEvents Scale9 As Label
    Friend WithEvents Scale10 As Label
    Friend WithEvents Scale11 As Label
    Friend WithEvents Scale12 As Label
    Friend WithEvents Scale13 As Label
    Friend WithEvents Scale14 As Label
    Friend WithEvents Scale15 As Label
    Friend WithEvents Scale16 As Label
    Friend WithEvents Scale25 As Label
    Friend WithEvents Scale17 As Label
    Friend WithEvents Scale24 As Label
    Friend WithEvents Scale18 As Label
    Friend WithEvents Scale23 As Label
    Friend WithEvents Scale19 As Label
    Friend WithEvents Scale20 As Label
    Friend WithEvents Scale21 As Label
    Friend WithEvents Scale22 As Label
    Friend WithEvents MedianValue As TextBox
    Friend WithEvents MedianValueText As Label
    Friend WithEvents CheckBoxPPMenable As CheckBox
    Friend WithEvents RadioButtonDev1 As RadioButton
    Friend WithEvents RadioButtonDev2 As RadioButton
    Friend WithEvents PPMBox1 As GroupBox
    Friend WithEvents Label38 As Label
    Friend WithEvents MedianTempText As Label
    Friend WithEvents MedianTemp As TextBox
    Friend WithEvents PPMscaleText As Label
    Friend WithEvents PPMscalerangeentry As TextBox
    Friend WithEvents LabelPPMtop As Label
    Friend WithEvents GroupBoxMisc As GroupBox
    Friend WithEvents Label13 As Label
    Friend WithEvents RadioButtonPPMTempo As RadioButton
    Friend WithEvents RadioButtonPPMDev As RadioButton
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Xscale As Label
    Friend WithEvents MinsTotal As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents CheckBoxMedianV As CheckBox
    Friend WithEvents CheckBoxMedianT As CheckBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Screenshot As Button
    Friend WithEvents LogYaxis As CheckBox
    Friend WithEvents Xscaletotal As Label
    Friend WithEvents Loading As Label
    Friend WithEvents ButtonScrollLeftSMALL As Button
    Friend WithEvents ButtonScrollRightSMALL As Button
    Friend WithEvents CheckBoxYscaletidy As CheckBox
    Friend WithEvents RefreshFile As Button
    Friend WithEvents CheckDev1Point As CheckBox
    Friend WithEvents CheckDev1Line As CheckBox
    Friend WithEvents CheckDev2Line As CheckBox
    Friend WithEvents CheckDev2Point As CheckBox
    Friend WithEvents ShowFiles2 As Button
    Friend WithEvents Timer1 As Timer
    Friend WithEvents YaxisCheck1 As CheckBox
    Friend WithEvents YaxisCheck2 As CheckBox
    Friend WithEvents YaxisCheck3 As CheckBox
    Friend WithEvents YaxisCheck4 As CheckBox
    Friend WithEvents YaxisSave As Button
    Friend WithEvents YaxisBox1 As GroupBox
    Friend WithEvents Label19 As Label
    Friend WithEvents YaxisLoad As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label18 As Label
    Friend WithEvents Dev2MaxMin As TextBox
    Friend WithEvents Label22 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label20 As Label
    'Friend WithEvents ShapeContainer2 As PowerPacks.ShapeContainer
    'Friend WithEvents RectangleShape3 As PowerPacks.RectangleShape
    Friend WithEvents Label12 As Label
    Friend WithEvents Dev1MaxMin As TextBox
    Friend WithEvents CheckX1000000 As CheckBox
    Friend WithEvents CheckX1000 As CheckBox
    Friend WithEvents CheckBoxGraphOnly As CheckBox
    Friend WithEvents LabelPPMdegctop As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents DEV2avg As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents DEV1avg As TextBox
    Friend WithEvents GroupBoxMiscTempHum As GroupBox
    Friend WithEvents Label11 As Label
    Friend WithEvents TEMPavg As TextBox
    Friend WithEvents PleaseLoadCSV As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents MetadataChart As TextBox
    Friend WithEvents CheckBoxColours As CheckBox
    Friend WithEvents ScaleX28 As Label
    Friend WithEvents ScaleX1 As Label
    Friend WithEvents RMSaverageDev1 As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents RMSaverageDev2 As TextBox
    Friend WithEvents Label17 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents RMSwindow As TextBox
    Friend WithEvents Label23 As Label
    Friend WithEvents YaxisPerDiv As TextBox
    Friend WithEvents Label24 As Label
    Friend WithEvents Label25 As Label
    Friend WithEvents SampleRateSecs As TextBox
    Friend WithEvents Label26 As Label
    Friend WithEvents GroupBox3 As GroupBox
End Class
