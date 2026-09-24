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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Chart))
        Me.FormsPlot2 = New ScottPlot.WinForms.FormsPlot()
        Me.CurrentPosition = New System.Windows.Forms.TextBox()
        Me.TargetPosition = New System.Windows.Forms.TextBox()
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
        Me.MedianValue = New System.Windows.Forms.TextBox()
        Me.MedianValueText = New System.Windows.Forms.Label()
        Me.CheckBoxPPMenable = New System.Windows.Forms.CheckBox()
        Me.RadioButtonDev1 = New System.Windows.Forms.RadioButton()
        Me.RadioButtonDev2 = New System.Windows.Forms.RadioButton()
        Me.PPMBox1 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CheckBoxMedianT = New System.Windows.Forms.CheckBox()
        Me.CheckBoxMedianV = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.RadioButtonPPMTempoRolling = New System.Windows.Forms.RadioButton()
        Me.RadioButtonPPMTempoLinReg = New System.Windows.Forms.RadioButton()
        Me.RadioButtonPPMTempo = New System.Windows.Forms.RadioButton()
        Me.RadioButtonPPMDev = New System.Windows.Forms.RadioButton()
        Me.PPMscaleText = New System.Windows.Forms.Label()
        Me.PPMscalerangeentry = New System.Windows.Forms.TextBox()
        Me.MedianTempText = New System.Windows.Forms.Label()
        Me.MedianTemp = New System.Windows.Forms.TextBox()
        Me.LabelPPMtop = New System.Windows.Forms.Label()
        Me.GroupBoxMisc = New System.Windows.Forms.GroupBox()
        Me.CheckBoxColours = New System.Windows.Forms.CheckBox()
        Me.Xscale = New System.Windows.Forms.Label()
        Me.MinsTotal = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ShowFiles2 = New System.Windows.Forms.Button()
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
        Me.CheckPlaybackDev2SEM = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev2Stdev = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev2Mean = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev1Mean = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev1Stdev = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev1SEM = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev2Data = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev1Data = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev1ShortTermMean = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev1Deviation = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev1MaxDiff = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev2Deviation = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev2MaxDiff = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev2ShortTermMean = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev1Allan = New System.Windows.Forms.CheckBox()
        Me.CheckPlaybackDev2Allan = New System.Windows.Forms.CheckBox()
        Me.HUMavg = New System.Windows.Forms.TextBox()
        Me.PanelChartSplitter = New System.Windows.Forms.Panel()
        Me.ButtonSaveCSVMeta = New System.Windows.Forms.Button()
        Me.ButtonPlaybackHelp = New System.Windows.Forms.Button()
        Me.Xscaletotal = New System.Windows.Forms.Label()
        Me.Loading = New System.Windows.Forms.Label()
        Me.CheckDev1Point = New System.Windows.Forms.CheckBox()
        Me.CheckDev1Line = New System.Windows.Forms.CheckBox()
        Me.CheckDev2Line = New System.Windows.Forms.CheckBox()
        Me.CheckDev2Point = New System.Windows.Forms.CheckBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.YaxisBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxPBXYaxis = New System.Windows.Forms.CheckBox()
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
        Me.LabelPPMdegctop = New System.Windows.Forms.Label()
        Me.GroupBoxMiscTempHum = New System.Windows.Forms.GroupBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.ChartScaleHUMMax = New System.Windows.Forms.TextBox()
        Me.ChartScaleHUMMin = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.PleaseLoadCSV = New System.Windows.Forms.Label()
        Me.MetadataChart = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.LabelBottomChart = New System.Windows.Forms.Label()
        Me.LabelTopTopChart = New System.Windows.Forms.Label()
        Me.LabelPPMstats = New System.Windows.Forms.Label()
        Me.LabelDEV1 = New System.Windows.Forms.Label()
        Me.LabelSTATS = New System.Windows.Forms.Label()
        Me.LabelDEV2 = New System.Windows.Forms.Label()
        Me.LabelSTDEV = New System.Windows.Forms.Label()
        Me.LabelSEM = New System.Windows.Forms.Label()
        Me.LabelSTDEVscale = New System.Windows.Forms.Label()
        Me.LabelSEMscale = New System.Windows.Forms.Label()
        Me.LabelMean = New System.Windows.Forms.Label()
        Me.LabelSMean = New System.Windows.Forms.Label()
        Me.PPMBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBoxMisc.SuspendLayout()
        Me.YaxisBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBoxMiscTempHum.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'FormsPlot2
        '
        Me.FormsPlot2.BackColor = System.Drawing.SystemColors.Control
        Me.FormsPlot2.Location = New System.Drawing.Point(211, 287)
        Me.FormsPlot2.Name = "FormsPlot2"
        Me.FormsPlot2.Size = New System.Drawing.Size(938, 335)
        Me.FormsPlot2.TabIndex = 52
        '
        'CurrentPosition
        '
        Me.CurrentPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CurrentPosition.Location = New System.Drawing.Point(6, 63)
        Me.CurrentPosition.Name = "CurrentPosition"
        Me.CurrentPosition.ReadOnly = True
        Me.CurrentPosition.Size = New System.Drawing.Size(46, 20)
        Me.CurrentPosition.TabIndex = 56
        Me.CurrentPosition.Text = "0"
        '
        'TargetPosition
        '
        Me.TargetPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TargetPosition.Location = New System.Drawing.Point(6, 82)
        Me.TargetPosition.Name = "TargetPosition"
        Me.TargetPosition.ReadOnly = True
        Me.TargetPosition.Size = New System.Drawing.Size(46, 20)
        Me.TargetPosition.TabIndex = 57
        Me.TargetPosition.Text = "0"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(18, 63)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(53, 13)
        Me.Label27.TabIndex = 65
        Me.Label27.Text = "CSV File -"
        '
        'CSVfilenamePlayback
        '
        Me.CSVfilenamePlayback.Location = New System.Drawing.Point(74, 60)
        Me.CSVfilenamePlayback.Name = "CSVfilenamePlayback"
        Me.CSVfilenamePlayback.Size = New System.Drawing.Size(421, 20)
        Me.CSVfilenamePlayback.TabIndex = 64
        Me.CSVfilenamePlayback.WordWrap = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(50, 67)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 13)
        Me.Label1.TabIndex = 66
        Me.Label1.Text = "- Start Data"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(50, 86)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 67
        Me.Label2.Text = "- End Data"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(480, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 13)
        Me.Label3.TabIndex = 69
        Me.Label3.Text = "- Points (Dev)"
        '
        'CSVfileLines
        '
        Me.CSVfileLines.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSVfileLines.Location = New System.Drawing.Point(440, 25)
        Me.CSVfileLines.Name = "CSVfileLines"
        Me.CSVfileLines.ReadOnly = True
        Me.CSVfileLines.Size = New System.Drawing.Size(40, 20)
        Me.CSVfileLines.TabIndex = 70
        Me.CSVfileLines.Text = "0"
        Me.CSVfileLines.WordWrap = False
        '
        'YaxisMinimum
        '
        Me.YaxisMinimum.Location = New System.Drawing.Point(13, 160)
        Me.YaxisMinimum.Name = "YaxisMinimum"
        Me.YaxisMinimum.Size = New System.Drawing.Size(83, 20)
        Me.YaxisMinimum.TabIndex = 71
        Me.YaxisMinimum.Text = "0.000000"
        Me.YaxisMinimum.WordWrap = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(97, 137)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 13)
        Me.Label4.TabIndex = 72
        Me.Label4.Text = "- Y-axis Max."
        '
        'YaxisMax
        '
        Me.YaxisMax.AutoSize = True
        Me.YaxisMax.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.YaxisMax.Location = New System.Drawing.Point(91, 70)
        Me.YaxisMax.Name = "YaxisMax"
        Me.YaxisMax.Size = New System.Drawing.Size(64, 13)
        Me.YaxisMax.TabIndex = 74
        Me.YaxisMax.Text = "- Y-axis Min."
        '
        'YaxisMaximum
        '
        Me.YaxisMaximum.Location = New System.Drawing.Point(13, 134)
        Me.YaxisMaximum.Name = "YaxisMaximum"
        Me.YaxisMaximum.Size = New System.Drawing.Size(83, 20)
        Me.YaxisMaximum.TabIndex = 73
        Me.YaxisMaximum.Text = "15.000000"
        Me.YaxisMaximum.WordWrap = False
        '
        'BrowseToFile
        '
        Me.BrowseToFile.Location = New System.Drawing.Point(15, 9)
        Me.BrowseToFile.Name = "BrowseToFile"
        Me.BrowseToFile.Size = New System.Drawing.Size(69, 45)
        Me.BrowseToFile.TabIndex = 84
        Me.BrowseToFile.Text = "LOAD" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & ".CSV FILE"
        Me.ToolTip1.SetToolTip(Me.BrowseToFile, "Load CSV from disk.")
        Me.BrowseToFile.UseVisualStyleBackColor = True
        '
        'ButtonDisplayAll
        '
        Me.ButtonDisplayAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDisplayAll.Location = New System.Drawing.Point(164, 41)
        Me.ButtonDisplayAll.Name = "ButtonDisplayAll"
        Me.ButtonDisplayAll.Size = New System.Drawing.Size(88, 22)
        Me.ButtonDisplayAll.TabIndex = 85
        Me.ButtonDisplayAll.Text = "Zoom All"
        Me.ToolTip1.SetToolTip(Me.ButtonDisplayAll, "Display all of chart")
        Me.ButtonDisplayAll.UseVisualStyleBackColor = True
        '
        'DeviceName1
        '
        Me.DeviceName1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DeviceName1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.DeviceName1.Location = New System.Drawing.Point(6, 26)
        Me.DeviceName1.Name = "DeviceName1"
        Me.DeviceName1.ReadOnly = True
        Me.DeviceName1.Size = New System.Drawing.Size(121, 20)
        Me.DeviceName1.TabIndex = 86
        Me.DeviceName1.WordWrap = False
        '
        'DeviceName2
        '
        Me.DeviceName2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DeviceName2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.DeviceName2.Location = New System.Drawing.Point(221, 25)
        Me.DeviceName2.Name = "DeviceName2"
        Me.DeviceName2.ReadOnly = True
        Me.DeviceName2.Size = New System.Drawing.Size(123, 20)
        Me.DeviceName2.TabIndex = 91
        Me.DeviceName2.WordWrap = False
        '
        'PlaybackTemp
        '
        Me.PlaybackTemp.AutoSize = True
        Me.PlaybackTemp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PlaybackTemp.Location = New System.Drawing.Point(9, 24)
        Me.PlaybackTemp.Name = "PlaybackTemp"
        Me.PlaybackTemp.Size = New System.Drawing.Size(53, 17)
        Me.PlaybackTemp.TabIndex = 92
        Me.PlaybackTemp.Text = "Temp"
        Me.ToolTip1.SetToolTip(Me.PlaybackTemp, "From CSV")
        Me.PlaybackTemp.UseVisualStyleBackColor = True
        '
        'PlaybackHum
        '
        Me.PlaybackHum.AutoSize = True
        Me.PlaybackHum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PlaybackHum.Location = New System.Drawing.Point(122, 24)
        Me.PlaybackHum.Name = "PlaybackHum"
        Me.PlaybackHum.Size = New System.Drawing.Size(51, 17)
        Me.PlaybackHum.TabIndex = 93
        Me.PlaybackHum.Text = "Hum."
        Me.ToolTip1.SetToolTip(Me.PlaybackHum, "From CSV")
        Me.PlaybackHum.UseVisualStyleBackColor = True
        '
        'ChartScaleMax
        '
        Me.ChartScaleMax.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChartScaleMax.Location = New System.Drawing.Point(6, 65)
        Me.ChartScaleMax.Name = "ChartScaleMax"
        Me.ChartScaleMax.Size = New System.Drawing.Size(26, 20)
        Me.ChartScaleMax.TabIndex = 94
        Me.ChartScaleMax.Text = "50"
        Me.ChartScaleMax.WordWrap = False
        '
        'ChartScaleMin
        '
        Me.ChartScaleMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChartScaleMin.Location = New System.Drawing.Point(6, 84)
        Me.ChartScaleMin.Name = "ChartScaleMin"
        Me.ChartScaleMin.Size = New System.Drawing.Size(26, 20)
        Me.ChartScaleMin.TabIndex = 95
        Me.ChartScaleMin.Text = "15"
        Me.ChartScaleMin.WordWrap = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(32, 69)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 96
        Me.Label6.Text = "Temp Max."
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(32, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(57, 13)
        Me.Label7.TabIndex = 97
        Me.Label7.Text = "Temp Min."
        '
        'ButtonSaveSettings
        '
        Me.ButtonSaveSettings.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ButtonSaveSettings.Location = New System.Drawing.Point(391, 32)
        Me.ButtonSaveSettings.Name = "ButtonSaveSettings"
        Me.ButtonSaveSettings.Size = New System.Drawing.Size(50, 22)
        Me.ButtonSaveSettings.TabIndex = 98
        Me.ButtonSaveSettings.Text = "Save"
        Me.ToolTip1.SetToolTip(Me.ButtonSaveSettings, resources.GetString("ButtonSaveSettings.ToolTip"))
        Me.ButtonSaveSettings.UseVisualStyleBackColor = True
        '
        'LabelHum
        '
        Me.LabelHum.AutoSize = True
        Me.LabelHum.BackColor = System.Drawing.Color.Black
        Me.LabelHum.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelHum.ForeColor = System.Drawing.Color.DodgerBlue
        Me.LabelHum.Location = New System.Drawing.Point(1282, 205)
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
        Me.LabelTempC.Location = New System.Drawing.Point(1243, 205)
        Me.LabelTempC.Name = "LabelTempC"
        Me.LabelTempC.Size = New System.Drawing.Size(38, 15)
        Me.LabelTempC.TabIndex = 101
        Me.LabelTempC.Text = "DegC"
        '
        'MedianValue
        '
        Me.MedianValue.Location = New System.Drawing.Point(895, 42)
        Me.MedianValue.Name = "MedianValue"
        Me.MedianValue.Size = New System.Drawing.Size(75, 20)
        Me.MedianValue.TabIndex = 144
        Me.MedianValue.Text = "10.00000"
        Me.MedianValue.WordWrap = False
        '
        'MedianValueText
        '
        Me.MedianValueText.AutoSize = True
        Me.MedianValueText.Location = New System.Drawing.Point(971, 45)
        Me.MedianValueText.Name = "MedianValueText"
        Me.MedianValueText.Size = New System.Drawing.Size(67, 13)
        Me.MedianValueText.TabIndex = 145
        Me.MedianValueText.Text = "- Initial Value"
        '
        'CheckBoxPPMenable
        '
        Me.CheckBoxPPMenable.AutoSize = True
        Me.CheckBoxPPMenable.Location = New System.Drawing.Point(895, 20)
        Me.CheckBoxPPMenable.Name = "CheckBoxPPMenable"
        Me.CheckBoxPPMenable.Size = New System.Drawing.Size(85, 17)
        Me.CheckBoxPPMenable.TabIndex = 146
        Me.CheckBoxPPMenable.Text = "Enable PPM"
        Me.ToolTip1.SetToolTip(Me.CheckBoxPPMenable, "Calculated in Playback Chart")
        Me.CheckBoxPPMenable.UseVisualStyleBackColor = True
        '
        'RadioButtonDev1
        '
        Me.RadioButtonDev1.AutoSize = True
        Me.RadioButtonDev1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonDev1.Location = New System.Drawing.Point(106, 15)
        Me.RadioButtonDev1.Name = "RadioButtonDev1"
        Me.RadioButtonDev1.Size = New System.Drawing.Size(54, 17)
        Me.RadioButtonDev1.TabIndex = 147
        Me.RadioButtonDev1.TabStop = True
        Me.RadioButtonDev1.Text = "Dev 1"
        Me.RadioButtonDev1.UseVisualStyleBackColor = True
        '
        'RadioButtonDev2
        '
        Me.RadioButtonDev2.AutoSize = True
        Me.RadioButtonDev2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonDev2.Location = New System.Drawing.Point(180, 15)
        Me.RadioButtonDev2.Name = "RadioButtonDev2"
        Me.RadioButtonDev2.Size = New System.Drawing.Size(54, 17)
        Me.RadioButtonDev2.TabIndex = 148
        Me.RadioButtonDev2.TabStop = True
        Me.RadioButtonDev2.Text = "Dev 2"
        Me.RadioButtonDev2.UseVisualStyleBackColor = True
        '
        'PPMBox1
        '
        Me.PPMBox1.Controls.Add(Me.Label5)
        Me.PPMBox1.Controls.Add(Me.CheckBoxMedianT)
        Me.PPMBox1.Controls.Add(Me.CheckBoxMedianV)
        Me.PPMBox1.Controls.Add(Me.Panel1)
        Me.PPMBox1.Controls.Add(Me.PPMscaleText)
        Me.PPMBox1.Controls.Add(Me.PPMscalerangeentry)
        Me.PPMBox1.Controls.Add(Me.MedianTempText)
        Me.PPMBox1.Controls.Add(Me.RadioButtonDev2)
        Me.PPMBox1.Controls.Add(Me.MedianTemp)
        Me.PPMBox1.Controls.Add(Me.RadioButtonDev1)
        Me.PPMBox1.Enabled = False
        Me.PPMBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PPMBox1.Location = New System.Drawing.Point(888, 4)
        Me.PPMBox1.Name = "PPMBox1"
        Me.PPMBox1.Size = New System.Drawing.Size(467, 87)
        Me.PPMBox1.TabIndex = 149
        Me.PPMBox1.TabStop = False
        Me.PPMBox1.Text = "PPM DEVIATION / TEMPCO"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(175, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 13)
        Me.Label5.TabIndex = 579
        Me.Label5.Text = "(main chart)"
        '
        'CheckBoxMedianT
        '
        Me.CheckBoxMedianT.AutoSize = True
        Me.CheckBoxMedianT.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxMedianT.Location = New System.Drawing.Point(168, 65)
        Me.CheckBoxMedianT.Name = "CheckBoxMedianT"
        Me.CheckBoxMedianT.Size = New System.Drawing.Size(79, 17)
        Me.CheckBoxMedianT.TabIndex = 196
        Me.CheckBoxMedianT.Text = "- From CSV"
        Me.CheckBoxMedianT.UseVisualStyleBackColor = True
        '
        'CheckBoxMedianV
        '
        Me.CheckBoxMedianV.AutoSize = True
        Me.CheckBoxMedianV.Checked = True
        Me.CheckBoxMedianV.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxMedianV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxMedianV.Location = New System.Drawing.Point(168, 41)
        Me.CheckBoxMedianV.Name = "CheckBoxMedianV"
        Me.CheckBoxMedianV.Size = New System.Drawing.Size(79, 17)
        Me.CheckBoxMedianV.TabIndex = 195
        Me.CheckBoxMedianV.Text = "- From CSV"
        Me.CheckBoxMedianV.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.RadioButtonPPMTempoRolling)
        Me.Panel1.Controls.Add(Me.RadioButtonPPMTempoLinReg)
        Me.Panel1.Controls.Add(Me.RadioButtonPPMTempo)
        Me.Panel1.Controls.Add(Me.RadioButtonPPMDev)
        Me.Panel1.Location = New System.Drawing.Point(338, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(125, 75)
        Me.Panel1.TabIndex = 157
        '
        'RadioButtonPPMTempoRolling
        '
        Me.RadioButtonPPMTempoRolling.AutoSize = True
        Me.RadioButtonPPMTempoRolling.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonPPMTempoRolling.Location = New System.Drawing.Point(9, 58)
        Me.RadioButtonPPMTempoRolling.Name = "RadioButtonPPMTempoRolling"
        Me.RadioButtonPPMTempoRolling.Size = New System.Drawing.Size(117, 17)
        Me.RadioButtonPPMTempoRolling.TabIndex = 158
        Me.RadioButtonPPMTempoRolling.Text = "PPM/DegC (Trend)"
        Me.ToolTip1.SetToolTip(Me.RadioButtonPPMTempoRolling, resources.GetString("RadioButtonPPMTempoRolling.ToolTip"))
        Me.RadioButtonPPMTempoRolling.UseVisualStyleBackColor = True
        '
        'RadioButtonPPMTempoLinReg
        '
        Me.RadioButtonPPMTempoLinReg.AutoSize = True
        Me.RadioButtonPPMTempoLinReg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonPPMTempoLinReg.Location = New System.Drawing.Point(9, 40)
        Me.RadioButtonPPMTempoLinReg.Name = "RadioButtonPPMTempoLinReg"
        Me.RadioButtonPPMTempoLinReg.Size = New System.Drawing.Size(100, 17)
        Me.RadioButtonPPMTempoLinReg.TabIndex = 157
        Me.RadioButtonPPMTempoLinReg.Text = "PPM/DegC (Fit)"
        Me.ToolTip1.SetToolTip(Me.RadioButtonPPMTempoLinReg, resources.GetString("RadioButtonPPMTempoLinReg.ToolTip"))
        Me.RadioButtonPPMTempoLinReg.UseVisualStyleBackColor = True
        '
        'RadioButtonPPMTempo
        '
        Me.RadioButtonPPMTempo.AutoSize = True
        Me.RadioButtonPPMTempo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonPPMTempo.Location = New System.Drawing.Point(9, 22)
        Me.RadioButtonPPMTempo.Name = "RadioButtonPPMTempo"
        Me.RadioButtonPPMTempo.Size = New System.Drawing.Size(112, 17)
        Me.RadioButtonPPMTempo.TabIndex = 156
        Me.RadioButtonPPMTempo.Text = "PPM/DegC (point)"
        Me.ToolTip1.SetToolTip(Me.RadioButtonPPMTempo, "Tempco per point vs. the Initial Value/Initial Temp baseline")
        Me.RadioButtonPPMTempo.UseVisualStyleBackColor = True
        '
        'RadioButtonPPMDev
        '
        Me.RadioButtonPPMDev.AutoSize = True
        Me.RadioButtonPPMDev.Checked = True
        Me.RadioButtonPPMDev.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButtonPPMDev.Location = New System.Drawing.Point(9, 4)
        Me.RadioButtonPPMDev.Name = "RadioButtonPPMDev"
        Me.RadioButtonPPMDev.Size = New System.Drawing.Size(96, 17)
        Me.RadioButtonPPMDev.TabIndex = 155
        Me.RadioButtonPPMDev.TabStop = True
        Me.RadioButtonPPMDev.Text = "PPM Deviation"
        Me.ToolTip1.SetToolTip(Me.RadioButtonPPMDev, "Plots reading deviation from the Initial Value, in ppm - no temperature involved")
        Me.RadioButtonPPMDev.UseVisualStyleBackColor = True
        '
        'PPMscaleText
        '
        Me.PPMscaleText.AutoSize = True
        Me.PPMscaleText.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PPMscaleText.Location = New System.Drawing.Point(289, 16)
        Me.PPMscaleText.Name = "PPMscaleText"
        Me.PPMscaleText.Size = New System.Drawing.Size(40, 13)
        Me.PPMscaleText.TabIndex = 154
        Me.PPMscaleText.Text = "- Scale"
        '
        'PPMscalerangeentry
        '
        Me.PPMscalerangeentry.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PPMscalerangeentry.Location = New System.Drawing.Point(261, 13)
        Me.PPMscalerangeentry.Name = "PPMscalerangeentry"
        Me.PPMscalerangeentry.Size = New System.Drawing.Size(26, 20)
        Me.PPMscalerangeentry.TabIndex = 151
        Me.PPMscalerangeentry.Text = "6"
        Me.PPMscalerangeentry.WordWrap = False
        '
        'MedianTempText
        '
        Me.MedianTempText.AutoSize = True
        Me.MedianTempText.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MedianTempText.Location = New System.Drawing.Point(82, 65)
        Me.MedianTempText.Name = "MedianTempText"
        Me.MedianTempText.Size = New System.Drawing.Size(67, 13)
        Me.MedianTempText.TabIndex = 150
        Me.MedianTempText.Text = "- Initial Temp"
        '
        'MedianTemp
        '
        Me.MedianTemp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MedianTemp.Location = New System.Drawing.Point(6, 62)
        Me.MedianTemp.Name = "MedianTemp"
        Me.MedianTemp.Size = New System.Drawing.Size(75, 20)
        Me.MedianTemp.TabIndex = 150
        Me.MedianTemp.Text = "23.4"
        Me.MedianTemp.WordWrap = False
        '
        'LabelPPMtop
        '
        Me.LabelPPMtop.AutoSize = True
        Me.LabelPPMtop.BackColor = System.Drawing.Color.Black
        Me.LabelPPMtop.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelPPMtop.ForeColor = System.Drawing.Color.White
        Me.LabelPPMtop.Location = New System.Drawing.Point(1319, 205)
        Me.LabelPPMtop.Name = "LabelPPMtop"
        Me.LabelPPMtop.Size = New System.Drawing.Size(34, 15)
        Me.LabelPPMtop.TabIndex = 152
        Me.LabelPPMtop.Text = "PPM"
        '
        'GroupBoxMisc
        '
        Me.GroupBoxMisc.Controls.Add(Me.CheckBoxColours)
        Me.GroupBoxMisc.Enabled = False
        Me.GroupBoxMisc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBoxMisc.Location = New System.Drawing.Point(1279, 93)
        Me.GroupBoxMisc.Name = "GroupBoxMisc"
        Me.GroupBoxMisc.Size = New System.Drawing.Size(76, 109)
        Me.GroupBoxMisc.TabIndex = 155
        Me.GroupBoxMisc.TabStop = False
        Me.GroupBoxMisc.Text = "MISC."
        '
        'CheckBoxColours
        '
        Me.CheckBoxColours.AutoSize = True
        Me.CheckBoxColours.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxColours.Location = New System.Drawing.Point(8, 24)
        Me.CheckBoxColours.Name = "CheckBoxColours"
        Me.CheckBoxColours.Size = New System.Drawing.Size(64, 17)
        Me.CheckBoxColours.TabIndex = 117
        Me.CheckBoxColours.Text = "Light M."
        Me.ToolTip1.SetToolTip(Me.CheckBoxColours, "Set form to light mode, better for printing.")
        Me.CheckBoxColours.UseVisualStyleBackColor = True
        '
        'Xscale
        '
        Me.Xscale.AutoSize = True
        Me.Xscale.BackColor = System.Drawing.Color.Black
        Me.Xscale.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Xscale.ForeColor = System.Drawing.Color.Orange
        Me.Xscale.Location = New System.Drawing.Point(651, 210)
        Me.Xscale.Name = "Xscale"
        Me.Xscale.Size = New System.Drawing.Size(80, 16)
        Me.Xscale.TabIndex = 192
        Me.Xscale.Text = "Time (mins):"
        '
        'MinsTotal
        '
        Me.MinsTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(50, 48)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(62, 13)
        Me.Label16.TabIndex = 194
        Me.Label16.Text = "- Mins Total"
        '
        'ToolTip1
        '
        '
        'ShowFiles2
        '
        Me.ShowFiles2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ShowFiles2.Location = New System.Drawing.Point(391, 9)
        Me.ShowFiles2.Name = "ShowFiles2"
        Me.ShowFiles2.Size = New System.Drawing.Size(104, 22)
        Me.ShowFiles2.TabIndex = 558
        Me.ShowFiles2.Text = "\WinGPIBdata"
        Me.ToolTip1.SetToolTip(Me.ShowFiles2, "Launch Windows File Explorer")
        Me.ShowFiles2.UseVisualStyleBackColor = True
        '
        'CheckX1000
        '
        Me.CheckX1000.AutoSize = True
        Me.CheckX1000.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckX1000.Location = New System.Drawing.Point(444, 87)
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
        Me.CheckX1000000.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckX1000000.Location = New System.Drawing.Point(493, 87)
        Me.CheckX1000000.Name = "CheckX1000000"
        Me.CheckX1000000.Size = New System.Drawing.Size(61, 17)
        Me.CheckX1000000.TabIndex = 577
        Me.CheckX1000000.Text = "x1000k"
        Me.ToolTip1.SetToolTip(Me.CheckX1000000, "I.E. Vdc to uVdc")
        Me.CheckX1000000.UseVisualStyleBackColor = True
        '
        'DEV2avg
        '
        Me.DEV2avg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DEV2avg.Location = New System.Drawing.Point(348, 25)
        Me.DEV2avg.Name = "DEV2avg"
        Me.DEV2avg.Size = New System.Drawing.Size(26, 20)
        Me.DEV2avg.TabIndex = 579
        Me.DEV2avg.Text = "0"
        Me.ToolTip1.SetToolTip(Me.DEV2avg, "Set to '0' to disable averaging" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Range = 0 to 100.")
        '
        'DEV1avg
        '
        Me.DEV1avg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DEV1avg.Location = New System.Drawing.Point(131, 26)
        Me.DEV1avg.Name = "DEV1avg"
        Me.DEV1avg.Size = New System.Drawing.Size(26, 20)
        Me.DEV1avg.TabIndex = 199
        Me.DEV1avg.Text = "0"
        Me.ToolTip1.SetToolTip(Me.DEV1avg, "Set to '0' to disable averaging" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Range = 0 to 100.")
        '
        'TEMPavg
        '
        Me.TEMPavg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TEMPavg.Location = New System.Drawing.Point(6, 46)
        Me.TEMPavg.Name = "TEMPavg"
        Me.TEMPavg.Size = New System.Drawing.Size(26, 20)
        Me.TEMPavg.TabIndex = 582
        Me.TEMPavg.Text = "0"
        Me.ToolTip1.SetToolTip(Me.TEMPavg, "Set to '0' to disable averaging" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Range = 0 to 100.")
        Me.TEMPavg.WordWrap = False
        '
        'Dev2MaxMin
        '
        Me.Dev2MaxMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev2MaxMin.Location = New System.Drawing.Point(221, 54)
        Me.Dev2MaxMin.Name = "Dev2MaxMin"
        Me.Dev2MaxMin.ReadOnly = True
        Me.Dev2MaxMin.Size = New System.Drawing.Size(75, 20)
        Me.Dev2MaxMin.TabIndex = 570
        Me.ToolTip1.SetToolTip(Me.Dev2MaxMin, "Maximum - Minimum for the chart")
        Me.Dev2MaxMin.WordWrap = False
        '
        'RMSaverageDev2
        '
        Me.RMSaverageDev2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RMSaverageDev2.Location = New System.Drawing.Point(221, 81)
        Me.RMSaverageDev2.Name = "RMSaverageDev2"
        Me.RMSaverageDev2.ReadOnly = True
        Me.RMSaverageDev2.Size = New System.Drawing.Size(75, 20)
        Me.RMSaverageDev2.TabIndex = 583
        Me.ToolTip1.SetToolTip(Me.RMSaverageDev2, "Noise calculation whilst taking into consideration drift over time.")
        Me.RMSaverageDev2.WordWrap = False
        '
        'RMSaverageDev1
        '
        Me.RMSaverageDev1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RMSaverageDev1.Location = New System.Drawing.Point(6, 81)
        Me.RMSaverageDev1.Name = "RMSaverageDev1"
        Me.RMSaverageDev1.ReadOnly = True
        Me.RMSaverageDev1.Size = New System.Drawing.Size(75, 20)
        Me.RMSaverageDev1.TabIndex = 581
        Me.ToolTip1.SetToolTip(Me.RMSaverageDev1, "Noise calculation whilst taking into consideration drift over time.")
        Me.RMSaverageDev1.WordWrap = False
        '
        'Dev1MaxMin
        '
        Me.Dev1MaxMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev1MaxMin.Location = New System.Drawing.Point(6, 54)
        Me.Dev1MaxMin.Name = "Dev1MaxMin"
        Me.Dev1MaxMin.ReadOnly = True
        Me.Dev1MaxMin.Size = New System.Drawing.Size(75, 20)
        Me.Dev1MaxMin.TabIndex = 575
        Me.ToolTip1.SetToolTip(Me.Dev1MaxMin, "Maximum - Minimum for the chart")
        Me.Dev1MaxMin.WordWrap = False
        '
        'RMSwindow
        '
        Me.RMSwindow.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RMSwindow.Location = New System.Drawing.Point(440, 54)
        Me.RMSwindow.Name = "RMSwindow"
        Me.RMSwindow.Size = New System.Drawing.Size(40, 20)
        Me.RMSwindow.TabIndex = 586
        Me.ToolTip1.SetToolTip(Me.RMSwindow, resources.GetString("RMSwindow.ToolTip"))
        Me.RMSwindow.WordWrap = False
        '
        'CheckPlaybackDev2SEM
        '
        Me.CheckPlaybackDev2SEM.AutoSize = True
        Me.CheckPlaybackDev2SEM.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev2SEM.Location = New System.Drawing.Point(4, 66)
        Me.CheckPlaybackDev2SEM.Name = "CheckPlaybackDev2SEM"
        Me.CheckPlaybackDev2SEM.Size = New System.Drawing.Size(49, 17)
        Me.CheckPlaybackDev2SEM.TabIndex = 593
        Me.CheckPlaybackDev2SEM.Text = "SEM"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev2SEM, "From CSV")
        Me.CheckPlaybackDev2SEM.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev2Stdev
        '
        Me.CheckPlaybackDev2Stdev.AutoSize = True
        Me.CheckPlaybackDev2Stdev.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev2Stdev.Location = New System.Drawing.Point(4, 49)
        Me.CheckPlaybackDev2Stdev.Name = "CheckPlaybackDev2Stdev"
        Me.CheckPlaybackDev2Stdev.Size = New System.Drawing.Size(62, 17)
        Me.CheckPlaybackDev2Stdev.TabIndex = 592
        Me.CheckPlaybackDev2Stdev.Text = "STDEV"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev2Stdev, "From CSV")
        Me.CheckPlaybackDev2Stdev.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev2Mean
        '
        Me.CheckPlaybackDev2Mean.AutoSize = True
        Me.CheckPlaybackDev2Mean.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev2Mean.Location = New System.Drawing.Point(4, 32)
        Me.CheckPlaybackDev2Mean.Name = "CheckPlaybackDev2Mean"
        Me.CheckPlaybackDev2Mean.Size = New System.Drawing.Size(53, 17)
        Me.CheckPlaybackDev2Mean.TabIndex = 591
        Me.CheckPlaybackDev2Mean.Text = "Mean"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev2Mean, "From CSV")
        Me.CheckPlaybackDev2Mean.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev1Mean
        '
        Me.CheckPlaybackDev1Mean.AutoSize = True
        Me.CheckPlaybackDev1Mean.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev1Mean.Location = New System.Drawing.Point(5, 32)
        Me.CheckPlaybackDev1Mean.Name = "CheckPlaybackDev1Mean"
        Me.CheckPlaybackDev1Mean.Size = New System.Drawing.Size(53, 17)
        Me.CheckPlaybackDev1Mean.TabIndex = 590
        Me.CheckPlaybackDev1Mean.Text = "Mean"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev1Mean, "From CSV")
        Me.CheckPlaybackDev1Mean.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev1Stdev
        '
        Me.CheckPlaybackDev1Stdev.AutoSize = True
        Me.CheckPlaybackDev1Stdev.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev1Stdev.Location = New System.Drawing.Point(5, 49)
        Me.CheckPlaybackDev1Stdev.Name = "CheckPlaybackDev1Stdev"
        Me.CheckPlaybackDev1Stdev.Size = New System.Drawing.Size(62, 17)
        Me.CheckPlaybackDev1Stdev.TabIndex = 589
        Me.CheckPlaybackDev1Stdev.Text = "STDEV"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev1Stdev, "From CSV")
        Me.CheckPlaybackDev1Stdev.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev1SEM
        '
        Me.CheckPlaybackDev1SEM.AutoSize = True
        Me.CheckPlaybackDev1SEM.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev1SEM.Location = New System.Drawing.Point(5, 66)
        Me.CheckPlaybackDev1SEM.Name = "CheckPlaybackDev1SEM"
        Me.CheckPlaybackDev1SEM.Size = New System.Drawing.Size(49, 17)
        Me.CheckPlaybackDev1SEM.TabIndex = 588
        Me.CheckPlaybackDev1SEM.Text = "SEM"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev1SEM, "From CSV")
        Me.CheckPlaybackDev1SEM.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev2Data
        '
        Me.CheckPlaybackDev2Data.AutoSize = True
        Me.CheckPlaybackDev2Data.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev2Data.Location = New System.Drawing.Point(4, 15)
        Me.CheckPlaybackDev2Data.Name = "CheckPlaybackDev2Data"
        Me.CheckPlaybackDev2Data.Size = New System.Drawing.Size(49, 17)
        Me.CheckPlaybackDev2Data.TabIndex = 587
        Me.CheckPlaybackDev2Data.Text = "Data"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev2Data, "From CSV")
        Me.CheckPlaybackDev2Data.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev1Data
        '
        Me.CheckPlaybackDev1Data.AutoSize = True
        Me.CheckPlaybackDev1Data.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev1Data.Location = New System.Drawing.Point(5, 15)
        Me.CheckPlaybackDev1Data.Name = "CheckPlaybackDev1Data"
        Me.CheckPlaybackDev1Data.Size = New System.Drawing.Size(49, 17)
        Me.CheckPlaybackDev1Data.TabIndex = 578
        Me.CheckPlaybackDev1Data.Text = "Data"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev1Data, "From CSV")
        Me.CheckPlaybackDev1Data.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev1ShortTermMean
        '
        Me.CheckPlaybackDev1ShortTermMean.AutoSize = True
        Me.CheckPlaybackDev1ShortTermMean.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev1ShortTermMean.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CheckPlaybackDev1ShortTermMean.Location = New System.Drawing.Point(69, 49)
        Me.CheckPlaybackDev1ShortTermMean.Name = "CheckPlaybackDev1ShortTermMean"
        Me.CheckPlaybackDev1ShortTermMean.Size = New System.Drawing.Size(108, 17)
        Me.CheckPlaybackDev1ShortTermMean.TabIndex = 596
        Me.CheckPlaybackDev1ShortTermMean.Text = "Short Term Mean"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev1ShortTermMean, "Calculated in Playback Chart")
        Me.CheckPlaybackDev1ShortTermMean.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev1Deviation
        '
        Me.CheckPlaybackDev1Deviation.AutoSize = True
        Me.CheckPlaybackDev1Deviation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev1Deviation.Location = New System.Drawing.Point(69, 32)
        Me.CheckPlaybackDev1Deviation.Name = "CheckPlaybackDev1Deviation"
        Me.CheckPlaybackDev1Deviation.Size = New System.Drawing.Size(97, 17)
        Me.CheckPlaybackDev1Deviation.TabIndex = 595
        Me.CheckPlaybackDev1Deviation.Text = "PPM Deviation"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev1Deviation, "From CSV")
        Me.CheckPlaybackDev1Deviation.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev1MaxDiff
        '
        Me.CheckPlaybackDev1MaxDiff.AutoSize = True
        Me.CheckPlaybackDev1MaxDiff.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev1MaxDiff.Location = New System.Drawing.Point(69, 15)
        Me.CheckPlaybackDev1MaxDiff.Name = "CheckPlaybackDev1MaxDiff"
        Me.CheckPlaybackDev1MaxDiff.Size = New System.Drawing.Size(68, 17)
        Me.CheckPlaybackDev1MaxDiff.TabIndex = 594
        Me.CheckPlaybackDev1MaxDiff.Text = "Max Diff."
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev1MaxDiff, "From CSV")
        Me.CheckPlaybackDev1MaxDiff.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev2Deviation
        '
        Me.CheckPlaybackDev2Deviation.AutoSize = True
        Me.CheckPlaybackDev2Deviation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev2Deviation.Location = New System.Drawing.Point(68, 32)
        Me.CheckPlaybackDev2Deviation.Name = "CheckPlaybackDev2Deviation"
        Me.CheckPlaybackDev2Deviation.Size = New System.Drawing.Size(97, 17)
        Me.CheckPlaybackDev2Deviation.TabIndex = 597
        Me.CheckPlaybackDev2Deviation.Text = "PPM Deviation"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev2Deviation, "From CSV")
        Me.CheckPlaybackDev2Deviation.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev2MaxDiff
        '
        Me.CheckPlaybackDev2MaxDiff.AutoSize = True
        Me.CheckPlaybackDev2MaxDiff.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev2MaxDiff.Location = New System.Drawing.Point(68, 15)
        Me.CheckPlaybackDev2MaxDiff.Name = "CheckPlaybackDev2MaxDiff"
        Me.CheckPlaybackDev2MaxDiff.Size = New System.Drawing.Size(68, 17)
        Me.CheckPlaybackDev2MaxDiff.TabIndex = 596
        Me.CheckPlaybackDev2MaxDiff.Text = "Max Diff."
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev2MaxDiff, "From CSV")
        Me.CheckPlaybackDev2MaxDiff.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev2ShortTermMean
        '
        Me.CheckPlaybackDev2ShortTermMean.AutoSize = True
        Me.CheckPlaybackDev2ShortTermMean.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev2ShortTermMean.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CheckPlaybackDev2ShortTermMean.Location = New System.Drawing.Point(68, 49)
        Me.CheckPlaybackDev2ShortTermMean.Name = "CheckPlaybackDev2ShortTermMean"
        Me.CheckPlaybackDev2ShortTermMean.Size = New System.Drawing.Size(108, 17)
        Me.CheckPlaybackDev2ShortTermMean.TabIndex = 597
        Me.CheckPlaybackDev2ShortTermMean.Text = "Short Term Mean"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev2ShortTermMean, "Calculated in Playback Chart")
        Me.CheckPlaybackDev2ShortTermMean.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev1Allan
        '
        Me.CheckPlaybackDev1Allan.AutoSize = True
        Me.CheckPlaybackDev1Allan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev1Allan.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CheckPlaybackDev1Allan.Location = New System.Drawing.Point(69, 66)
        Me.CheckPlaybackDev1Allan.Name = "CheckPlaybackDev1Allan"
        Me.CheckPlaybackDev1Allan.Size = New System.Drawing.Size(97, 17)
        Me.CheckPlaybackDev1Allan.TabIndex = 597
        Me.CheckPlaybackDev1Allan.Text = "Allan Deviation"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev1Allan, "Calculated in Playback Chart")
        Me.CheckPlaybackDev1Allan.UseVisualStyleBackColor = True
        '
        'CheckPlaybackDev2Allan
        '
        Me.CheckPlaybackDev2Allan.AutoSize = True
        Me.CheckPlaybackDev2Allan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckPlaybackDev2Allan.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CheckPlaybackDev2Allan.Location = New System.Drawing.Point(68, 66)
        Me.CheckPlaybackDev2Allan.Name = "CheckPlaybackDev2Allan"
        Me.CheckPlaybackDev2Allan.Size = New System.Drawing.Size(97, 17)
        Me.CheckPlaybackDev2Allan.TabIndex = 598
        Me.CheckPlaybackDev2Allan.Text = "Allan Deviation"
        Me.ToolTip1.SetToolTip(Me.CheckPlaybackDev2Allan, "Calculated in Playback Chart")
        Me.CheckPlaybackDev2Allan.UseVisualStyleBackColor = True
        '
        'HUMavg
        '
        Me.HUMavg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HUMavg.Location = New System.Drawing.Point(119, 46)
        Me.HUMavg.Name = "HUMavg"
        Me.HUMavg.Size = New System.Drawing.Size(26, 20)
        Me.HUMavg.TabIndex = 588
        Me.HUMavg.Text = "0"
        Me.ToolTip1.SetToolTip(Me.HUMavg, "Set to '0' to disable averaging" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Range = 0 to 100.")
        Me.HUMavg.WordWrap = False
        '
        'PanelChartSplitter
        '
        Me.PanelChartSplitter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PanelChartSplitter.BackColor = System.Drawing.Color.DimGray
        Me.PanelChartSplitter.Cursor = System.Windows.Forms.Cursors.HSplit
        Me.PanelChartSplitter.Location = New System.Drawing.Point(585, 644)
        Me.PanelChartSplitter.Name = "PanelChartSplitter"
        Me.PanelChartSplitter.Size = New System.Drawing.Size(200, 5)
        Me.PanelChartSplitter.TabIndex = 585
        Me.ToolTip1.SetToolTip(Me.PanelChartSplitter, "Grab and move to re-size charts")
        '
        'ButtonSaveCSVMeta
        '
        Me.ButtonSaveCSVMeta.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ButtonSaveCSVMeta.Location = New System.Drawing.Point(333, 10)
        Me.ButtonSaveCSVMeta.Name = "ButtonSaveCSVMeta"
        Me.ButtonSaveCSVMeta.Size = New System.Drawing.Size(44, 44)
        Me.ButtonSaveCSVMeta.TabIndex = 589
        Me.ButtonSaveCSVMeta.Text = "Save" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Meta"
        Me.ToolTip1.SetToolTip(Me.ButtonSaveCSVMeta, "Save Metadata to CSV file")
        Me.ButtonSaveCSVMeta.UseVisualStyleBackColor = True
        '
        'ButtonPlaybackHelp
        '
        Me.ButtonPlaybackHelp.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ButtonPlaybackHelp.Location = New System.Drawing.Point(445, 32)
        Me.ButtonPlaybackHelp.Name = "ButtonPlaybackHelp"
        Me.ButtonPlaybackHelp.Size = New System.Drawing.Size(50, 22)
        Me.ButtonPlaybackHelp.TabIndex = 581
        Me.ButtonPlaybackHelp.Text = "Help"
        Me.ButtonPlaybackHelp.UseVisualStyleBackColor = True
        '
        'Xscaletotal
        '
        Me.Xscaletotal.AutoSize = True
        Me.Xscaletotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Xscaletotal.Location = New System.Drawing.Point(734, 210)
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
        Me.Loading.Location = New System.Drawing.Point(471, 464)
        Me.Loading.Name = "Loading"
        Me.Loading.Size = New System.Drawing.Size(440, 37)
        Me.Loading.TabIndex = 197
        Me.Loading.Text = "Loading CSV, Please Wait....."
        '
        'CheckDev1Point
        '
        Me.CheckDev1Point.AutoSize = True
        Me.CheckDev1Point.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckDev1Point.Location = New System.Drawing.Point(151, 86)
        Me.CheckDev1Point.Name = "CheckDev1Point"
        Me.CheckDev1Point.Size = New System.Drawing.Size(50, 17)
        Me.CheckDev1Point.TabIndex = 202
        Me.CheckDev1Point.Text = "Point"
        Me.CheckDev1Point.UseVisualStyleBackColor = True
        '
        'CheckDev1Line
        '
        Me.CheckDev1Line.AutoSize = True
        Me.CheckDev1Line.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckDev1Line.Location = New System.Drawing.Point(151, 67)
        Me.CheckDev1Line.Name = "CheckDev1Line"
        Me.CheckDev1Line.Size = New System.Drawing.Size(46, 17)
        Me.CheckDev1Line.TabIndex = 203
        Me.CheckDev1Line.Text = "Line"
        Me.CheckDev1Line.UseVisualStyleBackColor = True
        '
        'CheckDev2Line
        '
        Me.CheckDev2Line.AutoSize = True
        Me.CheckDev2Line.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckDev2Line.Location = New System.Drawing.Point(367, 68)
        Me.CheckDev2Line.Name = "CheckDev2Line"
        Me.CheckDev2Line.Size = New System.Drawing.Size(46, 17)
        Me.CheckDev2Line.TabIndex = 206
        Me.CheckDev2Line.Text = "Line"
        Me.CheckDev2Line.UseVisualStyleBackColor = True
        '
        'CheckDev2Point
        '
        Me.CheckDev2Point.AutoSize = True
        Me.CheckDev2Point.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckDev2Point.Location = New System.Drawing.Point(367, 87)
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
        'YaxisBox1
        '
        Me.YaxisBox1.Controls.Add(Me.CheckBoxPBXYaxis)
        Me.YaxisBox1.Controls.Add(Me.ButtonDisplayAll)
        Me.YaxisBox1.Controls.Add(Me.YaxisMax)
        Me.YaxisBox1.Enabled = False
        Me.YaxisBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.YaxisBox1.Location = New System.Drawing.Point(6, 93)
        Me.YaxisBox1.Name = "YaxisBox1"
        Me.YaxisBox1.Size = New System.Drawing.Size(258, 109)
        Me.YaxisBox1.TabIndex = 566
        Me.YaxisBox1.TabStop = False
        Me.YaxisBox1.Text = "X && Y-AXIS SCALES"
        '
        'CheckBoxPBXYaxis
        '
        Me.CheckBoxPBXYaxis.AutoSize = True
        Me.CheckBoxPBXYaxis.Checked = True
        Me.CheckBoxPBXYaxis.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxPBXYaxis.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxPBXYaxis.Location = New System.Drawing.Point(7, 19)
        Me.CheckBoxPBXYaxis.Name = "CheckBoxPBXYaxis"
        Me.CheckBoxPBXYaxis.Size = New System.Drawing.Size(146, 17)
        Me.CheckBoxPBXYaxis.TabIndex = 86
        Me.CheckBoxPBXYaxis.Text = "AutoScale X-axis && Y-axis"
        Me.CheckBoxPBXYaxis.UseVisualStyleBackColor = True
        '
        'SampleRateSecs
        '
        Me.SampleRateSecs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.Label25.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.Location = New System.Drawing.Point(50, 29)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(101, 13)
        Me.Label25.TabIndex = 570
        Me.Label25.Text = "- Sample Rate Secs"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(297, 60)
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
        Me.GroupBox2.Controls.Add(Me.CheckDev1Line)
        Me.GroupBox2.Controls.Add(Me.CSVfileLines)
        Me.GroupBox2.Controls.Add(Me.DeviceName2)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.DeviceName1)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox2.Location = New System.Drawing.Point(455, 93)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(568, 109)
        Me.GroupBox2.TabIndex = 572
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "DEVICES"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(480, 58)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(76, 13)
        Me.Label21.TabIndex = 585
        Me.Label21.Text = "- RMS window"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(297, 86)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(67, 13)
        Me.Label17.TabIndex = 584
        Me.Label17.Text = "- RMS Noise"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(81, 85)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(67, 13)
        Me.Label10.TabIndex = 582
        Me.Label10.Text = "- RMS Noise"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(375, 29)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(35, 13)
        Me.Label9.TabIndex = 580
        Me.Label9.Text = "- Avg."
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(158, 28)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(35, 13)
        Me.Label8.TabIndex = 200
        Me.Label8.Text = "- Avg."
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(81, 60)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(53, 13)
        Me.Label12.TabIndex = 576
        Me.Label12.Text = "- Max-Min"
        '
        'LabelPPMdegctop
        '
        Me.LabelPPMdegctop.AutoSize = True
        Me.LabelPPMdegctop.BackColor = System.Drawing.Color.Black
        Me.LabelPPMdegctop.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelPPMdegctop.ForeColor = System.Drawing.Color.White
        Me.LabelPPMdegctop.Location = New System.Drawing.Point(1319, 217)
        Me.LabelPPMdegctop.Name = "LabelPPMdegctop"
        Me.LabelPPMdegctop.Size = New System.Drawing.Size(34, 12)
        Me.LabelPPMdegctop.TabIndex = 573
        Me.LabelPPMdegctop.Text = " /DegC"
        '
        'GroupBoxMiscTempHum
        '
        Me.GroupBoxMiscTempHum.Controls.Add(Me.Label13)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.HUMavg)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.ChartScaleHUMMax)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.ChartScaleHUMMin)
        Me.GroupBoxMiscTempHum.Controls.Add(Me.Label14)
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
        Me.GroupBoxMiscTempHum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBoxMiscTempHum.Location = New System.Drawing.Point(1044, 93)
        Me.GroupBoxMiscTempHum.Name = "GroupBoxMiscTempHum"
        Me.GroupBoxMiscTempHum.Size = New System.Drawing.Size(217, 109)
        Me.GroupBoxMiscTempHum.TabIndex = 156
        Me.GroupBoxMiscTempHum.TabStop = False
        Me.GroupBoxMiscTempHum.Text = "TEMP/HUM"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(145, 50)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(54, 13)
        Me.Label13.TabIndex = 587
        Me.Label13.Text = "Hum Avg."
        '
        'ChartScaleHUMMax
        '
        Me.ChartScaleHUMMax.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChartScaleHUMMax.Location = New System.Drawing.Point(119, 65)
        Me.ChartScaleHUMMax.Name = "ChartScaleHUMMax"
        Me.ChartScaleHUMMax.Size = New System.Drawing.Size(26, 20)
        Me.ChartScaleHUMMax.TabIndex = 583
        Me.ChartScaleHUMMax.Text = "50"
        Me.ChartScaleHUMMax.WordWrap = False
        '
        'ChartScaleHUMMin
        '
        Me.ChartScaleHUMMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChartScaleHUMMin.Location = New System.Drawing.Point(119, 84)
        Me.ChartScaleHUMMin.Name = "ChartScaleHUMMin"
        Me.ChartScaleHUMMin.Size = New System.Drawing.Size(26, 20)
        Me.ChartScaleHUMMin.TabIndex = 584
        Me.ChartScaleHUMMin.Text = "15"
        Me.ChartScaleHUMMin.WordWrap = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(145, 69)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(55, 13)
        Me.Label14.TabIndex = 585
        Me.Label14.Text = "Hum Max."
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(145, 88)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(52, 13)
        Me.Label15.TabIndex = 586
        Me.Label15.Text = "Hum Min."
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(32, 50)
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
        Me.PleaseLoadCSV.Location = New System.Drawing.Point(494, 427)
        Me.PleaseLoadCSV.Name = "PleaseLoadCSV"
        Me.PleaseLoadCSV.Size = New System.Drawing.Size(343, 37)
        Me.PleaseLoadCSV.TabIndex = 574
        Me.PleaseLoadCSV.Text = "Please load a CSV file!"
        '
        'MetadataChart
        '
        Me.MetadataChart.Location = New System.Drawing.Point(99, 10)
        Me.MetadataChart.Multiline = True
        Me.MetadataChart.Name = "MetadataChart"
        Me.MetadataChart.Size = New System.Drawing.Size(233, 44)
        Me.MetadataChart.TabIndex = 575
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.CurrentPosition)
        Me.GroupBox3.Controls.Add(Me.MinsTotal)
        Me.GroupBox3.Controls.Add(Me.SampleRateSecs)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.Label25)
        Me.GroupBox3.Controls.Add(Me.Label16)
        Me.GroupBox3.Controls.Add(Me.TargetPosition)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(283, 93)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(153, 109)
        Me.GroupBox3.TabIndex = 577
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "CSV"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.CheckPlaybackDev1Allan)
        Me.GroupBox4.Controls.Add(Me.CheckPlaybackDev1ShortTermMean)
        Me.GroupBox4.Controls.Add(Me.CheckPlaybackDev1Deviation)
        Me.GroupBox4.Controls.Add(Me.CheckPlaybackDev1MaxDiff)
        Me.GroupBox4.Controls.Add(Me.CheckPlaybackDev1Data)
        Me.GroupBox4.Controls.Add(Me.CheckPlaybackDev1Mean)
        Me.GroupBox4.Controls.Add(Me.CheckPlaybackDev1SEM)
        Me.GroupBox4.Controls.Add(Me.CheckPlaybackDev1Stdev)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(502, 4)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(178, 87)
        Me.GroupBox4.TabIndex = 578
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "DEV 1 TRACES"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.CheckPlaybackDev2Allan)
        Me.GroupBox5.Controls.Add(Me.CheckPlaybackDev2ShortTermMean)
        Me.GroupBox5.Controls.Add(Me.CheckPlaybackDev2Deviation)
        Me.GroupBox5.Controls.Add(Me.CheckPlaybackDev2Mean)
        Me.GroupBox5.Controls.Add(Me.CheckPlaybackDev2Data)
        Me.GroupBox5.Controls.Add(Me.CheckPlaybackDev2MaxDiff)
        Me.GroupBox5.Controls.Add(Me.CheckPlaybackDev2Stdev)
        Me.GroupBox5.Controls.Add(Me.CheckPlaybackDev2SEM)
        Me.GroupBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(693, 4)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(178, 87)
        Me.GroupBox5.TabIndex = 580
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "DEV 2 TRACES"
        '
        'LabelBottomChart
        '
        Me.LabelBottomChart.AutoSize = True
        Me.LabelBottomChart.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.LabelBottomChart.Location = New System.Drawing.Point(542, 785)
        Me.LabelBottomChart.Name = "LabelBottomChart"
        Me.LabelBottomChart.Size = New System.Drawing.Size(292, 13)
        Me.LabelBottomChart.TabIndex = 582
        Me.LabelBottomChart.Text = "Initial baseline for PPM Deviation value is derived from Stats."
        Me.LabelBottomChart.Visible = False
        '
        'LabelTopTopChart
        '
        Me.LabelTopTopChart.AutoSize = True
        Me.LabelTopTopChart.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.LabelTopTopChart.Location = New System.Drawing.Point(517, 230)
        Me.LabelTopTopChart.Name = "LabelTopTopChart"
        Me.LabelTopTopChart.Size = New System.Drawing.Size(390, 13)
        Me.LabelTopTopChart.TabIndex = 583
        Me.LabelTopTopChart.Text = "Data, Mean, Short Term Mean. PPM Deviation / Tempco, Temperature, Humidity"
        Me.LabelTopTopChart.Visible = False
        '
        'LabelPPMstats
        '
        Me.LabelPPMstats.AutoSize = True
        Me.LabelPPMstats.BackColor = System.Drawing.Color.Black
        Me.LabelPPMstats.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelPPMstats.ForeColor = System.Drawing.Color.White
        Me.LabelPPMstats.Location = New System.Drawing.Point(1239, 644)
        Me.LabelPPMstats.Name = "LabelPPMstats"
        Me.LabelPPMstats.Size = New System.Drawing.Size(34, 15)
        Me.LabelPPMstats.TabIndex = 584
        Me.LabelPPMstats.Text = "PPM"
        '
        'LabelDEV1
        '
        Me.LabelDEV1.AutoSize = True
        Me.LabelDEV1.BackColor = System.Drawing.Color.Yellow
        Me.LabelDEV1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDEV1.ForeColor = System.Drawing.Color.Black
        Me.LabelDEV1.Location = New System.Drawing.Point(5, 205)
        Me.LabelDEV1.Name = "LabelDEV1"
        Me.LabelDEV1.Size = New System.Drawing.Size(38, 15)
        Me.LabelDEV1.TabIndex = 586
        Me.LabelDEV1.Text = "DEV1"
        '
        'LabelSTATS
        '
        Me.LabelSTATS.AutoSize = True
        Me.LabelSTATS.BackColor = System.Drawing.Color.Black
        Me.LabelSTATS.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelSTATS.ForeColor = System.Drawing.Color.White
        Me.LabelSTATS.Location = New System.Drawing.Point(12, 644)
        Me.LabelSTATS.Name = "LabelSTATS"
        Me.LabelSTATS.Size = New System.Drawing.Size(62, 15)
        Me.LabelSTATS.TabIndex = 587
        Me.LabelSTATS.Text = "MAX DIFF"
        '
        'LabelDEV2
        '
        Me.LabelDEV2.AutoSize = True
        Me.LabelDEV2.BackColor = System.Drawing.Color.Cyan
        Me.LabelDEV2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDEV2.ForeColor = System.Drawing.Color.Black
        Me.LabelDEV2.Location = New System.Drawing.Point(46, 205)
        Me.LabelDEV2.Name = "LabelDEV2"
        Me.LabelDEV2.Size = New System.Drawing.Size(38, 15)
        Me.LabelDEV2.TabIndex = 588
        Me.LabelDEV2.Text = "DEV2"
        '
        'LabelSTDEV
        '
        Me.LabelSTDEV.AutoSize = True
        Me.LabelSTDEV.BackColor = System.Drawing.Color.Black
        Me.LabelSTDEV.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelSTDEV.ForeColor = System.Drawing.Color.White
        Me.LabelSTDEV.Location = New System.Drawing.Point(1274, 644)
        Me.LabelSTDEV.Name = "LabelSTDEV"
        Me.LabelSTDEV.Size = New System.Drawing.Size(46, 15)
        Me.LabelSTDEV.TabIndex = 590
        Me.LabelSTDEV.Text = "STDEV"
        '
        'LabelSEM
        '
        Me.LabelSEM.AutoSize = True
        Me.LabelSEM.BackColor = System.Drawing.Color.Black
        Me.LabelSEM.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelSEM.ForeColor = System.Drawing.Color.White
        Me.LabelSEM.Location = New System.Drawing.Point(1321, 644)
        Me.LabelSEM.Name = "LabelSEM"
        Me.LabelSEM.Size = New System.Drawing.Size(34, 15)
        Me.LabelSEM.TabIndex = 591
        Me.LabelSEM.Text = "SEM"
        '
        'LabelSTDEVscale
        '
        Me.LabelSTDEVscale.AutoSize = True
        Me.LabelSTDEVscale.BackColor = System.Drawing.Color.Black
        Me.LabelSTDEVscale.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelSTDEVscale.ForeColor = System.Drawing.Color.White
        Me.LabelSTDEVscale.Location = New System.Drawing.Point(1279, 783)
        Me.LabelSTDEVscale.Name = "LabelSTDEVscale"
        Me.LabelSTDEVscale.Size = New System.Drawing.Size(32, 13)
        Me.LabelSTDEVscale.TabIndex = 592
        Me.LabelSTDEVscale.Text = "scale"
        '
        'LabelSEMscale
        '
        Me.LabelSEMscale.AutoSize = True
        Me.LabelSEMscale.BackColor = System.Drawing.Color.Black
        Me.LabelSEMscale.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelSEMscale.ForeColor = System.Drawing.Color.White
        Me.LabelSEMscale.Location = New System.Drawing.Point(1317, 783)
        Me.LabelSEMscale.Name = "LabelSEMscale"
        Me.LabelSEMscale.Size = New System.Drawing.Size(32, 13)
        Me.LabelSEMscale.TabIndex = 593
        Me.LabelSEMscale.Text = "scale"
        '
        'LabelMean
        '
        Me.LabelMean.AutoSize = True
        Me.LabelMean.BackColor = System.Drawing.Color.Black
        Me.LabelMean.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMean.ForeColor = System.Drawing.Color.White
        Me.LabelMean.Location = New System.Drawing.Point(5, 223)
        Me.LabelMean.Name = "LabelMean"
        Me.LabelMean.Size = New System.Drawing.Size(42, 15)
        Me.LabelMean.TabIndex = 594
        Me.LabelMean.Text = "MEAN"
        '
        'LabelSMean
        '
        Me.LabelSMean.AutoSize = True
        Me.LabelSMean.BackColor = System.Drawing.Color.Black
        Me.LabelSMean.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelSMean.ForeColor = System.Drawing.Color.White
        Me.LabelSMean.Location = New System.Drawing.Point(48, 223)
        Me.LabelSMean.Name = "LabelSMean"
        Me.LabelSMean.Size = New System.Drawing.Size(53, 15)
        Me.LabelSMean.TabIndex = 595
        Me.LabelSMean.Text = "S.MEAN"
        '
        'Chart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1359, 801)
        Me.Controls.Add(Me.LabelSMean)
        Me.Controls.Add(Me.LabelMean)
        Me.Controls.Add(Me.LabelSEMscale)
        Me.Controls.Add(Me.LabelSTDEVscale)
        Me.Controls.Add(Me.LabelSEM)
        Me.Controls.Add(Me.LabelSTDEV)
        Me.Controls.Add(Me.ButtonSaveCSVMeta)
        Me.Controls.Add(Me.LabelDEV2)
        Me.Controls.Add(Me.LabelSTATS)
        Me.Controls.Add(Me.LabelDEV1)
        Me.Controls.Add(Me.PanelChartSplitter)
        Me.Controls.Add(Me.LabelPPMstats)
        Me.Controls.Add(Me.LabelTopTopChart)
        Me.Controls.Add(Me.LabelBottomChart)
        Me.Controls.Add(Me.ButtonPlaybackHelp)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.MetadataChart)
        Me.Controls.Add(Me.PleaseLoadCSV)
        Me.Controls.Add(Me.GroupBoxMiscTempHum)
        Me.Controls.Add(Me.LabelPPMtop)
        Me.Controls.Add(Me.LabelPPMdegctop)
        Me.Controls.Add(Me.ButtonSaveSettings)
        Me.Controls.Add(Me.ShowFiles2)
        Me.Controls.Add(Me.Loading)
        Me.Controls.Add(Me.Xscaletotal)
        Me.Controls.Add(Me.Xscale)
        Me.Controls.Add(Me.CheckBoxPPMenable)
        Me.Controls.Add(Me.MedianValueText)
        Me.Controls.Add(Me.MedianValue)
        Me.Controls.Add(Me.LabelTempC)
        Me.Controls.Add(Me.LabelHum)
        Me.Controls.Add(Me.BrowseToFile)
        Me.Controls.Add(Me.YaxisMaximum)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.YaxisMinimum)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.CSVfilenamePlayback)
        Me.Controls.Add(Me.FormsPlot2)
        Me.Controls.Add(Me.PPMBox1)
        Me.Controls.Add(Me.GroupBoxMisc)
        Me.Controls.Add(Me.YaxisBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "Chart"
        Me.Text = "WinGPIB    Playback Chart    (Free for Non-Commercial Use • Support WinGPIB — see" &
    " About)"
        Me.PPMBox1.ResumeLayout(False)
        Me.PPMBox1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBoxMisc.ResumeLayout(False)
        Me.GroupBoxMisc.PerformLayout()
        Me.YaxisBox1.ResumeLayout(False)
        Me.YaxisBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBoxMiscTempHum.ResumeLayout(False)
        Me.GroupBoxMiscTempHum.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents FormsPlot2 As ScottPlot.WinForms.FormsPlot
    Friend WithEvents CurrentPosition As TextBox
    Friend WithEvents TargetPosition As TextBox
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
    Friend WithEvents MedianValue As TextBox
    Friend WithEvents MedianValueText As Label
    Friend WithEvents CheckBoxPPMenable As CheckBox
    Friend WithEvents RadioButtonDev1 As RadioButton
    Friend WithEvents RadioButtonDev2 As RadioButton
    Friend WithEvents PPMBox1 As GroupBox
    Friend WithEvents MedianTempText As Label
    Friend WithEvents MedianTemp As TextBox
    Friend WithEvents PPMscaleText As Label
    Friend WithEvents PPMscalerangeentry As TextBox
    Friend WithEvents LabelPPMtop As Label
    Friend WithEvents GroupBoxMisc As GroupBox
    Friend WithEvents RadioButtonPPMTempo As RadioButton
    Friend WithEvents RadioButtonPPMDev As RadioButton
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Xscale As Label
    Friend WithEvents MinsTotal As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents CheckBoxMedianV As CheckBox
    Friend WithEvents CheckBoxMedianT As CheckBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Xscaletotal As Label
    Friend WithEvents Loading As Label
    Friend WithEvents CheckDev1Point As CheckBox
    Friend WithEvents CheckDev1Line As CheckBox
    Friend WithEvents CheckDev2Line As CheckBox
    Friend WithEvents CheckDev2Point As CheckBox
    Friend WithEvents ShowFiles2 As Button
    Friend WithEvents Timer1 As Timer
    Friend WithEvents YaxisBox1 As GroupBox
    Friend WithEvents Dev2MaxMin As TextBox
    Friend WithEvents Label22 As Label
    Friend WithEvents GroupBox2 As GroupBox
    'Friend WithEvents ShapeContainer2 As PowerPacks.ShapeContainer
    'Friend WithEvents RectangleShape3 As PowerPacks.RectangleShape
    Friend WithEvents Label12 As Label
    Friend WithEvents Dev1MaxMin As TextBox
    Friend WithEvents CheckX1000000 As CheckBox
    Friend WithEvents CheckX1000 As CheckBox
    Friend WithEvents LabelPPMdegctop As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents DEV2avg As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents DEV1avg As TextBox
    Friend WithEvents GroupBoxMiscTempHum As GroupBox
    Friend WithEvents Label11 As Label
    Friend WithEvents TEMPavg As TextBox
    Friend WithEvents PleaseLoadCSV As Label
    Friend WithEvents MetadataChart As TextBox
    Friend WithEvents CheckBoxColours As CheckBox
    Friend WithEvents RMSaverageDev1 As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents RMSaverageDev2 As TextBox
    Friend WithEvents Label17 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents RMSwindow As TextBox
    Friend WithEvents Label25 As Label
    Friend WithEvents SampleRateSecs As TextBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents CheckPlaybackDev1Data As CheckBox
    Friend WithEvents CheckPlaybackDev2Data As CheckBox
    Friend WithEvents CheckPlaybackDev1Mean As CheckBox
    Friend WithEvents CheckPlaybackDev1Stdev As CheckBox
    Friend WithEvents CheckPlaybackDev1SEM As CheckBox
    Friend WithEvents CheckPlaybackDev2SEM As CheckBox
    Friend WithEvents CheckPlaybackDev2Stdev As CheckBox
    Friend WithEvents CheckPlaybackDev2Mean As CheckBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents CheckPlaybackDev2Deviation As CheckBox
    Friend WithEvents CheckPlaybackDev2MaxDiff As CheckBox
    Friend WithEvents CheckPlaybackDev1Deviation As CheckBox
    Friend WithEvents CheckPlaybackDev1MaxDiff As CheckBox
    Friend WithEvents Label5 As Label
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents CheckPlaybackDev1ShortTermMean As CheckBox
    Friend WithEvents CheckPlaybackDev2ShortTermMean As CheckBox
    Friend WithEvents CheckPlaybackDev1Allan As CheckBox
    Friend WithEvents CheckPlaybackDev2Allan As CheckBox
    Friend WithEvents ButtonPlaybackHelp As Button
    Friend WithEvents RadioButtonPPMTempoLinReg As RadioButton
    Friend WithEvents RadioButtonPPMTempoRolling As RadioButton
    Friend WithEvents LabelBottomChart As Label
    Friend WithEvents LabelTopTopChart As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents HUMavg As TextBox
    Friend WithEvents ChartScaleHUMMax As TextBox
    Friend WithEvents ChartScaleHUMMin As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents CheckBoxPBXYaxis As CheckBox
    Friend WithEvents LabelPPMstats As Label
    Friend WithEvents PanelChartSplitter As Panel
    Friend WithEvents LabelDEV1 As Label
    Friend WithEvents LabelSTATS As Label
    Friend WithEvents LabelDEV2 As Label
    Friend WithEvents ButtonSaveCSVMeta As Button
    Friend WithEvents LabelSTDEV As Label
    Friend WithEvents LabelSEM As Label
    Friend WithEvents LabelSTDEVscale As Label
    Friend WithEvents LabelSEMscale As Label
    Friend WithEvents LabelMean As Label
    Friend WithEvents LabelSMean As Label
End Class
