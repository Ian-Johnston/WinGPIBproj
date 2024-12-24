<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Formtest
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Formtest))
        Dim ChartArea3 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend3 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series3 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.SerialPort = New System.IO.Ports.SerialPort(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer3 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer4 = New System.Windows.Forms.Timer(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btndevlist = New System.Windows.Forms.Button()
        Me.ButtonReset = New System.Windows.Forms.Button()
        Me.ButtonDev12Run = New System.Windows.Forms.Button()
        Me.Dev2removeletters = New System.Windows.Forms.CheckBox()
        Me.CheckBoxSendBlockingDev2 = New System.Windows.Forms.CheckBox()
        Me.btnq2b = New System.Windows.Forms.Button()
        Me.btns2c = New System.Windows.Forms.Button()
        Me.btnq2a = New System.Windows.Forms.Button()
        Me.CommandStart2run = New System.Windows.Forms.TextBox()
        Me.CommandStop2 = New System.Windows.Forms.TextBox()
        Me.CommandStart2 = New System.Windows.Forms.TextBox()
        Me.ButtonDev2Run = New System.Windows.Forms.Button()
        Me.Dev1removeletters = New System.Windows.Forms.CheckBox()
        Me.CheckBoxSendBlockingDev1 = New System.Windows.Forms.CheckBox()
        Me.btnq1b = New System.Windows.Forms.Button()
        Me.btns1c = New System.Windows.Forms.Button()
        Me.btnq1a = New System.Windows.Forms.Button()
        Me.CommandStart1run = New System.Windows.Forms.TextBox()
        Me.CommandStop1 = New System.Windows.Forms.TextBox()
        Me.CommandStart1 = New System.Windows.Forms.TextBox()
        Me.ButtonDev1Run = New System.Windows.Forms.Button()
        Me.lstIntf3 = New System.Windows.Forms.ComboBox()
        Me.ButtonStart = New System.Windows.Forms.Button()
        Me.ButtonEnd = New System.Windows.Forms.Button()
        Me.ComboBoxPort = New System.Windows.Forms.ComboBox()
        Me.ShowFiles = New System.Windows.Forms.Button()
        Me.ResetCSV = New System.Windows.Forms.Button()
        Me.ButtonExportCSV = New System.Windows.Forms.Button()
        Me.ButtonClearChart = New System.Windows.Forms.Button()
        Me.ButtonIanWebsite = New System.Windows.Forms.Button()
        Me.ButtonSaveSettings = New System.Windows.Forms.Button()
        Me.ButtonNotePad2 = New System.Windows.Forms.Button()
        Me.ShowFilesCalRam = New System.Windows.Forms.Button()
        Me.ShowFiles2 = New System.Windows.Forms.Button()
        Me.exportBTN = New System.Windows.Forms.Button()
        Me.ButtonSetPrecalvars = New System.Windows.Forms.Button()
        Me.nplc4_BTN = New System.Windows.Forms.Button()
        Me.nplc3_BTN = New System.Windows.Forms.Button()
        Me.nplc2_BTN = New System.Windows.Forms.Button()
        Me.getcal_BTN = New System.Windows.Forms.Button()
        Me.nplc_BTN = New System.Windows.Forms.Button()
        Me.Abort_BTN = New System.Windows.Forms.Button()
        Me.precal_BTN = New System.Windows.Forms.Button()
        Me.SavePDVS2Eprom = New System.Windows.Forms.Button()
        Me.Dev23457Aseven = New System.Windows.Forms.CheckBox()
        Me.Dev13457Aseven = New System.Windows.Forms.CheckBox()
        Me.ShowFiles3 = New System.Windows.Forms.Button()
        Me.CSVfilename = New System.Windows.Forms.TextBox()
        Me.CSVfilepath = New System.Windows.Forms.TextBox()
        Me.Dev2TerminatorEnable2 = New System.Windows.Forms.CheckBox()
        Me.Dev2TerminatorEnable = New System.Windows.Forms.CheckBox()
        Me.Dev1TerminatorEnable2 = New System.Windows.Forms.CheckBox()
        Me.Dev1TerminatorEnable = New System.Windows.Forms.CheckBox()
        Me.EditMode = New System.Windows.Forms.CheckBox()
        Me.CheckboxCSVlimit = New System.Windows.Forms.CheckBox()
        Me.CheckboxCSVlimitMins = New System.Windows.Forms.CheckBox()
        Me.TextFilenameAppend = New System.Windows.Forms.TextBox()
        Me.Dev1K2001isolatedata = New System.Windows.Forms.CheckBox()
        Me.Dev2K2001isolatedata = New System.Windows.Forms.CheckBox()
        Me.Div1000Dev2 = New System.Windows.Forms.CheckBox()
        Me.Div1000Dev1 = New System.Windows.Forms.CheckBox()
        Me.Mult1000Dev1 = New System.Windows.Forms.CheckBox()
        Me.Mult1000Dev2 = New System.Windows.Forms.CheckBox()
        Me.CalOnExisting = New System.Windows.Forms.Button()
        Me.Label253 = New System.Windows.Forms.Label()
        Me.Label251 = New System.Windows.Forms.Label()
        Me.Label252 = New System.Windows.Forms.Label()
        Me.Label250 = New System.Windows.Forms.Label()
        Me.ButtonSetRetrievedVars = New System.Windows.Forms.Button()
        Me.ButtonAutoSET = New System.Windows.Forms.Button()
        Me.ButtonAutomV = New System.Windows.Forms.Button()
        Me.EnableAutoYChart1 = New System.Windows.Forms.CheckBox()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.ButtonPauseChart = New System.Windows.Forms.Button()
        Me.ButtonSaveLiveSettings = New System.Windows.Forms.Button()
        Me.CheckBoxAZERO = New System.Windows.Forms.CheckBox()
        Me.ClearLOGdisp = New System.Windows.Forms.Button()
        Me.ButtonSaveDefs = New System.Windows.Forms.Button()
        Me.ButtonLoadDefs = New System.Windows.Forms.Button()
        Me.Dev1IntEnable = New System.Windows.Forms.CheckBox()
        Me.Dev2IntEnable = New System.Windows.Forms.CheckBox()
        Me.ButtonRefreshPorts = New System.Windows.Forms.Button()
        Me.TextBoxProtocolInput = New System.Windows.Forms.TextBox()
        Me.TextBoxResult = New System.Windows.Forms.TextBox()
        Me.TextBoxFinalTempValue = New System.Windows.Forms.TextBox()
        Me.TextBoxRegex = New System.Windows.Forms.TextBox()
        Me.TextBoxSerialPortBaud = New System.Windows.Forms.TextBox()
        Me.TextBoxSerialPortBits = New System.Windows.Forms.TextBox()
        Me.TextBoxSerialPortParity = New System.Windows.Forms.TextBox()
        Me.TextBoxSerialPortStop = New System.Windows.Forms.TextBox()
        Me.TextBoxSerialPortHand = New System.Windows.Forms.TextBox()
        Me.ButtonSaveTempHumSettings = New System.Windows.Forms.Button()
        Me.Dev1Regex = New System.Windows.Forms.CheckBox()
        Me.Dev2Regex = New System.Windows.Forms.CheckBox()
        Me.Dev1DecimalNumDPs = New System.Windows.Forms.TextBox()
        Me.Dev2DecimalNumDPs = New System.Windows.Forms.TextBox()
        Me.Dev2pauseDurationInSeconds = New System.Windows.Forms.TextBox()
        Me.Dev2runStopwatchEveryInMins = New System.Windows.Forms.TextBox()
        Me.Dev2SampleRate = New System.Windows.Forms.TextBox()
        Me.Dev1pauseDurationInSeconds = New System.Windows.Forms.TextBox()
        Me.Dev1runStopwatchEveryInMins = New System.Windows.Forms.TextBox()
        Me.Dev1SampleRate = New System.Windows.Forms.TextBox()
        Me.XaxisPoints = New System.Windows.Forms.TextBox()
        Me.PDVS2miniSave = New System.Windows.Forms.Button()
        Me.Dev1TextResponse = New System.Windows.Forms.CheckBox()
        Me.Dev2TextResponse = New System.Windows.Forms.CheckBox()
        Me.ButtonDev1PreRun = New System.Windows.Forms.Button()
        Me.ButtonDev2PreRun = New System.Windows.Forms.Button()
        Me.txtq2d = New System.Windows.Forms.TextBox()
        Me.txtq1d = New System.Windows.Forms.TextBox()
        Me.Dev1SendQuery = New System.Windows.Forms.CheckBox()
        Me.Dev2SendQuery = New System.Windows.Forms.CheckBox()
        Me.ClearEventLOG = New System.Windows.Forms.Button()
        Me.Dev1Units = New System.Windows.Forms.TextBox()
        Me.Dev2Units = New System.Windows.Forms.TextBox()
        Me.TextBoxHumUnits = New System.Windows.Forms.TextBox()
        Me.TextBoxTempUnits = New System.Windows.Forms.TextBox()
        Me.TempOffset = New System.Windows.Forms.TextBox()
        Me.ButtonRefreshPorts1 = New System.Windows.Forms.Button()
        Me.WryTech = New System.Windows.Forms.CheckBox()
        Me.btncreate2 = New System.Windows.Forms.Button()
        Me.btncreate3 = New System.Windows.Forms.Button()
        Me.btncreate = New System.Windows.Forms.Button()
        Me.noEOI = New System.Windows.Forms.CheckBox()
        Me.TextBoxTempHumSample = New System.Windows.Forms.TextBox()
        Me.txtname3 = New System.Windows.Forms.TextBox()
        Me.ShowFilesCalRamR6581 = New System.Windows.Forms.Button()
        Me.ButtonAvailableComPorts = New System.Windows.Forms.Button()
        Me.ButtonJsonViewer = New System.Windows.Forms.Button()
        Me.ButtonOpenR6581fileSelectJson = New System.Windows.Forms.Button()
        Me.ButtonOpenR6581fileJson = New System.Windows.Forms.Button()
        Me.ButtonOpenR6581file = New System.Windows.Forms.Button()
        Me.ButtonJsonViewer2 = New System.Windows.Forms.Button()
        Me.btnRestore = New System.Windows.Forms.Button()
        Me.btnBackup = New System.Windows.Forms.Button()
        Me.CalRam3458APreRun = New System.Windows.Forms.TextBox()
        Me.BtnSave3458A = New System.Windows.Forms.Button()
        Me.CheckBoxR6581RetrieveREF = New System.Windows.Forms.CheckBox()
        Me.SerialPort1 = New System.IO.Ports.SerialPort(Me.components)
        Me.Timer6 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer7 = New System.Windows.Forms.Timer(Me.components)
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label133 = New System.Windows.Forms.Label()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.ProfDev2_8 = New System.Windows.Forms.CheckBox()
        Me.ProfDev2_7 = New System.Windows.Forms.CheckBox()
        Me.ProfDev2_6 = New System.Windows.Forms.CheckBox()
        Me.ProfDev2_5 = New System.Windows.Forms.CheckBox()
        Me.ProfDev2_4 = New System.Windows.Forms.CheckBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.ProfDev2_3 = New System.Windows.Forms.CheckBox()
        Me.ProfDev2_2 = New System.Windows.Forms.CheckBox()
        Me.ProfDev2_1 = New System.Windows.Forms.CheckBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.IODeviceLabel2 = New System.Windows.Forms.Label()
        Me.lstIntf2 = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtname2 = New System.Windows.Forms.TextBox()
        Me.txtaddr2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.ProfDev1_8 = New System.Windows.Forms.CheckBox()
        Me.ProfDev1_7 = New System.Windows.Forms.CheckBox()
        Me.ProfDev1_6 = New System.Windows.Forms.CheckBox()
        Me.ProfDev1_5 = New System.Windows.Forms.CheckBox()
        Me.ProfDev1_4 = New System.Windows.Forms.CheckBox()
        Me.IODeviceLabel1 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ProfDev1_3 = New System.Windows.Forms.CheckBox()
        Me.ProfDev1_2 = New System.Windows.Forms.CheckBox()
        Me.ProfDev1_1 = New System.Windows.Forms.CheckBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.lstIntf1 = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtaddr1 = New System.Windows.Forms.TextBox()
        Me.txtname1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label303 = New System.Windows.Forms.Label()
        Me.Label302 = New System.Windows.Forms.Label()
        Me.RunningTimeLogging = New System.Windows.Forms.Label()
        Me.Label56 = New System.Windows.Forms.Label()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.gbox12 = New System.Windows.Forms.GroupBox()
        Me.Label67 = New System.Windows.Forms.Label()
        Me.Label66 = New System.Windows.Forms.Label()
        Me.Dev12SampleRate = New System.Windows.Forms.TextBox()
        Me.gbox2 = New System.Windows.Forms.GroupBox()
        Me.Dev2INTb = New System.Windows.Forms.Label()
        Me.Label221 = New System.Windows.Forms.Label()
        Me.Label299 = New System.Windows.Forms.Label()
        Me.Label293 = New System.Windows.Forms.Label()
        Me.Label300 = New System.Windows.Forms.Label()
        Me.Dev2delayop = New System.Windows.Forms.TextBox()
        Me.Label297 = New System.Windows.Forms.Label()
        Me.Dev2Timeout = New System.Windows.Forms.TextBox()
        Me.Dev2INT = New System.Windows.Forms.Label()
        Me.Dev2K2001isolatedataCHAR = New System.Windows.Forms.TextBox()
        Me.Label191 = New System.Windows.Forms.Label()
        Me.Label101 = New System.Windows.Forms.Label()
        Me.Dev2PollingEnable = New System.Windows.Forms.CheckBox()
        Me.txtr2a_disp = New System.Windows.Forms.TextBox()
        Me.Dev2STBMask = New System.Windows.Forms.TextBox()
        Me.txtq2c = New System.Windows.Forms.TextBox()
        Me.txtq2b = New System.Windows.Forms.TextBox()
        Me.txtr2b = New System.Windows.Forms.TextBox()
        Me.txtq2a = New System.Windows.Forms.TextBox()
        Me.txtr2a = New System.Windows.Forms.TextBox()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Dev2Samples = New System.Windows.Forms.Label()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.IgnoreErrors2 = New System.Windows.Forms.CheckBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtr2astat = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.gbox1 = New System.Windows.Forms.GroupBox()
        Me.Dev1INTb = New System.Windows.Forms.Label()
        Me.Label220 = New System.Windows.Forms.Label()
        Me.Label294 = New System.Windows.Forms.Label()
        Me.Label295 = New System.Windows.Forms.Label()
        Me.Label296 = New System.Windows.Forms.Label()
        Me.Label292 = New System.Windows.Forms.Label()
        Me.Dev1INT = New System.Windows.Forms.Label()
        Me.Dev1delayop = New System.Windows.Forms.TextBox()
        Me.Dev1Timeout = New System.Windows.Forms.TextBox()
        Me.Dev1K2001isolatedataCHAR = New System.Windows.Forms.TextBox()
        Me.Label190 = New System.Windows.Forms.Label()
        Me.Dev1STBMask = New System.Windows.Forms.TextBox()
        Me.Label99 = New System.Windows.Forms.Label()
        Me.txtr1a_disp = New System.Windows.Forms.TextBox()
        Me.txtq1b = New System.Windows.Forms.TextBox()
        Me.Dev1PollingEnable = New System.Windows.Forms.CheckBox()
        Me.txtq1c = New System.Windows.Forms.TextBox()
        Me.txtr1a = New System.Windows.Forms.TextBox()
        Me.txtq1a = New System.Windows.Forms.TextBox()
        Me.txtr1b = New System.Windows.Forms.TextBox()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Dev1Samples = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.IgnoreErrors1 = New System.Windows.Forms.CheckBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtr1astat = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TabPage10 = New System.Windows.Forms.TabPage()
        Me.Label227 = New System.Windows.Forms.Label()
        Me.TextBoxDev1CMD = New System.Windows.Forms.TextBox()
        Me.CMD2clear = New System.Windows.Forms.Button()
        Me.CMD1clear = New System.Windows.Forms.Button()
        Me.CMDdev2 = New System.Windows.Forms.Label()
        Me.CMDdev1 = New System.Windows.Forms.Label()
        Me.Label189 = New System.Windows.Forms.Label()
        Me.CheckBoxDev2Async = New System.Windows.Forms.CheckBox()
        Me.CheckBoxDev2Query = New System.Windows.Forms.CheckBox()
        Me.TextBoxDev2CMD = New System.Windows.Forms.TextBox()
        Me.Label188 = New System.Windows.Forms.Label()
        Me.Label187 = New System.Windows.Forms.Label()
        Me.CheckBoxDev1Async = New System.Windows.Forms.CheckBox()
        Me.CheckBoxDev1Query = New System.Windows.Forms.CheckBox()
        Me.TabPage8 = New System.Windows.Forms.TabPage()
        Me.Label169 = New System.Windows.Forms.Label()
        Me.Label170 = New System.Windows.Forms.Label()
        Me.Device2name = New System.Windows.Forms.Label()
        Me.DeviceHumidity = New System.Windows.Forms.Label()
        Me.Label173 = New System.Windows.Forms.Label()
        Me.DeviceTemperature = New System.Windows.Forms.Label()
        Me.Label171 = New System.Windows.Forms.Label()
        Me.Dev2Meter = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Device1name = New System.Windows.Forms.Label()
        Me.Dev1Meter = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.gboxtemphum = New System.Windows.Forms.GroupBox()
        Me.Label231 = New System.Windows.Forms.Label()
        Me.Label230 = New System.Windows.Forms.Label()
        Me.Label198 = New System.Windows.Forms.Label()
        Me.CheckBoxParseLeftRight = New System.Windows.Forms.CheckBox()
        Me.CheckBoxArithmetic = New System.Windows.Forms.CheckBox()
        Me.CheckBoxRegex = New System.Windows.Forms.CheckBox()
        Me.Label219 = New System.Windows.Forms.Label()
        Me.Label218 = New System.Windows.Forms.Label()
        Me.Label217 = New System.Windows.Forms.Label()
        Me.Label216 = New System.Windows.Forms.Label()
        Me.Label215 = New System.Windows.Forms.Label()
        Me.Label214 = New System.Windows.Forms.Label()
        Me.Label213 = New System.Windows.Forms.Label()
        Me.Label212 = New System.Windows.Forms.Label()
        Me.Label211 = New System.Windows.Forms.Label()
        Me.Label210 = New System.Windows.Forms.Label()
        Me.Label209 = New System.Windows.Forms.Label()
        Me.Label208 = New System.Windows.Forms.Label()
        Me.Label207 = New System.Windows.Forms.Label()
        Me.Label206 = New System.Windows.Forms.Label()
        Me.Label205 = New System.Windows.Forms.Label()
        Me.Label204 = New System.Windows.Forms.Label()
        Me.Label195 = New System.Windows.Forms.Label()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.Label202 = New System.Windows.Forms.Label()
        Me.Label203 = New System.Windows.Forms.Label()
        Me.TextBoxTempArithmentic = New System.Windows.Forms.TextBox()
        Me.Label201 = New System.Windows.Forms.Label()
        Me.Label200 = New System.Windows.Forms.Label()
        Me.Label199 = New System.Windows.Forms.Label()
        Me.TempFinalValue = New System.Windows.Forms.Label()
        Me.Label197 = New System.Windows.Forms.Label()
        Me.Label196 = New System.Windows.Forms.Label()
        Me.TextBoxParseRight = New System.Windows.Forms.TextBox()
        Me.TextBoxParseLeft = New System.Windows.Forms.TextBox()
        Me.Label194 = New System.Windows.Forms.Label()
        Me.Label193 = New System.Windows.Forms.Label()
        Me.Label186 = New System.Windows.Forms.Label()
        Me.Label185 = New System.Windows.Forms.Label()
        Me.Label184 = New System.Windows.Forms.Label()
        Me.Label183 = New System.Windows.Forms.Label()
        Me.Label182 = New System.Windows.Forms.Label()
        Me.Label181 = New System.Windows.Forms.Label()
        Me.Label178 = New System.Windows.Forms.Label()
        Me.Label164 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.LabelHumidity = New System.Windows.Forms.Label()
        Me.LabelTemperature = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Label291 = New System.Windows.Forms.Label()
        Me.ListLog = New System.Windows.Forms.ListBox()
        Me.LogFileMetadata = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.bgoxdata = New System.Windows.Forms.GroupBox()
        Me.CSVsize = New System.Windows.Forms.Label()
        Me.CSVcounts = New System.Windows.Forms.Label()
        Me.CSVwrite = New System.Windows.Forms.Label()
        Me.Label228 = New System.Windows.Forms.Label()
        Me.Label226 = New System.Windows.Forms.Label()
        Me.Label177 = New System.Windows.Forms.Label()
        Me.Label176 = New System.Windows.Forms.Label()
        Me.CSVEntryLimitMins = New System.Windows.Forms.TextBox()
        Me.Label175 = New System.Windows.Forms.Label()
        Me.CSVEntryLimit = New System.Windows.Forms.TextBox()
        Me.Label58 = New System.Windows.Forms.Label()
        Me.Label71 = New System.Windows.Forms.Label()
        Me.CSVdelimiterSemiColon = New System.Windows.Forms.RadioButton()
        Me.CSVdelimiterComma = New System.Windows.Forms.RadioButton()
        Me.LabelCSVfilesize = New System.Windows.Forms.Label()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.Label51 = New System.Windows.Forms.Label()
        Me.Label50 = New System.Windows.Forms.Label()
        Me.Label49 = New System.Windows.Forms.Label()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.LabelCSVcounts = New System.Windows.Forms.Label()
        Me.CheckboxEnableLOG = New System.Windows.Forms.CheckBox()
        Me.ENotationDecimal = New System.Windows.Forms.CheckBox()
        Me.CheckboxEnableCSV = New System.Windows.Forms.CheckBox()
        Me.TempHumLogs = New System.Windows.Forms.CheckBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.ListBoxData = New System.Windows.Forms.ListBox()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.Label313 = New System.Windows.Forms.Label()
        Me.Label238 = New System.Windows.Forms.Label()
        Me.Label237 = New System.Windows.Forms.Label()
        Me.Label236 = New System.Windows.Forms.Label()
        Me.Label235 = New System.Windows.Forms.Label()
        Me.StartChartMessage = New System.Windows.Forms.Label()
        Me.LabeChartMinutes = New System.Windows.Forms.Label()
        Me.Label223 = New System.Windows.Forms.Label()
        Me.CheckBoxTempHide = New System.Windows.Forms.CheckBox()
        Me.CheckBoxDevice2Hide = New System.Windows.Forms.CheckBox()
        Me.CheckBoxDevice1Hide = New System.Windows.Forms.CheckBox()
        Me.Device2nameLive = New System.Windows.Forms.Label()
        Me.Device1nameLive = New System.Windows.Forms.Label()
        Me.DisableRollingChart = New System.Windows.Forms.CheckBox()
        Me.LabelChartPoints1 = New System.Windows.Forms.Label()
        Me.Label258 = New System.Windows.Forms.Label()
        Me.LabelChartPoints2 = New System.Windows.Forms.Label()
        Me.Label257 = New System.Windows.Forms.Label()
        Me.YaxisDiff = New System.Windows.Forms.Label()
        Me.Label256 = New System.Windows.Forms.Label()
        Me.Label179 = New System.Windows.Forms.Label()
        Me.Label180 = New System.Windows.Forms.Label()
        Me.LCTempMax = New System.Windows.Forms.TextBox()
        Me.LCTempMin = New System.Windows.Forms.TextBox()
        Me.Dev1Max = New System.Windows.Forms.TextBox()
        Me.Dev1Min = New System.Windows.Forms.TextBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Label72 = New System.Windows.Forms.Label()
        Me.ButtonDiffRecordedTempReset = New System.Windows.Forms.Button()
        Me.TemperatureDiffRecorded = New System.Windows.Forms.Label()
        Me.ButtonDiffRecorded2Reset = New System.Windows.Forms.Button()
        Me.EnableChart1 = New System.Windows.Forms.CheckBox()
        Me.ButtonDiffRecorded1Reset = New System.Windows.Forms.Button()
        Me.EnableChart3 = New System.Windows.Forms.CheckBox()
        Me.inst_value2FDiffRecorded = New System.Windows.Forms.Label()
        Me.inst_value1FDiffRecorded = New System.Windows.Forms.Label()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.EnableChart2 = New System.Windows.Forms.CheckBox()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.TabPage9 = New System.Windows.Forms.TabPage()
        Me.TabPage7 = New System.Windows.Forms.TabPage()
        Me.AddressRangeF = New System.Windows.Forms.RadioButton()
        Me.AddressRangeB = New System.Windows.Forms.RadioButton()
        Me.AddressRangeA = New System.Windows.Forms.RadioButton()
        Me.TextBoxCalRamFile3457A = New System.Windows.Forms.TextBox()
        Me.Label125 = New System.Windows.Forms.Label()
        Me.ButtonCalramDump3457A = New System.Windows.Forms.Button()
        Me.Label122 = New System.Windows.Forms.Label()
        Me.Label123 = New System.Windows.Forms.Label()
        Me.Label116 = New System.Windows.Forms.Label()
        Me.TextBoxCalRamFile = New System.Windows.Forms.TextBox()
        Me.Label62 = New System.Windows.Forms.Label()
        Me.Label63 = New System.Windows.Forms.Label()
        Me.Label60 = New System.Windows.Forms.Label()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.Label318 = New System.Windows.Forms.Label()
        Me.Label315 = New System.Windows.Forms.Label()
        Me.LabelCalRamAddressHex = New System.Windows.Forms.Label()
        Me.TextBoxCalRamFile2 = New System.Windows.Forms.TextBox()
        Me.ButtonCalramDump3458A = New System.Windows.Forms.Button()
        Me.Label249 = New System.Windows.Forms.Label()
        Me.Button3458Aabort = New System.Windows.Forms.Button()
        Me.Label141 = New System.Windows.Forms.Label()
        Me.Label140 = New System.Windows.Forms.Label()
        Me.Label139 = New System.Windows.Forms.Label()
        Me.Label138 = New System.Windows.Forms.Label()
        Me.Label64 = New System.Windows.Forms.Label()
        Me.AddressRangeD = New System.Windows.Forms.RadioButton()
        Me.AddressRangeC = New System.Windows.Forms.RadioButton()
        Me.LabelCounter = New System.Windows.Forms.Label()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.Label103 = New System.Windows.Forms.Label()
        Me.Label68 = New System.Windows.Forms.Label()
        Me.LabelCalRamAddress = New System.Windows.Forms.Label()
        Me.LabelByte = New System.Windows.Forms.Label()
        Me.Label102 = New System.Windows.Forms.Label()
        Me.Label61 = New System.Windows.Forms.Label()
        Me.LabelCalRamByte = New System.Windows.Forms.Label()
        Me.CalramStatus = New System.Windows.Forms.Label()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.LabelCalRamAddress3457AHex = New System.Windows.Forms.Label()
        Me.Label148 = New System.Windows.Forms.Label()
        Me.TextBox3457ATo = New System.Windows.Forms.TextBox()
        Me.Label147 = New System.Windows.Forms.Label()
        Me.Label146 = New System.Windows.Forms.Label()
        Me.TextBox3457AFrom = New System.Windows.Forms.TextBox()
        Me.Label142 = New System.Windows.Forms.Label()
        Me.Button3457Aabort = New System.Windows.Forms.Button()
        Me.Label143 = New System.Windows.Forms.Label()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.Label144 = New System.Windows.Forms.Label()
        Me.Label127 = New System.Windows.Forms.Label()
        Me.Label131 = New System.Windows.Forms.Label()
        Me.LabelCounter3457A = New System.Windows.Forms.Label()
        Me.Label129 = New System.Windows.Forms.Label()
        Me.Label126 = New System.Windows.Forms.Label()
        Me.Label130 = New System.Windows.Forms.Label()
        Me.LabelCalRamAddress3457A = New System.Windows.Forms.Label()
        Me.Label132 = New System.Windows.Forms.Label()
        Me.Label128 = New System.Windows.Forms.Label()
        Me.CalramStatus3457A = New System.Windows.Forms.Label()
        Me.LabelCalRamByte3457A = New System.Windows.Forms.Label()
        Me.TabPage11 = New System.Windows.Forms.TabPage()
        Me.GroupBox10 = New System.Windows.Forms.GroupBox()
        Me.TextBoxR6581GPIBlist = New System.Windows.Forms.TextBox()
        Me.Label111 = New System.Windows.Forms.Label()
        Me.CheckBoxR6581Upload9 = New System.Windows.Forms.CheckBox()
        Me.CheckBoxR6581Upload8 = New System.Windows.Forms.CheckBox()
        Me.Label305 = New System.Windows.Forms.Label()
        Me.Label248 = New System.Windows.Forms.Label()
        Me.Label301 = New System.Windows.Forms.Label()
        Me.LabelCalRamByte6581upload = New System.Windows.Forms.Label()
        Me.CalramStatus6581upload = New System.Windows.Forms.Label()
        Me.Label243 = New System.Windows.Forms.Label()
        Me.Label244 = New System.Windows.Forms.Label()
        Me.ButtonR6581commitEEprom = New System.Windows.Forms.Button()
        Me.ButtonR6581upload = New System.Windows.Forms.Button()
        Me.TextBoxCalRamFileJson6581Select = New System.Windows.Forms.TextBox()
        Me.CheckBoxR6581Upload5 = New System.Windows.Forms.CheckBox()
        Me.CheckBoxR6581Upload6 = New System.Windows.Forms.CheckBox()
        Me.CheckBoxR6581Upload7 = New System.Windows.Forms.CheckBox()
        Me.CheckBoxR6581Upload2 = New System.Windows.Forms.CheckBox()
        Me.CheckBoxR6581Upload3 = New System.Windows.Forms.CheckBox()
        Me.CheckBoxR6581Upload4 = New System.Windows.Forms.CheckBox()
        Me.CheckBoxR6581Upload1 = New System.Windows.Forms.CheckBox()
        Me.Label241 = New System.Windows.Forms.Label()
        Me.Label242 = New System.Windows.Forms.Label()
        Me.SendRegularConstantsReadR6581 = New System.Windows.Forms.RadioButton()
        Me.Label59 = New System.Windows.Forms.Label()
        Me.TextBoxCalRamFileJson6581 = New System.Windows.Forms.TextBox()
        Me.TextBoxCalRamFile6581 = New System.Windows.Forms.TextBox()
        Me.Label312 = New System.Windows.Forms.Label()
        Me.Label309 = New System.Windows.Forms.Label()
        Me.Label310 = New System.Windows.Forms.Label()
        Me.Label311 = New System.Windows.Forms.Label()
        Me.ButtonCalramDumpR6581 = New System.Windows.Forms.Button()
        Me.Label245 = New System.Windows.Forms.Label()
        Me.Label246 = New System.Windows.Forms.Label()
        Me.AllRegularConstantsReadR6581 = New System.Windows.Forms.RadioButton()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label298 = New System.Windows.Forms.Label()
        Me.Label304 = New System.Windows.Forms.Label()
        Me.Label306 = New System.Windows.Forms.Label()
        Me.LabelCalRamByte6581 = New System.Windows.Forms.Label()
        Me.CalramStatus6581 = New System.Windows.Forms.Label()
        Me.ButtonR6581abort = New System.Windows.Forms.Button()
        Me.TabPage12 = New System.Windows.Forms.TabPage()
        Me.Label259 = New System.Windows.Forms.Label()
        Me.Label260 = New System.Windows.Forms.Label()
        Me.Label261 = New System.Windows.Forms.Label()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.Label263 = New System.Windows.Forms.Label()
        Me.Timeout3458A = New System.Windows.Forms.TextBox()
        Me.Label273 = New System.Windows.Forms.Label()
        Me.PictureBox8 = New System.Windows.Forms.PictureBox()
        Me.Label272 = New System.Windows.Forms.Label()
        Me.Label267 = New System.Windows.Forms.Label()
        Me.RadioButton3245ADCVDCI = New System.Windows.Forms.RadioButton()
        Me.Label3245AWRI = New System.Windows.Forms.Label()
        Me.Label266 = New System.Windows.Forms.Label()
        Me.Button3245Aabort = New System.Windows.Forms.Button()
        Me.Label264 = New System.Windows.Forms.Label()
        Me.Label262 = New System.Windows.Forms.Label()
        Me.Code3245A = New System.Windows.Forms.TextBox()
        Me.Label3458123 = New System.Windows.Forms.Label()
        Me.Label3458ARDG = New System.Windows.Forms.Label()
        Me.ButtonCal3245A = New System.Windows.Forms.Button()
        Me.Label265 = New System.Windows.Forms.Label()
        Me.Label268 = New System.Windows.Forms.Label()
        Me.RadioButton3245ADCV = New System.Windows.Forms.RadioButton()
        Me.LabelRDG = New System.Windows.Forms.Label()
        Me.Label270 = New System.Windows.Forms.Label()
        Me.Label271 = New System.Windows.Forms.Label()
        Me.Label274 = New System.Windows.Forms.Label()
        Me.Label275 = New System.Windows.Forms.Label()
        Me.Cal3245status = New System.Windows.Forms.Label()
        Me.Label255 = New System.Windows.Forms.Label()
        Me.Label269 = New System.Windows.Forms.Label()
        Me.PictureBox9 = New System.Windows.Forms.PictureBox()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.Label145 = New System.Windows.Forms.Label()
        Me.Label79 = New System.Windows.Forms.Label()
        Me.Label81 = New System.Windows.Forms.Label()
        Me.Label232 = New System.Windows.Forms.Label()
        Me.Label233 = New System.Windows.Forms.Label()
        Me.Label154 = New System.Windows.Forms.Label()
        Me.Label153 = New System.Windows.Forms.Label()
        Me.Label152 = New System.Windows.Forms.Label()
        Me.Label151 = New System.Windows.Forms.Label()
        Me.LabelDeltaV = New System.Windows.Forms.Label()
        Me.volts10 = New System.Windows.Forms.TextBox()
        Me.volts9 = New System.Windows.Forms.TextBox()
        Me.volts8 = New System.Windows.Forms.TextBox()
        Me.volts7 = New System.Windows.Forms.TextBox()
        Me.volts6 = New System.Windows.Forms.TextBox()
        Me.volts5 = New System.Windows.Forms.TextBox()
        Me.volts4 = New System.Windows.Forms.TextBox()
        Me.volts3 = New System.Windows.Forms.TextBox()
        Me.volts2 = New System.Windows.Forms.TextBox()
        Me.volts1 = New System.Windows.Forms.TextBox()
        Me.volts0 = New System.Windows.Forms.TextBox()
        Me.Label93 = New System.Windows.Forms.Label()
        Me.LabeldacSpan9Cal = New System.Windows.Forms.Label()
        Me.DacSpan9 = New System.Windows.Forms.TextBox()
        Me.Label89 = New System.Windows.Forms.Label()
        Me.LabeldacSpan8Cal = New System.Windows.Forms.Label()
        Me.DacSpan8 = New System.Windows.Forms.TextBox()
        Me.Label85 = New System.Windows.Forms.Label()
        Me.LabeldacSpan7Cal = New System.Windows.Forms.Label()
        Me.DacSpan7 = New System.Windows.Forms.TextBox()
        Me.Label86 = New System.Windows.Forms.Label()
        Me.LabeldacSpan6Cal = New System.Windows.Forms.Label()
        Me.DacSpan6 = New System.Windows.Forms.TextBox()
        Me.Label90 = New System.Windows.Forms.Label()
        Me.LabeldacSpan5Cal = New System.Windows.Forms.Label()
        Me.DacSpan5 = New System.Windows.Forms.TextBox()
        Me.Label94 = New System.Windows.Forms.Label()
        Me.LabeldacSpan4Cal = New System.Windows.Forms.Label()
        Me.DacSpan4 = New System.Windows.Forms.TextBox()
        Me.Label96 = New System.Windows.Forms.Label()
        Me.LabeldacSpan3Cal = New System.Windows.Forms.Label()
        Me.DacSpan3 = New System.Windows.Forms.TextBox()
        Me.Label98 = New System.Windows.Forms.Label()
        Me.LabeldacSpan2Cal = New System.Windows.Forms.Label()
        Me.DacSpan2 = New System.Windows.Forms.TextBox()
        Me.Label100 = New System.Windows.Forms.Label()
        Me.LabeldacSpan1Cal = New System.Windows.Forms.Label()
        Me.DacSpan1 = New System.Windows.Forms.TextBox()
        Me.LabelBatteryMonICMult = New System.Windows.Forms.Label()
        Me.Label106 = New System.Windows.Forms.Label()
        Me.ChargeI = New System.Windows.Forms.TextBox()
        Me.Label107 = New System.Windows.Forms.Label()
        Me.Label108 = New System.Windows.Forms.Label()
        Me.Label109 = New System.Windows.Forms.Label()
        Me.Label110 = New System.Windows.Forms.Label()
        Me.LabelBatteryVFeedMult = New System.Windows.Forms.Label()
        Me.LabelOutputVFeedMult = New System.Windows.Forms.Label()
        Me.LabeldacSpan0Cal = New System.Windows.Forms.Label()
        Me.LabeldacZero0Cal = New System.Windows.Forms.Label()
        Me.BattVdc = New System.Windows.Forms.TextBox()
        Me.OutVdc = New System.Windows.Forms.TextBox()
        Me.DacSpan0 = New System.Windows.Forms.TextBox()
        Me.DacZero0 = New System.Windows.Forms.TextBox()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.TextBoxdegC = New System.Windows.Forms.TextBox()
        Me.Label97 = New System.Windows.Forms.Label()
        Me.TextBoxSer = New System.Windows.Forms.TextBox()
        Me.Label95 = New System.Windows.Forms.Label()
        Me.TextBoxSOAK = New System.Windows.Forms.TextBox()
        Me.TextBoxSERIAL = New System.Windows.Forms.TextBox()
        Me.TextBoxFULLMA = New System.Windows.Forms.TextBox()
        Me.TextBoxOLMA = New System.Windows.Forms.TextBox()
        Me.TextBoxCENABLE = New System.Windows.Forms.TextBox()
        Me.TextBoxLOWSHUT = New System.Windows.Forms.TextBox()
        Me.TextBoxDC = New System.Windows.Forms.TextBox()
        Me.TextBoxCD = New System.Windows.Forms.TextBox()
        Me.Label92 = New System.Windows.Forms.Label()
        Me.Label91 = New System.Windows.Forms.Label()
        Me.Label88 = New System.Windows.Forms.Label()
        Me.Label87 = New System.Windows.Forms.Label()
        Me.Label84 = New System.Windows.Forms.Label()
        Me.Label82 = New System.Windows.Forms.Label()
        Me.Label78 = New System.Windows.Forms.Label()
        Me.Label77 = New System.Windows.Forms.Label()
        Me.Label114 = New System.Windows.Forms.Label()
        Me.Label113 = New System.Windows.Forms.Label()
        Me.Label112 = New System.Windows.Forms.Label()
        Me.txtr1aBIG = New System.Windows.Forms.Label()
        Me.comPort_ComboBox = New System.Windows.Forms.ComboBox()
        Me.connect_BTN = New System.Windows.Forms.Button()
        Me.ButtonKeyVoltage = New System.Windows.Forms.Button()
        Me.KeyVoltage = New System.Windows.Forms.TextBox()
        Me.Label80 = New System.Windows.Forms.Label()
        Me.LabelKeyVoltage = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label174 = New System.Windows.Forms.Label()
        Me.Default11 = New System.Windows.Forms.TextBox()
        Me.Label172 = New System.Windows.Forms.Label()
        Me.Default10 = New System.Windows.Forms.TextBox()
        Me.Default9 = New System.Windows.Forms.TextBox()
        Me.Default8 = New System.Windows.Forms.TextBox()
        Me.Default7 = New System.Windows.Forms.TextBox()
        Me.Default6 = New System.Windows.Forms.TextBox()
        Me.Default5 = New System.Windows.Forms.TextBox()
        Me.Default4 = New System.Windows.Forms.TextBox()
        Me.Default3 = New System.Windows.Forms.TextBox()
        Me.Default2 = New System.Windows.Forms.TextBox()
        Me.Default1 = New System.Windows.Forms.TextBox()
        Me.Default0 = New System.Windows.Forms.TextBox()
        Me.Label168 = New System.Windows.Forms.Label()
        Me.Label158 = New System.Windows.Forms.Label()
        Me.Label157 = New System.Windows.Forms.Label()
        Me.Label156 = New System.Windows.Forms.Label()
        Me.LabelTemperature2 = New System.Windows.Forms.Label()
        Me.Label155 = New System.Windows.Forms.Label()
        Me.Label229 = New System.Windows.Forms.Label()
        Me.LabelTemperature3 = New System.Windows.Forms.TextBox()
        Me.LabeldacSpan10Delta = New System.Windows.Forms.Label()
        Me.Label149 = New System.Windows.Forms.Label()
        Me.Label150 = New System.Windows.Forms.Label()
        Me.ButtonDacSpan10Up = New System.Windows.Forms.Button()
        Me.ButtonDacSpan10 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan10down = New System.Windows.Forms.Button()
        Me.LabeldacSpan10Cal = New System.Windows.Forms.Label()
        Me.volts11 = New System.Windows.Forms.TextBox()
        Me.DacSpan10 = New System.Windows.Forms.TextBox()
        Me.ButtonDacSpan9Up = New System.Windows.Forms.Button()
        Me.ButtonDacSpan9 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan9down = New System.Windows.Forms.Button()
        Me.ButtonDacSpan8 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan8Up = New System.Windows.Forms.Button()
        Me.ButtonDacSpan7 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan8down = New System.Windows.Forms.Button()
        Me.ButtonDacSpan6 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan7Up = New System.Windows.Forms.Button()
        Me.ButtonDacSpan5 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan7down = New System.Windows.Forms.Button()
        Me.ButtonDacSpan4 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan6Up = New System.Windows.Forms.Button()
        Me.ButtonDacSpan6down = New System.Windows.Forms.Button()
        Me.ButtonDacSpan3 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan5Up = New System.Windows.Forms.Button()
        Me.ButtonDacSpan2 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan5down = New System.Windows.Forms.Button()
        Me.ButtonDacSpan1 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan4Up = New System.Windows.Forms.Button()
        Me.ButtonOutVdc = New System.Windows.Forms.Button()
        Me.ButtonDacSpan4down = New System.Windows.Forms.Button()
        Me.ButtonDacSpan0 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan3Up = New System.Windows.Forms.Button()
        Me.ButtonDacZero0 = New System.Windows.Forms.Button()
        Me.ButtonDacSpan3down = New System.Windows.Forms.Button()
        Me.ButtonBattVdc = New System.Windows.Forms.Button()
        Me.ButtonDacSpan2Up = New System.Windows.Forms.Button()
        Me.ButtonDacSpan2down = New System.Windows.Forms.Button()
        Me.Label162 = New System.Windows.Forms.Label()
        Me.ButtonDacSpan1Up = New System.Windows.Forms.Button()
        Me.ButtonDacSpan1down = New System.Windows.Forms.Button()
        Me.ButtonChargeIUp = New System.Windows.Forms.Button()
        Me.ButtonBattVdcUp = New System.Windows.Forms.Button()
        Me.Label225 = New System.Windows.Forms.Label()
        Me.ButtonOutVdcUp = New System.Windows.Forms.Button()
        Me.Label224 = New System.Windows.Forms.Label()
        Me.ButtonDacSpan0Up = New System.Windows.Forms.Button()
        Me.ButtonDacZero0Up = New System.Windows.Forms.Button()
        Me.TextBoxUser = New System.Windows.Forms.TextBox()
        Me.ButtonChargeIdown = New System.Windows.Forms.Button()
        Me.Label222 = New System.Windows.Forms.Label()
        Me.ButtonBattVdcdown = New System.Windows.Forms.Button()
        Me.TextBox3458Asn = New System.Windows.Forms.TextBox()
        Me.ButtonOutVdcdown = New System.Windows.Forms.Button()
        Me.Label73 = New System.Windows.Forms.Label()
        Me.ButtonDacSpan0down = New System.Windows.Forms.Button()
        Me.Label290 = New System.Windows.Forms.Label()
        Me.ButtonDacZero0down = New System.Windows.Forms.Button()
        Me.Label289 = New System.Windows.Forms.Label()
        Me.Label288 = New System.Windows.Forms.Label()
        Me.Label287 = New System.Windows.Forms.Label()
        Me.Label286 = New System.Windows.Forms.Label()
        Me.Label285 = New System.Windows.Forms.Label()
        Me.Label284 = New System.Windows.Forms.Label()
        Me.Label283 = New System.Windows.Forms.Label()
        Me.Label282 = New System.Windows.Forms.Label()
        Me.Label281 = New System.Windows.Forms.Label()
        Me.Label280 = New System.Windows.Forms.Label()
        Me.Label279 = New System.Windows.Forms.Label()
        Me.Label278 = New System.Windows.Forms.Label()
        Me.Label277 = New System.Windows.Forms.Label()
        Me.Label161 = New System.Windows.Forms.Label()
        Me.Label160 = New System.Windows.Forms.Label()
        Me.Label159 = New System.Windows.Forms.Label()
        Me.Label137 = New System.Windows.Forms.Label()
        Me.Label136 = New System.Windows.Forms.Label()
        Me.Label124 = New System.Windows.Forms.Label()
        Me.Label104 = New System.Windows.Forms.Label()
        Me.Label55 = New System.Windows.Forms.Label()
        Me.DACcalparamVAL = New System.Windows.Forms.Label()
        Me.RadioPDF = New System.Windows.Forms.RadioButton()
        Me.PDVS2counter = New System.Windows.Forms.Label()
        Me.RadioMSWord = New System.Windows.Forms.RadioButton()
        Me.Label254 = New System.Windows.Forms.Label()
        Me.Label163 = New System.Windows.Forms.Label()
        Me.Label167 = New System.Windows.Forms.Label()
        Me.Label166 = New System.Windows.Forms.Label()
        Me.Label165 = New System.Windows.Forms.Label()
        Me.TextBoxLabelBatteryCharge = New System.Windows.Forms.TextBox()
        Me.TextBoxLabelBatteryLowInd = New System.Windows.Forms.TextBox()
        Me.TextBoxLabelMode = New System.Windows.Forms.TextBox()
        Me.TextBoxLabelBatteryI = New System.Windows.Forms.TextBox()
        Me.TextBoxLabelBatteryV = New System.Windows.Forms.TextBox()
        Me.TextBoxLabelOutputVFeedback = New System.Windows.Forms.TextBox()
        Me.Label118 = New System.Windows.Forms.Label()
        Me.ButtonSetXYcalvars = New System.Windows.Forms.Button()
        Me.LabeldacSpan9Delta = New System.Windows.Forms.Label()
        Me.LabeldacSpan8Delta = New System.Windows.Forms.Label()
        Me.LabeldacSpan7Delta = New System.Windows.Forms.Label()
        Me.LabeldacSpan6Delta = New System.Windows.Forms.Label()
        Me.LabeldacSpan5Delta = New System.Windows.Forms.Label()
        Me.LabeldacSpan4Delta = New System.Windows.Forms.Label()
        Me.LabeldacSpan3Delta = New System.Windows.Forms.Label()
        Me.LabeldacSpan2Delta = New System.Windows.Forms.Label()
        Me.LabeldacSpan1Delta = New System.Windows.Forms.Label()
        Me.LabeldacSpan0Delta = New System.Windows.Forms.Label()
        Me.LabeldacZero0Delta = New System.Windows.Forms.Label()
        Me.volts030000 = New System.Windows.Forms.TextBox()
        Me.Button030000 = New System.Windows.Forms.Button()
        Me.volts050000 = New System.Windows.Forms.TextBox()
        Me.volts020000 = New System.Windows.Forms.TextBox()
        Me.volts010000 = New System.Windows.Forms.TextBox()
        Me.volts001000 = New System.Windows.Forms.TextBox()
        Me.volts000100 = New System.Windows.Forms.TextBox()
        Me.Button050000 = New System.Windows.Forms.Button()
        Me.Button020000 = New System.Windows.Forms.Button()
        Me.Button010000 = New System.Windows.Forms.Button()
        Me.Button001000 = New System.Windows.Forms.Button()
        Me.Button000010 = New System.Windows.Forms.Button()
        Me.volts000010 = New System.Windows.Forms.TextBox()
        Me.Button000100 = New System.Windows.Forms.Button()
        Me.NPLC = New System.Windows.Forms.Label()
        Me.Label83 = New System.Windows.Forms.Label()
        Me.Label76 = New System.Windows.Forms.Label()
        Me.Label75 = New System.Windows.Forms.Label()
        Me.Label74 = New System.Windows.Forms.Label()
        Me.LabelCalCount = New System.Windows.Forms.Label()
        Me.ButtonChargeI = New System.Windows.Forms.Button()
        Me.PDVS2delay = New System.Windows.Forms.TextBox()
        Me.CalStep = New System.Windows.Forms.TextBox()
        Me.Label115 = New System.Windows.Forms.Label()
        Me.CalAccuracy = New System.Windows.Forms.TextBox()
        Me.Label119 = New System.Windows.Forms.Label()
        Me.CalStepFinal = New System.Windows.Forms.TextBox()
        Me.Label120 = New System.Windows.Forms.Label()
        Me.CalAccuracyFinal = New System.Windows.Forms.TextBox()
        Me.Label117 = New System.Windows.Forms.Label()
        Me.Label121 = New System.Windows.Forms.Label()
        Me.TabPage13 = New System.Windows.Forms.TabPage()
        Me.GroupBox11 = New System.Windows.Forms.GroupBox()
        Me.Label314 = New System.Windows.Forms.Label()
        Me.Label307 = New System.Windows.Forms.Label()
        Me.CheckBoxEnableTooltips = New System.Windows.Forms.CheckBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label308 = New System.Windows.Forms.Label()
        Me.Label247 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TextBoxTextEditor = New System.Windows.Forms.TextBox()
        Me.CheckBoxAllowSaveAnytime = New System.Windows.Forms.CheckBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
        Me.Label316 = New System.Windows.Forms.Label()
        Me.Label240 = New System.Windows.Forms.Label()
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label239 = New System.Windows.Forms.Label()
        Me.Label234 = New System.Windows.Forms.Label()
        Me.URL3 = New System.Windows.Forms.Label()
        Me.Label192 = New System.Windows.Forms.Label()
        Me.URL2 = New System.Windows.Forms.Label()
        Me.Label134 = New System.Windows.Forms.Label()
        Me.PictureBox6 = New System.Windows.Forms.PictureBox()
        Me.Label135 = New System.Windows.Forms.Label()
        Me.URL1 = New System.Windows.Forms.Label()
        Me.Label105 = New System.Windows.Forms.Label()
        Me.Label70 = New System.Windows.Forms.Label()
        Me.Label69 = New System.Windows.Forms.Label()
        Me.Donate1 = New System.Windows.Forms.Label()
        Me.Timer8 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer9 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer10 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer11 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer12 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer13 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer14 = New System.Windows.Forms.Timer(Me.components)
        Me.OnOffLed2 = New WinGPIBproj.OnOffLed()
        Me.OnOffLed1 = New WinGPIBproj.OnOffLed()
        Me.OnOffLed4 = New WinGPIBproj.OnOffLed()
        Me.OnOffLed3 = New WinGPIBproj.OnOffLed()
        Me.TabControl1.SuspendLayout
        Me.TabPage1.SuspendLayout
        Me.Panel2.SuspendLayout
        Me.Panel1.SuspendLayout
        Me.GroupBox9.SuspendLayout
        Me.GroupBox8.SuspendLayout
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit
        Me.gbox12.SuspendLayout
        Me.gbox2.SuspendLayout
        Me.gbox1.SuspendLayout
        Me.TabPage10.SuspendLayout
        Me.TabPage8.SuspendLayout
        Me.GroupBox4.SuspendLayout
        Me.GroupBox3.SuspendLayout
        Me.TabPage2.SuspendLayout
        Me.gboxtemphum.SuspendLayout
        Me.TabPage3.SuspendLayout
        Me.bgoxdata.SuspendLayout
        Me.TabPage4.SuspendLayout
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit
        Me.TabPage7.SuspendLayout
        Me.GroupBox6.SuspendLayout
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit
        Me.GroupBox7.SuspendLayout
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit
        Me.TabPage11.SuspendLayout
        Me.GroupBox10.SuspendLayout
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit
        Me.TabPage12.SuspendLayout
        Me.GroupBox5.SuspendLayout
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.PictureBox9, System.ComponentModel.ISupportInitialize).BeginInit
        Me.TabPage5.SuspendLayout
        Me.GroupBox2.SuspendLayout
        Me.TabPage13.SuspendLayout
        Me.GroupBox11.SuspendLayout
        Me.TabPage6.SuspendLayout
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit
        Me.SuspendLayout
        '
        'SerialPort
        '
        Me.SerialPort.BaudRate = 300
        Me.SerialPort.DtrEnable = True
        Me.SerialPort.ReadTimeout = 500
        Me.SerialPort.WriteTimeout = 500
        '
        'Timer1
        '
        Me.Timer1.Interval = 5000
        '
        'Timer2
        '
        Me.Timer2.Interval = 5000
        '
        'Timer3
        '
        Me.Timer3.Interval = 5000
        '
        'Timer4
        '
        Me.Timer4.Interval = 5000
        '
        'btndevlist
        '
        Me.btndevlist.BackColor = System.Drawing.Color.WhiteSmoke
        Me.btndevlist.Location = New System.Drawing.Point(951, 134)
        Me.btndevlist.Name = "btndevlist"
        Me.btndevlist.Size = New System.Drawing.Size(90, 40)
        Me.btndevlist.TabIndex = 70
        Me.btndevlist.Text = "Show I/O" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Devices List"
        Me.ToolTip1.SetToolTip(Me.btndevlist, "Show the IO Devices pop-up")
        Me.btndevlist.UseVisualStyleBackColor = True
        '
        'ButtonReset
        '
        Me.ButtonReset.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ButtonReset.Location = New System.Drawing.Point(0, 0)
        Me.ButtonReset.Name = "ButtonReset"
        Me.ButtonReset.Size = New System.Drawing.Size(90, 40)
        Me.ButtonReset.TabIndex = 75
        Me.ButtonReset.Text = "Disconnect I/O" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Devices"
        Me.ToolTip1.SetToolTip(Me.ButtonReset, "Disconnect the currently connected devices")
        Me.ButtonReset.UseVisualStyleBackColor = True
        '
        'ButtonDev12Run
        '
        Me.ButtonDev12Run.Location = New System.Drawing.Point(95, 9)
        Me.ButtonDev12Run.Name = "ButtonDev12Run"
        Me.ButtonDev12Run.Size = New System.Drawing.Size(115, 35)
        Me.ButtonDev12Run.TabIndex = 76
        Me.ButtonDev12Run.Text = "Run"
        Me.ToolTip1.SetToolTip(Me.ButtonDev12Run, "START/STOP logging dual device data")
        Me.ButtonDev12Run.UseVisualStyleBackColor = True
        '
        'Dev2removeletters
        '
        Me.Dev2removeletters.AutoSize = True
        Me.Dev2removeletters.Location = New System.Drawing.Point(119, 26)
        Me.Dev2removeletters.Name = "Dev2removeletters"
        Me.Dev2removeletters.Size = New System.Drawing.Size(215, 17)
        Me.Dev2removeletters.TabIndex = 93
        Me.Dev2removeletters.Text = "Isolate data to right of leading two Chars"
        Me.ToolTip1.SetToolTip(Me.Dev2removeletters, "Remove first two letters from start of returned data leaving the numerical data")
        Me.Dev2removeletters.UseVisualStyleBackColor = True
        '
        'CheckBoxSendBlockingDev2
        '
        Me.CheckBoxSendBlockingDev2.AutoSize = True
        Me.CheckBoxSendBlockingDev2.Location = New System.Drawing.Point(8, 49)
        Me.CheckBoxSendBlockingDev2.Name = "CheckBoxSendBlockingDev2"
        Me.CheckBoxSendBlockingDev2.Size = New System.Drawing.Size(95, 17)
        Me.CheckBoxSendBlockingDev2.TabIndex = 77
        Me.CheckBoxSendBlockingDev2.Text = "Send Blocking"
        Me.ToolTip1.SetToolTip(Me.CheckBoxSendBlockingDev2, "Alternative method of sending" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "AT RUN to Device 2")
        Me.CheckBoxSendBlockingDev2.UseVisualStyleBackColor = True
        '
        'btnq2b
        '
        Me.btnq2b.Location = New System.Drawing.Point(262, 109)
        Me.btnq2b.Name = "btnq2b"
        Me.btnq2b.Size = New System.Drawing.Size(199, 40)
        Me.btnq2b.TabIndex = 17
        Me.btnq2b.Text = "Query Blocking"
        Me.ToolTip1.SetToolTip(Me.btnq2b, "Immediately executed and waits until it" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "gets response from the device")
        Me.btnq2b.UseVisualStyleBackColor = True
        '
        'btns2c
        '
        Me.btns2c.Location = New System.Drawing.Point(262, 194)
        Me.btns2c.Name = "btns2c"
        Me.btns2c.Size = New System.Drawing.Size(199, 22)
        Me.btns2c.TabIndex = 70
        Me.btns2c.Text = "Send Async"
        Me.ToolTip1.SetToolTip(Me.btns2c, "These are queued, i.e. appended to the current" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "queue but expect no reply from th" &
        "e device")
        Me.btns2c.UseVisualStyleBackColor = True
        '
        'btnq2a
        '
        Me.btnq2a.Location = New System.Drawing.Point(262, 151)
        Me.btnq2a.Name = "btnq2a"
        Me.btnq2a.Size = New System.Drawing.Size(199, 40)
        Me.btnq2a.TabIndex = 18
        Me.btnq2a.Text = "Query Async"
        Me.ToolTip1.SetToolTip(Me.btnq2a, "These are queued, i.e. appended to the current" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "queue and when actioned return a " &
        "response" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "immediately")
        Me.btnq2a.UseVisualStyleBackColor = True
        '
        'CommandStart2run
        '
        Me.CommandStart2run.Location = New System.Drawing.Point(262, 353)
        Me.CommandStart2run.Name = "CommandStart2run"
        Me.CommandStart2run.Size = New System.Drawing.Size(198, 20)
        Me.CommandStart2run.TabIndex = 67
        Me.CommandStart2run.Text = ":READ?"
        Me.ToolTip1.SetToolTip(Me.CommandStart2run, "A single command sent automatically after PRE-RUN" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "and repeated at the set Sample" &
        " Rate (seconds)." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "This command is sent as Async Query type command" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "i.e. expects" &
        " a reply from the device.")
        '
        'CommandStop2
        '
        Me.CommandStop2.Location = New System.Drawing.Point(262, 390)
        Me.CommandStop2.Multiline = True
        Me.CommandStop2.Name = "CommandStop2"
        Me.CommandStop2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.CommandStop2.Size = New System.Drawing.Size(198, 35)
        Me.CommandStop2.TabIndex = 52
        Me.ToolTip1.SetToolTip(Me.CommandStop2, "These batch commands are sent when STOP" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "is pressed and are Async Query type comm" &
        "ands" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "i.e. expect a reply from the device.")
        '
        'CommandStart2
        '
        Me.CommandStart2.Location = New System.Drawing.Point(262, 240)
        Me.CommandStart2.Multiline = True
        Me.CommandStart2.Name = "CommandStart2"
        Me.CommandStart2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.CommandStart2.Size = New System.Drawing.Size(198, 96)
        Me.CommandStart2.TabIndex = 52
        Me.CommandStart2.Text = "FUNC 'VOLT:DC'" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & ":INIT:CONT OFF"
        Me.ToolTip1.SetToolTip(Me.CommandStart2, "These batch commands are sent when RUN is" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "pressed and are Send Asynchronous type" &
        " commands" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "i.e. expect no reply.")
        '
        'ButtonDev2Run
        '
        Me.ButtonDev2Run.Location = New System.Drawing.Point(10, 391)
        Me.ButtonDev2Run.Name = "ButtonDev2Run"
        Me.ButtonDev2Run.Size = New System.Drawing.Size(115, 35)
        Me.ButtonDev2Run.TabIndex = 17
        Me.ButtonDev2Run.Text = "Run"
        Me.ToolTip1.SetToolTip(Me.ButtonDev2Run, "START/STOP logging Device 2 data")
        Me.ButtonDev2Run.UseVisualStyleBackColor = True
        '
        'Dev1removeletters
        '
        Me.Dev1removeletters.AutoSize = True
        Me.Dev1removeletters.Location = New System.Drawing.Point(121, 26)
        Me.Dev1removeletters.Name = "Dev1removeletters"
        Me.Dev1removeletters.Size = New System.Drawing.Size(215, 17)
        Me.Dev1removeletters.TabIndex = 92
        Me.Dev1removeletters.Text = "Isolate data to right of leading two Chars"
        Me.ToolTip1.SetToolTip(Me.Dev1removeletters, "Remove first two letters from start of returned data leaving the numerical data")
        Me.Dev1removeletters.UseVisualStyleBackColor = True
        '
        'CheckBoxSendBlockingDev1
        '
        Me.CheckBoxSendBlockingDev1.AutoSize = True
        Me.CheckBoxSendBlockingDev1.Location = New System.Drawing.Point(8, 49)
        Me.CheckBoxSendBlockingDev1.Name = "CheckBoxSendBlockingDev1"
        Me.CheckBoxSendBlockingDev1.Size = New System.Drawing.Size(95, 17)
        Me.CheckBoxSendBlockingDev1.TabIndex = 76
        Me.CheckBoxSendBlockingDev1.Text = "Send Blocking"
        Me.ToolTip1.SetToolTip(Me.CheckBoxSendBlockingDev1, "Alternative method of sending" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "AT RUN to Device 1")
        Me.CheckBoxSendBlockingDev1.UseVisualStyleBackColor = True
        '
        'btnq1b
        '
        Me.btnq1b.Location = New System.Drawing.Point(263, 109)
        Me.btnq1b.Name = "btnq1b"
        Me.btnq1b.Size = New System.Drawing.Size(199, 40)
        Me.btnq1b.TabIndex = 7
        Me.btnq1b.Text = "Query Blocking"
        Me.ToolTip1.SetToolTip(Me.btnq1b, "Immediately executed and waits until it" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "gets response from the device")
        Me.btnq1b.UseVisualStyleBackColor = True
        '
        'btns1c
        '
        Me.btns1c.Location = New System.Drawing.Point(263, 194)
        Me.btns1c.Name = "btns1c"
        Me.btns1c.Size = New System.Drawing.Size(199, 22)
        Me.btns1c.TabIndex = 69
        Me.btns1c.Text = "Send Async"
        Me.ToolTip1.SetToolTip(Me.btns1c, "These are queued, i.e. appended to the current" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "queue but expect no reply from th" &
        "e device")
        Me.btns1c.UseVisualStyleBackColor = True
        '
        'btnq1a
        '
        Me.btnq1a.Location = New System.Drawing.Point(263, 151)
        Me.btnq1a.Name = "btnq1a"
        Me.btnq1a.Size = New System.Drawing.Size(199, 40)
        Me.btnq1a.TabIndex = 8
        Me.btnq1a.Text = "Query Async"
        Me.ToolTip1.SetToolTip(Me.btnq1a, "These are queued, i.e. appended to the current" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "queue and when actioned return a " &
        "response" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "immediately")
        Me.btnq1a.UseVisualStyleBackColor = True
        '
        'CommandStart1run
        '
        Me.CommandStart1run.Location = New System.Drawing.Point(264, 353)
        Me.CommandStart1run.Name = "CommandStart1run"
        Me.CommandStart1run.Size = New System.Drawing.Size(197, 20)
        Me.CommandStart1run.TabIndex = 66
        Me.CommandStart1run.Text = "TARM SGL"
        Me.ToolTip1.SetToolTip(Me.CommandStart1run, "A single command sent automatically after PRE-RUN" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "and repeated at the set Sample" &
        " Rate (seconds)." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "This command is sent as Async Query type command" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "i.e. expects" &
        " a reply from the device." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        '
        'CommandStop1
        '
        Me.CommandStop1.Location = New System.Drawing.Point(264, 390)
        Me.CommandStop1.Multiline = True
        Me.CommandStop1.Name = "CommandStop1"
        Me.CommandStop1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.CommandStop1.Size = New System.Drawing.Size(197, 35)
        Me.CommandStop1.TabIndex = 49
        Me.CommandStop1.Text = "NRDGS 1,AUTO"
        Me.ToolTip1.SetToolTip(Me.CommandStop1, "These batch commands are sent when STOP" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "is pressed and are Async Query type comm" &
        "ands" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "i.e. expect a reply from the device.")
        '
        'CommandStart1
        '
        Me.CommandStart1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CommandStart1.Location = New System.Drawing.Point(264, 240)
        Me.CommandStart1.Multiline = True
        Me.CommandStart1.Name = "CommandStart1"
        Me.CommandStart1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.CommandStart1.Size = New System.Drawing.Size(197, 96)
        Me.CommandStart1.TabIndex = 48
        Me.CommandStart1.Text = "END ALWAYS" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "NPLC 100" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "NRDGS 1"
        Me.ToolTip1.SetToolTip(Me.CommandStart1, "These batch commands are sent when RUN is" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "pressed and are Send Asynchronous type" &
        " commands" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "i.e. expect no reply.")
        '
        'ButtonDev1Run
        '
        Me.ButtonDev1Run.Location = New System.Drawing.Point(8, 391)
        Me.ButtonDev1Run.Name = "ButtonDev1Run"
        Me.ButtonDev1Run.Size = New System.Drawing.Size(115, 35)
        Me.ButtonDev1Run.TabIndex = 16
        Me.ButtonDev1Run.Text = "Run"
        Me.ToolTip1.SetToolTip(Me.ButtonDev1Run, "START/STOP logging Device 1 data")
        Me.ButtonDev1Run.UseVisualStyleBackColor = True
        '
        'lstIntf3
        '
        Me.lstIntf3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstIntf3.FormattingEnabled = True
        Me.lstIntf3.Location = New System.Drawing.Point(96, 68)
        Me.lstIntf3.Name = "lstIntf3"
        Me.lstIntf3.Size = New System.Drawing.Size(206, 21)
        Me.lstIntf3.TabIndex = 24
        Me.ToolTip1.SetToolTip(Me.lstIntf3, "Select a Temp/Hum probe type")
        '
        'ButtonStart
        '
        Me.ButtonStart.Location = New System.Drawing.Point(19, 239)
        Me.ButtonStart.Name = "ButtonStart"
        Me.ButtonStart.Size = New System.Drawing.Size(115, 35)
        Me.ButtonStart.TabIndex = 27
        Me.ButtonStart.Text = "Start"
        Me.ToolTip1.SetToolTip(Me.ButtonStart, "Connect to the probe and" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "start receiving Temp/Hum data")
        Me.ButtonStart.UseVisualStyleBackColor = True
        '
        'ButtonEnd
        '
        Me.ButtonEnd.Enabled = False
        Me.ButtonEnd.Location = New System.Drawing.Point(162, 239)
        Me.ButtonEnd.Name = "ButtonEnd"
        Me.ButtonEnd.Size = New System.Drawing.Size(115, 35)
        Me.ButtonEnd.TabIndex = 28
        Me.ButtonEnd.Text = "Stop"
        Me.ToolTip1.SetToolTip(Me.ButtonEnd, "Disconnect from the Temp/Hum probe")
        Me.ButtonEnd.UseVisualStyleBackColor = True
        '
        'ComboBoxPort
        '
        Me.ComboBoxPort.DisplayMember = "gPortList"
        Me.ComboBoxPort.FormattingEnabled = True
        Me.ComboBoxPort.Location = New System.Drawing.Point(96, 96)
        Me.ComboBoxPort.Name = "ComboBoxPort"
        Me.ComboBoxPort.Size = New System.Drawing.Size(90, 21)
        Me.ComboBoxPort.TabIndex = 25
        Me.ToolTip1.SetToolTip(Me.ComboBoxPort, "Serial COM port for the probe")
        Me.ComboBoxPort.ValueMember = "gPortList"
        '
        'ShowFiles
        '
        Me.ShowFiles.BackColor = System.Drawing.Color.Thistle
        Me.ShowFiles.Location = New System.Drawing.Point(508, 539)
        Me.ShowFiles.Name = "ShowFiles"
        Me.ShowFiles.Size = New System.Drawing.Size(127, 37)
        Me.ShowFiles.TabIndex = 55
        Me.ShowFiles.Text = "\WinGPIBdata" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.ToolTip1.SetToolTip(Me.ShowFiles, "Launch Windows File Explorer")
        Me.ShowFiles.UseVisualStyleBackColor = True
        '
        'ResetCSV
        '
        Me.ResetCSV.Location = New System.Drawing.Point(493, 378)
        Me.ResetCSV.Name = "ResetCSV"
        Me.ResetCSV.Size = New System.Drawing.Size(126, 21)
        Me.ResetCSV.TabIndex = 54
        Me.ResetCSV.Text = "Clear CSV file contents"
        Me.ToolTip1.SetToolTip(Me.ResetCSV, "Clear the contents of the " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "current CSV (empty the file)")
        Me.ResetCSV.UseVisualStyleBackColor = True
        '
        'ButtonExportCSV
        '
        Me.ButtonExportCSV.Location = New System.Drawing.Point(472, 462)
        Me.ButtonExportCSV.Name = "ButtonExportCSV"
        Me.ButtonExportCSV.Size = New System.Drawing.Size(127, 37)
        Me.ButtonExportCSV.TabIndex = 53
        Me.ButtonExportCSV.Text = "Export CSV" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "[date prefixed name]"
        Me.ToolTip1.SetToolTip(Me.ButtonExportCSV, "Exports the current CSV to a new file using the date as the filename")
        Me.ButtonExportCSV.UseVisualStyleBackColor = True
        Me.ButtonExportCSV.Visible = False
        '
        'ButtonClearChart
        '
        Me.ButtonClearChart.Location = New System.Drawing.Point(11, 41)
        Me.ButtonClearChart.Name = "ButtonClearChart"
        Me.ButtonClearChart.Size = New System.Drawing.Size(91, 29)
        Me.ButtonClearChart.TabIndex = 89
        Me.ButtonClearChart.Text = "Clear Chart"
        Me.ToolTip1.SetToolTip(Me.ButtonClearChart, "Reset the chart")
        Me.ButtonClearChart.UseVisualStyleBackColor = True
        '
        'ButtonIanWebsite
        '
        Me.ButtonIanWebsite.Location = New System.Drawing.Point(204, 361)
        Me.ButtonIanWebsite.Name = "ButtonIanWebsite"
        Me.ButtonIanWebsite.Size = New System.Drawing.Size(172, 77)
        Me.ButtonIanWebsite.TabIndex = 64
        Me.ButtonIanWebsite.Text = "PayPal.me"
        Me.ToolTip1.SetToolTip(Me.ButtonIanWebsite, "Launch PayPal.me in your web browser and donate if you want to." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        Me.ButtonIanWebsite.UseVisualStyleBackColor = True
        '
        'ButtonSaveSettings
        '
        Me.ButtonSaveSettings.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ButtonSaveSettings.Location = New System.Drawing.Point(951, 415)
        Me.ButtonSaveSettings.Name = "ButtonSaveSettings"
        Me.ButtonSaveSettings.Size = New System.Drawing.Size(90, 37)
        Me.ButtonSaveSettings.TabIndex = 77
        Me.ButtonSaveSettings.Text = "Save All Profiles/Settings"
        Me.ToolTip1.SetToolTip(Me.ButtonSaveSettings, "Save settings for most of the user data")
        Me.ButtonSaveSettings.UseVisualStyleBackColor = True
        '
        'ButtonNotePad2
        '
        Me.ButtonNotePad2.BackColor = System.Drawing.Color.Wheat
        Me.ButtonNotePad2.Location = New System.Drawing.Point(951, 258)
        Me.ButtonNotePad2.Name = "ButtonNotePad2"
        Me.ButtonNotePad2.Size = New System.Drawing.Size(90, 41)
        Me.ButtonNotePad2.TabIndex = 78
        Me.ButtonNotePad2.Text = "Edit .txt" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "GPIBchannels"
        Me.ToolTip1.SetToolTip(Me.ButtonNotePad2, "Open: GPIBchannels.txt text file")
        Me.ButtonNotePad2.UseVisualStyleBackColor = True
        '
        'ShowFilesCalRam
        '
        Me.ShowFilesCalRam.BackColor = System.Drawing.Color.Thistle
        Me.ShowFilesCalRam.Location = New System.Drawing.Point(928, 16)
        Me.ShowFilesCalRam.Name = "ShowFilesCalRam"
        Me.ShowFilesCalRam.Size = New System.Drawing.Size(115, 37)
        Me.ShowFilesCalRam.TabIndex = 556
        Me.ShowFilesCalRam.Text = "\WinGPIBdata"
        Me.ToolTip1.SetToolTip(Me.ShowFilesCalRam, "Launch Windows File Explorer")
        Me.ShowFilesCalRam.UseVisualStyleBackColor = True
        '
        'ShowFiles2
        '
        Me.ShowFiles2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ShowFiles2.Location = New System.Drawing.Point(951, 319)
        Me.ShowFiles2.Name = "ShowFiles2"
        Me.ShowFiles2.Size = New System.Drawing.Size(90, 37)
        Me.ShowFiles2.TabIndex = 557
        Me.ShowFiles2.Text = "\WinGPIBdata"
        Me.ToolTip1.SetToolTip(Me.ShowFiles2, "Launch Windows File Explorer")
        Me.ShowFiles2.UseVisualStyleBackColor = True
        '
        'exportBTN
        '
        Me.exportBTN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exportBTN.Location = New System.Drawing.Point(23, 315)
        Me.exportBTN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.exportBTN.Name = "exportBTN"
        Me.exportBTN.Size = New System.Drawing.Size(138, 37)
        Me.exportBTN.TabIndex = 559
        Me.exportBTN.Text = "Generate Certificate"
        Me.ToolTip1.SetToolTip(Me.exportBTN, "Generate full Cal Cert - Requires MS Word install")
        Me.exportBTN.UseVisualStyleBackColor = True
        '
        'ButtonSetPrecalvars
        '
        Me.ButtonSetPrecalvars.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSetPrecalvars.Location = New System.Drawing.Point(23, 92)
        Me.ButtonSetPrecalvars.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonSetPrecalvars.Name = "ButtonSetPrecalvars"
        Me.ButtonSetPrecalvars.Size = New System.Drawing.Size(138, 37)
        Me.ButtonSetPrecalvars.TabIndex = 702
        Me.ButtonSetPrecalvars.Text = "Send X-Y Pre-Cal." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Data to PDVS2mini"
        Me.ToolTip1.SetToolTip(Me.ButtonSetPrecalvars, "Send Pre-Cal X/Y data to PDVS2mini")
        Me.ButtonSetPrecalvars.UseVisualStyleBackColor = True
        '
        'nplc4_BTN
        '
        Me.nplc4_BTN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nplc4_BTN.Location = New System.Drawing.Point(181, 87)
        Me.nplc4_BTN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.nplc4_BTN.Name = "nplc4_BTN"
        Me.nplc4_BTN.Size = New System.Drawing.Size(85, 25)
        Me.nplc4_BTN.TabIndex = 699
        Me.nplc4_BTN.Text = "Set 200 NPLC" & Global.Microsoft.VisualBasic.ChrW(13)
        Me.ToolTip1.SetToolTip(Me.nplc4_BTN, "Set 200 NPLC on 3458A")
        Me.nplc4_BTN.UseVisualStyleBackColor = True
        '
        'nplc3_BTN
        '
        Me.nplc3_BTN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nplc3_BTN.Location = New System.Drawing.Point(181, 152)
        Me.nplc3_BTN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.nplc3_BTN.Name = "nplc3_BTN"
        Me.nplc3_BTN.Size = New System.Drawing.Size(85, 25)
        Me.nplc3_BTN.TabIndex = 698
        Me.nplc3_BTN.Text = "Set 50 NPLC"
        Me.ToolTip1.SetToolTip(Me.nplc3_BTN, "Set 50 NPLC on 3458A")
        Me.nplc3_BTN.UseVisualStyleBackColor = True
        '
        'nplc2_BTN
        '
        Me.nplc2_BTN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nplc2_BTN.Location = New System.Drawing.Point(181, 184)
        Me.nplc2_BTN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.nplc2_BTN.Name = "nplc2_BTN"
        Me.nplc2_BTN.Size = New System.Drawing.Size(85, 25)
        Me.nplc2_BTN.TabIndex = 697
        Me.nplc2_BTN.Text = "Set 25 NPLC" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.ToolTip1.SetToolTip(Me.nplc2_BTN, "Set 25 NPLC on 3458A")
        Me.nplc2_BTN.UseVisualStyleBackColor = True
        '
        'getcal_BTN
        '
        Me.getcal_BTN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.getcal_BTN.Location = New System.Drawing.Point(23, 37)
        Me.getcal_BTN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.getcal_BTN.Name = "getcal_BTN"
        Me.getcal_BTN.Size = New System.Drawing.Size(138, 37)
        Me.getcal_BTN.TabIndex = 695
        Me.getcal_BTN.Text = "Get all Cal. Data" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "from PDVS2mini"
        Me.ToolTip1.SetToolTip(Me.getcal_BTN, "Get full Cal data from PDVS2mini")
        Me.getcal_BTN.UseVisualStyleBackColor = True
        '
        'nplc_BTN
        '
        Me.nplc_BTN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nplc_BTN.Location = New System.Drawing.Point(181, 120)
        Me.nplc_BTN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.nplc_BTN.Name = "nplc_BTN"
        Me.nplc_BTN.Size = New System.Drawing.Size(85, 25)
        Me.nplc_BTN.TabIndex = 695
        Me.nplc_BTN.Text = "Set 100 NPLC"
        Me.ToolTip1.SetToolTip(Me.nplc_BTN, "Set 100 NPLC on 3458A")
        Me.nplc_BTN.UseVisualStyleBackColor = True
        '
        'Abort_BTN
        '
        Me.Abort_BTN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Abort_BTN.Location = New System.Drawing.Point(496, 461)
        Me.Abort_BTN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Abort_BTN.Name = "Abort_BTN"
        Me.Abort_BTN.Size = New System.Drawing.Size(90, 50)
        Me.Abort_BTN.TabIndex = 546
        Me.Abort_BTN.Text = "Abort"
        Me.ToolTip1.SetToolTip(Me.Abort_BTN, "Abort Cal")
        Me.Abort_BTN.UseVisualStyleBackColor = True
        '
        'precal_BTN
        '
        Me.precal_BTN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.precal_BTN.Location = New System.Drawing.Point(23, 134)
        Me.precal_BTN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.precal_BTN.Name = "precal_BTN"
        Me.precal_BTN.Size = New System.Drawing.Size(138, 37)
        Me.precal_BTN.TabIndex = 544
        Me.precal_BTN.Text = "Run Full Auto-Cal" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "From Scratch"
        Me.ToolTip1.SetToolTip(Me.precal_BTN, "Run full Cal on PDVS2mini")
        Me.precal_BTN.UseVisualStyleBackColor = True
        '
        'SavePDVS2Eprom
        '
        Me.SavePDVS2Eprom.Location = New System.Drawing.Point(23, 273)
        Me.SavePDVS2Eprom.Name = "SavePDVS2Eprom"
        Me.SavePDVS2Eprom.Size = New System.Drawing.Size(138, 37)
        Me.SavePDVS2Eprom.TabIndex = 537
        Me.SavePDVS2Eprom.Text = "Save Cal. to EEprom (V1.3+ only)"
        Me.ToolTip1.SetToolTip(Me.SavePDVS2Eprom, "Save Cal currently on PDVS2mini to EEprom")
        Me.SavePDVS2Eprom.UseVisualStyleBackColor = True
        '
        'Dev23457Aseven
        '
        Me.Dev23457Aseven.AutoSize = True
        Me.Dev23457Aseven.Location = New System.Drawing.Point(119, 10)
        Me.Dev23457Aseven.Name = "Dev23457Aseven"
        Me.Dev23457Aseven.Size = New System.Drawing.Size(150, 17)
        Me.Dev23457Aseven.TabIndex = 95
        Me.Dev23457Aseven.Text = "HP3457A Enable 7th Digit"
        Me.ToolTip1.SetToolTip(Me.Dev23457Aseven, "Enable 7th digit on HP3457A DMM")
        Me.Dev23457Aseven.UseVisualStyleBackColor = True
        '
        'Dev13457Aseven
        '
        Me.Dev13457Aseven.AutoSize = True
        Me.Dev13457Aseven.Location = New System.Drawing.Point(121, 10)
        Me.Dev13457Aseven.Name = "Dev13457Aseven"
        Me.Dev13457Aseven.Size = New System.Drawing.Size(150, 17)
        Me.Dev13457Aseven.TabIndex = 94
        Me.Dev13457Aseven.Text = "HP3457A Enable 7th Digit"
        Me.ToolTip1.SetToolTip(Me.Dev13457Aseven, "Enable 7th digit on HP3457A DMM")
        Me.Dev13457Aseven.UseVisualStyleBackColor = True
        '
        'ShowFiles3
        '
        Me.ShowFiles3.BackColor = System.Drawing.Color.Thistle
        Me.ShowFiles3.Location = New System.Drawing.Point(11, 81)
        Me.ShowFiles3.Name = "ShowFiles3"
        Me.ShowFiles3.Size = New System.Drawing.Size(91, 29)
        Me.ShowFiles3.TabIndex = 686
        Me.ShowFiles3.Text = "\WinGPIBdata" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.ToolTip1.SetToolTip(Me.ShowFiles3, "Launch Windows File Explorer")
        Me.ShowFiles3.UseVisualStyleBackColor = True
        '
        'CSVfilename
        '
        Me.CSVfilename.Location = New System.Drawing.Point(100, 30)
        Me.CSVfilename.Name = "CSVfilename"
        Me.CSVfilename.Size = New System.Drawing.Size(271, 20)
        Me.CSVfilename.TabIndex = 50
        Me.CSVfilename.Text = "Log.csv"
        Me.ToolTip1.SetToolTip(Me.CSVfilename, "Edit this field to change the CSV filename")
        '
        'CSVfilepath
        '
        Me.CSVfilepath.Location = New System.Drawing.Point(100, 56)
        Me.CSVfilepath.Name = "CSVfilepath"
        Me.CSVfilepath.Size = New System.Drawing.Size(533, 20)
        Me.CSVfilepath.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.CSVfilepath, "Edit this field to change the CSV filepath")
        '
        'Dev2TerminatorEnable2
        '
        Me.Dev2TerminatorEnable2.AutoSize = True
        Me.Dev2TerminatorEnable2.Location = New System.Drawing.Point(355, 42)
        Me.Dev2TerminatorEnable2.Name = "Dev2TerminatorEnable2"
        Me.Dev2TerminatorEnable2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Dev2TerminatorEnable2.Size = New System.Drawing.Size(106, 17)
        Me.Dev2TerminatorEnable2.TabIndex = 96
        Me.Dev2TerminatorEnable2.Text = "Terminator CRLF"
        Me.ToolTip1.SetToolTip(Me.Dev2TerminatorEnable2, "\r, 0x0D, 13")
        Me.Dev2TerminatorEnable2.UseVisualStyleBackColor = True
        '
        'Dev2TerminatorEnable
        '
        Me.Dev2TerminatorEnable.AutoSize = True
        Me.Dev2TerminatorEnable.Location = New System.Drawing.Point(370, 26)
        Me.Dev2TerminatorEnable.Name = "Dev2TerminatorEnable"
        Me.Dev2TerminatorEnable.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Dev2TerminatorEnable.Size = New System.Drawing.Size(91, 17)
        Me.Dev2TerminatorEnable.TabIndex = 87
        Me.Dev2TerminatorEnable.Text = "Terminator LF"
        Me.ToolTip1.SetToolTip(Me.Dev2TerminatorEnable, "\n, 0x0A, 10, newline")
        Me.Dev2TerminatorEnable.UseVisualStyleBackColor = True
        '
        'Dev1TerminatorEnable2
        '
        Me.Dev1TerminatorEnable2.AutoSize = True
        Me.Dev1TerminatorEnable2.Location = New System.Drawing.Point(356, 42)
        Me.Dev1TerminatorEnable2.Name = "Dev1TerminatorEnable2"
        Me.Dev1TerminatorEnable2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Dev1TerminatorEnable2.Size = New System.Drawing.Size(106, 17)
        Me.Dev1TerminatorEnable2.TabIndex = 95
        Me.Dev1TerminatorEnable2.Text = "Terminator CRLF"
        Me.ToolTip1.SetToolTip(Me.Dev1TerminatorEnable2, "\r, 0x0D, 13")
        Me.Dev1TerminatorEnable2.UseVisualStyleBackColor = True
        '
        'Dev1TerminatorEnable
        '
        Me.Dev1TerminatorEnable.AutoSize = True
        Me.Dev1TerminatorEnable.Location = New System.Drawing.Point(371, 26)
        Me.Dev1TerminatorEnable.Name = "Dev1TerminatorEnable"
        Me.Dev1TerminatorEnable.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Dev1TerminatorEnable.Size = New System.Drawing.Size(91, 17)
        Me.Dev1TerminatorEnable.TabIndex = 88
        Me.Dev1TerminatorEnable.Text = "Terminator LF"
        Me.ToolTip1.SetToolTip(Me.Dev1TerminatorEnable, "\n, 0x0A, 10, newline")
        Me.Dev1TerminatorEnable.UseVisualStyleBackColor = True
        '
        'EditMode
        '
        Me.EditMode.AutoSize = True
        Me.EditMode.Location = New System.Drawing.Point(966, 191)
        Me.EditMode.Name = "EditMode"
        Me.EditMode.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.EditMode.Size = New System.Drawing.Size(74, 17)
        Me.EditMode.TabIndex = 491
        Me.EditMode.Text = "Edit Mode"
        Me.ToolTip1.SetToolTip(Me.EditMode, "Use Edit mode to change your profiles whilst offline")
        Me.EditMode.UseVisualStyleBackColor = True
        '
        'CheckboxCSVlimit
        '
        Me.CheckboxCSVlimit.AutoSize = True
        Me.CheckboxCSVlimit.Location = New System.Drawing.Point(494, 281)
        Me.CheckboxCSVlimit.Name = "CheckboxCSVlimit"
        Me.CheckboxCSVlimit.Size = New System.Drawing.Size(149, 17)
        Me.CheckboxCSVlimit.TabIndex = 86
        Me.CheckboxCSVlimit.Text = "CSV Limit (Device Entries)"
        Me.ToolTip1.SetToolTip(Me.CheckboxCSVlimit, "Put limit on entries recorded to the CSV")
        Me.CheckboxCSVlimit.UseVisualStyleBackColor = True
        '
        'CheckboxCSVlimitMins
        '
        Me.CheckboxCSVlimitMins.AutoSize = True
        Me.CheckboxCSVlimitMins.Location = New System.Drawing.Point(494, 329)
        Me.CheckboxCSVlimitMins.Name = "CheckboxCSVlimitMins"
        Me.CheckboxCSVlimitMins.Size = New System.Drawing.Size(103, 17)
        Me.CheckboxCSVlimitMins.TabIndex = 89
        Me.CheckboxCSVlimitMins.Text = "CSV Limit (Time)"
        Me.ToolTip1.SetToolTip(Me.CheckboxCSVlimitMins, "Put Time limit (Mins) on entries recorded to the CSV")
        Me.CheckboxCSVlimitMins.UseVisualStyleBackColor = True
        '
        'TextFilenameAppend
        '
        Me.TextFilenameAppend.Location = New System.Drawing.Point(472, 517)
        Me.TextFilenameAppend.MaxLength = 35
        Me.TextFilenameAppend.Name = "TextFilenameAppend"
        Me.TextFilenameAppend.Size = New System.Drawing.Size(196, 20)
        Me.TextFilenameAppend.TabIndex = 92
        Me.ToolTip1.SetToolTip(Me.TextFilenameAppend, "Add up to 35 chars to the filename when you Export")
        Me.TextFilenameAppend.Visible = False
        '
        'Dev1K2001isolatedata
        '
        Me.Dev1K2001isolatedata.AutoSize = True
        Me.Dev1K2001isolatedata.Location = New System.Drawing.Point(121, 42)
        Me.Dev1K2001isolatedata.Name = "Dev1K2001isolatedata"
        Me.Dev1K2001isolatedata.Size = New System.Drawing.Size(150, 17)
        Me.Dev1K2001isolatedata.TabIndex = 97
        Me.Dev1K2001isolatedata.Text = "Isolate data to left of Char:"
        Me.ToolTip1.SetToolTip(Me.Dev1K2001isolatedata, "Split the incoming string at Char and use the data to the left")
        Me.Dev1K2001isolatedata.UseVisualStyleBackColor = True
        '
        'Dev2K2001isolatedata
        '
        Me.Dev2K2001isolatedata.AutoSize = True
        Me.Dev2K2001isolatedata.Location = New System.Drawing.Point(119, 42)
        Me.Dev2K2001isolatedata.Name = "Dev2K2001isolatedata"
        Me.Dev2K2001isolatedata.Size = New System.Drawing.Size(150, 17)
        Me.Dev2K2001isolatedata.TabIndex = 98
        Me.Dev2K2001isolatedata.Text = "Isolate data to left of Char:"
        Me.ToolTip1.SetToolTip(Me.Dev2K2001isolatedata, "Split the incoming string at Char and use the data to the left")
        Me.Dev2K2001isolatedata.UseVisualStyleBackColor = True
        '
        'Div1000Dev2
        '
        Me.Div1000Dev2.AutoSize = True
        Me.Div1000Dev2.Location = New System.Drawing.Point(8, 65)
        Me.Div1000Dev2.Name = "Div1000Dev2"
        Me.Div1000Dev2.Size = New System.Drawing.Size(55, 17)
        Me.Div1000Dev2.TabIndex = 94
        Me.Div1000Dev2.Text = "/1000"
        Me.ToolTip1.SetToolTip(Me.Div1000Dev2, "Divide by 1000")
        Me.Div1000Dev2.UseVisualStyleBackColor = True
        '
        'Div1000Dev1
        '
        Me.Div1000Dev1.AutoSize = True
        Me.Div1000Dev1.Location = New System.Drawing.Point(8, 65)
        Me.Div1000Dev1.Name = "Div1000Dev1"
        Me.Div1000Dev1.Size = New System.Drawing.Size(55, 17)
        Me.Div1000Dev1.TabIndex = 93
        Me.Div1000Dev1.Text = "/1000"
        Me.ToolTip1.SetToolTip(Me.Div1000Dev1, "Divide by 1000")
        Me.Div1000Dev1.UseVisualStyleBackColor = True
        '
        'Mult1000Dev1
        '
        Me.Mult1000Dev1.AutoSize = True
        Me.Mult1000Dev1.Location = New System.Drawing.Point(8, 81)
        Me.Mult1000Dev1.Name = "Mult1000Dev1"
        Me.Mult1000Dev1.Size = New System.Drawing.Size(55, 17)
        Me.Mult1000Dev1.TabIndex = 99
        Me.Mult1000Dev1.Text = "x1000"
        Me.ToolTip1.SetToolTip(Me.Mult1000Dev1, "Multiply by 1000")
        Me.Mult1000Dev1.UseVisualStyleBackColor = True
        '
        'Mult1000Dev2
        '
        Me.Mult1000Dev2.AutoSize = True
        Me.Mult1000Dev2.Location = New System.Drawing.Point(8, 81)
        Me.Mult1000Dev2.Name = "Mult1000Dev2"
        Me.Mult1000Dev2.Size = New System.Drawing.Size(55, 17)
        Me.Mult1000Dev2.TabIndex = 100
        Me.Mult1000Dev2.Text = "x1000"
        Me.ToolTip1.SetToolTip(Me.Mult1000Dev2, "Multiply by 1000")
        Me.Mult1000Dev2.UseVisualStyleBackColor = True
        '
        'CalOnExisting
        '
        Me.CalOnExisting.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalOnExisting.Location = New System.Drawing.Point(181, 261)
        Me.CalOnExisting.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CalOnExisting.Name = "CalOnExisting"
        Me.CalOnExisting.Size = New System.Drawing.Size(168, 25)
        Me.CalOnExisting.TabIndex = 747
        Me.CalOnExisting.Text = "Run Auto-Cal on Current Counts"
        Me.ToolTip1.SetToolTip(Me.CalOnExisting, "Optional - Run Cal on current counts")
        Me.CalOnExisting.UseVisualStyleBackColor = True
        '
        'Label253
        '
        Me.Label253.AutoSize = True
        Me.Label253.Location = New System.Drawing.Point(345, 86)
        Me.Label253.Name = "Label253"
        Me.Label253.Size = New System.Drawing.Size(73, 13)
        Me.Label253.TabIndex = 105
        Me.Label253.Text = "Delay Op.(ms)"
        Me.ToolTip1.SetToolTip(Me.Label253, "Delay (ms) to wait between operations")
        '
        'Label251
        '
        Me.Label251.AutoSize = True
        Me.Label251.Location = New System.Drawing.Point(353, 63)
        Me.Label251.Name = "Label251"
        Me.Label251.Size = New System.Drawing.Size(67, 13)
        Me.Label251.TabIndex = 103
        Me.Label251.Text = "Timeout (ms)"
        Me.ToolTip1.SetToolTip(Me.Label251, "Time out (ms) wait time before abandon command")
        '
        'Label252
        '
        Me.Label252.AutoSize = True
        Me.Label252.Location = New System.Drawing.Point(348, 86)
        Me.Label252.Name = "Label252"
        Me.Label252.Size = New System.Drawing.Size(73, 13)
        Me.Label252.TabIndex = 103
        Me.Label252.Text = "Delay Op.(ms)"
        Me.ToolTip1.SetToolTip(Me.Label252, "Delay (ms) to wait between operations")
        '
        'Label250
        '
        Me.Label250.AutoSize = True
        Me.Label250.Location = New System.Drawing.Point(354, 63)
        Me.Label250.Name = "Label250"
        Me.Label250.Size = New System.Drawing.Size(67, 13)
        Me.Label250.TabIndex = 101
        Me.Label250.Text = "Timeout (ms)"
        Me.ToolTip1.SetToolTip(Me.Label250, "Time out (ms) wait time before abandon command")
        '
        'ButtonSetRetrievedVars
        '
        Me.ButtonSetRetrievedVars.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSetRetrievedVars.Location = New System.Drawing.Point(181, 231)
        Me.ButtonSetRetrievedVars.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonSetRetrievedVars.Name = "ButtonSetRetrievedVars"
        Me.ButtonSetRetrievedVars.Size = New System.Drawing.Size(168, 25)
        Me.ButtonSetRetrievedVars.TabIndex = 748
        Me.ButtonSetRetrievedVars.Text = "Send Counts to PDVS2mini"
        Me.ToolTip1.SetToolTip(Me.ButtonSetRetrievedVars, "Optional - Send current counts data to PDVS2mini")
        Me.ButtonSetRetrievedVars.UseVisualStyleBackColor = True
        '
        'ButtonAutoSET
        '
        Me.ButtonAutoSET.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonAutoSET.Location = New System.Drawing.Point(23, 189)
        Me.ButtonAutoSET.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonAutoSET.Name = "ButtonAutoSET"
        Me.ButtonAutoSET.Size = New System.Drawing.Size(138, 37)
        Me.ButtonAutoSET.TabIndex = 749
        Me.ButtonAutoSET.Text = "Auto SET all" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "X-Y (3 reads)"
        Me.ToolTip1.SetToolTip(Me.ButtonAutoSET, "Automatically run through all X-Y SET buttons")
        Me.ButtonAutoSET.UseVisualStyleBackColor = True
        '
        'ButtonAutomV
        '
        Me.ButtonAutomV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonAutomV.Location = New System.Drawing.Point(23, 231)
        Me.ButtonAutomV.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonAutomV.Name = "ButtonAutomV"
        Me.ButtonAutomV.Size = New System.Drawing.Size(138, 37)
        Me.ButtonAutomV.TabIndex = 750
        Me.ButtonAutomV.Text = "Auto mV" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(2 reads)"
        Me.ToolTip1.SetToolTip(Me.ButtonAutomV, "Automatically run through all mV data buttons")
        Me.ButtonAutomV.UseVisualStyleBackColor = True
        '
        'EnableAutoYChart1
        '
        Me.EnableAutoYChart1.AutoSize = True
        Me.EnableAutoYChart1.Location = New System.Drawing.Point(357, 26)
        Me.EnableAutoYChart1.Name = "EnableAutoYChart1"
        Me.EnableAutoYChart1.Size = New System.Drawing.Size(196, 17)
        Me.EnableAutoYChart1.TabIndex = 691
        Me.EnableAutoYChart1.Text = "Enable Autoscale Y-axis (5 samples)"
        Me.ToolTip1.SetToolTip(Me.EnableAutoYChart1, "Autoscale Device 1, 2 or both. Whatever combination is enabled")
        Me.EnableAutoYChart1.UseVisualStyleBackColor = True
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(433, 46)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(97, 13)
        Me.Label41.TabIndex = 96
        Me.Label41.Text = "X-axis Scale Points"
        Me.ToolTip1.SetToolTip(Me.Label41, "Once number of points on graph are achieved the graph will change to rolling type" &
        "")
        '
        'ButtonPauseChart
        '
        Me.ButtonPauseChart.Location = New System.Drawing.Point(11, 9)
        Me.ButtonPauseChart.Name = "ButtonPauseChart"
        Me.ButtonPauseChart.Size = New System.Drawing.Size(91, 29)
        Me.ButtonPauseChart.TabIndex = 695
        Me.ButtonPauseChart.Text = "Start Chart"
        Me.ToolTip1.SetToolTip(Me.ButtonPauseChart, "Run/Pause the chart")
        Me.ButtonPauseChart.UseVisualStyleBackColor = True
        '
        'ButtonSaveLiveSettings
        '
        Me.ButtonSaveLiveSettings.BackColor = System.Drawing.Color.PaleGreen
        Me.ButtonSaveLiveSettings.Location = New System.Drawing.Point(11, 113)
        Me.ButtonSaveLiveSettings.Name = "ButtonSaveLiveSettings"
        Me.ButtonSaveLiveSettings.Size = New System.Drawing.Size(91, 29)
        Me.ButtonSaveLiveSettings.TabIndex = 703
        Me.ButtonSaveLiveSettings.Text = "Save Settings"
        Me.ToolTip1.SetToolTip(Me.ButtonSaveLiveSettings, "Save Live Chart Settings")
        Me.ButtonSaveLiveSettings.UseVisualStyleBackColor = True
        '
        'CheckBoxAZERO
        '
        Me.CheckBoxAZERO.AutoSize = True
        Me.CheckBoxAZERO.Checked = True
        Me.CheckBoxAZERO.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxAZERO.Location = New System.Drawing.Point(187, 189)
        Me.CheckBoxAZERO.Name = "CheckBoxAZERO"
        Me.CheckBoxAZERO.Size = New System.Drawing.Size(116, 17)
        Me.CheckBoxAZERO.TabIndex = 611
        Me.CheckBoxAZERO.Text = "3458A AZERO ON"
        Me.ToolTip1.SetToolTip(Me.CheckBoxAZERO, "3458A Auto Zero function")
        Me.CheckBoxAZERO.UseVisualStyleBackColor = True
        '
        'ClearLOGdisp
        '
        Me.ClearLOGdisp.Location = New System.Drawing.Point(591, 228)
        Me.ClearLOGdisp.Name = "ClearLOGdisp"
        Me.ClearLOGdisp.Size = New System.Drawing.Size(64, 21)
        Me.ClearLOGdisp.TabIndex = 94
        Me.ClearLOGdisp.Text = "Clear LOG"
        Me.ToolTip1.SetToolTip(Me.ClearLOGdisp, "Clear the LOG display")
        Me.ClearLOGdisp.UseVisualStyleBackColor = True
        '
        'ButtonSaveDefs
        '
        Me.ButtonSaveDefs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSaveDefs.Location = New System.Drawing.Point(963, 312)
        Me.ButtonSaveDefs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonSaveDefs.Name = "ButtonSaveDefs"
        Me.ButtonSaveDefs.Size = New System.Drawing.Size(65, 37)
        Me.ButtonSaveDefs.TabIndex = 774
        Me.ButtonSaveDefs.Text = "Save " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "to PC"
        Me.ToolTip1.SetToolTip(Me.ButtonSaveDefs, "Save counts defaults to app.")
        Me.ButtonSaveDefs.UseVisualStyleBackColor = True
        '
        'ButtonLoadDefs
        '
        Me.ButtonLoadDefs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonLoadDefs.Location = New System.Drawing.Point(889, 312)
        Me.ButtonLoadDefs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonLoadDefs.Name = "ButtonLoadDefs"
        Me.ButtonLoadDefs.Size = New System.Drawing.Size(65, 37)
        Me.ButtonLoadDefs.TabIndex = 775
        Me.ButtonLoadDefs.Text = "Load" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "from PC" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.ToolTip1.SetToolTip(Me.ButtonLoadDefs, "Load saved counts from app.")
        Me.ButtonLoadDefs.UseVisualStyleBackColor = True
        '
        'Dev1IntEnable
        '
        Me.Dev1IntEnable.AutoSize = True
        Me.Dev1IntEnable.Location = New System.Drawing.Point(66, 357)
        Me.Dev1IntEnable.Name = "Dev1IntEnable"
        Me.Dev1IntEnable.Size = New System.Drawing.Size(101, 17)
        Me.Dev1IntEnable.TabIndex = 113
        Me.Dev1IntEnable.Text = "Interrupt Enable"
        Me.ToolTip1.SetToolTip(Me.Dev1IntEnable, "Interrupt Enable")
        Me.Dev1IntEnable.UseVisualStyleBackColor = True
        '
        'Dev2IntEnable
        '
        Me.Dev2IntEnable.AutoSize = True
        Me.Dev2IntEnable.Location = New System.Drawing.Point(61, 357)
        Me.Dev2IntEnable.Name = "Dev2IntEnable"
        Me.Dev2IntEnable.Size = New System.Drawing.Size(101, 17)
        Me.Dev2IntEnable.TabIndex = 114
        Me.Dev2IntEnable.Text = "Interrupt Enable"
        Me.ToolTip1.SetToolTip(Me.Dev2IntEnable, "Interrupt Enable")
        Me.Dev2IntEnable.UseVisualStyleBackColor = True
        '
        'ButtonRefreshPorts
        '
        Me.ButtonRefreshPorts.Enabled = False
        Me.ButtonRefreshPorts.Location = New System.Drawing.Point(248, 95)
        Me.ButtonRefreshPorts.Name = "ButtonRefreshPorts"
        Me.ButtonRefreshPorts.Size = New System.Drawing.Size(55, 23)
        Me.ButtonRefreshPorts.TabIndex = 495
        Me.ButtonRefreshPorts.Text = "Refresh"
        Me.ToolTip1.SetToolTip(Me.ButtonRefreshPorts, "Refresh available COM ports")
        Me.ButtonRefreshPorts.UseVisualStyleBackColor = True
        '
        'TextBoxProtocolInput
        '
        Me.TextBoxProtocolInput.Location = New System.Drawing.Point(108, 366)
        Me.TextBoxProtocolInput.Name = "TextBoxProtocolInput"
        Me.TextBoxProtocolInput.Size = New System.Drawing.Size(109, 20)
        Me.TextBoxProtocolInput.TabIndex = 496
        Me.TextBoxProtocolInput.Text = "GT"
        Me.ToolTip1.SetToolTip(Me.TextBoxProtocolInput, "Enter the serial command used to initiate a read from your sensor")
        '
        'TextBoxResult
        '
        Me.TextBoxResult.Location = New System.Drawing.Point(108, 392)
        Me.TextBoxResult.Name = "TextBoxResult"
        Me.TextBoxResult.ReadOnly = True
        Me.TextBoxResult.Size = New System.Drawing.Size(109, 20)
        Me.TextBoxResult.TabIndex = 497
        Me.ToolTip1.SetToolTip(Me.TextBoxResult, "Raw data as received from your sensor")
        '
        'TextBoxFinalTempValue
        '
        Me.TextBoxFinalTempValue.Location = New System.Drawing.Point(108, 522)
        Me.TextBoxFinalTempValue.Name = "TextBoxFinalTempValue"
        Me.TextBoxFinalTempValue.ReadOnly = True
        Me.TextBoxFinalTempValue.Size = New System.Drawing.Size(109, 20)
        Me.TextBoxFinalTempValue.TabIndex = 505
        Me.ToolTip1.SetToolTip(Me.TextBoxFinalTempValue, "Parsed final result")
        '
        'TextBoxRegex
        '
        Me.TextBoxRegex.Location = New System.Drawing.Point(129, 470)
        Me.TextBoxRegex.Name = "TextBoxRegex"
        Me.TextBoxRegex.Size = New System.Drawing.Size(88, 20)
        Me.TextBoxRegex.TabIndex = 510
        Me.TextBoxRegex.Text = "(\d+(\.\d+)?)"
        Me.ToolTip1.SetToolTip(Me.TextBoxRegex, "Enter the serial command used to initiate a read from your sensor")
        '
        'TextBoxSerialPortBaud
        '
        Me.TextBoxSerialPortBaud.Location = New System.Drawing.Point(655, 366)
        Me.TextBoxSerialPortBaud.Name = "TextBoxSerialPortBaud"
        Me.TextBoxSerialPortBaud.Size = New System.Drawing.Size(109, 20)
        Me.TextBoxSerialPortBaud.TabIndex = 519
        Me.TextBoxSerialPortBaud.Text = "250000"
        Me.ToolTip1.SetToolTip(Me.TextBoxSerialPortBaud, "Baud Rate - Valid integer value")
        '
        'TextBoxSerialPortBits
        '
        Me.TextBoxSerialPortBits.Location = New System.Drawing.Point(655, 392)
        Me.TextBoxSerialPortBits.Name = "TextBoxSerialPortBits"
        Me.TextBoxSerialPortBits.Size = New System.Drawing.Size(109, 20)
        Me.TextBoxSerialPortBits.TabIndex = 521
        Me.TextBoxSerialPortBits.Text = "8"
        Me.ToolTip1.SetToolTip(Me.TextBoxSerialPortBits, "Baud Rate - Valid integer value")
        '
        'TextBoxSerialPortParity
        '
        Me.TextBoxSerialPortParity.Location = New System.Drawing.Point(655, 418)
        Me.TextBoxSerialPortParity.Name = "TextBoxSerialPortParity"
        Me.TextBoxSerialPortParity.Size = New System.Drawing.Size(109, 20)
        Me.TextBoxSerialPortParity.TabIndex = 523
        Me.TextBoxSerialPortParity.Text = "NONE"
        Me.ToolTip1.SetToolTip(Me.TextBoxSerialPortParity, "NONE, ODD, or EVEN")
        '
        'TextBoxSerialPortStop
        '
        Me.TextBoxSerialPortStop.Location = New System.Drawing.Point(655, 444)
        Me.TextBoxSerialPortStop.Name = "TextBoxSerialPortStop"
        Me.TextBoxSerialPortStop.Size = New System.Drawing.Size(109, 20)
        Me.TextBoxSerialPortStop.TabIndex = 525
        Me.TextBoxSerialPortStop.Text = "1"
        Me.ToolTip1.SetToolTip(Me.TextBoxSerialPortStop, "No. of Stop Bits - 1, 1.5 or 2")
        '
        'TextBoxSerialPortHand
        '
        Me.TextBoxSerialPortHand.Location = New System.Drawing.Point(655, 470)
        Me.TextBoxSerialPortHand.Name = "TextBoxSerialPortHand"
        Me.TextBoxSerialPortHand.Size = New System.Drawing.Size(109, 20)
        Me.TextBoxSerialPortHand.TabIndex = 527
        Me.TextBoxSerialPortHand.Text = "NONE"
        Me.ToolTip1.SetToolTip(Me.TextBoxSerialPortHand, "NONE, XONXOFF, RTSCTS, or RTSXONXOFF")
        '
        'ButtonSaveTempHumSettings
        '
        Me.ButtonSaveTempHumSettings.BackColor = System.Drawing.Color.PaleGreen
        Me.ButtonSaveTempHumSettings.Location = New System.Drawing.Point(330, 40)
        Me.ButtonSaveTempHumSettings.Name = "ButtonSaveTempHumSettings"
        Me.ButtonSaveTempHumSettings.Size = New System.Drawing.Size(82, 29)
        Me.ButtonSaveTempHumSettings.TabIndex = 704
        Me.ButtonSaveTempHumSettings.Text = "Save Settings"
        Me.ToolTip1.SetToolTip(Me.ButtonSaveTempHumSettings, "Save COM Port Settings")
        Me.ButtonSaveTempHumSettings.UseVisualStyleBackColor = True
        '
        'Dev1Regex
        '
        Me.Dev1Regex.AutoSize = True
        Me.Dev1Regex.Location = New System.Drawing.Point(121, 58)
        Me.Dev1Regex.Name = "Dev1Regex"
        Me.Dev1Regex.Size = New System.Drawing.Size(138, 17)
        Me.Dev1Regex.TabIndex = 114
        Me.Dev1Regex.Text = "Regex isolate numerical"
        Me.ToolTip1.SetToolTip(Me.Dev1Regex, "Isolate numerical data" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Regex = [-+]?\d*\.?\d+([eE][-+]?\d+)?")
        Me.Dev1Regex.UseVisualStyleBackColor = True
        '
        'Dev2Regex
        '
        Me.Dev2Regex.AutoSize = True
        Me.Dev2Regex.Location = New System.Drawing.Point(119, 58)
        Me.Dev2Regex.Name = "Dev2Regex"
        Me.Dev2Regex.Size = New System.Drawing.Size(138, 17)
        Me.Dev2Regex.TabIndex = 118
        Me.Dev2Regex.Text = "Regex isolate numerical"
        Me.ToolTip1.SetToolTip(Me.Dev2Regex, "Isolate numerical data" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Regex = [-+]?\d*\.?\d+([eE][-+]?\d+)?" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        Me.Dev2Regex.UseVisualStyleBackColor = True
        '
        'Dev1DecimalNumDPs
        '
        Me.Dev1DecimalNumDPs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev1DecimalNumDPs.Location = New System.Drawing.Point(239, 221)
        Me.Dev1DecimalNumDPs.Name = "Dev1DecimalNumDPs"
        Me.Dev1DecimalNumDPs.Size = New System.Drawing.Size(20, 20)
        Me.Dev1DecimalNumDPs.TabIndex = 115
        Me.Dev1DecimalNumDPs.Text = "8"
        Me.Dev1DecimalNumDPs.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ToolTip1.SetToolTip(Me.Dev1DecimalNumDPs, "Number of decimal points for display including Device Meter")
        '
        'Dev2DecimalNumDPs
        '
        Me.Dev2DecimalNumDPs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev2DecimalNumDPs.Location = New System.Drawing.Point(238, 221)
        Me.Dev2DecimalNumDPs.Name = "Dev2DecimalNumDPs"
        Me.Dev2DecimalNumDPs.Size = New System.Drawing.Size(20, 20)
        Me.Dev2DecimalNumDPs.TabIndex = 119
        Me.Dev2DecimalNumDPs.Text = "8"
        Me.Dev2DecimalNumDPs.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ToolTip1.SetToolTip(Me.Dev2DecimalNumDPs, "Number of decimal points for display including Device Meter")
        '
        'Dev2pauseDurationInSeconds
        '
        Me.Dev2pauseDurationInSeconds.Location = New System.Drawing.Point(177, 336)
        Me.Dev2pauseDurationInSeconds.Name = "Dev2pauseDurationInSeconds"
        Me.Dev2pauseDurationInSeconds.Size = New System.Drawing.Size(35, 20)
        Me.Dev2pauseDurationInSeconds.TabIndex = 116
        Me.ToolTip1.SetToolTip(Me.Dev2pauseDurationInSeconds, "How long the interrupt should last." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Should be set long enough to more" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "than enou" &
        "gh for the device to complete." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        '
        'Dev2runStopwatchEveryInMins
        '
        Me.Dev2runStopwatchEveryInMins.Location = New System.Drawing.Point(223, 336)
        Me.Dev2runStopwatchEveryInMins.Name = "Dev2runStopwatchEveryInMins"
        Me.Dev2runStopwatchEveryInMins.Size = New System.Drawing.Size(35, 20)
        Me.Dev2runStopwatchEveryInMins.TabIndex = 113
        Me.ToolTip1.SetToolTip(Me.Dev2runStopwatchEveryInMins, "How often the interrupt should be activated.")
        '
        'Dev2SampleRate
        '
        Me.Dev2SampleRate.Location = New System.Drawing.Point(132, 392)
        Me.Dev2SampleRate.Name = "Dev2SampleRate"
        Me.Dev2SampleRate.Size = New System.Drawing.Size(42, 20)
        Me.Dev2SampleRate.TabIndex = 22
        Me.Dev2SampleRate.Text = "5"
        Me.ToolTip1.SetToolTip(Me.Dev2SampleRate, "GPIB interval/period.")
        '
        'Dev1pauseDurationInSeconds
        '
        Me.Dev1pauseDurationInSeconds.Location = New System.Drawing.Point(178, 336)
        Me.Dev1pauseDurationInSeconds.Name = "Dev1pauseDurationInSeconds"
        Me.Dev1pauseDurationInSeconds.Size = New System.Drawing.Size(35, 20)
        Me.Dev1pauseDurationInSeconds.TabIndex = 109
        Me.ToolTip1.SetToolTip(Me.Dev1pauseDurationInSeconds, "How long the interrupt should last." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Should be set long enough to more" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "than enou" &
        "gh for the device to complete.")
        '
        'Dev1runStopwatchEveryInMins
        '
        Me.Dev1runStopwatchEveryInMins.Location = New System.Drawing.Point(224, 336)
        Me.Dev1runStopwatchEveryInMins.Name = "Dev1runStopwatchEveryInMins"
        Me.Dev1runStopwatchEveryInMins.Size = New System.Drawing.Size(35, 20)
        Me.Dev1runStopwatchEveryInMins.TabIndex = 106
        Me.ToolTip1.SetToolTip(Me.Dev1runStopwatchEveryInMins, "How often the interrupt should be activated.")
        '
        'Dev1SampleRate
        '
        Me.Dev1SampleRate.Location = New System.Drawing.Point(130, 392)
        Me.Dev1SampleRate.Name = "Dev1SampleRate"
        Me.Dev1SampleRate.Size = New System.Drawing.Size(41, 20)
        Me.Dev1SampleRate.TabIndex = 24
        Me.Dev1SampleRate.Text = "5"
        Me.ToolTip1.SetToolTip(Me.Dev1SampleRate, "GPIB interval/period.")
        '
        'XaxisPoints
        '
        Me.XaxisPoints.Location = New System.Drawing.Point(357, 45)
        Me.XaxisPoints.Name = "XaxisPoints"
        Me.XaxisPoints.Size = New System.Drawing.Size(71, 20)
        Me.XaxisPoints.TabIndex = 91
        Me.XaxisPoints.Text = "500"
        Me.ToolTip1.SetToolTip(Me.XaxisPoints, "Resolution && scroll mode")
        '
        'PDVS2miniSave
        '
        Me.PDVS2miniSave.BackColor = System.Drawing.Color.PaleGreen
        Me.PDVS2miniSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PDVS2miniSave.Location = New System.Drawing.Point(496, 413)
        Me.PDVS2miniSave.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.PDVS2miniSave.Name = "PDVS2miniSave"
        Me.PDVS2miniSave.Size = New System.Drawing.Size(90, 29)
        Me.PDVS2miniSave.TabIndex = 789
        Me.PDVS2miniSave.Text = "Save Settings"
        Me.ToolTip1.SetToolTip(Me.PDVS2miniSave, "Save Settings")
        Me.PDVS2miniSave.UseVisualStyleBackColor = True
        '
        'Dev1TextResponse
        '
        Me.Dev1TextResponse.AutoSize = True
        Me.Dev1TextResponse.Location = New System.Drawing.Point(121, 74)
        Me.Dev1TextResponse.Name = "Dev1TextResponse"
        Me.Dev1TextResponse.Size = New System.Drawing.Size(157, 17)
        Me.Dev1TextResponse.TabIndex = 117
        Me.Dev1TextResponse.Text = "Query Async text responses"
        Me.ToolTip1.SetToolTip(Me.Dev1TextResponse, "Allow non-numerical responses using Query Async")
        Me.Dev1TextResponse.UseVisualStyleBackColor = True
        '
        'Dev2TextResponse
        '
        Me.Dev2TextResponse.AutoSize = True
        Me.Dev2TextResponse.Location = New System.Drawing.Point(119, 74)
        Me.Dev2TextResponse.Name = "Dev2TextResponse"
        Me.Dev2TextResponse.Size = New System.Drawing.Size(157, 17)
        Me.Dev2TextResponse.TabIndex = 121
        Me.Dev2TextResponse.Text = "Query Async text responses"
        Me.ToolTip1.SetToolTip(Me.Dev2TextResponse, "Allow non-numerical responses using Query Async")
        Me.Dev2TextResponse.UseVisualStyleBackColor = True
        '
        'ButtonDev1PreRun
        '
        Me.ButtonDev1PreRun.Location = New System.Drawing.Point(347, 220)
        Me.ButtonDev1PreRun.Name = "ButtonDev1PreRun"
        Me.ButtonDev1PreRun.Size = New System.Drawing.Size(115, 21)
        Me.ButtonDev1PreRun.TabIndex = 119
        Me.ButtonDev1PreRun.Text = "Send PRE RUN only"
        Me.ToolTip1.SetToolTip(Me.ButtonDev1PreRun, "Send Device 1 PRE RUN commands only")
        Me.ButtonDev1PreRun.UseVisualStyleBackColor = True
        '
        'ButtonDev2PreRun
        '
        Me.ButtonDev2PreRun.Location = New System.Drawing.Point(346, 220)
        Me.ButtonDev2PreRun.Name = "ButtonDev2PreRun"
        Me.ButtonDev2PreRun.Size = New System.Drawing.Size(115, 21)
        Me.ButtonDev2PreRun.TabIndex = 123
        Me.ButtonDev2PreRun.Text = "Send PRE RUN only"
        Me.ToolTip1.SetToolTip(Me.ButtonDev2PreRun, "Send Device 2 PRE RUN commands only")
        Me.ButtonDev2PreRun.UseVisualStyleBackColor = True
        '
        'txtq2d
        '
        Me.txtq2d.Location = New System.Drawing.Point(61, 336)
        Me.txtq2d.Name = "txtq2d"
        Me.txtq2d.Size = New System.Drawing.Size(108, 20)
        Me.txtq2d.TabIndex = 111
        Me.ToolTip1.SetToolTip(Me.txtq2d, "Send Async commands only")
        '
        'txtq1d
        '
        Me.txtq1d.Location = New System.Drawing.Point(66, 336)
        Me.txtq1d.Name = "txtq1d"
        Me.txtq1d.Size = New System.Drawing.Size(104, 20)
        Me.txtq1d.TabIndex = 104
        Me.ToolTip1.SetToolTip(Me.txtq1d, "Send Async commands only")
        '
        'Dev1SendQuery
        '
        Me.Dev1SendQuery.AutoSize = True
        Me.Dev1SendQuery.Location = New System.Drawing.Point(66, 372)
        Me.Dev1SendQuery.Name = "Dev1SendQuery"
        Me.Dev1SendQuery.Size = New System.Drawing.Size(86, 17)
        Me.Dev1SendQuery.TabIndex = 120
        Me.Dev1SendQuery.Text = "Query Async"
        Me.ToolTip1.SetToolTip(Me.Dev1SendQuery, "Interrupt Send or Query Async")
        Me.Dev1SendQuery.UseVisualStyleBackColor = True
        '
        'Dev2SendQuery
        '
        Me.Dev2SendQuery.AutoSize = True
        Me.Dev2SendQuery.Location = New System.Drawing.Point(61, 372)
        Me.Dev2SendQuery.Name = "Dev2SendQuery"
        Me.Dev2SendQuery.Size = New System.Drawing.Size(86, 17)
        Me.Dev2SendQuery.TabIndex = 124
        Me.Dev2SendQuery.Text = "Query Async"
        Me.ToolTip1.SetToolTip(Me.Dev2SendQuery, "Interrupt Send or Query Async")
        Me.Dev2SendQuery.UseVisualStyleBackColor = True
        '
        'ClearEventLOG
        '
        Me.ClearEventLOG.Location = New System.Drawing.Point(975, 17)
        Me.ClearEventLOG.Name = "ClearEventLOG"
        Me.ClearEventLOG.Size = New System.Drawing.Size(64, 21)
        Me.ClearEventLOG.TabIndex = 101
        Me.ClearEventLOG.Text = "Clear"
        Me.ToolTip1.SetToolTip(Me.ClearEventLOG, "Clear the LOG display")
        Me.ClearEventLOG.UseVisualStyleBackColor = True
        '
        'Dev1Units
        '
        Me.Dev1Units.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Dev1Units.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Dev1Units.Font = New System.Drawing.Font("Microsoft Sans Serif", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev1Units.Location = New System.Drawing.Point(885, 166)
        Me.Dev1Units.Name = "Dev1Units"
        Me.Dev1Units.Size = New System.Drawing.Size(142, 42)
        Me.Dev1Units.TabIndex = 73
        Me.Dev1Units.Text = "VDC"
        Me.ToolTip1.SetToolTip(Me.Dev1Units, "Click here to edit the units")
        '
        'Dev2Units
        '
        Me.Dev2Units.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Dev2Units.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Dev2Units.Font = New System.Drawing.Font("Microsoft Sans Serif", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev2Units.Location = New System.Drawing.Point(885, 166)
        Me.Dev2Units.Name = "Dev2Units"
        Me.Dev2Units.Size = New System.Drawing.Size(142, 42)
        Me.Dev2Units.TabIndex = 74
        Me.Dev2Units.Text = "kΩ"
        Me.ToolTip1.SetToolTip(Me.Dev2Units, "Click here to edit the units")
        '
        'TextBoxHumUnits
        '
        Me.TextBoxHumUnits.Location = New System.Drawing.Point(159, 168)
        Me.TextBoxHumUnits.Name = "TextBoxHumUnits"
        Me.TextBoxHumUnits.Size = New System.Drawing.Size(38, 20)
        Me.TextBoxHumUnits.TabIndex = 516
        Me.TextBoxHumUnits.Text = "%RH"
        Me.ToolTip1.SetToolTip(Me.TextBoxHumUnits, "Humidity units")
        '
        'TextBoxTempUnits
        '
        Me.TextBoxTempUnits.Location = New System.Drawing.Point(159, 144)
        Me.TextBoxTempUnits.Name = "TextBoxTempUnits"
        Me.TextBoxTempUnits.Size = New System.Drawing.Size(38, 20)
        Me.TextBoxTempUnits.TabIndex = 515
        Me.TextBoxTempUnits.Text = "DegC"
        Me.ToolTip1.SetToolTip(Me.TextBoxTempUnits, "Temperature units")
        '
        'TempOffset
        '
        Me.TempOffset.Location = New System.Drawing.Point(218, 144)
        Me.TempOffset.Name = "TempOffset"
        Me.TempOffset.Size = New System.Drawing.Size(37, 20)
        Me.TempOffset.TabIndex = 486
        Me.TempOffset.Text = "0.0"
        Me.ToolTip1.SetToolTip(Me.TempOffset, "Offset value for the Temp readout")
        '
        'ButtonRefreshPorts1
        '
        Me.ButtonRefreshPorts1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRefreshPorts1.Location = New System.Drawing.Point(93, 11)
        Me.ButtonRefreshPorts1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonRefreshPorts1.Name = "ButtonRefreshPorts1"
        Me.ButtonRefreshPorts1.Size = New System.Drawing.Size(22, 21)
        Me.ButtonRefreshPorts1.TabIndex = 800
        Me.ButtonRefreshPorts1.Text = "R"
        Me.ToolTip1.SetToolTip(Me.ButtonRefreshPorts1, "Refresh available COM ports")
        Me.ButtonRefreshPorts1.UseVisualStyleBackColor = True
        '
        'WryTech
        '
        Me.WryTech.AutoSize = True
        Me.WryTech.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WryTech.Location = New System.Drawing.Point(379, 15)
        Me.WryTech.Name = "WryTech"
        Me.WryTech.Size = New System.Drawing.Size(98, 20)
        Me.WryTech.TabIndex = 828
        Me.WryTech.Text = "Wrytech unit"
        Me.ToolTip1.SetToolTip(Me.WryTech, "Wrytech PDVS2mini" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Additional setpoint and" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "internal temperature")
        Me.WryTech.UseVisualStyleBackColor = True
        '
        'btncreate2
        '
        Me.btncreate2.Location = New System.Drawing.Point(324, 40)
        Me.btncreate2.Name = "btncreate2"
        Me.btncreate2.Size = New System.Drawing.Size(90, 40)
        Me.btncreate2.TabIndex = 500
        Me.btncreate2.Text = "Connect to" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "I/O Device 1"
        Me.ToolTip1.SetToolTip(Me.btncreate2, "Connect to Device 1 only")
        Me.btncreate2.UseVisualStyleBackColor = True
        '
        'btncreate3
        '
        Me.btncreate3.Location = New System.Drawing.Point(324, 38)
        Me.btncreate3.Name = "btncreate3"
        Me.btncreate3.Size = New System.Drawing.Size(90, 40)
        Me.btncreate3.TabIndex = 521
        Me.btncreate3.Text = "Connect to" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "I/O Device 2"
        Me.ToolTip1.SetToolTip(Me.btncreate3, "Connect to Device 2 only")
        Me.btncreate3.UseVisualStyleBackColor = True
        '
        'btncreate
        '
        Me.btncreate.BackColor = System.Drawing.Color.WhiteSmoke
        Me.btncreate.Location = New System.Drawing.Point(0, 0)
        Me.btncreate.Name = "btncreate"
        Me.btncreate.Size = New System.Drawing.Size(90, 40)
        Me.btncreate.TabIndex = 506
        Me.btncreate.Text = "Connect to I/O" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Devices 1 && 2"
        Me.ToolTip1.SetToolTip(Me.btncreate, "Connect to Device 1 & Device 2")
        Me.btncreate.UseVisualStyleBackColor = True
        '
        'noEOI
        '
        Me.noEOI.AutoSize = True
        Me.noEOI.Location = New System.Drawing.Point(952, 101)
        Me.noEOI.Name = "noEOI"
        Me.noEOI.Size = New System.Drawing.Size(94, 17)
        Me.noEOI.TabIndex = 507
        Me.noEOI.Text = "EOI Incapable"
        Me.ToolTip1.SetToolTip(Me.noEOI, "EOI incapable instruments , terminator will be set to 10")
        Me.noEOI.UseVisualStyleBackColor = True
        '
        'TextBoxTempHumSample
        '
        Me.TextBoxTempHumSample.Location = New System.Drawing.Point(159, 204)
        Me.TextBoxTempHumSample.Name = "TextBoxTempHumSample"
        Me.TextBoxTempHumSample.Size = New System.Drawing.Size(38, 20)
        Me.TextBoxTempHumSample.TabIndex = 710
        Me.TextBoxTempHumSample.Text = "1"
        Me.ToolTip1.SetToolTip(Me.TextBoxTempHumSample, "Sampling Frequency - Seconds")
        '
        'txtname3
        '
        Me.txtname3.Location = New System.Drawing.Point(96, 41)
        Me.txtname3.Name = "txtname3"
        Me.txtname3.Size = New System.Drawing.Size(206, 20)
        Me.txtname3.TabIndex = 25
        Me.txtname3.Text = "Temp&Humidity"
        Me.ToolTip1.SetToolTip(Me.txtname3, "Give your probe a name for easy reference")
        '
        'ShowFilesCalRamR6581
        '
        Me.ShowFilesCalRamR6581.BackColor = System.Drawing.Color.Thistle
        Me.ShowFilesCalRamR6581.Location = New System.Drawing.Point(928, 16)
        Me.ShowFilesCalRamR6581.Name = "ShowFilesCalRamR6581"
        Me.ShowFilesCalRamR6581.Size = New System.Drawing.Size(115, 37)
        Me.ShowFilesCalRamR6581.TabIndex = 590
        Me.ShowFilesCalRamR6581.Text = "\WinGPIBdata"
        Me.ToolTip1.SetToolTip(Me.ShowFilesCalRamR6581, "Launch Windows File Explorer")
        Me.ShowFilesCalRamR6581.UseVisualStyleBackColor = True
        '
        'ButtonAvailableComPorts
        '
        Me.ButtonAvailableComPorts.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ButtonAvailableComPorts.Location = New System.Drawing.Point(951, 221)
        Me.ButtonAvailableComPorts.Name = "ButtonAvailableComPorts"
        Me.ButtonAvailableComPorts.Size = New System.Drawing.Size(90, 22)
        Me.ButtonAvailableComPorts.TabIndex = 562
        Me.ButtonAvailableComPorts.Text = "COM ports"
        Me.ToolTip1.SetToolTip(Me.ButtonAvailableComPorts, "Export profiles & settings data to ProfilesData.dat")
        Me.ButtonAvailableComPorts.UseVisualStyleBackColor = True
        '
        'ButtonJsonViewer
        '
        Me.ButtonJsonViewer.Enabled = False
        Me.ButtonJsonViewer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonJsonViewer.Location = New System.Drawing.Point(718, 539)
        Me.ButtonJsonViewer.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonJsonViewer.Name = "ButtonJsonViewer"
        Me.ButtonJsonViewer.Size = New System.Drawing.Size(80, 24)
        Me.ButtonJsonViewer.TabIndex = 638
        Me.ButtonJsonViewer.Text = "JSON Viewer"
        Me.ToolTip1.SetToolTip(Me.ButtonJsonViewer, "Open the JSON file in the built in viewer")
        Me.ButtonJsonViewer.UseVisualStyleBackColor = True
        '
        'ButtonOpenR6581fileSelectJson
        '
        Me.ButtonOpenR6581fileSelectJson.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonOpenR6581fileSelectJson.Location = New System.Drawing.Point(718, 512)
        Me.ButtonOpenR6581fileSelectJson.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonOpenR6581fileSelectJson.Name = "ButtonOpenR6581fileSelectJson"
        Me.ButtonOpenR6581fileSelectJson.Size = New System.Drawing.Size(80, 24)
        Me.ButtonOpenR6581fileSelectJson.TabIndex = 621
        Me.ButtonOpenR6581fileSelectJson.Text = "Select JSON"
        Me.ToolTip1.SetToolTip(Me.ButtonOpenR6581fileSelectJson, "Select the JSON file you wish to upload")
        Me.ButtonOpenR6581fileSelectJson.UseVisualStyleBackColor = True
        '
        'ButtonOpenR6581fileJson
        '
        Me.ButtonOpenR6581fileJson.Enabled = False
        Me.ButtonOpenR6581fileJson.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonOpenR6581fileJson.Location = New System.Drawing.Point(719, 238)
        Me.ButtonOpenR6581fileJson.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonOpenR6581fileJson.Name = "ButtonOpenR6581fileJson"
        Me.ButtonOpenR6581fileJson.Size = New System.Drawing.Size(80, 24)
        Me.ButtonOpenR6581fileJson.TabIndex = 607
        Me.ButtonOpenR6581fileJson.Text = "Open JSON"
        Me.ToolTip1.SetToolTip(Me.ButtonOpenR6581fileJson, "Open the JSON file in Notepad")
        Me.ButtonOpenR6581fileJson.UseVisualStyleBackColor = True
        '
        'ButtonOpenR6581file
        '
        Me.ButtonOpenR6581file.Enabled = False
        Me.ButtonOpenR6581file.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonOpenR6581file.Location = New System.Drawing.Point(719, 212)
        Me.ButtonOpenR6581file.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonOpenR6581file.Name = "ButtonOpenR6581file"
        Me.ButtonOpenR6581file.Size = New System.Drawing.Size(80, 24)
        Me.ButtonOpenR6581file.TabIndex = 605
        Me.ButtonOpenR6581file.Text = "Open Txt"
        Me.ToolTip1.SetToolTip(Me.ButtonOpenR6581file, "Open the TXT file in Notepad")
        Me.ButtonOpenR6581file.UseVisualStyleBackColor = True
        '
        'ButtonJsonViewer2
        '
        Me.ButtonJsonViewer2.Enabled = False
        Me.ButtonJsonViewer2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonJsonViewer2.Location = New System.Drawing.Point(803, 238)
        Me.ButtonJsonViewer2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonJsonViewer2.Name = "ButtonJsonViewer2"
        Me.ButtonJsonViewer2.Size = New System.Drawing.Size(80, 24)
        Me.ButtonJsonViewer2.TabIndex = 639
        Me.ButtonJsonViewer2.Text = "JSON Viewer"
        Me.ToolTip1.SetToolTip(Me.ButtonJsonViewer2, "Open the JSON file in the built in viewer")
        Me.ButtonJsonViewer2.UseVisualStyleBackColor = True
        '
        'btnRestore
        '
        Me.btnRestore.BackColor = System.Drawing.Color.WhiteSmoke
        Me.btnRestore.Location = New System.Drawing.Point(10, 156)
        Me.btnRestore.Name = "btnRestore"
        Me.btnRestore.Size = New System.Drawing.Size(90, 22)
        Me.btnRestore.TabIndex = 603
        Me.btnRestore.Text = "Import Profiles"
        Me.ToolTip1.SetToolTip(Me.btnRestore, "Import profiles & settings data directly from ProfilesData.dat")
        Me.btnRestore.UseVisualStyleBackColor = True
        '
        'btnBackup
        '
        Me.btnBackup.BackColor = System.Drawing.Color.WhiteSmoke
        Me.btnBackup.Location = New System.Drawing.Point(10, 125)
        Me.btnBackup.Name = "btnBackup"
        Me.btnBackup.Size = New System.Drawing.Size(90, 22)
        Me.btnBackup.TabIndex = 602
        Me.btnBackup.Text = "Export Profiles"
        Me.ToolTip1.SetToolTip(Me.btnBackup, "Export profiles & settings data to ProfilesData.dat")
        Me.btnBackup.UseVisualStyleBackColor = True
        '
        'CalRam3458APreRun
        '
        Me.CalRam3458APreRun.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalRam3458APreRun.Location = New System.Drawing.Point(229, 162)
        Me.CalRam3458APreRun.Multiline = True
        Me.CalRam3458APreRun.Name = "CalRam3458APreRun"
        Me.CalRam3458APreRun.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.CalRam3458APreRun.Size = New System.Drawing.Size(197, 71)
        Me.CalRam3458APreRun.TabIndex = 600
        Me.CalRam3458APreRun.Text = "END ALWAYS" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "NPLC 0" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "NRDGS 1" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "TRIG HOLD" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "QFORMAT NUM"
        Me.ToolTip1.SetToolTip(Me.CalRam3458APreRun, "These batch commands are sent when 3458A READ is" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "pressed and are Send Asynchrono" &
        "us type commands" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "i.e. expect no reply.")
        '
        'BtnSave3458A
        '
        Me.BtnSave3458A.BackColor = System.Drawing.Color.PaleGreen
        Me.BtnSave3458A.Location = New System.Drawing.Point(345, 134)
        Me.BtnSave3458A.Name = "BtnSave3458A"
        Me.BtnSave3458A.Size = New System.Drawing.Size(82, 25)
        Me.BtnSave3458A.TabIndex = 705
        Me.BtnSave3458A.Text = "Save Settings"
        Me.ToolTip1.SetToolTip(Me.BtnSave3458A, "Save 3458A Pre-Run")
        Me.BtnSave3458A.UseVisualStyleBackColor = True
        '
        'CheckBoxR6581RetrieveREF
        '
        Me.CheckBoxR6581RetrieveREF.AutoSize = True
        Me.CheckBoxR6581RetrieveREF.Location = New System.Drawing.Point(11, 107)
        Me.CheckBoxR6581RetrieveREF.Name = "CheckBoxR6581RetrieveREF"
        Me.CheckBoxR6581RetrieveREF.Size = New System.Drawing.Size(398, 17)
        Me.CheckBoxR6581RetrieveREF.TabIndex = 591
        Me.CheckBoxR6581RetrieveREF.Text = "Retrieve history (20) of 7.2Vdc && 10kohm internal reference value drift to .txt " &
    "file"
        Me.CheckBoxR6581RetrieveREF.UseVisualStyleBackColor = True
        '
        'SerialPort1
        '
        Me.SerialPort1.BaudRate = 115200
        Me.SerialPort1.DtrEnable = True
        Me.SerialPort1.ReadTimeout = 500
        Me.SerialPort1.WriteTimeout = 500
        '
        'Timer6
        '
        Me.Timer6.Interval = 50
        '
        'Timer7
        '
        Me.Timer7.Interval = 50
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage10)
        Me.TabControl1.Controls.Add(Me.TabPage8)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage9)
        Me.TabControl1.Controls.Add(Me.TabPage7)
        Me.TabControl1.Controls.Add(Me.TabPage11)
        Me.TabControl1.Controls.Add(Me.TabPage12)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Controls.Add(Me.TabPage13)
        Me.TabControl1.Controls.Add(Me.TabPage6)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Multiline = True
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1055, 625)
        Me.TabControl1.TabIndex = 518
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage1.Controls.Add(Me.ButtonAvailableComPorts)
        Me.TabPage1.Controls.Add(Me.Panel2)
        Me.TabPage1.Controls.Add(Me.Panel1)
        Me.TabPage1.Controls.Add(Me.Label133)
        Me.TabPage1.Controls.Add(Me.GroupBox9)
        Me.TabPage1.Controls.Add(Me.noEOI)
        Me.TabPage1.Controls.Add(Me.GroupBox8)
        Me.TabPage1.Controls.Add(Me.btndevlist)
        Me.TabPage1.Controls.Add(Me.Label303)
        Me.TabPage1.Controls.Add(Me.Label302)
        Me.TabPage1.Controls.Add(Me.RunningTimeLogging)
        Me.TabPage1.Controls.Add(Me.EditMode)
        Me.TabPage1.Controls.Add(Me.ShowFiles2)
        Me.TabPage1.Controls.Add(Me.Label56)
        Me.TabPage1.Controls.Add(Me.PictureBox5)
        Me.TabPage1.Controls.Add(Me.ButtonNotePad2)
        Me.TabPage1.Controls.Add(Me.ButtonSaveSettings)
        Me.TabPage1.Controls.Add(Me.gbox12)
        Me.TabPage1.Controls.Add(Me.gbox2)
        Me.TabPage1.Controls.Add(Me.gbox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Device 1/2  "
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Lime
        Me.Panel2.Controls.Add(Me.btncreate)
        Me.Panel2.Location = New System.Drawing.Point(951, 12)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(90, 40)
        Me.Panel2.TabIndex = 126
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Red
        Me.Panel1.Controls.Add(Me.ButtonReset)
        Me.Panel1.Location = New System.Drawing.Point(951, 57)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(90, 40)
        Me.Panel1.TabIndex = 125
        '
        'Label133
        '
        Me.Label133.AutoSize = True
        Me.Label133.Location = New System.Drawing.Point(968, 115)
        Me.Label133.Name = "Label133"
        Me.Label133.Size = New System.Drawing.Size(61, 13)
        Me.Label133.TabIndex = 508
        Me.Label133.Text = "Instruments"
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.btncreate3)
        Me.GroupBox9.Controls.Add(Me.Panel4)
        Me.GroupBox9.Controls.Add(Me.ProfDev2_8)
        Me.GroupBox9.Controls.Add(Me.ProfDev2_7)
        Me.GroupBox9.Controls.Add(Me.ProfDev2_6)
        Me.GroupBox9.Controls.Add(Me.ProfDev2_5)
        Me.GroupBox9.Controls.Add(Me.ProfDev2_4)
        Me.GroupBox9.Controls.Add(Me.Label15)
        Me.GroupBox9.Controls.Add(Me.ProfDev2_3)
        Me.GroupBox9.Controls.Add(Me.ProfDev2_2)
        Me.GroupBox9.Controls.Add(Me.ProfDev2_1)
        Me.GroupBox9.Controls.Add(Me.Label18)
        Me.GroupBox9.Controls.Add(Me.IODeviceLabel2)
        Me.GroupBox9.Controls.Add(Me.lstIntf2)
        Me.GroupBox9.Controls.Add(Me.Label4)
        Me.GroupBox9.Controls.Add(Me.txtname2)
        Me.GroupBox9.Controls.Add(Me.txtaddr2)
        Me.GroupBox9.Controls.Add(Me.Label2)
        Me.GroupBox9.Location = New System.Drawing.Point(477, 3)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(465, 123)
        Me.GroupBox9.TabIndex = 121
        Me.GroupBox9.TabStop = False
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.Lime
        Me.Panel4.Location = New System.Drawing.Point(323, 37)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(92, 42)
        Me.Panel4.TabIndex = 127
        '
        'ProfDev2_8
        '
        Me.ProfDev2_8.AutoSize = True
        Me.ProfDev2_8.Location = New System.Drawing.Point(209, 16)
        Me.ProfDev2_8.Name = "ProfDev2_8"
        Me.ProfDev2_8.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev2_8.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev2_8.TabIndex = 530
        Me.ProfDev2_8.UseVisualStyleBackColor = True
        '
        'ProfDev2_7
        '
        Me.ProfDev2_7.AutoSize = True
        Me.ProfDev2_7.Location = New System.Drawing.Point(190, 16)
        Me.ProfDev2_7.Name = "ProfDev2_7"
        Me.ProfDev2_7.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev2_7.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev2_7.TabIndex = 529
        Me.ProfDev2_7.UseVisualStyleBackColor = True
        '
        'ProfDev2_6
        '
        Me.ProfDev2_6.AutoSize = True
        Me.ProfDev2_6.Location = New System.Drawing.Point(171, 16)
        Me.ProfDev2_6.Name = "ProfDev2_6"
        Me.ProfDev2_6.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev2_6.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev2_6.TabIndex = 528
        Me.ProfDev2_6.UseVisualStyleBackColor = True
        '
        'ProfDev2_5
        '
        Me.ProfDev2_5.AutoSize = True
        Me.ProfDev2_5.Location = New System.Drawing.Point(152, 16)
        Me.ProfDev2_5.Name = "ProfDev2_5"
        Me.ProfDev2_5.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev2_5.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev2_5.TabIndex = 527
        Me.ProfDev2_5.UseVisualStyleBackColor = True
        '
        'ProfDev2_4
        '
        Me.ProfDev2_4.AutoSize = True
        Me.ProfDev2_4.Location = New System.Drawing.Point(133, 16)
        Me.ProfDev2_4.Name = "ProfDev2_4"
        Me.ProfDev2_4.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev2_4.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev2_4.TabIndex = 526
        Me.ProfDev2_4.UseVisualStyleBackColor = True
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(23, 16)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(52, 13)
        Me.Label15.TabIndex = 525
        Me.Label15.Text = "PROFILE"
        '
        'ProfDev2_3
        '
        Me.ProfDev2_3.AutoSize = True
        Me.ProfDev2_3.Location = New System.Drawing.Point(114, 16)
        Me.ProfDev2_3.Name = "ProfDev2_3"
        Me.ProfDev2_3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev2_3.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev2_3.TabIndex = 524
        Me.ProfDev2_3.UseVisualStyleBackColor = True
        '
        'ProfDev2_2
        '
        Me.ProfDev2_2.AutoSize = True
        Me.ProfDev2_2.Location = New System.Drawing.Point(95, 16)
        Me.ProfDev2_2.Name = "ProfDev2_2"
        Me.ProfDev2_2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev2_2.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev2_2.TabIndex = 523
        Me.ProfDev2_2.UseVisualStyleBackColor = True
        '
        'ProfDev2_1
        '
        Me.ProfDev2_1.AutoSize = True
        Me.ProfDev2_1.Location = New System.Drawing.Point(76, 16)
        Me.ProfDev2_1.Name = "ProfDev2_1"
        Me.ProfDev2_1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev2_1.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev2_1.TabIndex = 522
        Me.ProfDev2_1.UseVisualStyleBackColor = True
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(9, 93)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(67, 13)
        Me.Label18.TabIndex = 520
        Me.Label18.Text = "INTERFACE"
        '
        'IODeviceLabel2
        '
        Me.IODeviceLabel2.AutoSize = True
        Me.IODeviceLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IODeviceLabel2.Location = New System.Drawing.Point(376, 10)
        Me.IODeviceLabel2.Name = "IODeviceLabel2"
        Me.IODeviceLabel2.Size = New System.Drawing.Size(86, 13)
        Me.IODeviceLabel2.TabIndex = 514
        Me.IODeviceLabel2.Text = "I/O DEVICE 2"
        '
        'lstIntf2
        '
        Me.lstIntf2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstIntf2.FormattingEnabled = True
        Me.lstIntf2.Location = New System.Drawing.Point(77, 89)
        Me.lstIntf2.Name = "lstIntf2"
        Me.lstIntf2.Size = New System.Drawing.Size(147, 21)
        Me.lstIntf2.TabIndex = 519
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 67)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 13)
        Me.Label4.TabIndex = 518
        Me.Label4.Text = "ADDRESS"
        '
        'txtname2
        '
        Me.txtname2.Location = New System.Drawing.Point(77, 38)
        Me.txtname2.Name = "txtname2"
        Me.txtname2.Size = New System.Drawing.Size(240, 20)
        Me.txtname2.TabIndex = 517
        Me.txtname2.Text = "KEITHLEY"
        '
        'txtaddr2
        '
        Me.txtaddr2.Location = New System.Drawing.Point(77, 63)
        Me.txtaddr2.Name = "txtaddr2"
        Me.txtaddr2.Size = New System.Drawing.Size(240, 20)
        Me.txtaddr2.TabIndex = 516
        Me.txtaddr2.Text = "GPIB1::23::INSTR"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(20, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 13)
        Me.Label2.TabIndex = 515
        Me.Label2.Text = "DEVICE 2"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.btncreate2)
        Me.GroupBox8.Controls.Add(Me.Panel3)
        Me.GroupBox8.Controls.Add(Me.ProfDev1_8)
        Me.GroupBox8.Controls.Add(Me.ProfDev1_7)
        Me.GroupBox8.Controls.Add(Me.ProfDev1_6)
        Me.GroupBox8.Controls.Add(Me.ProfDev1_5)
        Me.GroupBox8.Controls.Add(Me.ProfDev1_4)
        Me.GroupBox8.Controls.Add(Me.IODeviceLabel1)
        Me.GroupBox8.Controls.Add(Me.Label16)
        Me.GroupBox8.Controls.Add(Me.ProfDev1_3)
        Me.GroupBox8.Controls.Add(Me.ProfDev1_2)
        Me.GroupBox8.Controls.Add(Me.ProfDev1_1)
        Me.GroupBox8.Controls.Add(Me.Label17)
        Me.GroupBox8.Controls.Add(Me.lstIntf1)
        Me.GroupBox8.Controls.Add(Me.Label3)
        Me.GroupBox8.Controls.Add(Me.txtaddr1)
        Me.GroupBox8.Controls.Add(Me.txtname1)
        Me.GroupBox8.Controls.Add(Me.Label1)
        Me.GroupBox8.Location = New System.Drawing.Point(6, 3)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(466, 123)
        Me.GroupBox8.TabIndex = 493
        Me.GroupBox8.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Lime
        Me.Panel3.Location = New System.Drawing.Point(323, 39)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(92, 42)
        Me.Panel3.TabIndex = 127
        '
        'ProfDev1_8
        '
        Me.ProfDev1_8.AutoSize = True
        Me.ProfDev1_8.Location = New System.Drawing.Point(208, 19)
        Me.ProfDev1_8.Name = "ProfDev1_8"
        Me.ProfDev1_8.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev1_8.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev1_8.TabIndex = 509
        Me.ProfDev1_8.UseVisualStyleBackColor = True
        '
        'ProfDev1_7
        '
        Me.ProfDev1_7.AutoSize = True
        Me.ProfDev1_7.Location = New System.Drawing.Point(189, 19)
        Me.ProfDev1_7.Name = "ProfDev1_7"
        Me.ProfDev1_7.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev1_7.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev1_7.TabIndex = 508
        Me.ProfDev1_7.UseVisualStyleBackColor = True
        '
        'ProfDev1_6
        '
        Me.ProfDev1_6.AutoSize = True
        Me.ProfDev1_6.Location = New System.Drawing.Point(170, 19)
        Me.ProfDev1_6.Name = "ProfDev1_6"
        Me.ProfDev1_6.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev1_6.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev1_6.TabIndex = 507
        Me.ProfDev1_6.UseVisualStyleBackColor = True
        '
        'ProfDev1_5
        '
        Me.ProfDev1_5.AutoSize = True
        Me.ProfDev1_5.Location = New System.Drawing.Point(151, 19)
        Me.ProfDev1_5.Name = "ProfDev1_5"
        Me.ProfDev1_5.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev1_5.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev1_5.TabIndex = 506
        Me.ProfDev1_5.UseVisualStyleBackColor = True
        '
        'ProfDev1_4
        '
        Me.ProfDev1_4.AutoSize = True
        Me.ProfDev1_4.Location = New System.Drawing.Point(132, 19)
        Me.ProfDev1_4.Name = "ProfDev1_4"
        Me.ProfDev1_4.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev1_4.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev1_4.TabIndex = 505
        Me.ProfDev1_4.UseVisualStyleBackColor = True
        '
        'IODeviceLabel1
        '
        Me.IODeviceLabel1.AutoSize = True
        Me.IODeviceLabel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IODeviceLabel1.Location = New System.Drawing.Point(377, 10)
        Me.IODeviceLabel1.Name = "IODeviceLabel1"
        Me.IODeviceLabel1.Size = New System.Drawing.Size(86, 13)
        Me.IODeviceLabel1.TabIndex = 499
        Me.IODeviceLabel1.Text = "I/O DEVICE 1"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(22, 19)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(52, 13)
        Me.Label16.TabIndex = 504
        Me.Label16.Text = "PROFILE"
        '
        'ProfDev1_3
        '
        Me.ProfDev1_3.AutoSize = True
        Me.ProfDev1_3.Location = New System.Drawing.Point(113, 19)
        Me.ProfDev1_3.Name = "ProfDev1_3"
        Me.ProfDev1_3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev1_3.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev1_3.TabIndex = 503
        Me.ProfDev1_3.UseVisualStyleBackColor = True
        '
        'ProfDev1_2
        '
        Me.ProfDev1_2.AutoSize = True
        Me.ProfDev1_2.Location = New System.Drawing.Point(94, 19)
        Me.ProfDev1_2.Name = "ProfDev1_2"
        Me.ProfDev1_2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev1_2.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev1_2.TabIndex = 502
        Me.ProfDev1_2.UseVisualStyleBackColor = True
        '
        'ProfDev1_1
        '
        Me.ProfDev1_1.AutoSize = True
        Me.ProfDev1_1.Location = New System.Drawing.Point(75, 19)
        Me.ProfDev1_1.Name = "ProfDev1_1"
        Me.ProfDev1_1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProfDev1_1.Size = New System.Drawing.Size(15, 14)
        Me.ProfDev1_1.TabIndex = 501
        Me.ProfDev1_1.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(7, 96)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(67, 13)
        Me.Label17.TabIndex = 498
        Me.Label17.Text = "INTERFACE"
        '
        'lstIntf1
        '
        Me.lstIntf1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstIntf1.FormattingEnabled = True
        Me.lstIntf1.Location = New System.Drawing.Point(76, 91)
        Me.lstIntf1.Name = "lstIntf1"
        Me.lstIntf1.Size = New System.Drawing.Size(147, 21)
        Me.lstIntf1.TabIndex = 497
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(15, 69)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 13)
        Me.Label3.TabIndex = 496
        Me.Label3.Text = "ADDRESS"
        '
        'txtaddr1
        '
        Me.txtaddr1.Location = New System.Drawing.Point(76, 65)
        Me.txtaddr1.Name = "txtaddr1"
        Me.txtaddr1.Size = New System.Drawing.Size(240, 20)
        Me.txtaddr1.TabIndex = 495
        Me.txtaddr1.Text = "GPIB1::22::INSTR"
        '
        'txtname1
        '
        Me.txtname1.Location = New System.Drawing.Point(76, 40)
        Me.txtname1.Name = "txtname1"
        Me.txtname1.Size = New System.Drawing.Size(240, 20)
        Me.txtname1.TabIndex = 494
        Me.txtname1.Text = "HP3458A"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(19, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 493
        Me.Label1.Text = "DEVICE 1"
        '
        'Label303
        '
        Me.Label303.AutoSize = True
        Me.Label303.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label303.Location = New System.Drawing.Point(11, 579)
        Me.Label303.Name = "Label303"
        Me.Label303.Size = New System.Drawing.Size(115, 12)
        Me.Label303.TabIndex = 559
        Me.Label303.Text = "(DAYS:HRS:MINS:SECS)"
        '
        'Label302
        '
        Me.Label302.AutoSize = True
        Me.Label302.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label302.Location = New System.Drawing.Point(11, 561)
        Me.Label302.Name = "Label302"
        Me.Label302.Size = New System.Drawing.Size(129, 15)
        Me.Label302.TabIndex = 114
        Me.Label302.Text = "GPIB Running Time  ="
        '
        'RunningTimeLogging
        '
        Me.RunningTimeLogging.AutoSize = True
        Me.RunningTimeLogging.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunningTimeLogging.Location = New System.Drawing.Point(141, 561)
        Me.RunningTimeLogging.Name = "RunningTimeLogging"
        Me.RunningTimeLogging.Size = New System.Drawing.Size(72, 16)
        Me.RunningTimeLogging.TabIndex = 114
        Me.RunningTimeLogging.Text = "00:00:00:00"
        '
        'Label56
        '
        Me.Label56.AutoSize = True
        Me.Label56.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label56.Location = New System.Drawing.Point(912, 579)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(134, 16)
        Me.Label56.TabIndex = 89
        Me.Label56.Text = "www.ianjohnston.com"
        '
        'PictureBox5
        '
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(952, 475)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(95, 110)
        Me.PictureBox5.TabIndex = 79
        Me.PictureBox5.TabStop = False
        '
        'gbox12
        '
        Me.gbox12.Controls.Add(Me.Label67)
        Me.gbox12.Controls.Add(Me.ButtonDev12Run)
        Me.gbox12.Controls.Add(Me.Label66)
        Me.gbox12.Controls.Add(Me.Dev12SampleRate)
        Me.gbox12.Enabled = False
        Me.gbox12.Location = New System.Drawing.Point(323, 549)
        Me.gbox12.Name = "gbox12"
        Me.gbox12.Size = New System.Drawing.Size(352, 48)
        Me.gbox12.TabIndex = 76
        Me.gbox12.TabStop = False
        '
        'Label67
        '
        Me.Label67.AutoSize = True
        Me.Label67.Location = New System.Drawing.Point(5, 20)
        Me.Label67.Name = "Label67"
        Me.Label67.Size = New System.Drawing.Size(76, 13)
        Me.Label67.TabIndex = 78
        Me.Label67.Text = "DEVICE 1 && 2:"
        '
        'Label66
        '
        Me.Label66.AutoSize = True
        Me.Label66.Location = New System.Drawing.Point(262, 19)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(85, 13)
        Me.Label66.TabIndex = 77
        Me.Label66.Text = "- PERIOD (secs)"
        '
        'Dev12SampleRate
        '
        Me.Dev12SampleRate.Location = New System.Drawing.Point(218, 15)
        Me.Dev12SampleRate.Name = "Dev12SampleRate"
        Me.Dev12SampleRate.Size = New System.Drawing.Size(41, 20)
        Me.Dev12SampleRate.TabIndex = 76
        Me.Dev12SampleRate.Text = "5"
        '
        'gbox2
        '
        Me.gbox2.Controls.Add(Me.CommandStart2)
        Me.gbox2.Controls.Add(Me.Dev2SendQuery)
        Me.gbox2.Controls.Add(Me.ButtonDev2PreRun)
        Me.gbox2.Controls.Add(Me.Dev2INTb)
        Me.gbox2.Controls.Add(Me.Dev2TextResponse)
        Me.gbox2.Controls.Add(Me.Label221)
        Me.gbox2.Controls.Add(Me.Dev2DecimalNumDPs)
        Me.gbox2.Controls.Add(Me.Dev2Regex)
        Me.gbox2.Controls.Add(Me.Dev2IntEnable)
        Me.gbox2.Controls.Add(Me.Label299)
        Me.gbox2.Controls.Add(Me.Label293)
        Me.gbox2.Controls.Add(Me.Label300)
        Me.gbox2.Controls.Add(Me.Label253)
        Me.gbox2.Controls.Add(Me.Dev2pauseDurationInSeconds)
        Me.gbox2.Controls.Add(Me.Dev2delayop)
        Me.gbox2.Controls.Add(Me.Label251)
        Me.gbox2.Controls.Add(Me.Label297)
        Me.gbox2.Controls.Add(Me.Mult1000Dev2)
        Me.gbox2.Controls.Add(Me.Dev2runStopwatchEveryInMins)
        Me.gbox2.Controls.Add(Me.Dev2Timeout)
        Me.gbox2.Controls.Add(Me.txtq2d)
        Me.gbox2.Controls.Add(Me.Dev2INT)
        Me.gbox2.Controls.Add(Me.Dev2K2001isolatedataCHAR)
        Me.gbox2.Controls.Add(Me.Dev2K2001isolatedata)
        Me.gbox2.Controls.Add(Me.Label191)
        Me.gbox2.Controls.Add(Me.Dev2TerminatorEnable2)
        Me.gbox2.Controls.Add(Me.Dev23457Aseven)
        Me.gbox2.Controls.Add(Me.Div1000Dev2)
        Me.gbox2.Controls.Add(Me.Dev2removeletters)
        Me.gbox2.Controls.Add(Me.CheckBoxSendBlockingDev2)
        Me.gbox2.Controls.Add(Me.Label101)
        Me.gbox2.Controls.Add(Me.Dev2PollingEnable)
        Me.gbox2.Controls.Add(Me.Dev2TerminatorEnable)
        Me.gbox2.Controls.Add(Me.txtr2a_disp)
        Me.gbox2.Controls.Add(Me.Dev2STBMask)
        Me.gbox2.Controls.Add(Me.btnq2b)
        Me.gbox2.Controls.Add(Me.btns2c)
        Me.gbox2.Controls.Add(Me.txtq2c)
        Me.gbox2.Controls.Add(Me.btnq2a)
        Me.gbox2.Controls.Add(Me.txtq2b)
        Me.gbox2.Controls.Add(Me.txtr2b)
        Me.gbox2.Controls.Add(Me.txtq2a)
        Me.gbox2.Controls.Add(Me.txtr2a)
        Me.gbox2.Controls.Add(Me.Label45)
        Me.gbox2.Controls.Add(Me.Label54)
        Me.gbox2.Controls.Add(Me.CommandStart2run)
        Me.gbox2.Controls.Add(Me.Label37)
        Me.gbox2.Controls.Add(Me.CommandStop2)
        Me.gbox2.Controls.Add(Me.Label36)
        Me.gbox2.Controls.Add(Me.Dev2Samples)
        Me.gbox2.Controls.Add(Me.Label34)
        Me.gbox2.Controls.Add(Me.IgnoreErrors2)
        Me.gbox2.Controls.Add(Me.Label30)
        Me.gbox2.Controls.Add(Me.Dev2SampleRate)
        Me.gbox2.Controls.Add(Me.ButtonDev2Run)
        Me.gbox2.Controls.Add(Me.Label13)
        Me.gbox2.Controls.Add(Me.Label14)
        Me.gbox2.Controls.Add(Me.Label11)
        Me.gbox2.Controls.Add(Me.txtr2astat)
        Me.gbox2.Controls.Add(Me.Label12)
        Me.gbox2.Enabled = False
        Me.gbox2.Location = New System.Drawing.Point(477, 118)
        Me.gbox2.Name = "gbox2"
        Me.gbox2.Size = New System.Drawing.Size(465, 431)
        Me.gbox2.TabIndex = 69
        Me.gbox2.TabStop = False
        '
        'Dev2INTb
        '
        Me.Dev2INTb.AutoSize = True
        Me.Dev2INTb.Location = New System.Drawing.Point(10, 353)
        Me.Dev2INTb.Name = "Dev2INTb"
        Me.Dev2INTb.Size = New System.Drawing.Size(31, 13)
        Me.Dev2INTb.TabIndex = 122
        Me.Dev2INTb.Text = "CMD"
        '
        'Label221
        '
        Me.Label221.AutoSize = True
        Me.Label221.Location = New System.Drawing.Point(160, 225)
        Me.Label221.Name = "Label221"
        Me.Label221.Size = New System.Drawing.Size(75, 13)
        Me.Label221.TabIndex = 120
        Me.Label221.Text = "Max. No. DP's"
        '
        'Label299
        '
        Me.Label299.AutoSize = True
        Me.Label299.Location = New System.Drawing.Point(177, 368)
        Me.Label299.Name = "Label299"
        Me.Label299.Size = New System.Drawing.Size(35, 13)
        Me.Label299.TabIndex = 114
        Me.Label299.Text = "(secs)"
        '
        'Label293
        '
        Me.Label293.AutoSize = True
        Me.Label293.Location = New System.Drawing.Point(172, 356)
        Me.Label293.Name = "Label293"
        Me.Label293.Size = New System.Drawing.Size(47, 13)
        Me.Label293.TabIndex = 117
        Me.Label293.Text = "Duration"
        '
        'Label300
        '
        Me.Label300.AutoSize = True
        Me.Label300.Location = New System.Drawing.Point(225, 368)
        Me.Label300.Name = "Label300"
        Me.Label300.Size = New System.Drawing.Size(34, 13)
        Me.Label300.TabIndex = 113
        Me.Label300.Text = "(mins)"
        '
        'Dev2delayop
        '
        Me.Dev2delayop.Location = New System.Drawing.Point(423, 82)
        Me.Dev2delayop.Name = "Dev2delayop"
        Me.Dev2delayop.Size = New System.Drawing.Size(37, 20)
        Me.Dev2delayop.TabIndex = 104
        Me.Dev2delayop.Text = "1"
        Me.Dev2delayop.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label297
        '
        Me.Label297.AutoSize = True
        Me.Label297.Location = New System.Drawing.Point(223, 356)
        Me.Label297.Name = "Label297"
        Me.Label297.Size = New System.Drawing.Size(37, 13)
        Me.Label297.TabIndex = 114
        Me.Label297.Text = "Period"
        '
        'Dev2Timeout
        '
        Me.Dev2Timeout.Location = New System.Drawing.Point(423, 59)
        Me.Dev2Timeout.Name = "Dev2Timeout"
        Me.Dev2Timeout.Size = New System.Drawing.Size(37, 20)
        Me.Dev2Timeout.TabIndex = 102
        Me.Dev2Timeout.Text = "5000"
        Me.Dev2Timeout.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Dev2INT
        '
        Me.Dev2INT.AutoSize = True
        Me.Dev2INT.Location = New System.Drawing.Point(10, 340)
        Me.Dev2INT.Name = "Dev2INT"
        Me.Dev2INT.Size = New System.Drawing.Size(50, 13)
        Me.Dev2INT.TabIndex = 112
        Me.Dev2INT.Text = "INTRPT."
        '
        'Dev2K2001isolatedataCHAR
        '
        Me.Dev2K2001isolatedataCHAR.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev2K2001isolatedataCHAR.Location = New System.Drawing.Point(269, 42)
        Me.Dev2K2001isolatedataCHAR.Name = "Dev2K2001isolatedataCHAR"
        Me.Dev2K2001isolatedataCHAR.Size = New System.Drawing.Size(20, 18)
        Me.Dev2K2001isolatedataCHAR.TabIndex = 99
        Me.Dev2K2001isolatedataCHAR.Text = "N"
        Me.Dev2K2001isolatedataCHAR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label191
        '
        Me.Label191.AutoSize = True
        Me.Label191.Location = New System.Drawing.Point(7, 225)
        Me.Label191.Name = "Label191"
        Me.Label191.Size = New System.Drawing.Size(54, 13)
        Me.Label191.TabIndex = 97
        Me.Label191.Text = "DECIMAL"
        '
        'Label101
        '
        Me.Label101.AutoSize = True
        Me.Label101.Location = New System.Drawing.Point(39, 16)
        Me.Label101.Name = "Label101"
        Me.Label101.Size = New System.Drawing.Size(56, 13)
        Me.Label101.TabIndex = 94
        Me.Label101.Text = "STB mask"
        '
        'Dev2PollingEnable
        '
        Me.Dev2PollingEnable.AutoSize = True
        Me.Dev2PollingEnable.Checked = True
        Me.Dev2PollingEnable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Dev2PollingEnable.Location = New System.Drawing.Point(8, 33)
        Me.Dev2PollingEnable.Name = "Dev2PollingEnable"
        Me.Dev2PollingEnable.Size = New System.Drawing.Size(93, 17)
        Me.Dev2PollingEnable.TabIndex = 92
        Me.Dev2PollingEnable.Text = "Enable Polling"
        Me.Dev2PollingEnable.UseVisualStyleBackColor = True
        '
        'txtr2a_disp
        '
        Me.txtr2a_disp.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtr2a_disp.ForeColor = System.Drawing.Color.Lime
        Me.txtr2a_disp.Location = New System.Drawing.Point(10, 240)
        Me.txtr2a_disp.Name = "txtr2a_disp"
        Me.txtr2a_disp.ReadOnly = True
        Me.txtr2a_disp.Size = New System.Drawing.Size(248, 29)
        Me.txtr2a_disp.TabIndex = 81
        Me.txtr2a_disp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Dev2STBMask
        '
        Me.Dev2STBMask.Location = New System.Drawing.Point(8, 11)
        Me.Dev2STBMask.Name = "Dev2STBMask"
        Me.Dev2STBMask.Size = New System.Drawing.Size(27, 20)
        Me.Dev2STBMask.TabIndex = 93
        Me.Dev2STBMask.Text = "16"
        Me.Dev2STBMask.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtq2c
        '
        Me.txtq2c.Location = New System.Drawing.Point(72, 195)
        Me.txtq2c.Name = "txtq2c"
        Me.txtq2c.Size = New System.Drawing.Size(186, 20)
        Me.txtq2c.TabIndex = 68
        '
        'txtq2b
        '
        Me.txtq2b.Location = New System.Drawing.Point(72, 109)
        Me.txtq2b.Name = "txtq2b"
        Me.txtq2b.Size = New System.Drawing.Size(186, 20)
        Me.txtq2b.TabIndex = 16
        '
        'txtr2b
        '
        Me.txtr2b.Location = New System.Drawing.Point(72, 128)
        Me.txtr2b.Name = "txtr2b"
        Me.txtr2b.ReadOnly = True
        Me.txtr2b.Size = New System.Drawing.Size(186, 20)
        Me.txtr2b.TabIndex = 15
        '
        'txtq2a
        '
        Me.txtq2a.Location = New System.Drawing.Point(72, 152)
        Me.txtq2a.Name = "txtq2a"
        Me.txtq2a.Size = New System.Drawing.Size(186, 20)
        Me.txtq2a.TabIndex = 14
        '
        'txtr2a
        '
        Me.txtr2a.Location = New System.Drawing.Point(72, 171)
        Me.txtr2a.Name = "txtr2a"
        Me.txtr2a.ReadOnly = True
        Me.txtr2a.Size = New System.Drawing.Size(186, 20)
        Me.txtr2a.TabIndex = 13
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Location = New System.Drawing.Point(7, 199)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(63, 13)
        Me.Label45.TabIndex = 69
        Me.Label45.Text = "COMMAND"
        '
        'Label54
        '
        Me.Label54.AutoSize = True
        Me.Label54.Location = New System.Drawing.Point(265, 339)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(48, 13)
        Me.Label54.TabIndex = 68
        Me.Label54.Text = "AT RUN"
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(265, 376)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(53, 13)
        Me.Label37.TabIndex = 52
        Me.Label37.Text = "AT STOP"
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(263, 226)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(56, 13)
        Me.Label36.TabIndex = 52
        Me.Label36.Text = "PRE RUN"
        '
        'Dev2Samples
        '
        Me.Dev2Samples.AutoSize = True
        Me.Dev2Samples.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev2Samples.Location = New System.Drawing.Point(132, 412)
        Me.Dev2Samples.Name = "Dev2Samples"
        Me.Dev2Samples.Size = New System.Drawing.Size(14, 16)
        Me.Dev2Samples.TabIndex = 49
        Me.Dev2Samples.Text = "0"
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(183, 414)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(57, 13)
        Me.Label34.TabIndex = 48
        Me.Label34.Text = "SAMPLES"
        '
        'IgnoreErrors2
        '
        Me.IgnoreErrors2.AutoSize = True
        Me.IgnoreErrors2.Checked = True
        Me.IgnoreErrors2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.IgnoreErrors2.Location = New System.Drawing.Point(380, 10)
        Me.IgnoreErrors2.Name = "IgnoreErrors2"
        Me.IgnoreErrors2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.IgnoreErrors2.Size = New System.Drawing.Size(81, 17)
        Me.IgnoreErrors2.TabIndex = 64
        Me.IgnoreErrors2.Text = "Abort Errors"
        Me.IgnoreErrors2.UseVisualStyleBackColor = True
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(177, 395)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(85, 13)
        Me.Label30.TabIndex = 23
        Me.Label30.Text = "- PERIOD (secs)"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(4, 132)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(66, 13)
        Me.Label13.TabIndex = 21
        Me.Label13.Text = "RESPONSE"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(7, 112)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(63, 13)
        Me.Label14.TabIndex = 20
        Me.Label14.Text = "COMMAND"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(4, 175)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(66, 13)
        Me.Label11.TabIndex = 15
        Me.Label11.Text = "RESPONSE"
        '
        'txtr2astat
        '
        Me.txtr2astat.Location = New System.Drawing.Point(10, 274)
        Me.txtr2astat.Multiline = True
        Me.txtr2astat.Name = "txtr2astat"
        Me.txtr2astat.ReadOnly = True
        Me.txtr2astat.Size = New System.Drawing.Size(248, 58)
        Me.txtr2astat.TabIndex = 19
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(7, 155)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(63, 13)
        Me.Label12.TabIndex = 14
        Me.Label12.Text = "COMMAND"
        '
        'gbox1
        '
        Me.gbox1.Controls.Add(Me.CommandStart1)
        Me.gbox1.Controls.Add(Me.Dev1SendQuery)
        Me.gbox1.Controls.Add(Me.ButtonDev1PreRun)
        Me.gbox1.Controls.Add(Me.Dev1INTb)
        Me.gbox1.Controls.Add(Me.Dev1TextResponse)
        Me.gbox1.Controls.Add(Me.Label220)
        Me.gbox1.Controls.Add(Me.Dev1DecimalNumDPs)
        Me.gbox1.Controls.Add(Me.Dev1Regex)
        Me.gbox1.Controls.Add(Me.Dev1IntEnable)
        Me.gbox1.Controls.Add(Me.Label294)
        Me.gbox1.Controls.Add(Me.Label295)
        Me.gbox1.Controls.Add(Me.Label296)
        Me.gbox1.Controls.Add(Me.Dev1pauseDurationInSeconds)
        Me.gbox1.Controls.Add(Me.Label292)
        Me.gbox1.Controls.Add(Me.Dev1runStopwatchEveryInMins)
        Me.gbox1.Controls.Add(Me.txtq1d)
        Me.gbox1.Controls.Add(Me.Dev1INT)
        Me.gbox1.Controls.Add(Me.Label252)
        Me.gbox1.Controls.Add(Me.Dev1delayop)
        Me.gbox1.Controls.Add(Me.Label250)
        Me.gbox1.Controls.Add(Me.Dev1Timeout)
        Me.gbox1.Controls.Add(Me.Mult1000Dev1)
        Me.gbox1.Controls.Add(Me.Dev1K2001isolatedataCHAR)
        Me.gbox1.Controls.Add(Me.Dev1K2001isolatedata)
        Me.gbox1.Controls.Add(Me.Label190)
        Me.gbox1.Controls.Add(Me.Dev1TerminatorEnable2)
        Me.gbox1.Controls.Add(Me.Dev13457Aseven)
        Me.gbox1.Controls.Add(Me.Div1000Dev1)
        Me.gbox1.Controls.Add(Me.Dev1removeletters)
        Me.gbox1.Controls.Add(Me.CheckBoxSendBlockingDev1)
        Me.gbox1.Controls.Add(Me.Dev1STBMask)
        Me.gbox1.Controls.Add(Me.Label99)
        Me.gbox1.Controls.Add(Me.txtr1a_disp)
        Me.gbox1.Controls.Add(Me.btnq1b)
        Me.gbox1.Controls.Add(Me.txtq1b)
        Me.gbox1.Controls.Add(Me.btns1c)
        Me.gbox1.Controls.Add(Me.Dev1PollingEnable)
        Me.gbox1.Controls.Add(Me.txtq1c)
        Me.gbox1.Controls.Add(Me.btnq1a)
        Me.gbox1.Controls.Add(Me.Dev1TerminatorEnable)
        Me.gbox1.Controls.Add(Me.txtr1a)
        Me.gbox1.Controls.Add(Me.txtq1a)
        Me.gbox1.Controls.Add(Me.txtr1b)
        Me.gbox1.Controls.Add(Me.Label43)
        Me.gbox1.Controls.Add(Me.CommandStart1run)
        Me.gbox1.Controls.Add(Me.Label53)
        Me.gbox1.Controls.Add(Me.Label35)
        Me.gbox1.Controls.Add(Me.Label32)
        Me.gbox1.Controls.Add(Me.CommandStop1)
        Me.gbox1.Controls.Add(Me.Dev1Samples)
        Me.gbox1.Controls.Add(Me.Label31)
        Me.gbox1.Controls.Add(Me.Label33)
        Me.gbox1.Controls.Add(Me.ButtonDev1Run)
        Me.gbox1.Controls.Add(Me.IgnoreErrors1)
        Me.gbox1.Controls.Add(Me.Dev1SampleRate)
        Me.gbox1.Controls.Add(Me.Label7)
        Me.gbox1.Controls.Add(Me.Label8)
        Me.gbox1.Controls.Add(Me.txtr1astat)
        Me.gbox1.Controls.Add(Me.Label9)
        Me.gbox1.Controls.Add(Me.Label5)
        Me.gbox1.Enabled = False
        Me.gbox1.Location = New System.Drawing.Point(6, 118)
        Me.gbox1.Name = "gbox1"
        Me.gbox1.Size = New System.Drawing.Size(466, 431)
        Me.gbox1.TabIndex = 67
        Me.gbox1.TabStop = False
        '
        'Dev1INTb
        '
        Me.Dev1INTb.AutoSize = True
        Me.Dev1INTb.Location = New System.Drawing.Point(11, 353)
        Me.Dev1INTb.Name = "Dev1INTb"
        Me.Dev1INTb.Size = New System.Drawing.Size(31, 13)
        Me.Dev1INTb.TabIndex = 118
        Me.Dev1INTb.Text = "CMD"
        '
        'Label220
        '
        Me.Label220.AutoSize = True
        Me.Label220.Location = New System.Drawing.Point(161, 225)
        Me.Label220.Name = "Label220"
        Me.Label220.Size = New System.Drawing.Size(75, 13)
        Me.Label220.TabIndex = 116
        Me.Label220.Text = "Max. No. DP's"
        '
        'Label294
        '
        Me.Label294.AutoSize = True
        Me.Label294.Location = New System.Drawing.Point(178, 369)
        Me.Label294.Name = "Label294"
        Me.Label294.Size = New System.Drawing.Size(35, 13)
        Me.Label294.TabIndex = 112
        Me.Label294.Text = "(secs)"
        '
        'Label295
        '
        Me.Label295.AutoSize = True
        Me.Label295.Location = New System.Drawing.Point(225, 369)
        Me.Label295.Name = "Label295"
        Me.Label295.Size = New System.Drawing.Size(34, 13)
        Me.Label295.TabIndex = 111
        Me.Label295.Text = "(mins)"
        '
        'Label296
        '
        Me.Label296.AutoSize = True
        Me.Label296.Location = New System.Drawing.Point(173, 356)
        Me.Label296.Name = "Label296"
        Me.Label296.Size = New System.Drawing.Size(47, 13)
        Me.Label296.TabIndex = 110
        Me.Label296.Text = "Duration"
        '
        'Label292
        '
        Me.Label292.AutoSize = True
        Me.Label292.Location = New System.Drawing.Point(224, 356)
        Me.Label292.Name = "Label292"
        Me.Label292.Size = New System.Drawing.Size(37, 13)
        Me.Label292.TabIndex = 107
        Me.Label292.Text = "Period"
        '
        'Dev1INT
        '
        Me.Dev1INT.AutoSize = True
        Me.Dev1INT.Location = New System.Drawing.Point(11, 340)
        Me.Dev1INT.Name = "Dev1INT"
        Me.Dev1INT.Size = New System.Drawing.Size(50, 13)
        Me.Dev1INT.TabIndex = 105
        Me.Dev1INT.Text = "INTRPT."
        '
        'Dev1delayop
        '
        Me.Dev1delayop.Location = New System.Drawing.Point(424, 82)
        Me.Dev1delayop.Name = "Dev1delayop"
        Me.Dev1delayop.Size = New System.Drawing.Size(37, 20)
        Me.Dev1delayop.TabIndex = 102
        Me.Dev1delayop.Text = "1"
        Me.Dev1delayop.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Dev1Timeout
        '
        Me.Dev1Timeout.Location = New System.Drawing.Point(424, 59)
        Me.Dev1Timeout.Name = "Dev1Timeout"
        Me.Dev1Timeout.Size = New System.Drawing.Size(37, 20)
        Me.Dev1Timeout.TabIndex = 100
        Me.Dev1Timeout.Text = "5000"
        Me.Dev1Timeout.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Dev1K2001isolatedataCHAR
        '
        Me.Dev1K2001isolatedataCHAR.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev1K2001isolatedataCHAR.Location = New System.Drawing.Point(272, 42)
        Me.Dev1K2001isolatedataCHAR.Name = "Dev1K2001isolatedataCHAR"
        Me.Dev1K2001isolatedataCHAR.Size = New System.Drawing.Size(20, 18)
        Me.Dev1K2001isolatedataCHAR.TabIndex = 98
        Me.Dev1K2001isolatedataCHAR.Text = "N"
        Me.Dev1K2001isolatedataCHAR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label190
        '
        Me.Label190.AutoSize = True
        Me.Label190.Location = New System.Drawing.Point(6, 225)
        Me.Label190.Name = "Label190"
        Me.Label190.Size = New System.Drawing.Size(54, 13)
        Me.Label190.TabIndex = 96
        Me.Label190.Text = "DECIMAL"
        '
        'Dev1STBMask
        '
        Me.Dev1STBMask.Location = New System.Drawing.Point(8, 11)
        Me.Dev1STBMask.Name = "Dev1STBMask"
        Me.Dev1STBMask.Size = New System.Drawing.Size(27, 20)
        Me.Dev1STBMask.TabIndex = 90
        Me.Dev1STBMask.Text = "16"
        Me.Dev1STBMask.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label99
        '
        Me.Label99.AutoSize = True
        Me.Label99.Location = New System.Drawing.Point(39, 16)
        Me.Label99.Name = "Label99"
        Me.Label99.Size = New System.Drawing.Size(56, 13)
        Me.Label99.TabIndex = 91
        Me.Label99.Text = "STB mask"
        '
        'txtr1a_disp
        '
        Me.txtr1a_disp.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtr1a_disp.ForeColor = System.Drawing.Color.Lime
        Me.txtr1a_disp.Location = New System.Drawing.Point(9, 240)
        Me.txtr1a_disp.Name = "txtr1a_disp"
        Me.txtr1a_disp.ReadOnly = True
        Me.txtr1a_disp.Size = New System.Drawing.Size(250, 29)
        Me.txtr1a_disp.TabIndex = 86
        Me.txtr1a_disp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtq1b
        '
        Me.txtq1b.Location = New System.Drawing.Point(73, 109)
        Me.txtq1b.Name = "txtq1b"
        Me.txtq1b.Size = New System.Drawing.Size(186, 20)
        Me.txtq1b.TabIndex = 0
        '
        'Dev1PollingEnable
        '
        Me.Dev1PollingEnable.AutoSize = True
        Me.Dev1PollingEnable.Checked = True
        Me.Dev1PollingEnable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Dev1PollingEnable.Location = New System.Drawing.Point(8, 33)
        Me.Dev1PollingEnable.Name = "Dev1PollingEnable"
        Me.Dev1PollingEnable.Size = New System.Drawing.Size(93, 17)
        Me.Dev1PollingEnable.TabIndex = 89
        Me.Dev1PollingEnable.Text = "Enable Polling"
        Me.Dev1PollingEnable.UseVisualStyleBackColor = True
        '
        'txtq1c
        '
        Me.txtq1c.Location = New System.Drawing.Point(73, 195)
        Me.txtq1c.Name = "txtq1c"
        Me.txtq1c.Size = New System.Drawing.Size(186, 20)
        Me.txtq1c.TabIndex = 67
        '
        'txtr1a
        '
        Me.txtr1a.Location = New System.Drawing.Point(73, 171)
        Me.txtr1a.Name = "txtr1a"
        Me.txtr1a.ReadOnly = True
        Me.txtr1a.Size = New System.Drawing.Size(186, 20)
        Me.txtr1a.TabIndex = 5
        '
        'txtq1a
        '
        Me.txtq1a.Location = New System.Drawing.Point(73, 152)
        Me.txtq1a.Name = "txtq1a"
        Me.txtq1a.Size = New System.Drawing.Size(186, 20)
        Me.txtq1a.TabIndex = 4
        '
        'txtr1b
        '
        Me.txtr1b.Location = New System.Drawing.Point(73, 128)
        Me.txtr1b.Name = "txtr1b"
        Me.txtr1b.ReadOnly = True
        Me.txtr1b.Size = New System.Drawing.Size(186, 20)
        Me.txtr1b.TabIndex = 3
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(8, 199)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(63, 13)
        Me.Label43.TabIndex = 68
        Me.Label43.Text = "COMMAND"
        '
        'Label53
        '
        Me.Label53.AutoSize = True
        Me.Label53.Location = New System.Drawing.Point(268, 339)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(48, 13)
        Me.Label53.TabIndex = 65
        Me.Label53.Text = "AT RUN"
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(267, 376)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(53, 13)
        Me.Label35.TabIndex = 51
        Me.Label35.Text = "AT STOP"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(265, 226)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(56, 13)
        Me.Label32.TabIndex = 50
        Me.Label32.Text = "PRE RUN"
        '
        'Dev1Samples
        '
        Me.Dev1Samples.AutoSize = True
        Me.Dev1Samples.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev1Samples.Location = New System.Drawing.Point(129, 412)
        Me.Dev1Samples.Name = "Dev1Samples"
        Me.Dev1Samples.Size = New System.Drawing.Size(14, 16)
        Me.Dev1Samples.TabIndex = 47
        Me.Dev1Samples.Text = "0"
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(174, 395)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(85, 13)
        Me.Label31.TabIndex = 25
        Me.Label31.Text = "- PERIOD (secs)"
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(180, 414)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(57, 13)
        Me.Label33.TabIndex = 46
        Me.Label33.Text = "SAMPLES"
        '
        'IgnoreErrors1
        '
        Me.IgnoreErrors1.AutoSize = True
        Me.IgnoreErrors1.Checked = True
        Me.IgnoreErrors1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.IgnoreErrors1.Location = New System.Drawing.Point(381, 10)
        Me.IgnoreErrors1.Name = "IgnoreErrors1"
        Me.IgnoreErrors1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.IgnoreErrors1.Size = New System.Drawing.Size(81, 17)
        Me.IgnoreErrors1.TabIndex = 63
        Me.IgnoreErrors1.Text = "Abort Errors"
        Me.IgnoreErrors1.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 175)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(66, 13)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "RESPONSE"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 155)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(63, 13)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "COMMAND"
        '
        'txtr1astat
        '
        Me.txtr1astat.Location = New System.Drawing.Point(9, 274)
        Me.txtr1astat.Multiline = True
        Me.txtr1astat.Name = "txtr1astat"
        Me.txtr1astat.ReadOnly = True
        Me.txtr1astat.Size = New System.Drawing.Size(250, 58)
        Me.txtr1astat.TabIndex = 11
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 132)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(66, 13)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "RESPONSE"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "COMMAND"
        '
        'TabPage10
        '
        Me.TabPage10.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage10.Controls.Add(Me.Label227)
        Me.TabPage10.Controls.Add(Me.TextBoxDev1CMD)
        Me.TabPage10.Controls.Add(Me.CMD2clear)
        Me.TabPage10.Controls.Add(Me.CMD1clear)
        Me.TabPage10.Controls.Add(Me.CMDdev2)
        Me.TabPage10.Controls.Add(Me.CMDdev1)
        Me.TabPage10.Controls.Add(Me.Label189)
        Me.TabPage10.Controls.Add(Me.CheckBoxDev2Async)
        Me.TabPage10.Controls.Add(Me.CheckBoxDev2Query)
        Me.TabPage10.Controls.Add(Me.TextBoxDev2CMD)
        Me.TabPage10.Controls.Add(Me.Label188)
        Me.TabPage10.Controls.Add(Me.Label187)
        Me.TabPage10.Controls.Add(Me.CheckBoxDev1Async)
        Me.TabPage10.Controls.Add(Me.CheckBoxDev1Query)
        Me.TabPage10.Location = New System.Drawing.Point(4, 22)
        Me.TabPage10.Name = "TabPage10"
        Me.TabPage10.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage10.TabIndex = 9
        Me.TabPage10.Text = "Cmd Line  "
        '
        'Label227
        '
        Me.Label227.AutoSize = True
        Me.Label227.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label227.Location = New System.Drawing.Point(752, 5)
        Me.Label227.Name = "Label227"
        Me.Label227.Size = New System.Drawing.Size(267, 78)
        Me.Label227.TabIndex = 14
        Me.Label227.Text = resources.GetString("Label227.Text")
        '
        'TextBoxDev1CMD
        '
        Me.TextBoxDev1CMD.BackColor = System.Drawing.Color.White
        Me.TextBoxDev1CMD.ForeColor = System.Drawing.Color.Black
        Me.TextBoxDev1CMD.Location = New System.Drawing.Point(8, 128)
        Me.TextBoxDev1CMD.Multiline = True
        Me.TextBoxDev1CMD.Name = "TextBoxDev1CMD"
        Me.TextBoxDev1CMD.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxDev1CMD.Size = New System.Drawing.Size(491, 460)
        Me.TextBoxDev1CMD.TabIndex = 13
        '
        'CMD2clear
        '
        Me.CMD2clear.Location = New System.Drawing.Point(967, 90)
        Me.CMD2clear.Name = "CMD2clear"
        Me.CMD2clear.Size = New System.Drawing.Size(52, 23)
        Me.CMD2clear.TabIndex = 12
        Me.CMD2clear.Text = "Clear"
        Me.CMD2clear.UseVisualStyleBackColor = True
        '
        'CMD1clear
        '
        Me.CMD1clear.Location = New System.Drawing.Point(441, 90)
        Me.CMD1clear.Name = "CMD1clear"
        Me.CMD1clear.Size = New System.Drawing.Size(52, 23)
        Me.CMD1clear.TabIndex = 11
        Me.CMD1clear.Text = "Clear"
        Me.CMD1clear.UseVisualStyleBackColor = True
        '
        'CMDdev2
        '
        Me.CMDdev2.AutoSize = True
        Me.CMDdev2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CMDdev2.Location = New System.Drawing.Point(623, 55)
        Me.CMDdev2.Name = "CMDdev2"
        Me.CMDdev2.Size = New System.Drawing.Size(70, 16)
        Me.CMDdev2.TabIndex = 10
        Me.CMDdev2.Text = "#########"
        '
        'CMDdev1
        '
        Me.CMDdev1.AutoSize = True
        Me.CMDdev1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CMDdev1.Location = New System.Drawing.Point(107, 55)
        Me.CMDdev1.Name = "CMDdev1"
        Me.CMDdev1.Size = New System.Drawing.Size(63, 16)
        Me.CMDdev1.TabIndex = 9
        Me.CMDdev1.Text = "########"
        '
        'Label189
        '
        Me.Label189.AutoSize = True
        Me.Label189.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label189.Location = New System.Drawing.Point(527, 55)
        Me.Label189.Name = "Label189"
        Me.Label189.Size = New System.Drawing.Size(78, 16)
        Me.Label189.TabIndex = 8
        Me.Label189.Text = "DEVICE 2:"
        '
        'CheckBoxDev2Async
        '
        Me.CheckBoxDev2Async.AutoSize = True
        Me.CheckBoxDev2Async.Location = New System.Drawing.Point(530, 98)
        Me.CheckBoxDev2Async.Name = "CheckBoxDev2Async"
        Me.CheckBoxDev2Async.Size = New System.Drawing.Size(138, 17)
        Me.CheckBoxDev2Async.TabIndex = 7
        Me.CheckBoxDev2Async.Text = "Send Async    (no reply)"
        Me.CheckBoxDev2Async.UseVisualStyleBackColor = True
        '
        'CheckBoxDev2Query
        '
        Me.CheckBoxDev2Query.AutoSize = True
        Me.CheckBoxDev2Query.Location = New System.Drawing.Point(530, 79)
        Me.CheckBoxDev2Query.Name = "CheckBoxDev2Query"
        Me.CheckBoxDev2Query.Size = New System.Drawing.Size(170, 17)
        Me.CheckBoxDev2Query.TabIndex = 6
        Me.CheckBoxDev2Query.Text = "Query Async   (reply expected)"
        Me.CheckBoxDev2Query.UseVisualStyleBackColor = True
        '
        'TextBoxDev2CMD
        '
        Me.TextBoxDev2CMD.BackColor = System.Drawing.Color.White
        Me.TextBoxDev2CMD.ForeColor = System.Drawing.Color.Black
        Me.TextBoxDev2CMD.Location = New System.Drawing.Point(527, 128)
        Me.TextBoxDev2CMD.Multiline = True
        Me.TextBoxDev2CMD.Name = "TextBoxDev2CMD"
        Me.TextBoxDev2CMD.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxDev2CMD.Size = New System.Drawing.Size(491, 460)
        Me.TextBoxDev2CMD.TabIndex = 5
        '
        'Label188
        '
        Me.Label188.AutoSize = True
        Me.Label188.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label188.Location = New System.Drawing.Point(9, 55)
        Me.Label188.Name = "Label188"
        Me.Label188.Size = New System.Drawing.Size(78, 16)
        Me.Label188.TabIndex = 4
        Me.Label188.Text = "DEVICE 1:"
        '
        'Label187
        '
        Me.Label187.AutoSize = True
        Me.Label187.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label187.Location = New System.Drawing.Point(420, 7)
        Me.Label187.Name = "Label187"
        Me.Label187.Size = New System.Drawing.Size(192, 20)
        Me.Label187.TabIndex = 3
        Me.Label187.Text = "GPIB COMMAND LINE"
        '
        'CheckBoxDev1Async
        '
        Me.CheckBoxDev1Async.AutoSize = True
        Me.CheckBoxDev1Async.Location = New System.Drawing.Point(12, 98)
        Me.CheckBoxDev1Async.Name = "CheckBoxDev1Async"
        Me.CheckBoxDev1Async.Size = New System.Drawing.Size(138, 17)
        Me.CheckBoxDev1Async.TabIndex = 2
        Me.CheckBoxDev1Async.Text = "Send Async    (no reply)"
        Me.CheckBoxDev1Async.UseVisualStyleBackColor = True
        '
        'CheckBoxDev1Query
        '
        Me.CheckBoxDev1Query.AutoSize = True
        Me.CheckBoxDev1Query.Location = New System.Drawing.Point(12, 79)
        Me.CheckBoxDev1Query.Name = "CheckBoxDev1Query"
        Me.CheckBoxDev1Query.Size = New System.Drawing.Size(170, 17)
        Me.CheckBoxDev1Query.TabIndex = 1
        Me.CheckBoxDev1Query.Text = "Query Async   (reply expected)"
        Me.CheckBoxDev1Query.UseVisualStyleBackColor = True
        '
        'TabPage8
        '
        Me.TabPage8.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage8.Controls.Add(Me.Label169)
        Me.TabPage8.Controls.Add(Me.Label170)
        Me.TabPage8.Controls.Add(Me.Device2name)
        Me.TabPage8.Controls.Add(Me.DeviceHumidity)
        Me.TabPage8.Controls.Add(Me.Label173)
        Me.TabPage8.Controls.Add(Me.DeviceTemperature)
        Me.TabPage8.Controls.Add(Me.Label171)
        Me.TabPage8.Controls.Add(Me.Dev2Meter)
        Me.TabPage8.Controls.Add(Me.GroupBox1)
        Me.TabPage8.Controls.Add(Me.GroupBox4)
        Me.TabPage8.Controls.Add(Me.GroupBox3)
        Me.TabPage8.Location = New System.Drawing.Point(4, 22)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage8.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage8.TabIndex = 7
        Me.TabPage8.Text = "Device Meters  "
        '
        'Label169
        '
        Me.Label169.AutoSize = True
        Me.Label169.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label169.Location = New System.Drawing.Point(26, 130)
        Me.Label169.Name = "Label169"
        Me.Label169.Size = New System.Drawing.Size(115, 25)
        Me.Label169.TabIndex = 69
        Me.Label169.Text = "DEVICE 1"
        '
        'Label170
        '
        Me.Label170.AutoSize = True
        Me.Label170.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label170.Location = New System.Drawing.Point(26, 376)
        Me.Label170.Name = "Label170"
        Me.Label170.Size = New System.Drawing.Size(115, 25)
        Me.Label170.TabIndex = 71
        Me.Label170.Text = "DEVICE 2"
        '
        'Device2name
        '
        Me.Device2name.AutoSize = True
        Me.Device2name.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Device2name.Location = New System.Drawing.Point(179, 376)
        Me.Device2name.Name = "Device2name"
        Me.Device2name.Size = New System.Drawing.Size(92, 25)
        Me.Device2name.TabIndex = 73
        Me.Device2name.Text = "----------"
        '
        'DeviceHumidity
        '
        Me.DeviceHumidity.AutoSize = True
        Me.DeviceHumidity.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DeviceHumidity.Location = New System.Drawing.Point(569, 21)
        Me.DeviceHumidity.Name = "DeviceHumidity"
        Me.DeviceHumidity.Size = New System.Drawing.Size(81, 33)
        Me.DeviceHumidity.TabIndex = 77
        Me.DeviceHumidity.Text = "------"
        '
        'Label173
        '
        Me.Label173.AutoSize = True
        Me.Label173.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label173.Location = New System.Drawing.Point(365, 28)
        Me.Label173.Name = "Label173"
        Me.Label173.Size = New System.Drawing.Size(180, 25)
        Me.Label173.TabIndex = 76
        Me.Label173.Text = "HUMIDITY %RH"
        '
        'DeviceTemperature
        '
        Me.DeviceTemperature.AutoSize = True
        Me.DeviceTemperature.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DeviceTemperature.Location = New System.Drawing.Point(190, 21)
        Me.DeviceTemperature.Name = "DeviceTemperature"
        Me.DeviceTemperature.Size = New System.Drawing.Size(81, 33)
        Me.DeviceTemperature.TabIndex = 75
        Me.DeviceTemperature.Text = "------"
        '
        'Label171
        '
        Me.Label171.AutoSize = True
        Me.Label171.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label171.ForeColor = System.Drawing.Color.Black
        Me.Label171.Location = New System.Drawing.Point(26, 28)
        Me.Label171.Name = "Label171"
        Me.Label171.Size = New System.Drawing.Size(144, 25)
        Me.Label171.TabIndex = 74
        Me.Label171.Text = "TEMP. degC"
        '
        'Dev2Meter
        '
        Me.Dev2Meter.AutoSize = True
        Me.Dev2Meter.Font = New System.Drawing.Font("Microsoft Sans Serif", 110.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev2Meter.Location = New System.Drawing.Point(14, 391)
        Me.Dev2Meter.Name = "Dev2Meter"
        Me.Dev2Meter.Size = New System.Drawing.Size(805, 166)
        Me.Dev2Meter.TabIndex = 70
        Me.Dev2Meter.Text = "---------------"
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1034, 77)
        Me.GroupBox1.TabIndex = 78
        Me.GroupBox1.TabStop = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Dev2Units)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 362)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(1034, 217)
        Me.GroupBox4.TabIndex = 80
        Me.GroupBox4.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Dev1Units)
        Me.GroupBox3.Controls.Add(Me.Device1name)
        Me.GroupBox3.Controls.Add(Me.Dev1Meter)
        Me.GroupBox3.Location = New System.Drawing.Point(8, 116)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1034, 217)
        Me.GroupBox3.TabIndex = 79
        Me.GroupBox3.TabStop = False
        '
        'Device1name
        '
        Me.Device1name.AutoSize = True
        Me.Device1name.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Device1name.Location = New System.Drawing.Point(171, 14)
        Me.Device1name.Name = "Device1name"
        Me.Device1name.Size = New System.Drawing.Size(92, 25)
        Me.Device1name.TabIndex = 72
        Me.Device1name.Text = "----------"
        '
        'Dev1Meter
        '
        Me.Dev1Meter.AutoSize = True
        Me.Dev1Meter.Font = New System.Drawing.Font("Microsoft Sans Serif", 110.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dev1Meter.Location = New System.Drawing.Point(6, 29)
        Me.Dev1Meter.Name = "Dev1Meter"
        Me.Dev1Meter.Size = New System.Drawing.Size(805, 166)
        Me.Dev1Meter.TabIndex = 48
        Me.Dev1Meter.Text = "---------------"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage2.Controls.Add(Me.gboxtemphum)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Temp/Hum  "
        '
        'gboxtemphum
        '
        Me.gboxtemphum.Controls.Add(Me.Label231)
        Me.gboxtemphum.Controls.Add(Me.TextBoxTempHumSample)
        Me.gboxtemphum.Controls.Add(Me.Label230)
        Me.gboxtemphum.Controls.Add(Me.Label198)
        Me.gboxtemphum.Controls.Add(Me.CheckBoxParseLeftRight)
        Me.gboxtemphum.Controls.Add(Me.CheckBoxArithmetic)
        Me.gboxtemphum.Controls.Add(Me.CheckBoxRegex)
        Me.gboxtemphum.Controls.Add(Me.ButtonSaveTempHumSettings)
        Me.gboxtemphum.Controls.Add(Me.Label219)
        Me.gboxtemphum.Controls.Add(Me.Label218)
        Me.gboxtemphum.Controls.Add(Me.Label217)
        Me.gboxtemphum.Controls.Add(Me.Label216)
        Me.gboxtemphum.Controls.Add(Me.Label215)
        Me.gboxtemphum.Controls.Add(Me.Label214)
        Me.gboxtemphum.Controls.Add(Me.Label213)
        Me.gboxtemphum.Controls.Add(Me.OnOffLed2)
        Me.gboxtemphum.Controls.Add(Me.OnOffLed1)
        Me.gboxtemphum.Controls.Add(Me.Label212)
        Me.gboxtemphum.Controls.Add(Me.Label211)
        Me.gboxtemphum.Controls.Add(Me.Label210)
        Me.gboxtemphum.Controls.Add(Me.Label209)
        Me.gboxtemphum.Controls.Add(Me.Label208)
        Me.gboxtemphum.Controls.Add(Me.Label207)
        Me.gboxtemphum.Controls.Add(Me.TextBoxSerialPortHand)
        Me.gboxtemphum.Controls.Add(Me.Label206)
        Me.gboxtemphum.Controls.Add(Me.TextBoxSerialPortStop)
        Me.gboxtemphum.Controls.Add(Me.Label205)
        Me.gboxtemphum.Controls.Add(Me.TextBoxSerialPortParity)
        Me.gboxtemphum.Controls.Add(Me.Label204)
        Me.gboxtemphum.Controls.Add(Me.TextBoxSerialPortBits)
        Me.gboxtemphum.Controls.Add(Me.Label195)
        Me.gboxtemphum.Controls.Add(Me.TextBoxSerialPortBaud)
        Me.gboxtemphum.Controls.Add(Me.Label47)
        Me.gboxtemphum.Controls.Add(Me.Label46)
        Me.gboxtemphum.Controls.Add(Me.TextBoxHumUnits)
        Me.gboxtemphum.Controls.Add(Me.TextBoxTempUnits)
        Me.gboxtemphum.Controls.Add(Me.Label202)
        Me.gboxtemphum.Controls.Add(Me.Label203)
        Me.gboxtemphum.Controls.Add(Me.TextBoxTempArithmentic)
        Me.gboxtemphum.Controls.Add(Me.Label201)
        Me.gboxtemphum.Controls.Add(Me.TextBoxRegex)
        Me.gboxtemphum.Controls.Add(Me.Label200)
        Me.gboxtemphum.Controls.Add(Me.Label199)
        Me.gboxtemphum.Controls.Add(Me.TempFinalValue)
        Me.gboxtemphum.Controls.Add(Me.TextBoxFinalTempValue)
        Me.gboxtemphum.Controls.Add(Me.Label197)
        Me.gboxtemphum.Controls.Add(Me.Label196)
        Me.gboxtemphum.Controls.Add(Me.TextBoxParseRight)
        Me.gboxtemphum.Controls.Add(Me.TextBoxParseLeft)
        Me.gboxtemphum.Controls.Add(Me.Label194)
        Me.gboxtemphum.Controls.Add(Me.Label193)
        Me.gboxtemphum.Controls.Add(Me.TextBoxResult)
        Me.gboxtemphum.Controls.Add(Me.TextBoxProtocolInput)
        Me.gboxtemphum.Controls.Add(Me.ButtonRefreshPorts)
        Me.gboxtemphum.Controls.Add(Me.Label186)
        Me.gboxtemphum.Controls.Add(Me.Label185)
        Me.gboxtemphum.Controls.Add(Me.Label184)
        Me.gboxtemphum.Controls.Add(Me.Label183)
        Me.gboxtemphum.Controls.Add(Me.Label182)
        Me.gboxtemphum.Controls.Add(Me.Label181)
        Me.gboxtemphum.Controls.Add(Me.Label178)
        Me.gboxtemphum.Controls.Add(Me.Label164)
        Me.gboxtemphum.Controls.Add(Me.TempOffset)
        Me.gboxtemphum.Controls.Add(Me.Label23)
        Me.gboxtemphum.Controls.Add(Me.txtname3)
        Me.gboxtemphum.Controls.Add(Me.Label29)
        Me.gboxtemphum.Controls.Add(Me.Label28)
        Me.gboxtemphum.Controls.Add(Me.lstIntf3)
        Me.gboxtemphum.Controls.Add(Me.Label21)
        Me.gboxtemphum.Controls.Add(Me.LabelHumidity)
        Me.gboxtemphum.Controls.Add(Me.ButtonStart)
        Me.gboxtemphum.Controls.Add(Me.LabelTemperature)
        Me.gboxtemphum.Controls.Add(Me.ButtonEnd)
        Me.gboxtemphum.Controls.Add(Me.Label20)
        Me.gboxtemphum.Controls.Add(Me.ComboBoxPort)
        Me.gboxtemphum.Controls.Add(Me.Label19)
        Me.gboxtemphum.Enabled = False
        Me.gboxtemphum.Location = New System.Drawing.Point(6, 6)
        Me.gboxtemphum.Name = "gboxtemphum"
        Me.gboxtemphum.Size = New System.Drawing.Size(1037, 586)
        Me.gboxtemphum.TabIndex = 49
        Me.gboxtemphum.TabStop = False
        '
        'Label231
        '
        Me.Label231.AutoSize = True
        Me.Label231.Location = New System.Drawing.Point(200, 208)
        Me.Label231.Name = "Label231"
        Me.Label231.Size = New System.Drawing.Size(34, 13)
        Me.Label231.TabIndex = 711
        Me.Label231.Text = "Secs."
        '
        'Label230
        '
        Me.Label230.AutoSize = True
        Me.Label230.Location = New System.Drawing.Point(17, 208)
        Me.Label230.Name = "Label230"
        Me.Label230.Size = New System.Drawing.Size(103, 13)
        Me.Label230.TabIndex = 709
        Me.Label230.Text = "Sampling Frequency"
        '
        'Label198
        '
        Me.Label198.AutoSize = True
        Me.Label198.Location = New System.Drawing.Point(743, 165)
        Me.Label198.Name = "Label198"
        Me.Label198.Size = New System.Drawing.Size(111, 13)
        Me.Label198.TabIndex = 708
        Me.Label198.Text = "User Defined Protocol"
        '
        'CheckBoxParseLeftRight
        '
        Me.CheckBoxParseLeftRight.AutoSize = True
        Me.CheckBoxParseLeftRight.Location = New System.Drawing.Point(108, 435)
        Me.CheckBoxParseLeftRight.Name = "CheckBoxParseLeftRight"
        Me.CheckBoxParseLeftRight.Size = New System.Drawing.Size(15, 14)
        Me.CheckBoxParseLeftRight.TabIndex = 707
        Me.CheckBoxParseLeftRight.UseVisualStyleBackColor = True
        '
        'CheckBoxArithmetic
        '
        Me.CheckBoxArithmetic.AutoSize = True
        Me.CheckBoxArithmetic.Location = New System.Drawing.Point(108, 499)
        Me.CheckBoxArithmetic.Name = "CheckBoxArithmetic"
        Me.CheckBoxArithmetic.Size = New System.Drawing.Size(15, 14)
        Me.CheckBoxArithmetic.TabIndex = 706
        Me.CheckBoxArithmetic.UseVisualStyleBackColor = True
        '
        'CheckBoxRegex
        '
        Me.CheckBoxRegex.AutoSize = True
        Me.CheckBoxRegex.Location = New System.Drawing.Point(108, 473)
        Me.CheckBoxRegex.Name = "CheckBoxRegex"
        Me.CheckBoxRegex.Size = New System.Drawing.Size(15, 14)
        Me.CheckBoxRegex.TabIndex = 705
        Me.CheckBoxRegex.UseVisualStyleBackColor = True
        '
        'Label219
        '
        Me.Label219.AutoSize = True
        Me.Label219.Location = New System.Drawing.Point(770, 474)
        Me.Label219.Name = "Label219"
        Me.Label219.Size = New System.Drawing.Size(233, 13)
        Me.Label219.TabIndex = 543
        Me.Label219.Text = "NONE, XONXOFF, RTSCTS, or RTSXONXOFF"
        '
        'Label218
        '
        Me.Label218.AutoSize = True
        Me.Label218.Location = New System.Drawing.Point(770, 448)
        Me.Label218.Name = "Label218"
        Me.Label218.Size = New System.Drawing.Size(55, 13)
        Me.Label218.TabIndex = 542
        Me.Label218.Text = "1, 1.5 or 2"
        '
        'Label217
        '
        Me.Label217.AutoSize = True
        Me.Label217.Location = New System.Drawing.Point(770, 422)
        Me.Label217.Name = "Label217"
        Me.Label217.Size = New System.Drawing.Size(115, 13)
        Me.Label217.TabIndex = 541
        Me.Label217.Text = "NONE, ODD, or EVEN"
        '
        'Label216
        '
        Me.Label216.AutoSize = True
        Me.Label216.Location = New System.Drawing.Point(770, 396)
        Me.Label216.Name = "Label216"
        Me.Label216.Size = New System.Drawing.Size(69, 13)
        Me.Label216.TabIndex = 540
        Me.Label216.Text = "Integer value"
        '
        'Label215
        '
        Me.Label215.AutoSize = True
        Me.Label215.Location = New System.Drawing.Point(769, 370)
        Me.Label215.Name = "Label215"
        Me.Label215.Size = New System.Drawing.Size(69, 13)
        Me.Label215.TabIndex = 539
        Me.Label215.Text = "Integer value"
        '
        'Label214
        '
        Me.Label214.AutoSize = True
        Me.Label214.Location = New System.Drawing.Point(220, 115)
        Me.Label214.Name = "Label214"
        Me.Label214.Size = New System.Drawing.Size(20, 13)
        Me.Label214.TabIndex = 538
        Me.Label214.Text = "Rx"
        '
        'Label213
        '
        Me.Label213.AutoSize = True
        Me.Label213.Location = New System.Drawing.Point(196, 115)
        Me.Label213.Name = "Label213"
        Me.Label213.Size = New System.Drawing.Size(19, 13)
        Me.Label213.TabIndex = 537
        Me.Label213.Text = "Tx"
        '
        'Label212
        '
        Me.Label212.AutoSize = True
        Me.Label212.Location = New System.Drawing.Point(16, 566)
        Me.Label212.Name = "Label212"
        Me.Label212.Size = New System.Drawing.Size(386, 13)
        Me.Label212.TabIndex = 533
        Me.Label212.Text = "Note: If you have a sensor (protocol) that doesn't work here then pls contact me." &
    ""
        '
        'Label211
        '
        Me.Label211.AutoSize = True
        Me.Label211.Location = New System.Drawing.Point(223, 526)
        Me.Label211.Name = "Label211"
        Me.Label211.Size = New System.Drawing.Size(130, 13)
        Me.Label211.TabIndex = 532
        Me.Label211.Text = "Sensor value after parsing"
        '
        'Label210
        '
        Me.Label210.AutoSize = True
        Me.Label210.Location = New System.Drawing.Point(223, 396)
        Me.Label210.Name = "Label210"
        Me.Label210.Size = New System.Drawing.Size(156, 13)
        Me.Label210.TabIndex = 531
        Me.Label210.Text = "Recieved string from the sensor"
        '
        'Label209
        '
        Me.Label209.AutoSize = True
        Me.Label209.Location = New System.Drawing.Point(223, 370)
        Me.Label209.Name = "Label209"
        Me.Label209.Size = New System.Drawing.Size(264, 13)
        Me.Label209.TabIndex = 530
        Me.Label209.Text = "String to send out which commands the sensor to reply"
        '
        'Label208
        '
        Me.Label208.AutoSize = True
        Me.Label208.Location = New System.Drawing.Point(573, 507)
        Me.Label208.Name = "Label208"
        Me.Label208.Size = New System.Drawing.Size(157, 13)
        Me.Label208.TabIndex = 529
        Me.Label208.Text = "Note: Applies to USB-User only."
        '
        'Label207
        '
        Me.Label207.AutoSize = True
        Me.Label207.Location = New System.Drawing.Point(572, 474)
        Me.Label207.Name = "Label207"
        Me.Label207.Size = New System.Drawing.Size(62, 13)
        Me.Label207.TabIndex = 528
        Me.Label207.Text = "Handshake"
        '
        'Label206
        '
        Me.Label206.AutoSize = True
        Me.Label206.Location = New System.Drawing.Point(572, 448)
        Me.Label206.Name = "Label206"
        Me.Label206.Size = New System.Drawing.Size(49, 13)
        Me.Label206.TabIndex = 526
        Me.Label206.Text = "Stop Bits"
        '
        'Label205
        '
        Me.Label205.AutoSize = True
        Me.Label205.Location = New System.Drawing.Point(572, 422)
        Me.Label205.Name = "Label205"
        Me.Label205.Size = New System.Drawing.Size(33, 13)
        Me.Label205.TabIndex = 524
        Me.Label205.Text = "Parity"
        '
        'Label204
        '
        Me.Label204.AutoSize = True
        Me.Label204.Location = New System.Drawing.Point(572, 396)
        Me.Label204.Name = "Label204"
        Me.Label204.Size = New System.Drawing.Size(50, 13)
        Me.Label204.TabIndex = 522
        Me.Label204.Text = "Data Bits"
        '
        'Label195
        '
        Me.Label195.AutoSize = True
        Me.Label195.Location = New System.Drawing.Point(572, 370)
        Me.Label195.Name = "Label195"
        Me.Label195.Size = New System.Drawing.Size(58, 13)
        Me.Label195.TabIndex = 520
        Me.Label195.Text = "Baud Rate"
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label47.Location = New System.Drawing.Point(12, 337)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(334, 13)
        Me.Label47.TabIndex = 518
        Me.Label47.Text = "USB-User USER DEFINED PROTOCOL (TEMPERATURE)"
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Location = New System.Drawing.Point(222, 474)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(282, 13)
        Me.Label46.TabIndex = 517
        Me.Label46.Text = "Regular Expression i.e. extract numeric value (\d+(\.\d+)?)"
        '
        'Label202
        '
        Me.Label202.AutoSize = True
        Me.Label202.Location = New System.Drawing.Point(222, 500)
        Me.Label202.Name = "Label202"
        Me.Label202.Size = New System.Drawing.Size(186, 13)
        Me.Label202.TabIndex = 514
        Me.Label202.Text = "Comma delimit operations i.e. *1.8,+32"
        '
        'Label203
        '
        Me.Label203.AutoSize = True
        Me.Label203.Location = New System.Drawing.Point(17, 499)
        Me.Label203.Name = "Label203"
        Me.Label203.Size = New System.Drawing.Size(79, 13)
        Me.Label203.TabIndex = 513
        Me.Label203.Text = "Arithmentic Op."
        '
        'TextBoxTempArithmentic
        '
        Me.TextBoxTempArithmentic.Location = New System.Drawing.Point(129, 496)
        Me.TextBoxTempArithmentic.Name = "TextBoxTempArithmentic"
        Me.TextBoxTempArithmentic.Size = New System.Drawing.Size(88, 20)
        Me.TextBoxTempArithmentic.TabIndex = 512
        Me.TextBoxTempArithmentic.Text = "*1.8,+32"
        '
        'Label201
        '
        Me.Label201.AutoSize = True
        Me.Label201.Location = New System.Drawing.Point(17, 474)
        Me.Label201.Name = "Label201"
        Me.Label201.Size = New System.Drawing.Size(38, 13)
        Me.Label201.TabIndex = 511
        Me.Label201.Text = "Regex"
        '
        'Label200
        '
        Me.Label200.AutoSize = True
        Me.Label200.Location = New System.Drawing.Point(223, 448)
        Me.Label200.Name = "Label200"
        Me.Label200.Size = New System.Drawing.Size(229, 13)
        Me.Label200.TabIndex = 509
        Me.Label200.Text = "Numerical value to set right side cut off position"
        '
        'Label199
        '
        Me.Label199.AutoSize = True
        Me.Label199.Location = New System.Drawing.Point(223, 422)
        Me.Label199.Name = "Label199"
        Me.Label199.Size = New System.Drawing.Size(223, 13)
        Me.Label199.TabIndex = 508
        Me.Label199.Text = "Numerical value to set left side cut off position"
        '
        'TempFinalValue
        '
        Me.TempFinalValue.AutoSize = True
        Me.TempFinalValue.Location = New System.Drawing.Point(17, 525)
        Me.TempFinalValue.Name = "TempFinalValue"
        Me.TempFinalValue.Size = New System.Drawing.Size(59, 13)
        Me.TempFinalValue.TabIndex = 506
        Me.TempFinalValue.Text = "Final Value"
        '
        'Label197
        '
        Me.Label197.AutoSize = True
        Me.Label197.Location = New System.Drawing.Point(17, 447)
        Me.Label197.Name = "Label197"
        Me.Label197.Size = New System.Drawing.Size(62, 13)
        Me.Label197.TabIndex = 504
        Me.Label197.Text = "Parse Right"
        '
        'Label196
        '
        Me.Label196.AutoSize = True
        Me.Label196.Location = New System.Drawing.Point(17, 421)
        Me.Label196.Name = "Label196"
        Me.Label196.Size = New System.Drawing.Size(55, 13)
        Me.Label196.TabIndex = 503
        Me.Label196.Text = "Parse Left"
        '
        'TextBoxParseRight
        '
        Me.TextBoxParseRight.Location = New System.Drawing.Point(129, 444)
        Me.TextBoxParseRight.Name = "TextBoxParseRight"
        Me.TextBoxParseRight.Size = New System.Drawing.Size(38, 20)
        Me.TextBoxParseRight.TabIndex = 502
        Me.TextBoxParseRight.Text = "10"
        '
        'TextBoxParseLeft
        '
        Me.TextBoxParseLeft.Location = New System.Drawing.Point(129, 418)
        Me.TextBoxParseLeft.Name = "TextBoxParseLeft"
        Me.TextBoxParseLeft.Size = New System.Drawing.Size(38, 20)
        Me.TextBoxParseLeft.TabIndex = 501
        Me.TextBoxParseLeft.Text = "0"
        '
        'Label194
        '
        Me.Label194.AutoSize = True
        Me.Label194.Location = New System.Drawing.Point(17, 396)
        Me.Label194.Name = "Label194"
        Me.Label194.Size = New System.Drawing.Size(50, 13)
        Me.Label194.TabIndex = 499
        Me.Label194.Text = "Rx String"
        '
        'Label193
        '
        Me.Label193.AutoSize = True
        Me.Label193.Location = New System.Drawing.Point(17, 370)
        Me.Label193.Name = "Label193"
        Me.Label193.Size = New System.Drawing.Size(69, 13)
        Me.Label193.TabIndex = 498
        Me.Label193.Text = "Tx Command"
        '
        'Label186
        '
        Me.Label186.AutoSize = True
        Me.Label186.Location = New System.Drawing.Point(743, 145)
        Me.Label186.Name = "Label186"
        Me.Label186.Size = New System.Drawing.Size(235, 13)
        Me.Label186.TabIndex = 494
        Me.Label186.Text = "Adafruit MCP2221A/SHT40,41,45 (USB device)"
        '
        'Label185
        '
        Me.Label185.AutoSize = True
        Me.Label185.Location = New System.Drawing.Point(743, 125)
        Me.Label185.Name = "Label185"
        Me.Label185.Size = New System.Drawing.Size(139, 13)
        Me.Label185.TabIndex = 493
        Me.Label185.Text = "Dogratian USB-TnH SHT30"
        '
        'Label184
        '
        Me.Label184.AutoSize = True
        Me.Label184.Location = New System.Drawing.Point(743, 105)
        Me.Label184.Name = "Label184"
        Me.Label184.Size = New System.Drawing.Size(132, 13)
        Me.Label184.TabIndex = 492
        Me.Label184.Text = "Dogratian USB-TnH LM75"
        '
        'Label183
        '
        Me.Label183.AutoSize = True
        Me.Label183.Location = New System.Drawing.Point(743, 85)
        Me.Label183.Name = "Label183"
        Me.Label183.Size = New System.Drawing.Size(139, 13)
        Me.Label183.TabIndex = 491
        Me.Label183.Text = "Dogratian USB-PA BME280"
        '
        'Label182
        '
        Me.Label182.AutoSize = True
        Me.Label182.Location = New System.Drawing.Point(743, 65)
        Me.Label182.Name = "Label182"
        Me.Label182.Size = New System.Drawing.Size(139, 13)
        Me.Label182.TabIndex = 490
        Me.Label182.Text = "Dogratian USB-TnH SHT10"
        '
        'Label181
        '
        Me.Label181.AutoSize = True
        Me.Label181.Location = New System.Drawing.Point(743, 45)
        Me.Label181.Name = "Label181"
        Me.Label181.Size = New System.Drawing.Size(164, 13)
        Me.Label181.TabIndex = 489
        Me.Label181.Text = "Dogratian USB-TnH SHT10 V2.0"
        '
        'Label178
        '
        Me.Label178.AutoSize = True
        Me.Label178.Location = New System.Drawing.Point(581, 45)
        Me.Label178.Name = "Label178"
        Me.Label178.Size = New System.Drawing.Size(139, 13)
        Me.Label178.TabIndex = 488
        Me.Label178.Text = "SENSOR COMPATIBILITY:"
        '
        'Label164
        '
        Me.Label164.AutoSize = True
        Me.Label164.Location = New System.Drawing.Point(259, 147)
        Me.Label164.Name = "Label164"
        Me.Label164.Size = New System.Drawing.Size(56, 13)
        Me.Label164.TabIndex = 487
        Me.Label164.Text = "Offset Adj."
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.Location = New System.Drawing.Point(12, 16)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(175, 13)
        Me.Label23.TabIndex = 485
        Me.Label23.Text = "TEMPERATURE / HUMIDITY"
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(16, 45)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(35, 13)
        Me.Label29.TabIndex = 24
        Me.Label29.Text = "Name"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(16, 73)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(49, 13)
        Me.Label28.TabIndex = 24
        Me.Label28.Text = "Interface"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(17, 172)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(47, 13)
        Me.Label21.TabIndex = 44
        Me.Label21.Text = "Humidity"
        '
        'LabelHumidity
        '
        Me.LabelHumidity.AutoSize = True
        Me.LabelHumidity.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelHumidity.Location = New System.Drawing.Point(104, 170)
        Me.LabelHumidity.Name = "LabelHumidity"
        Me.LabelHumidity.Size = New System.Drawing.Size(42, 16)
        Me.LabelHumidity.TabIndex = 43
        Me.LabelHumidity.Text = "#####"
        '
        'LabelTemperature
        '
        Me.LabelTemperature.AutoSize = True
        Me.LabelTemperature.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelTemperature.Location = New System.Drawing.Point(104, 146)
        Me.LabelTemperature.Name = "LabelTemperature"
        Me.LabelTemperature.Size = New System.Drawing.Size(42, 16)
        Me.LabelTemperature.TabIndex = 42
        Me.LabelTemperature.Text = "#####"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(17, 149)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(67, 13)
        Me.Label20.TabIndex = 26
        Me.Label20.Text = "Temperature"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(17, 101)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(53, 13)
        Me.Label19.TabIndex = 24
        Me.Label19.Text = "COM Port"
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage3.Controls.Add(Me.ClearEventLOG)
        Me.TabPage3.Controls.Add(Me.Label291)
        Me.TabPage3.Controls.Add(Me.ListLog)
        Me.TabPage3.Controls.Add(Me.LogFileMetadata)
        Me.TabPage3.Controls.Add(Me.Label24)
        Me.TabPage3.Controls.Add(Me.bgoxdata)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Data & Event Log/CSV  "
        '
        'Label291
        '
        Me.Label291.AutoSize = True
        Me.Label291.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label291.Location = New System.Drawing.Point(685, 518)
        Me.Label291.Name = "Label291"
        Me.Label291.Size = New System.Drawing.Size(133, 13)
        Me.Label291.TabIndex = 102
        Me.Label291.Text = "LOG FILE METADATA"
        '
        'ListLog
        '
        Me.ListLog.BackColor = System.Drawing.Color.White
        Me.ListLog.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListLog.FormattingEnabled = True
        Me.ListLog.ItemHeight = 14
        Me.ListLog.Location = New System.Drawing.Point(685, 40)
        Me.ListLog.Name = "ListLog"
        Me.ListLog.Size = New System.Drawing.Size(353, 466)
        Me.ListLog.TabIndex = 100
        '
        'LogFileMetadata
        '
        Me.LogFileMetadata.Location = New System.Drawing.Point(685, 534)
        Me.LogFileMetadata.Multiline = True
        Me.LogFileMetadata.Name = "LogFileMetadata"
        Me.LogFileMetadata.Size = New System.Drawing.Size(354, 49)
        Me.LogFileMetadata.TabIndex = 103
        Me.LogFileMetadata.Text = "Sample 1" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Sample 2" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Sample 3"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.Location = New System.Drawing.Point(681, 17)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(77, 13)
        Me.Label24.TabIndex = 99
        Me.Label24.Text = "EVENT LOG"
        '
        'bgoxdata
        '
        Me.bgoxdata.Controls.Add(Me.CSVsize)
        Me.bgoxdata.Controls.Add(Me.CSVcounts)
        Me.bgoxdata.Controls.Add(Me.CSVwrite)
        Me.bgoxdata.Controls.Add(Me.Label228)
        Me.bgoxdata.Controls.Add(Me.Label226)
        Me.bgoxdata.Controls.Add(Me.ClearLOGdisp)
        Me.bgoxdata.Controls.Add(Me.Label177)
        Me.bgoxdata.Controls.Add(Me.TextFilenameAppend)
        Me.bgoxdata.Controls.Add(Me.Label176)
        Me.bgoxdata.Controls.Add(Me.CSVEntryLimitMins)
        Me.bgoxdata.Controls.Add(Me.CheckboxCSVlimitMins)
        Me.bgoxdata.Controls.Add(Me.Label175)
        Me.bgoxdata.Controls.Add(Me.CSVEntryLimit)
        Me.bgoxdata.Controls.Add(Me.CheckboxCSVlimit)
        Me.bgoxdata.Controls.Add(Me.Label58)
        Me.bgoxdata.Controls.Add(Me.Label71)
        Me.bgoxdata.Controls.Add(Me.CSVdelimiterSemiColon)
        Me.bgoxdata.Controls.Add(Me.CSVdelimiterComma)
        Me.bgoxdata.Controls.Add(Me.LabelCSVfilesize)
        Me.bgoxdata.Controls.Add(Me.Label52)
        Me.bgoxdata.Controls.Add(Me.Label51)
        Me.bgoxdata.Controls.Add(Me.Label50)
        Me.bgoxdata.Controls.Add(Me.Label49)
        Me.bgoxdata.Controls.Add(Me.Label48)
        Me.bgoxdata.Controls.Add(Me.LabelCSVcounts)
        Me.bgoxdata.Controls.Add(Me.CheckboxEnableLOG)
        Me.bgoxdata.Controls.Add(Me.ShowFiles)
        Me.bgoxdata.Controls.Add(Me.ResetCSV)
        Me.bgoxdata.Controls.Add(Me.ButtonExportCSV)
        Me.bgoxdata.Controls.Add(Me.ENotationDecimal)
        Me.bgoxdata.Controls.Add(Me.CheckboxEnableCSV)
        Me.bgoxdata.Controls.Add(Me.TempHumLogs)
        Me.bgoxdata.Controls.Add(Me.Label27)
        Me.bgoxdata.Controls.Add(Me.CSVfilename)
        Me.bgoxdata.Controls.Add(Me.Label25)
        Me.bgoxdata.Controls.Add(Me.CSVfilepath)
        Me.bgoxdata.Controls.Add(Me.Label26)
        Me.bgoxdata.Controls.Add(Me.ListBoxData)
        Me.bgoxdata.Enabled = False
        Me.bgoxdata.Location = New System.Drawing.Point(6, 6)
        Me.bgoxdata.Name = "bgoxdata"
        Me.bgoxdata.Size = New System.Drawing.Size(1039, 585)
        Me.bgoxdata.TabIndex = 51
        Me.bgoxdata.TabStop = False
        '
        'CSVsize
        '
        Me.CSVsize.AutoSize = True
        Me.CSVsize.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSVsize.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.CSVsize.Location = New System.Drawing.Point(643, 119)
        Me.CSVsize.Name = "CSVsize"
        Me.CSVsize.Size = New System.Drawing.Size(13, 13)
        Me.CSVsize.TabIndex = 64
        Me.CSVsize.Text = "0"
        '
        'CSVcounts
        '
        Me.CSVcounts.AutoSize = True
        Me.CSVcounts.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSVcounts.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.CSVcounts.Location = New System.Drawing.Point(643, 100)
        Me.CSVcounts.Name = "CSVcounts"
        Me.CSVcounts.Size = New System.Drawing.Size(13, 13)
        Me.CSVcounts.TabIndex = 56
        Me.CSVcounts.Text = "0"
        '
        'CSVwrite
        '
        Me.CSVwrite.AutoSize = True
        Me.CSVwrite.Location = New System.Drawing.Point(6, 563)
        Me.CSVwrite.Name = "CSVwrite"
        Me.CSVwrite.Size = New System.Drawing.Size(14, 13)
        Me.CSVwrite.TabIndex = 98
        Me.CSVwrite.Text = "#"
        '
        'Label228
        '
        Me.Label228.AutoSize = True
        Me.Label228.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label228.Location = New System.Drawing.Point(6, 545)
        Me.Label228.Name = "Label228"
        Me.Label228.Size = New System.Drawing.Size(119, 13)
        Me.Label228.TabIndex = 97
        Me.Label228.Text = "CSV WRITE (LAST)"
        '
        'Label226
        '
        Me.Label226.AutoSize = True
        Me.Label226.Location = New System.Drawing.Point(472, 405)
        Me.Label226.Name = "Label226"
        Me.Label226.Size = New System.Drawing.Size(183, 52)
        Me.Label226.TabIndex = 97
        Me.Label226.Text = "Note:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Set File name, Log File Metadata and" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Enable CSV before hitting RUN on" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "De" &
    "vice Tab."
        '
        'Label177
        '
        Me.Label177.AutoSize = True
        Me.Label177.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label177.Location = New System.Drawing.Point(470, 501)
        Me.Label177.Name = "Label177"
        Me.Label177.Size = New System.Drawing.Size(184, 13)
        Me.Label177.TabIndex = 93
        Me.Label177.Text = "Optional: Text to append to Filename:"
        Me.Label177.Visible = False
        '
        'Label176
        '
        Me.Label176.AutoSize = True
        Me.Label176.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label176.Location = New System.Drawing.Point(538, 351)
        Me.Label176.Name = "Label176"
        Me.Label176.Size = New System.Drawing.Size(44, 13)
        Me.Label176.TabIndex = 91
        Me.Label176.Text = "Minutes"
        '
        'CSVEntryLimitMins
        '
        Me.CSVEntryLimitMins.Location = New System.Drawing.Point(494, 347)
        Me.CSVEntryLimitMins.Name = "CSVEntryLimitMins"
        Me.CSVEntryLimitMins.Size = New System.Drawing.Size(42, 20)
        Me.CSVEntryLimitMins.TabIndex = 90
        Me.CSVEntryLimitMins.Text = "60"
        '
        'Label175
        '
        Me.Label175.AutoSize = True
        Me.Label175.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label175.Location = New System.Drawing.Point(548, 303)
        Me.Label175.Name = "Label175"
        Me.Label175.Size = New System.Drawing.Size(39, 13)
        Me.Label175.TabIndex = 88
        Me.Label175.Text = "Entries"
        '
        'CSVEntryLimit
        '
        Me.CSVEntryLimit.Location = New System.Drawing.Point(494, 299)
        Me.CSVEntryLimit.Name = "CSVEntryLimit"
        Me.CSVEntryLimit.Size = New System.Drawing.Size(52, 20)
        Me.CSVEntryLimit.TabIndex = 87
        Me.CSVEntryLimit.Text = "1000"
        '
        'Label58
        '
        Me.Label58.AutoSize = True
        Me.Label58.Location = New System.Drawing.Point(471, 147)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(71, 13)
        Me.Label58.TabIndex = 85
        Me.Label58.Text = "CSV Delimiter"
        '
        'Label71
        '
        Me.Label71.AutoSize = True
        Me.Label71.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label71.Location = New System.Drawing.Point(170, 86)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(28, 12)
        Me.Label71.TabIndex = 84
        Me.Label71.Text = "TIME"
        '
        'CSVdelimiterSemiColon
        '
        Me.CSVdelimiterSemiColon.AutoSize = True
        Me.CSVdelimiterSemiColon.Location = New System.Drawing.Point(604, 145)
        Me.CSVdelimiterSemiColon.Name = "CSVdelimiterSemiColon"
        Me.CSVdelimiterSemiColon.Size = New System.Drawing.Size(77, 17)
        Me.CSVdelimiterSemiColon.TabIndex = 83
        Me.CSVdelimiterSemiColon.Text = "Semi-colon"
        Me.CSVdelimiterSemiColon.UseVisualStyleBackColor = True
        '
        'CSVdelimiterComma
        '
        Me.CSVdelimiterComma.AutoSize = True
        Me.CSVdelimiterComma.Location = New System.Drawing.Point(544, 145)
        Me.CSVdelimiterComma.Name = "CSVdelimiterComma"
        Me.CSVdelimiterComma.Size = New System.Drawing.Size(60, 17)
        Me.CSVdelimiterComma.TabIndex = 82
        Me.CSVdelimiterComma.Text = "Comma"
        Me.CSVdelimiterComma.UseVisualStyleBackColor = True
        '
        'LabelCSVfilesize
        '
        Me.LabelCSVfilesize.AutoSize = True
        Me.LabelCSVfilesize.Location = New System.Drawing.Point(471, 119)
        Me.LabelCSVfilesize.Name = "LabelCSVfilesize"
        Me.LabelCSVfilesize.Size = New System.Drawing.Size(98, 13)
        Me.LabelCSVfilesize.TabIndex = 63
        Me.LabelCSVfilesize.Text = "CSV File No. Lines "
        '
        'Label52
        '
        Me.Label52.AutoSize = True
        Me.Label52.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label52.Location = New System.Drawing.Point(414, 86)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(31, 12)
        Me.Label52.TabIndex = 62
        Me.Label52.Text = "HUM."
        '
        'Label51
        '
        Me.Label51.AutoSize = True
        Me.Label51.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label51.Location = New System.Drawing.Point(359, 86)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(34, 12)
        Me.Label51.TabIndex = 61
        Me.Label51.Text = "TEMP."
        '
        'Label50
        '
        Me.Label50.AutoSize = True
        Me.Label50.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label50.Location = New System.Drawing.Point(247, 86)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(37, 12)
        Me.Label50.TabIndex = 60
        Me.Label50.Text = "VALUE"
        '
        'Label49
        '
        Me.Label49.AutoSize = True
        Me.Label49.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label49.Location = New System.Drawing.Point(93, 86)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(30, 12)
        Me.Label49.TabIndex = 59
        Me.Label49.Text = "DATE"
        '
        'Label48
        '
        Me.Label48.AutoSize = True
        Me.Label48.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label48.Location = New System.Drawing.Point(6, 86)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(41, 12)
        Me.Label48.TabIndex = 58
        Me.Label48.Text = "DEVICE"
        '
        'LabelCSVcounts
        '
        Me.LabelCSVcounts.AutoSize = True
        Me.LabelCSVcounts.Location = New System.Drawing.Point(471, 100)
        Me.LabelCSVcounts.Name = "LabelCSVcounts"
        Me.LabelCSVcounts.Size = New System.Drawing.Size(174, 13)
        Me.LabelCSVcounts.TabIndex = 56
        Me.LabelCSVcounts.Text = "CSV Entries Per Device (This Pass)"
        '
        'CheckboxEnableLOG
        '
        Me.CheckboxEnableLOG.AutoSize = True
        Me.CheckboxEnableLOG.Location = New System.Drawing.Point(472, 231)
        Me.CheckboxEnableLOG.Name = "CheckboxEnableLOG"
        Me.CheckboxEnableLOG.Size = New System.Drawing.Size(116, 17)
        Me.CheckboxEnableLOG.TabIndex = 57
        Me.CheckboxEnableLOG.Text = "Enable DATA LOG"
        Me.CheckboxEnableLOG.UseVisualStyleBackColor = True
        '
        'ENotationDecimal
        '
        Me.ENotationDecimal.AutoSize = True
        Me.ENotationDecimal.Location = New System.Drawing.Point(472, 175)
        Me.ENotationDecimal.Name = "ENotationDecimal"
        Me.ENotationDecimal.Size = New System.Drawing.Size(191, 17)
        Me.ENotationDecimal.TabIndex = 53
        Me.ENotationDecimal.Text = "E notation to Decimal (LOG && CSV)"
        Me.ENotationDecimal.UseVisualStyleBackColor = True
        '
        'CheckboxEnableCSV
        '
        Me.CheckboxEnableCSV.AutoSize = True
        Me.CheckboxEnableCSV.Location = New System.Drawing.Point(472, 259)
        Me.CheckboxEnableCSV.Name = "CheckboxEnableCSV"
        Me.CheckboxEnableCSV.Size = New System.Drawing.Size(188, 17)
        Me.CheckboxEnableCSV.TabIndex = 52
        Me.CheckboxEnableCSV.Text = "Enable CSV  (Enable before RUN)"
        Me.CheckboxEnableCSV.UseVisualStyleBackColor = True
        '
        'TempHumLogs
        '
        Me.TempHumLogs.AutoSize = True
        Me.TempHumLogs.Location = New System.Drawing.Point(472, 204)
        Me.TempHumLogs.Name = "TempHumLogs"
        Me.TempHumLogs.Size = New System.Drawing.Size(206, 17)
        Me.TempHumLogs.TabIndex = 22
        Me.TempHumLogs.Text = "Include Temp/Hum in Data LOG/CSV"
        Me.TempHumLogs.UseVisualStyleBackColor = True
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(9, 34)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(84, 13)
        Me.Label27.TabIndex = 51
        Me.Label27.Text = "CSV FILENAME"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(9, 60)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(85, 13)
        Me.Label25.TabIndex = 22
        Me.Label25.Text = "CSV FILE PATH"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.Location = New System.Drawing.Point(4, 11)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(107, 13)
        Me.Label26.TabIndex = 45
        Me.Label26.Text = "DATA LOG / CSV"
        '
        'ListBoxData
        '
        Me.ListBoxData.BackColor = System.Drawing.Color.White
        Me.ListBoxData.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBoxData.FormattingEnabled = True
        Me.ListBoxData.ItemHeight = 14
        Me.ListBoxData.Location = New System.Drawing.Point(6, 100)
        Me.ListBoxData.Name = "ListBoxData"
        Me.ListBoxData.Size = New System.Drawing.Size(462, 438)
        Me.ListBoxData.TabIndex = 49
        '
        'TabPage4
        '
        Me.TabPage4.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage4.Controls.Add(Me.Label313)
        Me.TabPage4.Controls.Add(Me.Label238)
        Me.TabPage4.Controls.Add(Me.Label237)
        Me.TabPage4.Controls.Add(Me.Label236)
        Me.TabPage4.Controls.Add(Me.Label235)
        Me.TabPage4.Controls.Add(Me.StartChartMessage)
        Me.TabPage4.Controls.Add(Me.LabeChartMinutes)
        Me.TabPage4.Controls.Add(Me.Label223)
        Me.TabPage4.Controls.Add(Me.CheckBoxTempHide)
        Me.TabPage4.Controls.Add(Me.CheckBoxDevice2Hide)
        Me.TabPage4.Controls.Add(Me.CheckBoxDevice1Hide)
        Me.TabPage4.Controls.Add(Me.ButtonSaveLiveSettings)
        Me.TabPage4.Controls.Add(Me.Device2nameLive)
        Me.TabPage4.Controls.Add(Me.Device1nameLive)
        Me.TabPage4.Controls.Add(Me.DisableRollingChart)
        Me.TabPage4.Controls.Add(Me.LabelChartPoints1)
        Me.TabPage4.Controls.Add(Me.Label258)
        Me.TabPage4.Controls.Add(Me.LabelChartPoints2)
        Me.TabPage4.Controls.Add(Me.Label257)
        Me.TabPage4.Controls.Add(Me.ButtonPauseChart)
        Me.TabPage4.Controls.Add(Me.YaxisDiff)
        Me.TabPage4.Controls.Add(Me.Label256)
        Me.TabPage4.Controls.Add(Me.EnableAutoYChart1)
        Me.TabPage4.Controls.Add(Me.Label179)
        Me.TabPage4.Controls.Add(Me.Label180)
        Me.TabPage4.Controls.Add(Me.LCTempMax)
        Me.TabPage4.Controls.Add(Me.LCTempMin)
        Me.TabPage4.Controls.Add(Me.ShowFiles3)
        Me.TabPage4.Controls.Add(Me.XaxisPoints)
        Me.TabPage4.Controls.Add(Me.Dev1Max)
        Me.TabPage4.Controls.Add(Me.Dev1Min)
        Me.TabPage4.Controls.Add(Me.Label39)
        Me.TabPage4.Controls.Add(Me.Label40)
        Me.TabPage4.Controls.Add(Me.Label72)
        Me.TabPage4.Controls.Add(Me.Label41)
        Me.TabPage4.Controls.Add(Me.ButtonDiffRecordedTempReset)
        Me.TabPage4.Controls.Add(Me.TemperatureDiffRecorded)
        Me.TabPage4.Controls.Add(Me.ButtonDiffRecorded2Reset)
        Me.TabPage4.Controls.Add(Me.EnableChart1)
        Me.TabPage4.Controls.Add(Me.ButtonDiffRecorded1Reset)
        Me.TabPage4.Controls.Add(Me.EnableChart3)
        Me.TabPage4.Controls.Add(Me.inst_value2FDiffRecorded)
        Me.TabPage4.Controls.Add(Me.inst_value1FDiffRecorded)
        Me.TabPage4.Controls.Add(Me.Label44)
        Me.TabPage4.Controls.Add(Me.Label42)
        Me.TabPage4.Controls.Add(Me.EnableChart2)
        Me.TabPage4.Controls.Add(Me.ButtonClearChart)
        Me.TabPage4.Controls.Add(Me.Chart1)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Live Chart  "
        '
        'Label313
        '
        Me.Label313.AutoSize = True
        Me.Label313.Location = New System.Drawing.Point(108, 95)
        Me.Label313.Name = "Label313"
        Me.Label313.Size = New System.Drawing.Size(157, 13)
        Me.Label313.TabIndex = 714
        Me.Label313.Text = "(Enable Dev1/Dev2 after RUN)"
        '
        'Label238
        '
        Me.Label238.AutoSize = True
        Me.Label238.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label238.Location = New System.Drawing.Point(825, 6)
        Me.Label238.Name = "Label238"
        Me.Label238.Size = New System.Drawing.Size(100, 13)
        Me.Label238.TabIndex = 713
        Me.Label238.Text = "TEMPERATURE"
        '
        'Label237
        '
        Me.Label237.AutoSize = True
        Me.Label237.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label237.Location = New System.Drawing.Point(557, 6)
        Me.Label237.Name = "Label237"
        Me.Label237.Size = New System.Drawing.Size(76, 13)
        Me.Label237.TabIndex = 712
        Me.Label237.Text = "DEVICE 1/2"
        '
        'Label236
        '
        Me.Label236.AutoSize = True
        Me.Label236.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label236.Location = New System.Drawing.Point(354, 6)
        Me.Label236.Name = "Label236"
        Me.Label236.Size = New System.Drawing.Size(61, 13)
        Me.Label236.TabIndex = 711
        Me.Label236.Text = "X/Y AXIS"
        '
        'Label235
        '
        Me.Label235.AutoSize = True
        Me.Label235.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label235.Location = New System.Drawing.Point(107, 6)
        Me.Label235.Name = "Label235"
        Me.Label235.Size = New System.Drawing.Size(104, 13)
        Me.Label235.TabIndex = 710
        Me.Label235.Text = "DEVICE ENABLE"
        '
        'StartChartMessage
        '
        Me.StartChartMessage.AutoSize = True
        Me.StartChartMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StartChartMessage.Location = New System.Drawing.Point(401, 340)
        Me.StartChartMessage.Name = "StartChartMessage"
        Me.StartChartMessage.Size = New System.Drawing.Size(193, 25)
        Me.StartChartMessage.TabIndex = 709
        Me.StartChartMessage.Text = "Please Start Chart!"
        '
        'LabeChartMinutes
        '
        Me.LabeChartMinutes.AutoSize = True
        Me.LabeChartMinutes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeChartMinutes.Location = New System.Drawing.Point(498, 583)
        Me.LabeChartMinutes.Name = "LabeChartMinutes"
        Me.LabeChartMinutes.Size = New System.Drawing.Size(117, 15)
        Me.LabeChartMinutes.TabIndex = 705
        Me.LabeChartMinutes.Text = "0hrs 00mins 00secs"
        '
        'Label223
        '
        Me.Label223.AutoSize = True
        Me.Label223.Location = New System.Drawing.Point(430, 585)
        Me.Label223.Name = "Label223"
        Me.Label223.Size = New System.Drawing.Size(77, 13)
        Me.Label223.TabIndex = 704
        Me.Label223.Text = "Visible Chart = "
        '
        'CheckBoxTempHide
        '
        Me.CheckBoxTempHide.AutoSize = True
        Me.CheckBoxTempHide.Location = New System.Drawing.Point(829, 94)
        Me.CheckBoxTempHide.Name = "CheckBoxTempHide"
        Me.CheckBoxTempHide.Size = New System.Drawing.Size(140, 17)
        Me.CheckBoxTempHide.TabIndex = 708
        Me.CheckBoxTempHide.Text = "Hide Chart Trace Temp."
        Me.CheckBoxTempHide.UseVisualStyleBackColor = True
        '
        'CheckBoxDevice2Hide
        '
        Me.CheckBoxDevice2Hide.AutoSize = True
        Me.CheckBoxDevice2Hide.Location = New System.Drawing.Point(560, 130)
        Me.CheckBoxDevice2Hide.Name = "CheckBoxDevice2Hide"
        Me.CheckBoxDevice2Hide.Size = New System.Drawing.Size(139, 17)
        Me.CheckBoxDevice2Hide.TabIndex = 707
        Me.CheckBoxDevice2Hide.Text = "Hide Chart Trace Dev 2"
        Me.CheckBoxDevice2Hide.UseVisualStyleBackColor = True
        '
        'CheckBoxDevice1Hide
        '
        Me.CheckBoxDevice1Hide.AutoSize = True
        Me.CheckBoxDevice1Hide.Location = New System.Drawing.Point(560, 63)
        Me.CheckBoxDevice1Hide.Name = "CheckBoxDevice1Hide"
        Me.CheckBoxDevice1Hide.Size = New System.Drawing.Size(139, 17)
        Me.CheckBoxDevice1Hide.TabIndex = 706
        Me.CheckBoxDevice1Hide.Text = "Hide Chart Trace Dev 1"
        Me.CheckBoxDevice1Hide.UseVisualStyleBackColor = True
        '
        'Device2nameLive
        '
        Me.Device2nameLive.AutoSize = True
        Me.Device2nameLive.Location = New System.Drawing.Point(250, 50)
        Me.Device2nameLive.Name = "Device2nameLive"
        Me.Device2nameLive.Size = New System.Drawing.Size(63, 13)
        Me.Device2nameLive.TabIndex = 702
        Me.Device2nameLive.Text = "########"
        '
        'Device1nameLive
        '
        Me.Device1nameLive.AutoSize = True
        Me.Device1nameLive.Location = New System.Drawing.Point(250, 27)
        Me.Device1nameLive.Name = "Device1nameLive"
        Me.Device1nameLive.Size = New System.Drawing.Size(63, 13)
        Me.Device1nameLive.TabIndex = 701
        Me.Device1nameLive.Text = "########"
        '
        'DisableRollingChart
        '
        Me.DisableRollingChart.AutoSize = True
        Me.DisableRollingChart.Location = New System.Drawing.Point(357, 115)
        Me.DisableRollingChart.Name = "DisableRollingChart"
        Me.DisableRollingChart.Size = New System.Drawing.Size(155, 17)
        Me.DisableRollingChart.TabIndex = 700
        Me.DisableRollingChart.Text = "Disable Y-axis Rolling Chart"
        Me.DisableRollingChart.UseVisualStyleBackColor = True
        '
        'LabelChartPoints1
        '
        Me.LabelChartPoints1.AutoSize = True
        Me.LabelChartPoints1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelChartPoints1.Location = New System.Drawing.Point(700, 44)
        Me.LabelChartPoints1.Name = "LabelChartPoints1"
        Me.LabelChartPoints1.Size = New System.Drawing.Size(14, 15)
        Me.LabelChartPoints1.TabIndex = 699
        Me.LabelChartPoints1.Text = "0"
        '
        'Label258
        '
        Me.Label258.AutoSize = True
        Me.Label258.Location = New System.Drawing.Point(557, 46)
        Me.Label258.Name = "Label258"
        Me.Label258.Size = New System.Drawing.Size(146, 13)
        Me.Label258.TabIndex = 698
        Me.Label258.Text = "Dev 1 Chart samples visible ="
        '
        'LabelChartPoints2
        '
        Me.LabelChartPoints2.AutoSize = True
        Me.LabelChartPoints2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelChartPoints2.Location = New System.Drawing.Point(701, 111)
        Me.LabelChartPoints2.Name = "LabelChartPoints2"
        Me.LabelChartPoints2.Size = New System.Drawing.Size(14, 15)
        Me.LabelChartPoints2.TabIndex = 697
        Me.LabelChartPoints2.Text = "0"
        '
        'Label257
        '
        Me.Label257.AutoSize = True
        Me.Label257.Location = New System.Drawing.Point(558, 113)
        Me.Label257.Name = "Label257"
        Me.Label257.Size = New System.Drawing.Size(146, 13)
        Me.Label257.TabIndex = 696
        Me.Label257.Text = "Dev 2 Chart samples visible ="
        '
        'YaxisDiff
        '
        Me.YaxisDiff.AutoSize = True
        Me.YaxisDiff.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.YaxisDiff.Location = New System.Drawing.Point(432, 136)
        Me.YaxisDiff.Name = "YaxisDiff"
        Me.YaxisDiff.Size = New System.Drawing.Size(63, 13)
        Me.YaxisDiff.TabIndex = 693
        Me.YaxisDiff.Text = "########"
        '
        'Label256
        '
        Me.Label256.AutoSize = True
        Me.Label256.Location = New System.Drawing.Point(354, 136)
        Me.Label256.Name = "Label256"
        Me.Label256.Size = New System.Drawing.Size(80, 13)
        Me.Label256.TabIndex = 694
        Me.Label256.Text = "Y-Axis Range ="
        '
        'Label179
        '
        Me.Label179.AutoSize = True
        Me.Label179.Location = New System.Drawing.Point(874, 73)
        Me.Label179.Name = "Label179"
        Me.Label179.Size = New System.Drawing.Size(118, 13)
        Me.Label179.TabIndex = 689
        Me.Label179.Text = "Temp. Y-axis Scale Min"
        '
        'Label180
        '
        Me.Label180.AutoSize = True
        Me.Label180.Location = New System.Drawing.Point(874, 50)
        Me.Label180.Name = "Label180"
        Me.Label180.Size = New System.Drawing.Size(121, 13)
        Me.Label180.TabIndex = 690
        Me.Label180.Text = "Temp. Y-axis Scale Max"
        '
        'LCTempMax
        '
        Me.LCTempMax.Location = New System.Drawing.Point(829, 46)
        Me.LCTempMax.Name = "LCTempMax"
        Me.LCTempMax.Size = New System.Drawing.Size(39, 20)
        Me.LCTempMax.TabIndex = 687
        Me.LCTempMax.Text = "30"
        '
        'LCTempMin
        '
        Me.LCTempMin.Location = New System.Drawing.Point(829, 69)
        Me.LCTempMin.Name = "LCTempMin"
        Me.LCTempMin.Size = New System.Drawing.Size(39, 20)
        Me.LCTempMin.TabIndex = 688
        Me.LCTempMin.Text = "15"
        '
        'Dev1Max
        '
        Me.Dev1Max.Location = New System.Drawing.Point(357, 68)
        Me.Dev1Max.Name = "Dev1Max"
        Me.Dev1Max.Size = New System.Drawing.Size(71, 20)
        Me.Dev1Max.TabIndex = 92
        Me.Dev1Max.Text = "10"
        '
        'Dev1Min
        '
        Me.Dev1Min.Location = New System.Drawing.Point(357, 91)
        Me.Dev1Min.Name = "Dev1Min"
        Me.Dev1Min.Size = New System.Drawing.Size(71, 20)
        Me.Dev1Min.TabIndex = 94
        Me.Dev1Min.Text = "0"
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(433, 93)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(85, 13)
        Me.Label39.TabIndex = 88
        Me.Label39.Text = "Y-axis Scale Min"
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(433, 70)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(88, 13)
        Me.Label40.TabIndex = 95
        Me.Label40.Text = "Y-axis Scale Max"
        '
        'Label72
        '
        Me.Label72.AutoSize = True
        Me.Label72.Location = New System.Drawing.Point(825, 27)
        Me.Label72.Name = "Label72"
        Me.Label72.Size = New System.Drawing.Size(138, 13)
        Me.Label72.TabIndex = 110
        Me.Label72.Text = "Temp. Max. Diff. Recorded:"
        '
        'ButtonDiffRecordedTempReset
        '
        Me.ButtonDiffRecordedTempReset.Location = New System.Drawing.Point(998, 23)
        Me.ButtonDiffRecordedTempReset.Name = "ButtonDiffRecordedTempReset"
        Me.ButtonDiffRecordedTempReset.Size = New System.Drawing.Size(44, 20)
        Me.ButtonDiffRecordedTempReset.TabIndex = 109
        Me.ButtonDiffRecordedTempReset.Text = "Reset"
        Me.ButtonDiffRecordedTempReset.UseVisualStyleBackColor = True
        '
        'TemperatureDiffRecorded
        '
        Me.TemperatureDiffRecorded.AutoSize = True
        Me.TemperatureDiffRecorded.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TemperatureDiffRecorded.Location = New System.Drawing.Point(961, 27)
        Me.TemperatureDiffRecorded.Name = "TemperatureDiffRecorded"
        Me.TemperatureDiffRecorded.Size = New System.Drawing.Size(34, 13)
        Me.TemperatureDiffRecorded.TabIndex = 108
        Me.TemperatureDiffRecorded.Text = "00.00"
        '
        'ButtonDiffRecorded2Reset
        '
        Me.ButtonDiffRecorded2Reset.Location = New System.Drawing.Point(772, 90)
        Me.ButtonDiffRecorded2Reset.Name = "ButtonDiffRecorded2Reset"
        Me.ButtonDiffRecorded2Reset.Size = New System.Drawing.Size(44, 20)
        Me.ButtonDiffRecorded2Reset.TabIndex = 105
        Me.ButtonDiffRecorded2Reset.Text = "Reset"
        Me.ButtonDiffRecorded2Reset.UseVisualStyleBackColor = True
        '
        'EnableChart1
        '
        Me.EnableChart1.AutoSize = True
        Me.EnableChart1.Enabled = False
        Me.EnableChart1.Location = New System.Drawing.Point(110, 26)
        Me.EnableChart1.Name = "EnableChart1"
        Me.EnableChart1.Size = New System.Drawing.Size(142, 17)
        Me.EnableChart1.TabIndex = 85
        Me.EnableChart1.Text = "Enable Chart Device 1 ="
        Me.EnableChart1.UseVisualStyleBackColor = True
        '
        'ButtonDiffRecorded1Reset
        '
        Me.ButtonDiffRecorded1Reset.Location = New System.Drawing.Point(771, 23)
        Me.ButtonDiffRecorded1Reset.Name = "ButtonDiffRecorded1Reset"
        Me.ButtonDiffRecorded1Reset.Size = New System.Drawing.Size(44, 20)
        Me.ButtonDiffRecorded1Reset.TabIndex = 104
        Me.ButtonDiffRecorded1Reset.Text = "Reset"
        Me.ButtonDiffRecorded1Reset.UseVisualStyleBackColor = True
        '
        'EnableChart3
        '
        Me.EnableChart3.AutoSize = True
        Me.EnableChart3.Enabled = False
        Me.EnableChart3.Location = New System.Drawing.Point(110, 72)
        Me.EnableChart3.Name = "EnableChart3"
        Me.EnableChart3.Size = New System.Drawing.Size(150, 17)
        Me.EnableChart3.TabIndex = 103
        Me.EnableChart3.Text = "Enable Chart Temperature"
        Me.EnableChart3.UseVisualStyleBackColor = True
        '
        'inst_value2FDiffRecorded
        '
        Me.inst_value2FDiffRecorded.AutoSize = True
        Me.inst_value2FDiffRecorded.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.inst_value2FDiffRecorded.Location = New System.Drawing.Point(700, 94)
        Me.inst_value2FDiffRecorded.Name = "inst_value2FDiffRecorded"
        Me.inst_value2FDiffRecorded.Size = New System.Drawing.Size(70, 13)
        Me.inst_value2FDiffRecorded.TabIndex = 100
        Me.inst_value2FDiffRecorded.Text = "#########"
        '
        'inst_value1FDiffRecorded
        '
        Me.inst_value1FDiffRecorded.AutoSize = True
        Me.inst_value1FDiffRecorded.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.inst_value1FDiffRecorded.Location = New System.Drawing.Point(700, 27)
        Me.inst_value1FDiffRecorded.Name = "inst_value1FDiffRecorded"
        Me.inst_value1FDiffRecorded.Size = New System.Drawing.Size(70, 13)
        Me.inst_value1FDiffRecorded.TabIndex = 90
        Me.inst_value1FDiffRecorded.Text = "#########"
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(558, 94)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(146, 13)
        Me.Label44.TabIndex = 99
        Me.Label44.Text = "Dev 2 Max. Diff. Recorded  ="
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(557, 27)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(146, 13)
        Me.Label42.TabIndex = 93
        Me.Label42.Text = "Dev 1 Max. Diff. Recorded  ="
        '
        'EnableChart2
        '
        Me.EnableChart2.AutoSize = True
        Me.EnableChart2.Enabled = False
        Me.EnableChart2.Location = New System.Drawing.Point(110, 49)
        Me.EnableChart2.Name = "EnableChart2"
        Me.EnableChart2.Size = New System.Drawing.Size(142, 17)
        Me.EnableChart2.TabIndex = 97
        Me.EnableChart2.Text = "Enable Chart Device 2 ="
        Me.EnableChart2.UseVisualStyleBackColor = True
        '
        'Chart1
        '
        Me.Chart1.BackColor = System.Drawing.SystemColors.Control
        ChartArea3.BackColor = System.Drawing.Color.Black
        ChartArea3.BorderColor = System.Drawing.Color.White
        ChartArea3.BorderWidth = 2
        ChartArea3.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea3)
        Legend3.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend3)
        Me.Chart1.Location = New System.Drawing.Point(-32, 140)
        Me.Chart1.Name = "Chart1"
        Series3.ChartArea = "ChartArea1"
        Series3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        Series3.Color = System.Drawing.Color.Yellow
        Series3.Enabled = False
        Series3.Legend = "Legend1"
        Series3.Name = "Series1"
        Me.Chart1.Series.Add(Series3)
        Me.Chart1.Size = New System.Drawing.Size(1120, 462)
        Me.Chart1.TabIndex = 87
        Me.Chart1.Text = "Chart1"
        '
        'TabPage9
        '
        Me.TabPage9.Location = New System.Drawing.Point(4, 22)
        Me.TabPage9.Name = "TabPage9"
        Me.TabPage9.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage9.TabIndex = 8
        Me.TabPage9.Text = "Playback Chart  "
        Me.TabPage9.UseVisualStyleBackColor = True
        '
        'TabPage7
        '
        Me.TabPage7.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage7.Controls.Add(Me.AddressRangeF)
        Me.TabPage7.Controls.Add(Me.AddressRangeB)
        Me.TabPage7.Controls.Add(Me.AddressRangeA)
        Me.TabPage7.Controls.Add(Me.TextBoxCalRamFile3457A)
        Me.TabPage7.Controls.Add(Me.Label125)
        Me.TabPage7.Controls.Add(Me.ButtonCalramDump3457A)
        Me.TabPage7.Controls.Add(Me.Label122)
        Me.TabPage7.Controls.Add(Me.Label123)
        Me.TabPage7.Controls.Add(Me.Label116)
        Me.TabPage7.Controls.Add(Me.TextBoxCalRamFile)
        Me.TabPage7.Controls.Add(Me.ShowFilesCalRam)
        Me.TabPage7.Controls.Add(Me.Label62)
        Me.TabPage7.Controls.Add(Me.Label63)
        Me.TabPage7.Controls.Add(Me.Label60)
        Me.TabPage7.Controls.Add(Me.Label38)
        Me.TabPage7.Controls.Add(Me.GroupBox6)
        Me.TabPage7.Controls.Add(Me.GroupBox7)
        Me.TabPage7.Location = New System.Drawing.Point(4, 22)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage7.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage7.TabIndex = 6
        Me.TabPage7.Text = "345xA Cal Extract  "
        '
        'AddressRangeF
        '
        Me.AddressRangeF.AutoSize = True
        Me.AddressRangeF.Location = New System.Drawing.Point(19, 442)
        Me.AddressRangeF.Name = "AddressRangeF"
        Me.AddressRangeF.Size = New System.Drawing.Size(148, 17)
        Me.AddressRangeF.TabIndex = 590
        Me.AddressRangeF.Text = "Memory (user adjustable) -"
        Me.AddressRangeF.UseVisualStyleBackColor = True
        '
        'AddressRangeB
        '
        Me.AddressRangeB.AutoSize = True
        Me.AddressRangeB.Location = New System.Drawing.Point(19, 420)
        Me.AddressRangeB.Name = "AddressRangeB"
        Me.AddressRangeB.Size = New System.Drawing.Size(362, 17)
        Me.AddressRangeB.TabIndex = 584
        Me.AddressRangeB.Text = "Cal Ram (new map) - Address Range: 0x5000 (20480) - 0x57FF (22527)"
        Me.AddressRangeB.UseVisualStyleBackColor = True
        '
        'AddressRangeA
        '
        Me.AddressRangeA.AutoSize = True
        Me.AddressRangeA.Checked = True
        Me.AddressRangeA.Location = New System.Drawing.Point(19, 398)
        Me.AddressRangeA.Name = "AddressRangeA"
        Me.AddressRangeA.Size = New System.Drawing.Size(308, 17)
        Me.AddressRangeA.TabIndex = 583
        Me.AddressRangeA.TabStop = True
        Me.AddressRangeA.Text = "Cal Ram (old map) - Address Range: 0x40 (64) - 0x1FF (511)"
        Me.AddressRangeA.UseVisualStyleBackColor = True
        '
        'TextBoxCalRamFile3457A
        '
        Me.TextBoxCalRamFile3457A.Location = New System.Drawing.Point(20, 541)
        Me.TextBoxCalRamFile3457A.Name = "TextBoxCalRamFile3457A"
        Me.TextBoxCalRamFile3457A.ReadOnly = True
        Me.TextBoxCalRamFile3457A.Size = New System.Drawing.Size(701, 20)
        Me.TextBoxCalRamFile3457A.TabIndex = 581
        '
        'Label125
        '
        Me.Label125.AutoSize = True
        Me.Label125.Location = New System.Drawing.Point(19, 523)
        Me.Label125.Name = "Label125"
        Me.Label125.Size = New System.Drawing.Size(92, 13)
        Me.Label125.TabIndex = 580
        Me.Label125.Text = "File is saved here:"
        '
        'ButtonCalramDump3457A
        '
        Me.ButtonCalramDump3457A.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonCalramDump3457A.Location = New System.Drawing.Point(19, 473)
        Me.ButtonCalramDump3457A.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonCalramDump3457A.Name = "ButtonCalramDump3457A"
        Me.ButtonCalramDump3457A.Size = New System.Drawing.Size(100, 37)
        Me.ButtonCalramDump3457A.TabIndex = 571
        Me.ButtonCalramDump3457A.Text = "3457A Read"
        Me.ButtonCalramDump3457A.UseVisualStyleBackColor = True
        '
        'Label122
        '
        Me.Label122.AutoSize = True
        Me.Label122.Location = New System.Drawing.Point(16, 370)
        Me.Label122.Name = "Label122"
        Me.Label122.Size = New System.Drawing.Size(282, 13)
        Me.Label122.TabIndex = 568
        Me.Label122.Text = "Connect Device 1 to your 3457A (leave in STOP position)."
        '
        'Label123
        '
        Me.Label123.AutoSize = True
        Me.Label123.Location = New System.Drawing.Point(16, 353)
        Me.Label123.Name = "Label123"
        Me.Label123.Size = New System.Drawing.Size(332, 13)
        Me.Label123.TabIndex = 567
        Me.Label123.Text = "This utility will read the CALRAM NVRam contents and dump it to file."
        '
        'Label116
        '
        Me.Label116.AutoSize = True
        Me.Label116.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label116.Location = New System.Drawing.Point(16, 331)
        Me.Label116.Name = "Label116"
        Me.Label116.Size = New System.Drawing.Size(265, 15)
        Me.Label116.TabIndex = 566
        Me.Label116.Text = "HP 3457A CALIBRATION RAM EXTRACT:"
        '
        'TextBoxCalRamFile
        '
        Me.TextBoxCalRamFile.Location = New System.Drawing.Point(20, 243)
        Me.TextBoxCalRamFile.Name = "TextBoxCalRamFile"
        Me.TextBoxCalRamFile.ReadOnly = True
        Me.TextBoxCalRamFile.Size = New System.Drawing.Size(701, 20)
        Me.TextBoxCalRamFile.TabIndex = 557
        '
        'Label62
        '
        Me.Label62.AutoSize = True
        Me.Label62.Location = New System.Drawing.Point(19, 225)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(93, 13)
        Me.Label62.TabIndex = 555
        Me.Label62.Text = "File(s) saved here:"
        '
        'Label63
        '
        Me.Label63.AutoSize = True
        Me.Label63.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label63.Location = New System.Drawing.Point(16, 16)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(265, 15)
        Me.Label63.TabIndex = 550
        Me.Label63.Text = "HP 3458A CALIBRATION RAM EXTRACT:"
        '
        'Label60
        '
        Me.Label60.AutoSize = True
        Me.Label60.Location = New System.Drawing.Point(16, 57)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(282, 13)
        Me.Label60.TabIndex = 546
        Me.Label60.Text = "Connect Device 1 to your 3458A (leave in STOP position)."
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Location = New System.Drawing.Point(16, 40)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(391, 13)
        Me.Label38.TabIndex = 24
        Me.Label38.Text = "This utility will read the CALRAM/SETTINGS NVRam contents and dump it to file."
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.BtnSave3458A)
        Me.GroupBox6.Controls.Add(Me.Label318)
        Me.GroupBox6.Controls.Add(Me.Label315)
        Me.GroupBox6.Controls.Add(Me.CalRam3458APreRun)
        Me.GroupBox6.Controls.Add(Me.LabelCalRamAddressHex)
        Me.GroupBox6.Controls.Add(Me.TextBoxCalRamFile2)
        Me.GroupBox6.Controls.Add(Me.ButtonCalramDump3458A)
        Me.GroupBox6.Controls.Add(Me.Label249)
        Me.GroupBox6.Controls.Add(Me.Button3458Aabort)
        Me.GroupBox6.Controls.Add(Me.Label141)
        Me.GroupBox6.Controls.Add(Me.Label140)
        Me.GroupBox6.Controls.Add(Me.Label139)
        Me.GroupBox6.Controls.Add(Me.Label138)
        Me.GroupBox6.Controls.Add(Me.Label64)
        Me.GroupBox6.Controls.Add(Me.AddressRangeD)
        Me.GroupBox6.Controls.Add(Me.AddressRangeC)
        Me.GroupBox6.Controls.Add(Me.LabelCounter)
        Me.GroupBox6.Controls.Add(Me.PictureBox3)
        Me.GroupBox6.Controls.Add(Me.Label103)
        Me.GroupBox6.Controls.Add(Me.Label68)
        Me.GroupBox6.Controls.Add(Me.LabelCalRamAddress)
        Me.GroupBox6.Controls.Add(Me.LabelByte)
        Me.GroupBox6.Controls.Add(Me.Label102)
        Me.GroupBox6.Controls.Add(Me.Label61)
        Me.GroupBox6.Controls.Add(Me.LabelCalRamByte)
        Me.GroupBox6.Controls.Add(Me.CalramStatus)
        Me.GroupBox6.Enabled = False
        Me.GroupBox6.Location = New System.Drawing.Point(8, 4)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(913, 307)
        Me.GroupBox6.TabIndex = 588
        Me.GroupBox6.TabStop = False
        '
        'Label318
        '
        Me.Label318.AutoSize = True
        Me.Label318.Location = New System.Drawing.Point(676, 113)
        Me.Label318.Name = "Label318"
        Me.Label318.Size = New System.Drawing.Size(221, 52)
        Me.Label318.TabIndex = 602
        Me.Label318.Text = "On '3458A READ' the Device 1 STB Mask" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "&& Polling checkbox are configured && Pre-" &
    "Run" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Commands sent before data is pulled from the" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "3458A."
        '
        'Label315
        '
        Me.Label315.AutoSize = True
        Me.Label315.Location = New System.Drawing.Point(227, 146)
        Me.Label315.Name = "Label315"
        Me.Label315.Size = New System.Drawing.Size(104, 13)
        Me.Label315.TabIndex = 601
        Me.Label315.Text = "Pre-Run Commands:"
        '
        'LabelCalRamAddressHex
        '
        Me.LabelCalRamAddressHex.AutoSize = True
        Me.LabelCalRamAddressHex.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCalRamAddressHex.Location = New System.Drawing.Point(570, 216)
        Me.LabelCalRamAddressHex.Name = "LabelCalRamAddressHex"
        Me.LabelCalRamAddressHex.Size = New System.Drawing.Size(13, 13)
        Me.LabelCalRamAddressHex.TabIndex = 598
        Me.LabelCalRamAddressHex.Text = "0"
        '
        'TextBoxCalRamFile2
        '
        Me.TextBoxCalRamFile2.Location = New System.Drawing.Point(12, 261)
        Me.TextBoxCalRamFile2.Name = "TextBoxCalRamFile2"
        Me.TextBoxCalRamFile2.ReadOnly = True
        Me.TextBoxCalRamFile2.Size = New System.Drawing.Size(701, 20)
        Me.TextBoxCalRamFile2.TabIndex = 591
        '
        'ButtonCalramDump3458A
        '
        Me.ButtonCalramDump3458A.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonCalramDump3458A.Location = New System.Drawing.Point(11, 148)
        Me.ButtonCalramDump3458A.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonCalramDump3458A.Name = "ButtonCalramDump3458A"
        Me.ButtonCalramDump3458A.Size = New System.Drawing.Size(100, 37)
        Me.ButtonCalramDump3458A.TabIndex = 591
        Me.ButtonCalramDump3458A.Text = "3458A Read"
        Me.ButtonCalramDump3458A.UseVisualStyleBackColor = True
        '
        'Label249
        '
        Me.Label249.AutoSize = True
        Me.Label249.Location = New System.Drawing.Point(28, 121)
        Me.Label249.Name = "Label249"
        Me.Label249.Size = New System.Drawing.Size(216, 13)
        Me.Label249.TabIndex = 591
        Me.Label249.Text = "Two .bin files are created, U121(L) & U122(U)"
        '
        'Button3458Aabort
        '
        Me.Button3458Aabort.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3458Aabort.Location = New System.Drawing.Point(126, 148)
        Me.Button3458Aabort.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button3458Aabort.Name = "Button3458Aabort"
        Me.Button3458Aabort.Size = New System.Drawing.Size(48, 37)
        Me.Button3458Aabort.TabIndex = 590
        Me.Button3458Aabort.Text = "Abort"
        Me.Button3458Aabort.UseVisualStyleBackColor = True
        '
        'Label141
        '
        Me.Label141.AutoSize = True
        Me.Label141.Location = New System.Drawing.Point(676, 91)
        Me.Label141.Name = "Label141"
        Me.Label141.Size = New System.Drawing.Size(151, 13)
        Me.Label141.TabIndex = 597
        Me.Label141.Text = "Size = 32k (32768 bytes) each"
        '
        'Label140
        '
        Me.Label140.AutoSize = True
        Me.Label140.Location = New System.Drawing.Point(676, 55)
        Me.Label140.Name = "Label140"
        Me.Label140.Size = New System.Drawing.Size(112, 13)
        Me.Label140.TabIndex = 596
        Me.Label140.Text = "Size = 2k (2048 bytes)"
        '
        'Label139
        '
        Me.Label139.AutoSize = True
        Me.Label139.Location = New System.Drawing.Point(676, 76)
        Me.Label139.Name = "Label139"
        Me.Label139.Size = New System.Drawing.Size(223, 13)
        Me.Label139.TabIndex = 595
        Me.Label139.Text = "Settings Ram 1/2 = DS1235 (DS1230Y-120+)"
        '
        'Label138
        '
        Me.Label138.AutoSize = True
        Me.Label138.Location = New System.Drawing.Point(676, 39)
        Me.Label138.Name = "Label138"
        Me.Label138.Size = New System.Drawing.Size(195, 13)
        Me.Label138.TabIndex = 594
        Me.Label138.Text = "Cal Ram = DS1220Y (DS1220AD-150+)"
        '
        'Label64
        '
        Me.Label64.AutoSize = True
        Me.Label64.Location = New System.Drawing.Point(676, 19)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(85, 13)
        Me.Label64.TabIndex = 590
        Me.Label64.Text = "INFORMATION:"
        '
        'AddressRangeD
        '
        Me.AddressRangeD.AutoSize = True
        Me.AddressRangeD.Location = New System.Drawing.Point(11, 102)
        Me.AddressRangeD.Name = "AddressRangeD"
        Me.AddressRangeD.Size = New System.Drawing.Size(321, 17)
        Me.AddressRangeD.TabIndex = 592
        Me.AddressRangeD.Text = "Settings Ram 1 and 2 - Address Range: 0x120000  - 0x12FFFF"
        Me.AddressRangeD.UseVisualStyleBackColor = True
        '
        'AddressRangeC
        '
        Me.AddressRangeC.AutoSize = True
        Me.AddressRangeC.Checked = True
        Me.AddressRangeC.Location = New System.Drawing.Point(11, 80)
        Me.AddressRangeC.Name = "AddressRangeC"
        Me.AddressRangeC.Size = New System.Drawing.Size(244, 17)
        Me.AddressRangeC.TabIndex = 591
        Me.AddressRangeC.TabStop = True
        Me.AddressRangeC.Text = "Cal Ram - Address Range: 0x60000 - 0x60FFF"
        Me.AddressRangeC.UseVisualStyleBackColor = True
        '
        'LabelCounter
        '
        Me.LabelCounter.AutoSize = True
        Me.LabelCounter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCounter.Location = New System.Drawing.Point(570, 197)
        Me.LabelCounter.Name = "LabelCounter"
        Me.LabelCounter.Size = New System.Drawing.Size(13, 13)
        Me.LabelCounter.TabIndex = 564
        Me.LabelCounter.Text = "0"
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(471, 19)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(201, 109)
        Me.PictureBox3.TabIndex = 590
        Me.PictureBox3.TabStop = False
        '
        'Label103
        '
        Me.Label103.AutoSize = True
        Me.Label103.Location = New System.Drawing.Point(469, 197)
        Me.Label103.Name = "Label103"
        Me.Label103.Size = New System.Drawing.Size(71, 13)
        Me.Label103.TabIndex = 563
        Me.Label103.Text = "Byte Counter:"
        '
        'Label68
        '
        Me.Label68.AutoSize = True
        Me.Label68.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label68.ForeColor = System.Drawing.Color.Red
        Me.Label68.Location = New System.Drawing.Point(3, 286)
        Me.Label68.Name = "Label68"
        Me.Label68.Size = New System.Drawing.Size(296, 16)
        Me.Label68.TabIndex = 565
        Me.Label68.Text = "This is experimental, please use at your own risk."
        '
        'LabelCalRamAddress
        '
        Me.LabelCalRamAddress.AutoSize = True
        Me.LabelCalRamAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCalRamAddress.Location = New System.Drawing.Point(644, 216)
        Me.LabelCalRamAddress.Name = "LabelCalRamAddress"
        Me.LabelCalRamAddress.Size = New System.Drawing.Size(49, 13)
        Me.LabelCalRamAddress.TabIndex = 562
        Me.LabelCalRamAddress.Text = "######"
        Me.LabelCalRamAddress.Visible = False
        '
        'LabelByte
        '
        Me.LabelByte.AutoSize = True
        Me.LabelByte.Location = New System.Drawing.Point(746, 213)
        Me.LabelByte.Name = "LabelByte"
        Me.LabelByte.Size = New System.Drawing.Size(42, 13)
        Me.LabelByte.TabIndex = 559
        Me.LabelByte.Text = "Byte(s):"
        Me.LabelByte.Visible = False
        '
        'Label102
        '
        Me.Label102.AutoSize = True
        Me.Label102.Location = New System.Drawing.Point(469, 216)
        Me.Label102.Name = "Label102"
        Me.Label102.Size = New System.Drawing.Size(48, 13)
        Me.Label102.TabIndex = 561
        Me.Label102.Text = "Address:"
        '
        'Label61
        '
        Me.Label61.AutoSize = True
        Me.Label61.Location = New System.Drawing.Point(469, 177)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(40, 13)
        Me.Label61.TabIndex = 548
        Me.Label61.Text = "Status:"
        '
        'LabelCalRamByte
        '
        Me.LabelCalRamByte.AutoSize = True
        Me.LabelCalRamByte.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCalRamByte.Location = New System.Drawing.Point(847, 213)
        Me.LabelCalRamByte.Name = "LabelCalRamByte"
        Me.LabelCalRamByte.Size = New System.Drawing.Size(35, 13)
        Me.LabelCalRamByte.TabIndex = 560
        Me.LabelCalRamByte.Text = "####"
        Me.LabelCalRamByte.Visible = False
        '
        'CalramStatus
        '
        Me.CalramStatus.AutoSize = True
        Me.CalramStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalramStatus.Location = New System.Drawing.Point(570, 177)
        Me.CalramStatus.Name = "CalramStatus"
        Me.CalramStatus.Size = New System.Drawing.Size(44, 13)
        Me.CalramStatus.TabIndex = 551
        Me.CalramStatus.Text = "READY"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.LabelCalRamAddress3457AHex)
        Me.GroupBox7.Controls.Add(Me.Label148)
        Me.GroupBox7.Controls.Add(Me.TextBox3457ATo)
        Me.GroupBox7.Controls.Add(Me.Label147)
        Me.GroupBox7.Controls.Add(Me.Label146)
        Me.GroupBox7.Controls.Add(Me.TextBox3457AFrom)
        Me.GroupBox7.Controls.Add(Me.Label142)
        Me.GroupBox7.Controls.Add(Me.Button3457Aabort)
        Me.GroupBox7.Controls.Add(Me.Label143)
        Me.GroupBox7.Controls.Add(Me.PictureBox4)
        Me.GroupBox7.Controls.Add(Me.Label144)
        Me.GroupBox7.Controls.Add(Me.Label127)
        Me.GroupBox7.Controls.Add(Me.Label131)
        Me.GroupBox7.Controls.Add(Me.LabelCounter3457A)
        Me.GroupBox7.Controls.Add(Me.Label129)
        Me.GroupBox7.Controls.Add(Me.Label126)
        Me.GroupBox7.Controls.Add(Me.Label130)
        Me.GroupBox7.Controls.Add(Me.LabelCalRamAddress3457A)
        Me.GroupBox7.Controls.Add(Me.Label132)
        Me.GroupBox7.Controls.Add(Me.Label128)
        Me.GroupBox7.Controls.Add(Me.CalramStatus3457A)
        Me.GroupBox7.Controls.Add(Me.LabelCalRamByte3457A)
        Me.GroupBox7.Enabled = False
        Me.GroupBox7.Location = New System.Drawing.Point(8, 317)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(913, 270)
        Me.GroupBox7.TabIndex = 589
        Me.GroupBox7.TabStop = False
        '
        'LabelCalRamAddress3457AHex
        '
        Me.LabelCalRamAddress3457AHex.AutoSize = True
        Me.LabelCalRamAddress3457AHex.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCalRamAddress3457AHex.Location = New System.Drawing.Point(544, 178)
        Me.LabelCalRamAddress3457AHex.Name = "LabelCalRamAddress3457AHex"
        Me.LabelCalRamAddress3457AHex.Size = New System.Drawing.Size(49, 13)
        Me.LabelCalRamAddress3457AHex.TabIndex = 605
        Me.LabelCalRamAddress3457AHex.Text = "######"
        '
        'Label148
        '
        Me.Label148.AutoSize = True
        Me.Label148.Location = New System.Drawing.Point(677, 74)
        Me.Label148.Name = "Label148"
        Me.Label148.Size = New System.Drawing.Size(170, 13)
        Me.Label148.TabIndex = 604
        Me.Label148.Text = "User adjustable range: 0 to 32767."
        '
        'TextBox3457ATo
        '
        Me.TextBox3457ATo.Location = New System.Drawing.Point(367, 126)
        Me.TextBox3457ATo.Name = "TextBox3457ATo"
        Me.TextBox3457ATo.Size = New System.Drawing.Size(41, 20)
        Me.TextBox3457ATo.TabIndex = 603
        Me.TextBox3457ATo.Text = "32767"
        '
        'Label147
        '
        Me.Label147.AutoSize = True
        Me.Label147.Location = New System.Drawing.Point(342, 128)
        Me.Label147.Name = "Label147"
        Me.Label147.Size = New System.Drawing.Size(23, 13)
        Me.Label147.TabIndex = 602
        Me.Label147.Text = "To:"
        '
        'Label146
        '
        Me.Label146.AutoSize = True
        Me.Label146.Location = New System.Drawing.Point(170, 128)
        Me.Label146.Name = "Label146"
        Me.Label146.Size = New System.Drawing.Size(118, 13)
        Me.Label146.TabIndex = 601
        Me.Label146.Text = "Address Range:   From:"
        '
        'TextBox3457AFrom
        '
        Me.TextBox3457AFrom.Location = New System.Drawing.Point(289, 126)
        Me.TextBox3457AFrom.Name = "TextBox3457AFrom"
        Me.TextBox3457AFrom.Size = New System.Drawing.Size(41, 20)
        Me.TextBox3457AFrom.TabIndex = 590
        Me.TextBox3457AFrom.Text = "0"
        '
        'Label142
        '
        Me.Label142.AutoSize = True
        Me.Label142.Location = New System.Drawing.Point(677, 56)
        Me.Label142.Name = "Label142"
        Me.Label142.Size = New System.Drawing.Size(235, 13)
        Me.Label142.TabIndex = 600
        Me.Label142.Text = "to older ones, therefore different address ranges."
        '
        'Button3457Aabort
        '
        Me.Button3457Aabort.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3457Aabort.Location = New System.Drawing.Point(126, 156)
        Me.Button3457Aabort.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button3457Aabort.Name = "Button3457Aabort"
        Me.Button3457Aabort.Size = New System.Drawing.Size(48, 37)
        Me.Button3457Aabort.TabIndex = 590
        Me.Button3457Aabort.Text = "Abort"
        Me.Button3457Aabort.UseVisualStyleBackColor = True
        '
        'Label143
        '
        Me.Label143.AutoSize = True
        Me.Label143.Location = New System.Drawing.Point(676, 40)
        Me.Label143.Name = "Label143"
        Me.Label143.Size = New System.Drawing.Size(218, 13)
        Me.Label143.TabIndex = 599
        Me.Label143.Text = "Newer 3457A's have a different memory map"
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(472, 19)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(200, 100)
        Me.PictureBox4.TabIndex = 590
        Me.PictureBox4.TabStop = False
        '
        'Label144
        '
        Me.Label144.AutoSize = True
        Me.Label144.Location = New System.Drawing.Point(676, 19)
        Me.Label144.Name = "Label144"
        Me.Label144.Size = New System.Drawing.Size(85, 13)
        Me.Label144.TabIndex = 598
        Me.Label144.Text = "INFORMATION:"
        '
        'Label127
        '
        Me.Label127.AutoSize = True
        Me.Label127.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label127.ForeColor = System.Drawing.Color.Red
        Me.Label127.Location = New System.Drawing.Point(3, 250)
        Me.Label127.Name = "Label127"
        Me.Label127.Size = New System.Drawing.Size(296, 16)
        Me.Label127.TabIndex = 585
        Me.Label127.Text = "This is experimental, please use at your own risk."
        '
        'Label131
        '
        Me.Label131.AutoSize = True
        Me.Label131.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label131.Location = New System.Drawing.Point(183, 192)
        Me.Label131.Name = "Label131"
        Me.Label131.Size = New System.Drawing.Size(14, 13)
        Me.Label131.TabIndex = 587
        Me.Label131.Text = "#"
        Me.Label131.Visible = False
        '
        'LabelCounter3457A
        '
        Me.LabelCounter3457A.AutoSize = True
        Me.LabelCounter3457A.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCounter3457A.Location = New System.Drawing.Point(544, 159)
        Me.LabelCounter3457A.Name = "LabelCounter3457A"
        Me.LabelCounter3457A.Size = New System.Drawing.Size(13, 13)
        Me.LabelCounter3457A.TabIndex = 579
        Me.LabelCounter3457A.Text = "0"
        '
        'Label129
        '
        Me.Label129.AutoSize = True
        Me.Label129.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label129.Location = New System.Drawing.Point(183, 168)
        Me.Label129.Name = "Label129"
        Me.Label129.Size = New System.Drawing.Size(14, 13)
        Me.Label129.TabIndex = 586
        Me.Label129.Text = "#"
        Me.Label129.Visible = False
        '
        'Label126
        '
        Me.Label126.AutoSize = True
        Me.Label126.Location = New System.Drawing.Point(469, 159)
        Me.Label126.Name = "Label126"
        Me.Label126.Size = New System.Drawing.Size(71, 13)
        Me.Label126.TabIndex = 578
        Me.Label126.Text = "Byte Counter:"
        '
        'Label130
        '
        Me.Label130.AutoSize = True
        Me.Label130.Location = New System.Drawing.Point(469, 198)
        Me.Label130.Name = "Label130"
        Me.Label130.Size = New System.Drawing.Size(36, 13)
        Me.Label130.TabIndex = 574
        Me.Label130.Text = "Bytes:"
        '
        'LabelCalRamAddress3457A
        '
        Me.LabelCalRamAddress3457A.AutoSize = True
        Me.LabelCalRamAddress3457A.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCalRamAddress3457A.Location = New System.Drawing.Point(623, 178)
        Me.LabelCalRamAddress3457A.Name = "LabelCalRamAddress3457A"
        Me.LabelCalRamAddress3457A.Size = New System.Drawing.Size(49, 13)
        Me.LabelCalRamAddress3457A.TabIndex = 577
        Me.LabelCalRamAddress3457A.Text = "######"
        Me.LabelCalRamAddress3457A.Visible = False
        '
        'Label132
        '
        Me.Label132.AutoSize = True
        Me.Label132.Location = New System.Drawing.Point(469, 139)
        Me.Label132.Name = "Label132"
        Me.Label132.Size = New System.Drawing.Size(40, 13)
        Me.Label132.TabIndex = 572
        Me.Label132.Text = "Status:"
        '
        'Label128
        '
        Me.Label128.AutoSize = True
        Me.Label128.Location = New System.Drawing.Point(468, 178)
        Me.Label128.Name = "Label128"
        Me.Label128.Size = New System.Drawing.Size(48, 13)
        Me.Label128.TabIndex = 576
        Me.Label128.Text = "Address:"
        '
        'CalramStatus3457A
        '
        Me.CalramStatus3457A.AutoSize = True
        Me.CalramStatus3457A.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalramStatus3457A.Location = New System.Drawing.Point(544, 139)
        Me.CalramStatus3457A.Name = "CalramStatus3457A"
        Me.CalramStatus3457A.Size = New System.Drawing.Size(14, 13)
        Me.CalramStatus3457A.TabIndex = 573
        Me.CalramStatus3457A.Text = "#"
        '
        'LabelCalRamByte3457A
        '
        Me.LabelCalRamByte3457A.AutoSize = True
        Me.LabelCalRamByte3457A.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCalRamByte3457A.Location = New System.Drawing.Point(544, 198)
        Me.LabelCalRamByte3457A.Name = "LabelCalRamByte3457A"
        Me.LabelCalRamByte3457A.Size = New System.Drawing.Size(35, 13)
        Me.LabelCalRamByte3457A.TabIndex = 575
        Me.LabelCalRamByte3457A.Text = "####"
        '
        'TabPage11
        '
        Me.TabPage11.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage11.Controls.Add(Me.ShowFilesCalRamR6581)
        Me.TabPage11.Controls.Add(Me.GroupBox10)
        Me.TabPage11.Controls.Add(Me.ButtonR6581abort)
        Me.TabPage11.Location = New System.Drawing.Point(4, 22)
        Me.TabPage11.Name = "TabPage11"
        Me.TabPage11.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage11.TabIndex = 12
        Me.TabPage11.Text = "R6581 Cal Extract  "
        '
        'GroupBox10
        '
        Me.GroupBox10.Controls.Add(Me.ButtonJsonViewer2)
        Me.GroupBox10.Controls.Add(Me.ButtonJsonViewer)
        Me.GroupBox10.Controls.Add(Me.TextBoxR6581GPIBlist)
        Me.GroupBox10.Controls.Add(Me.Label111)
        Me.GroupBox10.Controls.Add(Me.CheckBoxR6581Upload9)
        Me.GroupBox10.Controls.Add(Me.CheckBoxR6581Upload8)
        Me.GroupBox10.Controls.Add(Me.Label305)
        Me.GroupBox10.Controls.Add(Me.Label248)
        Me.GroupBox10.Controls.Add(Me.Label301)
        Me.GroupBox10.Controls.Add(Me.LabelCalRamByte6581upload)
        Me.GroupBox10.Controls.Add(Me.CalramStatus6581upload)
        Me.GroupBox10.Controls.Add(Me.Label243)
        Me.GroupBox10.Controls.Add(Me.Label244)
        Me.GroupBox10.Controls.Add(Me.ButtonR6581commitEEprom)
        Me.GroupBox10.Controls.Add(Me.ButtonR6581upload)
        Me.GroupBox10.Controls.Add(Me.ButtonOpenR6581fileSelectJson)
        Me.GroupBox10.Controls.Add(Me.TextBoxCalRamFileJson6581Select)
        Me.GroupBox10.Controls.Add(Me.CheckBoxR6581Upload5)
        Me.GroupBox10.Controls.Add(Me.CheckBoxR6581Upload6)
        Me.GroupBox10.Controls.Add(Me.CheckBoxR6581Upload7)
        Me.GroupBox10.Controls.Add(Me.CheckBoxR6581Upload2)
        Me.GroupBox10.Controls.Add(Me.CheckBoxR6581Upload3)
        Me.GroupBox10.Controls.Add(Me.CheckBoxR6581Upload4)
        Me.GroupBox10.Controls.Add(Me.CheckBoxR6581Upload1)
        Me.GroupBox10.Controls.Add(Me.Label241)
        Me.GroupBox10.Controls.Add(Me.Label242)
        Me.GroupBox10.Controls.Add(Me.SendRegularConstantsReadR6581)
        Me.GroupBox10.Controls.Add(Me.Label59)
        Me.GroupBox10.Controls.Add(Me.CheckBoxR6581RetrieveREF)
        Me.GroupBox10.Controls.Add(Me.ButtonOpenR6581fileJson)
        Me.GroupBox10.Controls.Add(Me.TextBoxCalRamFileJson6581)
        Me.GroupBox10.Controls.Add(Me.ButtonOpenR6581file)
        Me.GroupBox10.Controls.Add(Me.TextBoxCalRamFile6581)
        Me.GroupBox10.Controls.Add(Me.Label312)
        Me.GroupBox10.Controls.Add(Me.Label309)
        Me.GroupBox10.Controls.Add(Me.Label310)
        Me.GroupBox10.Controls.Add(Me.Label311)
        Me.GroupBox10.Controls.Add(Me.ButtonCalramDumpR6581)
        Me.GroupBox10.Controls.Add(Me.Label245)
        Me.GroupBox10.Controls.Add(Me.Label246)
        Me.GroupBox10.Controls.Add(Me.AllRegularConstantsReadR6581)
        Me.GroupBox10.Controls.Add(Me.PictureBox1)
        Me.GroupBox10.Controls.Add(Me.Label298)
        Me.GroupBox10.Controls.Add(Me.Label304)
        Me.GroupBox10.Controls.Add(Me.Label306)
        Me.GroupBox10.Controls.Add(Me.LabelCalRamByte6581)
        Me.GroupBox10.Controls.Add(Me.CalramStatus6581)
        Me.GroupBox10.Enabled = False
        Me.GroupBox10.Location = New System.Drawing.Point(8, 4)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(913, 588)
        Me.GroupBox10.TabIndex = 589
        Me.GroupBox10.TabStop = False
        '
        'TextBoxR6581GPIBlist
        '
        Me.TextBoxR6581GPIBlist.Location = New System.Drawing.Point(542, 437)
        Me.TextBoxR6581GPIBlist.Multiline = True
        Me.TextBoxR6581GPIBlist.Name = "TextBoxR6581GPIBlist"
        Me.TextBoxR6581GPIBlist.ReadOnly = True
        Me.TextBoxR6581GPIBlist.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxR6581GPIBlist.Size = New System.Drawing.Size(363, 68)
        Me.TextBoxR6581GPIBlist.TabIndex = 635
        '
        'Label111
        '
        Me.Label111.AutoSize = True
        Me.Label111.Location = New System.Drawing.Point(540, 423)
        Me.Label111.Name = "Label111"
        Me.Label111.Size = New System.Drawing.Size(183, 13)
        Me.Label111.TabIndex = 636
        Me.Label111.Text = "Calibration commands sent to R6581:"
        '
        'CheckBoxR6581Upload9
        '
        Me.CheckBoxR6581Upload9.AutoSize = True
        Me.CheckBoxR6581Upload9.Location = New System.Drawing.Point(223, 392)
        Me.CheckBoxR6581Upload9.Name = "CheckBoxR6581Upload9"
        Me.CheckBoxR6581Upload9.Size = New System.Drawing.Size(161, 17)
        Me.CheckBoxR6581Upload9.TabIndex = 633
        Me.CheckBoxR6581Upload9.Text = "CAL:INT:AC:HOSEI (factory)"
        Me.CheckBoxR6581Upload9.UseVisualStyleBackColor = True
        '
        'CheckBoxR6581Upload8
        '
        Me.CheckBoxR6581Upload8.AutoSize = True
        Me.CheckBoxR6581Upload8.Location = New System.Drawing.Point(223, 372)
        Me.CheckBoxR6581Upload8.Name = "CheckBoxR6581Upload8"
        Me.CheckBoxR6581Upload8.Size = New System.Drawing.Size(169, 17)
        Me.CheckBoxR6581Upload8.TabIndex = 632
        Me.CheckBoxR6581Upload8.Text = "CAL:INT:DCV:HOSEI (factory)"
        Me.CheckBoxR6581Upload8.UseVisualStyleBackColor = True
        '
        'Label305
        '
        Me.Label305.AutoSize = True
        Me.Label305.ForeColor = System.Drawing.Color.Red
        Me.Label305.Location = New System.Drawing.Point(188, 544)
        Me.Label305.Name = "Label305"
        Me.Label305.Size = New System.Drawing.Size(399, 26)
        Me.Label305.TabIndex = 631
        Me.Label305.Text = "UPLOAD IS A WORK IN PROGRESS, BUTTONS ETC. WORK BUT NO ACTUAL" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "GPIB COMMANDS ARE " &
    "SENT. FEEL FREE TO PLAY!"
        Me.Label305.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label248
        '
        Me.Label248.AutoSize = True
        Me.Label248.Location = New System.Drawing.Point(221, 452)
        Me.Label248.Name = "Label248"
        Me.Label248.Size = New System.Drawing.Size(51, 13)
        Me.Label248.TabIndex = 629
        Me.Label248.Text = "ID Value:"
        '
        'Label301
        '
        Me.Label301.AutoSize = True
        Me.Label301.Location = New System.Drawing.Point(221, 431)
        Me.Label301.Name = "Label301"
        Me.Label301.Size = New System.Drawing.Size(40, 13)
        Me.Label301.TabIndex = 627
        Me.Label301.Text = "Status:"
        '
        'LabelCalRamByte6581upload
        '
        Me.LabelCalRamByte6581upload.AutoSize = True
        Me.LabelCalRamByte6581upload.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCalRamByte6581upload.Location = New System.Drawing.Point(277, 452)
        Me.LabelCalRamByte6581upload.Name = "LabelCalRamByte6581upload"
        Me.LabelCalRamByte6581upload.Size = New System.Drawing.Size(14, 13)
        Me.LabelCalRamByte6581upload.TabIndex = 630
        Me.LabelCalRamByte6581upload.Text = "#"
        '
        'CalramStatus6581upload
        '
        Me.CalramStatus6581upload.AutoSize = True
        Me.CalramStatus6581upload.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalramStatus6581upload.Location = New System.Drawing.Point(277, 431)
        Me.CalramStatus6581upload.Name = "CalramStatus6581upload"
        Me.CalramStatus6581upload.Size = New System.Drawing.Size(14, 13)
        Me.CalramStatus6581upload.TabIndex = 628
        Me.CalramStatus6581upload.Text = "#"
        '
        'Label243
        '
        Me.Label243.AutoSize = True
        Me.Label243.Location = New System.Drawing.Point(541, 298)
        Me.Label243.Name = "Label243"
        Me.Label243.Size = New System.Drawing.Size(362, 117)
        Me.Label243.TabIndex = 625
        Me.Label243.Text = resources.GetString("Label243.Text")
        '
        'Label244
        '
        Me.Label244.AutoSize = True
        Me.Label244.Location = New System.Drawing.Point(541, 283)
        Me.Label244.Name = "Label244"
        Me.Label244.Size = New System.Drawing.Size(138, 13)
        Me.Label244.TabIndex = 624
        Me.Label244.Text = "UPLOAD INSTRUCTIONS:"
        '
        'ButtonR6581commitEEprom
        '
        Me.ButtonR6581commitEEprom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonR6581commitEEprom.Location = New System.Drawing.Point(10, 544)
        Me.ButtonR6581commitEEprom.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonR6581commitEEprom.Name = "ButtonR6581commitEEprom"
        Me.ButtonR6581commitEEprom.Size = New System.Drawing.Size(143, 24)
        Me.ButtonR6581commitEEprom.TabIndex = 623
        Me.ButtonR6581commitEEprom.Text = "Commit Ram to EEprom"
        Me.ButtonR6581commitEEprom.UseVisualStyleBackColor = True
        '
        'ButtonR6581upload
        '
        Me.ButtonR6581upload.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonR6581upload.Location = New System.Drawing.Point(802, 512)
        Me.ButtonR6581upload.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonR6581upload.Name = "ButtonR6581upload"
        Me.ButtonR6581upload.Size = New System.Drawing.Size(103, 24)
        Me.ButtonR6581upload.TabIndex = 622
        Me.ButtonR6581upload.Text = "Upload to Ram"
        Me.ButtonR6581upload.UseVisualStyleBackColor = True
        '
        'TextBoxCalRamFileJson6581Select
        '
        Me.TextBoxCalRamFileJson6581Select.Location = New System.Drawing.Point(11, 514)
        Me.TextBoxCalRamFileJson6581Select.Name = "TextBoxCalRamFileJson6581Select"
        Me.TextBoxCalRamFileJson6581Select.ReadOnly = True
        Me.TextBoxCalRamFileJson6581Select.Size = New System.Drawing.Size(701, 20)
        Me.TextBoxCalRamFileJson6581Select.TabIndex = 620
        '
        'CheckBoxR6581Upload5
        '
        Me.CheckBoxR6581Upload5.AutoSize = True
        Me.CheckBoxR6581Upload5.Location = New System.Drawing.Point(11, 452)
        Me.CheckBoxR6581Upload5.Name = "CheckBoxR6581Upload5"
        Me.CheckBoxR6581Upload5.Size = New System.Drawing.Size(122, 17)
        Me.CheckBoxR6581Upload5.TabIndex = 619
        Me.CheckBoxR6581Upload5.Text = "CAL:INT:OHM:RAM"
        Me.CheckBoxR6581Upload5.UseVisualStyleBackColor = True
        '
        'CheckBoxR6581Upload6
        '
        Me.CheckBoxR6581Upload6.AutoSize = True
        Me.CheckBoxR6581Upload6.Location = New System.Drawing.Point(11, 472)
        Me.CheckBoxR6581Upload6.Name = "CheckBoxR6581Upload6"
        Me.CheckBoxR6581Upload6.Size = New System.Drawing.Size(111, 17)
        Me.CheckBoxR6581Upload6.TabIndex = 618
        Me.CheckBoxR6581Upload6.Text = "CAL:INT:AC:RAM"
        Me.CheckBoxR6581Upload6.UseVisualStyleBackColor = True
        '
        'CheckBoxR6581Upload7
        '
        Me.CheckBoxR6581Upload7.AutoSize = True
        Me.CheckBoxR6581Upload7.Location = New System.Drawing.Point(11, 492)
        Me.CheckBoxR6581Upload7.Name = "CheckBoxR6581Upload7"
        Me.CheckBoxR6581Upload7.Size = New System.Drawing.Size(119, 17)
        Me.CheckBoxR6581Upload7.TabIndex = 617
        Me.CheckBoxR6581Upload7.Text = "CAL:INT:DCV:RAM"
        Me.CheckBoxR6581Upload7.UseVisualStyleBackColor = True
        '
        'CheckBoxR6581Upload2
        '
        Me.CheckBoxR6581Upload2.AutoSize = True
        Me.CheckBoxR6581Upload2.Location = New System.Drawing.Point(11, 392)
        Me.CheckBoxR6581Upload2.Name = "CheckBoxR6581Upload2"
        Me.CheckBoxR6581Upload2.Size = New System.Drawing.Size(163, 17)
        Me.CheckBoxR6581Upload2.TabIndex = 616
        Me.CheckBoxR6581Upload2.Text = "CAL:EXT:ZERO:REAR:RAM"
        Me.CheckBoxR6581Upload2.UseVisualStyleBackColor = True
        '
        'CheckBoxR6581Upload3
        '
        Me.CheckBoxR6581Upload3.AutoSize = True
        Me.CheckBoxR6581Upload3.Location = New System.Drawing.Point(11, 412)
        Me.CheckBoxR6581Upload3.Name = "CheckBoxR6581Upload3"
        Me.CheckBoxR6581Upload3.Size = New System.Drawing.Size(122, 17)
        Me.CheckBoxR6581Upload3.TabIndex = 615
        Me.CheckBoxR6581Upload3.Text = "CAL:EXT:DCV:RAM"
        Me.CheckBoxR6581Upload3.UseVisualStyleBackColor = True
        '
        'CheckBoxR6581Upload4
        '
        Me.CheckBoxR6581Upload4.AutoSize = True
        Me.CheckBoxR6581Upload4.Location = New System.Drawing.Point(11, 432)
        Me.CheckBoxR6581Upload4.Name = "CheckBoxR6581Upload4"
        Me.CheckBoxR6581Upload4.Size = New System.Drawing.Size(125, 17)
        Me.CheckBoxR6581Upload4.TabIndex = 614
        Me.CheckBoxR6581Upload4.Text = "CAL:EXT:OHM:RAM"
        Me.CheckBoxR6581Upload4.UseVisualStyleBackColor = True
        '
        'CheckBoxR6581Upload1
        '
        Me.CheckBoxR6581Upload1.AutoSize = True
        Me.CheckBoxR6581Upload1.Location = New System.Drawing.Point(11, 372)
        Me.CheckBoxR6581Upload1.Name = "CheckBoxR6581Upload1"
        Me.CheckBoxR6581Upload1.Size = New System.Drawing.Size(170, 17)
        Me.CheckBoxR6581Upload1.TabIndex = 613
        Me.CheckBoxR6581Upload1.Text = "CAL:EXT:ZERO:FRONT:RAM"
        Me.CheckBoxR6581Upload1.UseVisualStyleBackColor = True
        '
        'Label241
        '
        Me.Label241.AutoSize = True
        Me.Label241.Location = New System.Drawing.Point(8, 321)
        Me.Label241.Name = "Label241"
        Me.Label241.Size = New System.Drawing.Size(283, 13)
        Me.Label241.TabIndex = 612
        Me.Label241.Text = "Connect Device 1 to your R6581 (leave in STOP position)."
        '
        'Label242
        '
        Me.Label242.AutoSize = True
        Me.Label242.Location = New System.Drawing.Point(8, 304)
        Me.Label242.Name = "Label242"
        Me.Label242.Size = New System.Drawing.Size(478, 13)
        Me.Label242.TabIndex = 611
        Me.Label242.Text = "This utility will upload the Json file contents to R6581 Ram. Select the section(" &
    "s) you want to upload."
        '
        'SendRegularConstantsReadR6581
        '
        Me.SendRegularConstantsReadR6581.AutoSize = True
        Me.SendRegularConstantsReadR6581.Checked = True
        Me.SendRegularConstantsReadR6581.Location = New System.Drawing.Point(11, 349)
        Me.SendRegularConstantsReadR6581.Name = "SendRegularConstantsReadR6581"
        Me.SendRegularConstantsReadR6581.Size = New System.Drawing.Size(380, 17)
        Me.SendRegularConstantsReadR6581.TabIndex = 610
        Me.SendRegularConstantsReadR6581.TabStop = True
        Me.SendRegularConstantsReadR6581.Text = "Send regular && factory calibration constants (select sections required below)"
        Me.SendRegularConstantsReadR6581.UseVisualStyleBackColor = True
        '
        'Label59
        '
        Me.Label59.AutoSize = True
        Me.Label59.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label59.Location = New System.Drawing.Point(8, 281)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(371, 15)
        Me.Label59.TabIndex = 609
        Me.Label59.Text = "ADVANTEST R6581 CALIBRATION CONSTANTS UPLOAD:"
        '
        'TextBoxCalRamFileJson6581
        '
        Me.TextBoxCalRamFileJson6581.Location = New System.Drawing.Point(12, 240)
        Me.TextBoxCalRamFileJson6581.Name = "TextBoxCalRamFileJson6581"
        Me.TextBoxCalRamFileJson6581.ReadOnly = True
        Me.TextBoxCalRamFileJson6581.Size = New System.Drawing.Size(701, 20)
        Me.TextBoxCalRamFileJson6581.TabIndex = 606
        '
        'TextBoxCalRamFile6581
        '
        Me.TextBoxCalRamFile6581.Location = New System.Drawing.Point(12, 214)
        Me.TextBoxCalRamFile6581.Name = "TextBoxCalRamFile6581"
        Me.TextBoxCalRamFile6581.ReadOnly = True
        Me.TextBoxCalRamFile6581.Size = New System.Drawing.Size(701, 20)
        Me.TextBoxCalRamFile6581.TabIndex = 604
        '
        'Label312
        '
        Me.Label312.AutoSize = True
        Me.Label312.Location = New System.Drawing.Point(11, 196)
        Me.Label312.Name = "Label312"
        Me.Label312.Size = New System.Drawing.Size(93, 13)
        Me.Label312.TabIndex = 603
        Me.Label312.Text = "File(s) saved here:"
        '
        'Label309
        '
        Me.Label309.AutoSize = True
        Me.Label309.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label309.Location = New System.Drawing.Point(8, 12)
        Me.Label309.Name = "Label309"
        Me.Label309.Size = New System.Drawing.Size(394, 15)
        Me.Label309.TabIndex = 602
        Me.Label309.Text = "ADVANTEST R6581 CALIBRATION CONSTANTS DOWNLOAD:"
        '
        'Label310
        '
        Me.Label310.AutoSize = True
        Me.Label310.Location = New System.Drawing.Point(8, 53)
        Me.Label310.Name = "Label310"
        Me.Label310.Size = New System.Drawing.Size(283, 13)
        Me.Label310.TabIndex = 601
        Me.Label310.Text = "Connect Device 1 to your R6581 (leave in STOP position)."
        '
        'Label311
        '
        Me.Label311.AutoSize = True
        Me.Label311.Location = New System.Drawing.Point(8, 36)
        Me.Label311.Name = "Label311"
        Me.Label311.Size = New System.Drawing.Size(366, 13)
        Me.Label311.TabIndex = 600
        Me.Label311.Text = "This utility will download the current calibration constants to a text/JSON file." &
    ""
        '
        'ButtonCalramDumpR6581
        '
        Me.ButtonCalramDumpR6581.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonCalramDumpR6581.Location = New System.Drawing.Point(11, 148)
        Me.ButtonCalramDumpR6581.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonCalramDumpR6581.Name = "ButtonCalramDumpR6581"
        Me.ButtonCalramDumpR6581.Size = New System.Drawing.Size(100, 37)
        Me.ButtonCalramDumpR6581.TabIndex = 591
        Me.ButtonCalramDumpR6581.Text = "R6581 Read"
        Me.ButtonCalramDumpR6581.UseVisualStyleBackColor = True
        '
        'Label245
        '
        Me.Label245.AutoSize = True
        Me.Label245.Location = New System.Drawing.Point(651, 34)
        Me.Label245.Name = "Label245"
        Me.Label245.Size = New System.Drawing.Size(257, 169)
        Me.Label245.TabIndex = 594
        Me.Label245.Text = resources.GetString("Label245.Text")
        '
        'Label246
        '
        Me.Label246.AutoSize = True
        Me.Label246.Location = New System.Drawing.Point(651, 19)
        Me.Label246.Name = "Label246"
        Me.Label246.Size = New System.Drawing.Size(85, 13)
        Me.Label246.TabIndex = 590
        Me.Label246.Text = "INFORMATION:"
        '
        'AllRegularConstantsReadR6581
        '
        Me.AllRegularConstantsReadR6581.AutoSize = True
        Me.AllRegularConstantsReadR6581.Checked = True
        Me.AllRegularConstantsReadR6581.Location = New System.Drawing.Point(11, 84)
        Me.AllRegularConstantsReadR6581.Name = "AllRegularConstantsReadR6581"
        Me.AllRegularConstantsReadR6581.Size = New System.Drawing.Size(257, 17)
        Me.AllRegularConstantsReadR6581.TabIndex = 591
        Me.AllRegularConstantsReadR6581.TabStop = True
        Me.AllRegularConstantsReadR6581.Text = "Retrieve all regular && factory calibration constants"
        Me.AllRegularConstantsReadR6581.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(431, 21)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(201, 102)
        Me.PictureBox1.TabIndex = 590
        Me.PictureBox1.TabStop = False
        '
        'Label298
        '
        Me.Label298.AutoSize = True
        Me.Label298.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label298.ForeColor = System.Drawing.Color.Red
        Me.Label298.Location = New System.Drawing.Point(613, 567)
        Me.Label298.Name = "Label298"
        Me.Label298.Size = New System.Drawing.Size(296, 16)
        Me.Label298.TabIndex = 565
        Me.Label298.Text = "This is experimental, please use at your own risk."
        '
        'Label304
        '
        Me.Label304.AutoSize = True
        Me.Label304.Location = New System.Drawing.Point(131, 171)
        Me.Label304.Name = "Label304"
        Me.Label304.Size = New System.Drawing.Size(51, 13)
        Me.Label304.TabIndex = 559
        Me.Label304.Text = "ID Value:"
        '
        'Label306
        '
        Me.Label306.AutoSize = True
        Me.Label306.Location = New System.Drawing.Point(131, 150)
        Me.Label306.Name = "Label306"
        Me.Label306.Size = New System.Drawing.Size(40, 13)
        Me.Label306.TabIndex = 548
        Me.Label306.Text = "Status:"
        '
        'LabelCalRamByte6581
        '
        Me.LabelCalRamByte6581.AutoSize = True
        Me.LabelCalRamByte6581.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCalRamByte6581.Location = New System.Drawing.Point(187, 171)
        Me.LabelCalRamByte6581.Name = "LabelCalRamByte6581"
        Me.LabelCalRamByte6581.Size = New System.Drawing.Size(14, 13)
        Me.LabelCalRamByte6581.TabIndex = 560
        Me.LabelCalRamByte6581.Text = "#"
        '
        'CalramStatus6581
        '
        Me.CalramStatus6581.AutoSize = True
        Me.CalramStatus6581.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalramStatus6581.Location = New System.Drawing.Point(187, 150)
        Me.CalramStatus6581.Name = "CalramStatus6581"
        Me.CalramStatus6581.Size = New System.Drawing.Size(14, 13)
        Me.CalramStatus6581.TabIndex = 551
        Me.CalramStatus6581.Text = "#"
        '
        'ButtonR6581abort
        '
        Me.ButtonR6581abort.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonR6581abort.Location = New System.Drawing.Point(928, 78)
        Me.ButtonR6581abort.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonR6581abort.Name = "ButtonR6581abort"
        Me.ButtonR6581abort.Size = New System.Drawing.Size(115, 37)
        Me.ButtonR6581abort.TabIndex = 590
        Me.ButtonR6581abort.Text = "Abort"
        Me.ButtonR6581abort.UseVisualStyleBackColor = True
        '
        'TabPage12
        '
        Me.TabPage12.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage12.Controls.Add(Me.Label259)
        Me.TabPage12.Controls.Add(Me.Label260)
        Me.TabPage12.Controls.Add(Me.Label261)
        Me.TabPage12.Controls.Add(Me.GroupBox5)
        Me.TabPage12.Location = New System.Drawing.Point(4, 22)
        Me.TabPage12.Name = "TabPage12"
        Me.TabPage12.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage12.TabIndex = 11
        Me.TabPage12.Text = "3245A Cal  "
        '
        'Label259
        '
        Me.Label259.AutoSize = True
        Me.Label259.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label259.Location = New System.Drawing.Point(16, 16)
        Me.Label259.Name = "Label259"
        Me.Label259.Size = New System.Drawing.Size(247, 15)
        Me.Label259.TabIndex = 591
        Me.Label259.Text = "3245A CALIBRATION ADJUSTMENTS:"
        '
        'Label260
        '
        Me.Label260.AutoSize = True
        Me.Label260.Location = New System.Drawing.Point(16, 66)
        Me.Label260.Name = "Label260"
        Me.Label260.Size = New System.Drawing.Size(316, 13)
        Me.Label260.TabIndex = 590
        Me.Label260.Text = "Execute ACAL DCV or ACAL ALL immediately prior to running Cal."
        '
        'Label261
        '
        Me.Label261.AutoSize = True
        Me.Label261.Location = New System.Drawing.Point(16, 40)
        Me.Label261.Name = "Label261"
        Me.Label261.Size = New System.Drawing.Size(441, 13)
        Me.Label261.TabIndex = 589
        Me.Label261.Text = "This utility will automate the calibration adjustments of an HP 3245A by using a " &
    "3458A DMM."
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Label263)
        Me.GroupBox5.Controls.Add(Me.Timeout3458A)
        Me.GroupBox5.Controls.Add(Me.Label273)
        Me.GroupBox5.Controls.Add(Me.PictureBox8)
        Me.GroupBox5.Controls.Add(Me.Label272)
        Me.GroupBox5.Controls.Add(Me.CheckBoxAZERO)
        Me.GroupBox5.Controls.Add(Me.Label267)
        Me.GroupBox5.Controls.Add(Me.RadioButton3245ADCVDCI)
        Me.GroupBox5.Controls.Add(Me.Label3245AWRI)
        Me.GroupBox5.Controls.Add(Me.Label266)
        Me.GroupBox5.Controls.Add(Me.Button3245Aabort)
        Me.GroupBox5.Controls.Add(Me.Label264)
        Me.GroupBox5.Controls.Add(Me.Label262)
        Me.GroupBox5.Controls.Add(Me.Code3245A)
        Me.GroupBox5.Controls.Add(Me.Label3458123)
        Me.GroupBox5.Controls.Add(Me.Label3458ARDG)
        Me.GroupBox5.Controls.Add(Me.ButtonCal3245A)
        Me.GroupBox5.Controls.Add(Me.Label265)
        Me.GroupBox5.Controls.Add(Me.Label268)
        Me.GroupBox5.Controls.Add(Me.RadioButton3245ADCV)
        Me.GroupBox5.Controls.Add(Me.LabelRDG)
        Me.GroupBox5.Controls.Add(Me.Label270)
        Me.GroupBox5.Controls.Add(Me.Label271)
        Me.GroupBox5.Controls.Add(Me.Label274)
        Me.GroupBox5.Controls.Add(Me.Label275)
        Me.GroupBox5.Controls.Add(Me.Cal3245status)
        Me.GroupBox5.Controls.Add(Me.Label255)
        Me.GroupBox5.Controls.Add(Me.Label269)
        Me.GroupBox5.Controls.Add(Me.PictureBox9)
        Me.GroupBox5.Enabled = False
        Me.GroupBox5.Location = New System.Drawing.Point(8, 4)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(1036, 274)
        Me.GroupBox5.TabIndex = 594
        Me.GroupBox5.TabStop = False
        '
        'Label263
        '
        Me.Label263.AutoSize = True
        Me.Label263.Location = New System.Drawing.Point(236, 164)
        Me.Label263.Name = "Label263"
        Me.Label263.Size = New System.Drawing.Size(101, 13)
        Me.Label263.TabIndex = 691
        Me.Label263.Text = "3458A Timeout (ms)"
        '
        'Timeout3458A
        '
        Me.Timeout3458A.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Timeout3458A.Location = New System.Drawing.Point(187, 160)
        Me.Timeout3458A.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Timeout3458A.MaxLength = 8
        Me.Timeout3458A.Name = "Timeout3458A"
        Me.Timeout3458A.Size = New System.Drawing.Size(46, 22)
        Me.Timeout3458A.TabIndex = 692
        Me.Timeout3458A.Text = "15000"
        '
        'Label273
        '
        Me.Label273.AutoSize = True
        Me.Label273.Location = New System.Drawing.Point(8, 99)
        Me.Label273.Name = "Label273"
        Me.Label273.Size = New System.Drawing.Size(283, 13)
        Me.Label273.TabIndex = 596
        Me.Label273.Text = "Connect Device 2 to your 3245A - Leave in STOP position"
        '
        'PictureBox8
        '
        Me.PictureBox8.Image = CType(resources.GetObject("PictureBox8.Image"), System.Drawing.Image)
        Me.PictureBox8.Location = New System.Drawing.Point(456, 80)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(204, 65)
        Me.PictureBox8.TabIndex = 612
        Me.PictureBox8.TabStop = False
        '
        'Label272
        '
        Me.Label272.AutoSize = True
        Me.Label272.Location = New System.Drawing.Point(8, 81)
        Me.Label272.Name = "Label272"
        Me.Label272.Size = New System.Drawing.Size(283, 13)
        Me.Label272.TabIndex = 595
        Me.Label272.Text = "Connect Device 1 to your 3458A - Leave in STOP position"
        '
        'Label267
        '
        Me.Label267.AutoSize = True
        Me.Label267.Location = New System.Drawing.Point(669, 58)
        Me.Label267.Name = "Label267"
        Me.Label267.Size = New System.Drawing.Size(351, 104)
        Me.Label267.TabIndex = 610
        Me.Label267.Text = resources.GetString("Label267.Text")
        '
        'RadioButton3245ADCVDCI
        '
        Me.RadioButton3245ADCVDCI.AutoSize = True
        Me.RadioButton3245ADCVDCI.Location = New System.Drawing.Point(11, 156)
        Me.RadioButton3245ADCVDCI.Name = "RadioButton3245ADCVDCI"
        Me.RadioButton3245ADCVDCI.Size = New System.Drawing.Size(129, 17)
        Me.RadioButton3245ADCVDCI.TabIndex = 608
        Me.RadioButton3245ADCVDCI.Text = "DCV && DCI Calibration"
        Me.RadioButton3245ADCVDCI.UseVisualStyleBackColor = True
        '
        'Label3245AWRI
        '
        Me.Label3245AWRI.AutoSize = True
        Me.Label3245AWRI.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3245AWRI.Location = New System.Drawing.Point(558, 223)
        Me.Label3245AWRI.Name = "Label3245AWRI"
        Me.Label3245AWRI.Size = New System.Drawing.Size(14, 13)
        Me.Label3245AWRI.TabIndex = 607
        Me.Label3245AWRI.Text = "#"
        '
        'Label266
        '
        Me.Label266.AutoSize = True
        Me.Label266.Location = New System.Drawing.Point(457, 223)
        Me.Label266.Name = "Label266"
        Me.Label266.Size = New System.Drawing.Size(69, 13)
        Me.Label266.TabIndex = 606
        Me.Label266.Text = "3245A Write:"
        '
        'Button3245Aabort
        '
        Me.Button3245Aabort.Enabled = False
        Me.Button3245Aabort.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3245Aabort.Location = New System.Drawing.Point(126, 206)
        Me.Button3245Aabort.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button3245Aabort.Name = "Button3245Aabort"
        Me.Button3245Aabort.Size = New System.Drawing.Size(48, 37)
        Me.Button3245Aabort.TabIndex = 605
        Me.Button3245Aabort.Text = "Abort"
        Me.Button3245Aabort.UseVisualStyleBackColor = True
        '
        'Label264
        '
        Me.Label264.AutoSize = True
        Me.Label264.Location = New System.Drawing.Point(578, 187)
        Me.Label264.Name = "Label264"
        Me.Label264.Size = New System.Drawing.Size(34, 13)
        Me.Label264.TabIndex = 604
        Me.Label264.Text = "of  47"
        '
        'Label262
        '
        Me.Label262.AutoSize = True
        Me.Label262.Location = New System.Drawing.Point(236, 136)
        Me.Label262.Name = "Label262"
        Me.Label262.Size = New System.Drawing.Size(202, 13)
        Me.Label262.TabIndex = 595
        Me.Label262.Text = "3245A Security code (3245 is HP default)"
        '
        'Code3245A
        '
        Me.Code3245A.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Code3245A.Location = New System.Drawing.Point(187, 132)
        Me.Code3245A.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Code3245A.MaxLength = 8
        Me.Code3245A.Name = "Code3245A"
        Me.Code3245A.Size = New System.Drawing.Size(46, 22)
        Me.Code3245A.TabIndex = 602
        Me.Code3245A.Text = "3245"
        '
        'Label3458123
        '
        Me.Label3458123.AutoSize = True
        Me.Label3458123.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3458123.Location = New System.Drawing.Point(686, 205)
        Me.Label3458123.Name = "Label3458123"
        Me.Label3458123.Size = New System.Drawing.Size(14, 13)
        Me.Label3458123.TabIndex = 601
        Me.Label3458123.Text = "#"
        '
        'Label3458ARDG
        '
        Me.Label3458ARDG.AutoSize = True
        Me.Label3458ARDG.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3458ARDG.Location = New System.Drawing.Point(558, 205)
        Me.Label3458ARDG.Name = "Label3458ARDG"
        Me.Label3458ARDG.Size = New System.Drawing.Size(14, 13)
        Me.Label3458ARDG.TabIndex = 598
        Me.Label3458ARDG.Text = "#"
        '
        'ButtonCal3245A
        '
        Me.ButtonCal3245A.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonCal3245A.Location = New System.Drawing.Point(11, 206)
        Me.ButtonCal3245A.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonCal3245A.Name = "ButtonCal3245A"
        Me.ButtonCal3245A.Size = New System.Drawing.Size(100, 37)
        Me.ButtonCal3245A.TabIndex = 591
        Me.ButtonCal3245A.Text = "3245A Calibration" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Adjustments"
        Me.ButtonCal3245A.UseVisualStyleBackColor = True
        '
        'Label265
        '
        Me.Label265.AutoSize = True
        Me.Label265.Location = New System.Drawing.Point(669, 36)
        Me.Label265.Name = "Label265"
        Me.Label265.Size = New System.Drawing.Size(334, 13)
        Me.Label265.TabIndex = 596
        Me.Label265.Text = "3245A Security code only req'd if hardware lockout jumper is in place."
        '
        'Label268
        '
        Me.Label268.AutoSize = True
        Me.Label268.Location = New System.Drawing.Point(669, 17)
        Me.Label268.Name = "Label268"
        Me.Label268.Size = New System.Drawing.Size(35, 13)
        Me.Label268.TabIndex = 590
        Me.Label268.Text = "INFO:"
        '
        'RadioButton3245ADCV
        '
        Me.RadioButton3245ADCV.AutoSize = True
        Me.RadioButton3245ADCV.Checked = True
        Me.RadioButton3245ADCV.Location = New System.Drawing.Point(11, 133)
        Me.RadioButton3245ADCV.Name = "RadioButton3245ADCV"
        Me.RadioButton3245ADCV.Size = New System.Drawing.Size(99, 17)
        Me.RadioButton3245ADCV.TabIndex = 591
        Me.RadioButton3245ADCV.TabStop = True
        Me.RadioButton3245ADCV.Text = "DCV Calibration"
        Me.RadioButton3245ADCV.UseVisualStyleBackColor = True
        '
        'LabelRDG
        '
        Me.LabelRDG.AutoSize = True
        Me.LabelRDG.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelRDG.Location = New System.Drawing.Point(558, 187)
        Me.LabelRDG.Name = "LabelRDG"
        Me.LabelRDG.Size = New System.Drawing.Size(14, 13)
        Me.LabelRDG.TabIndex = 564
        Me.LabelRDG.Text = "#"
        '
        'Label270
        '
        Me.Label270.AutoSize = True
        Me.Label270.Location = New System.Drawing.Point(457, 187)
        Me.Label270.Name = "Label270"
        Me.Label270.Size = New System.Drawing.Size(60, 13)
        Me.Label270.TabIndex = 563
        Me.Label270.Text = "Reading #:"
        '
        'Label271
        '
        Me.Label271.AutoSize = True
        Me.Label271.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label271.ForeColor = System.Drawing.Color.Red
        Me.Label271.Location = New System.Drawing.Point(3, 254)
        Me.Label271.Name = "Label271"
        Me.Label271.Size = New System.Drawing.Size(296, 16)
        Me.Label271.TabIndex = 565
        Me.Label271.Text = "This is experimental, please use at your own risk."
        '
        'Label274
        '
        Me.Label274.AutoSize = True
        Me.Label274.Location = New System.Drawing.Point(457, 205)
        Me.Label274.Name = "Label274"
        Me.Label274.Size = New System.Drawing.Size(70, 13)
        Me.Label274.TabIndex = 561
        Me.Label274.Text = "3458A Read:"
        '
        'Label275
        '
        Me.Label275.AutoSize = True
        Me.Label275.Location = New System.Drawing.Point(457, 169)
        Me.Label275.Name = "Label275"
        Me.Label275.Size = New System.Drawing.Size(92, 13)
        Me.Label275.TabIndex = 548
        Me.Label275.Text = "Command/Status:"
        '
        'Cal3245status
        '
        Me.Cal3245status.AutoSize = True
        Me.Cal3245status.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cal3245status.Location = New System.Drawing.Point(558, 169)
        Me.Cal3245status.Name = "Cal3245status"
        Me.Cal3245status.Size = New System.Drawing.Size(14, 13)
        Me.Cal3245status.TabIndex = 551
        Me.Cal3245status.Text = "#"
        '
        'Label255
        '
        Me.Label255.AutoSize = True
        Me.Label255.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label255.Location = New System.Drawing.Point(463, 58)
        Me.Label255.Name = "Label255"
        Me.Label255.Size = New System.Drawing.Size(23, 24)
        Me.Label255.TabIndex = 689
        Me.Label255.Text = "▲"
        '
        'Label269
        '
        Me.Label269.AutoSize = True
        Me.Label269.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label269.Location = New System.Drawing.Point(629, 58)
        Me.Label269.Name = "Label269"
        Me.Label269.Size = New System.Drawing.Size(23, 24)
        Me.Label269.TabIndex = 690
        Me.Label269.Text = "▼"
        '
        'PictureBox9
        '
        Me.PictureBox9.Image = CType(resources.GetObject("PictureBox9.Image"), System.Drawing.Image)
        Me.PictureBox9.Location = New System.Drawing.Point(456, 14)
        Me.PictureBox9.Name = "PictureBox9"
        Me.PictureBox9.Size = New System.Drawing.Size(204, 52)
        Me.PictureBox9.TabIndex = 600
        Me.PictureBox9.TabStop = False
        '
        'TabPage5
        '
        Me.TabPage5.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage5.Controls.Add(Me.OnOffLed4)
        Me.TabPage5.Controls.Add(Me.ButtonRefreshPorts1)
        Me.TabPage5.Controls.Add(Me.Label145)
        Me.TabPage5.Controls.Add(Me.Label79)
        Me.TabPage5.Controls.Add(Me.Label81)
        Me.TabPage5.Controls.Add(Me.Label232)
        Me.TabPage5.Controls.Add(Me.Label233)
        Me.TabPage5.Controls.Add(Me.Label154)
        Me.TabPage5.Controls.Add(Me.Label153)
        Me.TabPage5.Controls.Add(Me.Label152)
        Me.TabPage5.Controls.Add(Me.Label151)
        Me.TabPage5.Controls.Add(Me.LabelDeltaV)
        Me.TabPage5.Controls.Add(Me.volts10)
        Me.TabPage5.Controls.Add(Me.volts9)
        Me.TabPage5.Controls.Add(Me.volts8)
        Me.TabPage5.Controls.Add(Me.volts7)
        Me.TabPage5.Controls.Add(Me.volts6)
        Me.TabPage5.Controls.Add(Me.volts5)
        Me.TabPage5.Controls.Add(Me.volts4)
        Me.TabPage5.Controls.Add(Me.volts3)
        Me.TabPage5.Controls.Add(Me.volts2)
        Me.TabPage5.Controls.Add(Me.volts1)
        Me.TabPage5.Controls.Add(Me.volts0)
        Me.TabPage5.Controls.Add(Me.Label93)
        Me.TabPage5.Controls.Add(Me.LabeldacSpan9Cal)
        Me.TabPage5.Controls.Add(Me.DacSpan9)
        Me.TabPage5.Controls.Add(Me.Label89)
        Me.TabPage5.Controls.Add(Me.LabeldacSpan8Cal)
        Me.TabPage5.Controls.Add(Me.DacSpan8)
        Me.TabPage5.Controls.Add(Me.Label85)
        Me.TabPage5.Controls.Add(Me.LabeldacSpan7Cal)
        Me.TabPage5.Controls.Add(Me.DacSpan7)
        Me.TabPage5.Controls.Add(Me.Label86)
        Me.TabPage5.Controls.Add(Me.LabeldacSpan6Cal)
        Me.TabPage5.Controls.Add(Me.DacSpan6)
        Me.TabPage5.Controls.Add(Me.Label90)
        Me.TabPage5.Controls.Add(Me.LabeldacSpan5Cal)
        Me.TabPage5.Controls.Add(Me.DacSpan5)
        Me.TabPage5.Controls.Add(Me.Label94)
        Me.TabPage5.Controls.Add(Me.LabeldacSpan4Cal)
        Me.TabPage5.Controls.Add(Me.DacSpan4)
        Me.TabPage5.Controls.Add(Me.Label96)
        Me.TabPage5.Controls.Add(Me.LabeldacSpan3Cal)
        Me.TabPage5.Controls.Add(Me.DacSpan3)
        Me.TabPage5.Controls.Add(Me.Label98)
        Me.TabPage5.Controls.Add(Me.LabeldacSpan2Cal)
        Me.TabPage5.Controls.Add(Me.DacSpan2)
        Me.TabPage5.Controls.Add(Me.Label100)
        Me.TabPage5.Controls.Add(Me.LabeldacSpan1Cal)
        Me.TabPage5.Controls.Add(Me.DacSpan1)
        Me.TabPage5.Controls.Add(Me.LabelBatteryMonICMult)
        Me.TabPage5.Controls.Add(Me.Label106)
        Me.TabPage5.Controls.Add(Me.ChargeI)
        Me.TabPage5.Controls.Add(Me.Label107)
        Me.TabPage5.Controls.Add(Me.Label108)
        Me.TabPage5.Controls.Add(Me.Label109)
        Me.TabPage5.Controls.Add(Me.Label110)
        Me.TabPage5.Controls.Add(Me.LabelBatteryVFeedMult)
        Me.TabPage5.Controls.Add(Me.LabelOutputVFeedMult)
        Me.TabPage5.Controls.Add(Me.LabeldacSpan0Cal)
        Me.TabPage5.Controls.Add(Me.LabeldacZero0Cal)
        Me.TabPage5.Controls.Add(Me.BattVdc)
        Me.TabPage5.Controls.Add(Me.OutVdc)
        Me.TabPage5.Controls.Add(Me.DacSpan0)
        Me.TabPage5.Controls.Add(Me.DacZero0)
        Me.TabPage5.Controls.Add(Me.Label57)
        Me.TabPage5.Controls.Add(Me.TextBoxdegC)
        Me.TabPage5.Controls.Add(Me.Label97)
        Me.TabPage5.Controls.Add(Me.TextBoxSer)
        Me.TabPage5.Controls.Add(Me.Label95)
        Me.TabPage5.Controls.Add(Me.TextBoxSOAK)
        Me.TabPage5.Controls.Add(Me.TextBoxSERIAL)
        Me.TabPage5.Controls.Add(Me.TextBoxFULLMA)
        Me.TabPage5.Controls.Add(Me.TextBoxOLMA)
        Me.TabPage5.Controls.Add(Me.TextBoxCENABLE)
        Me.TabPage5.Controls.Add(Me.TextBoxLOWSHUT)
        Me.TabPage5.Controls.Add(Me.TextBoxDC)
        Me.TabPage5.Controls.Add(Me.TextBoxCD)
        Me.TabPage5.Controls.Add(Me.Label92)
        Me.TabPage5.Controls.Add(Me.Label91)
        Me.TabPage5.Controls.Add(Me.Label88)
        Me.TabPage5.Controls.Add(Me.Label87)
        Me.TabPage5.Controls.Add(Me.Label84)
        Me.TabPage5.Controls.Add(Me.Label82)
        Me.TabPage5.Controls.Add(Me.Label78)
        Me.TabPage5.Controls.Add(Me.Label77)
        Me.TabPage5.Controls.Add(Me.Label114)
        Me.TabPage5.Controls.Add(Me.Label113)
        Me.TabPage5.Controls.Add(Me.Label112)
        Me.TabPage5.Controls.Add(Me.txtr1aBIG)
        Me.TabPage5.Controls.Add(Me.comPort_ComboBox)
        Me.TabPage5.Controls.Add(Me.connect_BTN)
        Me.TabPage5.Controls.Add(Me.ButtonKeyVoltage)
        Me.TabPage5.Controls.Add(Me.KeyVoltage)
        Me.TabPage5.Controls.Add(Me.Label80)
        Me.TabPage5.Controls.Add(Me.LabelKeyVoltage)
        Me.TabPage5.Controls.Add(Me.GroupBox2)
        Me.TabPage5.Controls.Add(Me.OnOffLed3)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "PDVS2mini  "
        '
        'Label145
        '
        Me.Label145.AutoSize = True
        Me.Label145.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label145.Location = New System.Drawing.Point(754, 3)
        Me.Label145.Name = "Label145"
        Me.Label145.Size = New System.Drawing.Size(56, 12)
        Me.Label145.TabIndex = 800
        Me.Label145.Text = "PDVS2mini/"
        '
        'Label79
        '
        Me.Label79.AutoSize = True
        Me.Label79.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label79.Location = New System.Drawing.Point(894, 28)
        Me.Label79.Name = "Label79"
        Me.Label79.Size = New System.Drawing.Size(73, 20)
        Me.Label79.TabIndex = 529
        Me.Label79.Text = "OUTPUT"
        '
        'Label81
        '
        Me.Label81.AutoSize = True
        Me.Label81.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label81.Location = New System.Drawing.Point(12, 1)
        Me.Label81.Name = "Label81"
        Me.Label81.Size = New System.Drawing.Size(73, 9)
        Me.Label81.TabIndex = 536
        Me.Label81.Text = "BAUD: 250k, N, 8, 1"
        '
        'Label232
        '
        Me.Label232.AutoSize = True
        Me.Label232.Location = New System.Drawing.Point(138, 38)
        Me.Label232.Name = "Label232"
        Me.Label232.Size = New System.Drawing.Size(20, 13)
        Me.Label232.TabIndex = 799
        Me.Label232.Text = "Rx"
        '
        'Label233
        '
        Me.Label233.AutoSize = True
        Me.Label233.Location = New System.Drawing.Point(138, 15)
        Me.Label233.Name = "Label233"
        Me.Label233.Size = New System.Drawing.Size(19, 13)
        Me.Label233.TabIndex = 798
        Me.Label233.Text = "Tx"
        '
        'Label154
        '
        Me.Label154.AutoSize = True
        Me.Label154.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label154.Location = New System.Drawing.Point(754, 15)
        Me.Label154.Name = "Label154"
        Me.Label154.Size = New System.Drawing.Size(32, 12)
        Me.Label154.TabIndex = 699
        Me.Label154.Text = "3458A"
        '
        'Label153
        '
        Me.Label153.AutoSize = True
        Me.Label153.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label153.Location = New System.Drawing.Point(911, 15)
        Me.Label153.Name = "Label153"
        Me.Label153.Size = New System.Drawing.Size(53, 12)
        Me.Label153.TabIndex = 698
        Me.Label153.Text = "PDVS2mini"
        '
        'Label152
        '
        Me.Label152.AutoSize = True
        Me.Label152.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label152.Location = New System.Drawing.Point(500, 15)
        Me.Label152.Name = "Label152"
        Me.Label152.Size = New System.Drawing.Size(53, 12)
        Me.Label152.TabIndex = 697
        Me.Label152.Text = "PDVS2mini"
        '
        'Label151
        '
        Me.Label151.AutoSize = True
        Me.Label151.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label151.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label151.Location = New System.Drawing.Point(757, 28)
        Me.Label151.Name = "Label151"
        Me.Label151.Size = New System.Drawing.Size(44, 20)
        Me.Label151.TabIndex = 696
        Me.Label151.Text = "Δ uV"
        Me.Label151.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelDeltaV
        '
        Me.LabelDeltaV.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.LabelDeltaV.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDeltaV.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LabelDeltaV.Location = New System.Drawing.Point(800, 7)
        Me.LabelDeltaV.Name = "LabelDeltaV"
        Me.LabelDeltaV.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LabelDeltaV.Size = New System.Drawing.Size(105, 48)
        Me.LabelDeltaV.TabIndex = 695
        Me.LabelDeltaV.Text = "####"
        Me.LabelDeltaV.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'volts10
        '
        Me.volts10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts10.Location = New System.Drawing.Point(629, 309)
        Me.volts10.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts10.MaxLength = 8
        Me.volts10.Name = "volts10"
        Me.volts10.Size = New System.Drawing.Size(76, 22)
        Me.volts10.TabIndex = 680
        '
        'volts9
        '
        Me.volts9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts9.Location = New System.Drawing.Point(629, 285)
        Me.volts9.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts9.MaxLength = 8
        Me.volts9.Name = "volts9"
        Me.volts9.Size = New System.Drawing.Size(76, 22)
        Me.volts9.TabIndex = 679
        '
        'volts8
        '
        Me.volts8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts8.Location = New System.Drawing.Point(629, 261)
        Me.volts8.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts8.MaxLength = 8
        Me.volts8.Name = "volts8"
        Me.volts8.Size = New System.Drawing.Size(76, 22)
        Me.volts8.TabIndex = 678
        '
        'volts7
        '
        Me.volts7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts7.Location = New System.Drawing.Point(629, 237)
        Me.volts7.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts7.MaxLength = 8
        Me.volts7.Name = "volts7"
        Me.volts7.Size = New System.Drawing.Size(76, 22)
        Me.volts7.TabIndex = 677
        '
        'volts6
        '
        Me.volts6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts6.Location = New System.Drawing.Point(629, 213)
        Me.volts6.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts6.MaxLength = 8
        Me.volts6.Name = "volts6"
        Me.volts6.Size = New System.Drawing.Size(76, 22)
        Me.volts6.TabIndex = 676
        '
        'volts5
        '
        Me.volts5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts5.Location = New System.Drawing.Point(629, 189)
        Me.volts5.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts5.MaxLength = 8
        Me.volts5.Name = "volts5"
        Me.volts5.Size = New System.Drawing.Size(76, 22)
        Me.volts5.TabIndex = 675
        '
        'volts4
        '
        Me.volts4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts4.Location = New System.Drawing.Point(629, 165)
        Me.volts4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts4.MaxLength = 8
        Me.volts4.Name = "volts4"
        Me.volts4.Size = New System.Drawing.Size(76, 22)
        Me.volts4.TabIndex = 674
        '
        'volts3
        '
        Me.volts3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts3.Location = New System.Drawing.Point(629, 141)
        Me.volts3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts3.MaxLength = 8
        Me.volts3.Name = "volts3"
        Me.volts3.Size = New System.Drawing.Size(76, 22)
        Me.volts3.TabIndex = 673
        '
        'volts2
        '
        Me.volts2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts2.Location = New System.Drawing.Point(629, 117)
        Me.volts2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts2.MaxLength = 8
        Me.volts2.Name = "volts2"
        Me.volts2.Size = New System.Drawing.Size(76, 22)
        Me.volts2.TabIndex = 672
        '
        'volts1
        '
        Me.volts1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts1.Location = New System.Drawing.Point(629, 93)
        Me.volts1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts1.MaxLength = 8
        Me.volts1.Name = "volts1"
        Me.volts1.Size = New System.Drawing.Size(76, 22)
        Me.volts1.TabIndex = 671
        '
        'volts0
        '
        Me.volts0.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts0.Location = New System.Drawing.Point(629, 69)
        Me.volts0.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts0.MaxLength = 8
        Me.volts0.Name = "volts0"
        Me.volts0.Size = New System.Drawing.Size(76, 22)
        Me.volts0.TabIndex = 670
        '
        'Label93
        '
        Me.Label93.AutoSize = True
        Me.Label93.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label93.Location = New System.Drawing.Point(489, 312)
        Me.Label93.Name = "Label93"
        Me.Label93.Size = New System.Drawing.Size(68, 16)
        Me.Label93.TabIndex = 667
        Me.Label93.Text = "10.00000V"
        Me.Label93.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabeldacSpan9Cal
        '
        Me.LabeldacSpan9Cal.AutoSize = True
        Me.LabeldacSpan9Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan9Cal.Location = New System.Drawing.Point(796, 312)
        Me.LabeldacSpan9Cal.Name = "LabeldacSpan9Cal"
        Me.LabeldacSpan9Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan9Cal.TabIndex = 665
        Me.LabeldacSpan9Cal.Text = "#######"
        '
        'DacSpan9
        '
        Me.DacSpan9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan9.Location = New System.Drawing.Point(564, 309)
        Me.DacSpan9.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan9.MaxLength = 8
        Me.DacSpan9.Name = "DacSpan9"
        Me.DacSpan9.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan9.TabIndex = 664
        '
        'Label89
        '
        Me.Label89.AutoSize = True
        Me.Label89.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label89.Location = New System.Drawing.Point(496, 288)
        Me.Label89.Name = "Label89"
        Me.Label89.Size = New System.Drawing.Size(61, 16)
        Me.Label89.TabIndex = 661
        Me.Label89.Text = "9.00000V"
        Me.Label89.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabeldacSpan8Cal
        '
        Me.LabeldacSpan8Cal.AutoSize = True
        Me.LabeldacSpan8Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan8Cal.Location = New System.Drawing.Point(796, 288)
        Me.LabeldacSpan8Cal.Name = "LabeldacSpan8Cal"
        Me.LabeldacSpan8Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan8Cal.TabIndex = 659
        Me.LabeldacSpan8Cal.Text = "#######"
        '
        'DacSpan8
        '
        Me.DacSpan8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan8.Location = New System.Drawing.Point(564, 285)
        Me.DacSpan8.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan8.MaxLength = 8
        Me.DacSpan8.Name = "DacSpan8"
        Me.DacSpan8.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan8.TabIndex = 658
        '
        'Label85
        '
        Me.Label85.AutoSize = True
        Me.Label85.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label85.Location = New System.Drawing.Point(496, 264)
        Me.Label85.Name = "Label85"
        Me.Label85.Size = New System.Drawing.Size(61, 16)
        Me.Label85.TabIndex = 655
        Me.Label85.Text = "8.00000V"
        Me.Label85.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabeldacSpan7Cal
        '
        Me.LabeldacSpan7Cal.AutoSize = True
        Me.LabeldacSpan7Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan7Cal.Location = New System.Drawing.Point(796, 264)
        Me.LabeldacSpan7Cal.Name = "LabeldacSpan7Cal"
        Me.LabeldacSpan7Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan7Cal.TabIndex = 653
        Me.LabeldacSpan7Cal.Text = "#######"
        '
        'DacSpan7
        '
        Me.DacSpan7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan7.Location = New System.Drawing.Point(564, 261)
        Me.DacSpan7.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan7.MaxLength = 8
        Me.DacSpan7.Name = "DacSpan7"
        Me.DacSpan7.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan7.TabIndex = 652
        '
        'Label86
        '
        Me.Label86.AutoSize = True
        Me.Label86.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label86.Location = New System.Drawing.Point(496, 240)
        Me.Label86.Name = "Label86"
        Me.Label86.Size = New System.Drawing.Size(61, 16)
        Me.Label86.TabIndex = 649
        Me.Label86.Text = "7.00000V"
        Me.Label86.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabeldacSpan6Cal
        '
        Me.LabeldacSpan6Cal.AutoSize = True
        Me.LabeldacSpan6Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan6Cal.Location = New System.Drawing.Point(796, 240)
        Me.LabeldacSpan6Cal.Name = "LabeldacSpan6Cal"
        Me.LabeldacSpan6Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan6Cal.TabIndex = 647
        Me.LabeldacSpan6Cal.Text = "#######"
        '
        'DacSpan6
        '
        Me.DacSpan6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan6.Location = New System.Drawing.Point(564, 237)
        Me.DacSpan6.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan6.MaxLength = 8
        Me.DacSpan6.Name = "DacSpan6"
        Me.DacSpan6.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan6.TabIndex = 646
        '
        'Label90
        '
        Me.Label90.AutoSize = True
        Me.Label90.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label90.Location = New System.Drawing.Point(496, 216)
        Me.Label90.Name = "Label90"
        Me.Label90.Size = New System.Drawing.Size(61, 16)
        Me.Label90.TabIndex = 643
        Me.Label90.Text = "6.00000V"
        Me.Label90.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabeldacSpan5Cal
        '
        Me.LabeldacSpan5Cal.AutoSize = True
        Me.LabeldacSpan5Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan5Cal.Location = New System.Drawing.Point(796, 216)
        Me.LabeldacSpan5Cal.Name = "LabeldacSpan5Cal"
        Me.LabeldacSpan5Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan5Cal.TabIndex = 641
        Me.LabeldacSpan5Cal.Text = "#######"
        '
        'DacSpan5
        '
        Me.DacSpan5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan5.Location = New System.Drawing.Point(564, 213)
        Me.DacSpan5.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan5.MaxLength = 8
        Me.DacSpan5.Name = "DacSpan5"
        Me.DacSpan5.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan5.TabIndex = 640
        '
        'Label94
        '
        Me.Label94.AutoSize = True
        Me.Label94.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label94.Location = New System.Drawing.Point(496, 192)
        Me.Label94.Name = "Label94"
        Me.Label94.Size = New System.Drawing.Size(61, 16)
        Me.Label94.TabIndex = 637
        Me.Label94.Text = "5.00000V"
        Me.Label94.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabeldacSpan4Cal
        '
        Me.LabeldacSpan4Cal.AutoSize = True
        Me.LabeldacSpan4Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan4Cal.Location = New System.Drawing.Point(796, 192)
        Me.LabeldacSpan4Cal.Name = "LabeldacSpan4Cal"
        Me.LabeldacSpan4Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan4Cal.TabIndex = 635
        Me.LabeldacSpan4Cal.Text = "#######"
        '
        'DacSpan4
        '
        Me.DacSpan4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan4.Location = New System.Drawing.Point(564, 189)
        Me.DacSpan4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan4.MaxLength = 8
        Me.DacSpan4.Name = "DacSpan4"
        Me.DacSpan4.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan4.TabIndex = 634
        '
        'Label96
        '
        Me.Label96.AutoSize = True
        Me.Label96.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label96.Location = New System.Drawing.Point(496, 168)
        Me.Label96.Name = "Label96"
        Me.Label96.Size = New System.Drawing.Size(61, 16)
        Me.Label96.TabIndex = 631
        Me.Label96.Text = "4.00000V"
        Me.Label96.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabeldacSpan3Cal
        '
        Me.LabeldacSpan3Cal.AutoSize = True
        Me.LabeldacSpan3Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan3Cal.Location = New System.Drawing.Point(796, 168)
        Me.LabeldacSpan3Cal.Name = "LabeldacSpan3Cal"
        Me.LabeldacSpan3Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan3Cal.TabIndex = 629
        Me.LabeldacSpan3Cal.Text = "#######"
        '
        'DacSpan3
        '
        Me.DacSpan3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan3.Location = New System.Drawing.Point(564, 165)
        Me.DacSpan3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan3.MaxLength = 8
        Me.DacSpan3.Name = "DacSpan3"
        Me.DacSpan3.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan3.TabIndex = 628
        '
        'Label98
        '
        Me.Label98.AutoSize = True
        Me.Label98.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label98.Location = New System.Drawing.Point(496, 144)
        Me.Label98.Name = "Label98"
        Me.Label98.Size = New System.Drawing.Size(61, 16)
        Me.Label98.TabIndex = 625
        Me.Label98.Text = "3.00000V"
        Me.Label98.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabeldacSpan2Cal
        '
        Me.LabeldacSpan2Cal.AutoSize = True
        Me.LabeldacSpan2Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan2Cal.Location = New System.Drawing.Point(796, 144)
        Me.LabeldacSpan2Cal.Name = "LabeldacSpan2Cal"
        Me.LabeldacSpan2Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan2Cal.TabIndex = 623
        Me.LabeldacSpan2Cal.Text = "#######"
        '
        'DacSpan2
        '
        Me.DacSpan2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan2.Location = New System.Drawing.Point(564, 141)
        Me.DacSpan2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan2.MaxLength = 8
        Me.DacSpan2.Name = "DacSpan2"
        Me.DacSpan2.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan2.TabIndex = 622
        '
        'Label100
        '
        Me.Label100.AutoSize = True
        Me.Label100.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label100.Location = New System.Drawing.Point(496, 120)
        Me.Label100.Name = "Label100"
        Me.Label100.Size = New System.Drawing.Size(61, 16)
        Me.Label100.TabIndex = 619
        Me.Label100.Text = "2.00000V"
        Me.Label100.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabeldacSpan1Cal
        '
        Me.LabeldacSpan1Cal.AutoSize = True
        Me.LabeldacSpan1Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan1Cal.Location = New System.Drawing.Point(796, 120)
        Me.LabeldacSpan1Cal.Name = "LabeldacSpan1Cal"
        Me.LabeldacSpan1Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan1Cal.TabIndex = 617
        Me.LabeldacSpan1Cal.Text = "#######"
        '
        'DacSpan1
        '
        Me.DacSpan1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan1.Location = New System.Drawing.Point(564, 117)
        Me.DacSpan1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan1.MaxLength = 8
        Me.DacSpan1.Name = "DacSpan1"
        Me.DacSpan1.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan1.TabIndex = 616
        '
        'LabelBatteryMonICMult
        '
        Me.LabelBatteryMonICMult.AutoSize = True
        Me.LabelBatteryMonICMult.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelBatteryMonICMult.Location = New System.Drawing.Point(796, 408)
        Me.LabelBatteryMonICMult.Name = "LabelBatteryMonICMult"
        Me.LabelBatteryMonICMult.Size = New System.Drawing.Size(56, 16)
        Me.LabelBatteryMonICMult.TabIndex = 604
        Me.LabelBatteryMonICMult.Text = "#######"
        '
        'Label106
        '
        Me.Label106.AutoSize = True
        Me.Label106.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label106.Location = New System.Drawing.Point(483, 408)
        Me.Label106.Name = "Label106"
        Me.Label106.Size = New System.Drawing.Size(74, 16)
        Me.Label106.TabIndex = 603
        Me.Label106.Text = "Charge mA"
        Me.Label106.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ChargeI
        '
        Me.ChargeI.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChargeI.Location = New System.Drawing.Point(564, 405)
        Me.ChargeI.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ChargeI.MaxLength = 8
        Me.ChargeI.Name = "ChargeI"
        Me.ChargeI.Size = New System.Drawing.Size(61, 22)
        Me.ChargeI.TabIndex = 601
        '
        'Label107
        '
        Me.Label107.AutoSize = True
        Me.Label107.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label107.Location = New System.Drawing.Point(481, 384)
        Me.Label107.Name = "Label107"
        Me.Label107.Size = New System.Drawing.Size(76, 16)
        Me.Label107.TabIndex = 600
        Me.Label107.Text = "Battery Vdc"
        Me.Label107.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label108
        '
        Me.Label108.AutoSize = True
        Me.Label108.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label108.Location = New System.Drawing.Point(471, 360)
        Me.Label108.Name = "Label108"
        Me.Label108.Size = New System.Drawing.Size(86, 16)
        Me.Label108.TabIndex = 599
        Me.Label108.Text = "Out Mon. Vdc"
        Me.Label108.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label109
        '
        Me.Label109.AutoSize = True
        Me.Label109.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label109.Location = New System.Drawing.Point(496, 96)
        Me.Label109.Name = "Label109"
        Me.Label109.Size = New System.Drawing.Size(61, 16)
        Me.Label109.TabIndex = 598
        Me.Label109.Text = "1.00000V"
        Me.Label109.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label110
        '
        Me.Label110.AutoSize = True
        Me.Label110.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label110.Location = New System.Drawing.Point(496, 72)
        Me.Label110.Name = "Label110"
        Me.Label110.Size = New System.Drawing.Size(61, 16)
        Me.Label110.TabIndex = 597
        Me.Label110.Text = "0.00000V"
        Me.Label110.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelBatteryVFeedMult
        '
        Me.LabelBatteryVFeedMult.AutoSize = True
        Me.LabelBatteryVFeedMult.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelBatteryVFeedMult.Location = New System.Drawing.Point(796, 384)
        Me.LabelBatteryVFeedMult.Name = "LabelBatteryVFeedMult"
        Me.LabelBatteryVFeedMult.Size = New System.Drawing.Size(56, 16)
        Me.LabelBatteryVFeedMult.TabIndex = 591
        Me.LabelBatteryVFeedMult.Text = "#######"
        '
        'LabelOutputVFeedMult
        '
        Me.LabelOutputVFeedMult.AutoSize = True
        Me.LabelOutputVFeedMult.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelOutputVFeedMult.Location = New System.Drawing.Point(796, 360)
        Me.LabelOutputVFeedMult.Name = "LabelOutputVFeedMult"
        Me.LabelOutputVFeedMult.Size = New System.Drawing.Size(56, 16)
        Me.LabelOutputVFeedMult.TabIndex = 590
        Me.LabelOutputVFeedMult.Text = "#######"
        '
        'LabeldacSpan0Cal
        '
        Me.LabeldacSpan0Cal.AutoSize = True
        Me.LabeldacSpan0Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan0Cal.Location = New System.Drawing.Point(796, 96)
        Me.LabeldacSpan0Cal.Name = "LabeldacSpan0Cal"
        Me.LabeldacSpan0Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan0Cal.TabIndex = 589
        Me.LabeldacSpan0Cal.Text = "#######"
        '
        'LabeldacZero0Cal
        '
        Me.LabeldacZero0Cal.AutoSize = True
        Me.LabeldacZero0Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacZero0Cal.Location = New System.Drawing.Point(796, 72)
        Me.LabeldacZero0Cal.Name = "LabeldacZero0Cal"
        Me.LabeldacZero0Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacZero0Cal.TabIndex = 588
        Me.LabeldacZero0Cal.Text = "#######"
        '
        'BattVdc
        '
        Me.BattVdc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BattVdc.Location = New System.Drawing.Point(564, 381)
        Me.BattVdc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.BattVdc.MaxLength = 8
        Me.BattVdc.Name = "BattVdc"
        Me.BattVdc.Size = New System.Drawing.Size(61, 22)
        Me.BattVdc.TabIndex = 587
        '
        'OutVdc
        '
        Me.OutVdc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OutVdc.Location = New System.Drawing.Point(564, 357)
        Me.OutVdc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.OutVdc.MaxLength = 8
        Me.OutVdc.Name = "OutVdc"
        Me.OutVdc.Size = New System.Drawing.Size(61, 22)
        Me.OutVdc.TabIndex = 586
        '
        'DacSpan0
        '
        Me.DacSpan0.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan0.Location = New System.Drawing.Point(564, 93)
        Me.DacSpan0.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan0.MaxLength = 8
        Me.DacSpan0.Name = "DacSpan0"
        Me.DacSpan0.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan0.TabIndex = 585
        '
        'DacZero0
        '
        Me.DacZero0.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacZero0.Location = New System.Drawing.Point(564, 69)
        Me.DacZero0.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacZero0.MaxLength = 8
        Me.DacZero0.Name = "DacZero0"
        Me.DacZero0.Size = New System.Drawing.Size(61, 22)
        Me.DacZero0.TabIndex = 584
        '
        'Label57
        '
        Me.Label57.AutoSize = True
        Me.Label57.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label57.Location = New System.Drawing.Point(161, 15)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(48, 12)
        Me.Label57.TabIndex = 582
        Me.Label57.Text = "DEVICE 1"
        '
        'TextBoxdegC
        '
        Me.TextBoxdegC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxdegC.Location = New System.Drawing.Point(13, 515)
        Me.TextBoxdegC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxdegC.MaxLength = 8
        Me.TextBoxdegC.Name = "TextBoxdegC"
        Me.TextBoxdegC.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxdegC.TabIndex = 579
        '
        'Label97
        '
        Me.Label97.AutoSize = True
        Me.Label97.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label97.Location = New System.Drawing.Point(107, 518)
        Me.Label97.Name = "Label97"
        Me.Label97.Size = New System.Drawing.Size(42, 16)
        Me.Label97.TabIndex = 580
        Me.Label97.Text = "DegC"
        '
        'TextBoxSer
        '
        Me.TextBoxSer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxSer.Location = New System.Drawing.Point(176, 411)
        Me.TextBoxSer.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxSer.MaxLength = 8
        Me.TextBoxSer.Name = "TextBoxSer"
        Me.TextBoxSer.Size = New System.Drawing.Size(59, 22)
        Me.TextBoxSer.TabIndex = 577
        '
        'Label95
        '
        Me.Label95.AutoSize = True
        Me.Label95.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label95.Location = New System.Drawing.Point(237, 414)
        Me.Label95.Name = "Label95"
        Me.Label95.Size = New System.Drawing.Size(66, 16)
        Me.Label95.TabIndex = 578
        Me.Label95.Text = "Serial No."
        '
        'TextBoxSOAK
        '
        Me.TextBoxSOAK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxSOAK.Location = New System.Drawing.Point(175, 463)
        Me.TextBoxSOAK.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxSOAK.MaxLength = 8
        Me.TextBoxSOAK.Name = "TextBoxSOAK"
        Me.TextBoxSOAK.Size = New System.Drawing.Size(60, 22)
        Me.TextBoxSOAK.TabIndex = 570
        Me.TextBoxSOAK.Text = "168hrs"
        '
        'TextBoxSERIAL
        '
        Me.TextBoxSERIAL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxSERIAL.Location = New System.Drawing.Point(176, 437)
        Me.TextBoxSERIAL.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxSERIAL.MaxLength = 8
        Me.TextBoxSERIAL.Name = "TextBoxSERIAL"
        Me.TextBoxSERIAL.Size = New System.Drawing.Size(59, 22)
        Me.TextBoxSERIAL.TabIndex = 568
        Me.TextBoxSERIAL.Text = "OK"
        '
        'TextBoxFULLMA
        '
        Me.TextBoxFULLMA.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxFULLMA.Location = New System.Drawing.Point(13, 489)
        Me.TextBoxFULLMA.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxFULLMA.MaxLength = 8
        Me.TextBoxFULLMA.Name = "TextBoxFULLMA"
        Me.TextBoxFULLMA.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxFULLMA.TabIndex = 566
        Me.TextBoxFULLMA.Text = "70mA"
        '
        'TextBoxOLMA
        '
        Me.TextBoxOLMA.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxOLMA.Location = New System.Drawing.Point(13, 463)
        Me.TextBoxOLMA.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxOLMA.MaxLength = 8
        Me.TextBoxOLMA.Name = "TextBoxOLMA"
        Me.TextBoxOLMA.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxOLMA.TabIndex = 564
        Me.TextBoxOLMA.Text = "400mA"
        '
        'TextBoxCENABLE
        '
        Me.TextBoxCENABLE.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxCENABLE.Location = New System.Drawing.Point(13, 437)
        Me.TextBoxCENABLE.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxCENABLE.MaxLength = 8
        Me.TextBoxCENABLE.Name = "TextBoxCENABLE"
        Me.TextBoxCENABLE.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxCENABLE.TabIndex = 562
        Me.TextBoxCENABLE.Text = "YES"
        '
        'TextBoxLOWSHUT
        '
        Me.TextBoxLOWSHUT.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxLOWSHUT.Location = New System.Drawing.Point(13, 411)
        Me.TextBoxLOWSHUT.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxLOWSHUT.MaxLength = 20
        Me.TextBoxLOWSHUT.Name = "TextBoxLOWSHUT"
        Me.TextBoxLOWSHUT.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxLOWSHUT.TabIndex = 560
        Me.TextBoxLOWSHUT.Text = "13.0/12.5Vdc"
        '
        'TextBoxDC
        '
        Me.TextBoxDC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxDC.Location = New System.Drawing.Point(175, 489)
        Me.TextBoxDC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxDC.MaxLength = 8
        Me.TextBoxDC.Name = "TextBoxDC"
        Me.TextBoxDC.Size = New System.Drawing.Size(60, 22)
        Me.TextBoxDC.TabIndex = 572
        Me.TextBoxDC.Text = "OK"
        '
        'TextBoxCD
        '
        Me.TextBoxCD.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxCD.Location = New System.Drawing.Point(175, 515)
        Me.TextBoxCD.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxCD.MaxLength = 8
        Me.TextBoxCD.Name = "TextBoxCD"
        Me.TextBoxCD.Size = New System.Drawing.Size(60, 22)
        Me.TextBoxCD.TabIndex = 574
        Me.TextBoxCD.Text = "OK"
        '
        'Label92
        '
        Me.Label92.AutoSize = True
        Me.Label92.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label92.Location = New System.Drawing.Point(237, 518)
        Me.Label92.Name = "Label92"
        Me.Label92.Size = New System.Drawing.Size(58, 16)
        Me.Label92.TabIndex = 575
        Me.Label92.Text = "Chg/Dis."
        '
        'Label91
        '
        Me.Label91.AutoSize = True
        Me.Label91.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label91.Location = New System.Drawing.Point(237, 492)
        Me.Label91.Name = "Label91"
        Me.Label91.Size = New System.Drawing.Size(50, 16)
        Me.Label91.TabIndex = 573
        Me.Label91.Text = "DC Inp."
        '
        'Label88
        '
        Me.Label88.AutoSize = True
        Me.Label88.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label88.Location = New System.Drawing.Point(237, 466)
        Me.Label88.Name = "Label88"
        Me.Label88.Size = New System.Drawing.Size(39, 16)
        Me.Label88.TabIndex = 571
        Me.Label88.Text = "Soak"
        '
        'Label87
        '
        Me.Label87.AutoSize = True
        Me.Label87.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label87.Location = New System.Drawing.Point(237, 440)
        Me.Label87.Name = "Label87"
        Me.Label87.Size = New System.Drawing.Size(42, 16)
        Me.Label87.TabIndex = 569
        Me.Label87.Text = "Serial"
        '
        'Label84
        '
        Me.Label84.AutoSize = True
        Me.Label84.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label84.Location = New System.Drawing.Point(106, 492)
        Me.Label84.Name = "Label84"
        Me.Label84.Size = New System.Drawing.Size(51, 16)
        Me.Label84.TabIndex = 567
        Me.Label84.Text = "Full mA"
        '
        'Label82
        '
        Me.Label82.AutoSize = True
        Me.Label82.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label82.Location = New System.Drawing.Point(107, 466)
        Me.Label82.Name = "Label82"
        Me.Label82.Size = New System.Drawing.Size(51, 16)
        Me.Label82.TabIndex = 565
        Me.Label82.Text = "O/L mA"
        '
        'Label78
        '
        Me.Label78.AutoSize = True
        Me.Label78.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label78.Location = New System.Drawing.Point(106, 440)
        Me.Label78.Name = "Label78"
        Me.Label78.Size = New System.Drawing.Size(62, 16)
        Me.Label78.TabIndex = 563
        Me.Label78.Text = "C.Enable"
        '
        'Label77
        '
        Me.Label77.AutoSize = True
        Me.Label77.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label77.Location = New System.Drawing.Point(107, 414)
        Me.Label77.Name = "Label77"
        Me.Label77.Size = New System.Drawing.Size(61, 16)
        Me.Label77.TabIndex = 561
        Me.Label77.Text = "Low/Shut"
        '
        'Label114
        '
        Me.Label114.AutoSize = True
        Me.Label114.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label114.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label114.Location = New System.Drawing.Point(446, 24)
        Me.Label114.Name = "Label114"
        Me.Label114.Size = New System.Drawing.Size(54, 29)
        Me.Label114.TabIndex = 541
        Me.Label114.Text = "Vdc"
        '
        'Label113
        '
        Me.Label113.AutoSize = True
        Me.Label113.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label113.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label113.Location = New System.Drawing.Point(163, 28)
        Me.Label113.Name = "Label113"
        Me.Label113.Size = New System.Drawing.Size(47, 20)
        Me.Label113.TabIndex = 540
        Me.Label113.Text = "DMM"
        Me.Label113.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label112
        '
        Me.Label112.AutoSize = True
        Me.Label112.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label112.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label112.Location = New System.Drawing.Point(509, 28)
        Me.Label112.Name = "Label112"
        Me.Label112.Size = New System.Drawing.Size(40, 20)
        Me.Label112.TabIndex = 539
        Me.Label112.Text = "SET"
        Me.Label112.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtr1aBIG
        '
        Me.txtr1aBIG.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.txtr1aBIG.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtr1aBIG.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.txtr1aBIG.Location = New System.Drawing.Point(209, 7)
        Me.txtr1aBIG.Name = "txtr1aBIG"
        Me.txtr1aBIG.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtr1aBIG.Size = New System.Drawing.Size(267, 48)
        Me.txtr1aBIG.TabIndex = 538
        Me.txtr1aBIG.Text = "###########"
        Me.txtr1aBIG.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'comPort_ComboBox
        '
        Me.comPort_ComboBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comPort_ComboBox.FormattingEnabled = True
        Me.comPort_ComboBox.Location = New System.Drawing.Point(9, 11)
        Me.comPort_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.comPort_ComboBox.Name = "comPort_ComboBox"
        Me.comPort_ComboBox.Size = New System.Drawing.Size(80, 21)
        Me.comPort_ComboBox.TabIndex = 534
        '
        'connect_BTN
        '
        Me.connect_BTN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.connect_BTN.Location = New System.Drawing.Point(8, 33)
        Me.connect_BTN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.connect_BTN.Name = "connect_BTN"
        Me.connect_BTN.Size = New System.Drawing.Size(82, 22)
        Me.connect_BTN.TabIndex = 535
        Me.connect_BTN.Text = "Connect"
        Me.connect_BTN.UseVisualStyleBackColor = True
        '
        'ButtonKeyVoltage
        '
        Me.ButtonKeyVoltage.Location = New System.Drawing.Point(966, 10)
        Me.ButtonKeyVoltage.Name = "ButtonKeyVoltage"
        Me.ButtonKeyVoltage.Size = New System.Drawing.Size(76, 22)
        Me.ButtonKeyVoltage.TabIndex = 532
        Me.ButtonKeyVoltage.Text = "SET"
        Me.ButtonKeyVoltage.UseVisualStyleBackColor = True
        '
        'KeyVoltage
        '
        Me.KeyVoltage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyVoltage.Location = New System.Drawing.Point(967, 33)
        Me.KeyVoltage.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.KeyVoltage.MaxLength = 8
        Me.KeyVoltage.Name = "KeyVoltage"
        Me.KeyVoltage.Size = New System.Drawing.Size(74, 20)
        Me.KeyVoltage.TabIndex = 528
        '
        'Label80
        '
        Me.Label80.AutoSize = True
        Me.Label80.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label80.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label80.Location = New System.Drawing.Point(698, 24)
        Me.Label80.Name = "Label80"
        Me.Label80.Size = New System.Drawing.Size(54, 29)
        Me.Label80.TabIndex = 531
        Me.Label80.Text = "Vdc"
        '
        'LabelKeyVoltage
        '
        Me.LabelKeyVoltage.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.LabelKeyVoltage.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelKeyVoltage.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LabelKeyVoltage.Location = New System.Drawing.Point(548, 7)
        Me.LabelKeyVoltage.Name = "LabelKeyVoltage"
        Me.LabelKeyVoltage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LabelKeyVoltage.Size = New System.Drawing.Size(182, 48)
        Me.LabelKeyVoltage.TabIndex = 530
        Me.LabelKeyVoltage.Text = "#######"
        Me.LabelKeyVoltage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label174)
        Me.GroupBox2.Controls.Add(Me.Default11)
        Me.GroupBox2.Controls.Add(Me.Label172)
        Me.GroupBox2.Controls.Add(Me.Default10)
        Me.GroupBox2.Controls.Add(Me.Default9)
        Me.GroupBox2.Controls.Add(Me.Default8)
        Me.GroupBox2.Controls.Add(Me.Default7)
        Me.GroupBox2.Controls.Add(Me.Default6)
        Me.GroupBox2.Controls.Add(Me.Default5)
        Me.GroupBox2.Controls.Add(Me.Default4)
        Me.GroupBox2.Controls.Add(Me.Default3)
        Me.GroupBox2.Controls.Add(Me.Default2)
        Me.GroupBox2.Controls.Add(Me.Default1)
        Me.GroupBox2.Controls.Add(Me.Default0)
        Me.GroupBox2.Controls.Add(Me.Label168)
        Me.GroupBox2.Controls.Add(Me.Label158)
        Me.GroupBox2.Controls.Add(Me.Label157)
        Me.GroupBox2.Controls.Add(Me.Label156)
        Me.GroupBox2.Controls.Add(Me.LabelTemperature2)
        Me.GroupBox2.Controls.Add(Me.Label155)
        Me.GroupBox2.Controls.Add(Me.Label229)
        Me.GroupBox2.Controls.Add(Me.LabelTemperature3)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan10Delta)
        Me.GroupBox2.Controls.Add(Me.Label149)
        Me.GroupBox2.Controls.Add(Me.Label150)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan10Up)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan10)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan10down)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan10Cal)
        Me.GroupBox2.Controls.Add(Me.volts11)
        Me.GroupBox2.Controls.Add(Me.DacSpan10)
        Me.GroupBox2.Controls.Add(Me.WryTech)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan9Up)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan9)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan9down)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan8)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan8Up)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan7)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan8down)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan6)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan7Up)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan5)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan7down)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan4)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan6Up)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan6down)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan3)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan5Up)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan2)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan5down)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan1)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan4Up)
        Me.GroupBox2.Controls.Add(Me.ButtonOutVdc)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan4down)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan0)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan3Up)
        Me.GroupBox2.Controls.Add(Me.ButtonDacZero0)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan3down)
        Me.GroupBox2.Controls.Add(Me.ButtonBattVdc)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan2Up)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan2down)
        Me.GroupBox2.Controls.Add(Me.Label162)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan1Up)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan1down)
        Me.GroupBox2.Controls.Add(Me.ButtonChargeIUp)
        Me.GroupBox2.Controls.Add(Me.ButtonBattVdcUp)
        Me.GroupBox2.Controls.Add(Me.Label225)
        Me.GroupBox2.Controls.Add(Me.ButtonOutVdcUp)
        Me.GroupBox2.Controls.Add(Me.Label224)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan0Up)
        Me.GroupBox2.Controls.Add(Me.PDVS2miniSave)
        Me.GroupBox2.Controls.Add(Me.ButtonDacZero0Up)
        Me.GroupBox2.Controls.Add(Me.TextBoxUser)
        Me.GroupBox2.Controls.Add(Me.ButtonChargeIdown)
        Me.GroupBox2.Controls.Add(Me.Label222)
        Me.GroupBox2.Controls.Add(Me.ButtonBattVdcdown)
        Me.GroupBox2.Controls.Add(Me.TextBox3458Asn)
        Me.GroupBox2.Controls.Add(Me.ButtonOutVdcdown)
        Me.GroupBox2.Controls.Add(Me.Label73)
        Me.GroupBox2.Controls.Add(Me.ButtonDacSpan0down)
        Me.GroupBox2.Controls.Add(Me.Label290)
        Me.GroupBox2.Controls.Add(Me.ButtonDacZero0down)
        Me.GroupBox2.Controls.Add(Me.Label289)
        Me.GroupBox2.Controls.Add(Me.Label288)
        Me.GroupBox2.Controls.Add(Me.Label287)
        Me.GroupBox2.Controls.Add(Me.Label286)
        Me.GroupBox2.Controls.Add(Me.Label285)
        Me.GroupBox2.Controls.Add(Me.Label284)
        Me.GroupBox2.Controls.Add(Me.Label283)
        Me.GroupBox2.Controls.Add(Me.Label282)
        Me.GroupBox2.Controls.Add(Me.Label281)
        Me.GroupBox2.Controls.Add(Me.Label280)
        Me.GroupBox2.Controls.Add(Me.ButtonLoadDefs)
        Me.GroupBox2.Controls.Add(Me.ButtonSaveDefs)
        Me.GroupBox2.Controls.Add(Me.Label279)
        Me.GroupBox2.Controls.Add(Me.Label278)
        Me.GroupBox2.Controls.Add(Me.Label277)
        Me.GroupBox2.Controls.Add(Me.Label161)
        Me.GroupBox2.Controls.Add(Me.Label160)
        Me.GroupBox2.Controls.Add(Me.Label159)
        Me.GroupBox2.Controls.Add(Me.Label137)
        Me.GroupBox2.Controls.Add(Me.Label136)
        Me.GroupBox2.Controls.Add(Me.Label124)
        Me.GroupBox2.Controls.Add(Me.Label104)
        Me.GroupBox2.Controls.Add(Me.Label55)
        Me.GroupBox2.Controls.Add(Me.DACcalparamVAL)
        Me.GroupBox2.Controls.Add(Me.RadioPDF)
        Me.GroupBox2.Controls.Add(Me.PDVS2counter)
        Me.GroupBox2.Controls.Add(Me.RadioMSWord)
        Me.GroupBox2.Controls.Add(Me.Label254)
        Me.GroupBox2.Controls.Add(Me.ButtonAutomV)
        Me.GroupBox2.Controls.Add(Me.ButtonAutoSET)
        Me.GroupBox2.Controls.Add(Me.ButtonSetRetrievedVars)
        Me.GroupBox2.Controls.Add(Me.CalOnExisting)
        Me.GroupBox2.Controls.Add(Me.Label163)
        Me.GroupBox2.Controls.Add(Me.Label167)
        Me.GroupBox2.Controls.Add(Me.Label166)
        Me.GroupBox2.Controls.Add(Me.Label165)
        Me.GroupBox2.Controls.Add(Me.TextBoxLabelBatteryCharge)
        Me.GroupBox2.Controls.Add(Me.TextBoxLabelBatteryLowInd)
        Me.GroupBox2.Controls.Add(Me.TextBoxLabelMode)
        Me.GroupBox2.Controls.Add(Me.TextBoxLabelBatteryI)
        Me.GroupBox2.Controls.Add(Me.TextBoxLabelBatteryV)
        Me.GroupBox2.Controls.Add(Me.TextBoxLabelOutputVFeedback)
        Me.GroupBox2.Controls.Add(Me.Label118)
        Me.GroupBox2.Controls.Add(Me.ButtonSetXYcalvars)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan9Delta)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan8Delta)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan7Delta)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan6Delta)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan5Delta)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan4Delta)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan3Delta)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan2Delta)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan1Delta)
        Me.GroupBox2.Controls.Add(Me.LabeldacSpan0Delta)
        Me.GroupBox2.Controls.Add(Me.LabeldacZero0Delta)
        Me.GroupBox2.Controls.Add(Me.volts030000)
        Me.GroupBox2.Controls.Add(Me.Button030000)
        Me.GroupBox2.Controls.Add(Me.volts050000)
        Me.GroupBox2.Controls.Add(Me.volts020000)
        Me.GroupBox2.Controls.Add(Me.volts010000)
        Me.GroupBox2.Controls.Add(Me.volts001000)
        Me.GroupBox2.Controls.Add(Me.volts000100)
        Me.GroupBox2.Controls.Add(Me.Button050000)
        Me.GroupBox2.Controls.Add(Me.Button020000)
        Me.GroupBox2.Controls.Add(Me.Button010000)
        Me.GroupBox2.Controls.Add(Me.Button001000)
        Me.GroupBox2.Controls.Add(Me.Button000010)
        Me.GroupBox2.Controls.Add(Me.volts000010)
        Me.GroupBox2.Controls.Add(Me.Button000100)
        Me.GroupBox2.Controls.Add(Me.ButtonSetPrecalvars)
        Me.GroupBox2.Controls.Add(Me.nplc4_BTN)
        Me.GroupBox2.Controls.Add(Me.nplc3_BTN)
        Me.GroupBox2.Controls.Add(Me.nplc2_BTN)
        Me.GroupBox2.Controls.Add(Me.getcal_BTN)
        Me.GroupBox2.Controls.Add(Me.nplc_BTN)
        Me.GroupBox2.Controls.Add(Me.NPLC)
        Me.GroupBox2.Controls.Add(Me.Abort_BTN)
        Me.GroupBox2.Controls.Add(Me.precal_BTN)
        Me.GroupBox2.Controls.Add(Me.SavePDVS2Eprom)
        Me.GroupBox2.Controls.Add(Me.Label83)
        Me.GroupBox2.Controls.Add(Me.Label76)
        Me.GroupBox2.Controls.Add(Me.Label75)
        Me.GroupBox2.Controls.Add(Me.Label74)
        Me.GroupBox2.Controls.Add(Me.LabelCalCount)
        Me.GroupBox2.Controls.Add(Me.ButtonChargeI)
        Me.GroupBox2.Controls.Add(Me.exportBTN)
        Me.GroupBox2.Controls.Add(Me.PDVS2delay)
        Me.GroupBox2.Controls.Add(Me.CalStep)
        Me.GroupBox2.Controls.Add(Me.Label115)
        Me.GroupBox2.Controls.Add(Me.CalAccuracy)
        Me.GroupBox2.Controls.Add(Me.Label119)
        Me.GroupBox2.Controls.Add(Me.CalStepFinal)
        Me.GroupBox2.Controls.Add(Me.Label120)
        Me.GroupBox2.Controls.Add(Me.CalAccuracyFinal)
        Me.GroupBox2.Controls.Add(Me.Label117)
        Me.GroupBox2.Controls.Add(Me.Label121)
        Me.GroupBox2.Enabled = False
        Me.GroupBox2.Location = New System.Drawing.Point(8, 51)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1038, 548)
        Me.GroupBox2.TabIndex = 684
        Me.GroupBox2.TabStop = False
        '
        'Label174
        '
        Me.Label174.AutoSize = True
        Me.Label174.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label174.Location = New System.Drawing.Point(179, 39)
        Me.Label174.Name = "Label174"
        Me.Label174.Size = New System.Drawing.Size(38, 16)
        Me.Label174.TabIndex = 842
        Me.Label174.Text = "Step:"
        Me.Label174.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.Label174.Visible = False
        '
        'Default11
        '
        Me.Default11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default11.Location = New System.Drawing.Point(887, 283)
        Me.Default11.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default11.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default11.MaxLength = 8
        Me.Default11.Name = "Default11"
        Me.Default11.Size = New System.Drawing.Size(61, 22)
        Me.Default11.TabIndex = 833
        Me.Default11.Text = "0"
        '
        'Label172
        '
        Me.Label172.AutoSize = True
        Me.Label172.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label172.Location = New System.Drawing.Point(900, 7)
        Me.Label172.Name = "Label172"
        Me.Label172.Size = New System.Drawing.Size(35, 12)
        Me.Label172.TabIndex = 841
        Me.Label172.Text = "Counts"
        '
        'Default10
        '
        Me.Default10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default10.Location = New System.Drawing.Point(887, 259)
        Me.Default10.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default10.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default10.MaxLength = 8
        Me.Default10.Name = "Default10"
        Me.Default10.Size = New System.Drawing.Size(61, 22)
        Me.Default10.TabIndex = 763
        Me.Default10.Text = "0"
        '
        'Default9
        '
        Me.Default9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default9.Location = New System.Drawing.Point(887, 235)
        Me.Default9.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default9.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default9.MaxLength = 8
        Me.Default9.Name = "Default9"
        Me.Default9.Size = New System.Drawing.Size(61, 22)
        Me.Default9.TabIndex = 762
        Me.Default9.Text = "0"
        '
        'Default8
        '
        Me.Default8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default8.Location = New System.Drawing.Point(887, 211)
        Me.Default8.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default8.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default8.MaxLength = 8
        Me.Default8.Name = "Default8"
        Me.Default8.Size = New System.Drawing.Size(61, 22)
        Me.Default8.TabIndex = 761
        Me.Default8.Text = "0"
        '
        'Default7
        '
        Me.Default7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default7.Location = New System.Drawing.Point(887, 187)
        Me.Default7.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default7.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default7.MaxLength = 8
        Me.Default7.Name = "Default7"
        Me.Default7.Size = New System.Drawing.Size(61, 22)
        Me.Default7.TabIndex = 760
        Me.Default7.Text = "0"
        '
        'Default6
        '
        Me.Default6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default6.Location = New System.Drawing.Point(887, 163)
        Me.Default6.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default6.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default6.MaxLength = 8
        Me.Default6.Name = "Default6"
        Me.Default6.Size = New System.Drawing.Size(61, 22)
        Me.Default6.TabIndex = 759
        Me.Default6.Text = "0"
        '
        'Default5
        '
        Me.Default5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default5.Location = New System.Drawing.Point(887, 139)
        Me.Default5.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default5.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default5.MaxLength = 8
        Me.Default5.Name = "Default5"
        Me.Default5.Size = New System.Drawing.Size(61, 22)
        Me.Default5.TabIndex = 758
        Me.Default5.Text = "0"
        '
        'Default4
        '
        Me.Default4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default4.Location = New System.Drawing.Point(887, 115)
        Me.Default4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default4.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default4.MaxLength = 8
        Me.Default4.Name = "Default4"
        Me.Default4.Size = New System.Drawing.Size(61, 22)
        Me.Default4.TabIndex = 757
        Me.Default4.Text = "0"
        '
        'Default3
        '
        Me.Default3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default3.Location = New System.Drawing.Point(887, 91)
        Me.Default3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default3.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default3.MaxLength = 8
        Me.Default3.Name = "Default3"
        Me.Default3.Size = New System.Drawing.Size(61, 22)
        Me.Default3.TabIndex = 756
        Me.Default3.Text = "0"
        '
        'Default2
        '
        Me.Default2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default2.Location = New System.Drawing.Point(887, 67)
        Me.Default2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default2.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default2.MaxLength = 8
        Me.Default2.Name = "Default2"
        Me.Default2.Size = New System.Drawing.Size(61, 22)
        Me.Default2.TabIndex = 755
        Me.Default2.Text = "0"
        '
        'Default1
        '
        Me.Default1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default1.Location = New System.Drawing.Point(887, 43)
        Me.Default1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default1.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default1.MaxLength = 8
        Me.Default1.Name = "Default1"
        Me.Default1.Size = New System.Drawing.Size(61, 22)
        Me.Default1.TabIndex = 754
        Me.Default1.Text = "0"
        '
        'Default0
        '
        Me.Default0.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Default0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Default0.Location = New System.Drawing.Point(887, 19)
        Me.Default0.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Default0.MaximumSize = New System.Drawing.Size(61, 22)
        Me.Default0.MaxLength = 8
        Me.Default0.Name = "Default0"
        Me.Default0.Size = New System.Drawing.Size(61, 22)
        Me.Default0.TabIndex = 700
        Me.Default0.Text = "0"
        '
        'Label168
        '
        Me.Label168.AutoSize = True
        Me.Label168.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label168.Location = New System.Drawing.Point(641, 7)
        Me.Label168.Name = "Label168"
        Me.Label168.Size = New System.Drawing.Size(36, 12)
        Me.Label168.TabIndex = 840
        Me.Label168.Text = "Voltage"
        '
        'Label158
        '
        Me.Label158.AutoSize = True
        Me.Label158.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label158.Location = New System.Drawing.Point(569, 7)
        Me.Label158.Name = "Label158"
        Me.Label158.Size = New System.Drawing.Size(35, 12)
        Me.Label158.TabIndex = 839
        Me.Label158.Text = "Counts"
        '
        'Label157
        '
        Me.Label157.AutoSize = True
        Me.Label157.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label157.Location = New System.Drawing.Point(847, 7)
        Me.Label157.Name = "Label157"
        Me.Label157.Size = New System.Drawing.Size(27, 12)
        Me.Label157.TabIndex = 838
        Me.Label157.Text = "Delta"
        '
        'Label156
        '
        Me.Label156.AutoSize = True
        Me.Label156.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label156.Location = New System.Drawing.Point(789, 7)
        Me.Label156.Name = "Label156"
        Me.Label156.Size = New System.Drawing.Size(45, 12)
        Me.Label156.TabIndex = 801
        Me.Label156.Text = "Received"
        '
        'LabelTemperature2
        '
        Me.LabelTemperature2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelTemperature2.ForeColor = System.Drawing.Color.Black
        Me.LabelTemperature2.Location = New System.Drawing.Point(274, 9)
        Me.LabelTemperature2.Name = "LabelTemperature2"
        Me.LabelTemperature2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.LabelTemperature2.Size = New System.Drawing.Size(50, 30)
        Me.LabelTemperature2.TabIndex = 746
        Me.LabelTemperature2.Text = "#####"
        Me.LabelTemperature2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label155
        '
        Me.Label155.AutoSize = True
        Me.Label155.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label155.Location = New System.Drawing.Point(792, 487)
        Me.Label155.Name = "Label155"
        Me.Label155.Size = New System.Drawing.Size(47, 16)
        Me.Label155.TabIndex = 801
        Me.Label155.Text = "Status:"
        Me.Label155.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label229
        '
        Me.Label229.AutoSize = True
        Me.Label229.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label229.Location = New System.Drawing.Point(392, 519)
        Me.Label229.Name = "Label229"
        Me.Label229.Size = New System.Drawing.Size(130, 16)
        Me.Label229.TabIndex = 837
        Me.Label229.Text = "Internal Temp. DegC"
        '
        'LabelTemperature3
        '
        Me.LabelTemperature3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelTemperature3.Location = New System.Drawing.Point(298, 516)
        Me.LabelTemperature3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.LabelTemperature3.MaxLength = 8
        Me.LabelTemperature3.Name = "LabelTemperature3"
        Me.LabelTemperature3.ReadOnly = True
        Me.LabelTemperature3.Size = New System.Drawing.Size(91, 22)
        Me.LabelTemperature3.TabIndex = 836
        '
        'LabeldacSpan10Delta
        '
        Me.LabeldacSpan10Delta.AutoSize = True
        Me.LabeldacSpan10Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan10Delta.Location = New System.Drawing.Point(846, 285)
        Me.LabeldacSpan10Delta.Name = "LabeldacSpan10Delta"
        Me.LabeldacSpan10Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan10Delta.TabIndex = 835
        Me.LabeldacSpan10Delta.Text = "####"
        '
        'Label149
        '
        Me.Label149.AutoSize = True
        Me.Label149.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label149.Location = New System.Drawing.Point(947, 286)
        Me.Label149.Name = "Label149"
        Me.Label149.Size = New System.Drawing.Size(90, 16)
        Me.Label149.TabIndex = 834
        Me.Label149.Text = "Def. 11 counts"
        '
        'Label150
        '
        Me.Label150.AutoSize = True
        Me.Label150.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label150.Location = New System.Drawing.Point(481, 285)
        Me.Label150.Name = "Label150"
        Me.Label150.Size = New System.Drawing.Size(68, 16)
        Me.Label150.TabIndex = 801
        Me.Label150.Text = "10.22222V"
        Me.Label150.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ButtonDacSpan10Up
        '
        Me.ButtonDacSpan10Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan10Up.Location = New System.Drawing.Point(766, 281)
        Me.ButtonDacSpan10Up.Name = "ButtonDacSpan10Up"
        Me.ButtonDacSpan10Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan10Up.TabIndex = 832
        Me.ButtonDacSpan10Up.Text = ">"
        Me.ButtonDacSpan10Up.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan10
        '
        Me.ButtonDacSpan10.Location = New System.Drawing.Point(722, 281)
        Me.ButtonDacSpan10.Name = "ButtonDacSpan10"
        Me.ButtonDacSpan10.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan10.TabIndex = 830
        Me.ButtonDacSpan10.Text = "SET"
        Me.ButtonDacSpan10.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan10down
        '
        Me.ButtonDacSpan10down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan10down.Location = New System.Drawing.Point(701, 281)
        Me.ButtonDacSpan10down.Name = "ButtonDacSpan10down"
        Me.ButtonDacSpan10down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan10down.TabIndex = 831
        Me.ButtonDacSpan10down.Text = "<"
        Me.ButtonDacSpan10down.UseVisualStyleBackColor = True
        '
        'LabeldacSpan10Cal
        '
        Me.LabeldacSpan10Cal.AutoSize = True
        Me.LabeldacSpan10Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan10Cal.Location = New System.Drawing.Point(788, 286)
        Me.LabeldacSpan10Cal.Name = "LabeldacSpan10Cal"
        Me.LabeldacSpan10Cal.Size = New System.Drawing.Size(56, 16)
        Me.LabeldacSpan10Cal.TabIndex = 829
        Me.LabeldacSpan10Cal.Text = "#######"
        '
        'volts11
        '
        Me.volts11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts11.Location = New System.Drawing.Point(621, 282)
        Me.volts11.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts11.MaxLength = 8
        Me.volts11.Name = "volts11"
        Me.volts11.Size = New System.Drawing.Size(76, 22)
        Me.volts11.TabIndex = 801
        '
        'DacSpan10
        '
        Me.DacSpan10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DacSpan10.Location = New System.Drawing.Point(556, 282)
        Me.DacSpan10.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DacSpan10.MaxLength = 8
        Me.DacSpan10.Name = "DacSpan10"
        Me.DacSpan10.Size = New System.Drawing.Size(61, 22)
        Me.DacSpan10.TabIndex = 801
        '
        'ButtonDacSpan9Up
        '
        Me.ButtonDacSpan9Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan9Up.Location = New System.Drawing.Point(766, 257)
        Me.ButtonDacSpan9Up.Name = "ButtonDacSpan9Up"
        Me.ButtonDacSpan9Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan9Up.TabIndex = 827
        Me.ButtonDacSpan9Up.Text = ">"
        Me.ButtonDacSpan9Up.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan9
        '
        Me.ButtonDacSpan9.Location = New System.Drawing.Point(722, 257)
        Me.ButtonDacSpan9.Name = "ButtonDacSpan9"
        Me.ButtonDacSpan9.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan9.TabIndex = 808
        Me.ButtonDacSpan9.Text = "SET"
        Me.ButtonDacSpan9.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan9down
        '
        Me.ButtonDacSpan9down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan9down.Location = New System.Drawing.Point(701, 257)
        Me.ButtonDacSpan9down.Name = "ButtonDacSpan9down"
        Me.ButtonDacSpan9down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan9down.TabIndex = 826
        Me.ButtonDacSpan9down.Text = "<"
        Me.ButtonDacSpan9down.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan8
        '
        Me.ButtonDacSpan8.Location = New System.Drawing.Point(722, 233)
        Me.ButtonDacSpan8.Name = "ButtonDacSpan8"
        Me.ButtonDacSpan8.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan8.TabIndex = 807
        Me.ButtonDacSpan8.Text = "SET"
        Me.ButtonDacSpan8.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan8Up
        '
        Me.ButtonDacSpan8Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan8Up.Location = New System.Drawing.Point(766, 233)
        Me.ButtonDacSpan8Up.Name = "ButtonDacSpan8Up"
        Me.ButtonDacSpan8Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan8Up.TabIndex = 825
        Me.ButtonDacSpan8Up.Text = ">"
        Me.ButtonDacSpan8Up.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan7
        '
        Me.ButtonDacSpan7.Location = New System.Drawing.Point(722, 209)
        Me.ButtonDacSpan7.Name = "ButtonDacSpan7"
        Me.ButtonDacSpan7.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan7.TabIndex = 806
        Me.ButtonDacSpan7.Text = "SET"
        Me.ButtonDacSpan7.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan8down
        '
        Me.ButtonDacSpan8down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan8down.Location = New System.Drawing.Point(701, 233)
        Me.ButtonDacSpan8down.Name = "ButtonDacSpan8down"
        Me.ButtonDacSpan8down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan8down.TabIndex = 824
        Me.ButtonDacSpan8down.Text = "<"
        Me.ButtonDacSpan8down.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan6
        '
        Me.ButtonDacSpan6.Location = New System.Drawing.Point(722, 185)
        Me.ButtonDacSpan6.Name = "ButtonDacSpan6"
        Me.ButtonDacSpan6.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan6.TabIndex = 805
        Me.ButtonDacSpan6.Text = "SET"
        Me.ButtonDacSpan6.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan7Up
        '
        Me.ButtonDacSpan7Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan7Up.Location = New System.Drawing.Point(766, 209)
        Me.ButtonDacSpan7Up.Name = "ButtonDacSpan7Up"
        Me.ButtonDacSpan7Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan7Up.TabIndex = 823
        Me.ButtonDacSpan7Up.Text = ">"
        Me.ButtonDacSpan7Up.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan5
        '
        Me.ButtonDacSpan5.Location = New System.Drawing.Point(722, 161)
        Me.ButtonDacSpan5.Name = "ButtonDacSpan5"
        Me.ButtonDacSpan5.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan5.TabIndex = 804
        Me.ButtonDacSpan5.Text = "SET"
        Me.ButtonDacSpan5.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan7down
        '
        Me.ButtonDacSpan7down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan7down.Location = New System.Drawing.Point(701, 209)
        Me.ButtonDacSpan7down.Name = "ButtonDacSpan7down"
        Me.ButtonDacSpan7down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan7down.TabIndex = 822
        Me.ButtonDacSpan7down.Text = "<"
        Me.ButtonDacSpan7down.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan4
        '
        Me.ButtonDacSpan4.Location = New System.Drawing.Point(722, 137)
        Me.ButtonDacSpan4.Name = "ButtonDacSpan4"
        Me.ButtonDacSpan4.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan4.TabIndex = 803
        Me.ButtonDacSpan4.Text = "SET"
        Me.ButtonDacSpan4.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan6Up
        '
        Me.ButtonDacSpan6Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan6Up.Location = New System.Drawing.Point(766, 185)
        Me.ButtonDacSpan6Up.Name = "ButtonDacSpan6Up"
        Me.ButtonDacSpan6Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan6Up.TabIndex = 821
        Me.ButtonDacSpan6Up.Text = ">"
        Me.ButtonDacSpan6Up.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan6down
        '
        Me.ButtonDacSpan6down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan6down.Location = New System.Drawing.Point(701, 185)
        Me.ButtonDacSpan6down.Name = "ButtonDacSpan6down"
        Me.ButtonDacSpan6down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan6down.TabIndex = 820
        Me.ButtonDacSpan6down.Text = "<"
        Me.ButtonDacSpan6down.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan3
        '
        Me.ButtonDacSpan3.Location = New System.Drawing.Point(722, 113)
        Me.ButtonDacSpan3.Name = "ButtonDacSpan3"
        Me.ButtonDacSpan3.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan3.TabIndex = 802
        Me.ButtonDacSpan3.Text = "SET"
        Me.ButtonDacSpan3.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan5Up
        '
        Me.ButtonDacSpan5Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan5Up.Location = New System.Drawing.Point(766, 161)
        Me.ButtonDacSpan5Up.Name = "ButtonDacSpan5Up"
        Me.ButtonDacSpan5Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan5Up.TabIndex = 819
        Me.ButtonDacSpan5Up.Text = ">"
        Me.ButtonDacSpan5Up.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan2
        '
        Me.ButtonDacSpan2.Location = New System.Drawing.Point(722, 89)
        Me.ButtonDacSpan2.Name = "ButtonDacSpan2"
        Me.ButtonDacSpan2.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan2.TabIndex = 801
        Me.ButtonDacSpan2.Text = "SET"
        Me.ButtonDacSpan2.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan5down
        '
        Me.ButtonDacSpan5down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan5down.Location = New System.Drawing.Point(701, 161)
        Me.ButtonDacSpan5down.Name = "ButtonDacSpan5down"
        Me.ButtonDacSpan5down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan5down.TabIndex = 818
        Me.ButtonDacSpan5down.Text = "<"
        Me.ButtonDacSpan5down.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan1
        '
        Me.ButtonDacSpan1.Location = New System.Drawing.Point(722, 65)
        Me.ButtonDacSpan1.Name = "ButtonDacSpan1"
        Me.ButtonDacSpan1.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan1.TabIndex = 800
        Me.ButtonDacSpan1.Text = "SET"
        Me.ButtonDacSpan1.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan4Up
        '
        Me.ButtonDacSpan4Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan4Up.Location = New System.Drawing.Point(766, 137)
        Me.ButtonDacSpan4Up.Name = "ButtonDacSpan4Up"
        Me.ButtonDacSpan4Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan4Up.TabIndex = 817
        Me.ButtonDacSpan4Up.Text = ">"
        Me.ButtonDacSpan4Up.UseVisualStyleBackColor = True
        '
        'ButtonOutVdc
        '
        Me.ButtonOutVdc.Location = New System.Drawing.Point(722, 305)
        Me.ButtonOutVdc.Name = "ButtonOutVdc"
        Me.ButtonOutVdc.Size = New System.Drawing.Size(40, 24)
        Me.ButtonOutVdc.TabIndex = 799
        Me.ButtonOutVdc.Text = "SET"
        Me.ButtonOutVdc.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan4down
        '
        Me.ButtonDacSpan4down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan4down.Location = New System.Drawing.Point(701, 137)
        Me.ButtonDacSpan4down.Name = "ButtonDacSpan4down"
        Me.ButtonDacSpan4down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan4down.TabIndex = 816
        Me.ButtonDacSpan4down.Text = "<"
        Me.ButtonDacSpan4down.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan0
        '
        Me.ButtonDacSpan0.Location = New System.Drawing.Point(722, 41)
        Me.ButtonDacSpan0.Name = "ButtonDacSpan0"
        Me.ButtonDacSpan0.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacSpan0.TabIndex = 798
        Me.ButtonDacSpan0.Text = "SET"
        Me.ButtonDacSpan0.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan3Up
        '
        Me.ButtonDacSpan3Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan3Up.Location = New System.Drawing.Point(766, 113)
        Me.ButtonDacSpan3Up.Name = "ButtonDacSpan3Up"
        Me.ButtonDacSpan3Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan3Up.TabIndex = 815
        Me.ButtonDacSpan3Up.Text = ">"
        Me.ButtonDacSpan3Up.UseVisualStyleBackColor = True
        '
        'ButtonDacZero0
        '
        Me.ButtonDacZero0.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacZero0.Location = New System.Drawing.Point(722, 17)
        Me.ButtonDacZero0.Name = "ButtonDacZero0"
        Me.ButtonDacZero0.Size = New System.Drawing.Size(40, 24)
        Me.ButtonDacZero0.TabIndex = 797
        Me.ButtonDacZero0.Text = "SET"
        Me.ButtonDacZero0.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan3down
        '
        Me.ButtonDacSpan3down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan3down.Location = New System.Drawing.Point(701, 113)
        Me.ButtonDacSpan3down.Name = "ButtonDacSpan3down"
        Me.ButtonDacSpan3down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan3down.TabIndex = 814
        Me.ButtonDacSpan3down.Text = "<"
        Me.ButtonDacSpan3down.UseVisualStyleBackColor = True
        '
        'ButtonBattVdc
        '
        Me.ButtonBattVdc.Location = New System.Drawing.Point(722, 329)
        Me.ButtonBattVdc.Name = "ButtonBattVdc"
        Me.ButtonBattVdc.Size = New System.Drawing.Size(40, 24)
        Me.ButtonBattVdc.TabIndex = 796
        Me.ButtonBattVdc.Text = "SET"
        Me.ButtonBattVdc.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan2Up
        '
        Me.ButtonDacSpan2Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan2Up.Location = New System.Drawing.Point(766, 89)
        Me.ButtonDacSpan2Up.Name = "ButtonDacSpan2Up"
        Me.ButtonDacSpan2Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan2Up.TabIndex = 813
        Me.ButtonDacSpan2Up.Text = ">"
        Me.ButtonDacSpan2Up.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan2down
        '
        Me.ButtonDacSpan2down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan2down.Location = New System.Drawing.Point(701, 89)
        Me.ButtonDacSpan2down.Name = "ButtonDacSpan2down"
        Me.ButtonDacSpan2down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan2down.TabIndex = 812
        Me.ButtonDacSpan2down.Text = "<"
        Me.ButtonDacSpan2down.UseVisualStyleBackColor = True
        '
        'Label162
        '
        Me.Label162.AutoSize = True
        Me.Label162.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label162.Location = New System.Drawing.Point(179, 16)
        Me.Label162.Name = "Label162"
        Me.Label162.Size = New System.Drawing.Size(100, 16)
        Me.Label162.TabIndex = 703
        Me.Label162.Text = "External Temp.:"
        Me.Label162.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ButtonDacSpan1Up
        '
        Me.ButtonDacSpan1Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan1Up.Location = New System.Drawing.Point(766, 65)
        Me.ButtonDacSpan1Up.Name = "ButtonDacSpan1Up"
        Me.ButtonDacSpan1Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan1Up.TabIndex = 811
        Me.ButtonDacSpan1Up.Text = ">"
        Me.ButtonDacSpan1Up.UseVisualStyleBackColor = True
        '
        'ButtonDacSpan1down
        '
        Me.ButtonDacSpan1down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan1down.Location = New System.Drawing.Point(701, 65)
        Me.ButtonDacSpan1down.Name = "ButtonDacSpan1down"
        Me.ButtonDacSpan1down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan1down.TabIndex = 810
        Me.ButtonDacSpan1down.Text = "<"
        Me.ButtonDacSpan1down.UseVisualStyleBackColor = True
        '
        'ButtonChargeIUp
        '
        Me.ButtonChargeIUp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonChargeIUp.Location = New System.Drawing.Point(766, 353)
        Me.ButtonChargeIUp.Name = "ButtonChargeIUp"
        Me.ButtonChargeIUp.Size = New System.Drawing.Size(17, 24)
        Me.ButtonChargeIUp.TabIndex = 809
        Me.ButtonChargeIUp.Text = ">"
        Me.ButtonChargeIUp.UseVisualStyleBackColor = True
        '
        'ButtonBattVdcUp
        '
        Me.ButtonBattVdcUp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonBattVdcUp.Location = New System.Drawing.Point(766, 329)
        Me.ButtonBattVdcUp.Name = "ButtonBattVdcUp"
        Me.ButtonBattVdcUp.Size = New System.Drawing.Size(17, 24)
        Me.ButtonBattVdcUp.TabIndex = 808
        Me.ButtonBattVdcUp.Text = ">"
        Me.ButtonBattVdcUp.UseVisualStyleBackColor = True
        '
        'Label225
        '
        Me.Label225.AutoSize = True
        Me.Label225.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label225.Location = New System.Drawing.Point(351, 266)
        Me.Label225.Name = "Label225"
        Me.Label225.Size = New System.Drawing.Size(65, 16)
        Me.Label225.TabIndex = 791
        Me.Label225.Text = "(Optional)"
        '
        'ButtonOutVdcUp
        '
        Me.ButtonOutVdcUp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonOutVdcUp.Location = New System.Drawing.Point(766, 305)
        Me.ButtonOutVdcUp.Name = "ButtonOutVdcUp"
        Me.ButtonOutVdcUp.Size = New System.Drawing.Size(17, 24)
        Me.ButtonOutVdcUp.TabIndex = 807
        Me.ButtonOutVdcUp.Text = ">"
        Me.ButtonOutVdcUp.UseVisualStyleBackColor = True
        '
        'Label224
        '
        Me.Label224.AutoSize = True
        Me.Label224.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label224.Location = New System.Drawing.Point(351, 236)
        Me.Label224.Name = "Label224"
        Me.Label224.Size = New System.Drawing.Size(65, 16)
        Me.Label224.TabIndex = 790
        Me.Label224.Text = "(Optional)"
        '
        'ButtonDacSpan0Up
        '
        Me.ButtonDacSpan0Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan0Up.Location = New System.Drawing.Point(766, 41)
        Me.ButtonDacSpan0Up.Name = "ButtonDacSpan0Up"
        Me.ButtonDacSpan0Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan0Up.TabIndex = 806
        Me.ButtonDacSpan0Up.Text = ">"
        Me.ButtonDacSpan0Up.UseVisualStyleBackColor = True
        '
        'ButtonDacZero0Up
        '
        Me.ButtonDacZero0Up.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacZero0Up.Location = New System.Drawing.Point(766, 17)
        Me.ButtonDacZero0Up.Name = "ButtonDacZero0Up"
        Me.ButtonDacZero0Up.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacZero0Up.TabIndex = 805
        Me.ButtonDacZero0Up.Text = ">"
        Me.ButtonDacZero0Up.UseVisualStyleBackColor = True
        '
        'TextBoxUser
        '
        Me.TextBoxUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxUser.Location = New System.Drawing.Point(5, 516)
        Me.TextBoxUser.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxUser.MaxLength = 25
        Me.TextBoxUser.Name = "TextBoxUser"
        Me.TextBoxUser.Size = New System.Drawing.Size(168, 22)
        Me.TextBoxUser.TabIndex = 787
        '
        'ButtonChargeIdown
        '
        Me.ButtonChargeIdown.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonChargeIdown.Location = New System.Drawing.Point(701, 353)
        Me.ButtonChargeIdown.Name = "ButtonChargeIdown"
        Me.ButtonChargeIdown.Size = New System.Drawing.Size(17, 24)
        Me.ButtonChargeIdown.TabIndex = 804
        Me.ButtonChargeIdown.Text = "<"
        Me.ButtonChargeIdown.UseVisualStyleBackColor = True
        '
        'Label222
        '
        Me.Label222.AutoSize = True
        Me.Label222.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label222.Location = New System.Drawing.Point(176, 519)
        Me.Label222.Name = "Label222"
        Me.Label222.Size = New System.Drawing.Size(44, 16)
        Me.Label222.TabIndex = 788
        Me.Label222.Text = "Name"
        '
        'ButtonBattVdcdown
        '
        Me.ButtonBattVdcdown.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonBattVdcdown.Location = New System.Drawing.Point(701, 329)
        Me.ButtonBattVdcdown.Name = "ButtonBattVdcdown"
        Me.ButtonBattVdcdown.Size = New System.Drawing.Size(17, 24)
        Me.ButtonBattVdcdown.TabIndex = 803
        Me.ButtonBattVdcdown.Text = "<"
        Me.ButtonBattVdcdown.UseVisualStyleBackColor = True
        '
        'TextBox3458Asn
        '
        Me.TextBox3458Asn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox3458Asn.Location = New System.Drawing.Point(5, 490)
        Me.TextBox3458Asn.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBox3458Asn.MaxLength = 15
        Me.TextBox3458Asn.Name = "TextBox3458Asn"
        Me.TextBox3458Asn.Size = New System.Drawing.Size(91, 22)
        Me.TextBox3458Asn.TabIndex = 700
        '
        'ButtonOutVdcdown
        '
        Me.ButtonOutVdcdown.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonOutVdcdown.Location = New System.Drawing.Point(701, 305)
        Me.ButtonOutVdcdown.Name = "ButtonOutVdcdown"
        Me.ButtonOutVdcdown.Size = New System.Drawing.Size(17, 24)
        Me.ButtonOutVdcdown.TabIndex = 802
        Me.ButtonOutVdcdown.Text = "<"
        Me.ButtonOutVdcdown.UseVisualStyleBackColor = True
        '
        'Label73
        '
        Me.Label73.AutoSize = True
        Me.Label73.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label73.Location = New System.Drawing.Point(95, 493)
        Me.Label73.Name = "Label73"
        Me.Label73.Size = New System.Drawing.Size(133, 16)
        Me.Label73.TabIndex = 701
        Me.Label73.Text = "3458A Serial Number"
        '
        'ButtonDacSpan0down
        '
        Me.ButtonDacSpan0down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacSpan0down.Location = New System.Drawing.Point(701, 41)
        Me.ButtonDacSpan0down.Name = "ButtonDacSpan0down"
        Me.ButtonDacSpan0down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacSpan0down.TabIndex = 801
        Me.ButtonDacSpan0down.Text = "<"
        Me.ButtonDacSpan0down.UseVisualStyleBackColor = True
        '
        'Label290
        '
        Me.Label290.AutoSize = True
        Me.Label290.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label290.ForeColor = System.Drawing.Color.Red
        Me.Label290.Location = New System.Drawing.Point(3, 284)
        Me.Label290.Name = "Label290"
        Me.Label290.Size = New System.Drawing.Size(19, 16)
        Me.Label290.TabIndex = 786
        Me.Label290.Text = "8."
        Me.Label290.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ButtonDacZero0down
        '
        Me.ButtonDacZero0down.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDacZero0down.Location = New System.Drawing.Point(701, 17)
        Me.ButtonDacZero0down.Name = "ButtonDacZero0down"
        Me.ButtonDacZero0down.Size = New System.Drawing.Size(17, 24)
        Me.ButtonDacZero0down.TabIndex = 800
        Me.ButtonDacZero0down.Text = "<"
        Me.ButtonDacZero0down.UseVisualStyleBackColor = True
        '
        'Label289
        '
        Me.Label289.AutoSize = True
        Me.Label289.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label289.ForeColor = System.Drawing.Color.Red
        Me.Label289.Location = New System.Drawing.Point(3, 326)
        Me.Label289.Name = "Label289"
        Me.Label289.Size = New System.Drawing.Size(19, 16)
        Me.Label289.TabIndex = 785
        Me.Label289.Text = "9."
        Me.Label289.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label288
        '
        Me.Label288.AutoSize = True
        Me.Label288.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label288.ForeColor = System.Drawing.Color.Red
        Me.Label288.Location = New System.Drawing.Point(3, 242)
        Me.Label288.Name = "Label288"
        Me.Label288.Size = New System.Drawing.Size(19, 16)
        Me.Label288.TabIndex = 784
        Me.Label288.Text = "7."
        Me.Label288.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label287
        '
        Me.Label287.AutoSize = True
        Me.Label287.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label287.ForeColor = System.Drawing.Color.Red
        Me.Label287.Location = New System.Drawing.Point(3, 200)
        Me.Label287.Name = "Label287"
        Me.Label287.Size = New System.Drawing.Size(19, 16)
        Me.Label287.TabIndex = 783
        Me.Label287.Text = "6."
        Me.Label287.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label286
        '
        Me.Label286.AutoSize = True
        Me.Label286.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label286.ForeColor = System.Drawing.Color.Red
        Me.Label286.Location = New System.Drawing.Point(3, 172)
        Me.Label286.Name = "Label286"
        Me.Label286.Size = New System.Drawing.Size(19, 16)
        Me.Label286.TabIndex = 782
        Me.Label286.Text = "5."
        Me.Label286.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label285
        '
        Me.Label285.AutoSize = True
        Me.Label285.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label285.ForeColor = System.Drawing.Color.Red
        Me.Label285.Location = New System.Drawing.Point(3, 145)
        Me.Label285.Name = "Label285"
        Me.Label285.Size = New System.Drawing.Size(19, 16)
        Me.Label285.TabIndex = 781
        Me.Label285.Text = "4."
        Me.Label285.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label284
        '
        Me.Label284.AutoSize = True
        Me.Label284.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label284.ForeColor = System.Drawing.Color.Red
        Me.Label284.Location = New System.Drawing.Point(3, 75)
        Me.Label284.Name = "Label284"
        Me.Label284.Size = New System.Drawing.Size(19, 16)
        Me.Label284.TabIndex = 780
        Me.Label284.Text = "2."
        Me.Label284.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label283
        '
        Me.Label283.AutoSize = True
        Me.Label283.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label283.ForeColor = System.Drawing.Color.Red
        Me.Label283.Location = New System.Drawing.Point(3, 103)
        Me.Label283.Name = "Label283"
        Me.Label283.Size = New System.Drawing.Size(19, 16)
        Me.Label283.TabIndex = 779
        Me.Label283.Text = "3."
        Me.Label283.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label282
        '
        Me.Label282.AutoSize = True
        Me.Label282.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label282.ForeColor = System.Drawing.Color.Red
        Me.Label282.Location = New System.Drawing.Point(3, 48)
        Me.Label282.Name = "Label282"
        Me.Label282.Size = New System.Drawing.Size(19, 16)
        Me.Label282.TabIndex = 778
        Me.Label282.Text = "1."
        Me.Label282.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label281
        '
        Me.Label281.AutoSize = True
        Me.Label281.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label281.Location = New System.Drawing.Point(27, 77)
        Me.Label281.Name = "Label281"
        Me.Label281.Size = New System.Drawing.Size(131, 13)
        Me.Label281.TabIndex = 777
        Me.Label281.Text = "Load Default Counts (opt.)"
        '
        'Label280
        '
        Me.Label280.AutoSize = True
        Me.Label280.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label280.Location = New System.Drawing.Point(889, 355)
        Me.Label280.Name = "Label280"
        Me.Label280.Size = New System.Drawing.Size(153, 130)
        Me.Label280.TabIndex = 776
        Me.Label280.Text = resources.GetString("Label280.Text")
        '
        'Label279
        '
        Me.Label279.AutoSize = True
        Me.Label279.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label279.Location = New System.Drawing.Point(947, 262)
        Me.Label279.Name = "Label279"
        Me.Label279.Size = New System.Drawing.Size(90, 16)
        Me.Label279.TabIndex = 773
        Me.Label279.Text = "Def. 10 counts"
        '
        'Label278
        '
        Me.Label278.AutoSize = True
        Me.Label278.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label278.Location = New System.Drawing.Point(947, 238)
        Me.Label278.Name = "Label278"
        Me.Label278.Size = New System.Drawing.Size(83, 16)
        Me.Label278.TabIndex = 772
        Me.Label278.Text = "Def. 9 counts"
        '
        'Label277
        '
        Me.Label277.AutoSize = True
        Me.Label277.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label277.Location = New System.Drawing.Point(947, 214)
        Me.Label277.Name = "Label277"
        Me.Label277.Size = New System.Drawing.Size(83, 16)
        Me.Label277.TabIndex = 771
        Me.Label277.Text = "Def. 8 counts"
        '
        'Label161
        '
        Me.Label161.AutoSize = True
        Me.Label161.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label161.Location = New System.Drawing.Point(947, 190)
        Me.Label161.Name = "Label161"
        Me.Label161.Size = New System.Drawing.Size(83, 16)
        Me.Label161.TabIndex = 770
        Me.Label161.Text = "Def. 7 counts"
        '
        'Label160
        '
        Me.Label160.AutoSize = True
        Me.Label160.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label160.Location = New System.Drawing.Point(947, 166)
        Me.Label160.Name = "Label160"
        Me.Label160.Size = New System.Drawing.Size(83, 16)
        Me.Label160.TabIndex = 769
        Me.Label160.Text = "Def. 6 counts"
        '
        'Label159
        '
        Me.Label159.AutoSize = True
        Me.Label159.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label159.Location = New System.Drawing.Point(947, 142)
        Me.Label159.Name = "Label159"
        Me.Label159.Size = New System.Drawing.Size(83, 16)
        Me.Label159.TabIndex = 768
        Me.Label159.Text = "Def. 5 counts"
        '
        'Label137
        '
        Me.Label137.AutoSize = True
        Me.Label137.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label137.Location = New System.Drawing.Point(947, 118)
        Me.Label137.Name = "Label137"
        Me.Label137.Size = New System.Drawing.Size(83, 16)
        Me.Label137.TabIndex = 767
        Me.Label137.Text = "Def. 4 counts"
        '
        'Label136
        '
        Me.Label136.AutoSize = True
        Me.Label136.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label136.Location = New System.Drawing.Point(947, 94)
        Me.Label136.Name = "Label136"
        Me.Label136.Size = New System.Drawing.Size(83, 16)
        Me.Label136.TabIndex = 766
        Me.Label136.Text = "Def. 3 counts"
        '
        'Label124
        '
        Me.Label124.AutoSize = True
        Me.Label124.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label124.Location = New System.Drawing.Point(947, 70)
        Me.Label124.Name = "Label124"
        Me.Label124.Size = New System.Drawing.Size(83, 16)
        Me.Label124.TabIndex = 765
        Me.Label124.Text = "Def. 2 counts"
        '
        'Label104
        '
        Me.Label104.AutoSize = True
        Me.Label104.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label104.Location = New System.Drawing.Point(947, 46)
        Me.Label104.Name = "Label104"
        Me.Label104.Size = New System.Drawing.Size(83, 16)
        Me.Label104.TabIndex = 764
        Me.Label104.Text = "Def. 1 counts"
        '
        'Label55
        '
        Me.Label55.AutoSize = True
        Me.Label55.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label55.Location = New System.Drawing.Point(947, 22)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(83, 16)
        Me.Label55.TabIndex = 700
        Me.Label55.Text = "Def. 0 counts"
        '
        'DACcalparamVAL
        '
        Me.DACcalparamVAL.AutoSize = True
        Me.DACcalparamVAL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DACcalparamVAL.ForeColor = System.Drawing.Color.Black
        Me.DACcalparamVAL.Location = New System.Drawing.Point(217, 39)
        Me.DACcalparamVAL.Name = "DACcalparamVAL"
        Me.DACcalparamVAL.Size = New System.Drawing.Size(21, 16)
        Me.DACcalparamVAL.TabIndex = 556
        Me.DACcalparamVAL.Text = "##"
        Me.DACcalparamVAL.Visible = False
        '
        'RadioPDF
        '
        Me.RadioPDF.AutoSize = True
        Me.RadioPDF.Checked = True
        Me.RadioPDF.Location = New System.Drawing.Point(164, 334)
        Me.RadioPDF.Name = "RadioPDF"
        Me.RadioPDF.Size = New System.Drawing.Size(62, 17)
        Me.RadioPDF.TabIndex = 753
        Me.RadioPDF.TabStop = True
        Me.RadioPDF.Text = "PDF file"
        Me.RadioPDF.UseVisualStyleBackColor = True
        '
        'PDVS2counter
        '
        Me.PDVS2counter.AutoSize = True
        Me.PDVS2counter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PDVS2counter.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.PDVS2counter.Location = New System.Drawing.Point(358, 334)
        Me.PDVS2counter.Name = "PDVS2counter"
        Me.PDVS2counter.Size = New System.Drawing.Size(35, 16)
        Me.PDVS2counter.TabIndex = 555
        Me.PDVS2counter.Text = "####"
        Me.PDVS2counter.Visible = False
        '
        'RadioMSWord
        '
        Me.RadioMSWord.AutoSize = True
        Me.RadioMSWord.Location = New System.Drawing.Point(164, 317)
        Me.RadioMSWord.Name = "RadioMSWord"
        Me.RadioMSWord.Size = New System.Drawing.Size(170, 17)
        Me.RadioMSWord.TabIndex = 752
        Me.RadioMSWord.Text = "MS Word file (req's MS Offfice)"
        Me.RadioMSWord.UseVisualStyleBackColor = True
        '
        'Label254
        '
        Me.Label254.AutoSize = True
        Me.Label254.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label254.Location = New System.Drawing.Point(63, 174)
        Me.Label254.Name = "Label254"
        Me.Label254.Size = New System.Drawing.Size(54, 13)
        Me.Label254.TabIndex = 751
        Me.Label254.Text = "Set NPLC"
        '
        'Label163
        '
        Me.Label163.AutoSize = True
        Me.Label163.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label163.Location = New System.Drawing.Point(322, 16)
        Me.Label163.Name = "Label163"
        Me.Label163.Size = New System.Drawing.Size(42, 16)
        Me.Label163.TabIndex = 736
        Me.Label163.Text = "DegC"
        '
        'Label167
        '
        Me.Label167.AutoSize = True
        Me.Label167.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label167.Location = New System.Drawing.Point(392, 493)
        Me.Label167.Name = "Label167"
        Me.Label167.Size = New System.Drawing.Size(83, 16)
        Me.Label167.TabIndex = 745
        Me.Label167.Text = "Power Mode"
        '
        'Label166
        '
        Me.Label166.AutoSize = True
        Me.Label166.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label166.Location = New System.Drawing.Point(392, 467)
        Me.Label166.Name = "Label166"
        Me.Label166.Size = New System.Drawing.Size(49, 16)
        Me.Label166.TabIndex = 744
        Me.Label166.Text = "Battery"
        '
        'Label165
        '
        Me.Label165.AutoSize = True
        Me.Label165.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label165.Location = New System.Drawing.Point(392, 441)
        Me.Label165.Name = "Label165"
        Me.Label165.Size = New System.Drawing.Size(42, 16)
        Me.Label165.TabIndex = 743
        Me.Label165.Text = "Mode"
        '
        'TextBoxLabelBatteryCharge
        '
        Me.TextBoxLabelBatteryCharge.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxLabelBatteryCharge.Location = New System.Drawing.Point(298, 490)
        Me.TextBoxLabelBatteryCharge.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxLabelBatteryCharge.MaxLength = 8
        Me.TextBoxLabelBatteryCharge.Name = "TextBoxLabelBatteryCharge"
        Me.TextBoxLabelBatteryCharge.ReadOnly = True
        Me.TextBoxLabelBatteryCharge.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxLabelBatteryCharge.TabIndex = 742
        '
        'TextBoxLabelBatteryLowInd
        '
        Me.TextBoxLabelBatteryLowInd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxLabelBatteryLowInd.Location = New System.Drawing.Point(298, 464)
        Me.TextBoxLabelBatteryLowInd.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxLabelBatteryLowInd.MaxLength = 8
        Me.TextBoxLabelBatteryLowInd.Name = "TextBoxLabelBatteryLowInd"
        Me.TextBoxLabelBatteryLowInd.ReadOnly = True
        Me.TextBoxLabelBatteryLowInd.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxLabelBatteryLowInd.TabIndex = 741
        '
        'TextBoxLabelMode
        '
        Me.TextBoxLabelMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxLabelMode.Location = New System.Drawing.Point(298, 438)
        Me.TextBoxLabelMode.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxLabelMode.MaxLength = 8
        Me.TextBoxLabelMode.Name = "TextBoxLabelMode"
        Me.TextBoxLabelMode.ReadOnly = True
        Me.TextBoxLabelMode.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxLabelMode.TabIndex = 740
        '
        'TextBoxLabelBatteryI
        '
        Me.TextBoxLabelBatteryI.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxLabelBatteryI.Location = New System.Drawing.Point(298, 412)
        Me.TextBoxLabelBatteryI.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxLabelBatteryI.MaxLength = 8
        Me.TextBoxLabelBatteryI.Name = "TextBoxLabelBatteryI"
        Me.TextBoxLabelBatteryI.ReadOnly = True
        Me.TextBoxLabelBatteryI.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxLabelBatteryI.TabIndex = 739
        '
        'TextBoxLabelBatteryV
        '
        Me.TextBoxLabelBatteryV.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxLabelBatteryV.Location = New System.Drawing.Point(298, 386)
        Me.TextBoxLabelBatteryV.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxLabelBatteryV.MaxLength = 8
        Me.TextBoxLabelBatteryV.Name = "TextBoxLabelBatteryV"
        Me.TextBoxLabelBatteryV.ReadOnly = True
        Me.TextBoxLabelBatteryV.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxLabelBatteryV.TabIndex = 738
        '
        'TextBoxLabelOutputVFeedback
        '
        Me.TextBoxLabelOutputVFeedback.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxLabelOutputVFeedback.Location = New System.Drawing.Point(298, 360)
        Me.TextBoxLabelOutputVFeedback.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBoxLabelOutputVFeedback.MaxLength = 8
        Me.TextBoxLabelOutputVFeedback.Name = "TextBoxLabelOutputVFeedback"
        Me.TextBoxLabelOutputVFeedback.ReadOnly = True
        Me.TextBoxLabelOutputVFeedback.Size = New System.Drawing.Size(91, 22)
        Me.TextBoxLabelOutputVFeedback.TabIndex = 703
        '
        'Label118
        '
        Me.Label118.AutoSize = True
        Me.Label118.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label118.Location = New System.Drawing.Point(2, 10)
        Me.Label118.Name = "Label118"
        Me.Label118.Size = New System.Drawing.Size(166, 18)
        Me.Label118.TabIndex = 737
        Me.Label118.Text = "CALIBRATION / ADJ."
        Me.Label118.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ButtonSetXYcalvars
        '
        Me.ButtonSetXYcalvars.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSetXYcalvars.Location = New System.Drawing.Point(337, 294)
        Me.ButtonSetXYcalvars.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.ButtonSetXYcalvars.Name = "ButtonSetXYcalvars"
        Me.ButtonSetXYcalvars.Size = New System.Drawing.Size(120, 37)
        Me.ButtonSetXYcalvars.TabIndex = 734
        Me.ButtonSetXYcalvars.Text = "Send XY Cal. from" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "app to PDVS2mini"
        Me.ButtonSetXYcalvars.UseVisualStyleBackColor = True
        Me.ButtonSetXYcalvars.Visible = False
        '
        'LabeldacSpan9Delta
        '
        Me.LabeldacSpan9Delta.AutoSize = True
        Me.LabeldacSpan9Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan9Delta.Location = New System.Drawing.Point(846, 261)
        Me.LabeldacSpan9Delta.Name = "LabeldacSpan9Delta"
        Me.LabeldacSpan9Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan9Delta.TabIndex = 733
        Me.LabeldacSpan9Delta.Text = "####"
        '
        'LabeldacSpan8Delta
        '
        Me.LabeldacSpan8Delta.AutoSize = True
        Me.LabeldacSpan8Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan8Delta.Location = New System.Drawing.Point(846, 237)
        Me.LabeldacSpan8Delta.Name = "LabeldacSpan8Delta"
        Me.LabeldacSpan8Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan8Delta.TabIndex = 732
        Me.LabeldacSpan8Delta.Text = "####"
        '
        'LabeldacSpan7Delta
        '
        Me.LabeldacSpan7Delta.AutoSize = True
        Me.LabeldacSpan7Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan7Delta.Location = New System.Drawing.Point(846, 213)
        Me.LabeldacSpan7Delta.Name = "LabeldacSpan7Delta"
        Me.LabeldacSpan7Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan7Delta.TabIndex = 731
        Me.LabeldacSpan7Delta.Text = "####"
        '
        'LabeldacSpan6Delta
        '
        Me.LabeldacSpan6Delta.AutoSize = True
        Me.LabeldacSpan6Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan6Delta.Location = New System.Drawing.Point(846, 189)
        Me.LabeldacSpan6Delta.Name = "LabeldacSpan6Delta"
        Me.LabeldacSpan6Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan6Delta.TabIndex = 730
        Me.LabeldacSpan6Delta.Text = "####"
        '
        'LabeldacSpan5Delta
        '
        Me.LabeldacSpan5Delta.AutoSize = True
        Me.LabeldacSpan5Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan5Delta.Location = New System.Drawing.Point(846, 165)
        Me.LabeldacSpan5Delta.Name = "LabeldacSpan5Delta"
        Me.LabeldacSpan5Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan5Delta.TabIndex = 729
        Me.LabeldacSpan5Delta.Text = "####"
        '
        'LabeldacSpan4Delta
        '
        Me.LabeldacSpan4Delta.AutoSize = True
        Me.LabeldacSpan4Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan4Delta.Location = New System.Drawing.Point(846, 141)
        Me.LabeldacSpan4Delta.Name = "LabeldacSpan4Delta"
        Me.LabeldacSpan4Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan4Delta.TabIndex = 728
        Me.LabeldacSpan4Delta.Text = "####"
        '
        'LabeldacSpan3Delta
        '
        Me.LabeldacSpan3Delta.AutoSize = True
        Me.LabeldacSpan3Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan3Delta.Location = New System.Drawing.Point(846, 117)
        Me.LabeldacSpan3Delta.Name = "LabeldacSpan3Delta"
        Me.LabeldacSpan3Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan3Delta.TabIndex = 727
        Me.LabeldacSpan3Delta.Text = "####"
        '
        'LabeldacSpan2Delta
        '
        Me.LabeldacSpan2Delta.AutoSize = True
        Me.LabeldacSpan2Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan2Delta.Location = New System.Drawing.Point(846, 93)
        Me.LabeldacSpan2Delta.Name = "LabeldacSpan2Delta"
        Me.LabeldacSpan2Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan2Delta.TabIndex = 726
        Me.LabeldacSpan2Delta.Text = "####"
        '
        'LabeldacSpan1Delta
        '
        Me.LabeldacSpan1Delta.AutoSize = True
        Me.LabeldacSpan1Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan1Delta.Location = New System.Drawing.Point(846, 69)
        Me.LabeldacSpan1Delta.Name = "LabeldacSpan1Delta"
        Me.LabeldacSpan1Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan1Delta.TabIndex = 725
        Me.LabeldacSpan1Delta.Text = "####"
        '
        'LabeldacSpan0Delta
        '
        Me.LabeldacSpan0Delta.AutoSize = True
        Me.LabeldacSpan0Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacSpan0Delta.Location = New System.Drawing.Point(846, 45)
        Me.LabeldacSpan0Delta.Name = "LabeldacSpan0Delta"
        Me.LabeldacSpan0Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacSpan0Delta.TabIndex = 724
        Me.LabeldacSpan0Delta.Text = "####"
        '
        'LabeldacZero0Delta
        '
        Me.LabeldacZero0Delta.AutoSize = True
        Me.LabeldacZero0Delta.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabeldacZero0Delta.Location = New System.Drawing.Point(846, 21)
        Me.LabeldacZero0Delta.Name = "LabeldacZero0Delta"
        Me.LabeldacZero0Delta.Size = New System.Drawing.Size(35, 16)
        Me.LabeldacZero0Delta.TabIndex = 703
        Me.LabeldacZero0Delta.Text = "####"
        '
        'volts030000
        '
        Me.volts030000.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts030000.Location = New System.Drawing.Point(621, 498)
        Me.volts030000.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts030000.MaxLength = 8
        Me.volts030000.Name = "volts030000"
        Me.volts030000.Size = New System.Drawing.Size(76, 22)
        Me.volts030000.TabIndex = 721
        '
        'Button030000
        '
        Me.Button030000.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button030000.Location = New System.Drawing.Point(701, 497)
        Me.Button030000.Name = "Button030000"
        Me.Button030000.Size = New System.Drawing.Size(82, 24)
        Me.Button030000.TabIndex = 720
        Me.Button030000.Text = "SET 300mV"
        Me.Button030000.UseVisualStyleBackColor = True
        '
        'volts050000
        '
        Me.volts050000.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts050000.Location = New System.Drawing.Point(621, 522)
        Me.volts050000.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts050000.MaxLength = 8
        Me.volts050000.Name = "volts050000"
        Me.volts050000.Size = New System.Drawing.Size(76, 22)
        Me.volts050000.TabIndex = 719
        '
        'volts020000
        '
        Me.volts020000.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts020000.Location = New System.Drawing.Point(621, 474)
        Me.volts020000.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts020000.MaxLength = 8
        Me.volts020000.Name = "volts020000"
        Me.volts020000.Size = New System.Drawing.Size(76, 22)
        Me.volts020000.TabIndex = 718
        '
        'volts010000
        '
        Me.volts010000.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts010000.Location = New System.Drawing.Point(621, 450)
        Me.volts010000.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts010000.MaxLength = 8
        Me.volts010000.Name = "volts010000"
        Me.volts010000.Size = New System.Drawing.Size(76, 22)
        Me.volts010000.TabIndex = 717
        '
        'volts001000
        '
        Me.volts001000.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts001000.Location = New System.Drawing.Point(621, 426)
        Me.volts001000.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts001000.MaxLength = 8
        Me.volts001000.Name = "volts001000"
        Me.volts001000.Size = New System.Drawing.Size(76, 22)
        Me.volts001000.TabIndex = 716
        '
        'volts000100
        '
        Me.volts000100.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts000100.Location = New System.Drawing.Point(621, 402)
        Me.volts000100.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts000100.MaxLength = 8
        Me.volts000100.Name = "volts000100"
        Me.volts000100.Size = New System.Drawing.Size(76, 22)
        Me.volts000100.TabIndex = 715
        '
        'Button050000
        '
        Me.Button050000.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button050000.Location = New System.Drawing.Point(701, 521)
        Me.Button050000.Name = "Button050000"
        Me.Button050000.Size = New System.Drawing.Size(82, 24)
        Me.Button050000.TabIndex = 713
        Me.Button050000.Text = "SET 500mV"
        Me.Button050000.UseVisualStyleBackColor = True
        '
        'Button020000
        '
        Me.Button020000.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button020000.Location = New System.Drawing.Point(701, 473)
        Me.Button020000.Name = "Button020000"
        Me.Button020000.Size = New System.Drawing.Size(82, 24)
        Me.Button020000.TabIndex = 708
        Me.Button020000.Text = "SET 200mV"
        Me.Button020000.UseVisualStyleBackColor = True
        '
        'Button010000
        '
        Me.Button010000.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button010000.Location = New System.Drawing.Point(701, 449)
        Me.Button010000.Name = "Button010000"
        Me.Button010000.Size = New System.Drawing.Size(82, 24)
        Me.Button010000.TabIndex = 700
        Me.Button010000.Text = "SET 100mV"
        Me.Button010000.UseVisualStyleBackColor = True
        '
        'Button001000
        '
        Me.Button001000.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button001000.Location = New System.Drawing.Point(701, 425)
        Me.Button001000.Name = "Button001000"
        Me.Button001000.Size = New System.Drawing.Size(82, 24)
        Me.Button001000.TabIndex = 700
        Me.Button001000.Text = "SET 10mV"
        Me.Button001000.UseVisualStyleBackColor = True
        '
        'Button000010
        '
        Me.Button000010.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button000010.Location = New System.Drawing.Point(701, 377)
        Me.Button000010.Name = "Button000010"
        Me.Button000010.Size = New System.Drawing.Size(82, 24)
        Me.Button000010.TabIndex = 707
        Me.Button000010.Text = "SET 0.1mV"
        Me.Button000010.UseVisualStyleBackColor = True
        '
        'volts000010
        '
        Me.volts000010.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volts000010.Location = New System.Drawing.Point(621, 378)
        Me.volts000010.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.volts000010.MaxLength = 8
        Me.volts000010.Name = "volts000010"
        Me.volts000010.Size = New System.Drawing.Size(76, 22)
        Me.volts000010.TabIndex = 700
        '
        'Button000100
        '
        Me.Button000100.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button000100.Location = New System.Drawing.Point(701, 401)
        Me.Button000100.Name = "Button000100"
        Me.Button000100.Size = New System.Drawing.Size(82, 24)
        Me.Button000100.TabIndex = 695
        Me.Button000100.Text = "SET 1mV"
        Me.Button000100.UseVisualStyleBackColor = True
        '
        'NPLC
        '
        Me.NPLC.AutoSize = True
        Me.NPLC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NPLC.ForeColor = System.Drawing.Color.Black
        Me.NPLC.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.NPLC.Location = New System.Drawing.Point(243, 68)
        Me.NPLC.Name = "NPLC"
        Me.NPLC.Size = New System.Drawing.Size(28, 16)
        Me.NPLC.TabIndex = 557
        Me.NPLC.Text = "###"
        '
        'Label83
        '
        Me.Label83.AutoSize = True
        Me.Label83.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label83.ForeColor = System.Drawing.Color.Black
        Me.Label83.Location = New System.Drawing.Point(178, 68)
        Me.Label83.Name = "Label83"
        Me.Label83.Size = New System.Drawing.Size(52, 16)
        Me.Label83.TabIndex = 558
        Me.Label83.Text = "NPLC ="
        '
        'Label76
        '
        Me.Label76.AutoSize = True
        Me.Label76.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label76.Location = New System.Drawing.Point(392, 389)
        Me.Label76.Name = "Label76"
        Me.Label76.Size = New System.Drawing.Size(31, 16)
        Me.Label76.TabIndex = 525
        Me.Label76.Text = "Vdc"
        '
        'Label75
        '
        Me.Label75.AutoSize = True
        Me.Label75.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label75.Location = New System.Drawing.Point(392, 415)
        Me.Label75.Name = "Label75"
        Me.Label75.Size = New System.Drawing.Size(81, 16)
        Me.Label75.TabIndex = 526
        Me.Label75.Text = "mA - Charge"
        '
        'Label74
        '
        Me.Label74.AutoSize = True
        Me.Label74.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label74.Location = New System.Drawing.Point(392, 363)
        Me.Label74.Name = "Label74"
        Me.Label74.Size = New System.Drawing.Size(84, 16)
        Me.Label74.TabIndex = 527
        Me.Label74.Text = "Vdc (low res)"
        '
        'LabelCalCount
        '
        Me.LabelCalCount.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCalCount.ForeColor = System.Drawing.Color.LimeGreen
        Me.LabelCalCount.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.LabelCalCount.Location = New System.Drawing.Point(791, 504)
        Me.LabelCalCount.Name = "LabelCalCount"
        Me.LabelCalCount.Size = New System.Drawing.Size(242, 37)
        Me.LabelCalCount.TabIndex = 615
        Me.LabelCalCount.Text = "######"
        '
        'ButtonChargeI
        '
        Me.ButtonChargeI.Location = New System.Drawing.Point(722, 353)
        Me.ButtonChargeI.Name = "ButtonChargeI"
        Me.ButtonChargeI.Size = New System.Drawing.Size(40, 24)
        Me.ButtonChargeI.TabIndex = 602
        Me.ButtonChargeI.Text = "SET"
        Me.ButtonChargeI.UseVisualStyleBackColor = True
        '
        'PDVS2delay
        '
        Me.PDVS2delay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PDVS2delay.Location = New System.Drawing.Point(286, 192)
        Me.PDVS2delay.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.PDVS2delay.MaxLength = 8
        Me.PDVS2delay.Name = "PDVS2delay"
        Me.PDVS2delay.Size = New System.Drawing.Size(63, 22)
        Me.PDVS2delay.TabIndex = 553
        Me.PDVS2delay.Text = "125"
        '
        'CalStep
        '
        Me.CalStep.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalStep.Location = New System.Drawing.Point(286, 71)
        Me.CalStep.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CalStep.MaxLength = 8
        Me.CalStep.Name = "CalStep"
        Me.CalStep.Size = New System.Drawing.Size(63, 22)
        Me.CalStep.TabIndex = 543
        Me.CalStep.Text = "5"
        '
        'Label115
        '
        Me.Label115.AutoSize = True
        Me.Label115.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label115.Location = New System.Drawing.Point(352, 75)
        Me.Label115.Name = "Label115"
        Me.Label115.Size = New System.Drawing.Size(118, 16)
        Me.Label115.TabIndex = 545
        Me.Label115.Text = "Init. Step Size (bits)"
        '
        'CalAccuracy
        '
        Me.CalAccuracy.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalAccuracy.Location = New System.Drawing.Point(286, 101)
        Me.CalAccuracy.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CalAccuracy.MaxLength = 8
        Me.CalAccuracy.Name = "CalAccuracy"
        Me.CalAccuracy.Size = New System.Drawing.Size(63, 22)
        Me.CalAccuracy.TabIndex = 547
        Me.CalAccuracy.Text = "0.00010"
        '
        'Label119
        '
        Me.Label119.AutoSize = True
        Me.Label119.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label119.Location = New System.Drawing.Point(352, 105)
        Me.Label119.Name = "Label119"
        Me.Label119.Size = New System.Drawing.Size(85, 16)
        Me.Label119.TabIndex = 548
        Me.Label119.Text = "Init. Accuracy"
        '
        'CalStepFinal
        '
        Me.CalStepFinal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalStepFinal.Location = New System.Drawing.Point(286, 131)
        Me.CalStepFinal.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CalStepFinal.MaxLength = 8
        Me.CalStepFinal.Name = "CalStepFinal"
        Me.CalStepFinal.Size = New System.Drawing.Size(63, 22)
        Me.CalStepFinal.TabIndex = 549
        Me.CalStepFinal.Text = "1"
        '
        'Label120
        '
        Me.Label120.AutoSize = True
        Me.Label120.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label120.Location = New System.Drawing.Point(350, 135)
        Me.Label120.Name = "Label120"
        Me.Label120.Size = New System.Drawing.Size(128, 16)
        Me.Label120.TabIndex = 550
        Me.Label120.Text = "Step Size Final (bits)"
        '
        'CalAccuracyFinal
        '
        Me.CalAccuracyFinal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CalAccuracyFinal.Location = New System.Drawing.Point(286, 161)
        Me.CalAccuracyFinal.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CalAccuracyFinal.MaxLength = 8
        Me.CalAccuracyFinal.Name = "CalAccuracyFinal"
        Me.CalAccuracyFinal.Size = New System.Drawing.Size(63, 22)
        Me.CalAccuracyFinal.TabIndex = 551
        Me.CalAccuracyFinal.Text = "0.00001"
        '
        'Label117
        '
        Me.Label117.AutoSize = True
        Me.Label117.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label117.Location = New System.Drawing.Point(353, 165)
        Me.Label117.Name = "Label117"
        Me.Label117.Size = New System.Drawing.Size(95, 16)
        Me.Label117.TabIndex = 552
        Me.Label117.Text = "Final Accuracy"
        '
        'Label121
        '
        Me.Label121.AutoSize = True
        Me.Label121.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label121.Location = New System.Drawing.Point(351, 195)
        Me.Label121.Name = "Label121"
        Me.Label121.Size = New System.Drawing.Size(123, 16)
        Me.Label121.TabIndex = 554
        Me.Label121.Text = "Comms Delay (mS)"
        '
        'TabPage13
        '
        Me.TabPage13.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage13.Controls.Add(Me.GroupBox11)
        Me.TabPage13.Location = New System.Drawing.Point(4, 22)
        Me.TabPage13.Name = "TabPage13"
        Me.TabPage13.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage13.TabIndex = 13
        Me.TabPage13.Text = "Settings  "
        '
        'GroupBox11
        '
        Me.GroupBox11.Controls.Add(Me.Label314)
        Me.GroupBox11.Controls.Add(Me.Label307)
        Me.GroupBox11.Controls.Add(Me.btnRestore)
        Me.GroupBox11.Controls.Add(Me.btnBackup)
        Me.GroupBox11.Controls.Add(Me.CheckBoxEnableTooltips)
        Me.GroupBox11.Controls.Add(Me.TextBox2)
        Me.GroupBox11.Controls.Add(Me.TextBox1)
        Me.GroupBox11.Controls.Add(Me.Label308)
        Me.GroupBox11.Controls.Add(Me.Label247)
        Me.GroupBox11.Controls.Add(Me.Label10)
        Me.GroupBox11.Controls.Add(Me.TextBoxTextEditor)
        Me.GroupBox11.Controls.Add(Me.CheckBoxAllowSaveAnytime)
        Me.GroupBox11.Controls.Add(Me.Label6)
        Me.GroupBox11.Location = New System.Drawing.Point(8, 3)
        Me.GroupBox11.Name = "GroupBox11"
        Me.GroupBox11.Size = New System.Drawing.Size(1030, 589)
        Me.GroupBox11.TabIndex = 589
        Me.GroupBox11.TabStop = False
        '
        'Label314
        '
        Me.Label314.AutoSize = True
        Me.Label314.Location = New System.Drawing.Point(106, 161)
        Me.Label314.Name = "Label314"
        Me.Label314.Size = New System.Drawing.Size(303, 13)
        Me.Label314.TabIndex = 605
        Me.Label314.Text = "Import all Profiles and saved data (disconnect from device 1&&2)"
        '
        'Label307
        '
        Me.Label307.AutoSize = True
        Me.Label307.Location = New System.Drawing.Point(106, 130)
        Me.Label307.Name = "Label307"
        Me.Label307.Size = New System.Drawing.Size(394, 13)
        Me.Label307.TabIndex = 604
        Me.Label307.Text = "Export all Profiles and saved data to ProfilesData.dat (disconnect from device 1&" &
    "&2)"
        '
        'CheckBoxEnableTooltips
        '
        Me.CheckBoxEnableTooltips.AutoSize = True
        Me.CheckBoxEnableTooltips.Location = New System.Drawing.Point(11, 69)
        Me.CheckBoxEnableTooltips.Name = "CheckBoxEnableTooltips"
        Me.CheckBoxEnableTooltips.Size = New System.Drawing.Size(112, 17)
        Me.CheckBoxEnableTooltips.TabIndex = 601
        Me.CheckBoxEnableTooltips.Text = "Enable all Tooltips"
        Me.CheckBoxEnableTooltips.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox2.Location = New System.Drawing.Point(594, 113)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(251, 13)
        Me.TextBox2.TabIndex = 600
        Me.TextBox2.Text = "C:\Program Files\Notepad++\notepad++.exe"
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Location = New System.Drawing.Point(594, 96)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(251, 13)
        Me.TextBox1.TabIndex = 599
        Me.TextBox1.Text = "C:\Windows\System32\notepad.exe"
        '
        'Label308
        '
        Me.Label308.AutoSize = True
        Me.Label308.Location = New System.Drawing.Point(8, 566)
        Me.Label308.Name = "Label308"
        Me.Label308.Size = New System.Drawing.Size(240, 13)
        Me.Label308.TabIndex = 598
        Me.Label308.Text = "These settings are saved immediately on change."
        '
        'Label247
        '
        Me.Label247.AutoSize = True
        Me.Label247.Location = New System.Drawing.Point(532, 96)
        Me.Label247.Name = "Label247"
        Me.Label247.Size = New System.Drawing.Size(55, 13)
        Me.Label247.TabIndex = 596
        Me.Label247.Text = "Examples:"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(290, 96)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(225, 13)
        Me.Label10.TabIndex = 595
        Me.Label10.Text = "Text editor path for 'Edit GPIBchannels' button"
        '
        'TextBoxTextEditor
        '
        Me.TextBoxTextEditor.Location = New System.Drawing.Point(11, 93)
        Me.TextBoxTextEditor.Name = "TextBoxTextEditor"
        Me.TextBoxTextEditor.Size = New System.Drawing.Size(273, 20)
        Me.TextBoxTextEditor.TabIndex = 594
        '
        'CheckBoxAllowSaveAnytime
        '
        Me.CheckBoxAllowSaveAnytime.AutoSize = True
        Me.CheckBoxAllowSaveAnytime.Location = New System.Drawing.Point(11, 46)
        Me.CheckBoxAllowSaveAnytime.Name = "CheckBoxAllowSaveAnytime"
        Me.CheckBoxAllowSaveAnytime.Size = New System.Drawing.Size(264, 17)
        Me.CheckBoxAllowSaveAnytime.TabIndex = 593
        Me.CheckBoxAllowSaveAnytime.Text = "Allow 'Save All Profiles/Settings' button at any time"
        Me.CheckBoxAllowSaveAnytime.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(8, 14)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(78, 15)
        Me.Label6.TabIndex = 592
        Me.Label6.Text = "SETTINGS:"
        '
        'TabPage6
        '
        Me.TabPage6.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage6.Controls.Add(Me.Label316)
        Me.TabPage6.Controls.Add(Me.Label240)
        Me.TabPage6.Controls.Add(Me.Label65)
        Me.TabPage6.Controls.Add(Me.Label22)
        Me.TabPage6.Controls.Add(Me.Label239)
        Me.TabPage6.Controls.Add(Me.Label234)
        Me.TabPage6.Controls.Add(Me.URL3)
        Me.TabPage6.Controls.Add(Me.Label192)
        Me.TabPage6.Controls.Add(Me.URL2)
        Me.TabPage6.Controls.Add(Me.Label134)
        Me.TabPage6.Controls.Add(Me.PictureBox6)
        Me.TabPage6.Controls.Add(Me.Label135)
        Me.TabPage6.Controls.Add(Me.URL1)
        Me.TabPage6.Controls.Add(Me.Label105)
        Me.TabPage6.Controls.Add(Me.Label70)
        Me.TabPage6.Controls.Add(Me.Label69)
        Me.TabPage6.Controls.Add(Me.Donate1)
        Me.TabPage6.Controls.Add(Me.ButtonIanWebsite)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage6.Size = New System.Drawing.Size(1047, 599)
        Me.TabPage6.TabIndex = 5
        Me.TabPage6.Text = "About  "
        '
        'Label316
        '
        Me.Label316.AutoSize = True
        Me.Label316.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label316.Location = New System.Drawing.Point(36, 121)
        Me.Label316.Name = "Label316"
        Me.Label316.Size = New System.Drawing.Size(418, 16)
        Me.Label316.TabIndex = 572
        Me.Label316.Text = "NI-GPIB-232CT-A Serial device library mods by Florian Diederichsen."
        '
        'Label240
        '
        Me.Label240.AutoSize = True
        Me.Label240.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label240.Location = New System.Drawing.Point(36, 191)
        Me.Label240.Name = "Label240"
        Me.Label240.Size = New System.Drawing.Size(660, 48)
        Me.Label240.TabIndex = 570
        Me.Label240.Text = resources.GetString("Label240.Text")
        '
        'Label65
        '
        Me.Label65.AutoSize = True
        Me.Label65.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label65.Location = New System.Drawing.Point(142, 32)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(93, 12)
        Me.Label65.TabIndex = 568
        Me.Label65.Text = "TM     (UK/USA 2022)"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(17, 16)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(132, 31)
        Me.Label22.TabIndex = 551
        Me.Label22.Text = "WinGPIB"
        '
        'Label239
        '
        Me.Label239.AutoSize = True
        Me.Label239.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label239.Location = New System.Drawing.Point(36, 261)
        Me.Label239.Name = "Label239"
        Me.Label239.Size = New System.Drawing.Size(679, 64)
        Me.Label239.TabIndex = 569
        Me.Label239.Text = resources.GetString("Label239.Text")
        '
        'Label234
        '
        Me.Label234.AutoSize = True
        Me.Label234.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label234.Location = New System.Drawing.Point(36, 157)
        Me.Label234.Name = "Label234"
        Me.Label234.Size = New System.Drawing.Size(224, 16)
        Me.Label234.TabIndex = 567
        Me.Label234.Text = "Written in VB using MS Visual Studio."
        '
        'URL3
        '
        Me.URL3.AutoSize = True
        Me.URL3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.URL3.ForeColor = System.Drawing.Color.Blue
        Me.URL3.Location = New System.Drawing.Point(202, 508)
        Me.URL3.Name = "URL3"
        Me.URL3.Size = New System.Drawing.Size(184, 16)
        Me.URL3.TabIndex = 566
        Me.URL3.Text = "www.twitter.com/IanSJohnston"
        '
        'Label192
        '
        Me.Label192.AutoSize = True
        Me.Label192.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label192.Location = New System.Drawing.Point(32, 557)
        Me.Label192.Name = "Label192"
        Me.Label192.Size = New System.Drawing.Size(410, 16)
        Me.Label192.TabIndex = 565
        Me.Label192.Text = "If you wish to use this program commercially then please contact me."
        '
        'URL2
        '
        Me.URL2.AutoSize = True
        Me.URL2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.URL2.ForeColor = System.Drawing.Color.Blue
        Me.URL2.Location = New System.Drawing.Point(203, 489)
        Me.URL2.Name = "URL2"
        Me.URL2.Size = New System.Drawing.Size(250, 16)
        Me.URL2.TabIndex = 556
        Me.URL2.Text = "www.youtube.com/user/IanScottJohnston"
        '
        'Label134
        '
        Me.Label134.AutoSize = True
        Me.Label134.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label134.Location = New System.Drawing.Point(36, 101)
        Me.Label134.Name = "Label134"
        Me.Label134.Size = New System.Drawing.Size(265, 16)
        Me.Label134.TabIndex = 555
        Me.Label134.Text = "Prologix Serial device library mods by tppc."
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = CType(resources.GetObject("PictureBox6.Image"), System.Drawing.Image)
        Me.PictureBox6.Location = New System.Drawing.Point(43, 353)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(140, 171)
        Me.PictureBox6.TabIndex = 554
        Me.PictureBox6.TabStop = False
        '
        'Label135
        '
        Me.Label135.AutoSize = True
        Me.Label135.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label135.Location = New System.Drawing.Point(32, 538)
        Me.Label135.Name = "Label135"
        Me.Label135.Size = New System.Drawing.Size(503, 16)
        Me.Label135.TabIndex = 553
        Me.Label135.Text = "This program may be used and distributed freely for non-commercial purposes only." &
    ""
        '
        'URL1
        '
        Me.URL1.AutoSize = True
        Me.URL1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.URL1.ForeColor = System.Drawing.Color.Blue
        Me.URL1.Location = New System.Drawing.Point(203, 470)
        Me.URL1.Name = "URL1"
        Me.URL1.Size = New System.Drawing.Size(134, 16)
        Me.URL1.TabIndex = 88
        Me.URL1.Text = "www.ianjohnston.com"
        '
        'Label105
        '
        Me.Label105.AutoSize = True
        Me.Label105.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label105.Location = New System.Drawing.Point(204, 452)
        Me.Label105.Name = "Label105"
        Me.Label105.Size = New System.Drawing.Size(85, 16)
        Me.Label105.TabIndex = 87
        Me.Label105.Text = "Ian Johnston."
        '
        'Label70
        '
        Me.Label70.AutoSize = True
        Me.Label70.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label70.Location = New System.Drawing.Point(36, 81)
        Me.Label70.Name = "Label70"
        Me.Label70.Size = New System.Drawing.Size(219, 16)
        Me.Label70.TabIndex = 83
        Me.Label70.Text = "Main WinGPIB app by Ian Johnston."
        '
        'Label69
        '
        Me.Label69.AutoSize = True
        Me.Label69.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label69.Location = New System.Drawing.Point(36, 61)
        Me.Label69.Name = "Label69"
        Me.Label69.Size = New System.Drawing.Size(355, 16)
        Me.Label69.TabIndex = 82
        Me.Label69.Text = "Original GPIB device library (.DLL) code by Pawel Wzietek."
        '
        'Donate1
        '
        Me.Donate1.AutoSize = True
        Me.Donate1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Donate1.ForeColor = System.Drawing.Color.Red
        Me.Donate1.Location = New System.Drawing.Point(247, 339)
        Me.Donate1.Name = "Donate1"
        Me.Donate1.Size = New System.Drawing.Size(86, 20)
        Me.Donate1.TabIndex = 85
        Me.Donate1.Text = "DONATE!"
        '
        'Timer8
        '
        Me.Timer8.Interval = 50
        '
        'Timer9
        '
        Me.Timer9.Interval = 50
        '
        'Timer10
        '
        Me.Timer10.Interval = 50
        '
        'Timer11
        '
        Me.Timer11.Interval = 50
        '
        'Timer12
        '
        Me.Timer12.Interval = 50
        '
        'Timer13
        '
        Me.Timer13.Interval = 50
        '
        'Timer14
        '
        Me.Timer14.Interval = 50
        '
        'OnOffLed2
        '
        Me.OnOffLed2.Location = New System.Drawing.Point(219, 96)
        Me.OnOffLed2.Name = "OnOffLed2"
        Me.OnOffLed2.OffText = Nothing
        Me.OnOffLed2.OnText = Nothing
        Me.OnOffLed2.Size = New System.Drawing.Size(20, 20)
        Me.OnOffLed2.State = WinGPIBproj.OnOffLed.LedState.Off
        Me.OnOffLed2.TabIndex = 536
        '
        'OnOffLed1
        '
        Me.OnOffLed1.Location = New System.Drawing.Point(195, 96)
        Me.OnOffLed1.Name = "OnOffLed1"
        Me.OnOffLed1.OffText = Nothing
        Me.OnOffLed1.OnText = Nothing
        Me.OnOffLed1.Size = New System.Drawing.Size(20, 20)
        Me.OnOffLed1.State = WinGPIBproj.OnOffLed.LedState.Off
        Me.OnOffLed1.TabIndex = 535
        '
        'OnOffLed4
        '
        Me.OnOffLed4.Location = New System.Drawing.Point(120, 34)
        Me.OnOffLed4.Name = "OnOffLed4"
        Me.OnOffLed4.OffText = Nothing
        Me.OnOffLed4.OnText = Nothing
        Me.OnOffLed4.Size = New System.Drawing.Size(20, 20)
        Me.OnOffLed4.State = WinGPIBproj.OnOffLed.LedState.Off
        Me.OnOffLed4.TabIndex = 797
        '
        'OnOffLed3
        '
        Me.OnOffLed3.Location = New System.Drawing.Point(120, 11)
        Me.OnOffLed3.Name = "OnOffLed3"
        Me.OnOffLed3.OffText = Nothing
        Me.OnOffLed3.OnText = Nothing
        Me.OnOffLed3.Size = New System.Drawing.Size(20, 20)
        Me.OnOffLed3.State = WinGPIBproj.OnOffLed.LedState.Off
        Me.OnOffLed3.TabIndex = 796
        '
        'Formtest
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(1054, 626)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "Formtest"
        Me.Text = "WinGPIB"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox9.PerformLayout
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit
        Me.gbox12.ResumeLayout(False)
        Me.gbox12.PerformLayout
        Me.gbox2.ResumeLayout(False)
        Me.gbox2.PerformLayout
        Me.gbox1.ResumeLayout(False)
        Me.gbox1.PerformLayout
        Me.TabPage10.ResumeLayout(False)
        Me.TabPage10.PerformLayout
        Me.TabPage8.ResumeLayout(False)
        Me.TabPage8.PerformLayout
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout
        Me.TabPage2.ResumeLayout(False)
        Me.gboxtemphum.ResumeLayout(False)
        Me.gboxtemphum.PerformLayout
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout
        Me.bgoxdata.ResumeLayout(False)
        Me.bgoxdata.PerformLayout
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage7.ResumeLayout(False)
        Me.TabPage7.PerformLayout
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage11.ResumeLayout(False)
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox10.PerformLayout
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage12.ResumeLayout(False)
        Me.TabPage12.PerformLayout
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.PictureBox9, System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout
        Me.TabPage13.ResumeLayout(False)
        Me.GroupBox11.ResumeLayout(False)
        Me.GroupBox11.PerformLayout
        Me.TabPage6.ResumeLayout(False)
        Me.TabPage6.PerformLayout
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SerialPort As IO.Ports.SerialPort
    Friend WithEvents Timer1 As Timer
    Friend WithEvents Timer2 As Timer
    Friend WithEvents Timer3 As Timer
    Friend WithEvents Timer4 As Timer
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents SerialPort1 As IO.Ports.SerialPort
    Friend WithEvents Timer6 As Timer
    Friend WithEvents Timer7 As Timer
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents btndevlist As Button
    Friend WithEvents ButtonReset As Button
    Friend WithEvents gbox12 As GroupBox
    Friend WithEvents Label67 As Label
    Friend WithEvents ButtonDev12Run As Button
    Friend WithEvents Label66 As Label
    Friend WithEvents Dev12SampleRate As TextBox
    Friend WithEvents gbox2 As GroupBox
    Friend WithEvents Div1000Dev2 As CheckBox
    Friend WithEvents Dev2removeletters As CheckBox
    Friend WithEvents CheckBoxSendBlockingDev2 As CheckBox
    Friend WithEvents Label101 As Label
    Friend WithEvents Dev2PollingEnable As CheckBox
    Friend WithEvents Dev2TerminatorEnable As CheckBox
    Friend WithEvents txtr2a_disp As TextBox
    Friend WithEvents Dev2STBMask As TextBox
    Friend WithEvents btnq2b As Button
    Friend WithEvents btns2c As Button
    Friend WithEvents txtq2c As TextBox
    Friend WithEvents btnq2a As Button
    Friend WithEvents txtq2b As TextBox
    Friend WithEvents txtr2b As TextBox
    Friend WithEvents txtq2a As TextBox
    Friend WithEvents txtr2a As TextBox
    Friend WithEvents Label45 As Label
    Friend WithEvents Label54 As Label
    Friend WithEvents CommandStart2run As TextBox
    Friend WithEvents Label37 As Label
    Friend WithEvents CommandStop2 As TextBox
    Friend WithEvents Label36 As Label
    Friend WithEvents CommandStart2 As TextBox
    Friend WithEvents Dev2Samples As Label
    Friend WithEvents Label34 As Label
    Friend WithEvents IgnoreErrors2 As CheckBox
    Friend WithEvents Label30 As Label
    Friend WithEvents Dev2SampleRate As TextBox
    Friend WithEvents ButtonDev2Run As Button
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents txtr2astat As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents gbox1 As GroupBox
    Friend WithEvents Div1000Dev1 As CheckBox
    Friend WithEvents Dev1removeletters As CheckBox
    Friend WithEvents CheckBoxSendBlockingDev1 As CheckBox
    Friend WithEvents Dev1STBMask As TextBox
    Friend WithEvents Label99 As Label
    Friend WithEvents txtr1a_disp As TextBox
    Friend WithEvents btnq1b As Button
    Friend WithEvents txtq1b As TextBox
    Friend WithEvents btns1c As Button
    Friend WithEvents Dev1PollingEnable As CheckBox
    Friend WithEvents txtq1c As TextBox
    Friend WithEvents btnq1a As Button
    Friend WithEvents Dev1TerminatorEnable As CheckBox
    Friend WithEvents txtr1a As TextBox
    Friend WithEvents txtq1a As TextBox
    Friend WithEvents txtr1b As TextBox
    Friend WithEvents Label43 As Label
    Friend WithEvents CommandStart1run As TextBox
    Friend WithEvents Label53 As Label
    Friend WithEvents Label35 As Label
    Friend WithEvents Label32 As Label
    Friend WithEvents CommandStop1 As TextBox
    Friend WithEvents CommandStart1 As TextBox
    Friend WithEvents Dev1Samples As Label
    Friend WithEvents Label31 As Label
    Friend WithEvents Label33 As Label
    Friend WithEvents ButtonDev1Run As Button
    Friend WithEvents IgnoreErrors1 As CheckBox
    Friend WithEvents Dev1SampleRate As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents txtr1astat As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Label5 As Label
    'Friend WithEvents ShapeContainer2 As PowerPacks.ShapeContainer
    'Friend WithEvents RectangleShape1 As PowerPacks.RectangleShape
    'Friend WithEvents RectangleShape2 As PowerPacks.RectangleShape
    Friend WithEvents gboxtemphum As GroupBox
    Friend WithEvents Label23 As Label
    Friend WithEvents txtname3 As TextBox
    Friend WithEvents Label29 As Label
    Friend WithEvents Label28 As Label
    Friend WithEvents lstIntf3 As ComboBox
    Friend WithEvents Label21 As Label
    Friend WithEvents LabelHumidity As Label
    Friend WithEvents ButtonStart As Button
    Friend WithEvents LabelTemperature As Label
    Friend WithEvents ButtonEnd As Button
    Friend WithEvents Label20 As Label
    Friend WithEvents ComboBoxPort As ComboBox
    Friend WithEvents Label19 As Label
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents bgoxdata As GroupBox
    Friend WithEvents Label71 As Label
    Friend WithEvents CSVdelimiterSemiColon As RadioButton
    Friend WithEvents CSVdelimiterComma As RadioButton
    Friend WithEvents LabelCSVfilesize As Label
    Friend WithEvents CSVsize As Label
    Friend WithEvents Label52 As Label
    Friend WithEvents Label51 As Label
    Friend WithEvents Label50 As Label
    Friend WithEvents Label49 As Label
    Friend WithEvents Label48 As Label
    Friend WithEvents CheckboxEnableLOG As CheckBox
    Friend WithEvents ShowFiles As Button
    Friend WithEvents ResetCSV As Button
    Friend WithEvents ButtonExportCSV As Button
    Friend WithEvents ENotationDecimal As CheckBox
    Friend WithEvents CheckboxEnableCSV As CheckBox
    Friend WithEvents TempHumLogs As CheckBox
    Friend WithEvents Label27 As Label
    Friend WithEvents CSVfilename As TextBox
    Friend WithEvents Label25 As Label
    Friend WithEvents CSVfilepath As TextBox
    Friend WithEvents Label26 As Label
    Friend WithEvents ListBoxData As ListBox
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents Label72 As Label
    Friend WithEvents ButtonDiffRecordedTempReset As Button
    Friend WithEvents TemperatureDiffRecorded As Label
    Friend WithEvents ButtonDiffRecorded2Reset As Button
    Friend WithEvents EnableChart1 As CheckBox
    Friend WithEvents ButtonDiffRecorded1Reset As Button
    Friend WithEvents EnableChart3 As CheckBox
    Friend WithEvents inst_value2FDiffRecorded As Label
    Friend WithEvents inst_value1FDiffRecorded As Label
    Friend WithEvents Label44 As Label
    Friend WithEvents Label42 As Label
    Friend WithEvents EnableChart2 As CheckBox
    Friend WithEvents ButtonClearChart As Button
    Friend WithEvents Chart1 As DataVisualization.Charting.Chart
    Friend WithEvents TabPage5 As TabPage
    Friend WithEvents volts10 As TextBox
    Friend WithEvents volts9 As TextBox
    Friend WithEvents volts8 As TextBox
    Friend WithEvents volts7 As TextBox
    Friend WithEvents volts6 As TextBox
    Friend WithEvents volts5 As TextBox
    Friend WithEvents volts4 As TextBox
    Friend WithEvents volts3 As TextBox
    Friend WithEvents volts2 As TextBox
    Friend WithEvents volts1 As TextBox
    Friend WithEvents volts0 As TextBox
    Friend WithEvents Label93 As Label
    Friend WithEvents LabeldacSpan9Cal As Label
    Friend WithEvents DacSpan9 As TextBox
    Friend WithEvents Label89 As Label
    Friend WithEvents LabeldacSpan8Cal As Label
    Friend WithEvents DacSpan8 As TextBox
    Friend WithEvents Label85 As Label
    Friend WithEvents LabeldacSpan7Cal As Label
    Friend WithEvents DacSpan7 As TextBox
    Friend WithEvents Label86 As Label
    Friend WithEvents LabeldacSpan6Cal As Label
    Friend WithEvents DacSpan6 As TextBox
    Friend WithEvents Label90 As Label
    Friend WithEvents LabeldacSpan5Cal As Label
    Friend WithEvents DacSpan5 As TextBox
    Friend WithEvents Label94 As Label
    Friend WithEvents LabeldacSpan4Cal As Label
    Friend WithEvents DacSpan4 As TextBox
    Friend WithEvents Label96 As Label
    Friend WithEvents LabeldacSpan3Cal As Label
    Friend WithEvents DacSpan3 As TextBox
    Friend WithEvents Label98 As Label
    Friend WithEvents LabeldacSpan2Cal As Label
    Friend WithEvents DacSpan2 As TextBox
    Friend WithEvents Label100 As Label
    Friend WithEvents LabeldacSpan1Cal As Label
    Friend WithEvents DacSpan1 As TextBox
    Friend WithEvents LabelCalCount As Label
    Friend WithEvents LabelBatteryMonICMult As Label
    Friend WithEvents Label106 As Label
    Friend WithEvents ButtonChargeI As Button
    Friend WithEvents ChargeI As TextBox
    Friend WithEvents Label107 As Label
    Friend WithEvents Label108 As Label
    Friend WithEvents Label109 As Label
    Friend WithEvents Label110 As Label
    Friend WithEvents LabelBatteryVFeedMult As Label
    Friend WithEvents LabelOutputVFeedMult As Label
    Friend WithEvents LabeldacSpan0Cal As Label
    Friend WithEvents LabeldacZero0Cal As Label
    Friend WithEvents BattVdc As TextBox
    Friend WithEvents OutVdc As TextBox
    Friend WithEvents DacSpan0 As TextBox
    Friend WithEvents DacZero0 As TextBox
    Friend WithEvents Label57 As Label
    Friend WithEvents TextBoxdegC As TextBox
    Friend WithEvents Label97 As Label
    Friend WithEvents TextBoxSer As TextBox
    Friend WithEvents Label95 As Label
    Friend WithEvents TextBoxSOAK As TextBox
    Friend WithEvents TextBoxSERIAL As TextBox
    Friend WithEvents TextBoxFULLMA As TextBox
    Friend WithEvents TextBoxOLMA As TextBox
    Friend WithEvents TextBoxCENABLE As TextBox
    Friend WithEvents TextBoxLOWSHUT As TextBox
    Friend WithEvents TextBoxDC As TextBox
    Friend WithEvents TextBoxCD As TextBox
    Friend WithEvents Label92 As Label
    Friend WithEvents Label91 As Label
    Friend WithEvents Label88 As Label
    Friend WithEvents Label87 As Label
    Friend WithEvents Label84 As Label
    Friend WithEvents Label82 As Label
    Friend WithEvents Label78 As Label
    Friend WithEvents Label77 As Label
    Friend WithEvents exportBTN As Button
    Friend WithEvents DACcalparamVAL As Label
    Friend WithEvents PDVS2counter As Label
    Friend WithEvents Label121 As Label
    Friend WithEvents PDVS2delay As TextBox
    Friend WithEvents Label117 As Label
    Friend WithEvents CalAccuracyFinal As TextBox
    Friend WithEvents Label120 As Label
    Friend WithEvents CalStepFinal As TextBox
    Friend WithEvents Label119 As Label
    Friend WithEvents CalAccuracy As TextBox
    Friend WithEvents Label115 As Label
    Friend WithEvents CalStep As TextBox
    Friend WithEvents Label113 As Label
    Friend WithEvents Label112 As Label
    Friend WithEvents txtr1aBIG As Label
    Friend WithEvents comPort_ComboBox As ComboBox
    Friend WithEvents Label81 As Label
    Friend WithEvents connect_BTN As Button
    Friend WithEvents ButtonKeyVoltage As Button
    Friend WithEvents Label79 As Label
    Friend WithEvents KeyVoltage As TextBox
    Friend WithEvents LabelKeyVoltage As Label
    Friend WithEvents TabPage6 As TabPage
    Friend WithEvents Donate1 As Label
    Friend WithEvents Label70 As Label
    Friend WithEvents Label69 As Label
    Friend WithEvents ButtonIanWebsite As Button
    Friend WithEvents TabPage7 As TabPage
    Friend WithEvents Label38 As Label
    Friend WithEvents ButtonSaveSettings As Button
    Friend WithEvents ButtonNotePad2 As Button
    Friend WithEvents Label58 As Label
    Friend WithEvents Label61 As Label
    Friend WithEvents Label60 As Label
    Friend WithEvents Label63 As Label
    Friend WithEvents CalramStatus As Label
    Friend WithEvents Label62 As Label
    Friend WithEvents ShowFilesCalRam As Button
    Friend WithEvents TextBoxCalRamFile As TextBox
    Friend WithEvents LabelCalRamAddress As Label
    Friend WithEvents Label102 As Label
    Friend WithEvents LabelCalRamByte As Label
    Friend WithEvents LabelByte As Label
    Friend WithEvents LabelCounter As Label
    Friend WithEvents Label103 As Label
    Friend WithEvents Label68 As Label
    Friend WithEvents Label105 As Label
    Friend WithEvents URL1 As Label
    Friend WithEvents Label122 As Label
    Friend WithEvents Label123 As Label
    Friend WithEvents Label116 As Label
    Friend WithEvents ButtonCalramDump3457A As Button
    Friend WithEvents LabelCounter3457A As Label
    Friend WithEvents Label126 As Label
    Friend WithEvents LabelCalRamAddress3457A As Label
    Friend WithEvents Label128 As Label
    Friend WithEvents LabelCalRamByte3457A As Label
    Friend WithEvents Label130 As Label
    Friend WithEvents CalramStatus3457A As Label
    Friend WithEvents Label132 As Label
    Friend WithEvents TextBoxCalRamFile3457A As TextBox
    Friend WithEvents Label125 As Label
    Friend WithEvents AddressRangeB As RadioButton
    Friend WithEvents AddressRangeA As RadioButton
    Friend WithEvents Label127 As Label
    Friend WithEvents Label131 As Label
    Friend WithEvents Label129 As Label
    Friend WithEvents GroupBox6 As GroupBox
    Friend WithEvents GroupBox7 As GroupBox
    Friend WithEvents PictureBox3 As PictureBox
    Friend WithEvents PictureBox4 As PictureBox
    Friend WithEvents Label22 As Label
    Friend WithEvents Label135 As Label
    Friend WithEvents PictureBox5 As PictureBox
    Friend WithEvents PictureBox6 As PictureBox
    Friend WithEvents Label56 As Label
    Friend WithEvents ShowFiles2 As Button
    Friend WithEvents Label134 As Label
    Friend WithEvents URL2 As Label
    Friend WithEvents XaxisPoints As TextBox
    Friend WithEvents Dev1Max As TextBox
    Friend WithEvents Dev1Min As TextBox
    Friend WithEvents Label39 As Label
    Friend WithEvents Label40 As Label
    Friend WithEvents Label41 As Label
    Friend WithEvents AddressRangeD As RadioButton
    Friend WithEvents AddressRangeC As RadioButton
    Friend WithEvents Label139 As Label
    Friend WithEvents Label138 As Label
    Friend WithEvents Label64 As Label
    Friend WithEvents Label141 As Label
    Friend WithEvents Label140 As Label
    Friend WithEvents Button3458Aabort As Button
    Friend WithEvents Button3457Aabort As Button
    Friend WithEvents Label142 As Label
    Friend WithEvents Label143 As Label
    Friend WithEvents Label144 As Label
    Friend WithEvents TextBox3457ATo As TextBox
    Friend WithEvents Label147 As Label
    Friend WithEvents Label146 As Label
    Friend WithEvents TextBox3457AFrom As TextBox
    Friend WithEvents AddressRangeF As RadioButton
    Friend WithEvents Label148 As Label
    Friend WithEvents Label151 As Label
    Friend WithEvents LabelDeltaV As Label
    Friend WithEvents Label152 As Label
    Friend WithEvents Label153 As Label
    Friend WithEvents Label154 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Button020000 As Button
    Friend WithEvents Button010000 As Button
    Friend WithEvents Button001000 As Button
    Friend WithEvents Button000010 As Button
    Friend WithEvents volts000010 As TextBox
    Friend WithEvents Button000100 As Button
    Friend WithEvents ButtonSetPrecalvars As Button
    Friend WithEvents nplc4_BTN As Button
    Friend WithEvents nplc3_BTN As Button
    Friend WithEvents nplc2_BTN As Button
    Friend WithEvents getcal_BTN As Button
    Friend WithEvents nplc_BTN As Button
    Friend WithEvents NPLC As Label
    Friend WithEvents Abort_BTN As Button
    Friend WithEvents precal_BTN As Button
    Friend WithEvents SavePDVS2Eprom As Button
    Friend WithEvents Label83 As Label
    Friend WithEvents Label76 As Label
    Friend WithEvents Label75 As Label
    Friend WithEvents Label74 As Label
    Friend WithEvents Button050000 As Button
    Friend WithEvents volts050000 As TextBox
    Friend WithEvents volts020000 As TextBox
    Friend WithEvents volts010000 As TextBox
    Friend WithEvents volts001000 As TextBox
    Friend WithEvents volts000100 As TextBox
    Friend WithEvents volts030000 As TextBox
    Friend WithEvents Button030000 As Button
    Friend WithEvents LabeldacZero0Delta As Label
    Friend WithEvents LabeldacSpan9Delta As Label
    Friend WithEvents LabeldacSpan8Delta As Label
    Friend WithEvents LabeldacSpan7Delta As Label
    Friend WithEvents LabeldacSpan6Delta As Label
    Friend WithEvents LabeldacSpan5Delta As Label
    Friend WithEvents LabeldacSpan4Delta As Label
    Friend WithEvents LabeldacSpan3Delta As Label
    Friend WithEvents LabeldacSpan2Delta As Label
    Friend WithEvents LabeldacSpan1Delta As Label
    Friend WithEvents LabeldacSpan0Delta As Label
    Friend WithEvents ButtonSetXYcalvars As Button
    Friend WithEvents Label162 As Label
    Friend WithEvents Label163 As Label
    Friend WithEvents Label164 As Label
    Friend WithEvents TempOffset As TextBox
    Friend WithEvents Label118 As Label
    Friend WithEvents TextBoxLabelBatteryCharge As TextBox
    Friend WithEvents TextBoxLabelBatteryLowInd As TextBox
    Friend WithEvents TextBoxLabelMode As TextBox
    Friend WithEvents TextBoxLabelBatteryI As TextBox
    Friend WithEvents TextBoxLabelBatteryV As TextBox
    Friend WithEvents TextBoxLabelOutputVFeedback As TextBox
    Friend WithEvents Label167 As Label
    Friend WithEvents Label166 As Label
    Friend WithEvents Label165 As Label
    Friend WithEvents Dev23457Aseven As CheckBox
    Friend WithEvents Dev13457Aseven As CheckBox
    Friend WithEvents TabPage8 As TabPage
    Friend WithEvents Label170 As Label
    Friend WithEvents Dev2Meter As Label
    Friend WithEvents Label169 As Label
    Friend WithEvents Dev1Meter As Label
    Friend WithEvents Device2name As Label
    Friend WithEvents Device1name As Label
    Friend WithEvents DeviceHumidity As Label
    Friend WithEvents Label173 As Label
    Friend WithEvents DeviceTemperature As Label
    Friend WithEvents Label171 As Label
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents ShowFiles3 As Button
    Friend WithEvents LabelTemperature2 As Label
    Friend WithEvents Dev2TerminatorEnable2 As CheckBox
    Friend WithEvents Dev1TerminatorEnable2 As CheckBox
    Friend WithEvents TabPage9 As TabPage
    Friend WithEvents Label179 As Label
    Friend WithEvents Label180 As Label
    Friend WithEvents LCTempMax As TextBox
    Friend WithEvents LCTempMin As TextBox
    Friend WithEvents EditMode As CheckBox
    Friend WithEvents Label175 As Label
    Friend WithEvents CSVEntryLimit As TextBox
    Friend WithEvents CheckboxCSVlimit As CheckBox
    Friend WithEvents Label176 As Label
    Friend WithEvents CSVEntryLimitMins As TextBox
    Friend WithEvents CheckboxCSVlimitMins As CheckBox
    Friend WithEvents Label177 As Label
    Friend WithEvents TextFilenameAppend As TextBox
    Friend WithEvents Label183 As Label
    Friend WithEvents Label182 As Label
    Friend WithEvents Label181 As Label
    Friend WithEvents Label178 As Label
    Friend WithEvents Label186 As Label
    Friend WithEvents Label185 As Label
    Friend WithEvents Label184 As Label
    Friend WithEvents TabPage10 As TabPage
    Friend WithEvents CheckBoxDev1Async As CheckBox
    Friend WithEvents CheckBoxDev1Query As CheckBox
    Friend WithEvents Label189 As Label
    Friend WithEvents CheckBoxDev2Async As CheckBox
    Friend WithEvents CheckBoxDev2Query As CheckBox
    Friend WithEvents Label188 As Label
    Friend WithEvents Label187 As Label
    Friend WithEvents CMDdev2 As Label
    Friend WithEvents CMDdev1 As Label
    Friend WithEvents CMD2clear As Button
    Friend WithEvents CMD1clear As Button
    Friend WithEvents TextBoxDev2CMD As TextBox
    Friend WithEvents TextBoxDev1CMD As TextBox
    Friend WithEvents Label191 As Label
    Friend WithEvents Label190 As Label
    Friend WithEvents Dev1K2001isolatedata As CheckBox
    Friend WithEvents Dev2K2001isolatedata As CheckBox
    Friend WithEvents Dev2K2001isolatedataCHAR As TextBox
    Friend WithEvents Dev1K2001isolatedataCHAR As TextBox
    Friend WithEvents Mult1000Dev2 As CheckBox
    Friend WithEvents Mult1000Dev1 As CheckBox
    Friend WithEvents Label192 As Label
    Friend WithEvents CalOnExisting As Button
    Friend WithEvents Label249 As Label
    Friend WithEvents ButtonCalramDump3458A As Button
    Friend WithEvents TextBoxCalRamFile2 As TextBox
    Friend WithEvents LabelCalRamAddressHex As Label
    Friend WithEvents LabelCalRamAddress3457AHex As Label
    Friend WithEvents Label251 As Label
    Friend WithEvents Dev2Timeout As TextBox
    Friend WithEvents Label250 As Label
    Friend WithEvents Dev1Timeout As TextBox
    Friend WithEvents Label253 As Label
    Friend WithEvents Dev2delayop As TextBox
    Friend WithEvents Label252 As Label
    Friend WithEvents Dev1delayop As TextBox
    Friend WithEvents ButtonSetRetrievedVars As Button
    Friend WithEvents ButtonAutoSET As Button
    Friend WithEvents ButtonAutomV As Button
    Friend WithEvents Label254 As Label
    Friend WithEvents EnableAutoYChart1 As CheckBox
    Friend WithEvents YaxisDiff As Label
    Friend WithEvents Label256 As Label
    Friend WithEvents ButtonPauseChart As Button
    Friend WithEvents LabelChartPoints2 As Label
    Friend WithEvents Label257 As Label
    Friend WithEvents LabelChartPoints1 As Label
    Friend WithEvents Label258 As Label
    Friend WithEvents DisableRollingChart As CheckBox
    Friend WithEvents Device2nameLive As Label
    Friend WithEvents Device1nameLive As Label
    Friend WithEvents ButtonSaveLiveSettings As Button
    Friend WithEvents TabPage12 As TabPage
    Friend WithEvents Label259 As Label
    Friend WithEvents Label260 As Label
    Friend WithEvents Label261 As Label
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents Label3458ARDG As Label
    Friend WithEvents ButtonCal3245A As Button
    Friend WithEvents Label265 As Label
    Friend WithEvents Label268 As Label
    Friend WithEvents RadioButton3245ADCV As RadioButton
    Friend WithEvents LabelRDG As Label
    Friend WithEvents Label270 As Label
    Friend WithEvents Label271 As Label
    Friend WithEvents Label274 As Label
    Friend WithEvents Label275 As Label
    Friend WithEvents Cal3245status As Label
    Friend WithEvents PictureBox9 As PictureBox
    Friend WithEvents Label3458123 As Label
    Friend WithEvents Label262 As Label
    Friend WithEvents Code3245A As TextBox
    Friend WithEvents Label264 As Label
    Friend WithEvents Button3245Aabort As Button
    Friend WithEvents Label3245AWRI As Label
    Friend WithEvents Label266 As Label
    Friend WithEvents RadioButton3245ADCVDCI As RadioButton
    Friend WithEvents Label267 As Label
    Friend WithEvents CheckBoxAZERO As CheckBox
    Friend WithEvents PictureBox8 As PictureBox
    Friend WithEvents Label255 As Label
    Friend WithEvents Label269 As Label
    Friend WithEvents Label273 As Label
    Friend WithEvents Label272 As Label
    Friend WithEvents URL3 As Label
    Friend WithEvents Label263 As Label
    Friend WithEvents Timeout3458A As TextBox
    Friend WithEvents RadioPDF As RadioButton
    Friend WithEvents RadioMSWord As RadioButton
    Friend WithEvents ClearLOGdisp As Button
    Friend WithEvents Label279 As Label
    Friend WithEvents Label278 As Label
    Friend WithEvents Label277 As Label
    Friend WithEvents Label161 As Label
    Friend WithEvents Label160 As Label
    Friend WithEvents Label159 As Label
    Friend WithEvents Label137 As Label
    Friend WithEvents Label136 As Label
    Friend WithEvents Label124 As Label
    Friend WithEvents Label104 As Label
    Friend WithEvents Label55 As Label
    Friend WithEvents ButtonSaveDefs As Button
    Friend WithEvents Default10 As TextBox
    Friend WithEvents Default9 As TextBox
    Friend WithEvents Default8 As TextBox
    Friend WithEvents Default7 As TextBox
    Friend WithEvents Default6 As TextBox
    Friend WithEvents Default5 As TextBox
    Friend WithEvents Default4 As TextBox
    Friend WithEvents Default3 As TextBox
    Friend WithEvents Default2 As TextBox
    Friend WithEvents Default1 As TextBox
    Friend WithEvents Default0 As TextBox
    Friend WithEvents ButtonLoadDefs As Button
    Friend WithEvents Label280 As Label
    Friend WithEvents Label281 As Label
    Friend WithEvents Label288 As Label
    Friend WithEvents Label287 As Label
    Friend WithEvents Label286 As Label
    Friend WithEvents Label285 As Label
    Friend WithEvents Label284 As Label
    Friend WithEvents Label283 As Label
    Friend WithEvents Label282 As Label
    Friend WithEvents Label289 As Label
    Friend WithEvents Label290 As Label
    Friend WithEvents Label292 As Label
    Friend WithEvents Dev1runStopwatchEveryInMins As TextBox
    Friend WithEvents txtq1d As TextBox
    Friend WithEvents Dev1INT As Label
    Friend WithEvents Label296 As Label
    Friend WithEvents Dev1pauseDurationInSeconds As TextBox
    Friend WithEvents Label293 As Label
    Friend WithEvents Dev2pauseDurationInSeconds As TextBox
    Friend WithEvents Label297 As Label
    Friend WithEvents Dev2runStopwatchEveryInMins As TextBox
    Friend WithEvents txtq2d As TextBox
    Friend WithEvents Dev2INT As Label
    Friend WithEvents Label299 As Label
    Friend WithEvents Label300 As Label
    Friend WithEvents Label294 As Label
    Friend WithEvents Label295 As Label
    Friend WithEvents Dev2IntEnable As CheckBox
    Friend WithEvents Dev1IntEnable As CheckBox
    Friend WithEvents LabelCSVcounts As Label
    Friend WithEvents CSVcounts As Label
    Friend WithEvents Label303 As Label
    Friend WithEvents Label302 As Label
    Friend WithEvents RunningTimeLogging As Label
    Friend WithEvents Timer8 As Timer
    Friend WithEvents ButtonRefreshPorts As Button
    Friend WithEvents TextBoxResult As TextBox
    Friend WithEvents TextBoxProtocolInput As TextBox
    Friend WithEvents Label194 As Label
    Friend WithEvents Label193 As Label
    Friend WithEvents TempFinalValue As Label
    Friend WithEvents TextBoxFinalTempValue As TextBox
    Friend WithEvents Label197 As Label
    Friend WithEvents Label196 As Label
    Friend WithEvents TextBoxParseRight As TextBox
    Friend WithEvents TextBoxParseLeft As TextBox
    Friend WithEvents Label200 As Label
    Friend WithEvents Label199 As Label
    Friend WithEvents Label201 As Label
    Friend WithEvents TextBoxRegex As TextBox
    Friend WithEvents Label202 As Label
    Friend WithEvents Label203 As Label
    Friend WithEvents TextBoxTempArithmentic As TextBox
    Friend WithEvents TextBoxHumUnits As TextBox
    Friend WithEvents TextBoxTempUnits As TextBox
    Friend WithEvents Label46 As Label
    Friend WithEvents Label47 As Label
    Friend WithEvents Label207 As Label
    Friend WithEvents TextBoxSerialPortHand As TextBox
    Friend WithEvents Label206 As Label
    Friend WithEvents TextBoxSerialPortStop As TextBox
    Friend WithEvents Label205 As Label
    Friend WithEvents TextBoxSerialPortParity As TextBox
    Friend WithEvents Label204 As Label
    Friend WithEvents TextBoxSerialPortBits As TextBox
    Friend WithEvents Label195 As Label
    Friend WithEvents TextBoxSerialPortBaud As TextBox
    Friend WithEvents Label208 As Label
    Friend WithEvents Label211 As Label
    Friend WithEvents Label210 As Label
    Friend WithEvents Label209 As Label
    Friend WithEvents Label212 As Label
    Friend WithEvents Timer9 As Timer
    Friend WithEvents OnOffLed1 As OnOffLed
    Friend WithEvents Timer10 As Timer
    Friend WithEvents OnOffLed2 As OnOffLed
    Friend WithEvents Label214 As Label
    Friend WithEvents Label213 As Label
    Friend WithEvents Label219 As Label
    Friend WithEvents Label218 As Label
    Friend WithEvents Label217 As Label
    Friend WithEvents Label216 As Label
    Friend WithEvents Label215 As Label
    Friend WithEvents ButtonSaveTempHumSettings As Button
    Friend WithEvents CheckBoxParseLeftRight As CheckBox
    Friend WithEvents CheckBoxArithmetic As CheckBox
    Friend WithEvents CheckBoxRegex As CheckBox
    Friend WithEvents Label198 As Label
    Friend WithEvents Dev1Regex As CheckBox
    Friend WithEvents Dev2Regex As CheckBox
    Friend WithEvents Label220 As Label
    Friend WithEvents Dev1DecimalNumDPs As TextBox
    Friend WithEvents Dev2DecimalNumDPs As TextBox
    Friend WithEvents Label221 As Label
    Friend WithEvents LabeChartMinutes As Label
    Friend WithEvents Label223 As Label
    Friend WithEvents CheckBoxDevice1Hide As CheckBox
    Friend WithEvents CheckBoxDevice2Hide As CheckBox
    Friend WithEvents CheckBoxTempHide As CheckBox
    Friend WithEvents TextBox3458Asn As TextBox
    Friend WithEvents Label73 As Label
    Friend WithEvents TextBoxUser As TextBox
    Friend WithEvents Label222 As Label
    Friend WithEvents PDVS2miniSave As Button
    Friend WithEvents Label225 As Label
    Friend WithEvents Label224 As Label
    Friend WithEvents Dev1TextResponse As CheckBox
    Friend WithEvents Dev2TextResponse As CheckBox
    Friend WithEvents Label226 As Label
    Friend WithEvents Label227 As Label
    Friend WithEvents Dev2INTb As Label
    Friend WithEvents Dev1INTb As Label
    Friend WithEvents ButtonDev1PreRun As Button
    Friend WithEvents ButtonDev2PreRun As Button
    Friend WithEvents Dev2SendQuery As CheckBox
    Friend WithEvents Dev1SendQuery As CheckBox
    Friend WithEvents Timer11 As Timer
    Friend WithEvents Timer12 As Timer
    Friend WithEvents CSVwrite As Label
    Friend WithEvents Label228 As Label
    Friend WithEvents ClearEventLOG As Button
    Friend WithEvents Label291 As Label
    Friend WithEvents ListLog As ListBox
    Friend WithEvents LogFileMetadata As TextBox
    Friend WithEvents Label24 As Label
    Friend WithEvents Dev1Units As TextBox
    Friend WithEvents Dev2Units As TextBox
    Friend WithEvents StartChartMessage As Label
    Friend WithEvents Timer13 As Timer
    Friend WithEvents OnOffLed3 As OnOffLed
    Friend WithEvents OnOffLed4 As OnOffLed
    Friend WithEvents Timer14 As Timer
    Friend WithEvents Label232 As Label
    Friend WithEvents Label233 As Label
    Friend WithEvents ButtonBattVdc As Button
    Friend WithEvents ButtonDacSpan9 As Button
    Friend WithEvents ButtonDacSpan8 As Button
    Friend WithEvents ButtonDacSpan7 As Button
    Friend WithEvents ButtonDacSpan6 As Button
    Friend WithEvents ButtonDacSpan5 As Button
    Friend WithEvents ButtonDacSpan4 As Button
    Friend WithEvents ButtonDacSpan3 As Button
    Friend WithEvents ButtonDacSpan2 As Button
    Friend WithEvents ButtonDacSpan1 As Button
    Friend WithEvents ButtonOutVdc As Button
    Friend WithEvents ButtonDacSpan0 As Button
    Friend WithEvents ButtonDacZero0 As Button
    Friend WithEvents ButtonDacSpan9Up As Button
    Friend WithEvents ButtonDacSpan9down As Button
    Friend WithEvents ButtonDacSpan8Up As Button
    Friend WithEvents ButtonDacSpan8down As Button
    Friend WithEvents ButtonDacSpan7Up As Button
    Friend WithEvents ButtonDacSpan7down As Button
    Friend WithEvents ButtonDacSpan6Up As Button
    Friend WithEvents ButtonDacSpan6down As Button
    Friend WithEvents ButtonDacSpan5Up As Button
    Friend WithEvents ButtonDacSpan5down As Button
    Friend WithEvents ButtonDacSpan4Up As Button
    Friend WithEvents ButtonDacSpan4down As Button
    Friend WithEvents ButtonDacSpan3Up As Button
    Friend WithEvents ButtonDacSpan3down As Button
    Friend WithEvents ButtonDacSpan2Up As Button
    Friend WithEvents ButtonDacSpan2down As Button
    Friend WithEvents ButtonDacSpan1Up As Button
    Friend WithEvents ButtonDacSpan1down As Button
    Friend WithEvents ButtonChargeIUp As Button
    Friend WithEvents ButtonBattVdcUp As Button
    Friend WithEvents ButtonOutVdcUp As Button
    Friend WithEvents ButtonDacSpan0Up As Button
    Friend WithEvents ButtonDacZero0Up As Button
    Friend WithEvents ButtonChargeIdown As Button
    Friend WithEvents ButtonBattVdcdown As Button
    Friend WithEvents ButtonOutVdcdown As Button
    Friend WithEvents ButtonDacSpan0down As Button
    Friend WithEvents ButtonDacZero0down As Button
    Friend WithEvents Label114 As Label
    Friend WithEvents Label80 As Label
    Friend WithEvents ButtonRefreshPorts1 As Button
    Friend WithEvents Label145 As Label
    Friend WithEvents WryTech As CheckBox
    Friend WithEvents ButtonDacSpan10Up As Button
    Friend WithEvents ButtonDacSpan10 As Button
    Friend WithEvents ButtonDacSpan10down As Button
    Friend WithEvents LabeldacSpan10Cal As Label
    Friend WithEvents volts11 As TextBox
    Friend WithEvents DacSpan10 As TextBox
    Friend WithEvents Label150 As Label
    Friend WithEvents LabeldacSpan10Delta As Label
    Friend WithEvents Label149 As Label
    Friend WithEvents Default11 As TextBox
    Friend WithEvents Label229 As Label
    Friend WithEvents LabelTemperature3 As TextBox
    Friend WithEvents Label155 As Label
    Friend WithEvents Label156 As Label
    Friend WithEvents Label157 As Label
    Friend WithEvents Label168 As Label
    Friend WithEvents Label158 As Label
    Friend WithEvents Label172 As Label
    Friend WithEvents Label174 As Label
    Friend WithEvents GroupBox8 As GroupBox
    Friend WithEvents ProfDev1_8 As CheckBox
    Friend WithEvents ProfDev1_7 As CheckBox
    Friend WithEvents ProfDev1_6 As CheckBox
    Friend WithEvents ProfDev1_5 As CheckBox
    Friend WithEvents ProfDev1_4 As CheckBox
    Friend WithEvents IODeviceLabel1 As Label
    Friend WithEvents btncreate2 As Button
    Friend WithEvents Label16 As Label
    Friend WithEvents ProfDev1_3 As CheckBox
    Friend WithEvents ProfDev1_2 As CheckBox
    Friend WithEvents ProfDev1_1 As CheckBox
    Friend WithEvents Label17 As Label
    Friend WithEvents lstIntf1 As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtaddr1 As TextBox
    Friend WithEvents txtname1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox9 As GroupBox
    Friend WithEvents ProfDev2_8 As CheckBox
    Friend WithEvents ProfDev2_7 As CheckBox
    Friend WithEvents ProfDev2_6 As CheckBox
    Friend WithEvents ProfDev2_5 As CheckBox
    Friend WithEvents ProfDev2_4 As CheckBox
    Friend WithEvents btncreate3 As Button
    Friend WithEvents Label15 As Label
    Friend WithEvents ProfDev2_3 As CheckBox
    Friend WithEvents ProfDev2_2 As CheckBox
    Friend WithEvents ProfDev2_1 As CheckBox
    Friend WithEvents Label18 As Label
    Friend WithEvents IODeviceLabel2 As Label
    Friend WithEvents lstIntf2 As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txtname2 As TextBox
    Friend WithEvents txtaddr2 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label133 As Label
    Friend WithEvents noEOI As CheckBox
    Friend WithEvents btncreate As Button
    Friend WithEvents TextBoxTempHumSample As TextBox
    Friend WithEvents Label230 As Label
    Friend WithEvents Label231 As Label
    Friend WithEvents Label234 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label238 As Label
    Friend WithEvents Label237 As Label
    Friend WithEvents Label236 As Label
    Friend WithEvents Label235 As Label
    Friend WithEvents Label65 As Label
    Friend WithEvents Label239 As Label
    Friend WithEvents Label240 As Label
    Friend WithEvents TabPage11 As TabPage
    Friend WithEvents GroupBox10 As GroupBox
    Friend WithEvents TextBoxCalRamFile6581 As TextBox
    Friend WithEvents Label312 As Label
    Friend WithEvents Label309 As Label
    Friend WithEvents Label310 As Label
    Friend WithEvents Label311 As Label
    Friend WithEvents ButtonCalramDumpR6581 As Button
    Friend WithEvents ButtonR6581abort As Button
    Friend WithEvents Label245 As Label
    Friend WithEvents Label246 As Label
    Friend WithEvents AllRegularConstantsReadR6581 As RadioButton
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label298 As Label
    Friend WithEvents Label304 As Label
    Friend WithEvents Label306 As Label
    Friend WithEvents LabelCalRamByte6581 As Label
    Friend WithEvents CalramStatus6581 As Label
    Friend WithEvents ShowFilesCalRamR6581 As Button
    Friend WithEvents ButtonOpenR6581file As Button
    Friend WithEvents ButtonOpenR6581fileJson As Button
    Friend WithEvents TextBoxCalRamFileJson6581 As TextBox
    Friend WithEvents CheckBoxR6581RetrieveREF As CheckBox
    Friend WithEvents Label59 As Label
    Friend WithEvents Label241 As Label
    Friend WithEvents Label242 As Label
    Friend WithEvents SendRegularConstantsReadR6581 As RadioButton
    Friend WithEvents CheckBoxR6581Upload1 As CheckBox
    Friend WithEvents CheckBoxR6581Upload5 As CheckBox
    Friend WithEvents CheckBoxR6581Upload6 As CheckBox
    Friend WithEvents CheckBoxR6581Upload7 As CheckBox
    Friend WithEvents CheckBoxR6581Upload2 As CheckBox
    Friend WithEvents CheckBoxR6581Upload3 As CheckBox
    Friend WithEvents CheckBoxR6581Upload4 As CheckBox
    Friend WithEvents ButtonOpenR6581fileSelectJson As Button
    Friend WithEvents TextBoxCalRamFileJson6581Select As TextBox
    Friend WithEvents ButtonR6581commitEEprom As Button
    Friend WithEvents ButtonR6581upload As Button
    Friend WithEvents Label243 As Label
    Friend WithEvents Label244 As Label
    Friend WithEvents Label248 As Label
    Friend WithEvents Label301 As Label
    Friend WithEvents LabelCalRamByte6581upload As Label
    Friend WithEvents CalramStatus6581upload As Label
    Friend WithEvents Label305 As Label
    Friend WithEvents CheckBoxR6581Upload9 As CheckBox
    Friend WithEvents CheckBoxR6581Upload8 As CheckBox
    Friend WithEvents TextBoxR6581GPIBlist As TextBox
    Friend WithEvents Label111 As Label
    Friend WithEvents ButtonJsonViewer As Button
    Friend WithEvents ButtonAvailableComPorts As Button
    Friend WithEvents ButtonJsonViewer2 As Button
    Friend WithEvents TabPage13 As TabPage
    Friend WithEvents GroupBox11 As GroupBox
    Friend WithEvents CheckBoxAllowSaveAnytime As CheckBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents TextBoxTextEditor As TextBox
    Friend WithEvents Label247 As Label
    Friend WithEvents Label308 As Label
    Friend WithEvents Label313 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents CheckBoxEnableTooltips As CheckBox
    Friend WithEvents Label314 As Label
    Friend WithEvents Label307 As Label
    Friend WithEvents btnRestore As Button
    Friend WithEvents btnBackup As Button
    Friend WithEvents Label316 As Label
    Friend WithEvents CalRam3458APreRun As TextBox
    Friend WithEvents Label315 As Label
    Friend WithEvents Label318 As Label
    Friend WithEvents BtnSave3458A As Button
End Class
