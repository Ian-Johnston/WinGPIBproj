' Multithreaded-communication-for-GPIB-Visa-Serial
' Base GPIB device code by Pawel Wzietek and modified/extended by Ian Johnston
'
' Pawel's original thread - www.codeproject.com/Articles/1166996/Multithreaded-communication-for-GPIB-Visa-Serial-i
' Ian's thread - www.eevblog.com/forum/metrology/3458a-logging-via-windows-app-revisited/
' Prologix code by Marco
'
' Disclaimer from Ian:
' I am not a VB programmer, heck I'm not even a programmer!!!......but I usually manage to hack things together and that includes writing apps for Windows.
' You'll see some 'methods' in this source which you may laugh at or say "OMG".......well, all I can say is that despite the iffy programming, this Windows app WORKS!!!
' Ian Johnston
'
' Main form AutoScaleMode was originally set to FONT, now set to NONE
'
' Example console
' Console.WriteLine("Rolling Average Value for Device " & PPMdevice & " - " & tempcounter & " - " & PPMdegCrollingAverageValue)
'


'Imports System.Threading
Imports System.IO
'Imports System
Imports System.IO.Ports
'Imports System.Drawing
Imports System.Management
Imports System.Net
'Imports System.Xml.Serialization
'Imports System.Configuration
'Imports System.Text.RegularExpressions
'Imports System.Configuration
'Imports WinGPIBproj.Formtest
Imports System.Text
Imports System.Text.RegularExpressions
Imports IODevices
'Imports System.Runtime.InteropServices


Public Class Formtest

    'Inherits Form

    Dim strParts() As String
    Dim intPos As Integer
    Dim strData As String = ""
    Dim comPORT As String
    Dim receivedData As String = ""
    Dim KeyV As Double = 0
    Dim BattV As Double = 0
    Dim OutV As Double = 0
    Dim alternate As Boolean = 0
    Dim CommsStatus As Boolean = 0
    Dim VrefStatus As Boolean = 0
    Dim number As Integer
    Dim GPIBactive As Boolean = 0
    Dim GPIBPDVS2index As Integer = 0

    Public dev1 As IODevice
    Public dev2 As IODevice

    Dim gWithHumi As Boolean = False
    Dim gCurrHumi As Double = 0
    Dim gCurrTemp As Double = 0
    Dim ThisMoment As Date

    Dim gChartTemp As Array = Array.CreateInstance(GetType(Double), 500)  ' chart1

    Dim gPortList As String() = SerialPort.GetPortNames

    'Dim gPortListPDVS2mini As String() = SerialPort1.GetPortNames      ' BC42025 warning
    Dim gPortListPDVS2mini As String() = SerialPort.GetPortNames()


    Dim ExpandStatus As Boolean = False ' form width toggle var

    Dim CSVdelimit As String = My.Settings.data29

    Dim GPIBdone As Boolean = False

    ' Os version
    Dim osVer As Version = Environment.OSVersion.Version    ' Operating system

    ' Get system Documents folder and OneDrive folder locations
    'Dim strPath As String = "C:\Users\" & String.Format("{0}", Environment.UserName) & "\Documents\WinGPIBdata"    ' users data folder.
    'Dim strPathOD As String = "C:\Users\" & String.Format("{0}", Environment.UserName) & "\OneDrive\Documents\WinGPIBdata"    ' users data folder on OneDrive.
    Dim documentsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Dim strPath As String = Path.Combine(documentsPath, "WinGPIBdata")                          ' users data folder.
    Dim userProfilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
    Dim strPathOD As String = Path.Combine(userProfilePath, "OneDrive\Documents\WinGPIBdata")   ' users data folder on OneDrive.

    Dim inst_value2_3457A_1 As Double
    Dim inst_value2_3457A_2 As Double
    Dim inst_value2_3457A_sum As Double
    Dim temp3457A_Dev2 As String
    Dim inst_value1_3457A_1 As Double
    Dim inst_value1_3457A_2 As Double
    Dim inst_value1_3457A_sum As Double
    Dim temp3457A_Dev1 As String
    Dim TESTcounter As Integer = 0

    Dim TabUsed As Integer = 1

    Public Property ConfigurationManager As Object
    Public Property ConfigurationUserLevel As Object

    'Dim Dev1interruptactive As Boolean = False
    'Dim Dev2interruptactive As Boolean = False
    'Dim RunChart As Boolean = False

    Dim Dev1GPIBActivity As Boolean = False
    Dim Dev2GPIBActivity As Boolean = False

    Dim BannerText1 As String
    Dim BannerText2 As String

    Dim PDVS2miniCalAvailable As Boolean = 0


    Private Sub Formtest_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        ' Tidy up
        'Me.Timer1.Stop()
        Me.Timer2.Stop()
        Me.Timer3.Stop()
        Me.Timer4.Stop()
        Me.Timer5.Stop()
        Me.Timer6.Stop()
        Me.Timer7.Stop()
        Me.Timer8.Stop()
        Me.Timer9.Stop()
        Me.Timer10.Stop()
        Me.Timer11.Stop()
        Me.Timer12.Stop()
        Me.Timer13.Stop()
        Me.Timer14.Stop()

        ' Manually run RESET I/O DEVICES button
        Dim reset As New EventArgs()
        ButtonReset_Click(Me, reset)

        'shut down gracefully: (may take time)
        IODevice.DisposeAll()

        ' Close serial ports
        If Me.SerialPort.IsOpen = True Then
            Me.SerialPort.Close()
        End If
        If Me.SerialPort1.IsOpen = True Then
            Me.SerialPort1.Close()
        End If

    End Sub

    'Sub formtest_close()
    '    IODevice.DisposeAll()
    'End Sub

    'Protected Overrides Sub OnLoad(e As EventArgs)
    'MyBase.OnLoad(e) ' Ensure this is called to trigger the Load event
    ' Additional logic
    'End Sub

    Private Sub Formtest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            ' Banner Text animation - See Timer8                                                                                                       Please DONATE if you find this app useful. See the ABOUT tab"
            BannerText1 = "WinGPIB   V4.028"
            BannerText2 = "Non-Commercial Use Only  -  Please DONATE if you find this app useful, see the ABOUT tab  -  Non-Commercial Use Only"

            ' Check for the existance of the WinGPIBdata folder at C:\Users\[username]\Documents and if it
            ' doesn't exist then copy it from C:\Users\[username]\OneDrive\Documents to C:\Users\[username]\Documents folder
            'If My.Computer.FileSystem.DirectoryExists(strPathOD) Then
            ' The OneDrive folder exists so now copy if the target folder doesn't exist
            'If Not My.Computer.FileSystem.DirectoryExists(strPath) Then
            'My.Computer.FileSystem.CopyDirectory(strPathOD, strPath, True)
            'End If
            'End If

            ' Hide the Advantest R6581 tab as it's not finished yet
            'TabControl1.TabPages.Remove(TabPage11)
            AllRegularConstantsReadR6581.Checked = True         ' default checked radio button

            ' Check for the existence of the WinGPIBdata folder in the OneDrive folder and if it exists then copy it to the documents folder
            ' This is because some install incorrectly install the WinGPIBData folder under the OneDrive structure
            If Directory.Exists(strPathOD) Then
                ' The OneDrive folder exists, so now copy if the target folder doesn't exist
                If Not Directory.Exists(strPath) Then
                    My.Computer.FileSystem.CopyDirectory(strPathOD, strPath, True)
                End If
            End If

            ' get available COM ports
            comPORT = ""
            For Each sp As String In My.Computer.Ports.SerialPortNames
                comPort_ComboBox.Items.Add(sp)
            Next

            ' Initially enable all controls on Calram groupboxes
            GroupBox6.Enabled = True
            GroupBox7.Enabled = True
            GroupBox10.Enabled = True
            GroupBox15.Enabled = True

            ' Initially enable all controls on 3245A Cal groupbox
            GroupBox5.Enabled = True

            ' Initially enable all controls on PDVS2mini groupbox
            GroupBox2.Enabled = True

            ResetCSV.Enabled = True
            TempHumLogs.Enabled = False

            lstIntf1.Items.Add("VISA")
            lstIntf1.Items.Add("GPIB:ADLink")
            lstIntf1.Items.Add("gpib488.dll")
            lstIntf1.Items.Add("Serial COM port")
            lstIntf1.Items.Add("Prologix Serial")
            lstIntf1.Items.Add("Prologix Ethernet")
            lstIntf1.Items.Add("NI-GPIB-232CT-A")
            lstIntf1.SelectedIndex = 0

            lstIntf2.Items.Add("VISA")
            lstIntf2.Items.Add("GPIB:ADLink")
            lstIntf2.Items.Add("gpib488.dll")
            lstIntf2.Items.Add("Serial COM port")
            lstIntf2.Items.Add("Prologix Serial")
            lstIntf2.Items.Add("Prologix Ethernet")
            lstIntf2.Items.Add("NI-GPIB-232CT-A")
            lstIntf2.SelectedIndex = 0

            ' Temp/Hum sensor
            lstIntf3.Items.Add("USB-TnH SHT10 V2.00")
            lstIntf3.Items.Add("USB-TnH (SHT10)")
            lstIntf3.Items.Add("USB-PA (BME280)")
            lstIntf3.Items.Add("USB-TnH (LM75)")
            lstIntf3.Items.Add("USB-TnH (SHT30)")
            lstIntf3.Items.Add("Adafruit (MCP2221A/SHT40,41,45)")
            lstIntf3.Items.Add("USB-User")
            lstIntf3.SelectedIndex = 0
            'TextBoxProtocolInput.ReadOnly = True
            'TextBoxParseLeft.ReadOnly = True
            'TextBoxParseRight.ReadOnly = True

            ' Check if the PDVS2mini stored port is still available
            If gPortListPDVS2mini.Contains(My.Settings.data333) Then
                ' The port is available, you can use it, put it in focus
                comPort_ComboBox.SelectedItem = My.Settings.data333
                comPORT = My.Settings.data333
            End If

            ' Recall all saved data
            txtname1.Text = My.Settings.data1
            txtaddr1.Text = My.Settings.data3
            CommandStart1.Text = My.Settings.data5
            CommandStop1.Text = My.Settings.data7
            Dev1SampleRate.Text = My.Settings.data9
            CSVfilename.Text = My.Settings.data11
            'CSVfilepath.Text = My.Settings.data12
            XaxisPoints.Text = My.Settings.data13
            Dev1Max.Text = My.Settings.data14
            Dev1Min.Text = My.Settings.data15
            txtname3.Text = My.Settings.data16
            CommandStart1run.Text = My.Settings.data17
            Dev12SampleRate.Text = My.Settings.data19
            CSVdelimit = My.Settings.data29

            TempOffset.Text = My.Settings.data78

            TextBoxAvgWindow.Text = My.Settings.data525

            ' Load Device 1 default profile 1
            lstIntf1.SelectedIndex = My.Settings.data33
            txtname1.Text = My.Settings.data1
            txtaddr1.Text = My.Settings.data3
            CommandStart1.Text = My.Settings.data5
            CommandStart1run.Text = My.Settings.data17
            CommandStop1.Text = My.Settings.data7
            Dev1SampleRate.Text = My.Settings.data9
            Dev1STBMask.Text = My.Settings.data66
            Div1000Dev1.Checked = My.Settings.data72
            Dev1PollingEnable.Checked = My.Settings.data36
            Dev1removeletters.Checked = My.Settings.data37
            IgnoreErrors1.Checked = My.Settings.data38
            Dev1TerminatorEnable.Checked = My.Settings.data39
            CheckBoxSendBlockingDev1.Checked = My.Settings.data40
            Dev13457Aseven.Checked = My.Settings.data79
            Dev1TerminatorEnable2.Checked = My.Settings.data85
            Dev1K2001isolatedata.Checked = My.Settings.data207
            Dev1K2001isolatedataCHAR.Text = My.Settings.data208
            Mult1000Dev1.Checked = My.Settings.data225
            Dev1Timeout.Text = My.Settings.data231
            Dev1delayop.Text = My.Settings.data243
            txtq1d.Text = My.Settings.data271
            Dev1pauseDurationInSeconds.Text = My.Settings.data272
            Dev1runStopwatchEveryInMins.Text = My.Settings.data273
            Dev1IntEnable.Checked = My.Settings.data274
            Dev1Regex.Checked = My.Settings.data337
            Dev1DecimalNumDPs.Text = My.Settings.data453
            txtOperationDev1.Text = My.Settings.data509


            ' Load Device 2 default profile 1
            lstIntf2.SelectedIndex = My.Settings.data30
            txtname2.Text = My.Settings.data2
            txtaddr2.Text = My.Settings.data4
            CommandStart2.Text = My.Settings.data6
            CommandStart2run.Text = My.Settings.data18
            CommandStop2.Text = My.Settings.data8
            Dev2SampleRate.Text = My.Settings.data10
            Dev2STBMask.Text = My.Settings.data67
            Div1000Dev2.Checked = My.Settings.data75
            Dev2PollingEnable.Checked = My.Settings.data51
            Dev2removeletters.Checked = My.Settings.data52
            IgnoreErrors2.Checked = My.Settings.data53
            Dev2TerminatorEnable.Checked = My.Settings.data54
            CheckBoxSendBlockingDev2.Checked = My.Settings.data55
            Dev23457Aseven.Checked = My.Settings.data80
            Dev2TerminatorEnable2.Checked = My.Settings.data86
            Dev2K2001isolatedata.Checked = My.Settings.data195
            Dev2K2001isolatedataCHAR.Text = My.Settings.data196
            Mult1000Dev2.Checked = My.Settings.data219
            Dev2Timeout.Text = My.Settings.data237
            Dev2delayop.Text = My.Settings.data249
            txtq2d.Text = My.Settings.data295
            Dev2pauseDurationInSeconds.Text = My.Settings.data296
            Dev2runStopwatchEveryInMins.Text = My.Settings.data297
            Dev2IntEnable.Checked = My.Settings.data298
            Dev2Regex.Checked = My.Settings.data343
            Dev2DecimalNumDPs.Text = My.Settings.data461
            txtOperationDev2.Text = My.Settings.data510

            ' Load Live Chart settings
            XaxisPoints.Text = My.Settings.data255

            ' recall with X amount of dp's and if no DP saved then recall as X.0
            Dim storedValue As Double = My.Settings.data256
            Dim formattedValue As String
            If storedValue - Math.Floor(storedValue) = 0 Then
                formattedValue = My.Settings.data256.ToString("0.0")
            Else
                formattedValue = My.Settings.data256.ToString()
            End If
            Dev1Max.Text = formattedValue

            ' recall with X amount of dp's and if no DP saved then recall as X.0
            storedValue = My.Settings.data257
            If storedValue - Math.Floor(storedValue) = 0 Then
                formattedValue = My.Settings.data257.ToString("0.0")
            Else
                formattedValue = My.Settings.data257.ToString()
            End If
            Dev1Min.Text = formattedValue

            ' recall with X amount of dp's and if no DP saved then recall as X.0
            storedValue = My.Settings.data258
            If storedValue - Math.Floor(storedValue) = 0 Then
                formattedValue = My.Settings.data258.ToString("0.0")
            Else
                formattedValue = My.Settings.data258.ToString()
            End If
            LCTempMax.Text = formattedValue

            ' recall with X amount of dp's and if no DP saved then recall as X.0
            storedValue = My.Settings.data259
            If storedValue - Math.Floor(storedValue) = 0 Then
                formattedValue = My.Settings.data259.ToString("0.0")
            Else
                formattedValue = My.Settings.data259.ToString()
            End If
            LCTempMin.Text = formattedValue

            CheckBoxDevice1Hide.Enabled = False
            CheckBoxDevice2Hide.Enabled = False
            CheckBoxTempHide.Enabled = False


            Device1nameLive.Text = ""
            Device2nameLive.Text = ""

            ' recall default PDVS2mini counts saved
            Default0.Text = My.Settings.data260
            Default1.Text = My.Settings.data261
            Default2.Text = My.Settings.data262
            Default3.Text = My.Settings.data263
            Default4.Text = My.Settings.data264
            Default5.Text = My.Settings.data265
            Default6.Text = My.Settings.data266
            Default7.Text = My.Settings.data267
            Default8.Text = My.Settings.data268
            Default9.Text = My.Settings.data269
            Default10.Text = My.Settings.data270
            Default11.Text = My.Settings.data502

            If CSVdelimit = "," Then
                CSVdelimiterComma.Checked = True
            Else
                CSVdelimiterSemiColon.Checked = True
                CSVdelimit = ";"
            End If

            ' If CSV file path is empty then set location
            If (CSVfilepath.Text = "") Then
                CSVfilepath.Text = strPath
            End If

            ' Check that the CSV file specified exists, if it doesn't then create a new blank CSV file
            ' On my Win10 laptop the install wizard pops up and it itself regenerates the Log.csv file......so looks like this sub not req'd?
            'If System.IO.File.Exists(strPath & "\" & "Log.csv") Then
            'the file exists
            'Else
            'the file doesn't exist
            'System.IO.File.Create(CSVfilepath.Text & "\" & CSVfilename.Text).Dispose()

            'Dialog2.Warning1 = "CSV file has been created:"
            'Dialog2.Warning2 = "File = " & CSVfilename.Text
            'Dialog2.Warning3 = ""
            ''Dialog2.Show() ' this method positions anywhere!
            'Dialog2.ShowDialog(Me)  ' this method positions centre of parent form, and requires to hit OK to return back to parent
            'End If


            ' Set Timer4 duration
            Me.Timer4.Interval = 100
            Me.Timer4.Start()

            ' Set Timer1 duration
            Me.Timer1.Interval = 1000
            Me.Timer1.Stop()

            ' Set Timer7 duration - PDVS2mini
            Me.Timer7.Interval = 500
            Me.Timer7.Stop()

            ' Set Timer8 duration
            Me.Timer8.Interval = 1000
            Me.Timer8.Start()

            ' Setup ComboBox of port
            gboxtemphum.Enabled = True
            Me.ComboBoxPort.Items.AddRange(gPortList)

            ' Enable logging boxe
            bgoxdata.Enabled = True

            ' 3245A button
            ButtonCal3245A.Enabled = False

            ' CalRam button 3458A & 3457A
            ButtonCalramDump3457A.Enabled = False
            ButtonCalramDump3458A.Enabled = False

            ' label style
            Chart1.ChartAreas(0).AxisY.LabelStyle.Font = New Font("Microsoft Sans Serif", 9)
            Chart1.ChartAreas(0).AxisY2.LabelStyle.Font = New Font("Microsoft Sans Serif", 9)
            Chart1.ChartAreas(0).AxisX.LabelStyle.Enabled = False
            Chart1.ChartAreas(0).AxisY.LabelStyle.Enabled = True
            ' tick marks
            Chart1.ChartAreas(0).AxisX.MajorTickMark.Enabled = True
            Chart1.ChartAreas(0).AxisY.MajorTickMark.Enabled = True
            Chart1.ChartAreas(0).AxisX.MinorTickMark.Enabled = False
            Chart1.ChartAreas(0).AxisY.MinorTickMark.Enabled = False
            ' grid
            Chart1.ChartAreas(0).AxisX.MajorGrid.Enabled = True
            Chart1.ChartAreas(0).AxisY.MajorGrid.Enabled = True
            Chart1.ChartAreas(0).AxisX.MinorGrid.Enabled = True
            Chart1.ChartAreas(0).AxisY.MinorGrid.Enabled = True
            Chart1.ChartAreas(0).AxisX.MajorGrid.LineColor = Color.FromArgb(255, 85, 85, 85)
            Chart1.ChartAreas(0).AxisY.MajorGrid.LineColor = Color.FromArgb(255, 85, 85, 85)
            Chart1.ChartAreas(0).AxisX.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)
            Chart1.ChartAreas(0).AxisY.MinorGrid.LineColor = Color.FromArgb(150, 85, 85, 85)

            Chart1.DataBindTable(gChartTemp)
            Chart1.Series(0).ChartType = 2
            Chart1.Series.Clear()
            Chart1.ChartAreas(0).BorderWidth = 1

            Chart1.Series.Add("Device 1")
            Chart1.Series.Add("Device 2")
            Chart1.Series.Add("Temperature")

            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series(2).ChartType = DataVisualization.Charting.SeriesChartType.Line

            Chart1.Series(2).YAxisType = DataVisualization.Charting.AxisType.Secondary
            Chart1.ChartAreas(0).AxisY2.Enabled = True
            Chart1.ChartAreas(0).AxisY2.Minimum = 15
            Chart1.ChartAreas(0).AxisY2.Maximum = 30
            Chart1.ChartAreas(0).AxisY2.Enabled = DataVisualization.Charting.AxisEnabled.True
            Chart1.ChartAreas(0).AxisY2.LabelStyle.Enabled = True

            Chart1.Series(0).YValueType = DataVisualization.Charting.ChartValueType.Single

            Chart1.Legends(0).Enabled = False 'set true to see channel colour labels

            Chart1.ChartAreas(0).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
            Chart1.ChartAreas(0).AxisY.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
            Chart1.ChartAreas(0).AxisY.LabelAutoFitStyle = DataVisualization.Charting.LabelAutoFitStyles.DecreaseFont 'default is staggered

            Chart1.ChartAreas(0).AxisX.IntervalOffset = 0
            Chart1.Series(0).Color = Color.Yellow
            Chart1.Series(1).Color = Color.Cyan
            Chart1.Series(2).Color = Color.Red

            Chart1.ChartAreas(0).AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot
            Chart1.ChartAreas(0).AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot
            'Chart1.ChartAreas(0).AxisY2.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

            Chart1.Series(0).YValueMembers = inst_value1FChart
            Chart1.Series(1).YValueMembers = inst_value2FChart
            Chart1.Series(2).YValueMembers = inst_value3FChart

            Chart1.ChartAreas(0).AxisX.LabelStyle.Enabled = False   'disable X-axis scale

            Chart1.Visible = False              ' hide chart on boot
            StartChartMessage.Visible = True

            IODeviceLabel1.BackColor = Color.Yellow
            IODeviceLabel2.BackColor = Color.Aqua

            'EnableChart1.BackColor = Color.Yellow
            'EnableChart2.BackColor = Color.Aqua
            'EnableChart3.BackColor = Color.Red

            TextBoxDev1CMD.Enabled = False
            TextBoxDev2CMD.Enabled = False
            CMD1clear.Enabled = False
            CMD2clear.Enabled = False
            CheckBoxDev1Query.Checked = True
            CheckBoxDev2Query.Checked = True
            TextBoxDev1CMD.BackColor = Color.DarkGray
            TextBoxDev2CMD.BackColor = Color.DarkGray
            Me.TextBoxDev1CMD.Font = New Font("System", 10)
            Me.TextBoxDev2CMD.Font = New Font("System", 10)

            ButtonRefreshPorts.Enabled = True

            TextBoxProtocolInput.Text = My.Settings.data319
            TextBoxParseLeft.Text = My.Settings.data320
            TextBoxParseRight.Text = My.Settings.data321
            TextBoxRegex.Text = My.Settings.data322
            TextBoxTempArithmentic.Text = My.Settings.data323
            TextBoxTempUnits.Text = My.Settings.data324
            TextBoxHumUnits.Text = My.Settings.data325
            lstIntf3.SelectedItem = My.Settings.data326
            TextBoxSerialPortBaud.Text = My.Settings.data327
            TextBoxSerialPortBits.Text = My.Settings.data328
            TextBoxSerialPortParity.Text = My.Settings.data329
            TextBoxSerialPortStop.Text = My.Settings.data330
            TextBoxSerialPortHand.Text = My.Settings.data331
            ComboBoxPort.SelectedItem = My.Settings.data332         ' temp/hum
            comPort_ComboBox.SelectedItem = My.Settings.data333     ' PDVS2mini
            CheckBoxParseLeftRight.Checked = My.Settings.data334    ' temp/hum checkbox
            CheckBoxRegex.Checked = My.Settings.data335             ' temp/hum checkbox
            CheckBoxArithmetic.Checked = My.Settings.data336        ' temp/hum checkbox
            TextBoxTempHumSample.Text = My.Settings.data504

            'myVariable1 = TextBoxTempUnits.Text
            'myVariable2 = TextBoxHumUnits.Text


            ' Get saved settings Dev 1
            ProfDev1_1.Checked = My.Settings.Dev1Prof1
            ProfDev1_2.Checked = My.Settings.Dev1Prof2
            ProfDev1_3.Checked = My.Settings.Dev1Prof3
            ProfDev1_4.Checked = My.Settings.Dev1Prof4
            ProfDev1_5.Checked = My.Settings.Dev1Prof5
            ProfDev1_6.Checked = My.Settings.Dev1Prof6
            ProfDev1_7.Checked = My.Settings.Dev1Prof7
            ProfDev1_8.Checked = My.Settings.Dev1Prof8

            ToolTip1.SetToolTip(ProfDev1_1, My.Settings.data1)
            ToolTip1.SetToolTip(ProfDev1_2, My.Settings.data1b)
            ToolTip1.SetToolTip(ProfDev1_3, My.Settings.data1c)
            ToolTip1.SetToolTip(ProfDev1_4, My.Settings.data139)
            ToolTip1.SetToolTip(ProfDev1_5, My.Settings.data155)
            ToolTip1.SetToolTip(ProfDev1_6, My.Settings.data171)
            ToolTip1.SetToolTip(ProfDev1_7, My.Settings.data349)
            ToolTip1.SetToolTip(ProfDev1_8, My.Settings.data375)

            ' Check to make sure that one of them is set TRUE, and also detect if more than one is set TRUE
            Dim checkboxesDev1() As CheckBox = {ProfDev1_1, ProfDev1_2, ProfDev1_3, ProfDev1_4, ProfDev1_5, ProfDev1_6, ProfDev1_7, ProfDev1_8}
            Dim checkedCountDev1 As Integer = 0
            For i As Integer = 0 To checkboxesDev1.Length - 1
                checkboxesDev1(i).Checked = My.Settings($"Dev1Prof{i + 1}")
                If checkboxesDev1(i).Checked Then
                    checkedCountDev1 += 1
                End If
            Next
            If checkedCountDev1 = 0 Then
                ' If none of the checkboxes are checked, set ProfDev1_1 to true
                ProfDev1_1.Checked = True
            ElseIf checkedCountDev1 > 1 Then
                ' If more than one checkbox is checked, set ProfDev1_1 to true and the rest to false
                ProfDev1_1.Checked = True
                For i As Integer = 1 To checkboxesDev1.Length - 1
                    checkboxesDev1(i).Checked = False
                Next
            End If

            ' Load settings per Profile selected
            If ProfDev1_1.Checked = True Then
                ProfDev1_1_Click(ProfDev1_1, EventArgs.Empty)
            End If
            If ProfDev1_2.Checked = True Then
                ProfDev1_2_Click(ProfDev1_2, EventArgs.Empty)
            End If
            If ProfDev1_3.Checked = True Then
                ProfDev1_3_Click(ProfDev1_3, EventArgs.Empty)
            End If
            If ProfDev1_4.Checked = True Then
                ProfDev1_4_Click(ProfDev1_4, EventArgs.Empty)
            End If
            If ProfDev1_5.Checked = True Then
                ProfDev1_5_Click(ProfDev1_5, EventArgs.Empty)
            End If
            If ProfDev1_6.Checked = True Then
                ProfDev1_6_Click(ProfDev1_6, EventArgs.Empty)
            End If
            If ProfDev1_7.Checked = True Then
                ProfDev1_7_Click(ProfDev1_7, EventArgs.Empty)
            End If
            If ProfDev1_8.Checked = True Then
                ProfDev1_8_Click(ProfDev1_8, EventArgs.Empty)
            End If


            ' Get saved settings Dev 2
            ProfDev2_1.Checked = My.Settings.Dev2Prof1
            ProfDev2_2.Checked = My.Settings.Dev2Prof2
            ProfDev2_3.Checked = My.Settings.Dev2Prof3
            ProfDev2_4.Checked = My.Settings.Dev2Prof4
            ProfDev2_5.Checked = My.Settings.Dev2Prof5
            ProfDev2_6.Checked = My.Settings.Dev2Prof6
            ProfDev2_7.Checked = My.Settings.Dev2Prof7
            ProfDev2_8.Checked = My.Settings.Dev2Prof8

            ToolTip1.SetToolTip(ProfDev2_1, My.Settings.data2)
            ToolTip1.SetToolTip(ProfDev2_2, My.Settings.data2b)
            ToolTip1.SetToolTip(ProfDev2_3, My.Settings.data2c)
            ToolTip1.SetToolTip(ProfDev2_4, My.Settings.data91)
            ToolTip1.SetToolTip(ProfDev2_5, My.Settings.data107)
            ToolTip1.SetToolTip(ProfDev2_6, My.Settings.data123)
            ToolTip1.SetToolTip(ProfDev2_7, My.Settings.data401)
            ToolTip1.SetToolTip(ProfDev2_8, My.Settings.data427)

            ' Check to make sure that one of them is set TRUE, and also detect if more than one is set TRUE
            Dim checkboxesDev2() As CheckBox = {ProfDev2_1, ProfDev2_2, ProfDev2_3, ProfDev2_4, ProfDev2_5, ProfDev2_6, ProfDev2_7, ProfDev2_8}
            Dim checkedCountDev2 As Integer = 0
            For i As Integer = 0 To checkboxesDev2.Length - 1
                checkboxesDev2(i).Checked = My.Settings($"Dev2Prof{i + 1}")
                If checkboxesDev2(i).Checked Then
                    checkedCountDev2 += 1
                End If
            Next
            If checkedCountDev2 = 0 Then
                ' If none of the checkboxes are checked, set ProfDev2_1 to true
                ProfDev2_1.Checked = True
            ElseIf checkedCountDev2 > 1 Then
                ' If more than one checkbox is checked, set ProfDev2_1 to true and the rest to false
                ProfDev2_1.Checked = True
                For i As Integer = 1 To checkboxesDev2.Length - 1
                    checkboxesDev2(i).Checked = False
                Next
            End If

            ' Load settings per Profile selected
            If ProfDev2_1.Checked = True Then
                ProfDev2_1_Click(ProfDev2_1, EventArgs.Empty)
            End If
            If ProfDev2_2.Checked = True Then
                ProfDev2_2_Click(ProfDev2_2, EventArgs.Empty)
            End If
            If ProfDev2_3.Checked = True Then
                ProfDev2_3_Click(ProfDev2_3, EventArgs.Empty)
            End If
            If ProfDev2_4.Checked = True Then
                ProfDev2_4_Click(ProfDev2_4, EventArgs.Empty)
            End If
            If ProfDev2_5.Checked = True Then
                ProfDev2_5_Click(ProfDev2_5, EventArgs.Empty)
            End If
            If ProfDev2_6.Checked = True Then
                ProfDev2_6_Click(ProfDev2_6, EventArgs.Empty)
            End If
            If ProfDev2_7.Checked = True Then
                ProfDev2_7_Click(ProfDev2_7, EventArgs.Empty)
            End If
            If ProfDev2_8.Checked = True Then
                ProfDev2_8_Click(ProfDev2_8, EventArgs.Empty)
            End If

            ' PDVS2mini
            TextBox3458Asn.Text = My.Settings.data469       ' 3458A serial number
            TextBoxUser.Text = My.Settings.data470          ' User/Company
            TextBoxLOWSHUT.Text = My.Settings.data471
            TextBoxCENABLE.Text = My.Settings.data472
            TextBoxOLMA.Text = My.Settings.data473
            TextBoxFULLMA.Text = My.Settings.data474
            TextBoxSERIAL.Text = My.Settings.data475
            TextBoxSOAK.Text = My.Settings.data476
            TextBoxDC.Text = My.Settings.data477
            TextBoxCD.Text = My.Settings.data478
            CalStep.Text = My.Settings.data479
            CalAccuracy.Text = My.Settings.data480
            CalStepFinal.Text = My.Settings.data481
            CalAccuracyFinal.Text = My.Settings.data482
            PDVS2delay.Text = My.Settings.data483
            TextBoxSer.Text = ""
            TextBoxdegC.Text = ""
            Dev1Units.Text = My.Settings.data500
            Dev2Units.Text = My.Settings.data501
            DisableAllButtonsInGroupBox2ExceptPDVS2miniSave()
            Label229.Enabled = False
            LabelTemperature3.Enabled = False
            volts11.Enabled = False
            ButtonDacSpan10down.Enabled = False
            ButtonDacSpan10.Enabled = False
            ButtonDacSpan10Up.Enabled = False
            LabeldacSpan10Cal.Enabled = False
            Label150.Enabled = False
            DacSpan10.Enabled = False
            LabeldacSpan10Delta.Enabled = False
            Default11.Enabled = False
            Label149.Enabled = False
            WryTech.Checked = My.Settings.data503

            'btncreate2.Enabled = True
            'btncreate3.Enabled = True

            ButtonR6581upload.Enabled = False
            ButtonR6581commitEEprom.Enabled = False

            ' Settings
            CheckBoxAllowSaveAnytime.Checked = My.Settings.data505
            TextBoxTextEditor.Text = My.Settings.data506
            CheckBoxEnableTooltips.Checked = My.Settings.data507

            '3458A CalRam
            CalRam3458APreRun.Text = My.Settings.data508

            If String.IsNullOrWhiteSpace(TextBoxTextEditor.Text) Then
                TextBoxTextEditor.Text = "C:\Windows\System32\notepad.exe"
                My.Settings.data506 = "C:\Windows\System32\notepad.exe"
                My.Settings.Save()
            End If

            If CheckBoxEnableTooltips.Checked = True Then
                ToolTip1.Active = True
            Else
                ToolTip1.Active = False
            End If

            ' Tooltip durations - If enabled
            ToolTip1.AutoPopDelay = 10000   ' Time in milliseconds tooltip stays visible
            ToolTip1.InitialDelay = 100     ' Delay before tooltip appears
            ToolTip1.ReshowDelay = 100      ' Delay before tooltip reappears if user moves away and back

            Dev1ChartValue.Text = ""
            Dev2ChartValue.Text = ""

            ' 3245A Dev2
            Disable3245AControls()



        Catch ex As Exception
            MessageBox.Show($"Error during load: {ex.Message}")
        End Try

    End Sub


    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        ' Create a solid red brush
        Dim redBrush As New SolidBrush(Color.Red)

        ' Define the rectangle's location and size
        Dim rect As New Rectangle(50, 50, 200, 100)

        ' Fill the rectangle with the red brush
        e.Graphics.FillRectangle(redBrush, rect)

        ' Dispose of the brush when done
        redBrush.Dispose()
    End Sub


    Function CreateDevice(ByVal name As String, ByVal address As String, ByVal interfacetype As Integer) As IODevice

        ' This function returns a generic IODevice object: here we can define devices polymorphically 
        ' because in this test code we don't use many interface-specific methods

        ' This is convenient if there are not too many interface-dependent options  (eg visa attributes, see below)
        ' (but otherwise it is still possible to bind statically without any change in the program other than initialization)

        Dim dev As IODevice

        Try

            Select Case interfacetype

                'Case 0 : dev = New VisaDevice(name, address)           ' not activated interlock
                'Case 0 : dev = New VisaDevice(name, address, True)      ' activated interlock. Before IanJ mod 


                Case 0 : If (noEOI.Checked = True) Then
                        dev = New VisaDevice(name, address, True)
                    Else
                        dev = New VisaDevice_noEOI(name, address, True, 10)       ' terminator set to LF (decimal 10)
                        'dev = New VisaDevice(name, address, True, 10)       ' terminator set to LF (decimal 10)
                    End If

                Case 1 : dev = New GPIBDevice_ADLink(name, address)
                Case 2 : dev = New GPIBDevice_gpib488(name, address)
                Case 3 : dev = New SerialDevice(name, address)
                Case 4 : dev = New PrologixDeviceSerial(name, address)
                Case 5 : dev = New PrologixDeviceEthernet(name, address)
                Case 6 : dev = New GPIB_NI_232CT_A(name, address)
                Case Else : Return Nothing
            End Select

        Catch ex As Exception  '(constructor exception: most often "dll not found")
            Dim msg As String = " cannot create device " & name & vbCrLf & ex.Message
            If ex.InnerException IsNot Nothing AndAlso ex.Message <> ex.InnerException.Message Then msg = msg & vbCrLf & ex.InnerException.Message
            MessageBox.Show(msg)
            dev = Nothing
            IODevice.statusmsg = ""
        End Try

        ' Option to debug interface routines, may remove this line in final version:
        ' If dev IsNot Nothing Then dev.catchinterfaceexceptions = False 

        Return dev


    End Function


    Private Sub Setnotify(ByVal dev As IODevice)

        Try
            dev.EnableNotify = True 'default implementation will throw an exception if not available for the selected interface
            Dim result As Integer = dev.SendBlocking("*SRE 16", False) ' set bit 4 in Service Request Enable Register, so that the MAV status will set SRQ
            If result = 0 Then
                dev.delayread = 1000 'set long wait delays (will be interrupted anyway)
                dev.delayrereadontimeout = 1000
            End If
        Catch ex As Exception
            Dim msg As String = " cannot set EnableNotify for device " & dev.devname & vbCrLf & ex.Message
            If ex.InnerException IsNot Nothing AndAlso ex.Message <> ex.InnerException.Message Then
                msg = msg & vbCrLf & ex.InnerException.Message
            End If
            MessageBox.Show(msg)
        End Try


    End Sub

    Private Sub Btncreate2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncreate2.Click

        dev1 = CreateDevice(txtname1.Text, txtaddr1.Text, lstIntf1.SelectedIndex)

        ButtonDev1Run.Enabled = True

        ButtonCalramDump3457A.Enabled = True
        ButtonCalramDump3458A.Enabled = True

        'to simulate slow response from device 2 we can set:
        'dev2.delayread = 2000;

        'example of interface-specific action: setting visa options in case visa is selected for dev1:
        Dim d As IODevice = TryCast(dev1, VisaDevice)
        If d IsNot Nothing Then
            'for example set attributes:
            'd.SetAttribute(  attr, value)
        End If

        If dev1 IsNot Nothing Then
            gbox1.Enabled = True

            'examples of some settings
            dev1.maxtasks = 10
            'dev1.readtimeout = 5000
            dev1.readtimeout = Val(Dev1Timeout.Text)

            dev1.delayop = Val(Dev1delayop.Text)

            dev1.showmessages = True
            dev1.catchcallbackexceptions = True

            EditMode.Enabled = False

            TextBoxDev1CMD.Enabled = True
            TextBoxDev1CMD.BackColor = Color.White
            CMD1clear.Enabled = True
            CMDdev1.Text = txtname1.Text
            Device1nameLive.Text = txtname1.Text
            TextBoxDev1CMD.Text = TextBoxDev1CMD.Text + "READY!" + Environment.NewLine  ' vbCrLf

            If CheckBoxAllowSaveAnytime.Checked = False Then
                ButtonSaveSettings.Enabled = False
            Else
                ButtonSaveSettings.Enabled = True
            End If

            btnBackup.Enabled = False
            btnRestore.Enabled = False
            ButtonAvailableComPorts.Enabled = False
            ProfDev1_1.Enabled = False
            ProfDev1_2.Enabled = False
            ProfDev1_3.Enabled = False
            ProfDev1_4.Enabled = False
            ProfDev1_5.Enabled = False
            ProfDev1_6.Enabled = False
            ProfDev1_7.Enabled = False
            ProfDev1_8.Enabled = False
            txtaddr1.Enabled = False
            lstIntf1.Enabled = False

        End If

        IODevice.ShowDevices()

        btncreate.Enabled = False
        btncreate2.Enabled = False
        btncreate3.Enabled = False
        'gboxdev.Enabled = False
        'btndevlist.Enabled = True

        ButtonReset.Enabled = True

        ' ButtonDev12Run.Enabled = False
        Dev1SampleRate.Enabled = True

        Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Dev.1 Created")


    End Sub

    Private Sub Btncreate3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncreate3.Click

        dev2 = CreateDevice(txtname2.Text, txtaddr2.Text, lstIntf2.SelectedIndex)

        ButtonDev2Run.Enabled = True

        'to simulate slow response from device 2 we can set:
        'dev2.delayread = 2000;

        'example of interface-specific action: setting visa options in case visa is selected for dev1:
        Dim d As IODevice = TryCast(dev2, VisaDevice)
        If d IsNot Nothing Then
            'for example set attributes:
            'd.SetAttribute(  attr, value)
        End If

        If dev2 IsNot Nothing Then
            gbox2.Enabled = True

            'examples of some settings
            dev2.maxtasks = 10
            'dev2.readtimeout = 5000
            dev2.readtimeout = Val(Dev2Timeout.Text)

            dev2.delayop = Val(Dev2delayop.Text)

            dev2.showmessages = True
            dev2.catchcallbackexceptions = True

            EditMode.Enabled = False

            TextBoxDev2CMD.Enabled = True
            TextBoxDev2CMD.BackColor = Color.White
            CMD2clear.Enabled = True
            CMDdev2.Text = txtname2.Text
            Device2nameLive.Text = txtname2.Text
            TextBoxDev2CMD.Text = TextBoxDev2CMD.Text + "READY!" + Environment.NewLine

            If CheckBoxAllowSaveAnytime.Checked = False Then
                ButtonSaveSettings.Enabled = False
            Else
                ButtonSaveSettings.Enabled = True
            End If

            btnBackup.Enabled = False
            btnRestore.Enabled = False
            ButtonAvailableComPorts.Enabled = False
            ProfDev2_1.Enabled = False
            ProfDev2_2.Enabled = False
            ProfDev2_3.Enabled = False
            ProfDev2_4.Enabled = False
            ProfDev2_5.Enabled = False
            ProfDev2_6.Enabled = False
            ProfDev2_7.Enabled = False
            ProfDev2_8.Enabled = False
            txtaddr2.Enabled = False
            lstIntf2.Enabled = False

            Enable3245AControls()

        End If

        IODevice.ShowDevices()

        btncreate.Enabled = False
        btncreate2.Enabled = False
        btncreate3.Enabled = False
        'gboxdev.Enabled = False
        'btndevlist.Enabled = True

        ButtonReset.Enabled = True

        Dev2SampleRate.Enabled = True

        Button3245A_DCV.Enabled = Enabled
        Button3245A_ACV.Enabled = Enabled
        Button3245A_DCI.Enabled = Enabled
        Button3245A_SQV.Enabled = Enabled
        Button3245A_RPV.Enabled = Enabled
        Button3245A_HIRES.Enabled = Enabled
        Button3245A_LORES.Enabled = Enabled
        Button3245A_SCRATCH.Enabled = Enabled
        Button3245A_ACI.Enabled = Enabled
        Button3245A_SQI.Enabled = Enabled
        Button3245A_RPI.Enabled = Enabled
        Button3245A_FREQ.Enabled = Enabled
        Button3245A_DUTY.Enabled = Enabled
        Button3245A_RESET.Enabled = Enabled

        Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Dev.2 Created")

    End Sub

    Private Sub Btncreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncreate.Click

        ' Dev 1 & Dev 2 devices

        'Dev1IntEnable.Checked = False
        'Dev2IntEnable.Checked = False
        'Dev1IntEnable.Enabled = False
        'Dev2IntEnable.Enabled = False

        gbox12.Enabled = True
        Dev1SampleRate.Enabled = False
        Dev2SampleRate.Enabled = False
        Dev12SampleRate.Enabled = True

        ButtonDev1Run.Enabled = False
        ButtonDev2Run.Enabled = False

        dev1 = CreateDevice(txtname1.Text, txtaddr1.Text, lstIntf1.SelectedIndex)
        dev2 = CreateDevice(txtname2.Text, txtaddr2.Text, lstIntf2.SelectedIndex)

        'to simulate slow response from device 2 we can set:
        'dev2.delayread = 2000;

        'example of interface-specific action: setting visa options in case visa is selected for dev1:
        Dim d As IODevice = TryCast(dev1, VisaDevice)
        If d IsNot Nothing Then
            'for example set attributes:
            'd.SetAttribute(  attr, value)
        End If

        If dev1 IsNot Nothing Then
            gbox1.Enabled = True

            'examples of some settings
            dev1.maxtasks = 10
            'dev1.readtimeout = 5000
            dev1.readtimeout = Val(Dev1Timeout.Text)

            dev1.delayop = Val(Dev1delayop.Text)

            dev1.showmessages = True
            dev1.catchcallbackexceptions = True

            'dev1.enablepoll = False  'uncomment this if a device does not support polling ("poll timeout" is signalled)  - IanJ 16/10/2020

            'dev1.MAVmask = ...  'set a different mask if your device supports polling but its status flags do not comply with 488.2

            'uncomment this to use SRQ notifying on device1 :
            'setnotify(dev1)

            EditMode.Enabled = False

            Device1nameLive.Text = txtname1.Text

            If CheckBoxAllowSaveAnytime.Checked = False Then
                ButtonSaveSettings.Enabled = False
            Else
                ButtonSaveSettings.Enabled = True
            End If

            btnBackup.Enabled = False
            btnRestore.Enabled = False
            ButtonAvailableComPorts.Enabled = False
            ProfDev1_1.Enabled = False
            ProfDev1_2.Enabled = False
            ProfDev1_3.Enabled = False
            ProfDev1_4.Enabled = False
            ProfDev1_5.Enabled = False
            ProfDev1_6.Enabled = False
            ProfDev1_7.Enabled = False
            ProfDev1_8.Enabled = False
            txtaddr1.Enabled = False
            lstIntf1.Enabled = False
            ProfDev2_1.Enabled = False
            ProfDev2_2.Enabled = False
            ProfDev2_3.Enabled = False
            ProfDev2_4.Enabled = False
            ProfDev2_5.Enabled = False
            ProfDev2_6.Enabled = False
            ProfDev2_7.Enabled = False
            ProfDev2_8.Enabled = False
            txtaddr2.Enabled = False
            lstIntf2.Enabled = False

        End If


        If dev2 IsNot Nothing Then

            gbox2.Enabled = True

            'examples of some settings
            dev2.maxtasks = 10
            'dev2.readtimeout = 5000
            dev2.readtimeout = Val(Dev2Timeout.Text)

            dev2.delayop = Val(Dev2delayop.Text)

            dev2.showmessages = True
            dev2.catchcallbackexceptions = True

            'uncomment this to use SRQ notifying on device2 :
            ' setnotify(dev2)

            Device2nameLive.Text = txtname2.Text

        End If


        ' enable 3245A Cal button
        If dev1 IsNot Nothing And dev2 IsNot Nothing Then
            ButtonCal3245A.Enabled = True
        End If



        IODevice.ShowDevices()

        btncreate.Enabled = False
        btncreate2.Enabled = False
        btncreate3.Enabled = False
        'gboxdev.Enabled = False
        'btndevlist.Enabled = True

        ButtonReset.Enabled = True

        TextBoxDev1CMD.Enabled = True
        CMDdev1.Text = txtname1.Text
        TextBoxDev1CMD.BackColor = Color.White
        TextBoxDev2CMD.Enabled = True
        CMDdev2.Text = txtname2.Text
        TextBoxDev2CMD.BackColor = Color.White
        TextBoxDev1CMD.Text = TextBoxDev1CMD.Text + "READY!" + Environment.NewLine
        TextBoxDev2CMD.Text = TextBoxDev2CMD.Text + "READY!" + Environment.NewLine
        CMD1clear.Enabled = True
        CMD2clear.Enabled = True

        Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " Dev.1 & Dev.2 Created")

    End Sub


    Private Sub Btndevlist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndevlist.Click

        IODevice.ShowDevices()

    End Sub


    Private Sub Btnq1b_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnq1b.Click

        ' DEV1 - Query BLOCKING

        Dim q As IOQuery = Nothing
        Dim result As Integer

        Dim s As String

        ' Set up options - Added by IanJ
        If Dev1PollingEnable.Checked = True Then
            dev1.enablepoll = True
        Else
            dev1.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
        End If

        If Dev1STBMask.Text = "" Then
            Dev1STBMask.Text = "16"
        End If
        dev1.MAVmask = Val(Dev1STBMask.Text)
        If Dev1STBMask.Text = "0" Then
            dev1.enablepoll = False
            Dev1PollingEnable.Checked = False
        End If

        btnq1b.Enabled = False
        'result = dev1.QueryBlocking(txtq1b.Text, q, True) 'standard version with IOQuery parameter
        result = dev1.QueryBlocking(txtq1b.Text & TermStr2(), q, True) 'simpler version with string parameter, modified for BB3 operation

        btnq1b.Enabled = True

        s = "blocking command:'" & q.cmd & "'" & vbCrLf

        If result = 0 Then

            txtr1b.Text = q.ResponseAsString
            s &= "device response time:" & Str(q.timeend.Subtract(q.timestart).TotalSeconds) & " s" & vbCrLf
            s &= "thread wait time:" & Str(q.timestart.Subtract(q.timecall).TotalSeconds) & " s" & vbCrLf

        Else
            s &= "status: " & result & vbCrLf
            s &= q.errmsg
        End If

        txtr1a_disp.Text = Format$(Val(txtr1b.Text), "0.0########################")     ' 0 = Digit placeholder. Displays a digit or a zero. # = Digit placeholder. Displays a digit or nothing. So, it's 0,0 bare minimum.

        txtr1astat.Text = s

    End Sub


    Private Sub Btnq2b_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnq2b.Click

        ' DEV2 - Query BLOCKING

        Dim q As IOQuery = Nothing
        Dim result As Integer

        Dim s As String

        'Dim resp As String = Nothing
        'Dim result As Integer

        'btnq2b.Enabled = False
        'result = dev2.QueryBlocking(txtq2b.Text, resp, True) 'simpler version with string parameter
        'btnq2b.Enabled = True

        'If result = 0 Then
        'txtr2b.Text = resp
        'End If

        ' Set up options - Added by IanJ
        If Dev2PollingEnable.Checked = True Then
            dev2.enablepoll = True      'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
        Else
            dev2.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
        End If

        If Dev2STBMask.Text = "" Then
            Dev2STBMask.Text = "16"
        End If
        dev2.MAVmask = Val(Dev2STBMask.Text)
        If Dev2STBMask.Text = "0" Then
            dev2.enablepoll = False
            Dev2PollingEnable.Checked = False
        End If

        ' Modified by IanJ in association with BB3 software guy
        'Dim resp As String = Nothing
        'Dim result As Integer

        btnq2b.Enabled = False

        result = dev2.QueryBlocking(txtq2b.Text & TermStr2(), q, True) 'simpler version with string parameter, modified for BB3 operation

        btnq2b.Enabled = True

        'result = dev2.QueryBlocking(txtq2b.Text & TermStr(), resp, True) 'simpler version with string parameter
        'btnq2b.Enabled = True

        'txtr2a_disp.Text = Format$(Val(txtr2b.Text), "0.0########################")     ' 0 = Digit placeholder. Displays a digit or a zero. # = Digit placeholder. Displays a digit or nothing. So, it's 0,0 bare minimum.

        'If result = 0 Then
        'txtr2b.Text = resp
        'End If

        s = "blocking command:'" & q.cmd & "'" & vbCrLf

        If result = 0 Then

            txtr2b.Text = q.ResponseAsString
            s &= "device response time:" & Str(q.timeend.Subtract(q.timestart).TotalSeconds) & " s" & vbCrLf
            s &= "thread wait time:" & Str(q.timestart.Subtract(q.timecall).TotalSeconds) & " s" & vbCrLf

        Else
            s &= "status: " & result & vbCrLf
            s &= q.errmsg
        End If

        txtr2a_disp.Text = Format$(Val(txtr2b.Text), "0.0########################")     ' 0 = Digit placeholder. Displays a digit or a zero. # = Digit placeholder. Displays a digit or nothing. So, it's 0,0 bare minimum.

        txtr2astat.Text = s



    End Sub


    Private Sub Btnq1a_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnq1a.Click

        RunBtnq1aCore()

    End Sub


    Private Sub RunBtnq1aCore()

        ' DEV1 - Query ASYNC

        ' Set up options - Added by IanJ
        If Dev1PollingEnable.Checked = True Then
            dev1.enablepoll = True
        Else
            dev1.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
        End If

        If Dev1STBMask.Text = "" Then
            Dev1STBMask.Text = "16"
        End If
        dev1.MAVmask = Val(Dev1STBMask.Text)
        If Dev1STBMask.Text = "0" Then
            dev1.enablepoll = False
            Dev1PollingEnable.Checked = False
        End If


        If dev1.PendingTasks(txtq1a.Text) <= 3 Then
            'example of using PendingTasks() method
            'dev1.QueryAsync(txtq1a.Text, AddressOf cbdev1, True)

            ' Modified by IanJ in association with BB3 software guy
            dev1.QueryAsync(txtq1a.Text & TermStr2(), AddressOf Cbdev1, True)

        Else
            txtr1astat.Text = "already 3 '" & txtq1a.Text & "' commands pending"
        End If

    End Sub


    Private Sub Btnq2a_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnq2a.Click

        RunBtnq2aCore()

    End Sub


    Private Sub RunBtnq2aCore()

        ' DEV2 - Query ASYNC

        ' Set up options - Added by IanJ
        If Dev2PollingEnable.Checked = True Then
            dev2.enablepoll = True      'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
        Else
            dev2.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
        End If

        If Dev2STBMask.Text = "" Then
            Dev2STBMask.Text = "16"
        End If
        dev2.MAVmask = Val(Dev2STBMask.Text)
        If Dev2STBMask.Text = "0" Then
            dev2.enablepoll = False
            Dev2PollingEnable.Checked = False
        End If


        If dev2.PendingTasks(txtq2a.Text) <= 3 Then
            'example of using PendingTasks() method
            'dev2.QueryAsync(txtq2a.Text, AddressOf cbdev2, True)

            ' Modified by IanJ in association with BB3 software guy
            dev2.QueryAsync(txtq2a.Text & TermStr2(), AddressOf Cbdev2, True)

        Else
            txtr2astat.Text = "already 3 '" & txtq2a.Text & "' commands pending"
        End If


        ' Modified by IanJ in association with BB3 software guy
        'dev2.QueryAsync(txtq2a.Text & TermStr(), AddressOf cbdev2, True)
        'dev2.QueryAsync(txtq2a.Text, AddressOf cbdev2, True)       ' original

    End Sub


    Private Sub Btns1c_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btns1c.Click
        RunBtns1cCore()
    End Sub

    Private Sub RunBtns1cCore()
        dev1.SendAsync(txtq1c.Text, True)
        txtr1astat.Text = "Send Async '" & txtq1c.Text
    End Sub


    ' Added by IanJ in association with BB3 software guy
    Private Function TermStr() As String
        Dim returnValue As String = "" ' Default return value

        If Dev2TerminatorEnable.Checked Then
            returnValue = ControlChars.NewLine
        End If

        If Dev2TerminatorEnable2.Checked Then
            returnValue = ControlChars.CrLf
        End If

        'If Dev2TerminatorEnable3.Checked Then
        'returnValue = ControlChars.CrLf & ControlChars.Lf
        'End If

        'If Not Dev2TerminatorEnable.Checked And Not Dev2TerminatorEnable2.Checked And Not Dev2TerminatorEnable3.Checked Then
        If Not Dev2TerminatorEnable.Checked And Not Dev2TerminatorEnable2.Checked Then
            returnValue = ""
        End If

        Return returnValue
    End Function


    ' Added by IanJ in association with BB3 software guy
    Private Function TermStr2() As String
        Dim returnValue As String = "" ' Default return value

        If Dev1TerminatorEnable.Checked Then
            returnValue = ControlChars.NewLine
        End If

        If Dev1TerminatorEnable2.Checked Then
            returnValue = ControlChars.CrLf
        End If

        'If Dev1TerminatorEnable3.Checked Then
        'returnValue = ControlChars.CrLf & ControlChars.Lf
        'End If

        'If Not Dev1TerminatorEnable.Checked And Not Dev1TerminatorEnable2.Checked And Not Dev1TerminatorEnable3.Checked Then
        If Not Dev1TerminatorEnable.Checked And Not Dev1TerminatorEnable2.Checked Then
            returnValue = ""
        End If

        Return returnValue
    End Function


    Private Sub Btns2c_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btns2c.Click
        RunBtns2cCore()
    End Sub

    Private Sub RunBtns2cCore()
        dev2.SendAsync(txtq2c.Text & TermStr(), True)
        txtr2astat.Text = "Send Async '" & txtq2c.Text
    End Sub


    'signature of callback functions:   Public Delegate Sub IOCallback(ByVal q As IOQuery) 

    Sub Cbdev1(ByVal q As IOQuery)

        Try

            Dim s As String = "async command:'" & q.cmd & "'" & vbCrLf
            If q.status = 0 And Dev1TextResponse.Checked = False Then

                inst_value1F = q.ResponseAsString   ' for PDVS2mini calibration

                ' Update CMD line only if it was used
                If CMDlineOp = True Then
                    ' load the command line variable with the response
                    TextBoxDev1CMD.AppendText(Environment.NewLine & ">    " & q.ResponseAsString & Environment.NewLine)
                    CMDlineOp = False
                End If

                If Dev13457Aseven.Checked = False Then
                    txtr1a.Text = q.ResponseAsString
                End If

                ' Enable 7th digit mode for 3457A DMM only. Send extra command after TARM SGL to get the additional data
                If Dev13457Aseven.Checked = True Then

                    temp3457A_Dev1 = q.ResponseAsString

                    If Dev1_3457A = True Then    ' 7th digit process
                        inst_value1_3457A_2 = CDbl(Val(temp3457A_Dev1))             ' convert digit 7 to double
                        Dev1_3457A = False
                    Else
                        inst_value1_3457A_1 = CDbl(Val(temp3457A_Dev1))             ' convert digits 1-6 to double
                        Exit Sub    ' Exit sub since we need digit 7 before can process
                    End If

                    inst_value1_3457A_sum = inst_value1_3457A_1 + inst_value1_3457A_2   ' add them together
                    txtr1a.Text = inst_value1_3457A_sum

                End If


                ' Remove letters A to Z from return from device, i.e. Racal-Dana 1991 returns +FA00000.0000 i.e. echos last command back ahead of numbers
                If Dev1removeletters.Checked = True Then
                    txtr1a.Text = txtr1a.Text.Remove(0, 2)  ' remove first 2 characters i.e. FA+00000.0000 (RACAL-DANA 1991 has 2 letters ahead of numerical data)
                End If


                ' Remove unwanted data from Keithley 2001 DMM, i.e.
                ' 1.2345678E+00NVDC,+47.354841SECS,+01096RDNG#,00EXTCHAN
                ' +12.34E-03NVAC,+211.709381SECS,+05638RDNG#,00EXTCHAN
                ' +0.01539E+00NOHM,+269.504250SECS,+06843RDNG#,00EXTCHAN
                If Dev1K2001isolatedata.Checked = True And Dev1K2001isolatedataCHAR.Text <> "" Then
                    Dim str As String = txtr1a.Text
                    Dim leftPart As String = str.Split(Dev1K2001isolatedataCHAR.Text)(0)  ' Isolate at CHAR, then accessing 0 gives you the first part
                    txtr1a.Text = leftPart
                End If


                ' Isolate numerical data even if text appears on left and right of data req'd. DP's are retained.
                If Dev1Regex.Checked = True Then
                    'txtr1a.Text = "abcde9.999873111E-01fghij"       ' for test only
                    ' Regular expression pattern to match numbers with optional decimal points
                    'Dim pattern As String = "(\d+(\.\d+)?)"
                    Dim pattern As String = "[-+]?\d*\.?\d+([eE][-+]?\d+)?"     ' includes capacity for scientific notation
                    ' Match the pattern in the input string
                    Dim match As Match = Regex.Match(txtr1a.Text, pattern)
                    ' Check if a match is found
                    If match.Success Then
                        ' Extract the matched numeric value
                        txtr1a.Text = match.Value
                    End If
                End If


                ' Divide value by 1000, i.e. when measuring 100K resistors etc
                If Div1000Dev1.Checked = True Then
                    txtr1a.Text = Val(txtr1a.Text) / 1000
                End If


                ' Multiply value by 1000
                If Mult1000Dev1.Checked = True Then
                    txtr1a.Text = Val(txtr1a.Text) * 1000
                End If


                If txtOperationDev1.Text.Trim() <> "" Then
                    Dim input As String = txtOperationDev1.Text.Trim()

                    If input.Length < 2 Then
                        MessageBox.Show("Arithmetic: Please enter an operator (*, /, +, -) followed by a number.",
                        "Arithmetic: Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        Dim op As Char = input(0)
                        Dim numberPart As String = input.Substring(1).Trim()

                        ' Check if the first character is one of the allowed operators
                        If op <> "*" AndAlso op <> "/" AndAlso op <> "+" AndAlso op <> "-" Then
                            MessageBox.Show("Arithmetic: Invalid operator. Please start with * / + or -.",
                            "Arithmetic: Invalid Operator", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            Dim operand As Double
                            If Not Double.TryParse(numberPart, operand) Then
                                MessageBox.Show("Arithmetic: The number entered is not valid.",
                                "Arithmetic: Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                Dim currentValue As Double
                                If Not Double.TryParse(txtr1a.Text, currentValue) Then
                                    MessageBox.Show("Arithmetic: The current value is not valid.",
                                    "Arithmetic: Invalid Value", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                Else
                                    Select Case op
                                        Case "*" ' Multiply
                                            currentValue *= operand
                                        Case "/" ' Divide
                                            If operand = 0 Then
                                                MessageBox.Show("Arithmetic: Division by zero is not allowed.",
                                                "Arithmetic: Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                            Else
                                                currentValue /= operand
                                            End If
                                        Case "+" ' Add
                                            currentValue += operand
                                        Case "-" ' Subtract
                                            currentValue -= operand
                                    End Select
                                    txtr1a.Text = currentValue.ToString()
                                End If
                            End If
                        End If
                    End If
                End If


                ' Convert to decimal for display (data incoming can be a mix of e-notation or decimal)
                'txtr1a_disp.Text = Format$(Val(txtr1a.Text), "0.0########################")     ' 0 = Digit placeholder. Displays a digit or a zero. # = Digit placeholder. Displays a digit or nothing. So, it's 0,0 bare minimum.
                Dim decimalPlaces As Integer
                If Integer.TryParse(Dev1DecimalNumDPs.Text, decimalPlaces) Then
                    ' Use the user input to format the number
                    Dim formatString As String = "0." & New String("0"c, decimalPlaces)
                    txtr1a_disp.Text = Format(Val(txtr1a.Text), formatString)
                Else
                    ' Handle invalid input (non-integer)
                    txtr1a_disp.Text = Format$(Val(txtr1a.Text), "0.0########################")
                End If


                ' Device Meter
                Dim Dev1Temp1 As String = Format(Val(txtr1a.Text), "#00.000000000")
                If Val(txtr1a.Text) >= 1 Then
                    Dev1Temp1 = Dev1Temp1.TrimStart("0"c)  ' remove leading zeros
                End If
                If Val(txtr1a.Text) < 1 Then
                    Dev1Temp1 = Format(Val(txtr1a.Text), "#0.000000000")    ' If less than "1" the only one leading zero
                End If
                Dev1Temp1 = Dev1Temp1.TrimEnd("0"c)  ' remove trailing zeros
                'Dev1Meter.Text = Dev1Temp1
                Device1name.Text = txtname1.Text
                If ButtonDev1Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop" Then
                    'Dev1Meter.Text = Dev1Temp1
                    ' Setr number of DP's
                    If Integer.TryParse(Dev1DecimalNumDPs.Text, decimalPlaces) Then
                        ' Use the user input to format the number
                        Dim formatString As String = "0." & New String("0"c, decimalPlaces)
                        Dev1Meter.Text = Format(Val(Dev1Temp1), formatString)
                    Else
                        ' Handle invalid input (non-integer)
                        Dev1Meter.Text = Dev1Temp1
                    End If
                Else
                    Dev1Meter.Text = "---------------"
                End If
                DeviceTemperature.Text = LabelTemperature.Text
                DeviceHumidity.Text = LabelHumidity.Text

                s &= "device response time:" & q.timeend.Subtract(q.timestart).TotalSeconds.ToString() & " s" & vbCrLf

                ' If timer running, i.e. not manual
                If ButtonDev1Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop" Then
                    Dev1GPIBActivity = True
                    'Dev1SampleCount = Dev1SampleCount + 1
                    Dev1SampleCount += 1
                    Dev1Samples.Text = Dev1SampleCount
                    CSVfile()       ' CSV file
                    LOGdisplay()    ' LOG display
                    LiveChart()     ' Live chart
                    Dev1GPIBActivity = False
                End If

            Else
                s &= "error " & q.errcode & vbCrLf
            End If


            ' Response expected is text not numerical
            If q.status = 0 And Dev1TextResponse.Checked = True Then
                txtr1a.Text = q.ResponseAsString
            Else
                s &= "error " & q.errcode & vbCrLf
            End If


            s &= "thread wait time:" & q.timestart.Subtract(q.timecall).TotalSeconds.ToString() & " s" & vbCrLf

            txtr1astat.Text = s

            'uncomment this to chain on dev2:
            ' dev2.QueryAsync(txtq2a.Text, AddressOf cbdev2, True)

            ' Update USER tab variable
            USERdev1output = txtr1a_disp.Text
            USERdev1output2 = txtr1a.Text           ' alternative
            OutputReceiveddev1 = True

        Catch ex As Exception
            txtr1astat.Text = q.cmd & " error in callback function:" & vbCrLf
            txtr1astat.Text &= ex.Message & vbCrLf
            If ex.InnerException IsNot Nothing Then
                txtr1astat.Text &= ex.InnerException.Message
            End If

        End Try

    End Sub


    Sub Cbdev2(ByVal q As IOQuery)

        Try

            Dim s As String = "async command:'" & q.cmd & "'" & vbCrLf
            If q.status = 0 And Dev2TextResponse.Checked = False Then
                'txtr2a.Text = q.ResponseAsString

                ' Update CMD line only if it was used
                If CMDlineOp = True Then
                    ' load the command line variable with the response
                    TextBoxDev2CMD.AppendText(Environment.NewLine & ">    " & q.ResponseAsString & Environment.NewLine)
                    CMDlineOp = False
                End If


                If Dev23457Aseven.Checked = False Then
                    txtr2a.Text = q.ResponseAsString
                End If


                ' Enable 7th digit mode for 3457A DMM only. Send extra command after TARM SGL to get the additional data
                If Dev23457Aseven.Checked = True Then

                    temp3457A_Dev2 = q.ResponseAsString

                    If Dev2_3457A = True Then    ' 7th digit process
                        inst_value2_3457A_2 = CDbl(Val(temp3457A_Dev2))             ' convert digit 7 to double
                        Dev2_3457A = False
                    Else
                        inst_value2_3457A_1 = CDbl(Val(temp3457A_Dev2))             ' convert digits 1-6 to double
                        Exit Sub    ' Exit sub since we need digit 7 before can process
                    End If

                    inst_value2_3457A_sum = inst_value2_3457A_1 + inst_value2_3457A_2   ' add them together
                    'txtr2a.Text = Format(inst_value2_3457A_sum, "#0.0000000")   ' limit to 7 DP's
                    txtr2a.Text = inst_value2_3457A_sum

                End If


                ' Remove letters A to Z from return from device, i.e. Racal-Dana 1991 returns +FA00000.0000 i.e. echos last command back ahead of numbers
                If Dev2removeletters.Checked = True Then
                    txtr2a.Text = txtr2a.Text.Remove(0, 2)  ' remove first 2 characters i.e. FA+00000.0000 (RACAL-DANA 1991 has 2 letters ahead of numerical data)
                End If


                ' Remove unwanted data from Keithley 2001 DMM, i.e.
                ' 1.2345678E+00NVDC,+47.354841SECS,+01096RDNG#,00EXTCHAN
                ' +12.34E-03NVAC,+211.709381SECS,+05638RDNG#,00EXTCHAN
                ' +0.01539E+00NOHM,+269.504250SECS,+06843RDNG#,00EXTCHAN
                If Dev2K2001isolatedata.Checked = True And Dev2K2001isolatedataCHAR.Text <> "" Then
                    Dim str As String = txtr2a.Text
                    Dim leftPart As String = str.Split(Dev2K2001isolatedataCHAR.Text)(0)  ' Isolate at CHAR, then accessing 0 gives you the first part
                    txtr2a.Text = leftPart
                End If


                ' Isolate numerical data even if text appears on left and right of data req'd. DP's are retained.
                If Dev2Regex.Checked = True Then
                    ' Regular expression pattern to match numbers with optional decimal points
                    'Dim pattern As String = "(\d+(\.\d+)?)"
                    Dim pattern As String = "[-+]?\d*\.?\d+([eE][-+]?\d+)?"
                    ' Match the pattern in the input string
                    Dim match As Match = Regex.Match(txtr2a.Text, pattern)
                    ' Check if a match is found
                    If match.Success Then
                        ' Extract the matched numeric value
                        txtr2a.Text = match.Value
                    End If
                End If


                ' Divide value by 1000, i.e. when measuring 100K resistors etc
                If Div1000Dev2.Checked = True Then
                    txtr2a.Text = Val(txtr2a.Text) / 1000
                End If


                ' Multiply value by 1000
                If Mult1000Dev2.Checked = True Then
                    txtr2a.Text = Val(txtr2a.Text) * 1000
                End If


                If txtOperationDev2.Text.Trim() <> "" Then
                    Dim input As String = txtOperationDev2.Text.Trim()

                    If input.Length < 2 Then
                        MessageBox.Show("Arithmetic: Please enter an operator (*, /, +, -) followed by a number.",
                        "Arithmetic: Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        Dim op As Char = input(0)
                        Dim numberPart As String = input.Substring(1).Trim()

                        ' Check if the first character is one of the allowed operators
                        If op <> "*" AndAlso op <> "/" AndAlso op <> "+" AndAlso op <> "-" Then
                            MessageBox.Show("Arithmetic: Invalid operator. Please start with * / + or -.",
                            "Arithmetic: Invalid Operator", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            Dim operand As Double
                            If Not Double.TryParse(numberPart, operand) Then
                                MessageBox.Show("Arithmetic: The number entered is not valid.",
                                "Arithmetic: Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                Dim currentValue As Double
                                If Not Double.TryParse(txtr2a.Text, currentValue) Then
                                    MessageBox.Show("Arithmetic: The current value is not valid.",
                                    "Arithmetic: Invalid Value", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                Else
                                    Select Case op
                                        Case "*" ' Multiply
                                            currentValue *= operand
                                        Case "/" ' Divide
                                            If operand = 0 Then
                                                MessageBox.Show("Arithmetic: Division by zero is not allowed.",
                                                "Arithmetic: Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                            Else
                                                currentValue /= operand
                                            End If
                                        Case "+" ' Add
                                            currentValue += operand
                                        Case "-" ' Subtract
                                            currentValue -= operand
                                    End Select
                                    txtr2a.Text = currentValue.ToString()
                                End If
                            End If
                        End If
                    End If
                End If


                ' Convert to decimal for display (data incoming can be a mix of e-notation or decimal)
                'txtr2a_disp.Text = Format$(Val(txtr2a.Text), "0.0########################")
                Dim decimalPlaces As Integer
                If Integer.TryParse(Dev2DecimalNumDPs.Text, decimalPlaces) Then
                    ' Use the user input to format the number
                    Dim formatString As String = "0." & New String("0"c, decimalPlaces)
                    txtr2a_disp.Text = Format(Val(txtr2a.Text), formatString)
                Else
                    ' Handle invalid input (non-integer)
                    txtr2a_disp.Text = Format$(Val(txtr2a.Text), "0.0########################")
                End If


                ' Device Meter
                Dim Dev2Temp1 As String = Format(Val(txtr2a.Text), "#00.000000000")
                If Val(txtr2a.Text) >= 1 Then
                    Dev2Temp1 = Dev2Temp1.TrimStart("0"c)  ' remove leading zeros
                End If
                If Val(txtr2a.Text) < 1 Then
                    Dev2Temp1 = Format(Val(txtr2a.Text), "#0.000000000")    ' If less than "1" the only one leading zero
                End If
                Dev2Temp1 = Dev2Temp1.TrimEnd("0"c)  ' remove trailing zeros
                'Dev2Meter.Text = Dev2Temp1
                Device2name.Text = txtname2.Text
                If ButtonDev2Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop" Then
                    'Dev2Meter.Text = Dev2Temp1
                    ' Setr number of DP's
                    If Integer.TryParse(Dev2DecimalNumDPs.Text, decimalPlaces) Then
                        ' Use the user input to format the number
                        Dim formatString As String = "0." & New String("0"c, decimalPlaces)
                        Dev2Meter.Text = Format(Val(Dev2Temp1), formatString)
                    Else
                        ' Handle invalid input (non-integer)
                        Dev2Meter.Text = Dev2Temp1
                    End If
                Else
                    Dev2Meter.Text = "---------------"
                End If
                DeviceTemperature.Text = LabelTemperature.Text
                DeviceHumidity.Text = LabelHumidity.Text

                s &= "device response time:" & q.timeend.Subtract(q.timestart).TotalSeconds.ToString() & " s" & vbCrLf

                ' If timer running, i.e. not manual
                If ButtonDev2Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop" Then
                    Dev2GPIBActivity = True
                    'Dev2SampleCount = Dev2SampleCount + 1
                    Dev2SampleCount += 1
                    Dev2Samples.Text = Dev2SampleCount
                    CSVfile()       ' CSV file
                    LOGdisplay()    ' LOG display
                    LiveChart()     ' Live chart
                    Dev2GPIBActivity = False
                End If


            Else
                s &= "error " & q.errcode & vbCrLf
            End If


            ' Response expected is text not numerical
            If q.status = 0 And Dev2TextResponse.Checked = True Then
                txtr2a.Text = q.ResponseAsString
            Else
                s &= "error " & q.errcode & vbCrLf
            End If


            s &= "thread wait time:" & q.timestart.Subtract(q.timecall).TotalSeconds.ToString() & " s" & vbCrLf

            txtr2astat.Text = s

            'uncomment this to chain on dev1:
            'dev1.QueryAsync(txtq1a.Text, AddressOf cbdev1, True)

            ' Update USER tab variable
            USERdev2output = txtr2a_disp.Text
            USERdev2output2 = txtr2a.Text           ' alternative
            OutputReceiveddev2 = True

        Catch ex As Exception
            txtr2astat.Text = "error in callback function:" & vbCrLf
            txtr2astat.Text &= ex.Message & vbCrLf
            If ex.InnerException IsNot Nothing Then
                txtr2astat.Text &= ex.InnerException.Message
            End If
        End Try

        'Jumpout:

    End Sub


    'suggest correct address format:
    Sub Formataddr(ByVal interfacenum As Integer, ByVal txtaddr As TextBox)

        Dim sa As String = Trim(txtaddr.Text)
        Select Case interfacenum
            Case 1, 2
                If Not IsNumeric(sa) Then
                    txtaddr.Text = "1"

                End If
            Case 0
                If IsNumeric(sa) Then
                    txtaddr.Text = "GPIB0::" & sa & "::INSTR"
                Else
                    txtaddr.Text = "GPIB1::22::INSTR"
                    '  or USB format like:  "USB0::xxxx::xxxx::xxxx::INSTR"

                End If
            Case 3
                If IsNumeric(sa) Then
                    txtaddr.Text = "COM" & sa & ":9600,N,8,1,CRLF"
                Else
                    txtaddr.Text = "COM1:9600,N,8,1,CRLF"

                End If
                'Case Else
        End Select


    End Sub


    Private Sub LstIntf1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Formataddr(lstIntf1.SelectedIndex, txtaddr1)

    End Sub

    Private Sub LstIntf2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Formataddr(lstIntf2.SelectedIndex, txtaddr2)

    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick

        ' This timer on all the time and set to 100mS, used for general stuff

        ' Chart control

        If RunChart = True Then     ' live Chart is running/paused?
            ChartControl()
        End If

        If (ButtonDev1Run.Text = "Stop" Or ButtonDev2Run.Text = "Stop") Then
            ButtonReset.Enabled = False
        Else
            ButtonReset.Enabled = True
        End If

        ' CSV file
        'CSVfile()

        'Donate2.Visible = Not Donate2.Visible

    End Sub

    Private Sub Timer8_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer8.Tick

        ' For logging stopwatch
        ' Update stopwatch HRS:MINS if logging
        If ButtonDev1Run.Text = "Stop" Or ButtonDev2Run.Text = "Stop" Or ButtonDev12Run.Text = "Stop" Then
            UpdateStopwatchLabel()
        End If



        ' Banner title text animation
        ' Generate a random number of spaces between the two parts (e.g., between 5 and 20 spaces)
        Dim random As New Random()
        Dim randomSpaces As Integer = random.Next(30, 36)

        ' Create a StringBuilder to construct the final text with random spaces
        Dim result As New StringBuilder()
        result.Append(BannerText1) ' Append the first part

        ' Append the random spaces
        For i As Integer = 1 To randomSpaces
            result.Append(" ")
        Next

        result.Append(BannerText2) ' Append the second part

        ' Set the form text
        Me.Text = result.ToString()

    End Sub

    Private Sub ButtonReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonReset.Click

        ButtonSaveSettings.Enabled = True
        btnBackup.Enabled = True
        btnRestore.Enabled = True
        ButtonAvailableComPorts.Enabled = True
        ProfDev1_1.Enabled = True
        ProfDev1_2.Enabled = True
        ProfDev1_3.Enabled = True
        ProfDev1_4.Enabled = True
        ProfDev1_5.Enabled = True
        ProfDev1_6.Enabled = True
        ProfDev1_7.Enabled = True
        ProfDev1_8.Enabled = True
        txtaddr1.Enabled = True
        lstIntf1.Enabled = True
        ProfDev2_1.Enabled = True
        ProfDev2_2.Enabled = True
        ProfDev2_3.Enabled = True
        ProfDev2_4.Enabled = True
        ProfDev2_5.Enabled = True
        ProfDev2_6.Enabled = True
        ProfDev2_7.Enabled = True
        ProfDev2_8.Enabled = True
        txtaddr2.Enabled = True
        lstIntf2.Enabled = True

        Dev1IntEnable.Enabled = True
        Dev2IntEnable.Enabled = True

        ' Interrupt stopwatch
        Dev1elapsedTime = 0
        Dev1runStopwatch = False
        Dev1stopWatch.Reset()
        Dev1isStopwatchRunning = False
        Dev2elapsedTime = 0
        Dev2runStopwatch = False
        Dev2stopWatch.Reset()
        Dev2isStopwatchRunning = False

        EditMode.Enabled = True
        EditMode.Checked = False

        ' If (Dev1Run.Text = "Run" And Dev2Run.Text = "Run") Then
        btncreate.Enabled = True
        btncreate2.Enabled = True
        btncreate3.Enabled = True
        'gboxdev.Enabled = True
        gbox1.Enabled = False
        gbox2.Enabled = False

        gbox12.Enabled = False
        IODevice.DisposeAll()

        TextBoxDev1CMD.Enabled = False
        TextBoxDev2CMD.Enabled = False
        TextBoxDev1CMD.BackColor = Color.DarkGray
        TextBoxDev2CMD.BackColor = Color.DarkGray

        CMD1clear.Enabled = False
        CMD2clear.Enabled = False
        CMDdev1.Text = "#########"
        CMDdev2.Text = "#########"

        ButtonCal3245A.Enabled = False      ' disable 3245A Cal button

        ButtonCalramDump3457A.Enabled = False      ' disable 3257A Cal dump button
        ButtonCalramDump3458A.Enabled = False      ' disable 3258A Cal dumpbutton

        EnableChart1.Checked = False
        EnableChart2.Checked = False
        EnableChart1.Enabled = False
        EnableChart2.Enabled = False
        CheckBoxDevice1Hide.Enabled = False
        CheckBoxDevice2Hide.Enabled = False

        ' 3245A Dev2
        Disable3245AControls()

    End Sub

    Private Sub ShowFiles_Click(sender As Object, e As EventArgs) Handles ShowFiles.Click

        Process.Start("explorer.exe", String.Format("/n, /e, {0}", CSVfilepath.Text))

    End Sub

    Private Sub ShowFiles2_Click(sender As Object, e As EventArgs) Handles ShowFiles2.Click

        Process.Start("explorer.exe", String.Format("/n, /e, {0}", CSVfilepath.Text))

    End Sub

    Private Sub ShowFiles3_Click(sender As Object, e As EventArgs) Handles ShowFiles3.Click

        Process.Start("explorer.exe", String.Format("/n, /e, {0}", CSVfilepath.Text))

    End Sub

    Private Sub ButtonPlaybackChart_Click_1(sender As Object, e As EventArgs)
        Dim externalchart = New Chart()
        externalchart.Show()
    End Sub

    Private Sub ButtonHelp_Click_1(sender As Object, e As EventArgs)
        Dialog1.Show()
    End Sub

    Private Sub CSVdelimiterComma_CheckedChanged(sender As Object, e As EventArgs)
        CSVdelimit = ","
        My.Settings.data29 = ","
    End Sub

    Private Sub CSVdelimiterSemiColon_CheckedChanged(sender As Object, e As EventArgs)
        CSVdelimit = ";"
        My.Settings.data29 = ";"
    End Sub

    'Private Sub ButtonNotepad_Click_1(sender As Object, e As EventArgs)
    '    Process.Start(TextEditorPath, CSVfilepath.Text & "\" & "GPIBchannels.txt")
    'End Sub

    'Private Sub ButtonNotePad2_Click(sender As Object, e As EventArgs) Handles ButtonNotePad2.Click
    '   Process.Start(TextEditorPath, strPath & "\" & "GPIBchannels.txt")
    'End Sub

    Private Sub ButtonNotePad2_Click(sender As Object, e As EventArgs) Handles ButtonNotePad2.Click
        'Dim TextEditorPath As String = "C:\Program Files\Notepad++\notepad++.exe" ' Adjust the path as needed
        Dim filePath As String = System.IO.Path.Combine(CSVfilepath.Text, "GPIBchannels.txt")

        ' Check if the text editor exists
        If Not System.IO.File.Exists(TextEditorPath) Then
            MessageBox.Show("Text editor not found at specified path: " & TextEditorPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Check if the file to open exists
        If Not System.IO.File.Exists(filePath) Then
            MessageBox.Show("File not found: " & filePath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Attempt to open the file in the text editor
        Try
            Process.Start(TextEditorPath, filePath)
        Catch ex As Exception
            MessageBox.Show("An error occurred while trying to open the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub ButtonIanWebsite_Click_1(sender As Object, e As EventArgs) Handles ButtonIanWebsite.Click
        Process.Start("www.paypal.me/IanSJohnston")
    End Sub

    Private Sub MainActivity_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    End Sub


    Private Sub Timer7_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer7.Tick

        If SerialPort1.IsOpen = True Then   ' only if PDVS2mini comms open

            ' Auto-Calibration of PDVS2mini - When value is received from 3458A. Run from here since the value of txtr1a_disp.Text will be newest and correct
            txtr1aBIG.Text = txtr1a_disp.Text   ' copy latest value to PDVS2 big text
            'txtr1aBIG.Text = Format(txtr1a_disp.Text, "#0.000000000")

            ' Timer7 starts out at 500mS, now up that to 2.5secs for the final tweak routine
            If (DACcalparam = 14) Then
                Me.Timer7.Stop()
                Me.Timer7.Interval = 2500
                Me.Timer7.Start()
            End If

            PDVS2mini()

        End If

    End Sub

    Private Sub URL1_Click(sender As Object, e As EventArgs) Handles URL1.Click
        System.Diagnostics.Process.Start("https://www.ianjohnston.com/")
    End Sub

    Private Sub URL2_Click(sender As Object, e As EventArgs) Handles URL2.Click
        System.Diagnostics.Process.Start("https://www.youtube.com/user/IanScottJohnston")
    End Sub

    Private Sub URL3_Click(sender As Object, e As EventArgs) Handles URL3.Click
        System.Diagnostics.Process.Start("https://twitter.com/IanSJohnston")
    End Sub

    Private Sub TabControl1_Selected(sender As Object, e As System.Windows.Forms.TabControlEventArgs) Handles TabControl1.Selected

        ' Record previous tab selected. Stored in integer var
        If TabControl1.SelectedTab Is TabPage1 Then
            TabUsed = 1
        End If
        If TabControl1.SelectedTab Is TabPage2 Then
            TabUsed = 2
        End If
        If TabControl1.SelectedTab Is TabPage3 Then
            TabUsed = 3
        End If
        If TabControl1.SelectedTab Is TabPage4 Then
            TabUsed = 4
        End If
        If TabControl1.SelectedTab Is TabPage5 Then
            TabUsed = 5
            Label163.Text = TextBoxTempUnits.Text       ' PDVS2mini tab
            Label97.Text = TextBoxTempUnits.Text        ' PDVS2mini tab

            ' Disable controls on PDVS2mini tab depending if Device 1 is enabled or not
            If btncreate2.Enabled = False Then
                EnableAllButtonsInGroupBox2()
            Else
                DisableAllButtonsInGroupBox2ExceptPDVS2miniSave()
            End If

        End If
        If TabControl1.SelectedTab Is TabPage6 Then
            TabUsed = 6
        End If
        If TabControl1.SelectedTab Is TabPage7 Then
            TabUsed = 7
        End If
        If TabControl1.SelectedTab Is TabPage8 Then
            TabUsed = 8
            Label171.Text = "TEMP. " & TextBoxTempUnits.Text
            Label173.Text = "HUMIDITY. " & TextBoxHumUnits.Text
        End If
        'If TabControl1.SelectedTab Is TabPage9 Then
        'TabUsed = 9
        'End If

        ' When user selectes the PlayBack Chart tab then it automatically open the playback screen and then re-directs the tab selected back to the previously used tab. The PlayBack chart is given focus
        If TabControl1.SelectedTab Is TabPage9 Then

            ' Save off Device 1 name to a file for later use with Chart.vb
            Dim filePath As String = Path.Combine(CSVfilepath.Text, "Device1Name.txt")
            Using file As New StreamWriter(filePath, False)
                file.WriteLine(txtname1.Text)
                file.Close()
            End Using

            'Dim CSVpath As String = CSVfilepath.Text & "\" & CSVfilename.Text
            'System.IO.File.WriteAllText(CSVpath, "")

            Dim externalchart = New Chart()

            'TabControl1.SelectedTab = TabPage1      ' set focus back to Device 1/2 tab

            If TabUsed = 1 Then
                TabControl1.SelectedTab = TabPage1      ' set focus back to previous used tab
            End If
            If TabUsed = 2 Then
                TabControl1.SelectedTab = TabPage2      ' set focus back to previous used tab
            End If
            If TabUsed = 3 Then
                TabControl1.SelectedTab = TabPage3      ' set focus back to previous used tab
            End If
            If TabUsed = 4 Then
                TabControl1.SelectedTab = TabPage4      ' set focus back to previous used tab
            End If
            If TabUsed = 5 Then
                TabControl1.SelectedTab = TabPage5      ' set focus back to previous used tab
            End If
            If TabUsed = 6 Then
                TabControl1.SelectedTab = TabPage6      ' set focus back to previous used tab
            End If
            If TabUsed = 7 Then
                TabControl1.SelectedTab = TabPage7      ' set focus back to previous used tab
            End If
            If TabUsed = 8 Then
                TabControl1.SelectedTab = TabPage8      ' set focus back to previous used tab
            End If
            If TabUsed = 1 Then
                TabControl1.SelectedTab = TabPage11     ' set focus back to previous used tab
            End If

            externalchart.Show()    ' show the PlayBack chart
        End If
    End Sub

    Private Sub Dev2TerminatorEnable_CheckedChanged(sender As Object, e As EventArgs) Handles Dev2TerminatorEnable.CheckedChanged

        ' Only LF or CRLF allowed at any one time, not both
        If Dev2TerminatorEnable.Checked Then
            Dev2TerminatorEnable2.Checked = False
        End If

    End Sub

    Private Sub Dev2TerminatorEnable2_CheckedChanged(sender As Object, e As EventArgs) Handles Dev2TerminatorEnable2.CheckedChanged

        ' Only LF or CRLF allowed at any one time, not both
        If Dev2TerminatorEnable2.Checked Then
            Dev2TerminatorEnable.Checked = False
        End If

    End Sub

    Private Sub Dev1TerminatorEnable_CheckedChanged(sender As Object, e As EventArgs) Handles Dev1TerminatorEnable.CheckedChanged

        ' Only LF or CRLF allowed at any one time, not both
        If Dev1TerminatorEnable.Checked Then
            Dev1TerminatorEnable2.Checked = False
        End If

    End Sub

    Private Sub Dev1TerminatorEnable2_CheckedChanged(sender As Object, e As EventArgs) Handles Dev1TerminatorEnable2.CheckedChanged

        ' Only LF or CRLF allowed at any one time, not both
        If Dev1TerminatorEnable2.Checked Then
            Dev1TerminatorEnable.Checked = False
        End If

    End Sub

    Private Sub EditMode_CheckedChanged(sender As Object, e As EventArgs) Handles EditMode.CheckedChanged

        ' Edit mode checkbox on main Device form allows editing of the profiles as long as no Devices are running.
        ' There are inhibits in place to disable the checkbox etc if any Devices are in operation

        If EditMode.Checked = True Then
            gbox1.Enabled = True
            gbox2.Enabled = True
            gbox12.Enabled = True
            ButtonDev1Run.Enabled = False
            ButtonDev2Run.Enabled = False
            btncreate2.Enabled = False
            btncreate3.Enabled = False
            btncreate.Enabled = False
            ButtonDev1PreRun.Enabled = False
            ButtonDev2PreRun.Enabled = False

            btnq1b.Enabled = False
            btnq1a.Enabled = False
            btns1c.Enabled = False
            btnq2b.Enabled = False
            btnq2a.Enabled = False
            btns2c.Enabled = False
            ButtonDev12Run.Enabled = False
        End If

        If EditMode.Checked = False Then
            gbox1.Enabled = False
            gbox2.Enabled = False
            gbox12.Enabled = False

            btncreate.Enabled = True
            btncreate2.Enabled = True
            btncreate3.Enabled = True

            ButtonDev1PreRun.Enabled = True
            ButtonDev2PreRun.Enabled = True

            'gboxdev.Enabled = True

            gbox12.Enabled = False

            btnq1b.Enabled = True
            btnq1a.Enabled = True
            btns1c.Enabled = True
            btnq2b.Enabled = True
            btnq2a.Enabled = True
            btns2c.Enabled = True
            ButtonDev12Run.Enabled = True
        End If


    End Sub


    Private Sub ShowExtendedSerialPortsInfo()
        ' String to hold the detailed port list with an extra line after the header
        Dim portList As String = "Available Serial COM Ports:" & Environment.NewLine & Environment.NewLine

        ' Query WMI for COM port descriptions
        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'")

        ' Iterate through each found port and extract information
        For Each port As ManagementObject In searcher.Get()
            Dim name As String = port("Name").ToString() ' Contains description, e.g., "USB Serial Device (COM13)"
            portList &= $"{name}" & Environment.NewLine
        Next

        ' Show in message box
        If portList = "Available Serial Ports:" & Environment.NewLine & Environment.NewLine Then
            MessageBox.Show("No serial ports available.", "Serial COM Ports", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MessageBox.Show(portList & Environment.NewLine & "(From Device Manager)", "Serial COM Ports", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub


    Private Sub ButtonAvailableComPorts_Click(sender As Object, e As EventArgs) Handles ButtonAvailableComPorts.Click

        ShowExtendedSerialPortsInfo()

    End Sub

    ' Tooltips for Device 1 checkboxes
    ' Event handler for txtname1.TextChanged to update tooltips as necessary
    Private Sub txtname1_TextChanged(sender As Object, e As EventArgs) Handles txtname1.TextChanged
        UpdateCheckedTooltips1()
    End Sub


    ' Method to update tooltips only for checked checkboxes
    Private Sub UpdateCheckedTooltips1()
        If ProfDev1_1.Checked Then ToolTip1.SetToolTip(ProfDev1_1, txtname1.Text)
        If ProfDev1_2.Checked Then ToolTip1.SetToolTip(ProfDev1_2, txtname1.Text)
        If ProfDev1_3.Checked Then ToolTip1.SetToolTip(ProfDev1_3, txtname1.Text)
        If ProfDev1_4.Checked Then ToolTip1.SetToolTip(ProfDev1_4, txtname1.Text)
        If ProfDev1_5.Checked Then ToolTip1.SetToolTip(ProfDev1_5, txtname1.Text)
        If ProfDev1_6.Checked Then ToolTip1.SetToolTip(ProfDev1_6, txtname1.Text)
        If ProfDev1_7.Checked Then ToolTip1.SetToolTip(ProfDev1_7, txtname1.Text)
        If ProfDev1_8.Checked Then ToolTip1.SetToolTip(ProfDev1_8, txtname1.Text)
    End Sub


    ' Tooltips for Device 2 checkboxes
    ' Event handler for txtname1.TextChanged to update tooltips as necessary
    Private Sub txtname2_TextChanged(sender As Object, e As EventArgs) Handles txtname2.TextChanged
        UpdateCheckedTooltips2()
    End Sub


    ' Method to update tooltips only for checked checkboxes
    Private Sub UpdateCheckedTooltips2()
        If ProfDev2_1.Checked Then ToolTip1.SetToolTip(ProfDev2_1, txtname2.Text)
        If ProfDev2_2.Checked Then ToolTip1.SetToolTip(ProfDev2_2, txtname2.Text)
        If ProfDev2_3.Checked Then ToolTip1.SetToolTip(ProfDev2_3, txtname2.Text)
        If ProfDev2_4.Checked Then ToolTip1.SetToolTip(ProfDev2_4, txtname2.Text)
        If ProfDev2_5.Checked Then ToolTip1.SetToolTip(ProfDev2_5, txtname2.Text)
        If ProfDev2_6.Checked Then ToolTip1.SetToolTip(ProfDev2_6, txtname2.Text)
        If ProfDev2_7.Checked Then ToolTip1.SetToolTip(ProfDev2_7, txtname2.Text)
        If ProfDev2_8.Checked Then ToolTip1.SetToolTip(ProfDev2_8, txtname2.Text)
    End Sub


    Private Sub IODeviceLabel1_MouseHover(sender As Object, e As EventArgs) Handles IODeviceLabel1.MouseHover
        Dim allTooltips As New List(Of String)

        ' Collect tooltips from each checkbox
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev1_1))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev1_2))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev1_3))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev1_4))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev1_5))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev1_6))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev1_7))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev1_8))

        ' Set combined tooltips as the tooltip for IODeviceLabel
        ToolTip1.SetToolTip(IODeviceLabel1, String.Join(Environment.NewLine, allTooltips))
    End Sub


    Private Sub IODeviceLabel2_MouseHover(sender As Object, e As EventArgs) Handles IODeviceLabel2.MouseHover
        Dim allTooltips As New List(Of String)

        ' Collect tooltips from each checkbox
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev2_1))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev2_2))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev2_3))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev2_4))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev2_5))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev2_6))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev2_7))
        allTooltips.Add(ToolTip1.GetToolTip(ProfDev2_8))

        ' Set combined tooltips as the tooltip for IODeviceLabel
        ToolTip1.SetToolTip(IODeviceLabel2, String.Join(Environment.NewLine, allTooltips))
    End Sub




    ' Check for program updates and download

    Private Const UpdateInfoUrl As String = "https://www.ianjohnston.com/WinGPIB/WinGPIBupdate.txt"
    Private Const DownloadBaseUrl As String = "https://www.ianjohnston.com/WinGPIB/"

    Private Sub ButtonCheckUpdates_Click(sender As Object, e As EventArgs) Handles ButtonCheckUpdates.Click
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Dim raw As String
            Using wc As New WebClient()
                wc.Headers(HttpRequestHeader.UserAgent) = "WinGPIB-Updater"
                wc.Encoding = Encoding.UTF8
                raw = wc.DownloadString(UpdateInfoUrl)
            End Using

            ' NEW: parse version + multi-line NOTES
            Dim parsed = ParseUpdateTxt(raw)
            Dim latest As Version = parsed.Item1
            Dim notesText As String = parsed.Item2

            Dim current As Version = CurrentVersionFromBanner(BannerText1)

            If latest > current Then
                Dim zipUrl As String = BuildZipUrlFromVersion(latest) ' e.g. .../WinGPIB_V3_284.zip
                Dim sb As New StringBuilder()
                sb.AppendLine($"A new version is available: V{latest.Major}.{latest.Minor}")
                sb.AppendLine($"You have: V{current.Major}.{current.Minor}")

                If Not String.IsNullOrWhiteSpace(notesText) Then
                    sb.AppendLine().AppendLine("Notes:")
                    sb.AppendLine(notesText) ' <- quotes NOTES exactly, multi-line preserved
                End If

                sb.AppendLine().AppendLine("Download ZIP file ?")

                If MessageBox.Show(Me, sb.ToString(), "Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start(New ProcessStartInfo(zipUrl) With {.UseShellExecute = True})
                End If
            Else
                MessageBox.Show(Me, $"You're up to date. (V{current.Major}.{current.Minor})", "Check for Updates", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show(Me, "Couldn't check for updates: " & ex.Message, "Check for Updates", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    ' Extracts version from your banner string (e.g., "WinGPIB   V3.283")
    Private Shared Function CurrentVersionFromBanner(banner As String) As Version
        If String.IsNullOrWhiteSpace(banner) Then Return New Version(0, 0)
        Dim i As Integer = banner.LastIndexOf("V"c)
        Dim s As String = If(i >= 0, banner.Substring(i + 1), banner)
        Return ParseLooseVersion(s)
    End Function

    ' Accepts "3.284", "V3.284", "3.284.0.0", etc.
    Private Shared Function ParseLooseVersion(s As String) As Version
        If String.IsNullOrWhiteSpace(s) Then Return New Version(0, 0)
        s = s.Trim()
        If s.StartsWith("v", StringComparison.OrdinalIgnoreCase) Then s = s.Substring(1)
        s = s.Replace("_", ".")
        Dim parts = s.Split("."c)
        Dim nums As New List(Of Integer)
        For Each p In parts
            Dim n As Integer
            If Integer.TryParse(p, n) Then nums.Add(n) Else nums.Add(0)
        Next
        While nums.Count < 2 : nums.Add(0) : End While
        While nums.Count < 4 : nums.Add(0) : End While
        Return New Version(nums(0), nums(1), nums(2), nums(3))
    End Function

    ' Builds "WinGPIB_V<major>_<minor>.zip"
    Private Shared Function BuildZipUrlFromVersion(ver As Version) As String
        Return $"{DownloadBaseUrl}WinGPIB_V{ver.Major}_{ver.Minor}.zip"
    End Function

    ' Parse version + multi-line NOTES (supports "notes=" or "NOTES=")
    Private Shared Function ParseUpdateTxt(raw As String) As (Version, String)
        If raw Is Nothing Then Return (New Version(0, 0), "")
        Dim text As String = raw.Replace(vbCrLf, vbLf)
        Dim lines = text.Split(vbLf)
        Dim latest As New Version(0, 0)
        Dim notes As New StringBuilder()
        Dim inNotes As Boolean = False

        For Each line In lines
            If line Is Nothing Then Continue For
            If Not inNotes Then
                Dim idx = line.IndexOf("="c)
                If idx > 0 Then
                    Dim key = line.Substring(0, idx).Trim()
                    Dim val = line.Substring(idx + 1) ' keep original spacing/formatting
                    If key.Equals("version", StringComparison.OrdinalIgnoreCase) Then
                        latest = ParseLooseVersion(val)
                    ElseIf key.Equals("notes", StringComparison.OrdinalIgnoreCase) OrElse key.Equals("NOTES", StringComparison.OrdinalIgnoreCase) Then
                        inNotes = True
                        notes.AppendLine(val)
                    End If
                End If
            Else
                ' Everything after the first notes= line is part of notes
                notes.AppendLine(line)
            End If
        Next

        Return (latest, notes.ToString().Trim())
    End Function


    ' SHUTTING DOWN WinGPIB

    Private Sub FormMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        For i As Integer = 1 To 15
            Dim t = DirectCast(Me.GetType().GetField(
                $"Timer{i}",
                Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic
            )?.GetValue(Me), System.Windows.Forms.Timer)

            If t IsNot Nothing Then
                t.Stop()
                t.Enabled = False
            End If
        Next

    End Sub





End Class




