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
'Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.IO
'Imports System
Imports System.IO.Ports
'Imports System.Drawing
Imports System.Management
Imports System.Net
Imports System.Reflection
Imports System.Runtime.InteropServices
'Imports System.Xml.Serialization
'Imports System.Configuration
'Imports System.Text.RegularExpressions
'Imports System.Configuration
'Imports WinGPIBproj.Formtest
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports IODevices
Imports MoonSharp.Interpreter


Public Class Formtest

    ' IODevices form tracker
    Private ioDevicesOffsetInitialized As Boolean = False
    Private ioDevicesOffsetX As Integer
    Private ioDevicesOffsetY As Integer
    ' IODevices tracking – remember last main window position
    Private lastMainLeft As Integer
    Private lastMainTop As Integer
    Private lastMainPosInit As Boolean = False

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowPos(hWnd As IntPtr,
                                         hWndInsertAfter As IntPtr,
                                         X As Integer,
                                         Y As Integer,
                                         cx As Integer,
                                         cy As Integer,
                                         uFlags As UInteger) As Boolean
    End Function
    Private Const SWP_NOSIZE As UInteger = &H1UI
    Private Const SWP_NOZORDER As UInteger = &H4UI

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure


    ' ##########################################################################################################################

    ' q.ResponseAsString replacement var
    Dim respRaw As String
    Dim respNorm As String

    ' Misc.
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

    ' Prevents dropdown->checkbox->dropdown event loops
    Private _suppressDev1Sync As Boolean = False
    Private _suppressDev2Sync As Boolean = False

    ' Selected profile (1..12) for each device (replaces the old checkbox matrix)
    Private _dev1Profile As Integer = 1
    Private _dev2Profile As Integer = 1

    Public Cal3245AinProgress As Boolean = False


    'Private _loading As Boolean = False


    ' SHUTTING DOWN WinGPIB
    Private Sub Formtest_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        ' USER tab
        ResetUsertab()

        ' Timers
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
        Me.Timer15.Stop()
        Me.Timer16.Stop()

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


    'Protected Overrides Sub OnHandleCreated(e As EventArgs)
    'MyBase.OnHandleCreated(e)
    'Me.DoubleBuffered = True
    'End Sub


    'Sub formtest_close()
    '    IODevice.DisposeAll()
    'End Sub

    'Protected Overrides Sub OnLoad(e As EventArgs)
    'MyBase.OnLoad(e) ' Ensure this is called to trigger the Load event
    ' Additional logic
    'End Sub

    Private Sub Formtest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            ' Theme adjustment for Win11, otherwise disabled controls are hardly visible!
            If My.Settings.ThemeSet = True Then
                EnhanceTextBoxBorders(Me)
                MakeButtonsWin10ish(Me)
            End If
            CheckBoxThemeSet.Checked = My.Settings.ThemeSet

            '_loading = True

            'Dim sw As New Stopwatch()
            'sw.Start()

            ' Banner Text animation - See Timer8                                                                                                       Please DONATE if you find this app useful. See the ABOUT tab"
            BannerText1 = "WinGPIB   V4.087"
            BannerText2 = "Non-Commercial Use Only  -  Please DONATE if you find this app useful, see the ABOUT tab"
            Me.Text = BannerText1 & "                                                        " & BannerText2.ToString()


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

            'Move/Copy disable at boot
            ButtonMOVEdev1.Enabled = False
            ButtonCOPYdev1.Enabled = False
            TextBoxMoveCopydev1.Enabled = False
            ButtonMOVEdev2.Enabled = False
            ButtonCOPYdev2.Enabled = False
            TextBoxMoveCopydev2.Enabled = False

            ResetCSV.Enabled = True
            TempHumLogs.Enabled = False

            lstIntf1.Items.Add("VISA")
            lstIntf1.Items.Add("GPIB:ADLink")
            lstIntf1.Items.Add("gpib488.dll")
            lstIntf1.Items.Add("Serial COM port")
            lstIntf1.Items.Add("Prologix Serial")
            lstIntf1.Items.Add("Prologix Ethernet")
            lstIntf1.Items.Add("NI-GPIB-232CT-A")
            'lstIntf1.Items.Add("XYPHRO UsbGpib (USBTMC)")
            lstIntf1.SelectedIndex = 0

            lstIntf2.Items.Add("VISA")
            lstIntf2.Items.Add("GPIB:ADLink")
            lstIntf2.Items.Add("gpib488.dll")
            lstIntf2.Items.Add("Serial COM port")
            lstIntf2.Items.Add("Prologix Serial")
            lstIntf2.Items.Add("Prologix Ethernet")
            lstIntf2.Items.Add("NI-GPIB-232CT-A")
            'lstIntf2.Items.Add("XYPHRO UsbGpib (USBTMC)")
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

            ' Prologix Serial Device - Set DTR behaviour for all Prologix serial devices
            IODevices.PrologixDeviceSerial.PrologixSerialDTREnable = My.Settings.PrologixSerialDTRenable
            CheckBoxPrologixSerialDTR.Checked = My.Settings.PrologixSerialDTRenable

            ' Serial COM Device - Set DTR/RTS behaviour for all serial COM port serial devices
            IODevices.SerialDevice.SerialCOMDTREnable = My.Settings.SerialCOMDTREnable
            CheckBoxSerialCOMDTREnable.Checked = My.Settings.SerialCOMDTREnable
            IODevices.SerialDevice.SerialCOMRTSEnable = My.Settings.SerialCOMRTSEnable
            CheckBoxSerialCOMRTSEnable.Checked = My.Settings.SerialCOMRTSEnable

            ' IODevices window x-y tracker
            CheckBoxIODevicesFormTracker.Checked = My.Settings.IODevicesFormTracker

            ' Recall misc saved data
            CSVfilename.Text = My.Settings.data11
            XaxisPoints.Text = My.Settings.data13
            Dev1Max.Text = My.Settings.data14
            Dev1Min.Text = My.Settings.data15
            txtname3.Text = My.Settings.data16
            Dev12SampleRate.Text = My.Settings.data19
            CSVdelimit = My.Settings.data29
            TempOffset.Text = My.Settings.data78
            TextBoxAvgWindow.Text = My.Settings.data525
            XaxisPoints.Text = My.Settings.data255
            CheckBoxAllowSaveAnytime.Checked = My.Settings.data505
            TextBoxTextEditor.Text = My.Settings.data506
            CheckBoxEnableTooltips.Checked = My.Settings.data507
            CalRam3458APreRun.Text = My.Settings.data508

            ' Recall with X amount of dp's and if no DP saved then recall as X.0
            Dim storedValue As Double = My.Settings.data256
            Dim formattedValue As String
            If storedValue - Math.Floor(storedValue) = 0 Then
                formattedValue = My.Settings.data256.ToString("0.0")
            Else
                formattedValue = My.Settings.data256.ToString()
            End If
            Dev1Max.Text = formattedValue

            ' Recall with X amount of dp's and if no DP saved then recall as X.0
            storedValue = My.Settings.data257
            If storedValue - Math.Floor(storedValue) = 0 Then
                formattedValue = My.Settings.data257.ToString("0.0")
            Else
                formattedValue = My.Settings.data257.ToString()
            End If
            Dev1Min.Text = formattedValue

            ' Recall with X amount of dp's and if no DP saved then recall as X.0
            storedValue = My.Settings.data258
            If storedValue - Math.Floor(storedValue) = 0 Then
                formattedValue = My.Settings.data258.ToString("0.0")
            Else
                formattedValue = My.Settings.data258.ToString()
            End If
            LCTempMax.Text = formattedValue

            ' Recall with X amount of dp's and if no DP saved then recall as X.0
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

            ' Live Chart
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

            ' Temp/Hum
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
            CheckBoxParseLeftRight.Checked = My.Settings.data334    ' temp/hum checkbox
            CheckBoxRegex.Checked = My.Settings.data335             ' temp/hum checkbox
            CheckBoxArithmetic.Checked = My.Settings.data336        ' temp/hum checkbox
            TextBoxTempHumSample.Text = My.Settings.data504

            ' PDVS2mini
            Default0.Text = My.Settings.data260            ' Recall default PDVS2mini counts saved
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
            WryTech.Checked = My.Settings.data503
            comPort_ComboBox.SelectedItem = My.Settings.data333

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

            ButtonR6581upload.Enabled = False
            ButtonR6581commitEEprom.Enabled = False

            If CheckBoxEnableTooltips.Checked = True Then
                ToolTip1.Active = True
            Else
                ToolTip1.Active = False
            End If

            ToolTip1.SetToolTip(cboDev1Device, "Select Dev1 profile")
            ToolTip1.SetToolTip(cboDev2Device, "Select Dev2 profile")

            ' Misc tooltips config
            ToolTip1.OwnerDraw = True
            ToolTip1.InitialDelay = 500   ' ms before first show (default is 1000)
            ToolTip1.ReshowDelay = 5    ' delay when moving between controls
            ToolTip1.AutoPopDelay = 15000 ' how long it stays visible

            Dev1ChartValue.Text = ""
            Dev2ChartValue.Text = ""

            ' 3245A Dev2
            Disable3245AControls()

            ' User tab
            ButtonLoadTxtRefresh.Enabled = False
            ButtonUserStart.Enabled = False

            ' For USER tab keypad
            Me.KeyPreview = True

            ' Dropdown init (profile selectors)
            AddHandler cboDev1Device.SelectedIndexChanged, AddressOf cboDev1Device_SelectedIndexChanged
            AddHandler cboDev2Device.SelectedIndexChanged, AddressOf cboDev2Device_SelectedIndexChanged

            _suppressDev1Sync = True
            _suppressDev2Sync = True

            PopulateDeviceDropdownsFromNames()

            Dim savedDev1 As Integer = GetSavedDev1Profile()
            Dim savedDev2 As Integer = GetSavedDev2Profile()

            If savedDev1 < 1 OrElse savedDev1 > 20 Then savedDev1 = 1
            If savedDev2 < 1 OrElse savedDev2 > 20 Then savedDev2 = 1

            cboDev1Device.SelectedIndex = savedDev1 - 1
            cboDev2Device.SelectedIndex = savedDev2 - 1

            SetDev1SelectedProfile(savedDev1, True)
            SetDev2SelectedProfile(savedDev2, True)

            _suppressDev1Sync = False
            _suppressDev2Sync = False

            MakeButtonsWin10ish(gbox1)
            MakeButtonsWin10ish(GroupBox8)
            MakeButtonsWin10ish(gbox2)
            MakeButtonsWin10ish(GroupBox9)
            MakeButtonsWin10ish(GroupBox12)

        Catch ex As Exception
            MessageBox.Show($"Error during load: {ex.Message}")
        Finally
            '_loading = False
        End Try

    End Sub








    ' SyncDev1DropdownFromCheckbox removed (checkboxes deleted)



    ' SyncDev2DropdownFromCheckbox removed (checkboxes deleted)
    Private Sub cboDev1Device_SelectedIndexChanged(sender As Object, e As EventArgs)

        If _suppressDev1Sync Then Return
        If cboDev1Device.Items.Count = 0 Then Return
        If cboDev1Device.SelectedIndex < 0 Then Return

        SetDev1SelectedProfile(cboDev1Device.SelectedIndex + 1, True)

    End Sub
    Private Sub cboDev2Device_SelectedIndexChanged(sender As Object, e As EventArgs)

        If _suppressDev2Sync Then Return
        If cboDev2Device.Items.Count = 0 Then Return
        If cboDev2Device.SelectedIndex < 0 Then Return

        SetDev2SelectedProfile(cboDev2Device.SelectedIndex + 1, True)

    End Sub

    Private Function Dev1ProfileNames() As String()
        Return New String() {
        My.Settings.data1,
        My.Settings.data1b,
        My.Settings.data1c,
        My.Settings.data139,
        My.Settings.data155,
        My.Settings.data171,
        My.Settings.data349,
        My.Settings.data375,
        My.Settings.data526,
        My.Settings.data555,
        My.Settings.data584,
        My.Settings.data613,
        My.Settings.data758,
        My.Settings.data787,
        My.Settings.data816,
        My.Settings.data845,
        My.Settings.data874,
        My.Settings.data903,
        My.Settings.data932,
        My.Settings.data961
    }
    End Function

    Private Function Dev2ProfileNames() As String()
        Return New String() {
        My.Settings.data2,
        My.Settings.data2b,
        My.Settings.data2c,
        My.Settings.data91,
        My.Settings.data107,
        My.Settings.data123,
        My.Settings.data401,
        My.Settings.data427,
        My.Settings.data642,
        My.Settings.data671,
        My.Settings.data700,
        My.Settings.data729,
        My.Settings.data990,
        My.Settings.data1019,
        My.Settings.data1048,
        My.Settings.data1077,
        My.Settings.data1106,
        My.Settings.data1135,
        My.Settings.data1164,
        My.Settings.data1193
    }
    End Function


    Private Sub PopulateDeviceDropdownsFromNames()

        cboDev1Device.DrawMode = DrawMode.OwnerDrawFixed
        cboDev2Device.DrawMode = DrawMode.OwnerDrawFixed

        AddHandler cboDev1Device.DrawItem, AddressOf ProfileCombo_DrawItem
        AddHandler cboDev2Device.DrawItem, AddressOf ProfileCombo_DrawItem

        cboDev1Device.DropDownStyle = ComboBoxStyle.DropDownList
        cboDev2Device.DropDownStyle = ComboBoxStyle.DropDownList

        cboDev1Device.Items.Clear()
        cboDev2Device.Items.Clear()

        Dim d1 = Dev1ProfileNames()
        For i As Integer = 0 To 19
            'Dim baseName = If(String.IsNullOrWhiteSpace(d1(i)), $"Dev1 Profile {i + 1}", d1(i))
            'Dim baseName = If(String.IsNullOrWhiteSpace(d1(i)), $"Dev1 Profile", d1(i))
            Dim baseName = If(String.IsNullOrWhiteSpace(d1(i)), $"  ", d1(i))
            cboDev1Device.Items.Add($"{i + 1}|{baseName}")
        Next

        Dim d2 = Dev2ProfileNames()
        For i As Integer = 0 To 19
            'Dim baseName = If(String.IsNullOrWhiteSpace(d2(i)), $"Dev2 Profile {i + 1}", d2(i))
            'Dim baseName = If(String.IsNullOrWhiteSpace(d2(i)), $"Dev2 Profile", d2(i))
            Dim baseName = If(String.IsNullOrWhiteSpace(d2(i)), $"  ", d2(i))
            cboDev2Device.Items.Add($"{i + 1}|{baseName}")
        Next

    End Sub

    Private Sub ProfileCombo_DrawItem(sender As Object, e As DrawItemEventArgs)
        e.DrawBackground()
        If e.Index < 0 Then Return

        Dim cb = DirectCast(sender, ComboBox)
        Dim raw As String = cb.Items(e.Index).ToString()

        Dim slotPart As String = raw
        Dim namePart As String = ""

        Dim p = raw.IndexOf("|"c)
        If p >= 0 Then
            slotPart = raw.Substring(0, p).Trim()
            namePart = raw.Substring(p + 1)
        End If

        ' Column layout
        Dim slotWidth As Integer = 20   ' adjust if you want (pixels)
        Dim gap As Integer = 2          ' adjust if you want (pixels)

        Dim rSlot As New Rectangle(e.Bounds.X, e.Bounds.Y, slotWidth, e.Bounds.Height)
        Dim rName As New Rectangle(e.Bounds.X + slotWidth + gap, e.Bounds.Y, e.Bounds.Width - slotWidth - gap, e.Bounds.Height)

        Dim fore As Color = If((e.State And DrawItemState.Selected) = DrawItemState.Selected, SystemColors.HighlightText, cb.ForeColor)

        ' Right-align slot, left-align name (perfect alignment regardless of font)
        TextRenderer.DrawText(e.Graphics, slotPart & ".", cb.Font, rSlot, fore, TextFormatFlags.Right Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix)
        TextRenderer.DrawText(e.Graphics, namePart, cb.Font, rName, fore, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix)

        e.DrawFocusRectangle()
    End Sub















    ' Large tooltips
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


    ' Large tooltips
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
                Case 7 : dev = New XyphroUsbGpibDevice(name, address, True)                 ' not currently implemented
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
            MakeButtonsWin10ish(gbox1)
            MakeButtonsWin10ish(GroupBox8)

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
            txtaddr1.Enabled = False
            lstIntf1.Enabled = False
            cboDev1Device.Enabled = False

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
            MakeButtonsWin10ish(gbox2)
            MakeButtonsWin10ish(GroupBox9)

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
            txtaddr2.Enabled = False
            lstIntf2.Enabled = False
            cboDev2Device.Enabled = False

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
        MakeButtonsWin10ish(gbox12)

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
            MakeButtonsWin10ish(gbox1)
            MakeButtonsWin10ish(GroupBox8)

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
            txtaddr1.Enabled = False
            lstIntf1.Enabled = False
            txtaddr2.Enabled = False
            lstIntf2.Enabled = False
            cboDev1Device.Enabled = False
            cboDev2Device.Enabled = False

        End If


        If dev2 IsNot Nothing Then

            gbox2.Enabled = True
            MakeButtonsWin10ish(gbox2)
            MakeButtonsWin10ish(GroupBox9)

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
        Debug.WriteLine("BLOCKING DetermineQuery: " & result)

        btnq1b.Enabled = True

        s = "blocking command:'" & q.cmd & "'" & vbCrLf

        If result = 0 Then

            txtr1b.Text = respNorm
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
        Debug.WriteLine("BLOCKING DetermineQuery: " & result)


        btnq2b.Enabled = True

        'result = dev2.QueryBlocking(txtq2b.Text & TermStr(), resp, True) 'simpler version with string parameter
        'btnq2b.Enabled = True

        'txtr2a_disp.Text = Format$(Val(txtr2b.Text), "0.0########################")     ' 0 = Digit placeholder. Displays a digit or a zero. # = Digit placeholder. Displays a digit or nothing. So, it's 0,0 bare minimum.

        'If result = 0 Then
        'txtr2b.Text = resp
        'End If

        s = "blocking command:'" & q.cmd & "'" & vbCrLf

        If result = 0 Then

            txtr2b.Text = respNorm
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

            ' Centralize raw + normalized response
            respRaw = q.ResponseAsString
            respNorm = NormalizeNumericResponse(respRaw)

            ' Fast query mode (used by User tab determine etc) - avoid slow processing
            If USERdev1fastquery Then
                Dim source As String

                ' If caller asked for raw text (resptext), give them the untouched instrument string
                If USERdev1rawoutput Then
                    source = If(respRaw, "")
                Else
                    ' Otherwise use the normalized form (locale-safe)
                    source = If(respNorm, "")
                End If

                USERdev1output2 = source            ' USERdev1output2 = raw/normalized instrument response
                USERdev1output = source             ' USERdev1output = what the USER tab will normally "display"
                txtr1a.Text = source                ' IMPORTANT: also update the DEVICES tab textbox so you see the value there

                OutputReceiveddev1 = True
                Exit Sub
            End If

            Dim s As String = "async command:'" & q.cmd & "'" & vbCrLf

            If q.status = 0 And Dev1TextResponse.Checked = False Then

                'inst_value1F = respNorm   ' for PDVS2mini calibration
                inst_value1F = Val(respNorm)

                ' Update CMD line only if it was used
                If CMDlineOp = True Then
                    ' load the command line variable with the response
                    TextBoxDev1CMD.AppendText(Environment.NewLine & ">    " & respNorm & Environment.NewLine)
                    CMDlineOp = False
                End If

                If Dev13457Aseven.Checked = False Then
                    txtr1a.Text = respNorm
                End If

                ' Enable 7th digit mode for 3457A DMM only. Send extra command after TARM SGL to get the additional data
                If Dev13457Aseven.Checked = True Then

                    temp3457A_Dev1 = respNorm

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
                    Dim pattern As String = "[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?"
                    Dim match As Match = Regex.Match(txtr1a.Text, pattern)
                    If match.Success Then
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

                USERdev1output = txtr1a_disp.Text       ' User tab processed/display

            Else
                s &= "error " & q.errcode & vbCrLf
            End If


            ' Response expected is text not numerical
            If q.status = 0 And Dev1TextResponse.Checked = True Then
                txtr1a.Text = respNorm
            Else
                s &= "error " & q.errcode & vbCrLf
            End If


            s &= "thread wait time:" & q.timestart.Subtract(q.timecall).TotalSeconds.ToString() & " s" & vbCrLf

            txtr1astat.Text = s

            'uncomment this to chain on dev2:
            ' dev2.QueryAsync(txtq2a.Text, AddressOf cbdev2, True)

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

            ' Centralize raw + normalized response
            respRaw = q.ResponseAsString
            respNorm = NormalizeNumericResponse(respRaw)

            ' Fast query mode (used by User tab determine etc) - avoid slow processing
            If USERdev2fastquery Then
                Dim source As String

                If USERdev2rawoutput Then
                    source = If(respRaw, "")
                Else
                    source = If(respNorm, "")
                End If

                USERdev2output2 = source
                USERdev2output = source
                txtr2a.Text = source

                OutputReceiveddev2 = True
                Exit Sub
            End If

            Dim s As String = "async command:'" & q.cmd & "'" & vbCrLf

            If q.status = 0 And Dev2TextResponse.Checked = False Then

                ' Update CMD line only if it was used
                If CMDlineOp = True Then
                    ' load the command line variable with the response
                    TextBoxDev2CMD.AppendText(Environment.NewLine & ">    " & respNorm & Environment.NewLine)
                    CMDlineOp = False
                End If


                If Dev23457Aseven.Checked = False Then
                    txtr2a.Text = respNorm
                End If


                ' Enable 7th digit mode for 3457A DMM only. Send extra command after TARM SGL to get the additional data
                If Dev23457Aseven.Checked = True Then

                    temp3457A_Dev2 = respNorm

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
                    Dim pattern As String = "[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?"
                    Dim match As Match = Regex.Match(txtr2a.Text, pattern)
                    If match.Success Then
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

                USERdev2output = txtr2a_disp.Text       ' User tab processed/display

            Else
                s &= "error " & q.errcode & vbCrLf
            End If


            ' Response expected is text not numerical
            If q.status = 0 And Dev2TextResponse.Checked = True Then
                txtr2a.Text = respNorm
            Else
                s &= "error " & q.errcode & vbCrLf
            End If


            s &= "thread wait time:" & q.timestart.Subtract(q.timecall).TotalSeconds.ToString() & " s" & vbCrLf

            txtr2astat.Text = s

            'uncomment this to chain on dev1:
            'dev1.QueryAsync(txtq1a.Text, AddressOf cbdev1, True)

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


    ' Normalize numeric device response to an invariant-style string
    Private Function NormalizeNumericResponse(raw As String) As String

        ' NOTE: This is string based manipulation.
        ' Might have been easier converting to double and manipulating there, but could get rounding issues.

        If String.IsNullOrWhiteSpace(raw) Then Return raw

        Dim trimmed As String = raw.Trim()

        ' Split exponent part (E or e) so we don't touch it
        Dim expIndex As Integer = trimmed.IndexOfAny(New Char() {"e"c, "E"c})
        Dim mainPart As String
        Dim expPart As String = ""

        If expIndex >= 0 Then
            mainPart = trimmed.Substring(0, expIndex)
            expPart = trimmed.Substring(expIndex)   ' includes 'E' and exponent
        Else
            mainPart = trimmed
        End If

        ' Remove spaces in the main numeric part only
        mainPart = mainPart.Replace(" ", "")

        Dim hasDot As Boolean = mainPart.Contains("."c)
        Dim hasComma As Boolean = mainPart.Contains(","c)

        Dim normalizedMain As String = mainPart

        If hasDot AndAlso hasComma Then
            ' Both '.' and ',' present:
            ' - whichever appears last is decimal separator
            ' - the other is thousands/grouping separator
            Dim lastDot As Integer = mainPart.LastIndexOf("."c)
            Dim lastComma As Integer = mainPart.LastIndexOf(","c)

            Dim decimalChar As Char
            Dim groupChar As Char

            If lastDot > lastComma Then
                decimalChar = "."c
                groupChar = ","c
            Else
                decimalChar = ","c
                groupChar = "."c
            End If

            ' Remove grouping char, keep digits identical
            normalizedMain = mainPart.Replace(groupChar.ToString(), "")

            ' If decimal is comma, convert it to dot
            If decimalChar = ","c Then
                normalizedMain = normalizedMain.Replace(","c, "."c)
            End If

        ElseIf hasComma AndAlso Not hasDot Then
            ' Only comma present → treat as decimal separator
            normalizedMain = mainPart.Replace(","c, "."c)

        Else
            ' Only dot or neither → leave as-is
            normalizedMain = mainPart
        End If

        ' Reattach exponent part unchanged
        Return normalizedMain & expPart
    End Function


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

    End Sub

    Private Sub ButtonReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonReset.Click

        ButtonSaveSettings.Enabled = True
        btnBackup.Enabled = True
        btnRestore.Enabled = True
        ButtonAvailableComPorts.Enabled = True
        txtaddr1.Enabled = True
        lstIntf1.Enabled = True
        txtaddr2.Enabled = True
        lstIntf2.Enabled = True

        cboDev1Device.Enabled = True
        cboDev2Device.Enabled = True

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
        MakeButtonsWin10ish(gbox1)
        MakeButtonsWin10ish(GroupBox8)
        gbox2.Enabled = False
        MakeButtonsWin10ish(gbox2)
        MakeButtonsWin10ish(GroupBox9)

        gbox12.Enabled = False
        MakeButtonsWin10ish(gbox12)


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

    Private Sub ShowFiles4_Click(sender As Object, e As EventArgs) Handles ShowFiles4.Click

        Dim basePath As String = CSVfilepath.Text
        Dim devicesPath As String = basePath & "\Devices"

        If Not Directory.Exists(devicesPath) Then
            devicesPath = basePath
        End If

        Process.Start("explorer.exe", String.Format("/n, /e, {0}", devicesPath))

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

            ButtonMOVEdev1.Enabled = True
            ButtonCOPYdev1.Enabled = True
            TextBoxMoveCopydev1.Enabled = True
            ButtonMOVEdev2.Enabled = True
            ButtonCOPYdev2.Enabled = True
            TextBoxMoveCopydev2.Enabled = True



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

            btnq1b.Enabled = True
            btnq1a.Enabled = True
            btns1c.Enabled = True
            btnq2b.Enabled = True
            btnq2a.Enabled = True
            btns2c.Enabled = True
            ButtonDev12Run.Enabled = True

            ButtonMOVEdev1.Enabled = False
            ButtonCOPYdev1.Enabled = False
            TextBoxMoveCopydev1.Enabled = False
            ButtonMOVEdev2.Enabled = False
            ButtonCOPYdev2.Enabled = False
            TextBoxMoveCopydev2.Enabled = False
        End If

        MakeButtonsWin10ish(gbox1)
        MakeButtonsWin10ish(GroupBox8)
        MakeButtonsWin10ish(gbox2)
        MakeButtonsWin10ish(GroupBox9)
        MakeButtonsWin10ish(gbox12)

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
    End Sub


    ' Tooltips for Device 2 checkboxes
    ' Event handler for txtname1.TextChanged to update tooltips as necessary
    Private Sub txtname2_TextChanged(sender As Object, e As EventArgs) Handles txtname2.TextChanged
        UpdateCheckedTooltips2()
    End Sub


    ' Method to update tooltips only for checked checkboxes
    Private Sub UpdateCheckedTooltips2()
    End Sub


    Private Sub IODeviceLabel1_MouseHover(sender As Object, e As EventArgs) Handles IODeviceLabel1.MouseHover
        Dim allTooltips As New List(Of String)

        ' Collect tooltips from each checkbox

        ' Set combined tooltips as the tooltip for IODeviceLabel
        ToolTip1.SetToolTip(IODeviceLabel1, String.Join(Environment.NewLine, allTooltips))
    End Sub


    Private Sub IODeviceLabel2_MouseHover(sender As Object, e As EventArgs) Handles IODeviceLabel2.MouseHover
        Dim allTooltips As New List(Of String)

        ' Collect tooltips from each checkbox

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

            ' Parse version + multi-line NOTES
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


    ' LUA copyright notice
    Private Sub ButtonLUA_Click(sender As Object, e As EventArgs) Handles ButtonLUA.Click

        Using noticeForm As New Form()
            noticeForm.Text = "Third-Party License - MoonSharp"
            noticeForm.Size = New Size(700, 500)
            noticeForm.StartPosition = FormStartPosition.CenterParent
            noticeForm.FormBorderStyle = FormBorderStyle.FixedDialog
            noticeForm.MaximizeBox = False
            noticeForm.MinimizeBox = False
            noticeForm.ShowInTaskbar = False

            ' License text
            Dim rtb As New RichTextBox()
            rtb.Dock = DockStyle.Fill
            rtb.ReadOnly = True
            rtb.BackColor = SystemColors.Window
            rtb.Font = New Font("Consolas", 10)
            rtb.WordWrap = False
            rtb.DetectUrls = True

            AddHandler rtb.LinkClicked, Sub(s, ev)
                                            Try
                                                Process.Start(New ProcessStartInfo(ev.LinkText) With {.UseShellExecute = True})
                                            Catch
                                                ' ignore
                                            End Try
                                        End Sub

            rtb.Text =
                "WinGPIB uses the LUA scripting language plugin, as follows:" & vbCrLf & vbCrLf &
                "MoonSharp Interpreter" & vbCrLf & vbCrLf &
                "Copyright (c) 2014-2016, Marco Mastropaolo" & vbCrLf &
                "All rights reserved." & vbCrLf & vbCrLf &
                "Parts of the string library are based on the KopiLua project: https://github.com/NLua/KopiLua" & vbCrLf &
                "Copyright (c) 2012 LoDC" & vbCrLf & vbCrLf &
                "Visual Studio Code debugger code is based on code from Microsoft vscode-mono-debug project: https://github.com/Microsoft/vscode-mono-debug" & vbCrLf &
                "Copyright (c) Microsoft Corporation - released under MIT license." & vbCrLf & vbCrLf &
                "Remote Debugger icons are from the Eclipse project: https://www.eclipse.org/" & vbCrLf &
                "Copyright of The Eclipse Foundation" & vbCrLf & vbCrLf &
                "The MoonSharp icon is (c) Isaac, 2014-2015" & vbCrLf & vbCrLf &
                "Redistribution and use in source and binary forms, with or without modification, " & vbCrLf &
                "are permitted provided that the following conditions are met:" & vbCrLf & vbCrLf &
                "* Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer." & vbCrLf &
                "* Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer " & vbCrLf &
                "in the documentation and/or other materials provided with the distribution." & vbCrLf &
                "* Neither the name of Marco Mastropaolo nor the names of its contributors may be used to endorse or promote products " & vbCrLf &
                "derived from this software without specific prior written permission." & vbCrLf & vbCrLf &
                "THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS ""AS IS"" AND ANY EXPRESS OR IMPLIED WARRANTIES, " & vbCrLf &
                "INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. " & vbCrLf &
                "IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, " & vbCrLf &
                "OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; " & vbCrLf &
                "OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT " & vbCrLf &
                "(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE."

            ' OK button
            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.DialogResult = DialogResult.OK
            btnOK.AutoSize = True

            noticeForm.AcceptButton = btnOK
            noticeForm.CancelButton = btnOK

            ' Layout so the RichTextBox doesn't cover the button
            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .RowCount = 2,
                .ColumnCount = 1
            }
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            Dim panelButtons As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .AutoSize = True,
                .Padding = New Padding(8)
            }
            panelButtons.Controls.Add(btnOK)

            layout.Controls.Add(rtb, 0, 0)
            layout.Controls.Add(panelButtons, 0, 1)

            noticeForm.Controls.Add(layout)

            ' Show the dialog
            noticeForm.ShowDialog(Me)
        End Using

    End Sub


    Private Sub CheckBoxPrologixSerialDRT_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxPrologixSerialDTR.CheckedChanged

        IODevices.PrologixDeviceSerial.PrologixSerialDTREnable = CheckBoxPrologixSerialDTR.Checked

        Dim state = CheckBoxPrologixSerialDTR.Checked

        IODevices.PrologixDeviceSerial.PrologixSerialDTREnable = state
        My.Settings.PrologixSerialDTRenable = state

    End Sub


    Private Sub CheckBoxSerialCOMDTREnable_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSerialCOMDTREnable.CheckedChanged

        IODevices.SerialDevice.SerialCOMDTREnable = CheckBoxSerialCOMDTREnable.Checked

        Dim state = CheckBoxSerialCOMDTREnable.Checked

        IODevices.SerialDevice.SerialCOMDTREnable = state
        My.Settings.SerialCOMDTREnable = state

    End Sub


    Private Sub CheckBoxSerialCOMRTSEnable_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSerialCOMRTSEnable.CheckedChanged

        IODevices.SerialDevice.SerialCOMRTSEnable = CheckBoxSerialCOMRTSEnable.Checked

        Dim state = CheckBoxSerialCOMRTSEnable.Checked

        IODevices.SerialDevice.SerialCOMRTSEnable = state
        My.Settings.SerialCOMRTSEnable = state

    End Sub


    Private Sub FormMain_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        If Not My.Settings.IODevicesFormTracker Then
            ' Even if tracking is off, keep lastMain* updated
            lastMainLeft = Me.Left
            lastMainTop = Me.Top
            lastMainPosInit = True
            Exit Sub
        End If

        Dim devicesCaption As String = "IO Devices"   ' exact title of the popup

        Dim hWnd = FindWindow(Nothing, devicesCaption)
        If hWnd = IntPtr.Zero Then
            ' Popup not open – just remember current main pos
            lastMainLeft = Me.Left
            lastMainTop = Me.Top
            lastMainPosInit = True
            Exit Sub
        End If

        ' If this is the first time we see a move, just initialise
        If Not lastMainPosInit Then
            lastMainLeft = Me.Left
            lastMainTop = Me.Top
            lastMainPosInit = True
            Exit Sub
        End If

        ' How far did the main form move since last time?
        Dim dx As Integer = Me.Left - lastMainLeft
        Dim dy As Integer = Me.Top - lastMainTop

        ' Get the current popup position
        Dim r As RECT
        If GetWindowRect(hWnd, r) Then
            Dim newX As Integer = r.Left + dx
            Dim newY As Integer = r.Top + dy

            SetWindowPos(hWnd, IntPtr.Zero, newX, newY, 0, 0, SWP_NOSIZE Or SWP_NOZORDER)
        End If

        ' Update stored main position for next move
        lastMainLeft = Me.Left
        lastMainTop = Me.Top
    End Sub


    Private Sub CheckBoxIODevicesFormTracker_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxIODevicesFormTracker.CheckedChanged

        Dim state = CheckBoxIODevicesFormTracker.Checked
        My.Settings.IODevicesFormTracker = state

    End Sub

    Private Function GetSavedDev1Profile() As Integer
        If My.Settings.Dev1Prof1 Then Return 1
        If My.Settings.Dev1Prof2 Then Return 2
        If My.Settings.Dev1Prof3 Then Return 3
        If My.Settings.Dev1Prof4 Then Return 4
        If My.Settings.Dev1Prof5 Then Return 5
        If My.Settings.Dev1Prof6 Then Return 6
        If My.Settings.Dev1Prof7 Then Return 7
        If My.Settings.Dev1Prof8 Then Return 8
        If My.Settings.Dev1Prof9 Then Return 9
        If My.Settings.Dev1Prof10 Then Return 10
        If My.Settings.Dev1Prof11 Then Return 11
        If My.Settings.Dev1Prof12 Then Return 12
        If My.Settings.Dev1Prof13 Then Return 13
        If My.Settings.Dev1Prof14 Then Return 14
        If My.Settings.Dev1Prof15 Then Return 15
        If My.Settings.Dev1Prof16 Then Return 16
        If My.Settings.Dev1Prof17 Then Return 17
        If My.Settings.Dev1Prof18 Then Return 18
        If My.Settings.Dev1Prof19 Then Return 19
        If My.Settings.Dev1Prof20 Then Return 20
        Return 1
    End Function

    Private Function GetSavedDev2Profile() As Integer
        If My.Settings.Dev2Prof1 Then Return 1
        If My.Settings.Dev2Prof2 Then Return 2
        If My.Settings.Dev2Prof3 Then Return 3
        If My.Settings.Dev2Prof4 Then Return 4
        If My.Settings.Dev2Prof5 Then Return 5
        If My.Settings.Dev2Prof6 Then Return 6
        If My.Settings.Dev2Prof7 Then Return 7
        If My.Settings.Dev2Prof8 Then Return 8
        If My.Settings.Dev2Prof9 Then Return 9
        If My.Settings.Dev2Prof10 Then Return 10
        If My.Settings.Dev2Prof11 Then Return 11
        If My.Settings.Dev2Prof12 Then Return 12
        If My.Settings.Dev2Prof13 Then Return 13
        If My.Settings.Dev2Prof14 Then Return 14
        If My.Settings.Dev2Prof15 Then Return 15
        If My.Settings.Dev2Prof16 Then Return 16
        If My.Settings.Dev2Prof17 Then Return 17
        If My.Settings.Dev2Prof18 Then Return 18
        If My.Settings.Dev2Prof19 Then Return 19
        If My.Settings.Dev2Prof20 Then Return 20
        Return 1
    End Function

    Private Function Dev1ProfileNumber() As Integer
        Return _dev1Profile
    End Function

    Private Function Dev2ProfileNumber() As Integer
        Return _dev2Profile
    End Function

    Private Sub SetDev1SelectedProfile(n As Integer, Optional loadNow As Boolean = True)
        If n < 1 OrElse n > 20 Then n = 1
        _dev1Profile = n

        ' Persist selection using existing Settings booleans (keeps compatibility with old config)
        My.Settings.Dev1Prof1 = (n = 1)
        My.Settings.Dev1Prof2 = (n = 2)
        My.Settings.Dev1Prof3 = (n = 3)
        My.Settings.Dev1Prof4 = (n = 4)
        My.Settings.Dev1Prof5 = (n = 5)
        My.Settings.Dev1Prof6 = (n = 6)
        My.Settings.Dev1Prof7 = (n = 7)
        My.Settings.Dev1Prof8 = (n = 8)
        My.Settings.Dev1Prof9 = (n = 9)
        My.Settings.Dev1Prof10 = (n = 10)
        My.Settings.Dev1Prof11 = (n = 11)
        My.Settings.Dev1Prof12 = (n = 12)

        My.Settings.Dev1Prof13 = (n = 13)
        My.Settings.Dev1Prof14 = (n = 14)
        My.Settings.Dev1Prof15 = (n = 15)
        My.Settings.Dev1Prof16 = (n = 16)
        My.Settings.Dev1Prof17 = (n = 17)
        My.Settings.Dev1Prof18 = (n = 18)
        My.Settings.Dev1Prof19 = (n = 19)
        My.Settings.Dev1Prof20 = (n = 20)
        If loadNow Then
            Select Case n
                Case 1 : LoadDev1Profile_1()
                Case 2 : LoadDev1Profile_2()
                Case 3 : LoadDev1Profile_3()
                Case 4 : LoadDev1Profile_4()
                Case 5 : LoadDev1Profile_5()
                Case 6 : LoadDev1Profile_6()
                Case 7 : LoadDev1Profile_7()
                Case 8 : LoadDev1Profile_8()
                Case 9 : LoadDev1Profile_9()
                Case 10 : LoadDev1Profile_10()
                Case 11 : LoadDev1Profile_11()
                Case 12 : LoadDev1Profile_12()
                Case 13 : LoadDev1Profile_13()
                Case 14 : LoadDev1Profile_14()
                Case 15 : LoadDev1Profile_15()
                Case 16 : LoadDev1Profile_16()
                Case 17 : LoadDev1Profile_17()
                Case 18 : LoadDev1Profile_18()
                Case 19 : LoadDev1Profile_19()
                Case 20 : LoadDev1Profile_20()

            End Select
        End If
    End Sub

    Private Sub SetDev2SelectedProfile(n As Integer, Optional loadNow As Boolean = True)
        If n < 1 OrElse n > 20 Then n = 1
        _dev2Profile = n

        My.Settings.Dev2Prof1 = (n = 1)
        My.Settings.Dev2Prof2 = (n = 2)
        My.Settings.Dev2Prof3 = (n = 3)
        My.Settings.Dev2Prof4 = (n = 4)
        My.Settings.Dev2Prof5 = (n = 5)
        My.Settings.Dev2Prof6 = (n = 6)
        My.Settings.Dev2Prof7 = (n = 7)
        My.Settings.Dev2Prof8 = (n = 8)
        My.Settings.Dev2Prof9 = (n = 9)
        My.Settings.Dev2Prof10 = (n = 10)
        My.Settings.Dev2Prof11 = (n = 11)
        My.Settings.Dev2Prof12 = (n = 12)

        My.Settings.Dev2Prof13 = (n = 13)
        My.Settings.Dev2Prof14 = (n = 14)
        My.Settings.Dev2Prof15 = (n = 15)
        My.Settings.Dev2Prof16 = (n = 16)
        My.Settings.Dev2Prof17 = (n = 17)
        My.Settings.Dev2Prof18 = (n = 18)
        My.Settings.Dev2Prof19 = (n = 19)
        My.Settings.Dev2Prof20 = (n = 20)
        If loadNow Then
            Select Case n
                Case 1 : LoadDev2Profile_1()
                Case 2 : LoadDev2Profile_2()
                Case 3 : LoadDev2Profile_3()
                Case 4 : LoadDev2Profile_4()
                Case 5 : LoadDev2Profile_5()
                Case 6 : LoadDev2Profile_6()
                Case 7 : LoadDev2Profile_7()
                Case 8 : LoadDev2Profile_8()
                Case 9 : LoadDev2Profile_9()
                Case 10 : LoadDev2Profile_10()
                Case 11 : LoadDev2Profile_11()
                Case 12 : LoadDev2Profile_12()
                Case 13 : LoadDev2Profile_13()
                Case 14 : LoadDev2Profile_14()
                Case 15 : LoadDev2Profile_15()
                Case 16 : LoadDev2Profile_16()
                Case 17 : LoadDev2Profile_17()
                Case 18 : LoadDev2Profile_18()
                Case 19 : LoadDev2Profile_19()
                Case 20 : LoadDev2Profile_20()

            End Select
        End If
    End Sub




    Private Sub EnhanceTextBoxBorders(root As Control)

        If My.Settings.ThemeSet = True Then
            For Each c As Control In AllControls(root)

                If TypeOf c Is TextBox Then
                    Dim tb = DirectCast(c, TextBox)

                    ' Skip if already wrapped (border panel)
                    If TypeOf tb.Parent Is Panel Then Continue For

                    Dim parent = tb.Parent

                    ' Outer border panel (grey)
                    Dim border As New Panel With {
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
                    .Padding = New Padding(0, 2, 0, 0) ' tweak 1..3 if needed
                }

                    ' TextBox inside
                    tb.BorderStyle = BorderStyle.None
                    tb.Multiline = True
                    tb.Dock = DockStyle.Fill
                    tb.Margin = New Padding(0)

                    ' Re-parent
                    parent.Controls.Add(border)
                    border.BringToFront()
                    border.Controls.Add(inner)
                    inner.Controls.Add(tb)
                End If

            Next
        End If

    End Sub


    Private Sub MakeButtonsWin10ish(root As Control)

        If My.Settings.ThemeSet = True Then

            For Each c As Control In AllControls(root)
                If TypeOf c Is Button Then
                    Dim b = DirectCast(c, Button)

                    b.FlatStyle = FlatStyle.Flat
                    b.UseVisualStyleBackColor = False

                    b.FlatAppearance.BorderSize = 1
                    b.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200)

                    ' NEW: hover / pressed colours (Win10/11 style)
                    b.FlatAppearance.MouseOverBackColor = Color.FromArgb(229, 241, 251)
                    b.FlatAppearance.MouseDownBackColor = Color.FromArgb(204, 228, 247)

                    If b.Enabled Then
                        b.BackColor = Color.White
                        b.ForeColor = Color.Black
                    Else
                        b.BackColor = Color.FromArgb(245, 245, 245)
                        b.ForeColor = Color.FromArgb(80, 80, 80)
                    End If
                End If
            Next

        End If

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


    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged

        If My.Settings.ThemeSet Then
            MakeButtonsWin10ish(TabControl1.SelectedTab)
        End If

    End Sub

End Class




