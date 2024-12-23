
' Temperature & Humidity serial interface
' SerialPort

Imports System.IO.Ports
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Microsoft.Office.Interop.Word
Imports MCP2221
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports System.Diagnostics
Imports System.Reflection.Emit
Imports System.Security.Cryptography

Imports System.Management

Partial Class Formtest
    Public DevTemp As String
    Dim TempHumConnected As Boolean = False
    Dim UsbI2c As MCP2221.MchpUsbI2c = New MchpUsbI2c()         ' MCP2221A / SHT40
    Private stopwatchSHT40 As New Stopwatch()

    Private Function GetTemperatureHumiditySHT40(ByVal portName As String) As Double

        'stopwatchSHT40.Start()

        ' This function takes about 50mS to execute. The Dogratian equivalent Functions are <10mS. Apparently the MCP2221A has a 30mS latency?

        Try
            ' Turn LED on
            OnOffLed1.State = OnOffLed.LedState.OnSmall
            OnOffLed1.Invalidate()
            OnOffLed1.Refresh()
            Timer9.Interval = 20 ' Set the interval
            AddHandler Timer9.Tick, AddressOf Timer9_Tick
            Timer9.Start()

            REM Check setting
            ' portName is not actual used with the MCP2221A / SHT series, this sensor is a USB device using VID/PID
            'If portName = vbNullString Then
            'ThisMoment = Now
            'Return Double.NaN
            'End If
            'If portName.Length = 0 Then
            'ThisMoment = Now
            'Return Double.NaN
            'End If

            ' Config Serial Port
            If (Me.SerialPort.IsOpen = False) Then
                Thread.Sleep(100)
                Me.SerialPort.PortName = portName
                Me.SerialPort.Open()
                Me.SerialPort.ReadTimeout = 500
            End If


            ' Make an instance of the MCP2221.MchpUsbI2c class. If using custom VID/PID, use VID and PID as arguments to the constructor.
            'Dim UsbI2c As MCP2221.MchpUsbI2c = New MchpUsbI2c()
            ' Navigate the DLL classes to find your desired function. Other examples are shown below.
            'Dim isConnected As Boolean = UsbI2c.Settings.GetConnectionStatus()

            'Console.WriteLine("The device is connected." + vbNewLine)
            ' Ex. Check for the total number of devices connected. Select first one.
            ' Get total number of devices plugged into PC
            'Dim devCount As Integer
            'devCount = UsbI2c.Management.GetDevCount()
            'Console.WriteLine("There are " + devCount.ToString() + " MCP2221 devices plugged into the PC." + vbNewLine)
            'UsbI2c.Management.SelectDev(0)

            ' Ex. Get USB descriptor string
            'Dim usbDescriptor As String = UsbI2c.Settings.GetUsbStringDescriptor()
            'Console.WriteLine("The USB descriptor string is: " + usbDescriptor + vbNewLine)

            ' Get the current (SRAM) setting of the clock pin divider value
            'Dim rslt As Integer = UsbI2c.Settings.GetClockPinDividerValue(DllConstants.CURRENT_SETTINGS_ONLY)
            'If (rslt > 0) Then
            'Console.WriteLine("The current value of clock pin divider is: " + (1 << rslt).ToString() + vbNewLine)
            'Else
            'Console.WriteLine("Encountered error " + rslt.ToString() + " when getting clock pin divider value.")
            'End If

            'System.Threading.Thread.Sleep(10)

            'If (isConnected = True) Then

            'Dim SHT40_I2C_ADDRESS As Byte = &H88            ' address is 7bit, so use 88 to indicate actual address 44

            Dim SHT40_I2C_ADDRESS As Byte = &H44                    ' Address of SHT40,41,45
            SHT40_I2C_ADDRESS = SHT40_I2C_ADDRESS << 1              ' Binary shift 1 to the left and add 0 at right, thus sending &H88 to enable READ from MCP2221A
            'SHT40_I2C_ADDRESS = (SHT40_I2C_ADDRESS << 1) Or &H1      ' Binary shift 1 to the left and add 1 at right, thus sending &H89 to enable WRITE to MCP2221A
            'Console.WriteLine("Shifted Result: " & SHT40_I2C_ADDRESS.ToString("X"))                '"X" is used as the format specifier in ToString("X") to format the integer as a hexadecimal string

            Dim SHT40_COMMAND As Byte = &HFD
            Dim I2C_SPEED As UInteger = 300 * 1000                  ' 300k. MCP2221 max = 115200, MCP2221A max = 460800.
            Dim numberOfBytesToWrite As UInteger = 1
            Dim setupData As Byte() = {SHT40_COMMAND}
            Dim numberOfBytesToRead As UInteger = 6

            ' Write to SHT40
            Dim SHT40Write As Integer = UsbI2c.Functions.WriteI2cData(SHT40_I2C_ADDRESS, setupData, numberOfBytesToWrite, I2C_SPEED)
            If SHT40Write = 0 Then
                'Console.WriteLine("Command FD written successfully ")
            Else
                Console.WriteLine($"Error writing command FD. Result: {SHT40Write}")
            End If

            System.Threading.Thread.Sleep(10)       ' delay 10mS

            ' Read from SHT40,41,45
            Dim receivedData As Byte() = New Byte(numberOfBytesToRead - 1) {}
            Dim resultRead As Integer = UsbI2c.Functions.ReadI2cData(SHT40_I2C_ADDRESS, receivedData, numberOfBytesToRead, I2C_SPEED)
            If resultRead = 0 Then
                ' Interpret and display data...
                Dim temperatureData As Byte() = {receivedData(1), receivedData(0)}
                Dim humidityData As Byte() = {receivedData(4), receivedData(3)}

                ' Convert temperature data to degC
                Dim temperatureValue As UInt16 = BitConverter.ToUInt16(temperatureData, 0)
                Dim temperatureDegC As Double = -45 + (175 * temperatureValue / 65535)
                'Console.WriteLine($"Temperature: {temperatureDegC} °C")
                gCurrTemp = temperatureDegC

                ' Convert humidity data to percentage relative humidity
                Dim humidityValue As UInt16 = BitConverter.ToUInt16(humidityData, 0)
                Dim humidityPercentRH As Double = -6 + (125 * humidityValue / 65535)
                'Console.WriteLine($"Humidity: {humidityPercentRH} %RH")
                gCurrHumi = humidityPercentRH
            Else
                Console.WriteLine($"Error reading data. Result: {resultRead}")
            End If

            ' Turn Rx LED on
            OnOffLed2.State = OnOffLed.LedState.On
            OnOffLed2.Invalidate()
            OnOffLed2.Refresh()
            Timer10.Interval = 20 ' Set the interval
            AddHandler Timer10.Tick, AddressOf Timer10_Tick
            Timer10.Start()

        Catch
            'Log(Err.Description)
        End Try

        'stopwatchSHT40.Stop()
        'Console.WriteLine($"Elapsed Time: {stopwatchSHT40.ElapsedMilliseconds} milliseconds")
        'stopwatchSHT40.Reset()

    End Function

    ' Function to calculate CRC-8....................Not Used
    Function CalculateCRC8(data As Byte()) As Byte

        Dim crc As Byte = &HFF ' Initialization with 0xFF

        For Each b As Byte In data
            crc = crc Xor b
            For i As Integer = 0 To 7
                If (crc And &H80) <> 0 Then
                    crc = (crc << 1) Xor &H31
                Else
                    crc <<= 1
                End If
            Next
        Next

        Return crc
    End Function


    Private Function GetTemperature(ByVal portName As String) As Double

        Dim str_temp As String
        Dim ret As Double = Double.NaN

        ret = 0
        Try

            ' Turn LED on
            OnOffLed1.State = OnOffLed.LedState.OnSmall
            OnOffLed1.Invalidate()
            OnOffLed1.Refresh()
            Timer9.Interval = 20 ' Set the interval
            AddHandler Timer9.Tick, AddressOf Timer9_Tick
            Timer9.Start()

            REM Check setting
            If portName = vbNullString Then
                ThisMoment = Now
                Return Double.NaN
            End If
            If portName.Length = 0 Then
                ThisMoment = Now
                Return Double.NaN
            End If

            ' Check serial port that was previously saved still exists

            ' Config Serial Port
            If (Me.SerialPort.IsOpen = False) Then
                Thread.Sleep(100)
                Me.SerialPort.PortName = portName
                Me.SerialPort.Open()
                Me.SerialPort.ReadTimeout = 500
            End If

            REM Get Temperature
            Me.SerialPort.ReadExisting()
            Me.SerialPort.Write("GT" + vbCr)
            str_temp = CleanUpString(Me.SerialPort.ReadLine)
            ret = str_temp

        Catch
            'Log(Err.Description)
        End Try

        Return ret

    End Function


    Private Function GetTemperatureUser(ByVal portName As String) As Double

        Dim str_temp As String
        Dim ret As Double = Double.NaN

        ret = 0
        Try

            ' Turn LED on
            OnOffLed1.State = OnOffLed.LedState.OnSmall
            OnOffLed1.Invalidate()
            OnOffLed1.Refresh()
            Timer9.Interval = 20 ' Set the interval
            AddHandler Timer9.Tick, AddressOf Timer9_Tick
            Timer9.Start()

            REM Check setting
            If portName = vbNullString Then
                ThisMoment = Now
                Return Double.NaN
            End If
            If portName.Length = 0 Then
                ThisMoment = Now
                Return Double.NaN
            End If

            REM Config Serial Port
            If (Me.SerialPort.IsOpen = False) Then
                Thread.Sleep(100)
                Me.SerialPort.PortName = portName
                Me.SerialPort.Open()
                Me.SerialPort.ReadTimeout = 500
            End If

            REM Get Temperature
            Me.SerialPort.ReadExisting()
            Me.SerialPort.Write(TextBoxProtocolInput.Text + vbCr)
            str_temp = CleanUpStringUser(Me.SerialPort.ReadLine)
            ret = str_temp

        Catch
            'Log(Err.Description)
        End Try

        Return ret

    End Function


    Private Function GetHumidity(ByVal portName As String) As Double
        Dim str_humi As String
        Dim ret As Double = Double.NaN

        ret = 0
        Try

            ' Turn LED on
            OnOffLed1.State = OnOffLed.LedState.OnSmall
            OnOffLed1.Invalidate()
            OnOffLed1.Refresh()
            Timer9.Interval = 20 ' Set the interval
            AddHandler Timer9.Tick, AddressOf Timer9_Tick
            Timer9.Start()

            REM Check setting
            If portName = vbNullString Then
                ThisMoment = Now
                Return Double.NaN
            End If
            If portName.Length = 0 Then
                ThisMoment = Now
                Return Double.NaN
            End If

            REM Config Serial Port
            If (Me.SerialPort.IsOpen = False) Then
                Thread.Sleep(100)

                Me.SerialPort.PortName = portName
                Me.SerialPort.Open()
                Me.SerialPort.ReadTimeout = 500
            End If

            REM Get Temperature
            Me.SerialPort.ReadExisting()
            Me.SerialPort.Write("GH" + vbCr)
            str_humi = CleanUpString(Me.SerialPort.ReadLine)
            ret = str_humi

        Catch
            'Log(Err.Description)
        End Try

        Return ret
    End Function


    Private Function TestPort(ByVal portName As String, Optional ByVal showError As Boolean = True) As String

        Dim str_info As String
        Dim str_version As String
        Dim str_temp As String
        Dim str_humi As String
        Dim ret As Boolean
        Dim ch As Char

        'Console.WriteLine("TestPort Called")

        REM init var
        ret = False
        str_info = "Unknown"
        str_version = ""
        Try

            REM Check setting
            If portName = vbNullString Then
                ThisMoment = Now
                Return vbNullString
            End If
            If portName.Length = 0 Then
                ThisMoment = Now
                Return vbNullString
            End If

            REM Config Serial Port
            If (Me.SerialPort.IsOpen) Then
                Me.SerialPort.Close()
                Thread.Sleep(100)
            End If

            ' Turn LED on
            OnOffLed1.State = OnOffLed.LedState.OnSmall
            OnOffLed1.Invalidate()
            OnOffLed1.Refresh()
            Timer9.Interval = 20 ' Set the interval
            AddHandler Timer9.Tick, AddressOf Timer9_Tick
            Timer9.Start()


            If (Me.lstIntf3.Text = "USB-User") Then

                ' BAUDRATE
                If Integer.TryParse(TextBoxSerialPortBaud.Text, SerialPort.BaudRate) Then
                    ' Baud rate successfully parsed
                    'MessageBox.Show("testbaud")
                    SerialPort.BaudRate = Val(TextBoxSerialPortBaud.Text)
                    Console.WriteLine("BaudRate = " & Val(TextBoxSerialPortBaud.Text))
                Else
                    ' Handle the case where the user entered an invalid value for BaudRate
                    ButtonEnd_Click(Nothing, EventArgs.Empty)
                    MessageBox.Show("Invalid Baud Rate. Please enter a valid integer value.")
                End If


                ' DATABITS
                If Integer.TryParse(TextBoxSerialPortBits.Text, SerialPort.DataBits) Then
                    ' Data bits successfully parsed
                    SerialPort.DataBits = Val(TextBoxSerialPortBits.Text)
                    Console.WriteLine("DataBits = " & TextBoxSerialPortBits.Text)
                Else
                    ' Handle the case where the user entered an invalid value for DataBits
                    ButtonEnd_Click(Nothing, EventArgs.Empty)
                    MessageBox.Show("Invalid Data Bits. Please enter a valid integer value.")
                End If


                Try
                    ' PARITY - Get the user input from the TextBox
                    Dim userInputParity As String = TextBoxSerialPortParity.Text.Trim()
                    ' Parse the input to set the Parity property
                    Select Case userInputParity.ToUpperInvariant()
                        Case "NONE"
                            SerialPort.Parity = Parity.None
                        Case "ODD"
                            SerialPort.Parity = Parity.Odd
                        Case "EVEN"
                            SerialPort.Parity = Parity.Even
                        Case Else
                            ' Handle the case where the user entered an invalid value
                            ButtonEnd_Click(Nothing, EventArgs.Empty)
                            MessageBox.Show("Invalid Parity. Please enter 'NONE', 'ODD', or 'EVEN'.")
                    End Select
                Catch ex As Exception
                    ' Handle exceptions, such as invalid port settings
                    MessageBox.Show("Error: " & ex.Message)
                End Try
                Console.WriteLine("Parity = " & TextBoxSerialPortParity.Text)


                Try
                    ' STOPBITS - Get the user input from the TextBox
                    Dim userInputStopBits As String = TextBoxSerialPortStop.Text.Trim()
                    ' Parse the input to set the StopBits property
                    Select Case userInputStopBits.ToUpperInvariant()
                        Case "1"
                            SerialPort.StopBits = StopBits.One
                        Case "1.5"
                            SerialPort.StopBits = StopBits.OnePointFive
                        Case "2"
                            SerialPort.StopBits = StopBits.Two
                        Case Else
                            ' Handle the case where the user entered an invalid value
                            ButtonEnd_Click(Nothing, EventArgs.Empty)
                            MessageBox.Show("Invalid Stop Bits. Please enter '1', '1.5', or '2'.")
                    End Select
                Catch ex As Exception
                    ' Handle exceptions, such as invalid port settings
                    MessageBox.Show("Error: " & ex.Message)
                End Try
                Console.WriteLine("StopBits = " & TextBoxSerialPortStop.Text)


                ' HANDSHAKE - Get the user input from the TextBox
                Dim userInputHandshake As String = TextBoxSerialPortHand.Text.Trim()
                Try
                    ' Parse the input to set the Handshake property
                    Select Case userInputHandshake.ToUpperInvariant()
                        Case "NONE"
                            SerialPort.Handshake = Handshake.None
                        Case "XONXOFF"
                            SerialPort.Handshake = Handshake.XOnXOff
                        Case "RTSCTS"
                            SerialPort.Handshake = Handshake.RequestToSend
                        Case "RTSXONXOFF"
                            SerialPort.Handshake = Handshake.RequestToSendXOnXOff
                        Case Else
                            ' Handle the case where the user entered an invalid value
                            ButtonEnd_Click(Nothing, EventArgs.Empty)
                            MessageBox.Show("Invalid Handshake. Please enter 'NONE', 'XONXOFF', 'RTSCTS', or 'RTSXONXOFF'.")
                    End Select
                Catch ex As Exception
                    ' Handle exceptions, such as invalid port settings
                    MessageBox.Show("Error: " & ex.Message)
                End Try
                Console.WriteLine("Handshake = " & TextBoxSerialPortHand.Text)

            End If


            Me.SerialPort.PortName = portName
            Me.SerialPort.Open()
            Me.SerialPort.ReadTimeout = 500


            ' Get info (Dogratian sensors only)
            If (Me.lstIntf3.Text <> "USB-User") Then
                ' Send GI command to get info from sensors and if it's a DogRatian sensor it will reply appropriately
                Me.SerialPort.ReadExisting()
                Me.SerialPort.Write("GI" + vbCr)
                str_info = CleanUpString(Me.SerialPort.ReadLine)

                If (str_info = "USB-TnH LM75") Then
                    REM Get version
                    Me.SerialPort.ReadExisting()
                    Me.SerialPort.Write("GV" + vbCr)
                    str_version = CleanUpString(Me.SerialPort.ReadLine)
                    Console.WriteLine("STR_VERSION: " & str_version)
                    ch = str_version.Chars(0)
                    If (ch = "V") Then
                        ret = True
                    End If
                End If

                If (str_info = "USB-TnH SHT10") Then
                    REM Get version
                    Me.SerialPort.ReadExisting()
                    Me.SerialPort.Write("GV" + vbCr)
                    str_version = CleanUpString(Me.SerialPort.ReadLine)
                    Console.WriteLine("STR_VERSION: " & str_version)
                    ch = str_version.Chars(0)
                    If (ch = "V") Then
                        ret = True
                    End If
                End If

                If (str_info = "USB-TnH SHT30") Then
                    REM Get version
                    Me.SerialPort.ReadExisting()
                    Me.SerialPort.Write("GV" + vbCr)
                    str_version = CleanUpString(Me.SerialPort.ReadLine)
                    Console.WriteLine("STR_VERSION: " & str_version)
                    ch = str_version.Chars(0)
                    If (ch = "V") Then
                        ret = True
                    End If
                End If

                If (str_info = "USB-PA") Then       ' BME-280
                    REM Get version
                    Me.SerialPort.ReadExisting()
                    Me.SerialPort.Write("GV" + vbCr)
                    str_version = CleanUpString(Me.SerialPort.ReadLine)
                    Console.WriteLine("STR_VERSION: " & str_version)
                    ch = str_version.Chars(0)
                    If (ch = "V") Then
                        ret = True
                    End If
                End If

                REM Get Temperature
                If (ret = True) Then
                    Me.SerialPort.ReadExisting()
                    Me.SerialPort.Write("GT" + vbCr)
                    str_temp = CleanUpString(Me.SerialPort.ReadLine)

                    ' incorprate offset
                    str_temp = str_temp + Val(TempOffset.Text)

                    gCurrTemp = str_temp
                End If

                REM Get Humidity
                If (ret = True) And (gWithHumi) Then
                    Me.SerialPort.ReadExisting()
                    Me.SerialPort.Write("GH" + vbCr)
                    str_humi = CleanUpString(Me.SerialPort.ReadLine)
                    gCurrHumi = str_humi
                End If

                Console.WriteLine("str_info =" & str_info)
                Console.WriteLine("str_version =" & str_version)
                ' End Get info (Dogratian sensors only)
            End If

            'Me.SerialPort.Close()

        Catch
            If (showError) Then
                'Log(Err.Description)
            End If
        End Try

        If (ret = True) Then
            Return (str_info & " " & str_version)
        Else
            Return ""
        End If
    End Function


    Private Sub ButtonRefreshPorts_Click(sender As Object, e As EventArgs) Handles ButtonRefreshPorts.Click

        ' The original method of clearing the list and refreshing the available COM ports didn't work properly if a device was used, then stopped, then unplugged, i.e.
        ' after hitting refresh it would still be listed and the only way to clear it was to restart WinGPIB. This method uses WMI and is more complex......but it works.

        ' Set the selected item to Nothing (empty) or an empty string
        Me.ComboBoxPort.SelectedItem = Nothing

        ' Close and dispose of the current serial port if it's open
        If SerialPort.IsOpen Then
            Try
                SerialPort.Close()
                ComboBoxPort.Enabled = True
            Catch ex As Exception
                ' Handle the exception as needed (log or display an error message)
            End Try
        End If

        SerialPort.Dispose()

        ' Clear existing items in the ComboBox
        Me.ComboBoxPort.Items.Clear()

        ' Use WMI to query the list of serial ports
        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'")
        Dim collection As ManagementObjectCollection = searcher.Get()

        ' Add unique ports to the ComboBox
        For Each device As ManagementObject In collection
            Dim name As String = TryCast(device("Name"), String)
            If name IsNot Nothing Then
                Dim portNameStart As Integer = name.IndexOf("(COM", StringComparison.OrdinalIgnoreCase)
                If portNameStart >= 0 Then
                    Dim portName As String = name.Substring(portNameStart)
                    portName = portName.Replace("(", "").Replace(")", "")               ' Remove parentheses
                    If Not Me.ComboBoxPort.Items.Contains(portName) Then
                        Me.ComboBoxPort.Items.Add(portName)
                    End If
                End If
            End If
        Next

        ' Set the selected item to Nothing (empty) or an empty string
        Me.ComboBoxPort.SelectedItem = Nothing

    End Sub


    Private Sub TextBoxTempHumSample_TextChanged(sender As Object, e As EventArgs) Handles TextBoxTempHumSample.TextChanged
        ValidateFloatingPointInput(TextBoxTempHumSample)
    End Sub


    Private Sub ValidateFloatingPointInput(textBox As TextBox)
        Dim regex As New Regex("^\d*\.?\d*$") ' This regex allows digits and a single decimal point
        If Not regex.IsMatch(textBox.Text) Then
            MessageBox.Show("Please enter a valid number.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            textBox.Text = Regex.Replace(textBox.Text, "[^\d.]", "") ' Remove non-digit and non-decimal point characters
            textBox.SelectionStart = textBox.Text.Length ' Set the cursor to the end of the text
        End If
    End Sub


    Private Sub ButtonStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStart.Click

        If ComboBoxPort.SelectedItem Is Nothing Or ComboBoxPort.SelectedItem = "" Then
            ' The port is not available, handle this situation accordingly
            MessageBox.Show($"No port is selected, please select one.{vbCrLf}Perhaps the previously saved port is not{vbCrLf}available, try hit Refresh and select a new port.")
            Exit Sub
        End If

        gCurrTemp = 0.0
        gCurrHumi = 0.0

        TempOffset.Enabled = False

        DevTemp = Me.lstIntf3.Text

        Me.Timer1.Interval = Val(TextBoxTempHumSample.Text) * 1000      ' mS
        Me.Timer1.Start()

        TempHumLogs.Enabled = True
        EnableChart3.Enabled = True

        Dim str As String

        Console.WriteLine("4: " & Me.lstIntf3.Text)

        If (DevTemp = "USB-TnH SHT10 V2.00") Then
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " USB-TnH SHT10 V2.00 FOUND AT " & Me.ComboBoxPort.Text)
            UpdateTempHumidityUSBTnHSHT10V200()
        End If

        If (DevTemp = "USB-TnH (SHT10)") Then
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " USB-TnH SHT10 FOUND AT " & Me.ComboBoxPort.Text)
            UpdateTempHumidityUSBTnHSHT10V200()
        End If

        If (DevTemp = "USB-PA (BME280)") Then
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " USB-PA (BME280) FOUND AT " & Me.ComboBoxPort.Text)
            UpdateTempHumidityUSBBME280()
        End If

        If (DevTemp = "USB-TnH (LM75)") Then
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " USB-TnH (LM75) FOUND AT " & Me.ComboBoxPort.Text)
            UpdateTempHumidityUSBLM75()
        End If

        If (DevTemp = "USB-TnH (SHT30)") Then
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " USB-TnH (SHT30) FOUND AT " & Me.ComboBoxPort.Text)
            UpdateTempHumidityUSBTnHSHT30()
        End If

        ' User based protocol
        If (DevTemp = "USB-User") Then
            str = TestPort(Me.ComboBoxPort.Text)
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " USB-User TEMP SENSOR FOUND AT " & Me.ComboBoxPort.Text)
            UpdateTempUSBUser()
        End If

        ' Adafruit SHT40,41,45      (TestPort not used)
        If (DevTemp = "Adafruit (MCP2221A/SHT40,41,45)") Then
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " MCP2221A/SHT40,41,45 FOUND AT " & Me.ComboBoxPort.Text)
            GetTemperatureHumiditySHT40(Me.ComboBoxPort.Text)
            'UpdateTempHumiUSBSHT40()
        End If

        Me.ComboBoxPort.Enabled = False
        Me.ButtonStart.Enabled = False
        Me.ButtonEnd.Enabled = True
        Me.lstIntf3.Enabled = False

        CheckBoxTempHide.Enabled = True

    End Sub


    Private Sub ButtonEnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEnd.Click

        TempOffset.Enabled = True

        TempHumLogs.Checked = False
        TempHumLogs.Enabled = False

        TempHumConnected = False

        LabelTemperature.Text = "0.00"
        LabelHumidity.Text = "0.00"

        LabelTemperature2.Text = "0.00"

        Me.ComboBoxPort.Enabled = True
        Me.ButtonStart.Enabled = True
        Me.ButtonEnd.Enabled = False
        Me.lstIntf3.Enabled = True

        ' Close port
        If Me.SerialPort.IsOpen = True Then
            Log(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " COM PORT CLOSED " & Me.ComboBoxPort.Text)
            Me.SerialPort.Close()
        End If

        Me.Timer1.Stop()

        EnableChart3.Checked = False
        EnableChart3.Enabled = False
        CheckBoxTempHide.Checked = False
        CheckBoxTempHide.Enabled = False

        Timer9.Stop()
        RemoveHandler Timer9.Tick, AddressOf Timer9_Tick
        OnOffLed1.State = OnOffLed.LedState.OffSmallBlack

        Timer10.Stop()
        RemoveHandler Timer10.Tick, AddressOf Timer10_Tick
        OnOffLed2.State = OnOffLed.LedState.OffSmallBlack

    End Sub

    ' DogRatian - USB-LM75 Sensor
    ' Spec TBA

    Private Sub UpdateTempHumidityUSBLM75()
        Dim str_temp As String
        Dim str_humi As String

        REM Update Temperature
        If gCurrTemp = Double.NaN Then
            str_temp = ""
        Else
            str_temp = String.Format("{0:F1}", gCurrTemp)
        End If

        ' incorprate offset
        str_temp = str_temp + Val(TempOffset.Text)

        LabelTemperature.Text = Format(Val(str_temp), "#0.0")   ' round to 1dp, # means any amount of chars, 0 means position must be filled
        LabelTemperature2.Text = Format(Val(str_temp), "#0.0")   ' round to 1dp, # means any amount of chars, 0 means position must be filled

        REM Update Humidity
        If gCurrHumi = Double.NaN Then
            str_humi = ""
        Else
            str_humi = String.Format("{0:F1}", gCurrHumi)
        End If

        LabelHumidity.Text = Format(Val(str_humi), "#0.0")   ' round to 1dp, # means any amount of chars, 0 means position must be filled

        TempHumConnected = True

    End Sub

    ' DogRatian - USB-PA-BME280 Usb Temperature & Humidity Sensor
    ' Temperature Accuracy: Operating Temperature(CPU) : 	-20°C ~ +70°C
    ' Connection to PC:	USB (CDC Virtual COM port)
    ' Power:	USB Powered
    ' Pressure Accuracy : 
    ' (0-65°C, 300 to 1100 hPa) 	
    ' +/-1 hPa (+/-100 Pa)
    ' Pressure Range : 	300 - 1100 hPa (30000 - 110000 Pa)

    Private Sub UpdateTempHumidityUSBBME280()
        Dim str_temp As String
        Dim str_humi As String

        REM Update Temperature
        If gCurrTemp = Double.NaN Then
            str_temp = ""
        Else
            str_temp = String.Format("{0:F1}", gCurrTemp)
        End If

        ' incorprate offset
        str_temp = str_temp + Val(TempOffset.Text)

        LabelTemperature.Text = Format(Val(str_temp), "#0.0")   ' round to 1dp, # means any amount of chars, 0 means position must be filled
        LabelTemperature2.Text = Format(Val(str_temp), "#0.0")   ' round to 1dp, # means any amount of chars, 0 means position must be filled

        REM Update Humidity
        If gCurrHumi = Double.NaN Then
            str_humi = ""
        Else
            str_humi = String.Format("{0:F1}", gCurrHumi)
        End If

        LabelHumidity.Text = Format(Val(str_humi), "#0.0")   ' round to 1dp, # means any amount of chars, 0 means position must be filled

        TempHumConnected = True

    End Sub

    ' DogRatian - USB-TnH SHT10 V2.00 Usb Temperature & Humidity Sensor
    ' Temperature Accuracy: 	+/-0.5°C at 25°C (max +/-3°C at 120°C)
    ' Temperature Resolution :  	14bits
    ' Temperature Range :  	-40°C to 120°C
    ' Humidity Accuracy :  	+/-4.5 %RH
    ' Humidity Resolution :  	12bits
    ' Humidity Range :  	5 %RH - 95 %RH

    Private Sub UpdateTempHumidityUSBTnHSHT10V200()
        Dim str_temp As String
        Dim str_humi As String

        REM Update Temperature
        If gCurrTemp = Double.NaN Then
            str_temp = ""
        Else
            str_temp = String.Format("{0:F1}", gCurrTemp)
        End If

        ' incorprate offset
        str_temp = str_temp + Val(TempOffset.Text)

        LabelTemperature.Text = Format(Val(str_temp), "#0.0")   ' round to 1dp, # means any amount of chars, 0 means position must be filled
        LabelTemperature2.Text = Format(Val(str_temp), "#0.0")   ' round to 1dp, # means any amount of chars, 0 means position must be filled

        REM Update Humidity
        If gCurrHumi = Double.NaN Then
            str_humi = ""
        Else
            str_humi = String.Format("{0:F1}", gCurrHumi)
        End If

        LabelHumidity.Text = Format(Val(str_humi), "#0.0")   ' round to 1dp, # means any amount of chars, 0 means position must be filled

        TempHumConnected = True

    End Sub

    ' DogRatian - USB-TnH SHT30 Usb Temperature & Humidity Sensor - 0.01 resolution
    ' Operating Temperature(CPU) :   	 -20°C ~ +70°C 
    ' Connection to PC:  	 USB (CDC Virtual COM port) 
    ' Power:  	 USB Powered 
    ' Temperature Accuracy :   	 +/-0.2°C (0°C to 65°C 
    ' Temperature Resolution :   	 0.01°C 
    ' Temperature Range :   	 -40°C to 125°C 
    ' Humidity Accuracy :   	 +/-2 %RH 
    ' Humidity Resolution :   	 0.01 %RH 
    ' Humidity Range :   	 0 %RH - 100 %RH 

    Private Sub UpdateTempHumidityUSBTnHSHT30()
        Dim str_temp As String
        Dim str_humi As String

        REM Update Temperature
        If gCurrTemp = Double.NaN Then
            str_temp = ""
        Else
            str_temp = String.Format("{0:F2}", gCurrTemp)
        End If

        ' incorprate offset
        str_temp = str_temp + Val(TempOffset.Text)

        LabelTemperature.Text = Format(Val(str_temp), "#0.00")   ' round to 2dp's, # means any amount of chars, 0 means position must be filled
        LabelTemperature2.Text = Format(Val(str_temp), "#0.00")   ' round to 2dp's, # means any amount of chars, 0 means position must be filled

        REM Update Humidity
        If gCurrHumi = Double.NaN Then
            str_humi = ""
        Else
            str_humi = String.Format("{0:F2}", gCurrHumi)
        End If

        LabelHumidity.Text = Format(Val(str_humi), "#0.00")   ' round to 2dp's, # means any amount of chars, 0 means position must be filled

        TempHumConnected = True

    End Sub

    ' User based manual protocol

    Private Sub UpdateTempUSBUser()
        Dim str_temp As String
        'Dim str_humi As String

        REM Update Temperature
        If gCurrTemp = Double.NaN Then
            str_temp = ""
        Else
            str_temp = String.Format("{0:F1}", gCurrTemp)
        End If

        ' incorprate offset
        str_temp = str_temp + Val(TempOffset.Text)

        LabelTemperature.Text = Format(Val(str_temp), "#0.00")   ' round to 2dp's, # means any amount of chars, 0 means position must be filled
        LabelTemperature2.Text = Format(Val(str_temp), "#0.00")   ' round to 2dp's, # means any amount of chars, 0 means position must be filled

        TempHumConnected = True

    End Sub

    ' Adafruit MCP2221A Breakout - General Purpose USB To GPIO ADC I2C - Stemma QT / Qwiic
    ' http://adafru.it/4471
    ' Adafruit Sensirion SHT40, SHT41, SHT45 Temperature & Humidity Sensor - STEMMA QT / Qwiic
    ' http://adafru.it/4885
    ' STEMMA QT / Qwiic JST SH 4-pin Cable - 100mm Long
    ' http://adafru.it/4210

    Private Sub UpdateTempHumiUSBSHT40()
        Dim str_temp As String
        Dim str_humi As String

        REM Update Temperature
        If gCurrTemp = Double.NaN Then
            str_temp = ""
        Else
            'str_temp = String.Format("{0:F2}", gCurrTemp)
            str_temp = gCurrTemp
            str_temp = str_temp + Val(TempOffset.Text)        ' incorprate offset
            LabelTemperature.Text = Format(Val(str_temp), "#0.00")   ' round to 2dp's, # means any amount of chars, 0 means position must be filled
            LabelTemperature2.Text = Format(Val(str_temp), "#0.00")   ' round to 2dp's, # means any amount of chars, 0 means position must be filled
        End If

        REM Update Humidity
        If gCurrHumi = Double.NaN Then
            str_humi = ""
        Else
            str_humi = gCurrHumi
            LabelHumidity.Text = Format(Val(str_humi), "#0.00")   ' round to 2dp's, # means any amount of chars, 0 means position must be filled
        End If

        TempHumConnected = True

    End Sub

    Private Function CleanUpString(ByVal str As String) As String

        ' Turn Rx LED on
        OnOffLed2.State = OnOffLed.LedState.On
        OnOffLed2.Invalidate()
        OnOffLed2.Refresh()
        Timer10.Interval = 50 ' Set the interval
        AddHandler Timer10.Tick, AddressOf Timer10_Tick
        Timer10.Start()

        str = str.Replace(vbLf, " ")
        str = str.Replace(vbCr, " ")
        str = str.Trim()
        Return str

    End Function

    Private Function CleanUpStringUser(ByVal str As String) As String

        ' Turn Rx LED on
        OnOffLed2.State = OnOffLed.LedState.On
        OnOffLed2.Invalidate()
        OnOffLed2.Refresh()
        Timer10.Interval = 50 ' Set the interval
        AddHandler Timer10.Tick, AddressOf Timer10_Tick
        Timer10.Start()

        'str = "Temp22.6degC"            ' Use this for testing only. Example LEFT=4, RIGHT=7

        TextBoxResult.Text = str

        If CheckBoxParseLeftRight.Checked = True Then

            ' Method 2 - Assuming TextBoxParseLeft and TextBoxParseRight contain the user-specified positions
            Dim startPosition As Integer = Integer.Parse(TextBoxParseLeft.Text)
            Dim endPosition As Integer = Integer.Parse(TextBoxParseRight.Text)
            If startPosition >= 0 AndAlso endPosition >= startPosition AndAlso endPosition < TextBoxResult.Text.Length Then
                ' Extract the substring based on user-specified positions
                Dim extractedValue As String = TextBoxResult.Text.Substring(startPosition, endPosition - startPosition + 1)
                str = extractedValue
            End If

        End If

        If CheckBoxRegex.Checked = True Then

            ' Method 1 - Define a regular expression pattern to match numbers with optional decimal points
            'Dim pattern As String = "(\d+(\.\d+)?)"
            Dim pattern As String = TextBoxRegex.Text
            ' Match the pattern in the input string
            Dim match As Match = Regex.Match(TextBoxResult.Text, pattern)
            ' Check if a match is found
            If match.Success Then
                ' Extract the matched numeric value
                str = match.Value
            Else
                ' No numeric value found
                str = "no value found"
            End If

        End If

        ' Arithmentic operations
        If CheckBoxArithmetic.Checked = True Then
            Dim inputValue As Double
            If Double.TryParse(str, inputValue) Then
                ' Split the input based on commas
                Dim operations As String() = TextBoxTempArithmentic.Text.Split(","c)

                ' Apply each operation to the input value
                For Each operation In operations
                    ' Trim any leading or trailing whitespaces
                    operation = operation.Trim()

                    ' Check if the operation is not empty
                    If Not String.IsNullOrEmpty(operation) Then
                        ' Evaluate the operation and update the input value
                        If operation.StartsWith("+") Then
                            Dim addValue As Double
                            If Double.TryParse(operation.Substring(1), addValue) Then
                                inputValue += addValue
                            End If
                        ElseIf operation.StartsWith("-") Then
                            Dim subtractValue As Double
                            If Double.TryParse(operation.Substring(1), subtractValue) Then
                                inputValue -= subtractValue
                            End If
                        ElseIf operation.StartsWith("*") Then
                            Dim multiplyValue As Double
                            If Double.TryParse(operation.Substring(1), multiplyValue) Then
                                inputValue *= multiplyValue
                            End If
                        ElseIf operation.StartsWith("/") Then
                            Dim divideValue As Double
                            If Double.TryParse(operation.Substring(1), divideValue) AndAlso divideValue <> 0 Then
                                inputValue /= divideValue
                            End If
                        End If
                    End If
                Next

                str = inputValue.ToString()
            End If
        End If


        TextBoxFinalTempValue.Text = str
        Return str
    End Function

    Private Sub lstIntf3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstIntf3.SelectedIndexChanged

        ' Set the USB-User textboxes read only if not currently selected on the pulldown

        If Me.lstIntf3.Text = "USB-User" Then

            TextBoxProtocolInput.ReadOnly = False
            TextBoxParseLeft.ReadOnly = False
            TextBoxParseRight.ReadOnly = False
            TextBoxRegex.ReadOnly = False
            TextBoxTempArithmentic.ReadOnly = False
            TextBoxSerialPortBaud.ReadOnly = False
            TextBoxSerialPortBits.ReadOnly = False
            TextBoxSerialPortParity.ReadOnly = False
            TextBoxSerialPortStop.ReadOnly = False
            TextBoxSerialPortHand.ReadOnly = False
            CheckBoxParseLeftRight.Enabled = True
            CheckBoxRegex.Enabled = True
            CheckBoxArithmetic.Enabled = True

        Else

            TextBoxProtocolInput.ReadOnly = True
            TextBoxParseLeft.ReadOnly = True
            TextBoxParseRight.ReadOnly = True
            TextBoxRegex.ReadOnly = True
            TextBoxTempArithmentic.ReadOnly = True
            TextBoxSerialPortBaud.ReadOnly = True
            TextBoxSerialPortBits.ReadOnly = True
            TextBoxSerialPortParity.ReadOnly = True
            TextBoxSerialPortStop.ReadOnly = True
            TextBoxSerialPortHand.ReadOnly = True
            CheckBoxParseLeftRight.Enabled = False
            CheckBoxRegex.Enabled = False
            CheckBoxArithmetic.Enabled = False
            TextBoxResult.Text = ""
            TextBoxFinalTempValue.Text = ""

        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        ' Get new readings for Temperature and Humidity every 1sec, provided its not the User protocol selected
        If (DevTemp <> "USB-User" And DevTemp <> "Adafruit (MCP2221A/SHT40,41,45)") Then
            gCurrTemp = GetTemperature(Me.ComboBoxPort.Text)
            gCurrHumi = GetHumidity(Me.ComboBoxPort.Text)
        End If

        If (DevTemp = "USB-TnH SHT10 V2.00") Or (DevTemp = "USB-TnH (SHT10)") Or (DevTemp = "USB-TnH SHT10 V2.00") Then
            UpdateTempHumidityUSBTnHSHT10V200()
        End If

        If (DevTemp = "USB-TnH (SHT30)") Then
            UpdateTempHumidityUSBTnHSHT30()
        End If

        If (DevTemp = "USB-PA (BME280)") Then
            UpdateTempHumidityUSBBME280()
        End If

        If (DevTemp = "USB-TnH (LM75)") Then
            UpdateTempHumidityUSBLM75()
        End If

        ' User based protocol
        If (DevTemp = "USB-User") Then
            gCurrTemp = GetTemperatureUser(Me.ComboBoxPort.Text)
            UpdateTempUSBUser()
        End If

        ' Adafruit MCP2221A / SHT40 combo
        If (DevTemp = "Adafruit (MCP2221A/SHT40,41,45)") Then
            GetTemperatureHumiditySHT40(Me.ComboBoxPort.Text)
            UpdateTempHumiUSBSHT40()
        End If

    End Sub

    Private Sub Timer9_Tick(sender As Object, e As EventArgs)

        ' Stop Timer9 and turn off the indicator after 0.1 seconds
        Timer9.Stop()
        RemoveHandler Timer9.Tick, AddressOf Timer9_Tick
        OnOffLed1.State = OnOffLed.LedState.OffSmallBlack
        OnOffLed1.Invalidate()
        OnOffLed1.Refresh()

    End Sub


    Private Sub Timer10_Tick(sender As Object, e As EventArgs)
        ' Stop Timer10 and turn off the indicator after 0.1 seconds
        Timer10.Stop()
        RemoveHandler Timer10.Tick, AddressOf Timer10_Tick
        OnOffLed2.State = OnOffLed.LedState.OffSmallBlack
        OnOffLed2.Invalidate()
        OnOffLed2.Refresh()
    End Sub


End Class