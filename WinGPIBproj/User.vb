' User customizeable tab

Imports IODevices
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.IO

Partial Class Formtest

    ' Auto-read state
    Private AutoReadDeviceName As String
    Private AutoReadCommand As String
    Private AutoReadResultControl As String
    Dim intervalMs As Integer = 2000

    ' TXT Format:
    ' BUTTON;Caption;Action;Device;CommandOrPrefix;ValueControl;ResultControl
    '
    ' Action:
    '   SEND        = send fixed command
    '   SENDVALUE   = send prefix + value from ValueControl
    '   QUERY       = blocking query -> result stored into ResultControl
    '
    ' Example:
    ' BUTTON;Reset dev1;SEND;dev1;RST;;
    ' BUTTON;Set DCV;SENDVALUE;dev2;APPLY DCV ;Button3245A_VALUE;
    ' BUTTON;Read ID;QUERY;dev1;*IDN?;;TextBoxID

    Private Sub BuildCustomGuiFromText(def As String)

        GroupBoxCustom.Controls.Clear()

        Dim autoY As Integer = 10   ' fallback vertical stacking

        For Each rawLine In def.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            Dim line = rawLine.Trim()
            If line = "" OrElse line.StartsWith(";") Then Continue For

            Dim parts = line.Split(";"c)
            If parts.Length < 2 Then Continue For

            Dim typeStr = parts(0).Trim().ToUpperInvariant()

            Select Case typeStr

            ' ============================
            '   SET GROUPBOX TITLE
            '   GROUPBOX;Some title text
            ' ============================
                Case "GROUPBOX"
                    Dim title = parts(1).Trim()
                    GroupBoxCustom.Text = title   ' <-- your custom GroupBox
                    Continue For

            ' ======================================
            '   CREATE A LABELED TEXTBOX FROM TXT
            '   TEXTBOX;Name;Label text;X;Y
            ' ======================================
                Case "TEXTBOX"
                    If parts.Length < 3 Then Continue For

                    Dim tbName = parts(1).Trim()
                    Dim labelText = parts(2).Trim()

                    Dim x As Integer = 10
                    Dim y As Integer = autoY

                    If parts.Length >= 5 Then
                        Integer.TryParse(parts(3).Trim(), x)
                        Integer.TryParse(parts(4).Trim(), y)
                    End If

                    ' Label
                    Dim lbl As New Label()
                    lbl.Text = labelText
                    lbl.AutoSize = True
                    lbl.Location = New Point(x, y)
                    GroupBoxCustom.Controls.Add(lbl)

                    ' TextBox to the right of label
                    Dim tb As New TextBox()
                    tb.Name = tbName
                    tb.Width = 120
                    tb.Location = New Point(lbl.Right + 5, y - 2)
                    GroupBoxCustom.Controls.Add(tb)

                    ' Only advance autoY if coordinates not explicitly given
                    If parts.Length < 5 Then
                        autoY += Math.Max(lbl.Height, tb.Height) + 5
                    End If

                    Continue For

            ' ======================================
            '   BUTTON LINES
            '   BUTTON;Caption;Action;Device;Command;ValueCtl;ResultCtl;X;Y
            ' ======================================
                Case "BUTTON"
                    If parts.Length < 5 Then Continue For

                    Dim caption = parts(1).Trim()
                    Dim action = parts(2).Trim().ToUpperInvariant()
                    Dim deviceName = parts(3).Trim()
                    Dim commandOrPrefix = parts(4).Trim()
                    Dim valueControlName As String = If(parts.Length > 5, parts(5).Trim(), "")
                    Dim resultControlName As String = If(parts.Length > 6, parts(6).Trim(), "")

                    ' --- Position (X,Y) ---
                    Dim x As Integer = 10
                    Dim y As Integer = autoY

                    If parts.Length >= 9 Then
                        Integer.TryParse(parts(7).Trim(), x)
                        Integer.TryParse(parts(8).Trim(), y)
                    End If

                    Dim b As New Button()
                    b.Text = caption
                    b.Tag = $"{action}|{deviceName}|{commandOrPrefix}|{valueControlName}|{resultControlName}"

                    ' --- Size (Width,Height) ---
                    Dim w As Integer = 160
                    Dim h As Integer = b.Height   ' default button height

                    If parts.Length >= 11 Then
                        Integer.TryParse(parts(9).Trim(), w)
                        Integer.TryParse(parts(10).Trim(), h)
                    End If

                    b.Location = New Point(x, y)
                    b.Size = New Size(w, h)

                    AddHandler b.Click, AddressOf CustomButton_Click
                    GroupBoxCustom.Controls.Add(b)

                    If parts.Length < 9 Then
                        autoY += b.Height + 5      ' uses final height (after any override)
                    End If

                Case "CHECKBOX"
                    ' CHECKBOX;ResultControlName;Caption;FunctionKey;Param;X;Y
                    ' e.g. CHECKBOX;TextBoxResult;Decimal:;FuncDecimal;;430;440
                    '      CHECKBOX;TextBoxResult;Auto 2secs:;FuncAuto;2secs;430;465

                    If parts.Length >= 7 Then

                        Dim resultName = parts(1).Trim()    ' e.g. "TextBoxResult"
                        Dim caption = parts(2).Trim()       ' e.g. "Decimal:"
                        Dim funcKey = parts(3).Trim()       ' e.g. "FuncDecimal" / "FuncAuto"
                        Dim param = parts(4).Trim()         ' e.g. "" or "2secs"

                        Dim x As Integer
                        Dim y As Integer
                        Integer.TryParse(parts(5).Trim(), x)
                        Integer.TryParse(parts(6).Trim(), y)

                        Dim cb As New CheckBox()
                        cb.Text = caption
                        cb.AutoSize = True
                        cb.Location = New Point(x, y)

                        ' Tag encodes: resultName|FUNCXXX|param
                        cb.Tag = resultName & "|" & funcKey.ToUpperInvariant() & "|" & param
                        cb.Name = "Chk_" & resultName & "_" & funcKey

                        GroupBoxCustom.Controls.Add(cb)

                    End If


            End Select

        Next

    End Sub


    Private Sub LoadCustomGuiFromFile(path As String)
        If IO.File.Exists(path) Then
            Dim text = IO.File.ReadAllText(path)
            BuildCustomGuiFromText(text)
        Else
            MessageBox.Show("Custom GUI file not found: " & path)
        End If
    End Sub


    Private Function GetControlByName(name As String) As Control
        If String.IsNullOrWhiteSpace(name) Then Return Nothing

        ' 1) First look inside PanelCustom (TXT-generated stuff)
        Dim matches = GroupBoxCustom.Controls.Find(name, True)
        If matches IsNot Nothing AndAlso matches.Length > 0 Then
            Return matches(0)
        End If

        ' 2) Fallback: search whole form
        matches = Me.Controls.Find(name, True)
        If matches IsNot Nothing AndAlso matches.Length > 0 Then
            Return matches(0)
        End If

        Return Nothing
    End Function


    Private Function GetDeviceByName(name As String) As IODevices.IODevice
        Select Case name.ToLowerInvariant()
            Case "dev1" : Return dev1
            Case "dev2" : Return dev2
            Case Else : Return Nothing
        End Select
    End Function


    Private Function GetCheckboxFor(resultName As String, funcKey As String) As CheckBox
        If String.IsNullOrWhiteSpace(resultName) OrElse String.IsNullOrWhiteSpace(funcKey) Then
            Return Nothing
        End If

        Dim funcKeyUpper = funcKey.ToUpperInvariant()

        For Each c As Control In GroupBoxCustom.Controls
            If TypeOf c Is CheckBox Then
                Dim cb = DirectCast(c, CheckBox)
                Dim tagStr = TryCast(cb.Tag, String)
                If String.IsNullOrEmpty(tagStr) Then Continue For

                Dim parts = tagStr.Split("|"c)
                If parts.Length >= 2 Then
                    Dim tagResult = parts(0)
                    Dim tagFunc = parts(1).ToUpperInvariant()

                    If String.Equals(tagResult, resultName, StringComparison.OrdinalIgnoreCase) AndAlso
                   tagFunc = funcKeyUpper Then
                        Return cb
                    End If
                End If
            End If
        Next

        Return Nothing
    End Function


    Private Sub RunQueryToResult(deviceName As String,
                             commandOrPrefix As String,
                             resultControlName As String)

        Dim target = GetControlByName(resultControlName)
        If target Is Nothing Then
            MessageBox.Show("Result control not found: " & resultControlName)
            Exit Sub
        End If

        ' Resolve device name (dev1, dev2, ...)
        ' IMPORTANT: dev must NOT be Object – it must be the same type as dev1/dev2.
        ' Assuming dev1/dev2 are IODevices.IODevice (or similar), type it accordingly.
        Dim dev As IODevices.IODevice = Nothing

        Select Case deviceName.ToLowerInvariant()
            Case "dev1"
                dev = dev1
            Case "dev2"
                dev = dev2
                ' add more devices here if you have them
        End Select

        If dev Is Nothing Then
            MessageBox.Show("Unknown device: " & deviceName)
            Exit Sub
        End If

        Dim q As IODevices.IOQuery = Nothing
        Dim status As Integer

        Dim fullCmd As String = commandOrPrefix & TermStr2()
        status = dev.QueryBlocking(fullCmd, q, False)   ' now a strongly-typed call

        Dim raw As String = ""
        Dim outText As String = ""

        If status = 0 AndAlso q IsNot Nothing Then
            raw = q.ResponseAsString.Trim()

            ' Optional decimal formatting via FuncDecimal checkbox
            Dim decCb As CheckBox = GetCheckboxFor(resultControlName, "FuncDecimal")

            If decCb IsNot Nothing AndAlso decCb.Checked Then
                Dim d As Double
                If Double.TryParse(raw,
                               Globalization.NumberStyles.Float,
                               Globalization.CultureInfo.InvariantCulture,
                               d) Then
                    outText = d.ToString("0.###############",
                                     Globalization.CultureInfo.InvariantCulture)
                Else
                    outText = raw
                End If
            Else
                outText = raw
            End If

        ElseIf q IsNot Nothing Then
            outText = "ERR " & status & ": " & q.errmsg
        Else
            outText = "ERR " & status & " (no IOQuery)"
        End If

        If TypeOf target Is TextBox Then
            DirectCast(target, TextBox).Text = outText
        ElseIf TypeOf target Is Label Then
            DirectCast(target, Label).Text = outText
        End If
    End Sub



    Private Function GetCheckboxParam(cb As CheckBox) As String
        If cb Is Nothing Then Return ""
        Dim tagStr = TryCast(cb.Tag, String)
        If String.IsNullOrEmpty(tagStr) Then Return ""
        Dim parts = tagStr.Split("|"c)
        If parts.Length >= 3 Then
            Return parts(2)
        Else
            Return ""
        End If
    End Function



    Private Sub CustomButton_Click(sender As Object, e As EventArgs)

        Dim b = DirectCast(sender, Button)
        Dim meta = CStr(b.Tag)
        Dim parts = meta.Split("|"c)
        If parts.Length < 3 Then Exit Sub

        Dim action = parts(0)                 ' SEND / SENDVALUE / QUERY
        Dim deviceName = parts(1)             ' "dev1" / "dev2"
        Dim commandOrPrefix = parts(2)
        Dim valueControlName = If(parts.Length > 3, parts(3), "")
        Dim resultControlName = If(parts.Length > 4, parts(4), "")

        Dim dev = GetDeviceByName(deviceName)
        If dev Is Nothing Then
            MessageBox.Show("Device not available: " & deviceName)
            Exit Sub
        End If

        Select Case action

            Case "SEND"
                dev.SendAsync(commandOrPrefix, True)

            Case "SENDVALUE"
                Dim valCtrl = TryCast(GetControlByName(valueControlName), TextBox)
                If valCtrl Is Nothing Then
                    MessageBox.Show("Value control not found: " & valueControlName)
                    Exit Sub
                End If

                Dim cmd = commandOrPrefix & valCtrl.Text.Trim()
                dev.SendAsync(cmd, True)


            Case "QUERY"

                ' Single immediate query
                RunQueryToResult(deviceName, commandOrPrefix, resultControlName)

                ' Check for an Auto checkbox (FuncAuto) for this result control
                Dim autoCb As CheckBox = GetCheckboxFor(resultControlName, "FuncAuto")

                If autoCb IsNot Nothing AndAlso autoCb.Checked Then
                    ' Default interval: 2000 ms
                    Dim intervalMs As Integer = 2000

                    Dim param As String = GetCheckboxParam(autoCb)   ' e.g. "2secs", "0.5", "0.5secs"
                    If Not String.IsNullOrEmpty(param) Then

                        ' Extract a numeric string allowing digits and decimal point
                        Dim numeric As String = ""
                        For Each ch As Char In param
                            If Char.IsDigit(ch) OrElse ch = "."c Then
                                numeric &= ch
                            End If
                        Next

                        Dim secs As Double
                        If Double.TryParse(numeric,
                       Globalization.NumberStyles.Float,
                       Globalization.CultureInfo.InvariantCulture,
                       secs) AndAlso secs > 0.0R Then

                            intervalMs = CInt(secs * 1000.0R)
                            Timer5.Interval = intervalMs
                        Else
                            ' fallback: keep existing intervalMs (e.g. default 2000ms)
                        End If
                    Else
                        ' no param → leave intervalMs as the default (e.g. 2000ms)
                    End If

                    AutoReadDeviceName = deviceName
                    AutoReadCommand = commandOrPrefix
                    AutoReadResultControl = resultControlName

                    Timer5.Interval = intervalMs
                    Timer5.Enabled = True
                Else
                    ' No auto-read requested - ensure timer is off for this path
                    Timer5.Enabled = False
                End If

        End Select

    End Sub


    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick

        ' If Auto checkbox is no longer present or unchecked, stop auto-read
        Dim autoCb As CheckBox = GetCheckboxFor(AutoReadResultControl, "FuncAuto")
        If autoCb Is Nothing OrElse Not autoCb.Checked Then
            Timer5.Enabled = False
            Exit Sub
        End If

        ' Re-run the last query definition
        If Not String.IsNullOrEmpty(AutoReadDeviceName) AndAlso
       Not String.IsNullOrEmpty(AutoReadCommand) AndAlso
       Not String.IsNullOrEmpty(AutoReadResultControl) Then

            RunQueryToResult(AutoReadDeviceName, AutoReadCommand, AutoReadResultControl)
        Else
            Timer5.Enabled = False
        End If
    End Sub



End Class
