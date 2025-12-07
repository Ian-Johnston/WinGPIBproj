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
    Private originalCustomControls As List(Of Control)


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
                        ParseIntField(parts(3), x)
                        ParseIntField(parts(4), y)
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
                        ParseIntField(parts(7), x)
                        ParseIntField(parts(8), y)
                    End If

                    Dim b As New Button()
                    b.Text = caption
                    b.Tag = $"{action}|{deviceName}|{commandOrPrefix}|{valueControlName}|{resultControlName}"

                    ' --- Size (Width,Height) ---
                    Dim w As Integer = 160
                    Dim h As Integer = b.Height   ' default button height

                    If parts.Length >= 11 Then
                        ParseIntField(parts(9), w)
                        ParseIntField(parts(10), h)
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

                        Dim x As Integer = 0
                        Dim y As Integer = 0
                        If parts.Length > 5 Then ParseIntField(parts(5), x)
                        If parts.Length > 6 Then ParseIntField(parts(6), y)

                        Dim cb As New CheckBox()
                        cb.Text = caption
                        cb.AutoSize = True
                        cb.Location = New Point(x, y)

                        ' Tag encodes: resultName|FUNCXXX|param
                        cb.Tag = resultName & "|" & funcKey.ToUpperInvariant() & "|" & param
                        cb.Name = "Chk_" & resultName & "_" & funcKey

                        GroupBoxCustom.Controls.Add(cb)

                    End If

                Case "LABEL"
                    ' LABEL;Caption;x=..;y=..;f=..
                    ' LABEL;Caption;X;Y;FontSize    (positional still supported)

                    If parts.Length >= 2 Then

                        Dim caption As String = parts(1).Trim()

                        Dim x As Integer = 0
                        Dim y As Integer = 0
                        Dim fontSize As Single = 10.0F   ' default

                        ' -----------------------
                        ' Positional form:
                        ' LABEL;cap;X;Y;F
                        ' -----------------------
                        Dim usedPositional As Boolean = False

                        If parts.Length >= 5 AndAlso Not parts(2).Contains("=") Then
                            ParseIntField(parts(2), x)
                            ParseIntField(parts(3), y)

                            Dim fs As Single
                            If Single.TryParse(parts(4).Trim(), fs) Then fontSize = fs

                            usedPositional = True
                        End If

                        ' -----------------------
                        ' Named-parameter form:
                        ' LABEL;cap;x=...;y=...;f=...
                        ' -----------------------
                        If Not usedPositional Then
                            For i As Integer = 2 To parts.Length - 1
                                Dim p = parts(i).Trim()
                                If p.Contains("=") Then
                                    Dim kv = p.Split("="c)
                                    If kv.Length = 2 Then
                                        Dim key = kv(0).Trim().ToLower()
                                        Dim val = kv(1).Trim()

                                        Select Case key
                                            Case "x" : ParseIntField(val, x)
                                            Case "y" : ParseIntField(val, y)
                                            Case "f"
                                                Dim fs As Single
                                                If Single.TryParse(val, fs) Then fontSize = fs
                                        End Select
                                    End If
                                End If
                            Next
                        End If

                        ' -----------------------
                        ' Create label
                        ' -----------------------
                        Dim lbl As New Label()
                        lbl.Text = caption
                        lbl.AutoSize = True
                        lbl.Location = New Point(x, y)
                        lbl.Font = New Font(lbl.Font.FontFamily, fontSize, FontStyle.Regular)

                        GroupBoxCustom.Controls.Add(lbl)
                    End If



                Case "DROPDOWN"
                    ' Format:
                    ' DROPDOWN;ControlName;Caption;DeviceName;CommandPrefix;X;Y;Width;Item1,Item2,Item3...
                    '
                    ' Example:
                    ' DROPDOWN;ComboNPLC;NPLC:;dev1;:VOLT:DC:NPLC ;400;40;120;1,10,100

                    If parts.Length < 9 Then Continue For

                    Dim ctrlName = parts(1).Trim()
                    Dim caption = parts(2).Trim()
                    Dim deviceName = parts(3).Trim()
                    Dim commandPrefix = parts(4).Trim()

                    Dim x As Integer = 0
                    Dim y As Integer = 0
                    Dim w As Integer = 120   ' sensible default width

                    If parts.Length > 5 Then ParseIntField(parts(5), x)
                    If parts.Length > 6 Then ParseIntField(parts(6), y)
                    If parts.Length > 7 Then ParseIntField(parts(7), w)

                    Dim itemsRaw = parts(8).Trim()
                    Dim items = itemsRaw.Split(","c)

                    ' Label next to dropdown
                    Dim lbl As New Label()
                    lbl.Text = caption
                    lbl.AutoSize = True
                    lbl.Location = New Point(x, y)
                    GroupBoxCustom.Controls.Add(lbl)

                    ' Create dropdown (height not used)
                    Dim cb As New ComboBox()
                    cb.Name = ctrlName
                    cb.DropDownStyle = ComboBoxStyle.DropDownList
                    cb.Location = New Point(x + lbl.PreferredWidth + 8, y - 3)
                    cb.Size = New Size(w, cb.Height)

                    ' Blank first item
                    cb.Items.Add("")

                    ' Add items
                    For Each it In items
                        If it.Trim() <> "" Then cb.Items.Add(it.Trim())
                    Next

                    cb.SelectedIndex = 0

                    ' Tag holds Device + CommandPrefix
                    cb.Tag = deviceName & "|" & commandPrefix

                    AddHandler cb.SelectedIndexChanged, AddressOf Dropdown_SelectedIndexChanged

                    GroupBoxCustom.Controls.Add(cb)

                Case "RADIOGROUP"
                    ' RADIOGROUP;GroupName;Caption;X;Y;Width;Height
                    ' Example:
                    ' RADIOGROUP;DCRange;DC Voltage Range;20;200;260;140

                    If parts.Length >= 7 Then

                        Dim groupName As String = parts(1).Trim()
                        Dim caption As String = parts(2).Trim()

                        Dim x As Integer = 0
                        Dim y As Integer = 0
                        Dim w As Integer = 100   ' sensible default
                        Dim h As Integer = 60    ' sensible default

                        If parts.Length > 3 Then ParseIntField(parts(3), x)
                        If parts.Length > 4 Then ParseIntField(parts(4), y)
                        If parts.Length > 5 Then ParseIntField(parts(5), w)
                        If parts.Length > 6 Then ParseIntField(parts(6), h)

                        Dim gb As New GroupBox()
                        gb.Name = "RG_" & groupName
                        gb.Text = caption
                        gb.Location = New Point(x, y)
                        gb.Size = New Size(w, h)

                        GroupBoxCustom.Controls.Add(gb)

                    End If

                Case "RADIO"
                    ' RADIO;GroupName;Caption;DeviceName;Command;X;Y
                    ' Example:
                    ' RADIO;DCRange;10 V;dev1;:SENS:VOLT:DC:RANG 10;10;20

                    If parts.Length >= 7 Then

                        Dim groupName As String = parts(1).Trim()
                        Dim caption As String = parts(2).Trim()
                        Dim deviceName As String = parts(3).Trim()
                        Dim command As String = parts(4).Trim()

                        Dim relX As Integer = 0
                        Dim relY As Integer = 0

                        If parts.Length > 5 Then ParseIntField(parts(5), relX)
                        If parts.Length > 6 Then ParseIntField(parts(6), relY)

                        ' Find the parent group box
                        Dim parentName As String = "RG_" & groupName
                        Dim found() As Control = GroupBoxCustom.Controls.Find(parentName, True)

                        If found Is Nothing OrElse found.Length = 0 Then
                            ' No matching RADIOGROUP found; ignore this RADIO
                            Exit Select
                        End If

                        Dim gb As GroupBox = TryCast(found(0), GroupBox)
                        If gb Is Nothing Then Exit Select

                        Dim rb As New RadioButton()
                        rb.Text = caption
                        rb.AutoSize = True
                        rb.Location = New Point(relX, relY)

                        ' Tag encodes: deviceName|command
                        rb.Tag = deviceName & "|" & command

                        AddHandler rb.CheckedChanged, AddressOf Radio_CheckedChanged

                        gb.Controls.Add(rb)

                    End If

                Case "SLIDER"
                    ' Supports:
                    ' SLIDER;Name;Caption;Device;Command;X;Y;W;Min;Max;Step;Scale
                    ' or
                    ' SLIDER;Name;Caption;Device;Command;x=..;y=..;w=..;min=..;max=..;step=..;scale=..;hint=..

                    If parts.Length >= 5 Then

                        Dim controlName As String = parts(1).Trim()
                        Dim caption As String = parts(2).Trim()
                        Dim deviceName As String = parts(3).Trim()
                        Dim commandPrefix As String = parts(4).Trim()

                        ' ---- defaults ----
                        Dim x As Integer = 20
                        Dim y As Integer = 20
                        Dim width As Integer = 200
                        Dim minVal As Integer = 0
                        Dim maxVal As Integer = 100
                        Dim stepVal As Integer = 1
                        Dim scale As Double = 1.0
                        Dim hintText As String = ""

                        ' --------------------------------------------------------
                        ' Try positional format first
                        ' --------------------------------------------------------
                        Dim usedPositional As Boolean = False

                        If parts.Length >= 12 AndAlso Not parts(5).Contains("="c) Then
                            ' Positional format:
                            ' SLIDER;Name;Caption;Device;Command;X;Y;W;Min;Max;Step;Scale

                            ParseIntField(parts(5), x)
                            ParseIntField(parts(6), y)
                            ParseIntField(parts(7), width)
                            ParseIntField(parts(8), minVal)
                            ParseIntField(parts(9), maxVal)
                            ParseIntField(parts(10), stepVal)

                            Double.TryParse(
                                   parts(11).Trim(),
                                   Globalization.NumberStyles.Float,
                                   Globalization.CultureInfo.InvariantCulture,
                                   scale
                            )

                            usedPositional = True
                        End If


                        ' --------------------------------------------------------
                        ' Named-parameter format
                        ' --------------------------------------------------------
                        If Not usedPositional Then
                            For i As Integer = 5 To parts.Length - 1
                                Dim p = parts(i).Trim()
                                If p.Contains("=") Then
                                    Dim kv = p.Split("="c)
                                    If kv.Length = 2 Then
                                        Dim key = kv(0).Trim().ToLower()
                                        Dim val = kv(1).Trim()

                                        Select Case key
                                            Case "x" : Integer.TryParse(val, x)
                                            Case "y" : Integer.TryParse(val, y)
                                            Case "w" : Integer.TryParse(val, width)
                                            Case "min" : Integer.TryParse(val, minVal)
                                            Case "max" : Integer.TryParse(val, maxVal)
                                            Case "step" : Integer.TryParse(val, stepVal)
                                            Case "scale" : Double.TryParse(val, scale)
                                            Case "hint" : hintText = val
                                        End Select
                                    End If
                                End If
                            Next
                        End If

                        ' --------------------------------------------------------
                        ' Build GroupBox
                        ' --------------------------------------------------------
                        Dim gb As New GroupBox()
                        gb.Text = caption
                        gb.Location = New Point(x, y)
                        gb.Size = New Size(width + 60, 70)
                        gb.Name = "GB_Slider_" & controlName
                        GroupBoxCustom.BackColor = Me.BackColor
                        GroupBoxCustom.Controls.Add(gb)

                        ' Value label
                        Dim lblValue As New Label()
                        lblValue.Name = "Lbl_" & controlName & "_Value"
                        lblValue.AutoSize = True
                        lblValue.Text = ""   ' blanks at start
                        lblValue.Location = New Point(width + 20, 30)
                        gb.Controls.Add(lblValue)
                        gb.Height = 72 ' groupbox height

                        ' Trackbar slider
                        Dim tb As New TrackBar()
                        tb.Name = controlName
                        tb.Location = New Point(10, 25)
                        tb.Width = width
                        tb.Minimum = minVal
                        tb.Maximum = maxVal
                        tb.SmallChange = stepVal
                        tb.LargeChange = stepVal * 5
                        tb.TickStyle = TickStyle.None

                        tb.Tag = deviceName & "|" &
                                 commandPrefix & "|" &
                                 scale.ToString(Globalization.CultureInfo.InvariantCulture) & "|" &
                                 lblValue.Name & "|" &
                                 stepVal.ToString()

                        AddHandler tb.Scroll, AddressOf Slider_Scroll
                        AddHandler tb.MouseUp, AddressOf Slider_MouseUp

                        gb.Controls.Add(tb)

                        ' Tooltip hint support
                        If hintText <> "" Then
                            Dim tt As New ToolTip()
                            tt.SetToolTip(gb, hintText)
                            tt.SetToolTip(tb, hintText)
                            tt.SetToolTip(lblValue, hintText)
                        End If

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


    Private Sub ButtonResetTxt_Click(sender As Object, e As EventArgs) Handles ButtonResetTxt.Click

        ' Remove all dynamically created controls
        GroupBoxCustom.Controls.Clear()

        ' Reset the title text
        GroupBoxCustom.Text = "User Defineable"

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


    Private Sub Dropdown_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim cb = TryCast(sender, ComboBox)
        If cb Is Nothing Then Exit Sub

        Dim meta As String = TryCast(cb.Tag, String)
        If String.IsNullOrEmpty(meta) Then Exit Sub

        Dim parts = meta.Split("|"c)
        If parts.Length < 2 Then Exit Sub

        Dim deviceName = parts(0)
        Dim commandPrefix = parts(1).TrimEnd()

        Dim dev As IODevices.IODevice = Nothing
        Select Case deviceName.ToLowerInvariant()
            Case "dev1" : dev = dev1
            Case "dev2" : dev = dev2
        End Select
        If dev Is Nothing Then
            MessageBox.Show("Unknown device: " & deviceName)
            Exit Sub
        End If

        Dim selected As String = ""
        If cb.SelectedItem IsNot Nothing Then
            selected = cb.SelectedItem.ToString().Trim()
        End If

        ' Blank entry → do nothing
        If selected = "" Then Exit Sub

        Dim cmd As String = commandPrefix & " " & selected
        dev.SendAsync(cmd, True)
    End Sub


    Private Sub Radio_CheckedChanged(sender As Object, e As EventArgs)
        Dim rb = TryCast(sender, RadioButton)
        If rb Is Nothing Then Exit Sub

        ' Only act when it becomes checked, not when unchecked
        If Not rb.Checked Then Exit Sub

        Dim meta As String = TryCast(rb.Tag, String)
        If String.IsNullOrEmpty(meta) Then Exit Sub

        Dim parts = meta.Split("|"c)
        If parts.Length < 2 Then Exit Sub

        Dim deviceName As String = parts(0)
        Dim command As String = parts(1)

        Dim dev As IODevices.IODevice = Nothing

        Select Case deviceName.ToLowerInvariant()
            Case "dev1" : dev = dev1
            Case "dev2" : dev = dev2
                ' add more if needed
        End Select

        If dev Is Nothing Then
            MessageBox.Show("Unknown device in RADIO: " & deviceName)
            Exit Sub
        End If

        dev.SendAsync(command, True)
    End Sub


    Private Sub Slider_Scroll(sender As Object, e As EventArgs)
        Dim tb = TryCast(sender, TrackBar)
        If tb Is Nothing Then Exit Sub

        Dim meta As String = TryCast(tb.Tag, String)
        If String.IsNullOrEmpty(meta) Then Exit Sub

        Dim parts = meta.Split("|"c)
        If parts.Length < 5 Then Exit Sub

        Dim scale As Double
        Double.TryParse(parts(2), Globalization.NumberStyles.Float,
                    Globalization.CultureInfo.InvariantCulture, scale)

        Dim valueLabelName As String = parts(3)

        Dim stepVal As Integer = 1
        Integer.TryParse(parts(4), stepVal)
        If stepVal < 1 Then stepVal = 1

        ' Snap to nearest multiple of stepVal
        Dim raw As Integer = tb.Value
        Dim snapped As Integer = CInt(Math.Round(raw / CDbl(stepVal))) * stepVal
        If snapped < tb.Minimum Then snapped = tb.Minimum
        If snapped > tb.Maximum Then snapped = tb.Maximum

        If snapped <> tb.Value Then
            tb.Value = snapped
        End If

        Dim scaledValue As Double = snapped * scale

        ' Update value label
        Dim lblArr = GroupBoxCustom.Controls.Find(valueLabelName, True)
        If lblArr IsNot Nothing AndAlso lblArr.Length > 0 Then
            Dim lbl = TryCast(lblArr(0), Label)
            If lbl IsNot Nothing Then
                lbl.Text = scaledValue.ToString("0.#####",
                        Globalization.CultureInfo.InvariantCulture)
            End If
        End If
    End Sub


    Private Sub Slider_MouseUp(sender As Object, e As MouseEventArgs)
        Dim tb = TryCast(sender, TrackBar)
        If tb Is Nothing Then Exit Sub

        Dim meta As String = TryCast(tb.Tag, String)
        If String.IsNullOrEmpty(meta) Then Exit Sub

        Dim parts = meta.Split("|"c)
        If parts.Length < 5 Then Exit Sub

        Dim deviceName As String = parts(0)
        Dim commandPrefix As String = parts(1)

        Dim scale As Double
        Double.TryParse(parts(2), Globalization.NumberStyles.Float,
                        Globalization.CultureInfo.InvariantCulture, scale)

        Dim stepVal As Integer = 1
        Integer.TryParse(parts(4), stepVal)
        If stepVal < 1 Then stepVal = 1

        ' Snap again on mouse-up (in case of any drift)
        Dim raw As Integer = tb.Value
        Dim snapped As Integer = CInt(Math.Round(raw / CDbl(stepVal))) * stepVal
        If snapped < tb.Minimum Then snapped = tb.Minimum
        If snapped > tb.Maximum Then snapped = tb.Maximum

        If snapped <> tb.Value Then
            tb.Value = snapped
        End If

        Dim scaledValue As Double = snapped * scale

        ' Resolve device
        Dim dev As IODevices.IODevice = Nothing
        Select Case deviceName.ToLowerInvariant()
            Case "dev1" : dev = dev1
            Case "dev2" : dev = dev2
        End Select
        If dev Is Nothing Then Exit Sub

        Dim cmd As String = commandPrefix.TrimEnd() & " " &
                            scaledValue.ToString("G", Globalization.CultureInfo.InvariantCulture)

        dev.SendAsync(cmd, True)
    End Sub


    Private Sub ParseIntField(token As String, ByRef target As Integer)
        If String.IsNullOrWhiteSpace(token) Then Exit Sub

        Dim t As String = token.Trim()

        ' If it has "=", strip everything up to and including "="
        Dim eqPos As Integer = t.IndexOf("="c)
        If eqPos >= 0 AndAlso eqPos < t.Length - 1 Then
            t = t.Substring(eqPos + 1).Trim()
        End If

        Dim v As Integer
        If Integer.TryParse(t, v) Then
            target = v
        End If
    End Sub




End Class
