' User customizeable tab

Imports IODevices
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Windows.Forms.DataVisualization.Charting


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

        ' Reset per-config runtime state
        Timer5.Enabled = False
        AutoReadDeviceName = ""
        AutoReadCommand = ""
        AutoReadResultControl = ""

        GroupBoxCustom.Controls.Clear()

        Dim autoY As Integer = 10

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
'   or with extras:
'   TEXTBOX;Name;Label text;X;Y;W;H;readonly=true
'   or token form:
'   TEXTBOX;Name;Label text;x=..;y=..;w=..;h=..;readonly=true
' ======================================
                Case "TEXTBOX"
                    If parts.Length < 3 Then Continue For

                    Dim tbName = parts(1).Trim()
                    Dim labelText = parts(2).Trim()

                    Dim x As Integer = 10
                    Dim y As Integer = autoY
                    Dim w As Integer = 120
                    Dim h As Integer = 22
                    Dim isReadOnly As Boolean = False
                    Dim hasExplicitCoords As Boolean = False

                    ' ---------- positional form ----------
                    ' TEXTBOX;Name;Caption;X;Y[;W;H;readonly=..]
                    If parts.Length >= 5 AndAlso Not parts(3).Contains("="c) Then
                        ParseIntField(parts(3), x)
                        ParseIntField(parts(4), y)
                        hasExplicitCoords = True

                        If parts.Length >= 7 Then ParseIntField(parts(5), w)
                        If parts.Length >= 8 Then ParseIntField(parts(6), h)

                        ' optional readonly in positional form
                        If parts.Length >= 9 Then
                            ParseBoolField(parts(7), isReadOnly)
                        End If

                    Else
                        ' ---------- token form ----------
                        ' TEXTBOX;Name;Caption;x=..;y=..;w=..;h=..;readonly=true
                        For i As Integer = 3 To parts.Length - 1
                            Dim token = parts(i).Trim()
                            If Not token.Contains("="c) Then Continue For

                            Dim kv = token.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLower()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "x"
                                    ParseIntField(val, x)
                                    hasExplicitCoords = True
                                Case "y"
                                    ParseIntField(val, y)
                                    hasExplicitCoords = True
                                Case "w"
                                    ParseIntField(val, w)
                                Case "h"
                                    ParseIntField(val, h)
                                Case "readonly"
                                    ParseBoolField(val, isReadOnly)
                            End Select
                        Next
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
                    tb.Location = New Point(lbl.Right + 5, y - 2)
                    tb.Size = New Size(w, h)
                    tb.ReadOnly = isReadOnly

                    GroupBoxCustom.Controls.Add(tb)

                    ' Only advance autoY if coordinates not explicitly given
                    If Not hasExplicitCoords Then
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

                Case "HR"
                    ' Supports:
                    ' HR;X;Y;Width
                    ' HR;x=..;y=..;w=..

                    Dim x As Integer = 0
                    Dim y As Integer = 0
                    Dim w As Integer = 200   ' default width

                    ' ---------- Positional format ----------
                    If parts.Length >= 4 AndAlso Not parts(1).Contains("=") Then
                        ParseIntField(parts(1), x)
                        ParseIntField(parts(2), y)
                        ParseIntField(parts(3), w)

                    Else
                        ' ---------- Token format ----------
                        For i As Integer = 1 To parts.Length - 1
                            Dim p = parts(i).Trim()
                            If p.Contains("=") Then
                                Dim kv = p.Split("="c)
                                If kv.Length = 2 Then
                                    Dim key = kv(0).Trim().ToLower()
                                    Dim val = kv(1).Trim()

                                    Select Case key
                                        Case "x" : ParseIntField(val, x)
                                        Case "y" : ParseIntField(val, y)
                                        Case "w" : ParseIntField(val, w)
                                    End Select
                                End If
                            End If
                        Next
                    End If

                    ' ---------- Draw the horizontal rule ----------
                    Dim hrLine As New Label()
                    hrLine.AutoSize = False
                    hrLine.BorderStyle = BorderStyle.Fixed3D
                    hrLine.Height = 2
                    hrLine.Width = w
                    hrLine.Location = New Point(x, y)

                    GroupBoxCustom.Controls.Add(hrLine)


                Case "VR"
                    ' VR;X;Y;Height
                    ' VR;x=20;y=200;h=150

                    Dim x As Integer = 10
                    Dim y As Integer = 10
                    Dim h As Integer = 100

                    ' Positional format?
                    If parts.Length >= 4 AndAlso Not parts(1).Contains("=") Then
                        ParseIntField(parts(1), x)
                        ParseIntField(parts(2), y)
                        ParseIntField(parts(3), h)
                    Else
                        ' Token format
                        For i As Integer = 1 To parts.Length - 1
                            Dim tok = parts(i).Trim()
                            Dim kv = tok.Split("="c)
                            If kv.Length <> 2 Then Continue For
                            Dim key = kv(0).Trim().ToLower()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "x" : ParseIntField(val, x)
                                Case "y" : ParseIntField(val, y)
                                Case "h" : ParseIntField(val, h)
                            End Select
                        Next
                    End If

                    ' Create vertical line
                    Dim ln As New Label()
                    ln.BorderStyle = BorderStyle.Fixed3D
                    ln.AutoSize = False
                    ln.Width = 2
                    ln.Height = h
                    ln.Location = New Point(x, y)

                    GroupBoxCustom.Controls.Add(ln)


                Case "SPINNER"
                    ' Supports:
                    ' SPINNER;Name;Caption;Device;Command;X;Y;W;Min;Max;Step;Scale
                    ' or
                    ' SPINNER;Name;Caption;Device;Command;x=..;y=..;w=..;min=..;max=..;step=..;scale=..

                    If parts.Length >= 5 Then

                        Dim controlName As String = parts(1).Trim()
                        Dim caption As String = parts(2).Trim()
                        Dim deviceName As String = parts(3).Trim()
                        Dim commandPrefix As String = parts(4).Trim()

                        Dim x As Integer = 20
                        Dim y As Integer = 20
                        Dim width As Integer = 80
                        Dim minVal As Integer = 0
                        Dim maxVal As Integer = 100
                        Dim stepVal As Integer = 1
                        Dim scale As Double = 1.0

                        Dim usedPositional As Boolean = False

                        ' Positional form
                        If parts.Length >= 12 AndAlso Not parts(5).Contains("="c) Then
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

                        ' Token form
                        If Not usedPositional Then
                            For i As Integer = 5 To parts.Length - 1
                                Dim p = parts(i).Trim()
                                If p.Contains("="c) Then
                                    Dim kv = p.Split("="c)
                                    If kv.Length = 2 Then
                                        Dim key = kv(0).Trim().ToLower()
                                        Dim val = kv(1).Trim()
                                        Select Case key
                                            Case "x" : ParseIntField(val, x)
                                            Case "y" : ParseIntField(val, y)
                                            Case "w" : ParseIntField(val, width)
                                            Case "min" : ParseIntField(val, minVal)
                                            Case "max" : ParseIntField(val, maxVal)
                                            Case "step" : ParseIntField(val, stepVal)
                                            Case "scale"
                                                Double.TryParse(
                                    val,
                                    Globalization.NumberStyles.Float,
                                    Globalization.CultureInfo.InvariantCulture,
                                    scale
                                )
                                        End Select
                                    End If
                                End If
                            Next
                        End If

                        ' Label
                        Dim lbl As New Label()
                        lbl.Text = caption
                        lbl.AutoSize = True
                        lbl.Location = New Point(x, y + 3)
                        GroupBoxCustom.Controls.Add(lbl)

                        ' Spinner (NumericUpDown)
                        Dim nud As New NumericUpDown()
                        nud.Name = controlName
                        nud.Location = New Point(x + lbl.PreferredWidth + 5, y)
                        nud.Width = width

                        nud.Minimum = minVal
                        nud.Maximum = maxVal
                        nud.Increment = stepVal
                        nud.DecimalPlaces = 0   ' using scale for fractional output

                        ' Start at min internally, but show blank
                        nud.Value = minVal
                        nud.Text = ""

                        ' Tag = DEVICE|COMMAND|SCALE|INITFLAG|MIN
                        nud.Tag = deviceName & "|" &
                                  commandPrefix & "|" &
                                  scale.ToString(Globalization.CultureInfo.InvariantCulture) & "|" &
                                  "0" & "|" &
                                  minVal.ToString(Globalization.CultureInfo.InvariantCulture)

                        AddHandler nud.ValueChanged, AddressOf Spinner_ValueChanged

                        GroupBoxCustom.Controls.Add(nud)

                    End If


                Case "LED"
                    ' LED;Name;Caption;X;Y;Size
                    ' or LED;Name;Caption;x=..;y=..;s=..;on=..;off=..;bad=..

                    If parts.Length >= 3 Then

                        Dim ledName As String = parts(1).Trim()
                        Dim caption As String = parts(2).Trim()

                        Dim x As Integer = 10
                        Dim y As Integer = 10
                        Dim size As Integer = 12

                        ' Default colours
                        Dim onColor As Color = Color.LimeGreen
                        Dim offColor As Color = Color.DarkGray
                        Dim badColor As Color = Color.Gold

                        ' --- positional X,Y,Size: LED;Name;Caption;X;Y;Size ---
                        If parts.Length >= 6 AndAlso Not parts(3).Contains("="c) Then
                            ParseIntField(parts(3), x)
                            ParseIntField(parts(4), y)
                            ParseIntField(parts(5), size)
                        End If

                        ' --- token style: x=..;y=..;s=..;on=..;off=..;bad=.. ---
                        For i As Integer = 3 To parts.Length - 1
                            Dim p = parts(i).Trim()
                            If p.Contains("="c) Then
                                Dim kv = p.Split("="c)
                                If kv.Length = 2 Then
                                    Dim key = kv(0).Trim().ToLower()
                                    Dim val = kv(1).Trim()
                                    Select Case key
                                        Case "x" : ParseIntField(val, x)
                                        Case "y" : ParseIntField(val, y)
                                        Case "s" : ParseIntField(val, size)
                                        Case "on"
                                            Dim c = Color.FromName(val)
                                            If c.ToArgb() <> 0 Then onColor = c
                                        Case "off"
                                            Dim c = Color.FromName(val)
                                            If c.ToArgb() <> 0 Then offColor = c
                                        Case "bad"
                                            Dim c = Color.FromName(val)
                                            If c.ToArgb() <> 0 Then badColor = c
                                    End Select
                                End If
                            End If
                        Next

                        ' Caption label
                        Dim lbl As New Label()
                        lbl.Text = caption
                        lbl.AutoSize = True
                        lbl.Location = New Point(x, y + 2)
                        GroupBoxCustom.Controls.Add(lbl)

                        ' LED panel
                        Dim led As New Panel()
                        led.Name = ledName
                        led.Size = New Size(size, size)
                        led.Location = New Point(lbl.Right + 5, y)
                        led.BorderStyle = BorderStyle.FixedSingle
                        led.BackColor = offColor    ' start OFF

                        ' Tag: marker + colours
                        led.Tag = $"LED|{onColor.ToArgb}|{offColor.ToArgb}|{badColor.ToArgb}"

                        GroupBoxCustom.Controls.Add(led)

                    End If


                Case "BIGTEXT"
                    ' BIGTEXT;ControlName;Caption(optional);x=..;y=..;f=..;w=..;h=..;border=on/off
                    If parts.Length >= 2 Then

                        Dim controlName As String = parts(1).Trim()
                        Dim caption As String = If(parts.Length > 2, parts(2).Trim(), "")

                        Dim x As Integer = 20
                        Dim y As Integer = 20
                        Dim fontSize As Single = 28.0F
                        Dim w As Integer = 200     ' default width
                        Dim h As Integer = 60      ' default height

                        ' NEW: border flag (default = ON, same as your original code)
                        Dim borderOn As Boolean = True

                        ' Parse named tokens
                        For i As Integer = 3 To parts.Length - 1
                            Dim p = parts(i).Trim()
                            If p.Contains("="c) Then
                                Dim kv = p.Split("="c)
                                If kv.Length = 2 Then
                                    Dim key = kv(0).Trim().ToLower()
                                    Dim val = kv(1).Trim()

                                    Select Case key
                                        Case "x" : Integer.TryParse(val, x)
                                        Case "y" : Integer.TryParse(val, y)
                                        Case "f" : Single.TryParse(val, fontSize)
                                        Case "w" : Integer.TryParse(val, w)
                                        Case "h" : Integer.TryParse(val, h)

                                        Case "border"
                                            Dim vLower = val.ToLower()
                                            Select Case vLower
                                                Case "0", "off", "false", "no", "none"
                                                    borderOn = False
                                                Case Else
                                                    borderOn = True
                                            End Select
                                    End Select
                                End If
                            End If
                        Next

                        ' --- Outer panel (the "box") ---
                        Dim panel As New Panel()
                        panel.Location = New Point(x, y)
                        panel.Size = New Size(w, h)
                        panel.Name = "Panel_" & controlName

                        panel.BorderStyle = If(borderOn, BorderStyle.FixedSingle, BorderStyle.None)
                        panel.BackColor = GroupBoxCustom.BackColor

                        GroupBoxCustom.Controls.Add(panel)

                        ' --- BIG TEXT LABEL ---
                        Dim lbl As New Label()
                        lbl.Name = controlName
                        lbl.Text = ""
                        lbl.AutoSize = False
                        lbl.TextAlign = ContentAlignment.MiddleCenter
                        lbl.Font = New Font(lbl.Font.FontFamily, fontSize, FontStyle.Bold)
                        lbl.Dock = DockStyle.Fill

                        lbl.Text = "#####" ' default text

                        panel.Controls.Add(lbl)
                    End If


                ' ======================================
                '   MULTILINE TEXT AREA
                '   TEXTAREA;Name;Caption;X;Y;W;H
                '   or token form:
                '   TEXTAREA;Name;Caption;x=..;y=..;w=..;h=..;init=line1|line2|...
                ' ======================================
                Case "TEXTAREA"
                    If parts.Length < 3 Then Continue For

                    Dim tbName = parts(1).Trim()
                    Dim labelText = parts(2).Trim()

                    Dim x As Integer = 10
                    Dim y As Integer = autoY
                    Dim w As Integer = 300
                    Dim h As Integer = 100
                    Dim hasExplicitCoords As Boolean = False
                    Dim initLines As String() = Nothing

                    ' ----- positional form? (X,Y[,W,H]) -----
                    If parts.Length >= 5 AndAlso Not parts(3).Contains("="c) Then
                        ParseIntField(parts(3), x)
                        ParseIntField(parts(4), y)
                        hasExplicitCoords = True

                        If parts.Length >= 7 Then ParseIntField(parts(5), w)
                        If parts.Length >= 8 Then ParseIntField(parts(6), h)

                        ' extra tokens may include init=
                        For i2 As Integer = 7 To parts.Length - 1
                            Dim p2 = parts(i2).Trim()
                            If p2.StartsWith("init=", StringComparison.OrdinalIgnoreCase) Then
                                Dim val = p2.Substring(5).Trim()
                                If val <> "" Then initLines = val.Split("|"c)
                            End If
                        Next

                    Else
                        ' ----- token form: x=..;y=..;w=..;h=..;init=... -----
                        For i2 As Integer = 3 To parts.Length - 1
                            Dim p2 = parts(i2).Trim()
                            If Not p2.Contains("="c) Then Continue For

                            Dim kv = p2.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLower()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "x" : ParseIntField(val, x) : hasExplicitCoords = True
                                Case "y" : ParseIntField(val, y) : hasExplicitCoords = True
                                Case "w" : ParseIntField(val, w)
                                Case "h" : ParseIntField(val, h)
                                Case "init"
                                    If val <> "" Then initLines = val.Split("|"c)
                            End Select
                        Next
                    End If

                    ' Label
                    Dim lbl As New Label()
                    lbl.Text = labelText
                    lbl.AutoSize = True
                    lbl.Location = New Point(x, y)
                    GroupBoxCustom.Controls.Add(lbl)

                    ' Multiline TextBox below the label
                    Dim tb As New TextBox()
                    tb.Name = tbName
                    tb.Multiline = True
                    tb.ScrollBars = ScrollBars.Vertical
                    tb.Location = New Point(x, y + lbl.Height + 3)
                    tb.Size = New Size(w, h)

                    ' apply init contents if provided
                    If initLines IsNot Nothing Then
                        tb.Lines = initLines
                    End If

                    GroupBoxCustom.Controls.Add(tb)

                    If Not hasExplicitCoords Then
                        autoY += lbl.Height + h + 8
                    End If

                    Continue For

                ' ======================================
                '   TOGGLE BUTTON
                '   TOGGLE;Name;Caption;Device;CommandOn;CommandOff;X;Y;W;H
                '   or token form:
                '   TOGGLE;Name;Caption;Device;CommandOn;CommandOff;x=..;y=..;w=..;h=..
                ' ======================================

                Case "TOGGLE"
                    ' TOGGLE;Name;Caption;Device;OnCmd;OffCmd;X;Y;W;H
                    ' OR token form: x=..;y=..;w=..;h=..

                    If parts.Length >= 6 Then

                        Dim name As String = parts(1).Trim()
                        Dim caption As String = parts(2).Trim()
                        Dim device As String = parts(3).Trim()
                        Dim cmdOn As String = parts(4).Trim()
                        Dim cmdOff As String = parts(5).Trim()

                        Dim x As Integer = 20, y As Integer = 20, w As Integer = 120, h As Integer = 35
                        Dim usedPositional As Boolean = False

                        ' Positional
                        If parts.Length >= 10 AndAlso Not parts(6).Contains("="c) Then
                            ParseIntField(parts(6), x)
                            ParseIntField(parts(7), y)
                            ParseIntField(parts(8), w)
                            ParseIntField(parts(9), h)
                            usedPositional = True
                        End If

                        ' Token format
                        If Not usedPositional Then
                            For i As Integer = 6 To parts.Length - 1
                                Dim p = parts(i).Trim()
                                If p.Contains("="c) Then
                                    Dim kv = p.Split("="c)
                                    If kv.Length = 2 Then
                                        Select Case kv(0).Trim().ToLower()
                                            Case "x" : ParseIntField(kv(1), x)
                                            Case "y" : ParseIntField(kv(1), y)
                                            Case "w" : ParseIntField(kv(1), w)
                                            Case "h" : ParseIntField(kv(1), h)
                                        End Select
                                    End If
                                End If
                            Next
                        End If

                        ' Create toggle button
                        Dim b As New Button()
                        b.Name = name
                        b.Text = caption
                        b.Location = New Point(x, y)
                        b.Size = New Size(w, h)

                        ' Tag holds: DEVICE|ONCMD|OFFCMD|STATE
                        b.Tag = device & "|" & cmdOn & "|" & cmdOff & "|0"   ' state 0 = OFF

                        AddHandler b.Click, AddressOf ToggleButton_Click

                        GroupBoxCustom.Controls.Add(b)
                    End If


                Case "CHART"
                    ' CHART;ChartName;Caption;ResultTarget;
                    '       x=..;y=..;w=..;h=..;ymin=..;ymax=..;xstep=..;maxpoints=..;color=..

                    If parts.Length < 4 Then Continue For

                    Dim chartName As String = parts(1).Trim()
                    Dim caption As String = parts(2).Trim()
                    Dim resultTarget As String = parts(3).Trim()

                    ' ---- defaults ----
                    Dim x As Integer = 20
                    Dim y As Integer = 20
                    Dim w As Integer = 250
                    Dim h As Integer = 120
                    Dim yMin As Double? = Nothing
                    Dim yMax As Double? = Nothing
                    Dim xStep As Double = 1.0R
                    Dim maxPoints As Integer = 100   ' rolling window length
                    Dim plotColor As Color = Color.Yellow   ' default trace colour

                    ' ---- parse named tokens ----
                    For i As Integer = 4 To parts.Length - 1
                        Dim p = parts(i).Trim()
                        If Not p.Contains("="c) Then Continue For

                        Dim kv = p.Split("="c)
                        If kv.Length <> 2 Then Continue For

                        Dim key = kv(0).Trim().ToLower()
                        Dim val = kv(1).Trim()

                        Select Case key
                            Case "x" : ParseIntField(val, x)
                            Case "y" : ParseIntField(val, y)
                            Case "w" : ParseIntField(val, w)
                            Case "h" : ParseIntField(val, h)

                            Case "ymin"
                                Dim d As Double
                                If Double.TryParse(val, Globalization.NumberStyles.Float,
                                   Globalization.CultureInfo.InvariantCulture, d) Then
                                    yMin = d
                                End If

                            Case "ymax"
                                Dim d As Double
                                If Double.TryParse(val, Globalization.NumberStyles.Float,
                                   Globalization.CultureInfo.InvariantCulture, d) Then
                                    yMax = d
                                End If

                            Case "xstep"
                                Double.TryParse(val, Globalization.NumberStyles.Float,
                                Globalization.CultureInfo.InvariantCulture, xStep)

                            Case "maxpoints"
                                Dim mp As Integer
                                If Integer.TryParse(val, mp) AndAlso mp > 0 Then maxPoints = mp

                            Case "color"
                                Dim c As Color = Color.FromName(val)
                                ' FromName returns Empty if invalid
                                If c.ToArgb() <> Color.Empty.ToArgb() Then
                                    plotColor = c
                                End If
                        End Select
                    Next

                    ' ---- Caption label above chart ----
                    Dim lblCap As New Label()
                    lblCap.Text = caption
                    lblCap.AutoSize = True
                    lblCap.Location = New Point(x, y)
                    GroupBoxCustom.Controls.Add(lblCap)

                    Dim chartTop As Integer = y + lblCap.Height + 3

                    ' ---- Create chart ----
                    Dim ch As New DataVisualization.Charting.Chart()
                    ch.Name = chartName
                    ch.Location = New Point(x, chartTop)
                    ch.Size = New Size(w, h)

                    Dim ca As New DataVisualization.Charting.ChartArea("Default")
                    ch.ChartAreas.Add(ca)

                    Dim s As New DataVisualization.Charting.Series("S1")
                    s.ChartType = DataVisualization.Charting.SeriesChartType.Line
                    s.BorderWidth = 2
                    ch.Series.Add(s)

                    ' ---- Apply Y range if specified ----
                    If yMin.HasValue Then
                        ca.AxisY.Minimum = yMin.Value
                    Else
                        ca.AxisY.Minimum = Double.NaN
                    End If

                    If yMax.HasValue Then
                        ca.AxisY.Maximum = yMax.Value
                    Else
                        ca.AxisY.Maximum = Double.NaN
                    End If

                    ' ================================
                    '   Dark theme, fixed-style grid
                    ' ================================
                    ca.BackColor = Color.Black
                    ch.BackColor = Color.Black

                    ' X-axis: vertical grid ON, but no labels
                    ca.AxisX.LabelStyle.Enabled = False
                    ca.AxisX.MajorTickMark.Enabled = False
                    ca.AxisX.MinorTickMark.Enabled = False

                    ca.AxisX.MajorGrid.Enabled = True
                    ca.AxisX.MajorGrid.LineColor = Color.FromArgb(60, 60, 60)
                    ca.AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot
                    ca.AxisX.MinorGrid.Enabled = False

                    ' Y-axis: horizontal grid + smaller font
                    ca.AxisY.MajorGrid.Enabled = True
                    ca.AxisY.MajorGrid.LineColor = Color.FromArgb(80, 80, 80)
                    ca.AxisY.MinorGrid.Enabled = False
                    ca.AxisY.LabelStyle.ForeColor = Color.White
                    ca.AxisY.LabelStyle.Font = New Font("Segoe UI", 7.0F)
                    ca.AxisY.MajorTickMark.Enabled = False

                    ' Series styling with INI colour
                    If ch.Series.Count > 0 Then
                        Dim s0 = ch.Series(0)
                        s0.Color = plotColor
                        s0.BorderWidth = 2
                        s0.MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
                        s0.MarkerSize = 3
                        s0.MarkerColor = plotColor
                    End If

                    GroupBoxCustom.Controls.Add(ch)

                    ' ---- Wire chart to the textbox that holds numeric data ----
                    Dim src = TryCast(GetControlByName(resultTarget), TextBox)
                    If src IsNot Nothing Then
                        ' Initial draw
                        UpdateChartFromText(ch, src.Text, yMin, yMax, xStep, maxPoints)

                        ' Update on every text change
                        AddHandler src.TextChanged,
            Sub(senderSrc As Object, eSrc As EventArgs)
                Dim tbSrc = DirectCast(senderSrc, TextBox)
                UpdateChartFromText(ch, tbSrc.Text, yMin, yMax, xStep, maxPoints)
            End Sub
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


    Private Sub RunQueryToResult(deviceName As String, commandOrPrefix As String, resultControlName As String)

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

        ElseIf status = -1 Then
            ' Instrument/driver reports "blocking" / busy.
            ' Just skip this cycle and leave existing display as-is.
            Exit Sub

        ElseIf q IsNot Nothing Then
            outText = "ERR " & status & ": " & q.errmsg
        Else
            outText = "ERR " & status & " (no IOQuery)"
        End If

        ' Find ALL controls with that name (textbox + BIGTEXT label + LED etc.)
        Dim targets = Me.Controls.Find(resultControlName, True)

        For Each target In targets
            If TypeOf target Is TextBox Then
                DirectCast(target, TextBox).Text = outText

            ElseIf TypeOf target Is Label Then
                DirectCast(target, Label).Text = outText

            ElseIf TypeOf target Is Panel Then
                Dim tagStr = TryCast(target.Tag, String)
                If Not String.IsNullOrEmpty(tagStr) AndAlso tagStr.StartsWith("LED", StringComparison.OrdinalIgnoreCase) Then
                    SetLedStateFromText(DirectCast(target, Panel), outText)
                End If
            End If
        Next



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


            Case "SENDLINES"
                If dev Is Nothing Then Exit Sub

                Dim valCtrl = TryCast(GetControlByName(valueControlName), TextBox)
                If valCtrl Is Nothing Then
                    MessageBox.Show("TEXTAREA not found: " & valueControlName)
                    Exit Sub
                End If

                For Each raw In valCtrl.Lines
                    Dim line = raw.Trim()
                    If line = "" Then Continue For
                    If line.StartsWith(";"c) Then Continue For   ' allow comments

                    ' commandOrPrefix may be blank or may include something like "APPLY DCV "
                    Dim cmd As String
                    If String.IsNullOrWhiteSpace(commandOrPrefix) Then
                        cmd = line
                    Else
                        cmd = commandOrPrefix & line
                    End If

                    dev.SendAsync(cmd & TermStr2(), True)
                Next

        End Select

    End Sub


    Private Sub ButtonResetTxt_Click(sender As Object, e As EventArgs) Handles ButtonResetTxt.Click

        ' Remove all dynamically created controls
        GroupBoxCustom.Controls.Clear()

        ' Reset the title text
        GroupBoxCustom.Text = "User Defineable"

    End Sub


    Private Sub ToggleButton_Click(sender As Object, e As EventArgs)
        Dim b As Button = DirectCast(sender, Button)

        Dim tagParts = CStr(b.Tag).Split("|"c)
        If tagParts.Length < 4 Then Exit Sub

        Dim device = tagParts(0)
        Dim cmdOn = tagParts(1)
        Dim cmdOff = tagParts(2)
        Dim state = CInt(tagParts(3))   ' 0 = OFF, 1 = ON

        ' Pick dev1/dev2
        Dim dev As Object = Nothing
        Select Case device.ToLower()
            Case "dev1" : dev = dev1
            Case "dev2" : dev = dev2
        End Select
        If dev Is Nothing Then Exit Sub

        If state = 0 Then
            ' Turn ON
            dev.SendAsync(cmdOn, True)
            b.BackColor = Color.LimeGreen
            b.ForeColor = Color.Black
            state = 1
        Else
            ' Turn OFF
            dev.SendAsync(cmdOff, True)
            b.BackColor = SystemColors.Control
            b.ForeColor = SystemColors.ControlText
            state = 0
        End If

        ' Update Tag with new state
        b.Tag = device & "|" & cmdOn & "|" & cmdOff & "|" & state
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


    Private Sub ParseBoolField(token As String, ByRef target As Boolean)
        token = token.Trim()

        ' Allow forms: "true", "false", "readonly=true", etc.
        If token.Contains("="c) Then
            token = token.Split("="c)(1).Trim()
        End If

        Select Case token.ToLower()
            Case "1", "true", "yes", "on"
                target = True
            Case "0", "false", "no", "off"
                target = False
        End Select
    End Sub


    Private Sub Spinner_ValueChanged(sender As Object, e As EventArgs)
        Dim nud = DirectCast(sender, NumericUpDown)
        Dim tagStr = TryCast(nud.Tag, String)
        If String.IsNullOrEmpty(tagStr) Then Exit Sub

        Dim parts = tagStr.Split("|"c)
        If parts.Length < 3 Then Exit Sub

        Dim deviceName = parts(0)
        Dim commandPrefix = parts(1)

        Dim scale As Double = 1.0
        Double.TryParse(parts(2),
                    Globalization.NumberStyles.Float,
                    Globalization.CultureInfo.InvariantCulture,
                    scale)

        Dim initFlag As String = If(parts.Length >= 4, parts(3), "1")
        Dim minValFromTag As Decimal = nud.Minimum
        If parts.Length >= 5 Then
            Decimal.TryParse(parts(4),
                         Globalization.NumberStyles.Float,
                         Globalization.CultureInfo.InvariantCulture,
                         minValFromTag)
        End If

        ' === FIRST USER CHANGE ===
        If initFlag = "0" Then
            ' Mark as initialised
            parts(3) = "1"
            nud.Tag = String.Join("|", parts)

            ' Force value to MIN on first change
            nud.Value = minValFromTag

            Dim scaledFirst As Double = CDbl(minValFromTag) * scale
            Dim firstStr As String = scaledFirst.ToString("G", Globalization.CultureInfo.InvariantCulture)
            Dim firstCmd As String = commandPrefix & " " & firstStr

            Dim dev As Object = Nothing
            Select Case deviceName.ToLowerInvariant()
                Case "dev1" : dev = dev1
                Case "dev2" : dev = dev2
            End Select

            If dev IsNot Nothing Then
                dev.SendAsync(firstCmd, True)
            End If

            Exit Sub
        End If

        ' === NORMAL CHANGES AFTER FIRST ===
        Dim scaledValue As Double = CDbl(nud.Value) * scale
        Dim valueStr As String = scaledValue.ToString("G", Globalization.CultureInfo.InvariantCulture)
        Dim cmd As String = commandPrefix & " " & valueStr

        Dim devNorm As Object = Nothing
        Select Case deviceName.ToLowerInvariant()
            Case "dev1" : devNorm = dev1
            Case "dev2" : devNorm = dev2
        End Select

        If devNorm IsNot Nothing Then
            devNorm.SendAsync(cmd, True)
        End If
    End Sub


    Private Sub Spinner_KeepBlank(sender As Object, e As EventArgs)
        Dim nud As NumericUpDown = DirectCast(sender, NumericUpDown)

        ' If spinner is tagged blank, keep it blank until the user changes it
        If nud.Tag IsNot Nothing AndAlso nud.Tag.ToString().Contains("|BLANK") Then
            nud.Text = ""
        End If
    End Sub


    Private Sub SetLedStateFromText(led As Control, reply As String)
        If led Is Nothing Then Exit Sub

        ' Defaults
        Dim onColor As Color = Color.LimeGreen
        Dim offColor As Color = Color.DarkGray
        Dim badColor As Color = Color.Gold

        ' Try to read colours from Tag: "LED|on|off|bad"
        Dim tagStr = TryCast(led.Tag, String)
        If Not String.IsNullOrEmpty(tagStr) Then
            Dim p = tagStr.Split("|"c)
            If p.Length = 4 AndAlso p(0) = "LED" Then
                Dim argb As Integer
                If Integer.TryParse(p(1), argb) Then onColor = Color.FromArgb(argb)
                If Integer.TryParse(p(2), argb) Then offColor = Color.FromArgb(argb)
                If Integer.TryParse(p(3), argb) Then badColor = Color.FromArgb(argb)
            End If
        End If

        Dim t As String = If(reply, "").Trim().ToUpperInvariant()
        Dim newColor As Color

        Select Case t
            Case "1", "ON", "TRUE", "HIGH"
                newColor = onColor
            Case "0", "OFF", "FALSE", "LOW", ""
                newColor = offColor
            Case Else
                Dim v As Double
                If Double.TryParse(t, Globalization.NumberStyles.Float,
                                   Globalization.CultureInfo.InvariantCulture, v) Then
                    newColor = If(v <> 0, onColor, offColor)
                Else
                    newColor = badColor
                End If
        End Select

        DirectCast(led, Panel).BackColor = newColor
    End Sub


    Private Sub SendLinesButton_Click(sender As Object, e As EventArgs)
        Dim b As Button = DirectCast(sender, Button)
        Dim meta = CStr(b.Tag).Split("|"c)

        If meta.Length < 2 Then Exit Sub

        Dim areaName = meta(0)
        Dim deviceName = meta(1)

        Dim tb As TextBox = TryCast(GetControlByName(areaName), TextBox)
        If tb Is Nothing Then
            MessageBox.Show("TEXTAREA not found: " & areaName)
            Exit Sub
        End If

        ' Get device object
        Dim dev = If(deviceName = "dev1", dev1, If(deviceName = "dev2", dev2, Nothing))
        If dev Is Nothing Then
            MessageBox.Show("Invalid device: " & deviceName)
            Exit Sub
        End If

        ' Process each line
        For Each rawLine In tb.Lines
            Dim line = rawLine.Trim()

            If line = "" Then Continue For          ' skip blanks
            If line.StartsWith(";") Then Continue For ' skip comments

            dev.SendAsync(line & TermStr2(), True)
        Next
    End Sub


    Private Sub UpdateChartFromText(ch As DataVisualization.Charting.Chart,
                                text As String,
                                yMin As Double?,
                                yMax As Double?,
                                xStepMinutes As Double,
                                maxPoints As Integer)

        If ch Is Nothing Then Exit Sub
        If maxPoints <= 0 Then maxPoints = 100
        If xStepMinutes <= 0 Then xStepMinutes = 1.0R

        ' Ensure a series exists
        If ch.Series Is Nothing OrElse ch.Series.Count = 0 Then
            Dim sNew As New DataVisualization.Charting.Series("S1")
            sNew.ChartType = DataVisualization.Charting.SeriesChartType.Line
            sNew.BorderWidth = 2
            sNew.MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            sNew.MarkerSize = 3
            sNew.MarkerColor = Color.Yellow   ' fallback
            ch.Series.Add(sNew)
        End If

        Dim s = ch.Series(0)

        If String.IsNullOrWhiteSpace(text) Then
            ch.Invalidate()
            Exit Sub
        End If

        ' ---- Extract numeric values from textbox text ----
        Dim tokens = text.Split({","c, ";"c, " "c, ControlChars.Cr, ControlChars.Lf},
                                StringSplitOptions.RemoveEmptyEntries)

        Dim values As New List(Of Double)
        For Each tok In tokens
            Dim d As Double
            If Double.TryParse(tok.Trim(),
                               Globalization.NumberStyles.Float,
                               Globalization.CultureInfo.InvariantCulture,
                               d) Then
                values.Add(d)
            End If
        Next

        If values.Count = 0 Then
            ch.Invalidate()
            Exit Sub
        End If

        Dim ca = ch.ChartAreas(0)

        ' ---- APPEND if single value, REPLACE if multiple values ----
        If values.Count = 1 Then
            ' Streaming DMM behaviour: append one new point
            Dim lastX As Double
            If s.Points.Count = 0 Then
                lastX = 0.0R
            Else
                lastX = s.Points(s.Points.Count - 1).XValue
            End If

            Dim newX As Double = lastX + xStepMinutes
            s.Points.AddXY(newX, values(0))

        Else
            ' Multi-value: treat as full history, then trim to last maxPoints
            s.Points.Clear()
            For i As Integer = 0 To values.Count - 1
                Dim xVal As Double = i * xStepMinutes
                s.Points.AddXY(xVal, values(i))
            Next
        End If

        ' ---- ROLLING BUFFER: keep only last maxPoints ----
        While s.Points.Count > maxPoints
            s.Points.RemoveAt(0)   ' drop oldest
        End While

        ' ---- Sliding X window so trace doesn’t squash or vanish ----
        If s.Points.Count > 0 Then
            Dim lastX As Double = s.Points(s.Points.Count - 1).XValue
            Dim window As Double = (maxPoints - 1) * xStepMinutes

            Dim xmin As Double = lastX - window
            If xmin < 0 Then xmin = 0

            ca.AxisX.Minimum = xmin
            ca.AxisX.Maximum = lastX

            ' Set a reasonable vertical grid spacing
            Dim domain As Double = ca.AxisX.Maximum - ca.AxisX.Minimum
            If domain <= 0 Then domain = xStepMinutes * 10
            ca.AxisX.Interval = domain / 10.0R
        End If

        ' ---- Y axis limits (only thing we touch vertically) ----
        If yMin.HasValue Then
            ca.AxisY.Minimum = yMin.Value
        Else
            ca.AxisY.Minimum = Double.NaN
        End If

        If yMax.HasValue Then
            ca.AxisY.Maximum = yMax.Value
        Else
            ca.AxisY.Maximum = Double.NaN
        End If

        ch.Invalidate()
    End Sub













End Class
