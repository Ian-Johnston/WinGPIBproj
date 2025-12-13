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
    Private AutoReadAction As String
    Private AutoReadValueControl As String
    Dim intervalMs As Integer = 2000
    Private originalCustomControls As List(Of Control)

    ' Current engineering unit (e.g. Ω, kΩ, MΩ, mA, µA) taken from the selected RADIO
    Private CurrentUserUnit As String = ""

    ' Current scale factor for USER-tab numeric result (e.g. Ω → kΩ)
    Private CurrentUserScale As Double = 1.0

    ' auto-scale state for USER tab
    Private CurrentUserScaleIsAuto As Boolean = False
    Private CurrentUserRangeQuery As String = ""

    ' Latest DEVICES-tab query outputs for USER tab to read
    Private USERdev1output As String = ""
    Private USERdev2output As String = ""
    Private USERdev1output2 As String = ""
    Private USERdev2output2 As String = ""
    Private OutputReceiveddev1 As Boolean = False
    Private OutputReceiveddev2 As Boolean = False

    ' --- Timer15 (QUERIESTOFILE) state ---
    Private Auto15DeviceName As String = ""
    Private Auto15ScriptBoxName As String = ""
    Private Auto15FilePathControl As String = ""
    Private Auto15ResultControl As String = ""

    ' Signals that a DEVICES-tab query has completed (reply captured)
    Private ReadOnly Dev1QueryDone As New Threading.AutoResetEvent(False)
    Private ReadOnly Dev2QueryDone As New Threading.AutoResetEvent(False)

    ' Per-device GPIB engine selection for USER tab ("standalone" or "native")
    Private GpibEngineDev1 As String = "standalone"   ' default if not specified
    Private GpibEngineDev2 As String = "standalone"   ' default if not specified

    ' NEW: stops Timer5 re-entering while a query is running
    Private UserAutoBusy As Integer = 0

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


    Private Function IsNativeEngine(deviceName As String) As Boolean
        If String.IsNullOrWhiteSpace(deviceName) Then Return False

        If deviceName.Equals("dev1", StringComparison.OrdinalIgnoreCase) Then
            Return String.Equals(GpibEngineDev1, "native", StringComparison.OrdinalIgnoreCase)
        ElseIf deviceName.Equals("dev2", StringComparison.OrdinalIgnoreCase) Then
            Return String.Equals(GpibEngineDev2, "native", StringComparison.OrdinalIgnoreCase)
        End If

        Return False
    End Function


    Private Sub NativeSend(deviceName As String, cmd As String)

        If deviceName.Equals("dev1", StringComparison.OrdinalIgnoreCase) Then
            txtq1c.Text = cmd
            RunBtns1cCore()     ' run core commands
            'dev1.SendAsync(txtq1c.Text, True)
            'txtr1astat.Text = "Send Async '" & txtq1c.Text

        ElseIf deviceName.Equals("dev2", StringComparison.OrdinalIgnoreCase) Then
            txtq2c.Text = cmd
            RunBtns2cCore()     ' run core commands
            'dev2.SendAsync(txtq2c.Text, True)
            'txtr2astat.Text = "Send Async '" & txtq2c.Text
        End If

    End Sub


    Private Function NativeQuery(deviceName As String, cmd As String) As String

        Const timeoutMs As Integer = 5000   ' adjust if needed

        If deviceName.Equals("dev1", StringComparison.OrdinalIgnoreCase) Then

            ' Arm flag for this query
            OutputReceiveddev1 = False

            txtq1a.Text = cmd
            RunBtnq1aCore()                 ' DEVICES-tab queues + sends

            ' Wait until DEVICES tab captures reply and sets flag TRUE
            Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()
            Do While Not OutputReceiveddev1 AndAlso sw.ElapsedMilliseconds < timeoutMs
                Application.DoEvents()
                Threading.Thread.Sleep(1)
            Loop

            Return If(USERdev1output, "").Trim()

        ElseIf deviceName.Equals("dev2", StringComparison.OrdinalIgnoreCase) Then

            OutputReceiveddev2 = False

            txtq2a.Text = cmd
            RunBtnq2aCore()

            Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()
            Do While Not OutputReceiveddev2 AndAlso sw.ElapsedMilliseconds < timeoutMs
                Application.DoEvents()
                Threading.Thread.Sleep(1)
            Loop

            Return If(USERdev2output, "").Trim()

        End If

        Return ""
    End Function



    Private Sub BuildCustomGuiFromText(def As String)

        ' Reset per-config runtime state
        Timer5.Enabled = False
        AutoReadDeviceName = ""
        AutoReadCommand = ""
        AutoReadResultControl = ""
        AutoReadAction = ""
        AutoReadValueControl = ""

        GroupBoxCustom.Controls.Clear()

        CurrentUserScale = 1.0   ' reset scale each time we load a layout

        Dim autoY As Integer = 10

        For Each rawLine In def.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            Dim line = rawLine.Trim()
            If line = "" OrElse line.StartsWith(";") Then Continue For

            ' Per-device GPIB engine selectors for USER tab
            ' e.g. GPIBenginedev1=native / standalone
            '      GPIBenginedev2=native / standalone
            If line.StartsWith("GPIBenginedev1=", StringComparison.OrdinalIgnoreCase) Then
                Dim val As String = line.Substring("GPIBenginedev1=".Length).Trim()
                If String.Equals(val, "native", StringComparison.OrdinalIgnoreCase) Then
                    GpibEngineDev1 = "native"
                Else
                    GpibEngineDev1 = "standalone"
                End If
                Continue For
            End If

            If line.StartsWith("GPIBenginedev2=", StringComparison.OrdinalIgnoreCase) Then
                Dim val As String = line.Substring("GPIBenginedev2=".Length).Trim()
                If String.Equals(val, "native", StringComparison.OrdinalIgnoreCase) Then
                    GpibEngineDev2 = "native"
                Else
                    GpibEngineDev2 = "standalone"
                End If
                Continue For
            End If

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

                        ' Only FuncAuto needs special behaviour when (un)checked
                        If funcKey.Equals("FuncAuto", StringComparison.OrdinalIgnoreCase) Then
                            AddHandler cb.CheckedChanged, AddressOf FuncAutoCheckbox_CheckedChanged
                        End If

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
                    ' DROPDOWN;ControlName;Caption;DeviceName;CommandPrefix;X;Y;Width;Item1,Item2,Item3...;captionpos=1/2
                    '
                    ' captionpos=1 (default) -> Caption label to the left, first combo item is blank.
                    ' captionpos=2           -> No label; caption used as first combo item (placeholder, no command).

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

                    ' Optional: caption position flag
                    ' 1 = label beside dropdown (default)
                    ' 2 = caption as first item in dropdown
                    Dim captionPos As Integer = 1

                    If parts.Length > 9 Then
                        For i As Integer = 9 To parts.Length - 1
                            Dim p As String = parts(i).Trim()
                            If p.Contains("="c) Then
                                Dim kv = p.Split("="c)
                                If kv.Length = 2 Then
                                    Dim key = kv(0).Trim().ToLower()
                                    Dim val = kv(1).Trim()

                                    If key = "captionpos" Then
                                        Integer.TryParse(val, captionPos)
                                        If captionPos <> 1 AndAlso captionPos <> 2 Then
                                            captionPos = 1
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If

                    ' Create dropdown (height not used)
                    Dim cb As New ComboBox()
                    cb.Name = ctrlName
                    cb.DropDownStyle = ComboBoxStyle.DropDownList

                    If captionPos = 1 Then
                        ' Label next to dropdown
                        Dim lbl As New Label()
                        lbl.Text = caption
                        lbl.AutoSize = True
                        lbl.Location = New Point(x, y)
                        GroupBoxCustom.Controls.Add(lbl)

                        cb.Location = New Point(x + lbl.PreferredWidth + 8, y - 3)
                    Else
                        ' No label; place combo directly at (x,y)
                        cb.Location = New Point(x, y)
                    End If

                    cb.Size = New Size(w, cb.Height)

                    ' First item:
                    '   captionPos=1 -> blank
                    '   captionPos=2 -> caption text as placeholder
                    If captionPos = 1 Then
                        cb.Items.Add("")                ' true blank
                    Else
                        cb.Items.Add(caption)           ' placeholder, acts like blank
                    End If

                    ' Add items
                    For Each it In items
                        Dim s As String = it.Trim()
                        If s <> "" Then cb.Items.Add(s)
                    Next

                    cb.SelectedIndex = 0

                    ' Tag holds Device + CommandPrefix (extra info can be added later if needed)
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
                    ' RADIO;DCRange;10 V;dev1;:SENS:VOLT:DC:RANG 10;x=10;y=70;scale=1
                    ' Or with auto-scale:
                    ' RADIO;DCIRange;Auto;dev1;:SENS:CURR:DC:RANG:AUTO ON;x=10;y=20;scale=auto|:SENS:CURR:DC:RANG?

                    If parts.Length >= 7 Then

                        Dim groupName As String = parts(1).Trim()
                        Dim caption As String = parts(2).Trim()
                        Dim deviceName As String = parts(3).Trim()
                        Dim command As String = parts(4).Trim()

                        Dim relX As Integer = 0
                        Dim relY As Integer = 0

                        ' X/Y (positional or token form)
                        If parts.Length > 5 Then ParseIntField(parts(5), relX)
                        If parts.Length > 6 Then ParseIntField(parts(6), relY)

                        ' Optional scale token
                        ' Supports:
                        '   scale=1.0        (numeric)
                        '   scale=auto       (auto-scale, no range query)
                        '   scale=auto|CMD?  (auto-scale, with per-meter range query)
                        Dim scale As Double = 1.0
                        Dim scaleIsAuto As Boolean = False
                        Dim rangeQueryForAuto As String = ""

                        If parts.Length > 7 Then
                            For i As Integer = 7 To parts.Length - 1
                                Dim tok = parts(i).Trim()
                                If tok.Contains("="c) Then
                                    Dim kv = tok.Split("="c)
                                    If kv.Length = 2 Then
                                        Dim key = kv(0).Trim().ToLower()
                                        Dim val = kv(1).Trim()

                                        If key = "scale" Then
                                            Dim lower = val.ToLowerInvariant()

                                            If lower.StartsWith("auto") Then
                                                ' auto or auto|<query>
                                                scaleIsAuto = True

                                                Dim pipeIdx As Integer = val.IndexOf("|"c)
                                                If pipeIdx >= 0 AndAlso pipeIdx < val.Length - 1 Then
                                                    rangeQueryForAuto = val.Substring(pipeIdx + 1).Trim()
                                                End If
                                            Else
                                                ' numeric scale (existing behaviour)
                                                Double.TryParse(val,
                                                    Globalization.NumberStyles.Float,
                                                    Globalization.CultureInfo.InvariantCulture,
                                                    scale)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If

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

                        ' Tag encodes:
                        '   device|command|scale            (numeric)
                        '   device|command|AUTO|rangeQuery  (auto)
                        If scaleIsAuto Then
                            rb.Tag = deviceName & "|" & command & "|" &
                                     "AUTO" & "|" & rangeQueryForAuto
                        Else
                            rb.Tag = deviceName & "|" & command & "|" &
                                     scale.ToString(Globalization.CultureInfo.InvariantCulture)
                        End If

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
                    ' BIGTEXT;ControlName;Caption(optional);x=..;y=..;f=..;w=..;h=..;border=on/off;units=on/off
                    If parts.Length >= 2 Then

                        Dim controlName As String = parts(1).Trim()
                        Dim caption As String = If(parts.Length > 2, parts(2).Trim(), "")

                        Dim x As Integer = 20
                        Dim y As Integer = 20
                        Dim fontSize As Single = 28.0F
                        Dim w As Integer = 200     ' default width
                        Dim h As Integer = 60      ' default height

                        ' Defaults:
                        Dim borderOn As Boolean = True
                        Dim unitsOn As Boolean = True

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
                                            Select Case val.Trim().ToLower()
                                                Case "0", "off", "false", "no", "none"
                                                    borderOn = False
                                                Case Else
                                                    borderOn = True
                                            End Select

                                        Case "units"
                                            Select Case val.Trim().ToLower()
                                                Case "0", "off", "false", "no", "none"
                                                    unitsOn = False
                                                Case Else
                                                    unitsOn = True
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

                        lbl.Text = "#####"

                        ' NEW: mark BIGTEXT labels that should NOT show units
                        If Not unitsOn Then
                            lbl.Tag = "BIGTEXT_UNITS_OFF"
                        End If

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
                    '       x=..;y=..;w=..;h=..;ymin=..;ymax=..;xstep=..;maxpoints=..;color=..;autoscale=..;linewidth=..

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
                    Dim maxPoints As Integer = 100           ' rolling window length
                    Dim plotColor As Color = Color.Yellow    ' default trace colour
                    Dim autoScaleY As Boolean = False        ' default off
                    Dim lineWidth As Integer = 2             ' NEW: default trace width

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

                            Case "autoscale"
                                Dim v = val.ToLower()
                                autoScaleY = (v = "yes" OrElse v = "true" OrElse v = "on" OrElse v = "1")

                            Case "linewidth"
                                Dim lw As Integer
                                If Integer.TryParse(val, lw) AndAlso lw >= 1 AndAlso lw <= 10 Then
                                    lineWidth = lw
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
                    s.BorderWidth = lineWidth   ' use configured line width
                    ch.Series.Add(s)

                    ' ---- Apply Y range if specified (fixed mode only) ----
                    If Not autoScaleY Then
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
                    Else
                        ' Let UpdateChartFromText handle Y limits
                        ca.AxisY.Minimum = Double.NaN
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
                        s0.BorderWidth = lineWidth   ' use configured line width
                        s0.MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
                        s0.MarkerSize = 3
                        s0.MarkerColor = plotColor
                    End If

                    GroupBoxCustom.Controls.Add(ch)

                    ' ---- Wire chart to the textbox that holds numeric data ----
                    Dim src = TryCast(GetControlByName(resultTarget), TextBox)
                    If src IsNot Nothing Then
                        ' Initial draw
                        UpdateChartFromText(ch, src.Text, yMin, yMax, xStep, maxPoints, autoScaleY)

                        ' Update on every text change
                        AddHandler src.TextChanged,
                            Sub(senderSrc As Object, eSrc As EventArgs)
                                Dim tbSrc = DirectCast(senderSrc, TextBox)
                                UpdateChartFromText(ch, tbSrc.Text, yMin, yMax, xStep, maxPoints, autoScaleY)
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


    Private Sub RunQueryToResult(deviceName As String,
                                 commandOrPrefix As String,
                                 resultControlName As String,
                                 Optional rawOverride As String = Nothing)

        Dim target = GetControlByName(resultControlName)
        If target Is Nothing Then
            MessageBox.Show("Result control not found: " & resultControlName)
            Exit Sub
        End If

        ' Resolve device name (we still resolve dev for standalone path + auto-scale query)
        Dim dev As IODevices.IODevice = Nothing
        Select Case deviceName.ToLowerInvariant()
            Case "dev1" : dev = dev1
            Case "dev2" : dev = dev2
        End Select

        If dev Is Nothing Then
            MessageBox.Show("Unknown device: " & deviceName)
            Exit Sub
        End If

        Dim raw As String = ""
        Dim outText As String = ""

        ' Extra: keep both numeric-only and numeric+unit text if we have them
        Dim numericText As String = ""
        Dim numericWithUnitText As String = ""

        ' =========================================================
        '   GET RAW RESPONSE (native or standalone)
        ' =========================================================
        If rawOverride IsNot Nothing Then
            raw = rawOverride
        Else
            If IsNativeEngine(deviceName) Then

                ' NATIVE PATH: use existing UI query buttons/textboxes
                raw = NativeQuery(deviceName, commandOrPrefix)

            Else

                ' STANDALONE PATH: original IODevices QueryBlocking behaviour
                Dim q As IODevices.IOQuery = Nothing
                Dim status As Integer

                Dim fullCmd As String = commandOrPrefix & TermStr2()
                status = dev.QueryBlocking(fullCmd, q, False)

                If status = 0 AndAlso q IsNot Nothing Then

                    raw = q.ResponseAsString.Trim()

                ElseIf q IsNot Nothing Then

                    ' Treat "Blocking" as a non-fatal "busy" condition – keep old value
                    If status = -1 AndAlso Not String.IsNullOrEmpty(q.errmsg) AndAlso
               q.errmsg.IndexOf("Blocking", StringComparison.OrdinalIgnoreCase) >= 0 Then

                        ' Just bail out without touching any controls
                        Exit Sub
                    End If

                    outText = "ERR " & status & ": " & q.errmsg

                Else
                    outText = "ERR " & status & " (no IOQuery)"
                End If

                ' If we already built an ERR string, skip numeric formatting and just output it
                If outText.StartsWith("ERR ", StringComparison.OrdinalIgnoreCase) Then
                    GoTo FanOut
                End If

            End If
        End If

        raw = If(raw, "").Trim()

        ' =========================================================
        '   EXISTING FORMAT / SCALE / UNITS LOGIC (unchanged)
        ' =========================================================
        Dim decCb As CheckBox = GetCheckboxFor(resultControlName, "FuncDecimal")

        Dim scale As Double = CurrentUserScale
        Dim d As Double

        If Double.TryParse(raw,
                       Globalization.NumberStyles.Float,
                       Globalization.CultureInfo.InvariantCulture,
                       d) Then

            Dim v As Double = d

            ' ============================
            '    AUTO SCALE MODE LOGIC
            ' ============================
            If CurrentUserScaleIsAuto AndAlso
           dev IsNot Nothing AndAlso
           Not String.IsNullOrWhiteSpace(CurrentUserRangeQuery) Then

                Try
                    Dim rngStr As String = ""

                    If IsNativeEngine(deviceName) Then
                        ' Range query via native engine path
                        rngStr = NativeQuery(deviceName, CurrentUserRangeQuery)
                    Else
                        ' Range query via standalone path
                        Dim rq As IODevices.IOQuery = Nothing
                        Dim st2 As Integer = dev.QueryBlocking(CurrentUserRangeQuery & TermStr2(), rq, False)
                        If st2 = 0 AndAlso rq IsNot Nothing Then
                            rngStr = rq.ResponseAsString.Trim()
                        End If
                    End If

                    Dim rngVal As Double
                    If Double.TryParse(rngStr,
                                   Globalization.NumberStyles.Float,
                                   Globalization.CultureInfo.InvariantCulture,
                                   rngVal) Then

                        Dim dynScale As Double = ComputeAutoScaleFromRange(rngVal)
                        v = d * dynScale
                    Else
                        v = d * scale    ' fallback
                    End If

                Catch
                    v = d * scale        ' fallback
                End Try

            Else
                ' FIXED SCALE MODE
                v = d * scale
            End If

            ' ============================
            '      OUTPUT FORMATTING
            ' ============================
            Dim valueStr As String

            If decCb IsNot Nothing AndAlso decCb.Checked Then
                valueStr = v.ToString("0.###############",
                                  Globalization.CultureInfo.InvariantCulture)
            ElseIf Math.Abs(v - d) > 0.0000000001R OrElse
               Math.Abs(scale - 1.0R) > 0.0000000001R OrElse
               CurrentUserScaleIsAuto Then

                valueStr = v.ToString("G",
                                  Globalization.CultureInfo.InvariantCulture)
            Else
                valueStr = raw
            End If

            numericText = valueStr

            If Not String.IsNullOrEmpty(CurrentUserUnit) Then
                numericWithUnitText = valueStr & "   " & CurrentUserUnit
            Else
                numericWithUnitText = valueStr
            End If

            outText = numericWithUnitText

        Else
            ' Non-numeric reply: pass through raw text
            outText = raw
        End If

FanOut:
        ' ============================
        '      FAN-OUT TO CONTROLS
        ' ============================
        Dim targets = GroupBoxCustom.Controls.Find(resultControlName, True)

        ' Fallback (only if the control isn't in GroupBoxCustom)
        If targets Is Nothing OrElse targets.Length = 0 Then
            targets = Me.Controls.Find(resultControlName, True)
        End If

        GroupBoxCustom.SuspendLayout()

        For Each targetCtrl In targets

            If TypeOf targetCtrl Is TextBox Then
                ' TextBox always gets value WITH units if we have it
                If numericWithUnitText <> "" Then
                    DirectCast(targetCtrl, TextBox).Text = numericWithUnitText
                Else
                    DirectCast(targetCtrl, TextBox).Text = outText
                End If

            ElseIf TypeOf targetCtrl Is Label Then
                Dim lbl = DirectCast(targetCtrl, Label)
                Dim tagStr = TryCast(lbl.Tag, String)

                ' BIGTEXT with units=off → numeric only
                If tagStr IsNot Nothing AndAlso tagStr = "BIGTEXT_UNITS_OFF" AndAlso numericText <> "" Then
                    lbl.Text = numericText
                ElseIf numericWithUnitText <> "" Then
                    lbl.Text = numericWithUnitText
                Else
                    lbl.Text = outText
                End If

            ElseIf TypeOf targetCtrl Is Panel Then
                Dim tagStr = TryCast(targetCtrl.Tag, String)
                If Not String.IsNullOrEmpty(tagStr) AndAlso
               tagStr.StartsWith("LED", StringComparison.OrdinalIgnoreCase) Then

                    SetLedStateFromText(DirectCast(targetCtrl, Panel), outText)
                End If
            End If

        Next

        GroupBoxCustom.ResumeLayout(False)

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

        Dim action = parts(0)                 ' SEND / SENDVALUE / QUERY / SENDLINES / CLEARCHART
        Dim deviceName = parts(1)             ' "dev1" / "dev2" / or ChartName for CLEARCHART
        Dim commandOrPrefix = parts(2)
        Dim valueControlName = If(parts.Length > 3, parts(3), "")
        Dim resultControlName = If(parts.Length > 4, parts(4), "")

        ' Only resolve a GPIB device for actions that actually need one
        Dim dev As IODevices.IODevice = Nothing
        If Not String.Equals(action, "CLEARCHART", StringComparison.OrdinalIgnoreCase) Then
            dev = GetDeviceByName(deviceName)
            If dev Is Nothing Then
                MessageBox.Show("Device not available: " & deviceName)
                Exit Sub
            End If
        End If

        Select Case action

            Case "SEND"
                If IsNativeEngine(deviceName) Then
                    NativeSend(deviceName, commandOrPrefix)
                Else
                    dev.SendAsync(commandOrPrefix, True)
                End If

            Case "SENDVALUE"
                Dim valCtrl = TryCast(GetControlByName(valueControlName), TextBox)
                If valCtrl Is Nothing Then
                    MessageBox.Show("Value control not found: " & valueControlName)
                    Exit Sub
                End If

                Dim cmd = commandOrPrefix & valCtrl.Text.Trim()

                If IsNativeEngine(deviceName) Then
                    NativeSend(deviceName, cmd)
                Else
                    dev.SendAsync(cmd, True)
                End If

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

                Dim valCtrl = TryCast(GetControlByName(valueControlName), TextBox)
                If valCtrl Is Nothing Then
                    MessageBox.Show("TEXTAREA not found: " & valueControlName)
                    Exit Sub
                End If

                ' ===============================
                ' NATIVE engine → use NativeSend()
                ' ===============================
                If IsNativeEngine(deviceName) Then

                    For Each raw In valCtrl.Lines
                        Dim line = raw.Trim()
                        If line = "" OrElse line.StartsWith(";"c) Then Continue For

                        ' If you want commandOrPrefix used (like standalone does), keep this:
                        Dim finalCmd As String = commandOrPrefix & line & TermStr2()

                        ' If you do NOT want a prefix (i.e. lines already contain full commands),
                        ' then use instead:
                        ' Dim finalCmd As String = line & TermStr2()

                        NativeSend(deviceName, finalCmd)
                    Next

                    Exit Sub
                End If

                ' ===============================
                ' STANDALONE engine → direct send
                ' ===============================
                For Each raw In valCtrl.Lines
                    Dim line = raw.Trim()
                    If line = "" OrElse line.StartsWith(";"c) Then Continue For

                    dev.SendAsync(commandOrPrefix & line & TermStr2(), True)
                Next

            Case "CLEARCHART"
                ' deviceName here is actually the ChartName (e.g. "ChartDMM")
                Dim ch = TryCast(GetControlByName(deviceName), DataVisualization.Charting.Chart)
                If ch IsNot Nothing AndAlso ch.Series.Count > 0 Then
                    ch.Series(0).Points.Clear()
                End If

            Case "QUERIESTOFILE"
                ' commandOrPrefix = TEXTAREA name (e.g. ScriptBox2)
                ' valueControlName = file path textbox (optional; may be blank)
                ' resultControlName = optional "latest reply" display

                Dim scriptBoxName As String = commandOrPrefix

                ' Run once immediately
                RunQueriesToFileFromTextArea(deviceName, scriptBoxName, valueControlName, resultControlName)

                ' Auto checkbox support (checkbox key = ScriptBox2)
                Dim autoCb As CheckBox = GetCheckboxFor(scriptBoxName, "FuncAuto")

                If autoCb IsNot Nothing AndAlso autoCb.Checked Then
                    Dim intervalMs As Integer = 2000

                    Dim param As String = GetCheckboxParam(autoCb)
                    If Not String.IsNullOrEmpty(param) Then
                        Dim numeric As String = ""
                        For Each ch As Char In param
                            If Char.IsDigit(ch) OrElse ch = "."c Then numeric &= ch
                        Next

                        Dim secs As Double
                        If Double.TryParse(numeric,
                                           Globalization.NumberStyles.Float,
                                           Globalization.CultureInfo.InvariantCulture,
                                           secs) AndAlso secs > 0.0R Then
                            intervalMs = CInt(secs * 1000.0R)
                        End If
                    End If

                    ' Store timer state
                    Auto15DeviceName = deviceName
                    Auto15ScriptBoxName = scriptBoxName
                    Auto15FilePathControl = valueControlName
                    Auto15ResultControl = resultControlName

                    Timer15.Interval = intervalMs
                    Timer15.Enabled = True
                Else
                    Timer15.Enabled = False
                End If

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

        Dim device As String = tagParts(0)
        Dim cmdOn As String = tagParts(1)
        Dim cmdOff As String = tagParts(2)
        Dim state As Integer
        Integer.TryParse(tagParts(3), state)   ' 0 = OFF, 1 = ON

        ' Pick dev1/dev2 + native/standalone flag
        Dim dev As IODevices.IODevice = Nothing
        Dim useNative As Boolean = False

        Select Case device.ToLowerInvariant()
            Case "dev1"
                dev = dev1
                useNative = String.Equals(GpibEngineDev1, "native", StringComparison.OrdinalIgnoreCase)
            Case "dev2"
                dev = dev2
                useNative = String.Equals(GpibEngineDev2, "native", StringComparison.OrdinalIgnoreCase)
        End Select
        If dev Is Nothing Then Exit Sub

        If state = 0 Then
            ' Turn ON
            If useNative Then
                NativeSend(device, cmdOn & TermStr2())
            Else
                dev.SendAsync(cmdOn & TermStr2(), True)
            End If

            b.BackColor = Color.LimeGreen
            b.ForeColor = Color.Black
            state = 1

        Else
            ' Turn OFF
            If useNative Then
                NativeSend(device, cmdOff & TermStr2())
            Else
                dev.SendAsync(cmdOff & TermStr2(), True)
            End If

            b.BackColor = SystemColors.Control
            b.ForeColor = SystemColors.ControlText
            state = 0
        End If

        ' Update Tag with new state
        b.Tag = device & "|" & cmdOn & "|" & cmdOff & "|" & state
    End Sub


    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick

        Dim autoCb As CheckBox = GetCheckboxFor(AutoReadResultControl, "FuncAuto")
        If autoCb Is Nothing OrElse Not autoCb.Checked Then
            Timer5.Enabled = False
            Exit Sub
        End If

        If String.IsNullOrEmpty(AutoReadDeviceName) OrElse String.IsNullOrEmpty(AutoReadCommand) Then
            Timer5.Enabled = False
            Exit Sub
        End If

        ' Prevent overlap (timer ticks while query still running)
        If Threading.Interlocked.Exchange(UserAutoBusy, 1) = 1 Then Exit Sub

        Dim devName As String = AutoReadDeviceName
        Dim cmd As String = AutoReadCommand
        Dim resultName As String = AutoReadResultControl

        Threading.Tasks.Task.Run(Sub()

                                     Dim raw As String = QueryRawOnly(devName, cmd)

                                     ' If busy/no update, just release lock
                                     If raw = "" Then
                                         Threading.Interlocked.Exchange(UserAutoBusy, 0)
                                         Return
                                     End If

                                     Me.BeginInvoke(Sub()
                                                        RunQueryToResult(devName, cmd, resultName, raw)
                                                        Threading.Interlocked.Exchange(UserAutoBusy, 0)
                                                    End Sub)

                                 End Sub)

    End Sub


    Private Sub Dropdown_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim cb = TryCast(sender, ComboBox)
        If cb Is Nothing Then Exit Sub

        Dim meta As String = TryCast(cb.Tag, String)
        If String.IsNullOrEmpty(meta) Then Exit Sub

        Dim parts = meta.Split("|"c)
        If parts.Length < 2 Then Exit Sub

        Dim deviceName As String = parts(0).Trim()
        Dim commandPrefix As String = parts(1).TrimEnd()

        ' First entry (index 0) is always a placeholder → do nothing
        If cb.SelectedIndex <= 0 Then Exit Sub

        Dim selected As String = ""
        If cb.SelectedItem IsNot Nothing Then selected = cb.SelectedItem.ToString().Trim()
        If selected = "" Then Exit Sub

        ' Build command
        Dim cmd As String = commandPrefix.TrimEnd() & " " & selected

        ' Resolve device + engine mode
        Dim dev As IODevices.IODevice = Nothing
        Dim useNative As Boolean = False

        Select Case deviceName.ToLowerInvariant()
            Case "dev1"
                dev = dev1
                useNative = String.Equals(GpibEngineDev1, "native", StringComparison.OrdinalIgnoreCase)

            Case "dev2"
                dev = dev2
                useNative = String.Equals(GpibEngineDev2, "native", StringComparison.OrdinalIgnoreCase)

            Case Else
                MessageBox.Show("Unknown device in DROPDOWN: " & deviceName)
                Exit Sub
        End Select

        If dev Is Nothing Then
            MessageBox.Show("Device not available: " & deviceName)
            Exit Sub
        End If

        ' Send (native uses the DEVICES-tab core send routine, not PerformClick)
        If useNative Then
            NativeSend(deviceName, cmd)     ' <-- calls RunBtns1cCore / RunBtns2cCore
        Else
            dev.SendAsync(cmd, True)
        End If
    End Sub



    Private Sub Radio_CheckedChanged(sender As Object, e As EventArgs)
        Dim rb As RadioButton = TryCast(sender, RadioButton)
        If rb Is Nothing Then Exit Sub
        If Not rb.Checked Then Exit Sub

        Dim meta As String = TryCast(rb.Tag, String)
        If String.IsNullOrEmpty(meta) Then Exit Sub

        Dim parts() As String = meta.Split("|"c)
        If parts.Length < 2 Then Exit Sub

        Dim deviceName As String = parts(0)
        Dim command As String = parts(1)

        ' ===============================
        '   RESET SCALE / AUTO STATE
        ' ===============================
        CurrentUserScaleIsAuto = False
        CurrentUserRangeQuery = ""
        Dim scale As Double = 1.0

        ' Tag formats:
        '   device|command|scale
        '   device|command|AUTO|rangeQuery
        If parts.Length >= 3 Then
            If String.Equals(parts(2), "AUTO", StringComparison.OrdinalIgnoreCase) Then
                CurrentUserScaleIsAuto = True
                If parts.Length >= 4 Then
                    CurrentUserRangeQuery = parts(3).Trim()
                End If
            Else
                Double.TryParse(parts(2),
                Globalization.NumberStyles.Float,
                Globalization.CultureInfo.InvariantCulture,
                scale)
            End If
        End If

        CurrentUserScale = scale

        ' ===============================
        '   MODE vs RANGE DETECTION
        ' ===============================
        Dim cap As String = rb.Text.Trim()
        Dim lastSpace As Integer = cap.LastIndexOf(" "c)
        Dim isModeRadio As Boolean = (lastSpace < 0)

        If isModeRadio Then
            CurrentUserUnit = ""

            ' Uncheck all other radios in all dynamic groups
            For Each ctrl As Control In GroupBoxCustom.Controls
                Dim gb As GroupBox = TryCast(ctrl, GroupBox)
                If gb Is Nothing Then Continue For

                For Each child As Control In gb.Controls
                    Dim r As RadioButton = TryCast(child, RadioButton)
                    If r Is Nothing Then Continue For
                    If r Is rb Then Continue For
                    r.Checked = False
                Next
            Next
        Else
            Dim unit As String = ""
            If lastSpace >= 0 AndAlso lastSpace < cap.Length - 1 Then
                unit = cap.Substring(lastSpace + 1).Trim()
            End If
            CurrentUserUnit = unit
        End If

        ' ===============================
        '   SEND THE RADIO COMMAND
        ' ===============================
        Dim dev As IODevices.IODevice = Nothing
        Dim useNative As Boolean = False

        Select Case deviceName.ToLowerInvariant()
            Case "dev1"
                dev = dev1
                useNative = String.Equals(GpibEngineDev1, "native", StringComparison.OrdinalIgnoreCase)
            Case "dev2"
                dev = dev2
                useNative = String.Equals(GpibEngineDev2, "native", StringComparison.OrdinalIgnoreCase)
        End Select

        If dev Is Nothing Then
            MessageBox.Show("Unknown device in RADIO: " & deviceName)
            Exit Sub
        End If

        If useNative Then
            NativeSend(deviceName, command)   ' <-- deferred click inside NativeSend
        Else
            dev.SendAsync(command, True)
        End If

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

        Dim deviceName As String = parts(0).Trim()
        Dim commandPrefix As String = parts(1).Trim()

        Dim scale As Double = 1.0
        Double.TryParse(parts(2), Globalization.NumberStyles.Float,
                    Globalization.CultureInfo.InvariantCulture, scale)

        Dim stepVal As Integer = 1
        Integer.TryParse(parts(4), stepVal)
        If stepVal < 1 Then stepVal = 1

        ' Snap again on mouse-up
        Dim raw As Integer = tb.Value
        Dim snapped As Integer = CInt(Math.Round(raw / CDbl(stepVal))) * stepVal
        If snapped < tb.Minimum Then snapped = tb.Minimum
        If snapped > tb.Maximum Then snapped = tb.Maximum
        If snapped <> tb.Value Then tb.Value = snapped

        Dim scaledValue As Double = snapped * scale

        Dim cmd As String =
        commandPrefix.TrimEnd() & " " &
        scaledValue.ToString("G", Globalization.CultureInfo.InvariantCulture)

        ' Resolve device + engine mode
        Dim dev As IODevices.IODevice = Nothing
        Dim useNative As Boolean = False

        Select Case deviceName.ToLowerInvariant()
            Case "dev1"
                dev = dev1
                useNative = String.Equals(GpibEngineDev1, "native", StringComparison.OrdinalIgnoreCase)

            Case "dev2"
                dev = dev2
                useNative = String.Equals(GpibEngineDev2, "native", StringComparison.OrdinalIgnoreCase)

            Case Else
                MessageBox.Show("Unknown device in SLIDER: " & deviceName)
                Exit Sub
        End Select

        If dev Is Nothing Then
            MessageBox.Show("Device not available: " & deviceName)
            Exit Sub
        End If

        ' Send (native uses the DEVICES-tab core send routine)
        If useNative Then
            NativeSend(deviceName, cmd)     ' <-- calls RunBtns1cCore / RunBtns2cCore
        Else
            dev.SendAsync(cmd, True)
        End If
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
        Dim nud = TryCast(sender, NumericUpDown)
        If nud Is Nothing Then Exit Sub

        Dim tagStr = TryCast(nud.Tag, String)
        If String.IsNullOrEmpty(tagStr) Then Exit Sub

        Dim parts = tagStr.Split("|"c)
        If parts.Length < 3 Then Exit Sub

        Dim deviceName As String = parts(0).Trim()
        Dim commandPrefix As String = parts(1).Trim()

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

        ' Resolve device + engine mode
        Dim dev As IODevices.IODevice = Nothing
        Dim useNative As Boolean = False

        Select Case deviceName.ToLowerInvariant()
            Case "dev1"
                dev = dev1
                useNative = String.Equals(GpibEngineDev1, "native", StringComparison.OrdinalIgnoreCase)

            Case "dev2"
                dev = dev2
                useNative = String.Equals(GpibEngineDev2, "native", StringComparison.OrdinalIgnoreCase)

            Case Else
                MessageBox.Show("Unknown device in SPINNER: " & deviceName)
                Exit Sub
        End Select

        If dev Is Nothing Then
            MessageBox.Show("Device not available: " & deviceName)
            Exit Sub
        End If

        ' === FIRST USER CHANGE ===
        If initFlag = "0" Then
            parts(3) = "1"
            nud.Tag = String.Join("|", parts)

            ' Set spinner to the configured min (this will re-enter this handler once)
            nud.Value = minValFromTag

            Dim scaledFirst As Double = CDbl(minValFromTag) * scale
            Dim firstStr As String = scaledFirst.ToString("G", Globalization.CultureInfo.InvariantCulture)
            Dim firstCmd As String = commandPrefix.TrimEnd() & " " & firstStr

            If useNative Then
                NativeSend(deviceName, firstCmd)   ' <-- IMPORTANT: use the working native path
            Else
                dev.SendAsync(firstCmd, True)
            End If

            Exit Sub
        End If

        ' === NORMAL CHANGES AFTER FIRST ===
        Dim scaledValue As Double = CDbl(nud.Value) * scale
        Dim valueStr As String = scaledValue.ToString("G", Globalization.CultureInfo.InvariantCulture)
        Dim cmd As String = commandPrefix.TrimEnd() & " " & valueStr

        If useNative Then
            NativeSend(deviceName, cmd)            ' <-- IMPORTANT: use the working native path
        Else
            dev.SendAsync(cmd, True)
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
        Dim b As Button = TryCast(sender, Button)
        If b Is Nothing Then Exit Sub

        Dim metaParts = CStr(b.Tag).Split("|"c)
        If metaParts.Length < 2 Then Exit Sub

        Dim areaName As String = metaParts(0).Trim()
        Dim deviceName As String = metaParts(1).Trim()

        Dim tb As TextBox = TryCast(GetControlByName(areaName), TextBox)
        If tb Is Nothing Then
            MessageBox.Show("TEXTAREA not found: " & areaName)
            Exit Sub
        End If

        ' Resolve device + engine mode
        Dim dev As IODevices.IODevice = Nothing
        Dim useNative As Boolean = False

        Select Case deviceName.ToLowerInvariant()
            Case "dev1"
                dev = dev1
                useNative = String.Equals(GpibEngineDev1, "native", StringComparison.OrdinalIgnoreCase)

            Case "dev2"
                dev = dev2
                useNative = String.Equals(GpibEngineDev2, "native", StringComparison.OrdinalIgnoreCase)

            Case Else
                MessageBox.Show("Invalid device: " & deviceName)
                Exit Sub
        End Select

        If dev Is Nothing Then
            MessageBox.Show("Device not available: " & deviceName)
            Exit Sub
        End If

        For Each rawLine In tb.Lines
            Dim line = rawLine.Trim()
            If line = "" Then Continue For
            If line.StartsWith(";"c) Then Continue For

            Dim finalCmd As String = line & TermStr2()

            If useNative Then
                ' IMPORTANT: don't PerformClick from USER tab; use the working native path
                NativeSend(deviceName, finalCmd)
            Else
                dev.SendAsync(finalCmd, True)
            End If
        Next
    End Sub


    Private Sub UpdateChartFromText(ch As DataVisualization.Charting.Chart,
                    text As String,
                    yMin As Double?,
                    yMax As Double?,
                    xStepMinutes As Double,
                    maxPoints As Integer,
                    autoScaleY As Boolean)

        If ch Is Nothing Then Exit Sub
        If maxPoints <= 0 Then maxPoints = 100
        If xStepMinutes <= 0 Then xStepMinutes = 1.0R

        ' Safety limit for chart values (protects against OVERLOAD / huge numbers)
        Const AXIS_LIMIT As Double = 1.0E+28

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

                ' Ignore NaN / Infinity / huge “OVERLOAD” values
                If Not Double.IsNaN(d) AndAlso
               Not Double.IsInfinity(d) AndAlso
               Math.Abs(d) <= AXIS_LIMIT Then

                    values.Add(d)
                End If
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

        ' ---- Sliding X window with fixed width (new points enter from the right) ----
        If s.Points.Count > 0 Then
            Dim lastX As Double = s.Points(s.Points.Count - 1).XValue
            Dim window As Double = (maxPoints - 1) * xStepMinutes

            Dim xmin As Double = lastX - window
            Dim xmax As Double = lastX

            ca.AxisX.Minimum = xmin
            ca.AxisX.Maximum = xmax

            ' Use the window for grid spacing
            Dim domain As Double = window
            If domain <= 0 Then domain = xStepMinutes * 10
            ca.AxisX.Interval = domain / 10.0R
        End If

        ' ---- Y axis limits ----
        If autoScaleY Then
            ' Compute min/max from current points, ignoring crazy values
            If s.Points.Count > 0 Then

                Dim anyValid As Boolean = False
                Dim minY As Double = Double.MaxValue
                Dim maxY As Double = Double.MinValue

                For Each pt As DataVisualization.Charting.DataPoint In s.Points
                    If pt.YValues Is Nothing OrElse pt.YValues.Length = 0 Then Continue For

                    Dim v As Double = pt.YValues(0)

                    ' Ignore NaN / Infinity / values outside Decimal-safe range
                    If Double.IsNaN(v) OrElse Double.IsInfinity(v) Then Continue For
                    If Math.Abs(v) > AXIS_LIMIT Then Continue For

                    anyValid = True
                    If v < minY Then minY = v
                    If v > maxY Then maxY = v
                Next

                If Not anyValid Then
                    ca.AxisY.Minimum = Double.NaN
                    ca.AxisY.Maximum = Double.NaN
                ElseIf minY = maxY Then
                    ' Flat line: add a small pad
                    Dim pad As Double = If(Math.Abs(minY) < 1.0R, 0.1R, Math.Abs(minY) * 0.1R)
                    Dim yLo As Double = minY - pad
                    Dim yHi As Double = maxY + pad

                    ' Clamp to safe range
                    If yLo < -AXIS_LIMIT Then yLo = -AXIS_LIMIT
                    If yHi > AXIS_LIMIT Then yHi = AXIS_LIMIT

                    ca.AxisY.Minimum = yLo
                    ca.AxisY.Maximum = yHi
                Else
                    Dim range As Double = maxY - minY
                    Dim pad As Double = range * 0.05R
                    Dim yLo As Double = minY - pad
                    Dim yHi As Double = maxY + pad

                    ' Clamp to safe range
                    If yLo < -AXIS_LIMIT Then yLo = -AXIS_LIMIT
                    If yHi > AXIS_LIMIT Then yHi = AXIS_LIMIT

                    ca.AxisY.Minimum = yLo
                    ca.AxisY.Maximum = yHi
                End If
            Else
                ca.AxisY.Minimum = Double.NaN
                ca.AxisY.Maximum = Double.NaN
            End If
        Else
            ' Fixed/manual Y range (existing behaviour)
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
        End If

        ch.Invalidate()
    End Sub


    Private Sub FuncAutoCheckbox_CheckedChanged(sender As Object, e As EventArgs)
        Dim cb = TryCast(sender, CheckBox)
        If cb Is Nothing Then Exit Sub

        Dim tagStr = TryCast(cb.Tag, String)
        If String.IsNullOrEmpty(tagStr) Then Exit Sub

        Dim parts = tagStr.Split("|"c)
        If parts.Length < 2 Then Exit Sub

        Dim resultName As String = parts(0).Trim()

        ' If unchecked → stop auto-read for this result (if it's the active one)
        If Not cb.Checked Then
            If String.Equals(AutoReadResultControl, resultName, StringComparison.OrdinalIgnoreCase) Then
                Timer5.Enabled = False
            End If
            Exit Sub
        End If

        ' Checked: find the QUERY button that feeds this result textbox
        For Each ctrl As Control In GroupBoxCustom.Controls
            Dim btn = TryCast(ctrl, Button)
            If btn Is Nothing Then Continue For

            Dim meta = TryCast(btn.Tag, String)
            If String.IsNullOrEmpty(meta) Then Continue For

            Dim bp = meta.Split("|"c)
            If bp.Length < 5 Then Continue For

            Dim action = bp(0)
            Dim deviceName = bp(1)
            Dim commandOrPrefix = bp(2)
            Dim resultControlName = bp(4)

            If String.Equals(action, "QUERY", StringComparison.OrdinalIgnoreCase) AndAlso
           String.Equals(resultControlName, resultName, StringComparison.OrdinalIgnoreCase) Then

                ' Simulate a click on this QUERY button
                CustomButton_Click(btn, EventArgs.Empty)
                Exit For
            End If
        Next
    End Sub


    ' Compute a display scale factor from a numeric range value
    ' Generic engineering-prefix style:
    '   very small ranges -> n / µ / m
    '   mid ranges        -> base units
    '   large ranges      -> k / M
    Private Function ComputeAutoScaleFromRange(rangeVal As Double) As Double
        Dim a As Double = Math.Abs(rangeVal)

        If a <= 0 OrElse Double.IsNaN(a) OrElse Double.IsInfinity(a) Then
            Return 1.0
        End If

        ' Below 1 µ → n-units
        If a < 0.000001 Then
            Return 1000000000.0     ' n
        End If

        ' 1 µ .. < 1 m → µ-units
        If a < 0.001 Then
            Return 1000000.0     ' µ
        End If

        ' 1 m .. < 1 → m-units
        If a < 1.0 Then
            Return 1000.0     ' m
        End If

        ' 1 .. < 1 k → base units
        If a < 1000.0 Then
            Return 1.0       ' base
        End If

        ' 1 k .. < 1 M → k-units
        If a < 1000000.0 Then
            Return 0.001    ' k
        End If

        ' >= 1 M → M-units
        Return 0.000001        ' M
    End Function


    Private Sub RunQueryToFile(deviceName As String,
                           queryCommand As String,
                           filePathControlName As String,
                           Optional resultControlName As String = "")

        ' queryCommand is a SINGLE query, e.g. ":READ?" or ":SENS:RES:RANG?"
        If String.IsNullOrWhiteSpace(queryCommand) Then Exit Sub

        Dim filePath As String = ""
        If Not String.IsNullOrWhiteSpace(filePathControlName) Then
            Dim fpCtrl = TryCast(GetControlByName(filePathControlName), TextBox)
            If fpCtrl IsNot Nothing Then filePath = fpCtrl.Text.Trim()
        End If

        ' Default file name if none provided
        If String.IsNullOrWhiteSpace(filePath) Then
            filePath = "QueriesToFile.csv"
        End If

        ' --- Run query (native or standalone) ---
        Dim dev As IODevices.IODevice = GetDeviceByName(deviceName)
        If dev Is Nothing Then
            MessageBox.Show("Unknown device: " & deviceName)
            Exit Sub
        End If

        Dim raw As String = ""

        If IsNativeEngine(deviceName) Then
            raw = NativeQuery(deviceName, queryCommand)
        Else
            Dim q As IODevices.IOQuery = Nothing
            Dim st As Integer = dev.QueryBlocking(queryCommand & TermStr2(), q, False)
            If st = 0 AndAlso q IsNot Nothing Then
                raw = q.ResponseAsString
            Else
                raw = "ERR " & st & If(q IsNot Nothing AndAlso Not String.IsNullOrEmpty(q.errmsg), ": " & q.errmsg, "")
            End If
        End If

        raw = If(raw, "").Trim()

        ' --- Write line to file (force into My Documents\WinGPIBdata unless rooted) ---
        Try
            Dim baseDir As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WinGPIBdata")
            If Not Directory.Exists(baseDir) Then Directory.CreateDirectory(baseDir)

            If Not Path.IsPathRooted(filePath) Then
                filePath = Path.Combine(baseDir, filePath)
            End If

            Dim dir As String = Path.GetDirectoryName(filePath)
            If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim ts As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
            Dim cmdEsc As String = queryCommand.Replace("""", """""")
            Dim respEsc As String = raw.Replace("""", """""")

            Dim line As String = $"""{ts}"",""{deviceName}"",""{cmdEsc}"",""{respEsc}"""

            SyncLock Me
                System.IO.File.AppendAllText(filePath, line & Environment.NewLine)
            End SyncLock

        Catch ex As Exception
            MessageBox.Show("QUERIESTOFILE write failed:" & vbCrLf & ex.Message)
        End Try

        ' Optional: update a result control with the latest reply
        If Not String.IsNullOrWhiteSpace(resultControlName) Then
            Dim targets() As Control = Me.Controls.Find(resultControlName, True)
            For Each c As Control In targets
                If TypeOf c Is TextBox Then
                    DirectCast(c, TextBox).Text = raw
                ElseIf TypeOf c Is Label Then
                    DirectCast(c, Label).Text = raw
                End If
            Next
        End If

    End Sub


    Private Sub RunSendToFile(deviceName As String,
                          cmdToSend As String,
                          filePathControlName As String,
                          Optional resultControlName As String = "")

        ' Resolve device
        Dim dev As IODevices.IODevice = GetDeviceByName(deviceName)
        If dev Is Nothing Then
            MessageBox.Show("Unknown device: " & deviceName)
            Exit Sub
        End If

        ' --- SEND (native or standalone) ---
        If IsNativeEngine(deviceName) Then
            NativeSend(deviceName, cmdToSend)   ' your working native send path
        Else
            dev.SendAsync(cmdToSend & TermStr2(), True)
        End If

        ' --- Determine file path (same rules as your RunQueryToFile) ---
        Dim filePath As String = ""
        If Not String.IsNullOrWhiteSpace(filePathControlName) Then
            Dim fpCtrl = TryCast(GetControlByName(filePathControlName), TextBox)
            If fpCtrl IsNot Nothing Then filePath = fpCtrl.Text.Trim()
        End If

        If String.IsNullOrWhiteSpace(filePath) Then
            filePath = "QueriesToFile.csv"
        End If

        ' --- Write CSV line with BLANK response ---
        Try
            Dim baseDir As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WinGPIBdata")
            If Not Directory.Exists(baseDir) Then Directory.CreateDirectory(baseDir)

            If Not System.IO.Path.IsPathRooted(filePath) Then
                filePath = System.IO.Path.Combine(baseDir, filePath)
            End If

            Dim dir As String = System.IO.Path.GetDirectoryName(filePath)
            If Not String.IsNullOrEmpty(dir) AndAlso Not System.IO.Directory.Exists(dir) Then
                System.IO.Directory.CreateDirectory(dir)
            End If

            Dim ts As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss",
                                                 Globalization.CultureInfo.InvariantCulture)

            Dim cmdEsc As String = cmdToSend.Replace("""", """""")
            Dim respEsc As String = "" ' no reply expected

            Dim lineOut As String = $"""{ts}"",""{deviceName}"",""{cmdEsc}"",""{respEsc}"""

            SyncLock Me
                System.IO.File.AppendAllText(filePath, lineOut & Environment.NewLine)
            End SyncLock

        Catch ex As Exception
            MessageBox.Show("SEND(no reply) write failed:" & vbCrLf & ex.Message)
        End Try

        ' Optional UI update: you can show "(noreply)" or blank
        If Not String.IsNullOrWhiteSpace(resultControlName) Then
            Dim targets() As Control = Me.Controls.Find(resultControlName, True)
            For Each c As Control In targets
                If TypeOf c Is TextBox Then
                    DirectCast(c, TextBox).Text = ""   ' no reply
                ElseIf TypeOf c Is Label Then
                    DirectCast(c, Label).Text = ""
                End If
            Next
        End If

    End Sub


    Private Function MakeSafeFileName(input As String) As String
        If input Is Nothing Then Return ""

        Dim s As String = input.Trim()

        ' 1) Remove invalid *path* characters first (prevents Path.GetFileName throwing)
        For Each ch As Char In System.IO.Path.GetInvalidPathChars()
            s = s.Replace(ch, "_"c)
        Next

        ' 2) Now it is safe to call GetFileName (also handles if user pasted a full path)
        Try
            s = System.IO.Path.GetFileName(s)
        Catch
            ' If anything weird still slips through, just use the raw string after path-char cleanup
            ' (and continue sanitising below)
        End Try

        ' 3) Remove invalid *file name* characters
        For Each ch As Char In System.IO.Path.GetInvalidFileNameChars()
            s = s.Replace(ch, "_"c)
        Next

        ' 4) Extra hardening: remove control chars explicitly
        Dim sb As New System.Text.StringBuilder()
        For Each ch As Char In s
            If Not Char.IsControl(ch) Then sb.Append(ch)
        Next
        s = sb.ToString().Trim()

        Return s
    End Function


    Private Sub RunQueriesToFileFromTextArea(deviceName As String,
                                        scriptControlName As String,
                                        filePathControlName As String,
                                        Optional resultControlName As String = "")

        Dim scriptTb As TextBox = TryCast(GetControlByName(scriptControlName), TextBox)
        If scriptTb Is Nothing Then
            Timer15.Enabled = False
            Exit Sub
        End If

        For Each rawLine As String In scriptTb.Lines
            Dim line As String = rawLine.Trim()
            If line = "" Then Continue For
            If line.StartsWith(";"c) Then Continue For

            ' Detect our explicit marker
            Dim noReply As Boolean = False
            If line.StartsWith("(noreply)", StringComparison.OrdinalIgnoreCase) Then
                noReply = True
                line = line.Substring(9).Trim() ' len("(noreply)") = 9
            End If

            If line = "" Then Continue For

            If noReply Then
                ' SEND only, but still log to CSV with blank response
                RunSendToFile(deviceName, line, filePathControlName, resultControlName)
            Else
                ' QUERY + log reply to CSV
                RunQueryToFile(deviceName, line, filePathControlName, resultControlName)
            End If
        Next

    End Sub


    Private Sub Timer15_Tick(sender As Object, e As EventArgs) Handles Timer15.Tick
        Dim autoCb As CheckBox = GetCheckboxFor(Auto15ScriptBoxName, "FuncAuto")
        If autoCb Is Nothing OrElse Not autoCb.Checked Then
            Timer15.Enabled = False
            Exit Sub
        End If

        RunQueriesToFileFromTextArea(Auto15DeviceName,
                                     Auto15ScriptBoxName,
                                     Auto15FilePathControl,
                                     Auto15ResultControl)
    End Sub


    ' Returns just the raw reply (no UI updates)
    Private Function QueryRawOnly(deviceName As String, cmd As String) As String

        If IsNativeEngine(deviceName) Then
            ' Runs your existing native path (but we’ll call it off-thread)
            Return NativeQuery(deviceName, cmd)
        End If

        Dim dev As IODevices.IODevice = GetDeviceByName(deviceName)
        If dev Is Nothing Then Return "ERR Unknown device: " & deviceName

        Dim q As IODevices.IOQuery = Nothing
        Dim st As Integer = dev.QueryBlocking(cmd & TermStr2(), q, False)

        If st = 0 AndAlso q IsNot Nothing Then
            Return q.ResponseAsString.Trim()
        End If

        If q IsNot Nothing Then
            If st = -1 AndAlso Not String.IsNullOrEmpty(q.errmsg) AndAlso
           q.errmsg.IndexOf("Blocking", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return ""   ' keep old value (busy)
            End If
            Return "ERR " & st & ": " & q.errmsg
        End If

        Return "ERR " & st & " (no IOQuery)"
    End Function




End Class
