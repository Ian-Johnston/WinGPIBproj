' User customizeable tab - Parsing

Imports System.ComponentModel


Partial Class Formtest


    Private Sub BuildCustomGuiFromText(def As String)

        ' Reset per-config runtime state
        Timer5.Enabled = False
        AutoReadDeviceName = ""
        AutoReadCommand = ""
        AutoReadResultControl = ""
        'AutoReadAction = ""
        AutoReadValueControl = ""

        Timer16.Enabled = False
        AutoReadDeviceName2 = ""
        AutoReadCommand2 = ""
        AutoReadResultControl2 = ""
        'AutoReadAction2 = ""
        AutoReadValueControl2 = ""

        Threading.Interlocked.Exchange(UserAutoBusy2, 0)

        GroupBoxCustom.Controls.Clear()
        LuaScriptsByName.Clear()

        CurrentUserScale = 1.0   ' reset scale each time we load a layout

        Dim autoY As Integer = 10

        For Each rawLine In def.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            Dim line = rawLine.Trim()
            If line = "" OrElse line.StartsWith(";") Then Continue For

            ' ================= LUA SCRIPT BLOCK =================
            If inLuaScript Then
                If line.StartsWith("LUASCRIPTEND", StringComparison.OrdinalIgnoreCase) Then
                    LuaScriptsByName(luaScriptName) = String.Join(vbCrLf, luaLines)
                    luaLines.Clear()
                    inLuaScript = False
                Else
                    ' store ORIGINAL rawLine to preserve indentation
                    luaLines.Add(rawLine)
                End If
                Continue For
            End If
            ' ===================================================

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

            ' LUASCRIPTBEGIN;name=MyScript
            If line.StartsWith("LUASCRIPTBEGIN", StringComparison.OrdinalIgnoreCase) Then
                luaScriptName = GetParam(line, "name")
                luaLines.Clear()
                inLuaScript = True
                Continue For
            End If

            Dim parts = line.Split(";"c)
            If parts.Length < 2 Then Continue For

            Dim typeStr = parts(0).Trim().ToUpperInvariant()

            Select Case typeStr

                Case "GROUPBOX"

                    ' Parse named parameters
                    Dim kv = ParseNamedParams(parts)

                    If kv.ContainsKey("caption") Then
                        GroupBoxCustom.Text = kv("caption")
                    End If

                    ' Optional future support
                    If kv.ContainsKey("x") Then GroupBoxCustom.Left = CInt(kv("x"))
                    If kv.ContainsKey("y") Then GroupBoxCustom.Top = CInt(kv("y"))
                    If kv.ContainsKey("w") Then GroupBoxCustom.Width = CInt(kv("w"))
                    If kv.ContainsKey("h") Then GroupBoxCustom.Height = CInt(kv("h"))

                    Continue For


                Case "TEXTBOX"
                    If parts.Length < 2 Then Continue For

                    Dim tbName As String = ""
                    Dim labelText As String = ""

                    Dim x As Integer = 10
                    Dim y As Integer = autoY
                    Dim w As Integer = 120
                    Dim h As Integer = 22
                    Dim isReadOnly As Boolean = False
                    Dim hasExplicitCoords As Boolean = False

                    ' -----------------------------
                    ' Detect all-named style early
                    ' -----------------------------
                    Dim firstLooksNamed As Boolean =
                        parts.Length >= 2 AndAlso parts(1).IndexOf("="c) >= 0

                    ' -----------------------------
                    ' Positional form (legacy)
                    ' TEXTBOX;Name;Caption;X;Y;[W;H;readonly=..]
                    ' -----------------------------
                    If Not firstLooksNamed AndAlso parts.Length >= 5 AndAlso Not parts(3).Contains("="c) Then

                        tbName = parts(1).Trim()
                        labelText = parts(2).Trim()

                        ParseIntField(parts(3), x)
                        ParseIntField(parts(4), y)
                        hasExplicitCoords = True

                        If parts.Length >= 6 Then ParseIntField(parts(5), w)
                        If parts.Length >= 7 Then ParseIntField(parts(6), h)

                        ' optional readonly (positional)
                        If parts.Length >= 8 Then
                            ParseBoolField(parts(7), isReadOnly)
                        End If

                    Else
                        ' -----------------------------
                        ' Named / hybrid token form
                        ' TEXTBOX;name=..;caption=..;x=..;y=..;...
                        ' OR TEXTBOX;Name;caption=..;x=..;...
                        ' -----------------------------

                        ' If user provided a plain Name in parts(1), keep it unless overridden by name=
                        If Not firstLooksNamed Then
                            tbName = parts(1).Trim()
                        End If

                        ' If user provided a plain Caption in parts(2) (hybrid), keep it unless overridden by caption=
                        If parts.Length >= 3 AndAlso Not parts(2).Contains("="c) Then
                            labelText = parts(2).Trim()
                        End If

                        For i As Integer = 1 To parts.Length - 1
                            Dim token = parts(i).Trim()
                            If Not token.Contains("="c) Then Continue For

                            Dim kv = token.Split({"="c}, 2)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key

                                Case "name"
                                    tbName = val

                                Case "caption"
                                    labelText = val

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

                    If String.IsNullOrWhiteSpace(tbName) Then Continue For

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

                    ' NEW: register so triggers/actions can reference this control by name
                    RegisterAnyControl(tbName, tb)

                    ' NEW: publish changes for trigger engine (user typing OR programmatic changes)
                    AddHandler tb.TextChanged,
                        Sub()
                            PublishTextBox(tb.Name, tb.Text)
                        End Sub

                    GroupBoxCustom.Controls.Add(tb)

                    ' NEW: publish initial (empty or preset) value once
                    PublishTextBox(tb.Name, tb.Text)

                    ' Only advance autoY if coordinates not explicitly given
                    If Not hasExplicitCoords Then
                        autoY += Math.Max(lbl.Height, tb.Height) + 5
                    End If
                    Continue For



                Case "BUTTON"
                    ' --------------------------------------------------
                    ' Detect ALL-NAMED mode (caption= / action= present)
                    ' --------------------------------------------------
                    Dim named As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                    For i As Integer = 1 To parts.Length - 1
                        Dim p = parts(i).Trim()
                        Dim eq As Integer = p.IndexOf("="c)
                        If eq > 0 Then
                            Dim k As String = p.Substring(0, eq).Trim()
                            Dim v As String = p.Substring(eq + 1).Trim()
                            If k <> "" Then named(k) = v
                        End If
                    Next

                    ' NEW: accept common aliases so config can use your doc terms
                    Dim hasCaption As Boolean = named.ContainsKey("caption")
                    Dim hasAction As Boolean = named.ContainsKey("action")
                    Dim hasDevice As Boolean = named.ContainsKey("device") OrElse named.ContainsKey("devicename")
                    Dim hasName As Boolean = named.ContainsKey("name") OrElse named.ContainsKey("controlname")

                    Dim isAllNamed As Boolean =
                        hasCaption OrElse hasAction OrElse hasDevice OrElse hasName   ' NEW

                    ' ==================================================
                    ' ALL-NAMED MODE
                    ' ==================================================
                    If isAllNamed Then

                        ' NEW: button internal id used by TRIGGER fire:...
                        Dim btnName As String = ""
                        If named.ContainsKey("name") Then btnName = named("name")
                        If btnName = "" AndAlso named.ContainsKey("controlname") Then btnName = named("controlname")

                        Dim caption As String = If(named.ContainsKey("caption"), named("caption"), "")

                        Dim action As String = If(named.ContainsKey("action"), named("action"), "").ToUpperInvariant()

                        ' NEW: accept device/deviceName
                        Dim deviceName As String = ""
                        If named.ContainsKey("device") Then deviceName = named("device")
                        If deviceName = "" AndAlso named.ContainsKey("devicename") Then deviceName = named("devicename")

                        ' NEW: accept command/commandprefix
                        Dim commandOrPrefix As String = ""
                        If named.ContainsKey("command") Then commandOrPrefix = named("command")
                        If commandOrPrefix = "" AndAlso named.ContainsKey("commandprefix") Then commandOrPrefix = named("commandprefix")

                        ' NEW: accept sendval/valuectl/valuecontrol
                        Dim valueControlName As String = ""
                        If named.ContainsKey("sendval") Then valueControlName = named("sendval")
                        If valueControlName = "" AndAlso named.ContainsKey("valuectl") Then valueControlName = named("valuectl")
                        If valueControlName = "" AndAlso named.ContainsKey("valuecontrol") Then valueControlName = named("valuecontrol")

                        ' NEW: accept result/resultctl/resulttarget
                        Dim resultControlName As String = ""
                        If named.ContainsKey("result") Then resultControlName = named("result")
                        If resultControlName = "" AndAlso named.ContainsKey("resultctl") Then resultControlName = named("resultctl")
                        If resultControlName = "" AndAlso named.ContainsKey("resulttarget") Then resultControlName = named("resulttarget")

                        Dim x As Integer = 10
                        Dim y As Integer = autoY
                        Dim w As Integer = 160
                        Dim h As Integer = 0

                        If named.ContainsKey("x") Then ParseIntField(named("x"), x)
                        If named.ContainsKey("y") Then ParseIntField(named("y"), y)
                        If named.ContainsKey("w") Then ParseIntField(named("w"), w)
                        If named.ContainsKey("h") Then ParseIntField(named("h"), h)

                        Dim b As New Button()

                        ' NEW: set internal name so TRIGGER fire:BtnHello can find it
                        If Not String.IsNullOrWhiteSpace(btnName) Then
                            b.Name = btnName
                        End If

                        b.Text = caption
                        b.Tag = $"{action}|{deviceName}|{commandOrPrefix}|{valueControlName}|{resultControlName}"
                        b.Location = New Point(x, y)

                        If h > 0 Then
                            b.Size = New Size(w, h)
                        Else
                            b.Width = w
                        End If

                        AddHandler b.Click, AddressOf CustomButton_Click
                        GroupBoxCustom.Controls.Add(b)

                        ' NEW: register button for trigger engine fire:
                        If Not String.IsNullOrWhiteSpace(b.Name) Then
                            RegisterButton(b.Name, b)
                            RegisterAnyControl(b.Name, b)
                        End If

                        ' Auto-flow only if Y not explicitly supplied
                        If Not named.ContainsKey("y") Then
                            autoY += b.Height + 5
                        End If

                        Continue For
                    End If

                    ' ==================================================
                    ' EXISTING POSITIONAL / HYBRID MODE (UNCHANGED)
                    ' ==================================================
                    If parts.Length < 5 Then Continue For

                    Dim captionPos = parts(1).Trim()
                    Dim actionPos = parts(2).Trim().ToUpperInvariant()
                    Dim deviceNamePos = parts(3).Trim()
                    Dim commandOrPrefixPos = parts(4).Trim()
                    Dim valueControlNamePos As String = If(parts.Length > 5, parts(5).Trim(), "")
                    Dim resultControlNamePos As String = If(parts.Length > 6, parts(6).Trim(), "")

                    ' --- Position (X,Y) ---
                    Dim xPos As Integer = 10
                    Dim yPos As Integer = autoY

                    If parts.Length >= 9 Then
                        ParseIntField(parts(7), xPos)
                        ParseIntField(parts(8), yPos)
                    End If

                    Dim bPos As New Button()
                    bPos.Text = captionPos
                    bPos.Tag = $"{actionPos}|{deviceNamePos}|{commandOrPrefixPos}|{valueControlNamePos}|{resultControlNamePos}"

                    ' --- Size (Width,Height) ---
                    Dim wPos As Integer = 160
                    Dim hPos As Integer = bPos.Height

                    If parts.Length >= 11 Then
                        ParseIntField(parts(9), wPos)
                        ParseIntField(parts(10), hPos)
                    End If

                    bPos.Location = New Point(xPos, yPos)
                    bPos.Size = New Size(wPos, hPos)

                    AddHandler bPos.Click, AddressOf CustomButton_Click
                    GroupBoxCustom.Controls.Add(bPos)

                    ' NEW: positional buttons don't have a name -> triggers can't fire them unless you invent one.
                    ' If you want, you can optionally auto-name them here, but leaving unchanged per your comment.

                    If parts.Length < 9 Then
                        autoY += bPos.Height + 5
                    End If



                Case "CHECKBOX"

                    If parts.Length < 2 Then Continue For

                    Dim resultName As String = ""
                    Dim caption As String = ""
                    Dim funcKey As String = ""
                    Dim param As String = ""

                    Dim x As Integer = 0
                    Dim y As Integer = 0

                    Dim firstLooksNamed As Boolean =
                        parts.Length >= 2 AndAlso parts(1).IndexOf("="c) >= 0

                    ' -----------------------------
                    ' Positional form (legacy)
                    ' CHECKBOX;ResultControlName;Caption;FunctionKey;Param;X;Y
                    ' -----------------------------
                    If Not firstLooksNamed AndAlso parts.Length >= 7 AndAlso Not parts(2).Contains("="c) Then

                        resultName = parts(1).Trim()
                        caption = parts(2).Trim()
                        funcKey = parts(3).Trim()
                        param = parts(4).Trim()

                        ParseIntField(parts(5), x)
                        ParseIntField(parts(6), y)

                    Else
                        ' -----------------------------
                        ' Named / hybrid token form
                        ' -----------------------------
                        If Not firstLooksNamed Then
                            resultName = parts(1).Trim()
                        End If

                        If parts.Length >= 3 AndAlso Not parts(2).Contains("="c) Then
                            caption = parts(2).Trim()
                        End If

                        If parts.Length >= 4 AndAlso Not parts(3).Contains("="c) Then
                            funcKey = parts(3).Trim()
                        End If

                        If parts.Length >= 5 AndAlso Not parts(4).Contains("="c) Then
                            param = parts(4).Trim()
                        End If

                        For i As Integer = 1 To parts.Length - 1
                            Dim t As String = parts(i).Trim()
                            If Not t.Contains("="c) Then Continue For

                            Dim kv = t.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "result", "target", "resulttarget"
                                    resultName = val

                                Case "caption", "text", "label"
                                    caption = val

                                Case "func", "function", "funckey"
                                    funcKey = val

                                Case "param"
                                    param = val

                                Case "x"
                                    ParseIntField(val, x)

                                Case "y"
                                    ParseIntField(val, y)
                            End Select
                        Next
                    End If

                    If String.IsNullOrWhiteSpace(funcKey) Then Continue For

                    Dim cb As New CheckBox()
                    cb.Text = caption
                    cb.AutoSize = True
                    cb.Location = New Point(x, y)

                    cb.Tag = resultName & "|" & funcKey.ToUpperInvariant() & "|" & param
                    cb.Name = "Chk_" & resultName & "_" & funcKey

                    If funcKey.Equals("FuncAuto", StringComparison.OrdinalIgnoreCase) Then
                        AddHandler cb.CheckedChanged, AddressOf FuncAutoCheckbox_CheckedChanged

                    ElseIf funcKey.Equals("FuncTrigEnable", StringComparison.OrdinalIgnoreCase) Then
                        AddHandler cb.CheckedChanged, AddressOf FuncTrigEnableCheckbox_CheckedChanged

                    End If

                    GroupBoxCustom.Controls.Add(cb)

                    Continue For


                Case "LABEL"
                    ' Supports:
                    '   LABEL;Caption;X;Y;FontSize                 (positional)
                    '   LABEL;caption=..;x=..;y=..;f=..           (all-named)
                    '   LABEL;Caption;x=..;y=..;f=..              (hybrid)
                    '   LABEL;name=Lbl1;caption=..;x=..;y=..;f=..

                    If parts.Length < 2 Then Continue For

                    Dim name As String = ""
                    Dim caption As String = ""
                    Dim x As Integer = 0
                    Dim y As Integer = 0
                    Dim fontSize As Single = 10.0F

                    Dim firstLooksNamed As Boolean =
                        parts.Length >= 2 AndAlso parts(1).IndexOf("="c) >= 0

                    ' -------------------------------------------------
                    ' Positional legacy form
                    ' LABEL;Caption;X;Y;FontSize
                    ' -------------------------------------------------
                    If Not firstLooksNamed AndAlso parts.Length >= 2 AndAlso Not parts(2).Contains("="c) Then

                        caption = parts(1).Trim()

                        If parts.Length >= 4 Then
                            ParseIntField(parts(2), x)
                            ParseIntField(parts(3), y)
                        End If

                        If parts.Length >= 5 Then
                            Dim fs As Single
                            If Single.TryParse(parts(4).Trim(),
                                               Globalization.NumberStyles.Float,
                                               Globalization.CultureInfo.InvariantCulture,
                                               fs) Then
                                fontSize = fs
                            End If
                        End If

                    Else
                        ' -------------------------------------------------
                        ' Named / hybrid form
                        ' -------------------------------------------------

                        ' Hybrid: allow positional caption
                        If Not firstLooksNamed Then
                            caption = parts(1).Trim()
                        End If

                        For i As Integer = 1 To parts.Length - 1
                            Dim t As String = parts(i).Trim()
                            If Not t.Contains("="c) Then Continue For

                            Dim kv = t.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "name"
                                    name = val

                                Case "caption", "text"
                                    caption = val

                                Case "x"
                                    ParseIntField(val, x)

                                Case "y"
                                    ParseIntField(val, y)

                                Case "f", "fontsize"
                                    Dim fs As Single
                                    If Single.TryParse(val,
                                                       Globalization.NumberStyles.Float,
                                                       Globalization.CultureInfo.InvariantCulture,
                                                       fs) Then
                                        fontSize = fs
                                    End If
                            End Select
                        Next
                    End If

                    If caption = "" Then Continue For

                    ' -------------------------------------------------
                    ' Create label
                    ' -------------------------------------------------
                    Dim lbl As New Label()
                    lbl.Text = caption
                    lbl.AutoSize = True
                    lbl.Location = New Point(x, y)
                    lbl.Font = New Font(lbl.Font.FontFamily, fontSize, FontStyle.Regular)

                    If name <> "" Then
                        lbl.Name = name
                        RegisterDynamicControl(name, lbl)
                    End If

                    GroupBoxCustom.Controls.Add(lbl)

                    Continue For


                Case "DROPDOWN"
                    ' Supports BOTH:
                    ' 1) Positional:
                    '    DROPDOWN;ControlName;Caption;DeviceName;CommandPrefix;X;Y;Width;Item1,Item2,Item3...;captionpos=1/2
                    '
                    ' 2) All-named:
                    '    DROPDOWN;name=Ctrl1;caption=Range;device=dev1;cmd=:SENS:VOLT:DC:RANG; x=20;y=60;w=140;items=0.1,1,10;captionpos=1
                    '
                    ' captionpos=1 (default) -> Caption label to the left, first combo item is blank.
                    ' captionpos=2           -> No label; caption used as first combo item (placeholder, no command).

                    Dim ctrlName As String = ""
                    Dim caption As String = ""
                    Dim deviceName As String = ""
                    Dim commandPrefix As String = ""

                    Dim x As Integer = 0
                    Dim y As Integer = 0
                    Dim w As Integer = 120
                    Dim itemsRaw As String = ""
                    Dim captionPos As Integer = 1

                    ' Decide whether this line is positional or all-named
                    Dim isAllNamed As Boolean = (parts.Length >= 2 AndAlso parts(1).Contains("="c))

                    If Not isAllNamed Then
                        ' ---------- positional ----------
                        If parts.Length < 9 Then Continue For

                        ctrlName = parts(1).Trim()
                        caption = parts(2).Trim()
                        deviceName = parts(3).Trim()
                        commandPrefix = parts(4).Trim()

                        If parts.Length > 5 Then ParseIntField(parts(5), x)
                        If parts.Length > 6 Then ParseIntField(parts(6), y)
                        If parts.Length > 7 Then ParseIntField(parts(7), w)

                        itemsRaw = parts(8).Trim()

                        If parts.Length > 9 Then
                            For i As Integer = 9 To parts.Length - 1
                                Dim p As String = parts(i).Trim()
                                If Not p.Contains("="c) Then Continue For
                                Dim kv = p.Split("="c)
                                If kv.Length <> 2 Then Continue For

                                Dim key = kv(0).Trim().ToLowerInvariant()
                                Dim val = kv(1).Trim()

                                If key = "captionpos" Then
                                    Integer.TryParse(val, captionPos)
                                    If captionPos <> 1 AndAlso captionPos <> 2 Then captionPos = 1
                                End If
                            Next
                        End If

                    Else
                        ' ---------- all-named ----------
                        ' DROPDOWN;name=...;caption=...;device=...;cmd=...;x=...;y=...;w=...;items=...;captionpos=...
                        For i As Integer = 1 To parts.Length - 1
                            Dim p As String = parts(i).Trim()
                            If Not p.Contains("="c) Then Continue For

                            Dim kv = p.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "name", "ctrl", "control", "controlname"
                                    ctrlName = val

                                Case "caption"
                                    caption = val

                                Case "device", "dev", "devicename"
                                    deviceName = val

                                Case "cmd", "command", "commandprefix"
                                    commandPrefix = val

                                Case "x"
                                    ParseIntField(val, x)

                                Case "y"
                                    ParseIntField(val, y)

                                Case "w", "width"
                                    ParseIntField(val, w)

                                Case "items"
                                    itemsRaw = val

                                Case "captionpos"
                                    Integer.TryParse(val, captionPos)
                                    If captionPos <> 1 AndAlso captionPos <> 2 Then captionPos = 1
                            End Select
                        Next

                        ' Minimal required fields
                        If String.IsNullOrWhiteSpace(ctrlName) OrElse
                           String.IsNullOrWhiteSpace(deviceName) OrElse
                           String.IsNullOrWhiteSpace(commandPrefix) OrElse
                           String.IsNullOrWhiteSpace(itemsRaw) Then
                            Continue For
                        End If
                    End If

                    Dim items = itemsRaw.Split(","c)

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
                        cb.Items.Add("")
                    Else
                        cb.Items.Add(caption)
                    End If

                    ' Add items
                    For Each it In items
                        Dim s As String = it.Trim()
                        If s <> "" Then cb.Items.Add(s)
                    Next

                    cb.SelectedIndex = 0

                    ' Tag holds Device + CommandPrefix
                    cb.Tag = deviceName & "|" & commandPrefix

                    AddHandler cb.SelectedIndexChanged, AddressOf Dropdown_SelectedIndexChanged

                    GroupBoxCustom.Controls.Add(cb)


                Case "RADIOGROUP"

                    If parts.Length < 2 Then Continue For

                    Dim groupName As String = ""
                    Dim caption As String = ""

                    Dim x As Integer = 0
                    Dim y As Integer = 0
                    Dim w As Integer = 100
                    Dim h As Integer = 60

                    ' collect named tokens (from anywhere after the keyword)
                    Dim tok As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    For i As Integer = 1 To parts.Length - 1
                        Dim t = parts(i).Trim()
                        Dim eq = t.IndexOf("="c)
                        If eq > 0 AndAlso eq < t.Length - 1 Then
                            Dim k = t.Substring(0, eq).Trim()
                            Dim v = t.Substring(eq + 1).Trim()
                            If k <> "" Then tok(k) = v
                        End If
                    Next

                    ' group name
                    If parts(1).Contains("="c) Then
                        If tok.ContainsKey("group") Then groupName = tok("group")
                        If groupName = "" AndAlso tok.ContainsKey("name") Then groupName = tok("name")
                    Else
                        groupName = parts(1).Trim()
                    End If
                    If groupName = "" Then Continue For

                    ' caption
                    If parts.Length >= 3 AndAlso Not parts(2).Contains("="c) Then
                        caption = parts(2).Trim()
                    ElseIf tok.ContainsKey("caption") Then
                        caption = tok("caption")
                    Else
                        caption = groupName
                    End If

                    ' positional coords still supported if present
                    If parts.Length >= 7 AndAlso Not parts(3).Contains("="c) Then
                        ParseIntField(parts(3), x)
                        ParseIntField(parts(4), y)
                        ParseIntField(parts(5), w)
                        ParseIntField(parts(6), h)
                    End If

                    ' named overrides
                    If tok.ContainsKey("x") Then ParseIntField(tok("x"), x)
                    If tok.ContainsKey("y") Then ParseIntField(tok("y"), y)
                    If tok.ContainsKey("w") Then ParseIntField(tok("w"), w)
                    If tok.ContainsKey("h") Then ParseIntField(tok("h"), h)

                    Dim gb As New GroupBox()
                    gb.Name = "RG_" & groupName
                    gb.Text = caption
                    gb.Location = New Point(x, y)
                    gb.Size = New Size(w, h)

                    GroupBoxCustom.Controls.Add(gb)

                    Continue For


                Case "RADIO"
                    If parts.Length < 2 Then Continue For

                    Dim groupName As String = ""
                    Dim caption As String = ""
                    Dim deviceName As String = ""
                    Dim command As String = ""

                    Dim relX As Integer = 0
                    Dim relY As Integer = 0

                    ' collect named tokens (from anywhere after the keyword)
                    Dim tok As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    For i As Integer = 1 To parts.Length - 1
                        Dim t = parts(i).Trim()
                        Dim eq = t.IndexOf("="c)
                        If eq > 0 AndAlso eq < t.Length - 1 Then
                            Dim k = t.Substring(0, eq).Trim()
                            Dim v = t.Substring(eq + 1).Trim()
                            If k <> "" Then tok(k) = v
                        End If
                    Next

                    ' group name
                    If parts(1).Contains("="c) Then
                        If tok.ContainsKey("group") Then groupName = tok("group")
                    Else
                        groupName = parts(1).Trim()
                    End If
                    If groupName = "" Then Continue For

                    ' caption
                    If parts.Length >= 3 AndAlso Not parts(2).Contains("="c) Then
                        caption = parts(2).Trim()
                    ElseIf tok.ContainsKey("caption") Then
                        caption = tok("caption")
                    End If

                    ' device + command
                    If parts.Length >= 5 AndAlso Not parts(3).Contains("="c) Then
                        deviceName = parts(3).Trim()
                        command = parts(4).Trim()
                    End If
                    If deviceName = "" Then
                        If tok.ContainsKey("device") Then deviceName = tok("device")
                        If deviceName = "" AndAlso tok.ContainsKey("dev") Then deviceName = tok("dev")
                    End If
                    If command = "" Then
                        If tok.ContainsKey("command") Then command = tok("command")
                        If command = "" AndAlso tok.ContainsKey("cmd") Then command = tok("cmd")
                    End If

                    If caption = "" OrElse deviceName = "" OrElse command = "" Then Continue For

                    ' positional X/Y
                    If parts.Length >= 7 AndAlso Not parts(5).Contains("="c) Then
                        ParseIntField(parts(5), relX)
                        ParseIntField(parts(6), relY)
                    End If

                    ' named X/Y override
                    If tok.ContainsKey("x") Then ParseIntField(tok("x"), relX)
                    If tok.ContainsKey("y") Then ParseIntField(tok("y"), relY)

                    ' scale token (named, or existing tail token form)
                    Dim scale As Double = 1.0
                    Dim scaleIsAuto As Boolean = False
                    Dim rangeQueryForAuto As String = ""

                    Dim scaleVal As String = ""
                    If tok.ContainsKey("scale") Then
                        scaleVal = tok("scale")
                    Else
                        ' also support older: ...;scale=...
                        For i As Integer = 7 To parts.Length - 1
                            Dim t = parts(i).Trim()
                            If t.StartsWith("scale=", StringComparison.OrdinalIgnoreCase) Then
                                scaleVal = t.Substring(6).Trim()
                                Exit For
                            End If
                        Next
                    End If

                    If scaleVal <> "" Then
                        Dim lower = scaleVal.ToLowerInvariant()
                        If lower.StartsWith("auto") Then
                            scaleIsAuto = True
                            Dim pipeIdx As Integer = scaleVal.IndexOf("|"c)
                            If pipeIdx >= 0 AndAlso pipeIdx < scaleVal.Length - 1 Then
                                rangeQueryForAuto = scaleVal.Substring(pipeIdx + 1).Trim()
                            End If
                        Else
                            Double.TryParse(scaleVal, Globalization.NumberStyles.Float,
                                            Globalization.CultureInfo.InvariantCulture, scale)
                        End If
                    End If

                    ' Find parent group box
                    Dim parentName As String = "RG_" & groupName
                    Dim found() As Control = GroupBoxCustom.Controls.Find(parentName, True)
                    If found Is Nothing OrElse found.Length = 0 Then Continue For

                    Dim gb As GroupBox = TryCast(found(0), GroupBox)
                    If gb Is Nothing Then Continue For

                    Dim rb As New RadioButton()
                    rb.Text = caption
                    rb.AutoSize = True
                    rb.Location = New Point(relX, relY)

                    If scaleIsAuto Then
                        rb.Tag = deviceName & "|" & command & "|AUTO|" & rangeQueryForAuto
                    Else
                        rb.Tag = deviceName & "|" & command & "|" &
                                 scale.ToString(Globalization.CultureInfo.InvariantCulture)
                    End If

                    AddHandler rb.CheckedChanged, AddressOf Radio_CheckedChanged
                    gb.Controls.Add(rb)
                    Continue For


                Case "SLIDER"
                    If parts.Length < 2 Then Continue For

                    ' ---- gather tokens (any key=value after SLIDER) ----
                    Dim tok As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    For i As Integer = 1 To parts.Length - 1
                        Dim t = parts(i).Trim()
                        Dim eq = t.IndexOf("="c)
                        If eq > 0 AndAlso eq < t.Length - 1 Then
                            Dim k = t.Substring(0, eq).Trim()
                            Dim v = t.Substring(eq + 1).Trim()
                            If k <> "" Then tok(k) = v
                        End If
                    Next

                    Dim controlName As String = ""
                    Dim caption As String = ""
                    Dim deviceName As String = ""
                    Dim commandPrefix As String = ""

                    ' ---- defaults ----
                    Dim x As Integer = 20
                    Dim y As Integer = 20
                    Dim width As Integer = 200
                    Dim minVal As Integer = 0
                    Dim maxVal As Integer = 100
                    Dim stepVal As Integer = 1
                    Dim scale As Double = 1.0
                    Dim hintText As String = ""

                    Dim usedPositional As Boolean = False
                    Dim isAllNamed As Boolean = (parts.Length >= 2 AndAlso parts(1).Contains("="c))

                    ' -------------------------------
                    ' Positional form (original)
                    ' -------------------------------
                    If Not isAllNamed AndAlso parts.Length >= 5 Then
                        controlName = parts(1).Trim()
                        caption = parts(2).Trim()
                        deviceName = parts(3).Trim()
                        commandPrefix = parts(4).Trim()

                        ' Try positional fields
                        If parts.Length >= 12 AndAlso Not parts(5).Contains("="c) Then
                            ParseIntField(parts(5), x)
                            ParseIntField(parts(6), y)
                            ParseIntField(parts(7), width)
                            ParseIntField(parts(8), minVal)
                            ParseIntField(parts(9), maxVal)
                            ParseIntField(parts(10), stepVal)

                            Double.TryParse(parts(11).Trim(),
                                            Globalization.NumberStyles.Float,
                                            Globalization.CultureInfo.InvariantCulture,
                                            scale)

                            usedPositional = True
                        End If
                    End If

                    ' -------------------------------
                    ' All-named or Hybrid-named
                    ' -------------------------------
                    If isAllNamed OrElse Not usedPositional Then

                        If controlName = "" Then
                            If tok.ContainsKey("name") Then controlName = tok("name")
                            If controlName = "" AndAlso tok.ContainsKey("control") Then controlName = tok("control")
                        End If

                        If caption = "" Then
                            If tok.ContainsKey("caption") Then caption = tok("caption")
                        End If

                        If deviceName = "" Then
                            If tok.ContainsKey("device") Then deviceName = tok("device")
                            If deviceName = "" AndAlso tok.ContainsKey("dev") Then deviceName = tok("dev")
                        End If

                        If commandPrefix = "" Then
                            If tok.ContainsKey("command") Then commandPrefix = tok("command")
                            If commandPrefix = "" AndAlso tok.ContainsKey("cmd") Then commandPrefix = tok("cmd")
                        End If

                        ' If hybrid named (Name;Caption;Device;Command; x=.. etc) but caption/device/command were positional already,
                        ' we still allow overriding geometry/params via tokens below.

                        If tok.ContainsKey("x") Then ParseIntField(tok("x"), x)
                        If tok.ContainsKey("y") Then ParseIntField(tok("y"), y)
                        If tok.ContainsKey("w") Then ParseIntField(tok("w"), width)

                        If tok.ContainsKey("min") Then ParseIntField(tok("min"), minVal)
                        If tok.ContainsKey("max") Then ParseIntField(tok("max"), maxVal)
                        If tok.ContainsKey("step") Then ParseIntField(tok("step"), stepVal)

                        If tok.ContainsKey("scale") Then
                            Double.TryParse(tok("scale"),
                                            Globalization.NumberStyles.Float,
                                            Globalization.CultureInfo.InvariantCulture,
                                            scale)
                        End If

                        If tok.ContainsKey("hint") Then hintText = tok("hint")
                    End If

                    ' basic required fields
                    If controlName = "" OrElse caption = "" OrElse deviceName = "" OrElse commandPrefix = "" Then Continue For
                    If maxVal < minVal Then
                        Dim tmp = minVal : minVal = maxVal : maxVal = tmp
                    End If
                    If stepVal <= 0 Then stepVal = 1

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
                    lblValue.Text = ""
                    lblValue.Location = New Point(width + 20, 30)
                    gb.Controls.Add(lblValue)
                    gb.Height = 72

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
                             stepVal.ToString(Globalization.CultureInfo.InvariantCulture)

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
                    Continue For


                Case "HR"
                    Dim x As Integer = 0
                    Dim y As Integer = 0
                    Dim w As Integer = 200   ' default width

                    ' ---------- Positional format ----------
                    If parts.Length >= 4 AndAlso Not parts(1).Contains("="c) Then
                        ParseIntField(parts(1), x)
                        ParseIntField(parts(2), y)
                        ParseIntField(parts(3), w)

                    Else
                        ' ---------- Named/token format ----------
                        For i As Integer = 1 To parts.Length - 1
                            Dim p = parts(i).Trim()
                            If Not p.Contains("="c) Then Continue For

                            Dim kv = p.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "x" : ParseIntField(val, x)
                                Case "y" : ParseIntField(val, y)
                                Case "w", "width" : ParseIntField(val, w)
                            End Select
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
                    Continue For


                Case "VR"
                    Dim x As Integer = 10
                    Dim y As Integer = 10
                    Dim h As Integer = 100

                    ' ---------- Positional format ----------
                    If parts.Length >= 4 AndAlso Not parts(1).Contains("="c) Then
                        ParseIntField(parts(1), x)
                        ParseIntField(parts(2), y)
                        ParseIntField(parts(3), h)

                    Else
                        ' ---------- Named/token format ----------
                        For i As Integer = 1 To parts.Length - 1
                            Dim tok = parts(i).Trim()
                            If Not tok.Contains("="c) Then Continue For

                            Dim kv = tok.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "x" : ParseIntField(val, x)
                                Case "y" : ParseIntField(val, y)
                                Case "h", "height" : ParseIntField(val, h)
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
                    Continue For



                Case "SPINNER"
                    If parts.Length < 2 Then Continue For

                    ' -----------------------
                    ' Defaults
                    ' -----------------------
                    Dim controlName As String = ""
                    Dim caption As String = ""
                    Dim deviceName As String = ""
                    Dim commandPrefix As String = ""

                    Dim x As Integer = 20
                    Dim y As Integer = 20
                    Dim width As Integer = 80
                    Dim minVal As Integer = 0
                    Dim maxVal As Integer = 100
                    Dim stepVal As Integer = 1
                    Dim scale As Double = 1.0

                    ' -----------------------
                    ' Detect if there are any named tokens early
                    ' -----------------------
                    Dim hasNamed As Boolean = False
                    For i As Integer = 1 To parts.Length - 1
                        If parts(i).Contains("="c) Then
                            hasNamed = True
                            Exit For
                        End If
                    Next

                    ' -----------------------
                    ' Positional header (classic)
                    ' -----------------------
                    If Not hasNamed Then
                        If parts.Length < 5 Then Continue For

                        controlName = parts(1).Trim()
                        caption = parts(2).Trim()
                        deviceName = parts(3).Trim()
                        commandPrefix = parts(4).Trim()

                        ' Positional coords/limits
                        If parts.Length >= 12 AndAlso Not parts(5).Contains("="c) Then
                            ParseIntField(parts(5), x)
                            ParseIntField(parts(6), y)
                            ParseIntField(parts(7), width)
                            ParseIntField(parts(8), minVal)
                            ParseIntField(parts(9), maxVal)
                            ParseIntField(parts(10), stepVal)

                            Double.TryParse(parts(11).Trim(),
                                            Globalization.NumberStyles.Float,
                                            Globalization.CultureInfo.InvariantCulture,
                                            scale)
                        Else
                            ' If they used classic header but then tokens for the rest, parse tokens from 5+
                            For i As Integer = 5 To parts.Length - 1
                                Dim p = parts(i).Trim()
                                If Not p.Contains("="c) Then Continue For

                                Dim kv = p.Split("="c)
                                If kv.Length <> 2 Then Continue For
                                Dim key = kv(0).Trim().ToLowerInvariant()
                                Dim val = kv(1).Trim()

                                Select Case key
                                    Case "x" : ParseIntField(val, x)
                                    Case "y" : ParseIntField(val, y)
                                    Case "w", "width" : ParseIntField(val, width)
                                    Case "min" : ParseIntField(val, minVal)
                                    Case "max" : ParseIntField(val, maxVal)
                                    Case "step" : ParseIntField(val, stepVal)
                                    Case "scale"
                                        Double.TryParse(val,
                                                        Globalization.NumberStyles.Float,
                                                        Globalization.CultureInfo.InvariantCulture,
                                                        scale)
                                End Select
                            Next
                        End If

                    Else
                        ' -----------------------
                        ' Fully token/named form (and hybrid)
                        ' -----------------------
                        ' We allow either:
                        '   SPINNER;Name;Caption;Device;Cmd; x=.. etc
                        ' or:
                        '   SPINNER;name=..;caption=..;device=..;cmd=..; x=.. etc

                        ' First, capture positional header if they provided it
                        If parts.Length >= 5 AndAlso
                           Not parts(1).Contains("="c) AndAlso
                           Not parts(2).Contains("="c) AndAlso
                           Not parts(3).Contains("="c) AndAlso
                           Not parts(4).Contains("="c) Then

                            controlName = parts(1).Trim()
                            caption = parts(2).Trim()
                            deviceName = parts(3).Trim()
                            commandPrefix = parts(4).Trim()
                        End If

                        ' Then parse all tokens (including possible name/caption/device/cmd overrides)
                        For i As Integer = 1 To parts.Length - 1
                            Dim p = parts(i).Trim()
                            If Not p.Contains("="c) Then Continue For

                            Dim kv = p.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "name" : controlName = val
                                Case "caption" : caption = val
                                Case "device" : deviceName = val
                                Case "cmd", "command" : commandPrefix = val

                                Case "x" : ParseIntField(val, x)
                                Case "y" : ParseIntField(val, y)
                                Case "w", "width" : ParseIntField(val, width)
                                Case "min" : ParseIntField(val, minVal)
                                Case "max" : ParseIntField(val, maxVal)
                                Case "step" : ParseIntField(val, stepVal)
                                Case "scale"
                                    Double.TryParse(val,
                                                    Globalization.NumberStyles.Float,
                                                    Globalization.CultureInfo.InvariantCulture,
                                                    scale)
                            End Select
                        Next

                        If controlName = "" OrElse deviceName = "" OrElse commandPrefix = "" Then
                            Continue For
                        End If
                    End If

                    ' -----------------------
                    ' Label
                    ' -----------------------
                    Dim lbl As New Label()
                    lbl.Text = caption
                    lbl.AutoSize = True
                    lbl.Location = New Point(x, y + 3)
                    GroupBoxCustom.Controls.Add(lbl)

                    ' -----------------------
                    ' Spinner (NumericUpDown)
                    ' -----------------------
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
                    Continue For


                Case "LED"
                    If parts.Length < 3 Then Continue For

                    Dim ledName As String = ""
                    Dim caption As String = ""

                    Dim x As Integer = 10
                    Dim y As Integer = 10
                    Dim size As Integer = 12

                    ' Default colours
                    Dim onColor As Color = Color.LimeGreen
                    Dim offColor As Color = Color.DarkGray
                    Dim badColor As Color = Color.Gold

                    Dim usedPositional As Boolean = False

                    ' -----------------------------
                    ' Positional form?
                    ' LED;Name;Caption;X;Y;Size
                    ' -----------------------------
                    If parts.Length >= 6 AndAlso Not parts(1).Contains("="c) AndAlso Not parts(3).Contains("="c) Then
                        ledName = parts(1).Trim()
                        caption = parts(2).Trim()
                        ParseIntField(parts(3), x)
                        ParseIntField(parts(4), y)
                        ParseIntField(parts(5), size)
                        usedPositional = True
                    End If

                    ' -----------------------------
                    ' Named / token form
                    ' LED;name=..;caption=..;x=..;y=..;s=..;on=..;off=..;bad=..
                    ' -----------------------------
                    If Not usedPositional Then
                        For i As Integer = 1 To parts.Length - 1
                            Dim p = parts(i).Trim()
                            If Not p.Contains("="c) Then Continue For

                            Dim kv = p.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "name" : ledName = val
                                Case "caption" : caption = val
                                Case "x" : ParseIntField(val, x)
                                Case "y" : ParseIntField(val, y)
                                Case "s", "size" : ParseIntField(val, size)

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
                        Next

                        ' fallback if someone still used legacy first two fields but then named coords
                        If ledName = "" AndAlso parts.Length > 1 AndAlso Not parts(1).Contains("="c) Then ledName = parts(1).Trim()
                        If caption = "" AndAlso parts.Length > 2 AndAlso Not parts(2).Contains("="c) Then caption = parts(2).Trim()
                    End If

                    If ledName = "" Then Continue For

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

                    led.Tag = $"LED|{onColor.ToArgb()}|{offColor.ToArgb()}|{badColor.ToArgb()}"
                    RegisterAnyControl(ledName, led)

                    ' Tag: marker + colours
                    'led.Tag = $"LED|{onColor.ToArgb}|{offColor.ToArgb}|{badColor.ToArgb}"

                    GroupBoxCustom.Controls.Add(led)
                    Continue For


                Case "BIGTEXT"
                    If parts.Length >= 2 Then

                        Dim controlName As String = ""
                        Dim caption As String = ""

                        ' Defaults
                        Dim x As Integer = 20
                        Dim y As Integer = 20
                        Dim fontSize As Single = 28.0F
                        Dim w As Integer = 200
                        Dim h As Integer = 60
                        Dim borderOn As Boolean = True
                        Dim unitsOn As Boolean = True
                        Dim fmt As String = ""

                        ' ------------------------------------------------------------
                        ' Parse first fields:
                        ' positional: parts(1)=name, parts(2)=caption
                        ' named:      parts(1)=name=..., parts(2)=caption=...
                        ' ------------------------------------------------------------
                        If parts(1).Contains("="c) Then
                            ' name/caption are given as tokens too
                            For i As Integer = 1 To parts.Length - 1
                                Dim p = parts(i).Trim()
                                If Not p.Contains("="c) Then Continue For

                                Dim kv = p.Split("="c)
                                If kv.Length <> 2 Then Continue For

                                Dim key = kv(0).Trim().ToLowerInvariant()
                                Dim val = kv(1).Trim()

                                Select Case key
                                    Case "name" : controlName = val
                                    Case "caption" : caption = val

                                    Case "x" : Integer.TryParse(val, x)
                                    Case "y" : Integer.TryParse(val, y)
                                    Case "f" : Single.TryParse(val, fontSize)
                                    Case "w" : Integer.TryParse(val, w)
                                    Case "h" : Integer.TryParse(val, h)

                                    Case "border"
                                        Select Case val.ToLowerInvariant()
                                            Case "0", "off", "false", "no", "none"
                                                borderOn = False
                                            Case Else
                                                borderOn = True
                                        End Select

                                    Case "units"
                                        Select Case val.ToLowerInvariant()
                                            Case "0", "off", "false", "no", "none"
                                                unitsOn = False
                                            Case Else
                                                unitsOn = True
                                        End Select

                                    Case "format"
                                        fmt = val
                                End Select
                            Next
                        Else
                            ' original positional fields
                            controlName = parts(1).Trim()
                            caption = If(parts.Length > 2, parts(2).Trim(), "")

                            ' named tokens start at index 3 (same as your existing code)
                            For i As Integer = 3 To parts.Length - 1
                                Dim p = parts(i).Trim()
                                If Not p.Contains("="c) Then Continue For

                                Dim kv = p.Split("="c)
                                If kv.Length <> 2 Then Continue For

                                Dim key = kv(0).Trim().ToLowerInvariant()
                                Dim val = kv(1).Trim()

                                Select Case key
                                    Case "x" : Integer.TryParse(val, x)
                                    Case "y" : Integer.TryParse(val, y)
                                    Case "f" : Single.TryParse(val, fontSize)
                                    Case "w" : Integer.TryParse(val, w)
                                    Case "h" : Integer.TryParse(val, h)

                                    Case "border"
                                        Select Case val.ToLowerInvariant()
                                            Case "0", "off", "false", "no", "none"
                                                borderOn = False
                                            Case Else
                                                borderOn = True
                                        End Select

                                    Case "units"
                                        Select Case val.ToLowerInvariant()
                                            Case "0", "off", "false", "no", "none"
                                                unitsOn = False
                                            Case Else
                                                unitsOn = True
                                        End Select

                                    Case "format"
                                        fmt = val
                                End Select
                            Next
                        End If

                        ' Must have a real name
                        If String.IsNullOrWhiteSpace(controlName) Then Exit Select

                        ' Guard: font size must be valid (>0)
                        If fontSize <= 0.0F Then fontSize = 28.0F

                        ' --- Outer panel ---
                        Dim panel As New Panel()
                        panel.Location = New Point(x, y)
                        panel.Size = New Size(w, h)
                        panel.Name = "Panel_" & controlName
                        panel.BorderStyle = If(borderOn, BorderStyle.FixedSingle, BorderStyle.None)
                        panel.BackColor = GroupBoxCustom.BackColor

                        GroupBoxCustom.Controls.Add(panel)

                        ' --- BIG TEXT label ---
                        Dim lbl As New Label()
                        lbl.Name = controlName            ' <-- this is what your updater looks for
                        lbl.Text = "#####"
                        lbl.AutoSize = False
                        lbl.TextAlign = ContentAlignment.MiddleCenter
                        lbl.Font = New Font(lbl.Font.FontFamily, fontSize, FontStyle.Bold)
                        lbl.Dock = DockStyle.Fill

                        ' Preserve legacy units tag, but allow format too
                        Dim tagParts As New List(Of String)

                        If Not unitsOn Then
                            tagParts.Add("BIGTEXT_UNITS_OFF") ' legacy behaviour
                        End If

                        If Not String.IsNullOrWhiteSpace(fmt) Then
                            tagParts.Add("BIGTEXT_FMT=" & fmt)
                        End If

                        If tagParts.Count > 0 Then
                            lbl.Tag = String.Join("|", tagParts)
                        End If

                        panel.Controls.Add(lbl)
                    End If
                    Continue For


                Case "TEXTAREA"
                    If parts.Length < 3 Then Continue For

                    Dim tbName As String = ""
                    Dim labelText As String = ""

                    Dim x As Integer = 10
                    Dim y As Integer = autoY
                    Dim w As Integer = 300
                    Dim h As Integer = 100
                    Dim hasExplicitCoords As Boolean = False
                    Dim initLines As String() = Nothing

                    ' ------------------------------------------------------------
                    ' If parts(1) is tokenised (name=...), treat as all-named mode
                    ' ------------------------------------------------------------
                    If parts(1).Contains("="c) Then
                        For i2 As Integer = 1 To parts.Length - 1
                            Dim p2 = parts(i2).Trim()
                            If Not p2.Contains("="c) Then Continue For

                            Dim kv = p2.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "name" : tbName = val
                                Case "caption" : labelText = val
                                Case "x" : ParseIntField(val, x) : hasExplicitCoords = True
                                Case "y" : ParseIntField(val, y) : hasExplicitCoords = True
                                Case "w" : ParseIntField(val, w)
                                Case "h" : ParseIntField(val, h)
                                Case "init"
                                    If val <> "" Then initLines = val.Split("|"c)
                            End Select
                        Next

                    Else
                        ' ------------------------------------------------------------
                        ' Positional name/caption
                        ' ------------------------------------------------------------
                        tbName = parts(1).Trim()
                        labelText = parts(2).Trim()

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

                                Dim key = kv(0).Trim().ToLowerInvariant()
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
                    End If

                    If String.IsNullOrWhiteSpace(tbName) Then Continue For

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

                    ' trigger
                    RegisterTextArea(tbName, tb)
                    RegisterAnyControl(tbName, tb)

                    ' apply init contents if provided
                    If initLines IsNot Nothing Then
                        tb.Lines = initLines
                    End If

                    GroupBoxCustom.Controls.Add(tb)

                    If Not hasExplicitCoords Then
                        autoY += lbl.Height + h + 8
                    End If

                    Continue For


                Case "TOGGLE"
                    If parts.Length < 2 Then Continue For

                    Dim name As String = ""
                    Dim caption As String = ""
                    Dim device As String = ""
                    Dim cmdOn As String = ""
                    Dim cmdOff As String = ""

                    Dim x As Integer = 20, y As Integer = 20, w As Integer = 120, h As Integer = 35

                    ' -------- detect named vs positional --------
                    Dim anyNamed As Boolean = False
                    For i As Integer = 1 To parts.Length - 1
                        If parts(i).Contains("="c) Then anyNamed = True : Exit For
                    Next

                    If Not anyNamed Then
                        ' ---- positional ----
                        If parts.Length < 6 Then Continue For

                        name = parts(1).Trim()
                        caption = parts(2).Trim()
                        device = parts(3).Trim()
                        cmdOn = parts(4).Trim()
                        cmdOff = parts(5).Trim()

                        If parts.Length >= 10 Then
                            ParseIntField(parts(6), x)
                            ParseIntField(parts(7), y)
                            ParseIntField(parts(8), w)
                            ParseIntField(parts(9), h)
                        End If

                    Else
                        ' ---- named / hybrid ----
                        ' allow first 5 fields positional if user kept them that way
                        If parts.Length >= 6 AndAlso Not parts(1).Contains("="c) Then
                            name = parts(1).Trim()
                            caption = parts(2).Trim()
                            device = parts(3).Trim()
                            cmdOn = parts(4).Trim()
                            cmdOff = parts(5).Trim()
                        End If

                        For i As Integer = 1 To parts.Length - 1
                            Dim p = parts(i).Trim()
                            If Not p.Contains("="c) Then Continue For

                            Dim kv = p.Split({"="c}, 2)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "name" : name = val
                                Case "caption" : caption = val
                                Case "device" : device = val
                                Case "on" : cmdOn = val
                                Case "off" : cmdOff = val
                                Case "x" : ParseIntField(val, x)
                                Case "y" : ParseIntField(val, y)
                                Case "w" : ParseIntField(val, w)
                                Case "h" : ParseIntField(val, h)
                            End Select
                        Next
                    End If

                    If name = "" OrElse device = "" OrElse cmdOn = "" OrElse cmdOff = "" Then Continue For

                    ' Create toggle button
                    Dim b As New Button()
                    b.Name = name
                    b.Text = If(caption <> "", caption, name)
                    b.Location = New Point(x, y)
                    b.Size = New Size(w, h)

                    ' Tag holds: DEVICE|ONCMD|OFFCMD|STATE
                    b.Tag = device & "|" & cmdOn & "|" & cmdOff & "|0"   ' state 0 = OFF

                    AddHandler b.Click, AddressOf ToggleButton_Click
                    GroupBoxCustom.Controls.Add(b)
                    Continue For


                Case "CHART"
                    If parts.Length < 2 Then Continue For

                    ' ---- defaults ----
                    Dim chartName As String = ""
                    Dim caption As String = ""
                    Dim resultTarget As String = ""

                    Dim x As Integer = 20
                    Dim y As Integer = 20
                    Dim w As Integer = 250
                    Dim h As Integer = 120
                    Dim yMin As Double? = Nothing
                    Dim yMax As Double? = Nothing
                    Dim xStep As Double = 1.0R
                    Dim maxPoints As Integer = 100
                    Dim plotColor As Color = Color.Yellow
                    Dim autoScaleY As Boolean = False
                    Dim lineWidth As Integer = 2

                    ' Inner plot area (%)
                    Dim innerX As Single = 2.0F
                    Dim innerY As Single = 2.0F
                    Dim innerW As Single = 96.0F
                    Dim innerH As Single = 96.0F

                    Dim usedPositional As Boolean = False

                    ' ------------------------------------------------------------
                    ' Positional header fields:
                    ' CHART;ChartName;Caption;ResultTarget;...
                    ' ------------------------------------------------------------
                    If parts.Length >= 4 AndAlso Not parts(1).Contains("="c) AndAlso Not parts(2).Contains("="c) AndAlso Not parts(3).Contains("="c) Then
                        chartName = parts(1).Trim()
                        caption = parts(2).Trim()
                        resultTarget = parts(3).Trim()
                        usedPositional = True
                    End If

                    ' ------------------------------------------------------------
                    ' Named header fields:
                    ' CHART;name=...;caption=...;target=...;...
                    ' ------------------------------------------------------------
                    If Not usedPositional Then
                        For i As Integer = 1 To Math.Min(parts.Length - 1, 6) ' header tokens usually early
                            Dim tok As String = parts(i).Trim()
                            If Not tok.Contains("="c) Then Continue For

                            Dim kv = tok.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "name"
                                    chartName = val
                                Case "caption"
                                    caption = val
                                Case "target", "resulttarget", "result", "source"
                                    resultTarget = val
                            End Select
                        Next
                    End If

                    ' Must have at least name + target to work
                    If chartName = "" Then Continue For
                    If resultTarget = "" Then Continue For

                    ' ------------------------------------------------------------
                    ' Parse parameters:
                    ' Positional body (optional) OR named tokens (always allowed)
                    ' ------------------------------------------------------------

                    If usedPositional Then
                        ' Positional extras start at index 4
                        ' X;Y;W;H;YMin;YMax;XStep;MaxPoints;Color;LineWidth;AutoScale;InnerX;InnerY;InnerW;InnerH
                        If parts.Length > 4 Then ParseIntField(parts(4), x)
                        If parts.Length > 5 Then ParseIntField(parts(5), y)
                        If parts.Length > 6 Then ParseIntField(parts(6), w)
                        If parts.Length > 7 Then ParseIntField(parts(7), h)

                        If parts.Length > 8 Then
                            Dim d As Double
                            If Double.TryParse(parts(8).Trim(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, d) Then yMin = d
                        End If
                        If parts.Length > 9 Then
                            Dim d As Double
                            If Double.TryParse(parts(9).Trim(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, d) Then yMax = d
                        End If
                        If parts.Length > 10 Then
                            Double.TryParse(parts(10).Trim(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, xStep)
                        End If
                        If parts.Length > 11 Then
                            Dim mp As Integer
                            If Integer.TryParse(parts(11).Trim(), mp) AndAlso mp > 0 Then maxPoints = mp
                        End If
                        If parts.Length > 12 Then
                            Dim c As Color = Color.FromName(parts(12).Trim())
                            If c.ToArgb() <> Color.Empty.ToArgb() Then plotColor = c
                        End If
                        If parts.Length > 13 Then
                            Dim lw As Integer
                            If Integer.TryParse(parts(13).Trim(), lw) AndAlso lw >= 1 AndAlso lw <= 10 Then lineWidth = lw
                        End If
                        If parts.Length > 14 Then
                            Dim v = parts(14).Trim().ToLowerInvariant()
                            autoScaleY = (v = "yes" OrElse v = "true" OrElse v = "on" OrElse v = "1")
                        End If
                        If parts.Length > 15 Then Single.TryParse(parts(15).Trim(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, innerX)
                        If parts.Length > 16 Then Single.TryParse(parts(16).Trim(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, innerY)
                        If parts.Length > 17 Then Single.TryParse(parts(17).Trim(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, innerW)
                        If parts.Length > 18 Then Single.TryParse(parts(18).Trim(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, innerH)
                    End If

                    ' Named tokens (allowed even if positional was used)
                    For i As Integer = If(usedPositional, 4, 1) To parts.Length - 1
                        Dim p = parts(i).Trim()
                        If Not p.Contains("="c) Then Continue For

                        Dim kv = p.Split("="c)
                        If kv.Length <> 2 Then Continue For

                        Dim key = kv(0).Trim().ToLowerInvariant()
                        Dim val = kv(1).Trim()

                        Select Case key
                            Case "x" : ParseIntField(val, x)
                            Case "y" : ParseIntField(val, y)
                            Case "w" : ParseIntField(val, w)
                            Case "h" : ParseIntField(val, h)

                            Case "ymin"
                                Dim d As Double
                                If Double.TryParse(val, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, d) Then yMin = d

                            Case "ymax"
                                Dim d As Double
                                If Double.TryParse(val, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, d) Then yMax = d

                            Case "xstep"
                                Double.TryParse(val, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, xStep)

                            Case "maxpoints"
                                Dim mp As Integer
                                If Integer.TryParse(val, mp) AndAlso mp > 0 Then maxPoints = mp

                            Case "color"
                                Dim c As Color = Color.FromName(val)
                                If c.ToArgb() <> Color.Empty.ToArgb() Then plotColor = c

                            Case "autoscale"
                                Dim v = val.ToLowerInvariant()
                                autoScaleY = (v = "yes" OrElse v = "true" OrElse v = "on" OrElse v = "1")

                            Case "linewidth"
                                Dim lw As Integer
                                If Integer.TryParse(val, lw) AndAlso lw >= 1 AndAlso lw <= 10 Then lineWidth = lw

                            Case "innerx"
                                Single.TryParse(val, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, innerX)
                            Case "innery"
                                Single.TryParse(val, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, innerY)
                            Case "innerw"
                                Single.TryParse(val, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, innerW)
                            Case "innerh"
                                Single.TryParse(val, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, innerH)
                        End Select
                    Next

                    ' ---- clamp inner plot values ----
                    innerX = Math.Max(0.0F, Math.Min(innerX, 100.0F))
                    innerY = Math.Max(0.0F, Math.Min(innerY, 100.0F))
                    innerW = Math.Max(1.0F, Math.Min(innerW, 100.0F))
                    innerH = Math.Max(1.0F, Math.Min(innerH, 100.0F))

                    ' ---- caption label (optional) ----
                    Dim lblCap As Label = Nothing
                    Dim chartTop As Integer = y

                    If caption <> "" Then
                        lblCap = New Label()
                        lblCap.Text = caption
                        lblCap.AutoSize = True
                        lblCap.Location = New Point(x, y)
                        lblCap.Name = "Lbl_" & chartName

                        RegisterDynamicControl(lblCap.Name, lblCap) ' so invisibility hides it too
                        GroupBoxCustom.Controls.Add(lblCap)

                        chartTop = y + lblCap.Height + 3
                    End If

                    ' ---- chart ----
                    Dim ch As New DataVisualization.Charting.Chart()
                    ch.Name = chartName

                    RegisterDynamicControl(chartName, ch)   ' register for invisibility

                    ch.Location = New Point(x, chartTop)
                    ch.Size = New Size(w, h)
                    ch.BackColor = Color.Black

                    Dim ca As New DataVisualization.Charting.ChartArea("Default")
                    ch.ChartAreas.Add(ca)

                    ' Fill chart completely
                    ca.Position.Auto = False
                    ca.Position = New DataVisualization.Charting.ElementPosition(0, 0, 100, 100)

                    ' Apply inner plot padding from config
                    ca.InnerPlotPosition.Auto = False
                    ca.InnerPlotPosition = New DataVisualization.Charting.ElementPosition(innerX, innerY, innerW, innerH)

                    ' ---- Y range ----
                    If Not autoScaleY Then
                        ca.AxisY.Minimum = If(yMin.HasValue, yMin.Value, Double.NaN)
                        ca.AxisY.Maximum = If(yMax.HasValue, yMax.Value, Double.NaN)
                    Else
                        ca.AxisY.Minimum = Double.NaN
                        ca.AxisY.Maximum = Double.NaN
                    End If

                    ' ---- grid / axes ----
                    ca.BackColor = Color.Black

                    ca.AxisX.LabelStyle.Enabled = False
                    ca.AxisX.MajorTickMark.Enabled = False
                    ca.AxisX.MinorTickMark.Enabled = False
                    ca.AxisX.MajorGrid.Enabled = True
                    ca.AxisX.MajorGrid.LineColor = Color.FromArgb(60, 60, 60)
                    ca.AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

                    ca.AxisY.MajorGrid.Enabled = True
                    ca.AxisY.MajorGrid.LineColor = Color.FromArgb(80, 80, 80)
                    ca.AxisY.LabelStyle.ForeColor = Color.White
                    ca.AxisY.LabelStyle.Font = New Font("Segoe UI", 7.0F)
                    ca.AxisY.MajorTickMark.Enabled = False

                    ' ---- series ----
                    Dim ser As New DataVisualization.Charting.Series("S1")
                    ser.ChartType = DataVisualization.Charting.SeriesChartType.Line
                    ser.BorderWidth = lineWidth
                    ser.Color = plotColor
                    ser.MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
                    ser.MarkerSize = 3
                    ser.MarkerColor = plotColor
                    ch.Series.Add(ser)

                    GroupBoxCustom.Controls.Add(ch)

                    ' Store chart binding so it can be updated even without a source TextBox
                    Dim yMinStr As String = If(yMin.HasValue, yMin.Value.ToString(Globalization.CultureInfo.InvariantCulture), "")
                    Dim yMaxStr As String = If(yMax.HasValue, yMax.Value.ToString(Globalization.CultureInfo.InvariantCulture), "")
                    Dim autoStr As String = If(autoScaleY, "1", "0")

                    ch.Tag = "CHART|" & resultTarget & "|" & yMinStr & "|" & yMaxStr & "|" & xStep.ToString(Globalization.CultureInfo.InvariantCulture) & "|" & maxPoints.ToString(Globalization.CultureInfo.InvariantCulture) & "|" & autoStr

                    ' ---- bind to textbox ----
                    Dim src = TryCast(GetControlByName(resultTarget), TextBox)
                    If src IsNot Nothing Then
                        UpdateChartFromText(ch, src.Text, yMin, yMax, xStep, maxPoints, autoScaleY)
                        AddHandler src.TextChanged,
                            Sub(senderSrc As Object, eSrc As EventArgs)
                                UpdateChartFromText(ch,
                                                    DirectCast(senderSrc, TextBox).Text,
                                                    yMin, yMax, xStep, maxPoints, autoScaleY)
                            End Sub
                    End If
                    Continue For


                Case "STATSPANEL"
                    ' Supports BOTH:
                    '   STATSPANEL;PanelName;Caption;ResultTarget;x=..;y=..;w=..;h=..;format=G7
                    ' and fully-named header:
                    '   STATSPANEL;name=Stats1;caption=Stats;target=TextBoxResult;x=..;y=..;w=..;h=..;format=G7

                    If parts.Length < 2 Then Continue For

                    ' ---- defaults ----
                    Dim panelName As String = ""
                    Dim caption As String = ""
                    Dim resultTarget As String = ""

                    Dim x As Integer = 20
                    Dim y As Integer = 20
                    Dim w As Integer = 260
                    Dim h As Integer = 160
                    Dim fmt As String = "G6"

                    ' ---- read header fields (positional OR named) ----
                    If parts.Length >= 4 Then
                        ' Try named header first
                        For i As Integer = 1 To 3
                            Dim p = parts(i).Trim()
                            If p.Contains("="c) Then
                                Dim kv = p.Split("="c)
                                If kv.Length = 2 Then
                                    Dim key = kv(0).Trim().ToLowerInvariant()
                                    Dim val = kv(1).Trim()
                                    Select Case key
                                        Case "name" : panelName = val
                                        Case "caption" : caption = val
                                        Case "target", "resulttarget" : resultTarget = val
                                    End Select
                                End If
                            End If
                        Next

                        ' If any are still blank, fall back to positional
                        If panelName = "" Then panelName = parts(1).Trim()
                        If caption = "" Then caption = parts(2).Trim()
                        If resultTarget = "" Then resultTarget = parts(3).Trim()
                    Else
                        ' Not enough for positional header
                        Continue For
                    End If

                    If panelName = "" OrElse resultTarget = "" Then Continue For

                    ' ---- parse tail tokens (x= y= w= h= format=) ----
                    For i As Integer = 4 To parts.Length - 1
                        Dim p = parts(i).Trim()
                        If Not p.Contains("="c) Then Continue For

                        Dim kv = p.Split("="c)
                        If kv.Length <> 2 Then Continue For

                        Dim key = kv(0).Trim().ToLowerInvariant()
                        Dim val = kv(1).Trim()

                        Select Case key
                            Case "x" : ParseIntField(val, x)
                            Case "y" : ParseIntField(val, y)
                            Case "w" : ParseIntField(val, w)
                            Case "h" : ParseIntField(val, h)
                            Case "format"
                                If val <> "" Then fmt = val
                        End Select
                    Next

                    ' ---- build panel ----
                    Dim gb As New GroupBox()
                    gb.Name = panelName

                    RegisterDynamicControl(panelName, gb)       ' Invisibility control

                    gb.Text = caption
                    gb.Location = New Point(x, y)
                    gb.Size = New Size(w, h)
                    gb.Tag = "STATSPANEL|" & resultTarget & "|" & fmt
                    GroupBoxCustom.Controls.Add(gb)

                    ' ---- init state/rows ----
                    If Not StatsState.ContainsKey(panelName) Then
                        StatsState(panelName) = New RunningStatsState()
                    Else
                        StatsState(panelName).Reset()
                    End If

                    If Not StatsRows.ContainsKey(panelName) Then
                        StatsRows(panelName) = New List(Of StatsRow)()
                    Else
                        StatsRows(panelName).Clear()
                    End If

                    ' ---- hook to source textbox ----
                    Dim src = TryCast(GetControlByName(resultTarget), TextBox)
                    If src IsNot Nothing Then
                        AddHandler src.TextChanged,
                            Sub(senderSrc As Object, eSrc As EventArgs)
                                Dim t As String = DirectCast(senderSrc, TextBox).Text
                                Dim d As Double
                                If TryExtractFirstDouble(t, d) Then
                                    UpdateStatsPanel(panelName, d)
                                End If
                            End Sub
                    End If
                    Continue For


                Case "STAT"
                    If parts.Length < 2 Then Continue For

                    Dim panelName As String = ""
                    Dim labelText As String = ""
                    Dim funcKey As String = ""
                    Dim refTok As String = ""
                    Dim fmtOverride As String = ""   ' "F" or "E" prefix override (uses panel fmt digits)

                    ' ---- header: named OR positional ----
                    If parts.Length >= 4 Then
                        ' try named header first (parts(1..3))
                        For i As Integer = 1 To 3
                            Dim p = parts(i).Trim()
                            If p.Contains("="c) Then
                                Dim kv = p.Split("="c)
                                If kv.Length = 2 Then
                                    Dim key = kv(0).Trim().ToLowerInvariant()
                                    Dim val = kv(1).Trim()
                                    Select Case key
                                        Case "panel", "panelname", "name" : panelName = val
                                        Case "label", "caption" : labelText = val
                                        Case "func", "funckey" : funcKey = val
                                    End Select
                                End If
                            End If
                        Next
                    End If

                    ' fall back to positional if needed
                    If panelName = "" AndAlso parts.Length >= 2 Then panelName = parts(1).Trim()
                    If labelText = "" AndAlso parts.Length >= 3 Then labelText = parts(2).Trim()
                    If funcKey = "" AndAlso parts.Length >= 4 Then funcKey = parts(3).Trim()

                    If panelName = "" OrElse funcKey = "" Then Continue For

                    funcKey = funcKey.Trim().ToUpperInvariant()

                    ' ---- parse tokens (including if user put them in "tail" OR in named form) ----
                    ' We parse from index 4 onwards in positional, but also allow users to put ref=/units=
                    ' right after STAT;... as named tokens (safe to just scan the entire remainder).
                    Dim startIdx As Integer = 4
                    If parts.Length < 4 Then startIdx = parts.Length ' (but we already Continued above)

                    For i As Integer = startIdx To parts.Length - 1
                        Dim p = parts(i).Trim()
                        If p = "" Then Continue For

                        If p.StartsWith("ref=", StringComparison.OrdinalIgnoreCase) Then
                            refTok = p.Substring(4).Trim()

                        ElseIf p.StartsWith("units=", StringComparison.OrdinalIgnoreCase) Then
                            Dim u = p.Substring(6).Trim().ToLowerInvariant()
                            If u = "decimal" Then
                                fmtOverride = "F"
                            ElseIf u = "scientific" Then
                                fmtOverride = "E"
                            ElseIf u = "general" OrElse u = "sig" OrElse u = "significant" Then
                                fmtOverride = "G"
                            End If

                        ElseIf p.Contains("="c) Then
                            ' also allow named tokens: panel=, label=, func=, ref=, units=
                            Dim kv = p.Split("="c)
                            If kv.Length = 2 Then
                                Dim key = kv(0).Trim().ToLowerInvariant()
                                Dim val = kv(1).Trim()

                                Select Case key
                                    Case "panel", "panelname", "name"
                                        If panelName = "" Then panelName = val
                                    Case "label", "caption"
                                        If labelText = "" Then labelText = val
                                    Case "func", "funckey"
                                        If funcKey = "" Then funcKey = val.ToUpperInvariant()
                                    Case "ref"
                                        refTok = val
                                    Case "units"
                                        Dim u = val.Trim().ToLowerInvariant()
                                        If u = "decimal" Then
                                            fmtOverride = "F"
                                        ElseIf u = "scientific" Then
                                            fmtOverride = "E"
                                        ElseIf u = "general" OrElse u = "sig" OrElse u = "significant" Then
                                            fmtOverride = "G"
                                        End If
                                End Select
                            End If
                        End If
                    Next

                    Dim gb = TryCast(GetControlByName(panelName), GroupBox)
                    If gb Is Nothing Then Continue For

                    ' Ensure dictionaries exist (defensive)
                    If Not StatsRows.ContainsKey(panelName) Then StatsRows(panelName) = New List(Of StatsRow)()
                    If Not StatsState.ContainsKey(panelName) Then StatsState(panelName) = New RunningStatsState()

                    Dim rowY As Integer = 18 + (StatsRows(panelName).Count * 18)

                    Dim lblName As New Label()
                    lblName.AutoSize = True
                    lblName.Text = labelText
                    lblName.Location = New Point(10, rowY)
                    lblName.ForeColor = SystemColors.ControlText

                    Dim lblVal As New Label()
                    lblVal.AutoSize = True
                    lblVal.Text = ""
                    lblVal.Location = New Point(gb.Width \ 2, rowY)
                    lblVal.ForeColor = SystemColors.ControlText
                    lblVal.Font = New Font("Consolas", 8.25F)

                    gb.Controls.Add(lblName)
                    gb.Controls.Add(lblVal)

                    StatsRows(panelName).Add(New StatsRow With {
                        .LabelText = labelText,
                        .FuncKey = funcKey,
                        .RefToken = refTok,
                        .FormatOverride = fmtOverride,
                        .ValueLabel = lblVal
                    })

                    ' Draw immediately if we already have samples
                    UpdateStatsPanel(panelName, Double.NaN)
                    Continue For


                Case "HISTORYGRID"
                    ' Supports BOTH:
                    ' Positional:
                    '   HISTORYGRID;GridName;Caption;ResultTarget; x=..;y=..;w=..;h=..;maxrows=..;format=F7;cols=...;ppmref=MEAN
                    '
                    ' Named:
                    '   HISTORYGRID;name=GridName;caption=...;result=TextBoxResult; x=..;y=..;w=..;h=..;maxrows=..;format=F7;cols=...;ppmref=MEAN

                    If parts.Length < 2 Then Continue For

                    Dim gridName As String = ""
                    Dim caption As String = ""
                    Dim resultTarget As String = ""

                    Dim x As Integer = 20
                    Dim y As Integer = 20
                    Dim w As Integer = 500
                    Dim h As Integer = 180
                    Dim maxRows As Integer = 50
                    Dim fmt As String = "F7"
                    Dim colsRaw As String = "Value,Time,Min,Max,PkPk,Mean,Std,PPM,Count"
                    Dim ppmRef As String = "MEAN"

                    Dim firstIsNamed As Boolean = parts.Length > 1 AndAlso parts(1).Contains("="c)

                    If Not firstIsNamed Then
                        ' ----- positional header -----
                        If parts.Length < 4 Then Continue For
                        gridName = parts(1).Trim()
                        caption = parts(2).Trim()
                        resultTarget = parts(3).Trim()

                        ' tokens start after the 3 header fields
                        For i As Integer = 4 To parts.Length - 1
                            Dim p = parts(i).Trim()
                            If Not p.Contains("="c) Then Continue For

                            Dim kv = p.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "x" : ParseIntField(val, x)
                                Case "y" : ParseIntField(val, y)
                                Case "w" : ParseIntField(val, w)
                                Case "h" : ParseIntField(val, h)
                                Case "maxrows"
                                    Dim mr As Integer
                                    If Integer.TryParse(val, mr) AndAlso mr > 0 Then maxRows = mr
                                Case "format"
                                    If val <> "" Then fmt = val
                                Case "cols"
                                    If val <> "" Then colsRaw = val
                                Case "ppmref"
                                    If val <> "" Then ppmRef = val
                            End Select
                        Next

                    Else
                        ' ----- named header + tokens -----
                        For i As Integer = 1 To parts.Length - 1
                            Dim p = parts(i).Trim()
                            If Not p.Contains("="c) Then Continue For

                            Dim kv = p.Split("="c)
                            If kv.Length <> 2 Then Continue For

                            Dim key = kv(0).Trim().ToLowerInvariant()
                            Dim val = kv(1).Trim()

                            Select Case key
                                Case "name" : gridName = val
                                Case "caption" : caption = val
                                Case "result" : resultTarget = val
                                Case "target" : resultTarget = val   ' allow synonym if you ever use it

                                Case "x" : ParseIntField(val, x)
                                Case "y" : ParseIntField(val, y)
                                Case "w" : ParseIntField(val, w)
                                Case "h" : ParseIntField(val, h)

                                Case "maxrows"
                                    Dim mr As Integer
                                    If Integer.TryParse(val, mr) AndAlso mr > 0 Then maxRows = mr

                                Case "format"
                                    If val <> "" Then fmt = val

                                Case "cols"
                                    If val <> "" Then colsRaw = val

                                Case "ppmref"
                                    If val <> "" Then ppmRef = val
                            End Select
                        Next

                        ' must have these in named mode
                        If gridName = "" OrElse resultTarget = "" Then Continue For
                    End If

                    ' ---------- Caption label (optional) ----------
                    Dim topY As Integer = y

                    If caption <> "" Then
                        Dim lbl As New Label()
                        lbl.Text = caption
                        lbl.AutoSize = True
                        lbl.Location = New Point(x, y)
                        GroupBoxCustom.Controls.Add(lbl)
                        topY = y + lbl.Height + 3
                    End If

                    ' ---------- Grid ----------
                    Dim dgv As New DataGridView()
                    dgv.Name = gridName

                    RegisterDynamicControl(gridName, dgv)       ' Invisibility control

                    dgv.Location = New Point(x, topY)
                    dgv.Size = New Size(w, h)
                    dgv.AllowUserToAddRows = False
                    dgv.AllowUserToDeleteRows = False
                    dgv.ReadOnly = True
                    dgv.RowHeadersVisible = False
                    dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
                    dgv.MultiSelect = False
                    dgv.AutoGenerateColumns = False
                    dgv.ScrollBars = ScrollBars.Vertical
                    dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                    ' Reduce flicker: DataGridView.DoubleBuffered (private property)
                    Dim ApplyDgvDoubleBuffer As Action =
                        Sub()
                            Try
                                Dim pi = GetType(DataGridView).GetProperty("DoubleBuffered",
                                    Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
                                If pi IsNot Nothing Then pi.SetValue(dgv, True, Nothing)
                            Catch
                            End Try
                        End Sub

                    ApplyDgvDoubleBuffer()
                    AddHandler dgv.HandleCreated, Sub() ApplyDgvDoubleBuffer()

                    ' Build columns from cols=
                    Dim colNames As New List(Of String)
                    For Each raw As String In colsRaw.Split(","c)
                        Dim s As String = raw.Trim()
                        If s <> "" Then colNames.Add(s)
                    Next
                    If colNames.Count = 0 Then colNames.AddRange(New String() {"Value", "Time"})

                    dgv.Columns.Clear()

                    For Each colName As String In colNames
                        Dim c As New DataGridViewTextBoxColumn()
                        c.Name = colName
                        c.HeaderText = colName
                        c.DataPropertyName = colName   ' MUST match HistoryRow property names
                        c.SortMode = DataGridViewColumnSortMode.NotSortable

                        c.Width = 70
                        If colName.Equals("Value", StringComparison.OrdinalIgnoreCase) Then c.Width = 95
                        If colName.Equals("Time", StringComparison.OrdinalIgnoreCase) Then c.Width = 120
                        If colName.Equals("PkPk", StringComparison.OrdinalIgnoreCase) Then c.Width = 90
                        If colName.Equals("Mean", StringComparison.OrdinalIgnoreCase) Then c.Width = 95

                        If colName.Equals("Time", StringComparison.OrdinalIgnoreCase) Then
                            c.DefaultCellStyle.Format = "HH:mm:ss"
                        ElseIf Not colName.Equals("Count", StringComparison.OrdinalIgnoreCase) Then
                            c.DefaultCellStyle.Format = fmt
                        End If

                        dgv.Columns.Add(c)
                    Next

                    If dgv.Columns.Count > 0 Then
                        dgv.Columns(0).DefaultCellStyle.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
                    End If

                    GroupBoxCustom.Controls.Add(dgv)

                    ' Per-grid history list
                    Dim list As BindingList(Of HistoryRow) = Nothing
                    If Not HistoryData.TryGetValue(gridName, list) Then
                        list = New BindingList(Of HistoryRow)()
                        HistoryData(gridName) = list
                    Else
                        list.Clear()
                    End If

                    dgv.DataSource = list

                    ' Per-grid running stats
                    Dim st As RunningStatsState = Nothing
                    If Not HistoryState.TryGetValue(gridName, st) Then
                        st = New RunningStatsState()
                        HistoryState(gridName) = st
                    Else
                        st.Reset()
                    End If

                    ' Hook to source textbox (NOTE: resultTarget must be just the control name)
                    Dim src = TryCast(GetControlByName(resultTarget), TextBox)
                    If src IsNot Nothing Then
                        AddHandler src.TextChanged,
                            Sub(senderSrc As Object, eSrc As EventArgs)

                                Dim dvSample As Double
                                If Not TryExtractFirstDouble(DirectCast(senderSrc, TextBox).Text, dvSample) Then Exit Sub

                                st.AddSample(dvSample)

                                Dim row As New HistoryRow()
                                row.Time = DateTime.Now
                                row.Value = dvSample
                                row.Min = If(st.Count > 0, st.Min, 0)
                                row.Max = If(st.Count > 0, st.Max, 0)
                                row.PkPk = If(st.Count > 0, st.Max - st.Min, 0)
                                row.Mean = If(st.Count > 0, st.Mean, 0)
                                row.Std = If(st.Count > 1, st.StdDevSample(), 0)

                                Dim refVal As Double = st.Mean
                                If ppmRef.Equals("FIRST", StringComparison.OrdinalIgnoreCase) AndAlso st.HasFirst Then
                                    refVal = st.First
                                ElseIf Not ppmRef.Equals("MEAN", StringComparison.OrdinalIgnoreCase) Then
                                    Dim tmp As Double
                                    If Double.TryParse(ppmRef, Globalization.NumberStyles.Float,
                                                       Globalization.CultureInfo.InvariantCulture, tmp) Then
                                        refVal = tmp
                                    End If
                                End If

                                row.PPM = If(refVal <> 0, ((st.Last - refVal) / refVal) * 1000000.0R, 0)
                                row.Count = CInt(st.Count)

                                list.Insert(0, row)

                                While list.Count > maxRows
                                    list.RemoveAt(list.Count - 1)
                                End While

                            End Sub
                    End If
                    ' NEW: register grid -> target so DATASOURCE updates can drive it (no hidden TextBox needed)
                    HistorySettings(gridName) = New HistoryGridConfig With {
                        .GridName = gridName,
                        .ResultTarget = resultTarget,
                        .MaxRows = maxRows,
                        .Format = fmt,
                        .Columns = colNames,
                        .PpmRef = ppmRef
                    }

                    Dim glist As List(Of String) = Nothing
                    If Not HistoryGridsByTarget.TryGetValue(resultTarget, glist) Then
                        glist = New List(Of String)()
                        HistoryGridsByTarget(resultTarget) = glist
                    End If
                    If Not glist.Contains(gridName, StringComparer.OrdinalIgnoreCase) Then
                        glist.Add(gridName)
                    End If
                    Continue For


                Case "INVISIBILITY"
                    ParseInvisibilityLine(parts)
                    Continue For


                Case "TRIGGER"
                    If parts.Length < 2 Then Continue For
                    If Not parts(1).Contains("="c) Then Continue For   ' named-only

                    Dim trigName As String = ""
                    Dim ifExpr As String = ""
                    Dim thenChain As String = ""

                    Dim periodSec As Double = 0.5
                    Dim cooldownSec As Double = 0.0
                    Dim needHits As Integer = 1
                    Dim oneShot As Boolean = False
                    Dim enabled As Boolean = True

                    For i2 As Integer = 1 To parts.Length - 1
                        Dim p2 = parts(i2).Trim()
                        If Not p2.Contains("="c) Then Continue For

                        Dim kv = p2.Split({"="c}, 2)
                        If kv.Length <> 2 Then Continue For

                        Dim key = kv(0).Trim().ToLowerInvariant()
                        Dim val = kv(1).Trim()

                        Select Case key
                            Case "name"
                                trigName = val

                            Case "if"
                                ifExpr = val

                            Case "then"
                                thenChain = val

                            Case "period"
                                Double.TryParse(val, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, periodSec)

                            Case "cooldown"
                                Double.TryParse(val, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, cooldownSec)

                            Case "need"
                                Integer.TryParse(val, needHits)

                            Case "oneshot"
                                oneShot = (val.Equals("1") OrElse val.Equals("true", StringComparison.OrdinalIgnoreCase) OrElse val.Equals("yes", StringComparison.OrdinalIgnoreCase))

                            Case "enabled"
                                enabled = (val.Equals("1") OrElse val.Equals("true", StringComparison.OrdinalIgnoreCase) OrElse val.Equals("yes", StringComparison.OrdinalIgnoreCase))
                        End Select
                    Next

                    If String.IsNullOrWhiteSpace(trigName) Then Continue For
                    If String.IsNullOrWhiteSpace(ifExpr) Then Continue For
                    If String.IsNullOrWhiteSpace(thenChain) Then Continue For

                    If needHits < 1 Then needHits = 1
                    If periodSec <= 0 Then periodSec = 0.5

                    ' NEW: your trigger engine must exist (or be created here)
                    EnsureTriggerEngine()

                    ' NEW: TriggerDef class/struct used by your engine
                    Dim td As New TriggerDef With {
                        .Name = trigName,
                        .Enabled = enabled,
                        .PeriodSec = periodSec,
                        .IfExpr = ifExpr,
                        .ThenChain = thenChain,
                        .CooldownSec = cooldownSec,
                        .NeedHits = needHits,
                        .OneShot = oneShot
                    }

                    ' Start trigger engine
                    EnsureTriggerEngine()
                    TriggerEng.AddOrUpdate(td)
                    Continue For


                Case "DATASOURCE"
                    Dim resultName As String = ""
                    Dim device As String = ""
                    Dim command As String = ""

                    For Each part In parts
                        Dim kv = part.Split("="c)
                        If kv.Length <> 2 Then Continue For

                        Dim key = kv(0).Trim().ToLowerInvariant()
                        Dim val = kv(1).Trim()

                        Select Case key
                            Case "result"
                                resultName = val
                            Case "device"
                                device = val
                            Case "command", "cmd"
                                command = val
                        End Select
                    Next

                    If resultName <> "" AndAlso device <> "" AndAlso command <> "" Then
                        DataSources(resultName) = New DataSourceDef With {.Device = device, .Command = command}
                    End If


                Case "CALC"
                    ParseCalcLine(parts)
                    Continue For





            End Select

        Next

        ' Apply INVISIBILITY defaults after all controls are created/registered.
        ApplyInvisibilityDefaults()

    End Sub


    ' Apply invisibility
    Private Sub ApplyInvisibilityDefaults()
        If InvisFuncDefaultVisible Is Nothing OrElse InvisFuncDefaultVisible.Count = 0 Then Exit Sub

        For Each kvp In InvisFuncDefaultVisible
            Dim funcName As String = kvp.Key
            Dim defaultVisible As Boolean = kvp.Value

            Dim targets As List(Of String) = Nothing
            If Not InvisFuncToTargets.TryGetValue(funcName, targets) OrElse targets Is Nothing Then Continue For

            For Each id In targets
                Dim c As Control = Nothing

                ' Prefer UiById if present
                If UiById IsNot Nothing AndAlso UiById.TryGetValue(id, c) AndAlso c IsNot Nothing Then
                    c.Visible = defaultVisible
                    If defaultVisible Then c.BringToFront()
                    Continue For
                End If

                ' Fallback: find by name in the GroupBox (recursive)
                Dim found = GroupBoxCustom.Controls.Find(id, True)
                If found IsNot Nothing AndAlso found.Length > 0 AndAlso found(0) IsNot Nothing Then
                    found(0).Visible = defaultVisible
                    If defaultVisible Then found(0).BringToFront()
                End If
            Next
        Next
    End Sub


    ' MINIMAL NAMED PARAM PARSER HELPERS
    Private Function ParseNamedParams(parts() As String) As Dictionary(Of String, String)
        Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For i As Integer = 1 To parts.Length - 1
            Dim s = parts(i).Trim()
            If s.Contains("=") Then
                Dim kv = s.Split({"="c}, 2)
                dict(kv(0).Trim()) = kv(1).Trim()
            End If
        Next
        Return dict
    End Function


    ' Which grids listen to which result name
    Public Class DataSourceDef
        Public Device As String
        Public Command As String
    End Class


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


    Private Sub ParseInvisibilityLine(parts() As String)
        ' Supports BOTH:
        '
        ' 1) Positional pairs (current working):
        '    INVISIBILITY;ChartDMM;ChartDMMhide;Hist1;Hist1hide;Stats1;Stats1hide
        '    INVISIBILITY;ChartDMM,Hist1,Stats1;AllHide
        '
        ' 2) Named tokens (new):
        '    INVISIBILITY;targets=ChartDMM;func=ChartDMMhide;targets=Hist1;func=Hist1hide
        '    INVISIBILITY;targets=ChartDMM,Hist1,Stats1;func=AllHide
        '    INVISIBILITY;func=AllHide;targets=ChartDMM,Hist1,Stats1   (order doesn't matter)

        Dim pendingHasDefault As Boolean = False
        Dim pendingDefaultVisible As Boolean = True

        If parts Is Nothing OrElse parts.Length < 3 Then Exit Sub

        ' Detect "named" usage if any token after INVISIBILITY contains '='
        Dim anyNamed As Boolean = False
        For i As Integer = 1 To parts.Length - 1
            If parts(i) IsNot Nothing AndAlso parts(i).Contains("="c) Then
                anyNamed = True
                Exit For
            End If
        Next

        If Not anyNamed Then
            ' Existing positional behaviour
            Dim i As Integer = 1
            While i + 1 < parts.Length
                Dim targetSpec As String = parts(i).Trim()
                Dim funcName As String = parts(i + 1).Trim()

                If targetSpec <> "" AndAlso funcName <> "" Then
                    Dim targets As New List(Of String)
                    For Each t In targetSpec.Split(","c)
                        Dim tid = t.Trim()
                        If tid <> "" Then targets.Add(tid)
                    Next

                    If targets.Count > 0 Then
                        InvisFuncToTargets(funcName) = targets
                    End If
                End If

                i += 2
            End While

            Exit Sub
        End If

        ' Named tokens behaviour
        Dim pendingTargets As String = ""
        Dim pendingFunc As String = ""

        For i As Integer = 1 To parts.Length - 1
            Dim tok As String = If(parts(i), "").Trim()
            If tok = "" Then Continue For
            If Not tok.Contains("="c) Then Continue For

            Dim kv = tok.Split(New Char() {"="c}, 2)
            If kv.Length <> 2 Then Continue For

            Dim key As String = kv(0).Trim().ToLowerInvariant()
            Dim val As String = kv(1).Trim()

            Select Case key
                Case "targets", "target", "id", "ids"
                    pendingTargets = val

                Case "func", "function", "name"
                    pendingFunc = val

                Case "default"
                    pendingHasDefault = True
                    Dim v As String = val.Trim().ToLowerInvariant()
                    pendingDefaultVisible = (v = "on" OrElse v = "1" OrElse v = "true" OrElse v = "yes")
            End Select

        Next

        ' If they supplied both but in a way that only resolves at the end, commit it
        If pendingTargets <> "" AndAlso pendingFunc <> "" Then
            Dim targets As New List(Of String)
            For Each t In pendingTargets.Split(","c)
                Dim tid = t.Trim()
                If tid <> "" Then targets.Add(tid)
            Next

            If targets.Count > 0 Then
                InvisFuncToTargets(pendingFunc.Trim()) = targets

                If pendingHasDefault Then
                    InvisFuncDefaultVisible(pendingFunc.Trim()) = pendingDefaultVisible
                End If
            End If

            pendingTargets = ""
            pendingFunc = ""
            pendingHasDefault = False
            pendingDefaultVisible = True
        End If

    End Sub


    Private Function GetDeviceByName(name As String) As IODevices.IODevice
        Select Case name.ToLowerInvariant()
            Case "dev1" : Return dev1
            Case "dev2" : Return dev2
            Case Else : Return Nothing
        End Select
    End Function


    Private Sub ParseCalcLine(parts() As String)
        ' CALC;result=PPM;expr=(YourDMM-Ref)/Ref*1000000
        If parts Is Nothing OrElse parts.Length < 2 Then Exit Sub

        Dim outResult As String = ""
        Dim expr As String = ""

        For i As Integer = 1 To parts.Length - 1
            Dim tok As String = If(parts(i), "").Trim()
            If tok = "" OrElse Not tok.Contains("="c) Then Continue For

            Dim kv = tok.Split(New Char() {"="c}, 2)
            If kv.Length <> 2 Then Continue For

            Dim key As String = kv(0).Trim().ToLowerInvariant()
            Dim val As String = kv(1).Trim()

            Select Case key
                Case "result", "name", "out"
                    outResult = val
                Case "expr", "expression"
                    expr = val
            End Select
        Next

        If outResult = "" OrElse expr = "" Then Exit Sub

        ' Build dependencies: identifiers in expr that are not numbers/operators
        Dim deps As List(Of String) = ExtractCalcDeps(expr)

        ' Store
        Dim cd As New CalcDef With {
        .OutResult = outResult.Trim(),
        .Expr = expr,
        .Deps = deps
    }
        CalcDefs(cd.OutResult) = cd

        ' Reverse map deps -> outResults
        For Each dep In deps
            Dim lst As List(Of String) = Nothing
            If Not CalcDeps.TryGetValue(dep, lst) Then
                lst = New List(Of String)()
                CalcDeps(dep) = lst
            End If
            If Not lst.Contains(cd.OutResult, StringComparer.OrdinalIgnoreCase) Then
                lst.Add(cd.OutResult)
            End If
        Next
    End Sub



End Class