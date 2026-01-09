' User customizeable tab - Parsing

Imports System.ComponentModel
Imports Newtonsoft.Json
Imports WinGPIBproj.OnOffLed


Partial Class Formtest

    Private Keypad1Panel As Control

    ' Sub-tab support
    Private UserTabControl As TabControl = Nothing
    Private UserTabPages As New Dictionary(Of String, TabPage)(StringComparer.OrdinalIgnoreCase)
    Private CurrentUserContainer As Control = Nothing


    Private Sub BuildCustomGuiFromText(def As String)

        UserConfig_DataSaveEnabled = True   ' reset to default for this config

        ' --- Check whether this config actually needs dev1/dev2 ---
        Dim needsDev1 As Boolean =
        def.IndexOf("device=dev1", StringComparison.OrdinalIgnoreCase) >= 0
        Dim needsDev2 As Boolean =
        def.IndexOf("device=dev2", StringComparison.OrdinalIgnoreCase) >= 0

        Dim missingDev1 As Boolean = needsDev1 AndAlso Not gbox1.Enabled
        Dim missingDev2 As Boolean = needsDev2 AndAlso Not gbox2.Enabled

        ' Only block if at least one required device is missing.
        If missingDev1 OrElse missingDev2 Then
            Dim msg As String =
        "No GPIB devices are currently connected." & vbCrLf & vbCrLf &
        "Start Device1 and/or Device2, then reload the User config."

            MessageBox.Show(msg, "WinGPIB",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)

            ' This already restores LabelUSERtab1 and the blank state
            ResetUsertab()

            Exit Sub
        End If


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

        ' Reset sub-tab layout state
        UserTabControl = Nothing
        UserTabPages.Clear()
        CurrentUserContainer = Nothing

        CurrentUserScale = 1.0   ' reset scale each time we load a layout
        CurrentUserUnit = ""

        Dim autoY As Integer = 10

        ' -----------------------------
        ' Multi-line logical line joiner
        ' -----------------------------
        Dim physicalLines = def.Split({vbCrLf, vbLf}, StringSplitOptions.None)
        Dim logicalLines As New List(Of String)
        Dim pending As String = Nothing

        For Each rawLine In physicalLines

            ' First detect LUA using the original text (no comment stripping)
            Dim trimmedEnd As String = rawLine.TrimEnd()
            Dim lineTrimmed As String = trimmedEnd.Trim()

            ' ================= LUA SCRIPT BLOCK =================
            If inLuaScript Then
                If lineTrimmed.StartsWith("LUASCRIPTEND", StringComparison.OrdinalIgnoreCase) Then
                    LuaScriptsByName(luaScriptName) = String.Join(vbCrLf, luaLines)
                    luaLines.Clear()
                    inLuaScript = False
                Else
                    ' store ORIGINAL rawLine to preserve indentation in LUA
                    luaLines.Add(rawLine)
                End If
                Continue For
            End If

            ' For NON-LUA lines, strip any inline ";;" comment BEFORE we do joins
            Dim effectiveLine As String = rawLine
            Dim ccIdx As Integer = effectiveLine.IndexOf(";;", StringComparison.Ordinal)
            If ccIdx >= 0 Then
                effectiveLine = effectiveLine.Substring(0, ccIdx)
            End If

            trimmedEnd = effectiveLine.TrimEnd()
            lineTrimmed = trimmedEnd.Trim()

            ' Skip empty/comment lines (do NOT flush pending)
            If lineTrimmed = "" OrElse lineTrimmed.StartsWith(";"c) Then
                Continue For
            End If

            Dim startsIndented As Boolean =
        effectiveLine.Length > 0 AndAlso Char.IsWhiteSpace(effectiveLine(0))

            If pending Is Nothing Then
                pending = lineTrimmed
            Else
                Dim pendingEndsWithSemi As Boolean =
            pending.TrimEnd().EndsWith(";", StringComparison.Ordinal)

                ' Detect if current pending block is a BOOTCOMMANDS block
                Dim isBootBlock As Boolean =
            pending.TrimStart().StartsWith("BOOTCOMMANDS", StringComparison.OrdinalIgnoreCase)

                ' Continuation rule:
                ' - continuation lines must be indented
                ' - AND (previous pending line ends with ";" OR we are inside BOOTCOMMANDS block)
                If startsIndented AndAlso (pendingEndsWithSemi OrElse isBootBlock) Then
                    pending &= lineTrimmed
                Else
                    ' flush previous logical line
                    logicalLines.Add(pending.Trim())
                    pending = lineTrimmed
                End If
            End If

            ' Start LUA block only when NOT already in one (must be checked here so it can begin on a logical line)
            If lineTrimmed.StartsWith("LUASCRIPTBEGIN", StringComparison.OrdinalIgnoreCase) Then
                ' flush any pending logical line before beginning LUA
                If pending IsNot Nothing Then
                    logicalLines.Add(pending.Trim())
                    pending = Nothing
                End If

                luaScriptName = GetParam(lineTrimmed, "name")
                luaLines.Clear()
                inLuaScript = True
                Continue For
            End If

        Next

        ' Flush final pending logical line
        If pending IsNot Nothing Then
            logicalLines.Add(pending.Trim())
        End If



        ' -----------------------------
        ' Existing parser works on logical lines
        ' -----------------------------
        For Each line In logicalLines

            If line = "" OrElse line.StartsWith(";") Then Continue For

            ' Per-device GPIB engine selectors for USER tab
            If line.StartsWith("GPIBEngineDev1", StringComparison.OrdinalIgnoreCase) Then
                Dim eq = line.IndexOf("="c)
                If eq > 0 Then
                    Dim val As String = line.Substring(eq + 1).Trim().TrimEnd(";"c)
                    GpibEngineDev1 = If(
            String.Equals(val, "native", StringComparison.OrdinalIgnoreCase),
            "native",
            "standalone"
        )
                End If
                Continue For
            End If

            If line.StartsWith("GPIBEngineDev2", StringComparison.OrdinalIgnoreCase) Then
                Dim eq = line.IndexOf("="c)
                If eq > 0 Then
                    Dim val As String = line.Substring(eq + 1).Trim().TrimEnd(";"c)
                    GpibEngineDev2 = If(
            String.Equals(val, "native", StringComparison.OrdinalIgnoreCase),
            "native",
            "standalone"
        )
                End If
                Continue For
            End If

            ' BOOTCOMMANDS; device=...; commandlist=...
            If line.StartsWith("BOOTCOMMANDS", StringComparison.OrdinalIgnoreCase) Then
                HandleBootCommandsLine(line)
                Continue For
            End If

            ' LUASCRIPTBEGIN;name=MyScript
            If line.StartsWith("LUASCRIPTBEGIN", StringComparison.OrdinalIgnoreCase) Then
                luaScriptName = GetParam(line, "name")
                luaLines.Clear()
                inLuaScript = True
                Continue For
            End If

            ' -----------------------------
            ' Simple KEY=VALUE directives (semicolon optional)
            ' e.g. DATASAVE=enabled   or   DATASAVE=enabled;
            ' -----------------------------
            If line.Contains("="c) Then
                Dim clean = line.Trim()

                ' remove trailing semicolon if present
                clean = clean.TrimEnd(";"c)

                Dim kv = clean.Split({"="c}, 2)
                If kv.Length = 2 Then
                    Dim k = kv(0).Trim().ToLowerInvariant()
                    Dim v = kv(1).Trim().ToLowerInvariant()

                    Select Case k
                        Case "datasave"
                            UserConfig_DataSaveEnabled =
                    (v = "enabled" OrElse v = "true" OrElse v = "1")
                            Continue For
                    End Select
                End If
            End If

            ' -------- INLINE COMMENT SUPPORT ( ;; ) --------
            Dim commentIdx As Integer = line.IndexOf(";;", StringComparison.Ordinal)
            If commentIdx >= 0 Then
                line = line.Substring(0, commentIdx).TrimEnd()
                If line = "" Then Continue For
            End If
            ' -----------------------------------------------

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


                Case "TAB"

                    Dim kv = ParseNamedParams(parts)

                    Dim tabName As String = Nothing
                    If kv.ContainsKey("name") Then
                        tabName = kv("name").Trim()
                    End If
                    If String.IsNullOrEmpty(tabName) Then
                        tabName = "Tab" & (UserTabPages.Count + 1).ToString()
                    End If

                    ' First TAB: create the TabControl and move existing controls to a "Main" tab
                    If UserTabControl Is Nothing Then

                        UserTabControl = New TabControl()
                        UserTabControl.Name = "UserSubTabs"
                        UserTabControl.Left = 5
                        UserTabControl.Top = 15
                        UserTabControl.Width = GroupBoxCustom.ClientSize.Width - 10
                        UserTabControl.Height = GroupBoxCustom.ClientSize.Height - 20
                        UserTabControl.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right

                        ' Default "Main" tab containing everything created before the first TAB;
                        Dim mainPage As New TabPage("Main")
                        mainPage.Name = "tabMain"

                        Dim existingCount As Integer = GroupBoxCustom.Controls.Count
                        If existingCount > 0 Then
                            Dim existing(existingCount - 1) As Control
                            GroupBoxCustom.Controls.CopyTo(existing, 0)
                            GroupBoxCustom.Controls.Clear()
                            For Each c As Control In existing
                                mainPage.Controls.Add(c)
                            Next
                        End If

                        UserTabPages("Main") = mainPage
                        UserTabControl.TabPages.Add(mainPage)

                        GroupBoxCustom.Controls.Add(UserTabControl)
                    End If

                    ' Reuse tab if name already seen, otherwise create
                    Dim targetPage As TabPage = Nothing
                    If Not UserTabPages.TryGetValue(tabName, targetPage) Then
                        targetPage = New TabPage(tabName)
                        targetPage.Name = "tab" & tabName
                        UserTabPages(tabName) = targetPage
                        UserTabControl.TabPages.Add(targetPage)
                    End If

                    ' From now on, controls go onto this tab
                    CurrentUserContainer = targetPage

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
                    AddToUserContainer(lbl)

                    ' TextBox to the right of label
                    Dim tb As New TextBox()
                    tb.Name = tbName
                    tb.Location = New Point(lbl.Right + 5, y - 2)
                    tb.Size = New Size(w, h)
                    tb.ReadOnly = isReadOnly

                    ' Register so triggers/actions can reference this control by name
                    RegisterAnyControl(tbName, tb)

                    ' Publish changes for trigger engine (user typing OR programmatic changes)
                    AddHandler tb.TextChanged,
        Sub()
            PublishTextBox(tb.Name, tb.Text)
        End Sub

                    ' Arm keypad target when user clicks/focuses this textbox
                    AddHandler tb.Enter, AddressOf UserKeypad_SetTarget
                    AddHandler tb.MouseDown, AddressOf UserKeypad_SetTarget

                    AddToUserContainer(tb)

                    ' Publish initial (empty or preset) value once
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

                    ' Accept common aliases so config can use your doc terms
                    Dim hasCaption As Boolean = named.ContainsKey("caption")
                    Dim hasAction As Boolean = named.ContainsKey("action")
                    Dim hasDevice As Boolean = named.ContainsKey("device") OrElse named.ContainsKey("devicename")
                    Dim hasName As Boolean = named.ContainsKey("name") OrElse named.ContainsKey("controlname")

                    Dim isAllNamed As Boolean =
                        hasCaption OrElse hasAction OrElse hasDevice OrElse hasName

                    ' ==================================================
                    ' ALL-NAMED MODE
                    ' ==================================================
                    If isAllNamed Then

                        ' Button internal id used by TRIGGER fire:...
                        Dim btnName As String = ""
                        If named.ContainsKey("name") Then btnName = named("name")
                        If btnName = "" AndAlso named.ContainsKey("controlname") Then btnName = named("controlname")

                        Dim caption As String = If(named.ContainsKey("caption"), named("caption"), "")

                        Dim action As String = If(named.ContainsKey("action"), named("action"), "").ToUpperInvariant()

                        ' Accept device/deviceName
                        Dim deviceName As String = ""
                        If named.ContainsKey("device") Then deviceName = named("device")
                        If deviceName = "" AndAlso named.ContainsKey("devicename") Then deviceName = named("devicename")

                        ' Accept command/commandprefix
                        Dim commandOrPrefix As String = ""
                        If named.ContainsKey("command") Then commandOrPrefix = named("command")
                        If commandOrPrefix = "" AndAlso named.ContainsKey("commandprefix") Then commandOrPrefix = named("commandprefix")

                        ' Accept sendval/valuectl/valuecontrol
                        Dim valueControlName As String = ""
                        If named.ContainsKey("sendval") Then valueControlName = named("sendval")
                        If valueControlName = "" AndAlso named.ContainsKey("valuectl") Then valueControlName = named("valuectl")
                        If valueControlName = "" AndAlso named.ContainsKey("valuecontrol") Then valueControlName = named("valuecontrol")

                        ' Accept result/resultctl/resulttarget
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

                        ' Set internal name so TRIGGER fire:BtnHello can find it
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
                        AddToUserContainer(b)

                        ' Register button for trigger engine fire:
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
                    AddToUserContainer(bPos)

                    ' Positional buttons don't have a name -> triggers can't fire them unless you invent one.
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

                    AddToUserContainer(cb)

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

                    AddToUserContainer(lbl)

                    Continue For


                Case "DROPDOWN"
                    ' Supports BOTH:
                    ' 1) Positional:
                    '    DROPDOWN;ControlName;Caption;DeviceName;CommandPrefix;X;Y;Width;Item1,Item2,...;captionpos=1/2;determine=...;commands=...;detmap=...
                    '
                    ' 2) All-named:
                    '    DROPDOWN;name=Ctrl1;caption=Range;device=dev1;command=:SENS:...;commands=cmd1,cmd2,...;x=..;y=..;w=..;items=..;captionpos=..;determine=..;detmap=..
                    '
                    ' determine syntax:
                    '   determine=<query>                 -> defaults respnum
                    '   determine=<query>|respnum
                    '   determine=<query>|resptext
                    '
                    ' For resptext dropdowns, you can add:
                    '   detmap=VOLT:AC,VOLT,=RES,=FRES,...
                    '   - map count must match the dropdown items count (excluding placeholder)
                    '   - "=TOKEN" means exact match; otherwise starts-with match

                    Dim ctrlName As String = ""
                    Dim caption As String = ""
                    Dim deviceName As String = ""
                    Dim commandPrefix As String = ""
                    Dim commandsList As String = ""
                    Dim detmapRaw As String = ""

                    Dim x As Integer = 0
                    Dim y As Integer = 0
                    Dim w As Integer = 120
                    Dim itemsRaw As String = ""
                    Dim captionPos As Integer = 1

                    Dim determineRaw As String = ""

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

                                Select Case key
                                    Case "captionpos"
                                        Integer.TryParse(val, captionPos)
                                        If captionPos <> 1 AndAlso captionPos <> 2 Then captionPos = 1

                                    Case "determine"
                                        determineRaw = val

                                    Case "command", "cmd", "commandprefix"
                                        commandPrefix = val

                                    Case "commands", "cmdlist"
                                        commandsList = val

                                    Case "detmap"
                                        detmapRaw = val
                                End Select
                            Next

                        End If

                    Else
                        ' ---------- all-named ----------
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

                                Case "command", "cmd", "commandprefix"
                                    commandPrefix = val

                                Case "commands", "cmdlist"
                                    commandsList = val

                                Case "detmap"
                                    detmapRaw = val

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

                                Case "determine"
                                    determineRaw = val
                            End Select
                        Next

                        If String.IsNullOrWhiteSpace(ctrlName) OrElse
                           String.IsNullOrWhiteSpace(deviceName) OrElse
                           (String.IsNullOrWhiteSpace(commandPrefix) AndAlso String.IsNullOrWhiteSpace(commandsList)) OrElse
                           String.IsNullOrWhiteSpace(itemsRaw) Then
                            Continue For
                        End If
                    End If

                    Dim items = itemsRaw.Split(","c)

                    Dim cb As New ComboBox()
                    cb.Name = ctrlName
                    cb.DropDownStyle = ComboBoxStyle.DropDownList

                    If captionPos = 1 Then
                        Dim lbl As New Label()
                        lbl.Text = caption
                        lbl.AutoSize = True
                        lbl.Location = New Point(x, y)
                        AddToUserContainer(lbl)

                        cb.Location = New Point(x + lbl.PreferredWidth + 8, y - 3)
                    Else
                        cb.Location = New Point(x, y)
                    End If

                    cb.Size = New Size(w, cb.Height)

                    If captionPos = 1 Then
                        cb.Items.Add("")
                    Else
                        cb.Items.Add(caption)
                    End If

                    For Each it In items
                        Dim s As String = it.Trim()
                        If s <> "" Then cb.Items.Add(s)
                    Next

                    cb.SelectedIndex = 0

                    ' ---------------------------------------------------------
                    ' Tag formats:
                    '   Prefix mode:  dev|<prefix>|DETQ=...|DETF=...|DETMAP=...
                    '   List mode:    dev|CMDLIST|cmd1,cmd2,...|DETQ=...|DETF=...|DETMAP=...
                    ' ---------------------------------------------------------
                    If Not String.IsNullOrWhiteSpace(commandsList) Then
                        cb.Tag = deviceName & "|CMDLIST|" & commandsList.Trim()
                    Else
                        cb.Tag = deviceName & "|" & commandPrefix.Trim()
                    End If

                    If Not String.IsNullOrWhiteSpace(determineRaw) Then
                        Dim dp() As String = determineRaw.Split("|"c)
                        Dim detQ As String = If(dp.Length >= 1, dp(0).Trim(), "")
                        Dim detF As String = "respnum"

                        ' take the LAST non-empty format token (handles CONF?||resptext)
                        For k As Integer = 1 To dp.Length - 1
                            Dim t As String = dp(k).Trim()
                            If t <> "" Then detF = t.ToLowerInvariant()
                        Next

                        If detQ <> "" Then
                            cb.Tag = CStr(cb.Tag) & "|DETQ=" & detQ & "|DETF=" & detF
                        End If
                    End If

                    If Not String.IsNullOrWhiteSpace(detmapRaw) Then
                        cb.Tag = CStr(cb.Tag) & "|DETMAP=" & detmapRaw.Trim()
                    End If

                    AddHandler cb.SelectedIndexChanged, AddressOf Dropdown_SelectedIndexChanged
                    AddToUserContainer(cb)


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

                    AddToUserContainer(gb)

                    Continue For


                Case "RADIO"
                    If parts.Length < 2 Then Continue For

                    Dim groupName As String = ""
                    Dim caption As String = ""
                    Dim deviceName As String = ""
                    Dim command As String = ""

                    Dim overloadToken As String = ""
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

                    ' ============================
                    '  NEW: multiline AUTO folding
                    '  scale=auto;
                    '     getrange=:VOLT:DC:RANG?;
                    '     range1=0.2,mV,6;
                    '     range2=2.0,V,7;
                    '     ...
                    ' becomes:
                    '  scale=auto|:VOLT:DC:RANG?|0.2:mV:6,2.0:V:7,...
                    ' ============================
                    If tok.ContainsKey("scale") Then
                        Dim s As String = tok("scale").Trim()
                        If s.Equals("auto", StringComparison.OrdinalIgnoreCase) AndAlso tok.ContainsKey("getrange") Then

                            Dim getRange As String = tok("getrange").Trim()

                            Dim specs As New List(Of Tuple(Of Integer, String))()

                            ' collect rangeN entries and sort by N
                            For Each kvp In tok
                                Dim key As String = kvp.Key
                                If key.Length > 5 AndAlso key.StartsWith("range", StringComparison.OrdinalIgnoreCase) Then
                                    Dim nStr As String = key.Substring(5)
                                    Dim n As Integer
                                    If Integer.TryParse(nStr, n) Then
                                        Dim v As String = kvp.Value.Trim()
                                        ' "0.2,mV,6" -> "0.2:mV:6"
                                        v = v.Replace(",", ":")
                                        specs.Add(Tuple.Create(n, v))
                                    End If
                                End If
                            Next

                            specs.Sort(Function(a, b) a.Item1.CompareTo(b.Item1))

                            Dim partsList As New List(Of String)()
                            For Each sItem In specs
                                If sItem.Item2 <> "" Then partsList.Add(sItem.Item2)
                            Next

                            Dim newScale As String
                            If partsList.Count > 0 Then
                                newScale = "auto|" & getRange & "|" & String.Join(",", partsList)
                            Else
                                newScale = "auto|" & getRange
                            End If

                            ' overwrite scale so the original code sees the old one-line form
                            tok("scale") = newScale
                        End If
                    End If
                    ' ============================
                    '  END multiline AUTO folding
                    ' ============================

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
                    Dim autoMapSpec As String = ""

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

                            ' ==============================
                            '  multiline AUTO support
                            '  scale=auto;
                            '      getrange=:VOLT:DC:RANG?;
                            '      range1=0.2,mV,6;
                            '      range2=2.0,V,7;
                            '      ...
                            ' ==============================
                            If tok.ContainsKey("getrange") Then
                                ' Multiline style
                                rangeQueryForAuto = tok("getrange").Trim()

                                Dim specs As New List(Of Tuple(Of Integer, String))()

                                ' collect rangeN entries and sort by N
                                For Each kvp In tok
                                    Dim key As String = kvp.Key
                                    If key.Length > 5 AndAlso key.StartsWith("range", StringComparison.OrdinalIgnoreCase) Then
                                        Dim nStr As String = key.Substring(5)
                                        Dim n As Integer
                                        If Integer.TryParse(nStr, n) Then
                                            Dim v As String = kvp.Value.Trim()
                                            ' "0.2,mV,6" -> "0.2:mV:6"
                                            v = v.Replace(",", ":")
                                            specs.Add(Tuple.Create(n, v))
                                        End If
                                    End If
                                Next

                                specs.Sort(Function(a, b) a.Item1.CompareTo(b.Item1))

                                Dim partsList As New List(Of String)()
                                For Each sItem In specs
                                    If sItem.Item2 <> "" Then partsList.Add(sItem.Item2)
                                Next

                                If partsList.Count > 0 Then
                                    autoMapSpec = String.Join(",", partsList)
                                Else
                                    autoMapSpec = ""
                                End If

                            Else
                                ' OLD one-line form: auto|<rangeQuery>|<mapSpec>
                                Dim autoParts() As String = scaleVal.Split("|"c)
                                If autoParts.Length >= 2 Then
                                    rangeQueryForAuto = autoParts(1).Trim()
                                End If
                                If autoParts.Length >= 3 Then
                                    autoMapSpec = autoParts(2).Trim().TrimEnd(";"c)
                                End If
                            End If

                        Else
                            ' Normal numeric scale
                            Double.TryParse(scaleVal,
                        Globalization.NumberStyles.Float,
                        Globalization.CultureInfo.InvariantCulture,
                        scale)
                        End If
                    End If

                    ' Decimal places (optional) dp=5
                    Dim dpVal As Integer = -1
                    If tok.ContainsKey("dp") Then
                        Integer.TryParse(tok("dp"), dpVal)
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

                    ' FIRST: set the base Tag (device|command|scale OR device|command|AUTO|rangeQuery|MAP=...)
                    If scaleIsAuto Then
                        rb.Tag = deviceName & "|" & command & "|AUTO|" & rangeQueryForAuto
                        If autoMapSpec <> "" Then
                            rb.Tag = CStr(rb.Tag) & "|MAP=" & autoMapSpec
                        End If
                    Else
                        rb.Tag = deviceName & "|" & command & "|" &
                                 scale.ToString(Globalization.CultureInfo.InvariantCulture)
                    End If

                    ' THEN: append determine (optional)
                    ' determine=<query>|<expected>|resptext
                    ' If resptext is omitted, numeric compare (respnum) is assumed.
                    If tok.ContainsKey("determine") Then
                        Dim det As String = tok("determine").Trim()
                        If det <> "" Then
                            Dim dp2() As String = det.Split("|"c)
                            Dim detQ As String = If(dp2.Length >= 1, dp2(0).Trim(), "")
                            Dim detE As String = If(dp2.Length >= 2, dp2(1).Trim(), "")
                            Dim detF As String = If(dp2.Length >= 3, dp2(2).Trim().ToLowerInvariant(), "") ' optional

                            If detQ <> "" AndAlso detE <> "" Then
                                rb.Tag = CStr(rb.Tag) & "|DETQ=" & detQ & "|DETE=" & detE
                                If detF <> "" Then rb.Tag = CStr(rb.Tag) & "|DETF=" & detF
                            End If
                        End If
                    End If

                    ' Append DP info to Tag if specified
                    If dpVal >= 0 Then
                        rb.Tag = CStr(rb.Tag) & "|DP=" & dpVal.ToString()
                    End If

                    AddHandler rb.CheckedChanged, AddressOf Radio_CheckedChanged
                    gb.Controls.Add(rb)
                    Continue For


                Case "SLIDER"

                    Dim determineRaw As String = ""

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

                        ' Look for determine= in any trailing key=value segments
                        If parts.Length > 12 Then
                            For i As Integer = 12 To parts.Length - 1
                                Dim p As String = parts(i).Trim()
                                If Not p.Contains("="c) Then Continue For
                                Dim kv = p.Split("="c)
                                If kv.Length <> 2 Then Continue For

                                Dim key = kv(0).Trim().ToLowerInvariant()
                                Dim val = kv(1).Trim()
                                If key = "determine" Then
                                    determineRaw = val
                                    Exit For
                                End If
                            Next
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

                        ' Determine from named tokens
                        If determineRaw = "" AndAlso tok.ContainsKey("determine") Then
                            determineRaw = tok("determine").Trim()
                        End If
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
                    AddToUserContainer(gb)

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

                    ' Append DETQ/DETF for determine support
                    If Not String.IsNullOrWhiteSpace(determineRaw) Then
                        Dim dp() As String = determineRaw.Split("|"c)
                        Dim detQ As String = If(dp.Length >= 1, dp(0).Trim(), "")
                        Dim detF As String = If(dp.Length >= 2 AndAlso dp(1).Trim() <> "",
                                dp(1).Trim().ToLowerInvariant(),
                                "respnum")
                        If detQ <> "" Then
                            tb.Tag = CStr(tb.Tag) & "|DETQ=" & detQ & "|DETF=" & detF
                        End If
                    End If

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

                    AddToUserContainer(hrLine)
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

                    AddToUserContainer(ln)
                    Continue For


                Case "SPINNER"

                    Dim determineRaw As String = ""

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
                                    Case "determine"
                                        determineRaw = val
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
                                Case "determine"
                                    determineRaw = val
                            End Select
                        Next

                        If controlName = "" OrElse deviceName = "" OrElse commandPrefix = "" Then
                            Continue For
                        End If
                    End If

                    ' -----------------------
                    ' Basic validation
                    ' -----------------------
                    If maxVal < minVal Then
                        Dim tmp = minVal : minVal = maxVal : maxVal = tmp
                    End If
                    If stepVal <= 0 Then stepVal = 1

                    ' -----------------------
                    ' Label
                    ' -----------------------
                    Dim lbl As New Label()
                    lbl.Text = caption
                    lbl.AutoSize = True
                    lbl.Location = New Point(x, y + 3)
                    AddToUserContainer(lbl)

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

                    ' Tag = DEVICE|COMMAND|SCALE|INITFLAG|MIN   (existing)
                    nud.Tag = deviceName & "|" &
              commandPrefix & "|" &
              scale.ToString(Globalization.CultureInfo.InvariantCulture) & "|" &
              "0" & "|" &
              minVal.ToString(Globalization.CultureInfo.InvariantCulture)

                    ' Append DETQ/DETF for determine support (optional)
                    If Not String.IsNullOrWhiteSpace(determineRaw) Then
                        Dim dp() As String = determineRaw.Split("|"c)
                        Dim detQ As String = If(dp.Length >= 1, dp(0).Trim(), "")
                        Dim detF As String = If(dp.Length >= 2 AndAlso dp(1).Trim() <> "",
                                dp(1).Trim().ToLowerInvariant(),
                                "respnum")
                        If detQ <> "" Then
                            nud.Tag = CStr(nud.Tag) & "|DETQ=" & detQ & "|DETF=" & detF
                        End If
                    End If

                    AddHandler nud.ValueChanged, AddressOf Spinner_ValueChanged

                    AddToUserContainer(nud)
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
                    AddToUserContainer(lbl)

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

                    AddToUserContainer(led)
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

                        AddToUserContainer(panel)

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
                    AddToUserContainer(lbl)

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

                    AddToUserContainer(tb)

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
                    Dim determineRaw As String = ""
                    Dim onColorName As String = "default"
                    Dim offColorName As String = "default"

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
                                Case "determine" : determineRaw = val
                                Case "oncolor" : onColorName = val.ToLowerInvariant()
                                Case "offcolor" : offColorName = val.ToLowerInvariant()
                            End Select
                        Next
                    End If

                    If name = "" OrElse device = "" OrElse cmdOn = "" OrElse cmdOff = "" Then Continue For

                    ' Create toggle button
                    Dim b As New Button()
                    b.UseVisualStyleBackColor = False
                    b.Name = name
                    b.Text = If(caption <> "", caption, name)
                    b.Location = New Point(x, y)
                    b.Size = New Size(w, h)

                    ' Tag holds: DEVICE|ONCMD|OFFCMD|STATE plus optional DETQ/DETE/DETF + ONCLR/OFFCLR
                    Dim tagStr As String = device & "|" & cmdOn & "|" & cmdOff & "|0"   ' state 0 = OFF

                    ' Optional: determine=<query>|<expected>|resptext
                    ' If resptext is omitted, numeric compare (respnum) is assumed.
                    If determineRaw <> "" Then
                        Dim dp() As String = determineRaw.Split("|"c)
                        Dim detQ As String = If(dp.Length >= 1, dp(0).Trim(), "")
                        Dim detE As String = If(dp.Length >= 2, dp(1).Trim(), "")
                        Dim detF As String = If(dp.Length >= 3, dp(2).Trim().ToLowerInvariant(), "") ' optional

                        If detQ <> "" AndAlso detE <> "" Then
                            tagStr &= "|DETQ=" & detQ & "|DETE=" & detE
                            If detF <> "" Then tagStr &= "|DETF=" & detF
                        End If
                    End If

                    ' Colours (optional)
                    tagStr &= "|ONCLR=" & onColorName & "|OFFCLR=" & offColorName

                    b.Tag = tagStr

                    AddHandler b.Click, AddressOf ToggleButton_Click
                    AddToUserContainer(b)
                    Continue For


                Case "TOGGLEDUAL"
                    If parts.Length < 2 Then Continue For

                    Dim name As String = ""
                    Dim caption As String = ""
                    Dim device As String = ""
                    Dim leftText As String = "ON"
                    Dim rightText As String = "OFF"
                    Dim cmdOn As String = ""
                    Dim cmdOff As String = ""
                    Dim determineRaw As String = ""
                    Dim onColorName As String = "default"
                    Dim offColorName As String = "default"

                    Dim x As Integer = 20, y As Integer = 20, w As Integer = 180, h As Integer = 35

                    ' tokens
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

                    If tok.ContainsKey("name") Then name = tok("name").Trim()
                    If tok.ContainsKey("caption") Then caption = tok("caption").Trim()
                    If tok.ContainsKey("device") Then device = tok("device").Trim()

                    If tok.ContainsKey("left") Then leftText = tok("left").Trim()
                    If tok.ContainsKey("right") Then rightText = tok("right").Trim()

                    If tok.ContainsKey("on") Then cmdOn = tok("on").Trim()
                    If tok.ContainsKey("off") Then cmdOff = tok("off").Trim()

                    If tok.ContainsKey("x") Then ParseIntField(tok("x"), x)
                    If tok.ContainsKey("y") Then ParseIntField(tok("y"), y)
                    If tok.ContainsKey("w") Then ParseIntField(tok("w"), w)
                    If tok.ContainsKey("h") Then ParseIntField(tok("h"), h)

                    If tok.ContainsKey("determine") Then determineRaw = tok("determine").Trim()

                    If tok.ContainsKey("oncolor") OrElse tok.ContainsKey("leftcolor") Then
                        onColorName = If(tok.ContainsKey("oncolor"), tok("oncolor"), tok("leftcolor")).Trim().ToLowerInvariant()
                    End If
                    If tok.ContainsKey("offcolor") OrElse tok.ContainsKey("rightcolor") Then
                        offColorName = If(tok.ContainsKey("offcolor"), tok("offcolor"), tok("rightcolor")).Trim().ToLowerInvariant()
                    End If

                    If name = "" OrElse device = "" OrElse cmdOn = "" OrElse cmdOff = "" Then Continue For

                    ' Optional caption label
                    Dim capW As Integer = 0
                    If caption <> "" Then
                        Dim lbl As New Label()
                        lbl.Text = caption
                        lbl.AutoSize = True
                        lbl.Location = New Point(x, y + 7)
                        AddToUserContainer(lbl)
                        capW = lbl.PreferredWidth + 8
                    End If

                    ' Host panel
                    Dim pnl As New Panel()
                    pnl.Name = "TD_" & name
                    pnl.Location = New Point(x + capW, y)
                    pnl.Size = New Size(w, h)
                    'pnl.BorderStyle = BorderStyle.FixedSingle
                    AddToUserContainer(pnl)

                    ' Two buttons
                    Dim bL As New Button()
                    bL.Name = "TDL_" & name
                    bL.Text = leftText
                    bL.Location = New Point(0, 0)
                    bL.Size = New Size(w \ 2, h)
                    bL.UseVisualStyleBackColor = False

                    Dim bR As New Button()
                    bR.Name = "TDR_" & name
                    bR.Text = rightText
                    bR.Location = New Point(w \ 2, 0)
                    bR.Size = New Size(w - (w \ 2), h)
                    bR.UseVisualStyleBackColor = False

                    ' Tag format (both buttons share the same Tag):
                    ' DEVICE|ONCMD|OFFCMD|STATE|PANELNAME plus optional DETQ/DETE/DETF + ONCLR/OFFCLR
                    ' STATE: 1 = LEFT active (ON), 0 = RIGHT active (OFF)
                    Dim tagStr As String = device & "|" & cmdOn & "|" & cmdOff & "|0|" & pnl.Name

                    ' append determine tokens if present
                    If determineRaw <> "" Then
                        Dim dp() As String = determineRaw.Split("|"c)
                        Dim detQ As String = If(dp.Length >= 1, dp(0).Trim(), "")
                        Dim detE As String = If(dp.Length >= 2, dp(1).Trim(), "")
                        Dim detF As String = If(dp.Length >= 3, dp(2).Trim().ToLowerInvariant(), "") ' optional
                        If detQ <> "" AndAlso detE <> "" Then
                            tagStr &= "|DETQ=" & detQ & "|DETE=" & detE
                            If detF <> "" Then tagStr &= "|DETF=" & detF
                        End If
                    End If

                    ' Colours (optional)
                    tagStr &= "|ONCLR=" & onColorName & "|OFFCLR=" & offColorName

                    bL.Tag = tagStr
                    bR.Tag = tagStr

                    AddHandler bL.Click, AddressOf ToggleDual_LeftClick
                    AddHandler bR.Click, AddressOf ToggleDual_RightClick

                    pnl.Controls.Add(bL)
                    pnl.Controls.Add(bR)

                    ' Default visual = OFF (right) (colour applied in ApplyToggleDualVisual)
                    ApplyToggleDualVisual(pnl, False)

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

                    ' font size for y-axis
                    Dim labelSize As Single = 7.0F

                    Dim usedPositional As Boolean = False
                    Dim popup As Boolean = False

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
                            Case "labelsize"
                                Single.TryParse(val, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, labelSize)

                            Case "popup"
                                Dim v = val.ToLowerInvariant()
                                popup = (v = "1" OrElse v = "true" OrElse v = "yes")

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
                        AddToUserContainer(lblCap)

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
                    ca.AxisY.LabelStyle.Font = New Font("Segoe UI", labelSize)
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

                    AddToUserContainer(ch)

                    ' Remember config for pop-out charts
                    ChartSettings(chartName) = New ChartConfig With {
        .ChartName = chartName,
        .ResultTarget = resultTarget,
        .YMin = yMin,
        .YMax = yMax,
        .XStep = xStep,
        .MaxPoints = maxPoints,
        .AutoScaleY = autoScaleY,
        .Popup = popup,
        .Width = w,
        .Height = h,
        .InnerX = innerX,
        .InnerY = innerY,
        .InnerW = innerW,
        .InnerH = innerH,
        .LabelSize = labelSize
    }

                    ' Store chart binding so it can be updated even without a source TextBox
                    Dim yMinStr As String = If(yMin.HasValue, yMin.Value.ToString(Globalization.CultureInfo.InvariantCulture), "")
                    Dim yMaxStr As String = If(yMax.HasValue, yMax.Value.ToString(Globalization.CultureInfo.InvariantCulture), "")
                    Dim autoStr As String = If(autoScaleY, "1", "0")

                    ch.Tag = "CHART|" & resultTarget & "|" & yMinStr & "|" & yMaxStr & "|" &
             xStep.ToString(Globalization.CultureInfo.InvariantCulture) & "|" &
             maxPoints.ToString(Globalization.CultureInfo.InvariantCulture) & "|" &
             autoStr

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
                    Dim fontSize As Single = 8.25F

                    Dim colLeftDefault As Integer = 0
                    Dim colRightDefault As Integer = 0

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

                    ' ---- parse tail tokens (x= y= w= h= format= fontsize= colleftpos= colrightpos=) ----
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
                            Case "fontsize"
                                Dim fsInt As Integer
                                If Integer.TryParse(val, fsInt) Then
                                    fontSize = CSng(fsInt)
                                End If
                            Case "colleftpos"
                                Integer.TryParse(val, colLeftDefault)
                            Case "colrightpos"
                                Integer.TryParse(val, colRightDefault)
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
                    AddToUserContainer(gb)

                    ' Per-panel settings
                    StatsFontSize(panelName) = fontSize
                    StatsColLeft(panelName) = colLeftDefault
                    StatsColRight(panelName) = colRightDefault

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
                    Dim fmtOverride As String = ""   ' "F" or "E" or "G" prefix override (uses panel fmt digits)

                    Dim colLeftAdjust As Integer = 0
                    Dim colRightAdjust As Integer = 0

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
                    Dim startIdx As Integer = 4
                    If parts.Length < 4 Then startIdx = parts.Length

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

                        ElseIf p.StartsWith("colleftpos=", StringComparison.OrdinalIgnoreCase) Then
                            Integer.TryParse(p.Substring(11).Trim(), colLeftAdjust)

                        ElseIf p.StartsWith("colrightpos=", StringComparison.OrdinalIgnoreCase) Then
                            Integer.TryParse(p.Substring(12).Trim(), colRightAdjust)

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

                    ' font size + dynamic line height
                    Dim fs As Single = 8.25F
                    If StatsFontSize.ContainsKey(panelName) Then
                        fs = StatsFontSize(panelName)
                    End If

                    Dim lineHeight As Integer = CInt(Math.Ceiling(fs * 2.0F))
                    Dim rowY As Integer = 18 + (StatsRows(panelName).Count * lineHeight)

                    ' Panel-wide base column offsets
                    Dim baseLeft As Integer = 0
                    Dim baseRight As Integer = 0

                    If StatsColLeft.ContainsKey(panelName) Then baseLeft = StatsColLeft(panelName)
                    If StatsColRight.ContainsKey(panelName) Then baseRight = StatsColRight(panelName)

                    ' Effective per-row offsets = panel defaults + STAT overrides
                    Dim effectiveLeft As Integer = baseLeft + colLeftAdjust
                    Dim effectiveRight As Integer = baseRight + colRightAdjust

                    ' dynamic column spacing for value label
                    Dim nameWidth As Integer = CInt(fs * 9)
                    Dim valueX As Integer = nameWidth + 15 + effectiveRight

                    Dim lblName As New Label()
                    lblName.AutoSize = True
                    lblName.Text = labelText
                    lblName.Location = New Point(10 + effectiveLeft, rowY)
                    lblName.ForeColor = SystemColors.ControlText
                    lblName.Font = New Font(gb.Font.FontFamily, fs, lblName.Font.Style)

                    Dim lblVal As New Label()
                    lblVal.AutoSize = True
                    lblVal.Text = ""
                    lblVal.Location = New Point(valueX, rowY)
                    lblVal.ForeColor = SystemColors.ControlText
                    lblVal.Font = New Font(gb.Font.FontFamily, fs)

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
                    Dim popup As Boolean = False

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
                                Case "popup"
                                    If val <> "" Then
                                        Dim lv = val.Trim().ToLowerInvariant()
                                        popup = (lv = "1" OrElse lv = "true" OrElse lv = "yes")
                                    End If
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

                                Case "popup"
                                    If val <> "" Then
                                        Dim lv = val.Trim().ToLowerInvariant()
                                        popup = (lv = "1" OrElse lv = "true" OrElse lv = "yes")
                                    End If
                            End Select
                        Next

                        ' must have these in named mode
                        If gridName = "" OrElse resultTarget = "" Then Continue For
                    End If

                    ' ---------- Caption label (optional) ----------
                    Dim topY As Integer = y

                    If caption <> "" AndAlso Not popup Then
                        Dim lbl As New Label()
                        lbl.Text = caption
                        lbl.AutoSize = True
                        lbl.Location = New Point(x, y)
                        AddToUserContainer(lbl)
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

                    AddToUserContainer(dgv)

                    ' Keep a reference to the grid for popup logic (CLEARHISTORY)
                    HistoryGrids(gridName) = dgv

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

                    ' Register grid -> target so DATASOURCE updates can drive it (no hidden TextBox needed)
                    HistorySettings(gridName) = New HistoryGridConfig With {
                    .GridName = gridName,
                    .ResultTarget = resultTarget,
                    .MaxRows = maxRows,
                    .Format = fmt,
                    .Columns = colNames,
                    .PpmRef = ppmRef,
                    .Popup = popup,
                    .Width = w,
                    .Height = h
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

                            Case "name"
                                trigName = val

                            Case "if", "when"
                                ifExpr = NormalizeTriggerWhen(val)

                            Case "then", "do"
                                thenChain = val

                        End Select
                    Next

                    If String.IsNullOrWhiteSpace(trigName) Then
                        TriggerAutoIndex += 1
                        trigName = "T" & TriggerAutoIndex.ToString()
                    End If

                    If String.IsNullOrWhiteSpace(ifExpr) Then Continue For
                    If String.IsNullOrWhiteSpace(thenChain) Then Continue For

                    If needHits < 1 Then needHits = 1
                    If periodSec <= 0 Then periodSec = 0.5

                    ' Trigger engine must exist (or be created here)
                    EnsureTriggerEngine()

                    ' TriggerDef class/struct used by your engine
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
                    Dim overloadToken As String = ""
                    Dim decimalFlag As Boolean = True

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
                            Case "overload"
                                overloadToken = val
                            Case "decimal"
                                decimalFlag =
            (val.Trim() = "1" OrElse
             val.Equals("true", StringComparison.OrdinalIgnoreCase) OrElse
             val.Equals("yes", StringComparison.OrdinalIgnoreCase))
                        End Select
                    Next

                    ' === Validate result name for internal variable use ===
                    If Not String.IsNullOrWhiteSpace(resultName) Then
                        Dim c As Char = resultName(0)
                        ' Must start with a letter or underscore
                        If Not (Char.IsLetter(c) OrElse c = "_"c) Then
                            ShowConfigWarningPopup(
    $"Invalid DATASOURCE result name '{resultName}'." & vbCrLf & vbCrLf &
    "Result names must start with a letter or underscore " &
    "for internal variables (e.g. CALC / LUA / trigger)."
)
                        End If
                    End If

                    If resultName <> "" AndAlso device <> "" AndAlso command <> "" Then
                        DataSources(resultName) = New DataSourceDef With {
                        .Device = device,
                        .Command = command,
                        .OverloadToken = overloadToken,
                        .ForceDecimal = decimalFlag
                        }
                    End If


                Case "CALC"
                    ParseCalcLine(parts)
                    Continue For


                Case "MULTIBUTTON"
                    Dim name As String = ""
                    Dim caption As String = ""
                    Dim device As String = ""
                    Dim commandPrefix As String = ""
                    Dim itemsRaw As String = ""
                    Dim commandsRaw As String = ""
                    Dim determineRaw As String = ""
                    Dim detmapRaw As String = ""

                    Dim x As Integer = 20, y As Integer = 20
                    Dim w As Integer = 60     ' per-button width
                    Dim h As Integer = 28     ' per-button height
                    Dim gap As Integer = 0

                    Dim onColorName As String = "limegreen"
                    Dim offColorName As String = "default"

                    ' --- tokens ---
                    Dim tok As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    For i As Integer = 1 To parts.Length - 1
                        Dim t = parts(i).Trim()
                        If t = "" Then Continue For

                        Dim eq = t.IndexOf("="c)
                        If eq > 0 AndAlso eq < t.Length - 1 Then
                            Dim k = t.Substring(0, eq).Trim()
                            Dim v = t.Substring(eq + 1).Trim().TrimEnd(";"c)
                            If k <> "" Then tok(k) = v
                        End If
                    Next

                    If tok.ContainsKey("name") Then name = tok("name")
                    If tok.ContainsKey("caption") Then caption = tok("caption")
                    If tok.ContainsKey("device") Then device = tok("device")

                    If tok.ContainsKey("command") Then commandPrefix = tok("command")
                    If tok.ContainsKey("cmd") AndAlso commandPrefix = "" Then commandPrefix = tok("cmd")

                    ' ---- multiline-aware for items / commands / detmap ----
                    itemsRaw = ParseMultiValueField(parts, "items")
                    'If itemsRaw = "" AndAlso tok.ContainsKey("items") Then
                    'itemsRaw = tok("items")
                    'End If

                    commandsRaw = ParseMultiValueField(parts, "commands")
                    'If commandsRaw = "" AndAlso tok.ContainsKey("commands") Then
                    'commandsRaw = tok("commands")
                    'End If

                    detmapRaw = ParseMultiValueField(parts, "detmap")
                    'If detmapRaw = "" AndAlso tok.ContainsKey("detmap") Then
                    'detmapRaw = tok("detmap")
                    'End If


                    'Debug.WriteLine("==== MULTIBUTTON DEBUG =====")
                    'Debug.WriteLine("NAME: " & name)

                    'Debug.WriteLine("ITEMS RAW:")
                    'Debug.WriteLine(If(itemsRaw, "<NULL>"))

                    'Debug.WriteLine("COMMANDS RAW:")
                    'Debug.WriteLine(If(commandsRaw, "<NULL>"))

                    'Debug.WriteLine("DETMAP RAW:")
                    'Debug.WriteLine(If(detmapRaw, "<NULL>"))


                    If tok.ContainsKey("x") Then ParseIntField(tok("x"), x)
                    If tok.ContainsKey("y") Then ParseIntField(tok("y"), y)
                    If tok.ContainsKey("w") Then ParseIntField(tok("w"), w)
                    If tok.ContainsKey("h") Then ParseIntField(tok("h"), h)
                    If tok.ContainsKey("gap") Then ParseIntField(tok("gap"), gap)

                    If tok.ContainsKey("determine") Then determineRaw = tok("determine")

                    If tok.ContainsKey("oncolor") Then onColorName = tok("oncolor").ToLowerInvariant()
                    If tok.ContainsKey("offcolor") Then offColorName = tok("offcolor").ToLowerInvariant()

                    If name = "" OrElse device = "" OrElse itemsRaw = "" Then Continue For

                    ' ====== build arrays first ======
                    Dim items = itemsRaw.Split(","c).
                        Select(Function(s) s.Trim()).
                        Where(Function(s) s <> "").
                        ToArray()
                    If items.Length < 2 OrElse items.Length > 10 Then Continue For

                    Dim cmds() As String = Nothing
                    If commandsRaw <> "" Then
                        cmds = commandsRaw.Split(","c).
                            Select(Function(s) s.Trim()).
                            Where(Function(s) s <> "").
                            ToArray()
                        If cmds.Length <> items.Length Then Continue For
                    Else
                        If commandPrefix = "" Then Continue For
                    End If

                    Dim detmap() As String = Nothing
                    If detmapRaw <> "" Then
                        detmap = detmapRaw.Split(","c).
                            Select(Function(s) s.Trim()).
                            Where(Function(s) s <> "").
                            ToArray()
                        If detmap.Length <> items.Length Then Continue For
                    End If


                    'Debug.WriteLine($"MB {name}: items={items.Length}, cmds={If(cmds Is Nothing, -1, cmds.Length)}, detmap={If(detmap Is Nothing, -1, detmap.Length)}")


                    ' --- caption label ---
                    Dim capW As Integer = 0
                    If caption <> "" Then
                        Dim lbl As New Label()
                        lbl.Text = caption
                        lbl.AutoSize = True
                        lbl.Location = New Point(x, y + 6)
                        AddToUserContainer(lbl)
                        capW = lbl.PreferredWidth + 8
                    End If

                    ' --- host panel ---
                    Dim pnl As New Panel()
                    pnl.Name = "MB_" & name
                    pnl.Location = New Point(x + capW, y)
                    pnl.Size = New Size(items.Length * w + (items.Length - 1) * gap, h)
                    AddToUserContainer(pnl)

                    ' --- base tag ---
                    Dim tagBase As String =
                        "MB|" & device & "|" & commandPrefix & "|" & pnl.Name & "|" &
                        items.Length & "|-1" &
                        "|ONCLR=" & onColorName & "|OFFCLR=" & offColorName

                    ' determine=<query>||resptext
                    If determineRaw <> "" Then
                        Dim dp() = determineRaw.Split("|"c)
                        Dim detQ As String = If(dp.Length >= 1, dp(0).Trim(), "")
                        Dim detE As String = If(dp.Length >= 2, dp(1).Trim(), "")
                        Dim detF As String = If(dp.Length >= 3, dp(2).Trim().ToLowerInvariant(), "")

                        If detQ <> "" Then
                            tagBase &= "|DETQ=" & detQ
                            If detE <> "" Then tagBase &= "|DETE=" & detE
                            If detF <> "" Then tagBase &= "|DETF=" & detF
                        End If
                    End If

                    If cmds IsNot Nothing Then
                        tagBase &= "|CMDS=" & String.Join("§", cmds)
                    End If

                    If detmap IsNot Nothing Then
                        tagBase &= "|DETMAP=" & String.Join("§", detmap)
                    End If

                    ' --- buttons ---
                    Dim xx As Integer = 0
                    For i As Integer = 0 To items.Length - 1
                        Dim b As New Button()
                        b.Name = "MBI_" & name & "_" & i
                        b.Text = items(i)
                        b.UseVisualStyleBackColor = False
                        b.Location = New Point(xx, 0)
                        b.Size = New Size(w, h)
                        b.Tag = tagBase & "|IDX=" & i & "|ITEM=" & items(i)

                        AddHandler b.Click, AddressOf MultiButton_Click
                        pnl.Controls.Add(b)

                        xx += w + gap
                    Next

                    ApplyMultiButtonVisual(pnl, -1)
                    Continue For



                Case "KEYPAD"
                    Dim name As String = "Keypad1"
                    Dim caption As String = "Keypad"
                    Dim mode As String = "fixed"
                    Dim x As Integer = 20, y As Integer = 20, w As Integer = 230, h As Integer = 280
                    Dim targetLabelOn As Boolean = True

                    Dim tok As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    For i As Integer = 1 To parts.Length - 1
                        Dim t = parts(i).Trim()
                        Dim eq = t.IndexOf("="c)
                        If eq > 0 AndAlso eq < t.Length - 1 Then
                            tok(t.Substring(0, eq).Trim()) = t.Substring(eq + 1).Trim()
                        End If
                    Next

                    If tok.ContainsKey("name") Then name = tok("name").Trim()
                    If tok.ContainsKey("caption") Then caption = tok("caption").Trim()
                    If tok.ContainsKey("mode") Then mode = tok("mode").Trim().ToLowerInvariant()

                    If tok.ContainsKey("x") Then ParseIntField(tok("x"), x)
                    If tok.ContainsKey("y") Then ParseIntField(tok("y"), y)
                    If tok.ContainsKey("w") Then ParseIntField(tok("w"), w)
                    If tok.ContainsKey("h") Then ParseIntField(tok("h"), h)

                    If tok.ContainsKey("targetlabel") Then
                        Dim v = tok("targetlabel").Trim().ToLowerInvariant()
                        targetLabelOn = (v = "on" OrElse v = "true" OrElse v = "1")
                    End If

                    ' Remember current keypad config so popup can be recreated if disposed
                    UserKeypadName = name
                    UserKeypadCaption = caption
                    UserKeypadW = w
                    UserKeypadH = h
                    UserKeypadTargetLabelOn = targetLabelOn

                    UserKeypadMode = If(mode = "popup", "popup", "fixed")

                    If UserKeypadMode = "fixed" Then

                        UserKeypadFixedPanel = CreateUserKeypadPanel(name, caption, w, h, targetLabelOn)
                        UserKeypadFixedPanel.Location = New Point(x, y)
                        AddToUserContainer(UserKeypadFixedPanel)

                    Else
                        ' Recreate-safe: EnsureKeypadPopup should internally handle Nothing/IsDisposed.
                        EnsureKeypadPopup(name, caption, w, h, targetLabelOn)

                        UserKeypadPopupForm.StartPosition = FormStartPosition.Manual

                        If Not UserKeypadPopupPosValid Then
                            UserKeypadPopupForm.Location = Me.PointToScreen(New Point(x, y))
                            UserKeypadPopupPos = UserKeypadPopupForm.Location
                            UserKeypadPopupPosValid = True
                        Else
                            UserKeypadPopupForm.Location = UserKeypadPopupPos
                        End If

                        UserKeypadPopupForm.Hide()

                    End If
                    Continue For


                Case "DATASAVE"
                    If parts.Length >= 2 Then
                        Dim v As String = parts(1).Trim().ToLowerInvariant()
                        UserConfig_DataSaveEnabled = (v = "enabled" OrElse v = "true" OrElse v = "1")
                    End If
                    Continue For





            End Select

        Next

        ' Apply INVISIBILITY defaults after all controls are created/registered.
        ApplyInvisibilityDefaults()

        ' Apply initial instrument state to controls that have determine=...
        ApplyDetermineRadios()
        ApplyDetermineDropdowns()
        ApplyDetermineSliders()
        ApplyDetermineSpinners()
        ApplyDetermineToggles()
        ApplyDetermineToggleDuals()
        ApplyDetermineMultiButtons()

    End Sub


    ' =========================================================
    '   DETERMINE (initial state readback)
    ' =========================================================
    '
    ' RADIO lines can optionally include:
    '   determine=<query>|<expected>
    '   determine=<query>|<expected>|resptext
    '
    ' If resptext is specified, WinGPIB will request a raw response for that query
    ' (USERdev1rawoutput/USERdev2rawoutput TRUE) and then do a text "contains" match.
    ' If omitted, numeric compare (respnum) is assumed (extract first number and compare).
    '
    Private Sub ApplyDetermineRadios()

        Dim cache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        For Each ctrl As Control In GetUserDetermineControls()
            Dim gb As GroupBox = TryCast(ctrl, GroupBox)
            If gb Is Nothing Then Continue For
            If Not gb.Name.StartsWith("RG_", StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim chosen As RadioButton = Nothing

            For Each child As Control In gb.Controls
                Dim rb As RadioButton = TryCast(child, RadioButton)
                If rb Is Nothing Then Continue For

                Dim tagStr As String = TryCast(rb.Tag, String)
                If String.IsNullOrEmpty(tagStr) Then Continue For

                ' Tag is: device|command|scale OR device|command|AUTO|rangeQuery|DETQ=...|DETE=...|DETF=...
                Dim parts() As String = tagStr.Split("|"c)
                If parts.Length < 2 Then Continue For

                Dim deviceName As String = parts(0).Trim()

                ' Look for DETQ= / DETE= / DETF=
                Dim detQuery As String = ""
                Dim detExpected As String = ""
                Dim detFmt As String = ""     ' optional

                For Each p In parts
                    If p.StartsWith("DETQ=", StringComparison.OrdinalIgnoreCase) Then
                        detQuery = p.Substring(5).Trim()
                    ElseIf p.StartsWith("DETE=", StringComparison.OrdinalIgnoreCase) Then
                        detExpected = p.Substring(5).Trim()
                    ElseIf p.StartsWith("DETF=", StringComparison.OrdinalIgnoreCase) Then
                        detFmt = p.Substring(5).Trim().ToLowerInvariant()
                    End If
                Next

                If detQuery = "" OrElse detExpected = "" Then Continue For

                If detFmt = "" Then detFmt = "respnum" ' default

                Dim cacheKey As String = deviceName & "||" & detQuery & "||" & detFmt
                Dim reply As String = ""

                If Not cache.TryGetValue(cacheKey, reply) Then
                    reply = DetermineQuery(deviceName, detQuery, detFmt)
                    cache(cacheKey) = reply
                End If

                If DetermineMatch(reply, detExpected, detFmt) Then
                    chosen = rb
                    Exit For
                End If

            Next

            If chosen IsNot Nothing Then
                ' Important: do NOT send SCPI when we tick the radio during determine
                UserInitSuppressSend = True
                SyncUserUnitsAndScaleFromCheckedRadios()
                chosen.Checked = True

                ' --- update CurrentUserUnit when determine selects a range radio ---
                Dim u As String = ExtractUnitFromCaption(chosen.Text)
                If Not String.IsNullOrEmpty(u) Then CurrentUserUnit = u

                UserInitSuppressSend = False
            End If
        Next

    End Sub


    Private Function DetermineMatch(reply As String, expected As String, detFmt As String) As Boolean
        If reply Is Nothing Then reply = ""
        If expected Is Nothing Then expected = ""

        If detFmt = "resptext" Then
            Dim r As String = reply.ToUpperInvariant()
            Dim eRaw As String = expected.Trim()

            If eRaw = "" Then Return False

            ' Exact token match if expected starts with "="
            If eRaw.StartsWith("=", StringComparison.Ordinal) Then
                Dim tok As String = eRaw.Substring(1).Trim().ToUpperInvariant()
                If tok = "" Then Return False

                ' Make tokenization robust (strip quotes, split on common SCPI delimiters)
                r = r.Replace("""", " ")

                Dim seps() As Char = {":"c, ","c, ";"c, " "c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf}
                For Each t In r.Split(seps, StringSplitOptions.RemoveEmptyEntries)
                    If t.Trim().ToUpperInvariant() = tok Then Return True
                Next
                Return False
            End If

            ' Default (legacy): substring match
            Dim e As String = eRaw.ToUpperInvariant()
            Return r.Contains(e)
        Else
            ' respnum (default)
            Dim dv As Double
            If Not TryExtractFirstDouble(reply, dv) Then Return False

            Dim ev As Double
            If Not Double.TryParse(expected, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, ev) Then
                Return False
            End If

            Return Math.Abs(dv - ev) < 0.0000001
        End If

    End Function


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
            If s = "" Then Continue For

            ' Allow optional trailing semicolon: caption=Foo;  → caption=Foo
            s = s.TrimEnd(";"c)

            If s.Contains("="c) Then
                Dim kv = s.Split({"="c}, 2)
                Dim key = kv(0).Trim()
                Dim val = kv(1).Trim().TrimEnd(";"c)   ' extra safety if value itself ends with ;

                If key <> "" Then
                    dict(key) = val
                End If
            End If
        Next

        Return dict
    End Function


    ' Which grids listen to which result name
    Public Class DataSourceDef
        Public Device As String
        Public Command As String
        Public OverloadToken As String
        Public Property ForceDecimal As Boolean  ' True = decimal, False = raw/sci
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
        ' Optional format suffix: expr=<math>|<fmt>   e.g. expr=HP3245ADatasource*1|F6

        If parts Is Nothing OrElse parts.Length < 2 Then Exit Sub

        Dim outResult As String = ""
        Dim expr As String = ""
        Dim outFmt As String = ""   ' NEW

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

        ' --- NEW: split expr into math + optional format suffix "math|F6" ---
        Dim p As Integer = expr.LastIndexOf("|"c)
        If p > 0 AndAlso p < expr.Length - 1 Then
            outFmt = expr.Substring(p + 1).Trim()
            expr = expr.Substring(0, p).Trim()
        End If

        ' Build dependencies from the math expression only
        Dim deps As List(Of String) = ExtractCalcDeps(expr)

        ' Store
        Dim cd As New CalcDef With {
        .OutResult = outResult.Trim(),
        .Expr = expr,
        .Deps = deps,
        .OutFmt = outFmt          ' NEW (requires CalcDef.OutFmt As String)
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


    Private Function NormalizeTriggerWhen(expr As String) As String
        If String.IsNullOrWhiteSpace(expr) Then Return ""

        Dim s = expr.Trim()

        ' Convert ON/OFF forms: LedMAV=ON  -> LedMAV==1
        ' Convert single '=' to '==' only when it's a comparison (not >=, <=, !=, ==)
        ' This is deliberately simple and safe.
        If s.Contains("="c) AndAlso Not s.Contains(">=") AndAlso Not s.Contains("<=") AndAlso Not s.Contains("==") AndAlso Not s.Contains("!=") Then
            ' change first '=' to '=='
            Dim p = s.Split({"="c}, 2)
            If p.Length = 2 Then
                Dim lhs = p(0).Trim()
                Dim rhs = p(1).Trim()

                ' Map ON/OFF/TRUE/FALSE to 1/0
                Select Case rhs.ToUpperInvariant()
                    Case "ON", "TRUE", "HIGH", "YES"
                        rhs = "1"
                    Case "OFF", "FALSE", "LOW", "NO"
                        rhs = "0"
                End Select

                s = lhs & " == " & rhs
            End If
        End If

        Return s
    End Function


    Private Sub ApplyDetermineDropdowns()

        ' test only
        '_detDbg.Clear()                ' TESTING ONLY

        For Each ctrl As Control In GetUserDetermineControls()

            Dim cb As ComboBox = TryCast(ctrl, ComboBox)
            If cb Is Nothing Then Continue For

            Dim tagStr As String = TryCast(cb.Tag, String)
            If String.IsNullOrEmpty(tagStr) Then Continue For

            Dim parts() As String = tagStr.Split("|"c)
            If parts.Length < 2 Then Continue For

            Dim deviceName As String = parts(0).Trim()

            Dim detQuery As String = ""
            Dim detFmt As String = "respnum"
            Dim detMap As String = ""

            For Each p In parts
                If p.StartsWith("DETQ=", StringComparison.OrdinalIgnoreCase) Then
                    detQuery = p.Substring(5).Trim()

                ElseIf p.StartsWith("DETF=", StringComparison.OrdinalIgnoreCase) Then
                    detFmt = p.Substring(5).Trim().ToLowerInvariant()

                ElseIf p.StartsWith("DETMAP=", StringComparison.OrdinalIgnoreCase) Then
                    detMap = p.Substring(7).Trim()
                End If
            Next

            If detQuery = "" Then Continue For

            Dim reply As String = DetermineQuery(deviceName, detQuery, detFmt)

            If reply = "" Then Continue For

            UserInitSuppressSend = True

            If detFmt = "resptext" AndAlso detMap <> "" Then
                ' ---------------------------
                ' TEXT + DETMAP MODE
                ' ---------------------------
                Dim maps() As String =
                detMap.Split(","c).
                Select(Function(s) s.Trim()).
                Where(Function(s) s <> "").
                ToArray()

                ' Extract token from reply: first token before whitespace
                ' Examples:
                '   "VOLT:AC +1.00000000E+01,+1.00000000E-05" -> "VOLT:AC"
                Dim rep As String = reply.Trim().Trim(""""c)

                ' first token up to first space
                Dim funcToken As String = rep
                Dim spIdx As Integer = rep.IndexOf(" "c)
                If spIdx > 0 Then funcToken = rep.Substring(0, spIdx).Trim()

                Dim funcU As String = funcToken.ToUpperInvariant()

                Dim matchIndex As Integer = -1

                For i As Integer = 0 To maps.Length - 1
                    Dim m As String = maps(i)
                    If m = "" Then Continue For

                    Dim exact As Boolean = False
                    If m.StartsWith("="c) Then
                        exact = True
                        m = m.Substring(1)
                    End If

                    Dim mU As String = m.ToUpperInvariant()

                    If exact Then
                        If String.Equals(funcU, mU, StringComparison.OrdinalIgnoreCase) Then
                            matchIndex = i
                            Exit For
                        End If
                    Else
                        ' starts-with match (order matters: put VOLT:AC before VOLT)
                        If funcU.StartsWith(mU, StringComparison.OrdinalIgnoreCase) Then
                            matchIndex = i
                            Exit For
                        End If
                    End If
                Next

                If matchIndex >= 0 Then
                    Dim sel As Integer = matchIndex + 1 ' +1 because index 0 is placeholder
                    If sel >= 0 AndAlso sel < cb.Items.Count Then cb.SelectedIndex = sel
                End If

            Else
                ' ---------------------------
                ' NUMERIC MODE (existing)
                ' ---------------------------
                Dim num As Double
                If TryExtractFirstDouble(reply, num) Then
                    For i As Integer = 0 To cb.Items.Count - 1
                        Dim itemVal As Double
                        If Double.TryParse(cb.Items(i).ToString(),
                                       Globalization.NumberStyles.Float,
                                       Globalization.CultureInfo.InvariantCulture,
                                       itemVal) Then
                            If Math.Abs(itemVal - num) < 0.0000001 Then
                                cb.SelectedIndex = i
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If

            UserInitSuppressSend = False

        Next

        'ShowCopyableDebug("ApplyDetermineDropdowns()", _detDbg.ToString())             ' TESTING ONLY

    End Sub


    Private Sub ApplyDetermineSliders()

        For Each ctrl As Control In GetUserDetermineControls()

            ' Your sliders live inside a GroupBox you create per slider
            Dim gb As GroupBox = TryCast(ctrl, GroupBox)
            If gb Is Nothing Then Continue For
            If Not gb.Name.StartsWith("GB_Slider_", StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim tb As TrackBar = Nothing
            For Each c As Control In gb.Controls
                tb = TryCast(c, TrackBar)
                If tb IsNot Nothing Then Exit For
            Next
            If tb Is Nothing Then Continue For

            Dim tagStr As String = TryCast(tb.Tag, String)
            If String.IsNullOrEmpty(tagStr) Then Continue For

            Dim parts() As String = tagStr.Split("|"c)
            If parts.Length < 2 Then Continue For

            Dim deviceName As String = parts(0).Trim()

            Dim detQuery As String = ""
            Dim detFmt As String = "respnum"

            For Each p In parts
                If p.StartsWith("DETQ=", StringComparison.OrdinalIgnoreCase) Then
                    detQuery = p.Substring(5).Trim()
                ElseIf p.StartsWith("DETF=", StringComparison.OrdinalIgnoreCase) Then
                    detFmt = p.Substring(5).Trim().ToLowerInvariant()
                End If
            Next

            If detQuery = "" Then Continue For

            Dim reply As String = DetermineQuery(deviceName, detQuery, detFmt)
            If reply = "" Then Continue For

            Dim num As Double
            If Not TryExtractFirstDouble(reply, num) Then Continue For

            Dim newVal As Integer = CInt(Math.Round(num))
            If newVal < tb.Minimum Then newVal = tb.Minimum
            If newVal > tb.Maximum Then newVal = tb.Maximum

            UserInitSuppressSend = True
            tb.Value = newVal
            UserInitSuppressSend = False

            ' Update the slider value label (normally done by Scroll/MouseUp, but suppressed during determine)
            Try
                Dim tagParts() As String = tagStr.Split("|"c)
                If tagParts.Length >= 5 Then
                    Dim scale As Double = 1.0
                    Double.TryParse(tagParts(2), Globalization.NumberStyles.Float,
                        Globalization.CultureInfo.InvariantCulture, scale)

                    Dim lblName As String = tagParts(3).Trim()

                    Dim lbl As Label = Nothing
                    For Each c As Control In gb.Controls
                        If TypeOf c Is Label AndAlso String.Equals(c.Name, lblName, StringComparison.OrdinalIgnoreCase) Then
                            lbl = DirectCast(c, Label)
                            Exit For
                        End If
                    Next

                    If lbl IsNot Nothing Then
                        Dim disp As Double = newVal * scale
                        lbl.Text = disp.ToString(Globalization.CultureInfo.InvariantCulture)
                    End If
                End If
            Catch
                ' ignore label update failures
            End Try

        Next

    End Sub


    Private Sub ApplyDetermineSpinners()

        For Each ctrl As Control In GetUserDetermineControls()

            Dim nud As NumericUpDown = TryCast(ctrl, NumericUpDown)
            If nud Is Nothing Then Continue For

            Dim tagStr As String = TryCast(nud.Tag, String)
            If String.IsNullOrEmpty(tagStr) Then Continue For

            Dim parts() As String = tagStr.Split("|"c)
            If parts.Length < 5 Then Continue For

            Dim deviceName As String = parts(0).Trim()

            ' Tag = DEVICE|COMMAND|SCALE|INITFLAG|MIN|DETQ=...|DETF=...
            Dim scale As Double = 1.0
            Double.TryParse(parts(2), Globalization.NumberStyles.Float,
                        Globalization.CultureInfo.InvariantCulture, scale)

            Dim detQuery As String = ""
            Dim detFmt As String = "respnum"

            For Each p In parts
                If p.StartsWith("DETQ=", StringComparison.OrdinalIgnoreCase) Then
                    detQuery = p.Substring(5).Trim()
                ElseIf p.StartsWith("DETF=", StringComparison.OrdinalIgnoreCase) Then
                    detFmt = p.Substring(5).Trim().ToLowerInvariant()
                End If
            Next

            If detQuery = "" Then Continue For

            Dim reply As String = DetermineQuery(deviceName, detQuery, detFmt)
            If reply = "" Then Continue For

            Dim num As Double
            If Not TryExtractFirstDouble(reply, num) Then Continue For

            ' Convert instrument value -> spinner internal value
            ' (assumes send uses: cmd & (nud.Value * scale))
            Dim internalVal As Double = num
            If scale <> 0 Then internalVal = num / scale

            Dim newVal As Decimal = CDec(Math.Round(internalVal, 0))

            If newVal < nud.Minimum Then newVal = nud.Minimum
            If newVal > nud.Maximum Then newVal = nud.Maximum

            UserInitSuppressSend = True
            nud.Value = newVal
            nud.Text = newVal.ToString(Globalization.CultureInfo.InvariantCulture)  ' prevent blank display
            UserInitSuppressSend = False

        Next

    End Sub


    Private Sub ApplyDetermineToggles()

        Dim cache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        For Each ctrl As Control In GetUserDetermineControls()

            Dim b As Button = TryCast(ctrl, Button)
            If b Is Nothing Then Continue For

            Dim tagStr As String = TryCast(b.Tag, String)
            If String.IsNullOrEmpty(tagStr) Then Continue For

            Dim parts() As String = tagStr.Split("|"c)
            If parts.Length < 4 Then Continue For   ' DEVICE|ONCMD|OFFCMD|STATE...

            Dim deviceName As String = parts(0).Trim()

            Dim detQuery As String = ""
            Dim detExpected As String = ""
            Dim detFmt As String = "respnum"

            Dim onClr As String = "default"
            Dim offClr As String = "default"

            For Each p In parts
                If p.StartsWith("DETQ=", StringComparison.OrdinalIgnoreCase) Then
                    detQuery = p.Substring(5).Trim()
                ElseIf p.StartsWith("DETE=", StringComparison.OrdinalIgnoreCase) Then
                    detExpected = p.Substring(5).Trim()
                ElseIf p.StartsWith("DETF=", StringComparison.OrdinalIgnoreCase) Then
                    detFmt = p.Substring(5).Trim().ToLowerInvariant()
                ElseIf p.StartsWith("ONCLR=", StringComparison.OrdinalIgnoreCase) Then
                    onClr = p.Substring(6).Trim().ToLowerInvariant()
                ElseIf p.StartsWith("OFFCLR=", StringComparison.OrdinalIgnoreCase) Then
                    offClr = p.Substring(7).Trim().ToLowerInvariant()
                End If
            Next

            If detQuery = "" OrElse detExpected = "" Then Continue For

            Dim cacheKey As String = deviceName & "||" & detQuery & "||" & detFmt
            Dim reply As String = ""
            If Not cache.TryGetValue(cacheKey, reply) Then
                reply = DetermineQuery(deviceName, detQuery, detFmt)
                cache(cacheKey) = reply
            End If

            Dim isOn As Boolean = DetermineMatch(reply, detExpected, detFmt)
            Dim newState As Integer = If(isOn, 1, 0)

            ' Update Tag state but preserve DET tokens
            parts(3) = newState.ToString(Globalization.CultureInfo.InvariantCulture)
            b.Tag = String.Join("|", parts)

            ApplyToggleVisual(b, newState, onClr, offClr)

        Next

    End Sub


    Private Function ToggleColorFromName(name As String) As Color
        Select Case name.ToLowerInvariant()
            Case "limegreen" : Return Color.LimeGreen
            Case "green" : Return Color.Green
            Case "red" : Return Color.Red
            Case "orange" : Return Color.Orange
            Case "yellow" : Return Color.Gold
            Case "blue" : Return Color.DeepSkyBlue
            Case Else : Return SystemColors.Control
        End Select
    End Function



    Private Sub ApplyToggleVisual(b As Button, state As Integer,
                             Optional onClr As String = "default",
                             Optional offClr As String = "default")

        b.UseVisualStyleBackColor = False

        If state = 1 Then
            b.BackColor = ToggleColorFromName(onClr)
            b.ForeColor = Color.Black
        Else
            b.BackColor = ToggleColorFromName(offClr)
            b.ForeColor = SystemColors.ControlText
        End If

    End Sub


    Private Sub ToggleDual_LeftClick(sender As Object, e As EventArgs)
        If UserInitSuppressSend Then Exit Sub
        ToggleDual_SendAndSet(DirectCast(sender, Button), True)
    End Sub


    Private Sub ToggleDual_RightClick(sender As Object, e As EventArgs)
        If UserInitSuppressSend Then Exit Sub
        ToggleDual_SendAndSet(DirectCast(sender, Button), False)
    End Sub


    Private Sub ToggleDual_SendAndSet(b As Button, setLeftActive As Boolean)

        Dim tagStr As String = TryCast(b.Tag, String)
        If String.IsNullOrEmpty(tagStr) Then Exit Sub

        Dim parts() As String = tagStr.Split("|"c)
        If parts.Length < 5 Then Exit Sub

        Dim device As String = parts(0).Trim()
        Dim cmdOn As String = parts(1).Trim()
        Dim cmdOff As String = parts(2).Trim()
        Dim pnlName As String = parts(4).Trim()

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

        Dim cmd As String = If(setLeftActive, cmdOn, cmdOff)

        If useNative Then
            NativeSend(device, cmd & TermStr2())
        Else
            dev.SendAsync(cmd & TermStr2(), True)
        End If

        ' Update STATE in BOTH buttons' tag (preserve DET tokens)
        parts(3) = If(setLeftActive, "1", "0")
        Dim newTag As String = String.Join("|", parts)

        Dim found() As Control = GroupBoxCustom.Controls.Find(pnlName, True)
        If found IsNot Nothing AndAlso found.Length > 0 Then
            Dim pnl As Panel = TryCast(found(0), Panel)
            If pnl IsNot Nothing Then
                For Each c As Control In pnl.Controls
                    Dim bb As Button = TryCast(c, Button)
                    If bb IsNot Nothing Then bb.Tag = newTag
                Next
                ApplyToggleDualVisual(pnl, setLeftActive)
            End If
        End If

    End Sub


    Private Sub ApplyToggleDualVisual(pnl As Panel, leftActive As Boolean)

        Dim bL As Button = Nothing
        Dim bR As Button = Nothing

        For Each c As Control In pnl.Controls
            Dim bb As Button = TryCast(c, Button)
            If bb Is Nothing Then Continue For
            If bb.Name.StartsWith("TDL_", StringComparison.OrdinalIgnoreCase) Then bL = bb
            If bb.Name.StartsWith("TDR_", StringComparison.OrdinalIgnoreCase) Then bR = bb
        Next
        If bL Is Nothing OrElse bR Is Nothing Then Exit Sub

        Dim tagStr As String = TryCast(bL.Tag, String)
        Dim onClr As String = "default"
        Dim offClr As String = "default"

        If Not String.IsNullOrEmpty(tagStr) Then
            Dim parts() As String = tagStr.Split("|"c)
            For Each p In parts
                If p.StartsWith("ONCLR=", StringComparison.OrdinalIgnoreCase) Then
                    onClr = p.Substring(6).Trim().ToLowerInvariant()
                ElseIf p.StartsWith("OFFCLR=", StringComparison.OrdinalIgnoreCase) Then
                    offClr = p.Substring(7).Trim().ToLowerInvariant()
                End If
            Next
        End If

        bL.UseVisualStyleBackColor = False
        bR.UseVisualStyleBackColor = False

        If leftActive Then
            bL.BackColor = ToggleColorFromName(onClr)
            bL.ForeColor = Color.Black
            bR.BackColor = ToggleColorFromName(offClr)
            bR.ForeColor = SystemColors.ControlText
        Else
            bR.BackColor = ToggleColorFromName(onClr)
            bR.ForeColor = Color.Black
            bL.BackColor = ToggleColorFromName(offClr)
            bL.ForeColor = SystemColors.ControlText
        End If

    End Sub


    Private Sub ApplyDetermineToggleDuals()

        Dim cache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        For Each ctrl As Control In GetUserDetermineControls()

            Dim pnl As Panel = TryCast(ctrl, Panel)
            If pnl Is Nothing Then Continue For
            If Not pnl.Name.StartsWith("TD_", StringComparison.OrdinalIgnoreCase) Then Continue For

            ' Grab tag from any button inside
            Dim tagStr As String = Nothing
            For Each c As Control In pnl.Controls
                Dim bb As Button = TryCast(c, Button)
                If bb IsNot Nothing Then
                    tagStr = TryCast(bb.Tag, String)
                    If Not String.IsNullOrEmpty(tagStr) Then Exit For
                End If
            Next
            If String.IsNullOrEmpty(tagStr) Then Continue For

            Dim parts() As String = tagStr.Split("|"c)
            If parts.Length < 5 Then Continue For

            Dim deviceName As String = parts(0).Trim()

            Dim detQuery As String = ""
            Dim detExpected As String = ""
            Dim detFmt As String = "respnum"

            For Each p In parts
                If p.StartsWith("DETQ=", StringComparison.OrdinalIgnoreCase) Then
                    detQuery = p.Substring(5).Trim()
                ElseIf p.StartsWith("DETE=", StringComparison.OrdinalIgnoreCase) Then
                    detExpected = p.Substring(5).Trim()
                ElseIf p.StartsWith("DETF=", StringComparison.OrdinalIgnoreCase) Then
                    detFmt = p.Substring(5).Trim().ToLowerInvariant()
                End If
            Next

            If detQuery = "" OrElse detExpected = "" Then Continue For

            Dim cacheKey As String = deviceName & "||" & detQuery & "||" & detFmt
            Dim reply As String = ""
            If Not cache.TryGetValue(cacheKey, reply) Then
                reply = DetermineQuery(deviceName, detQuery, detFmt)
                cache(cacheKey) = reply
            End If

            Dim isOnLeft As Boolean = DetermineMatch(reply, detExpected, detFmt)

            ' Update tag state + visual WITHOUT sending SCPI
            Dim newState As String = If(isOnLeft, "1", "0")
            parts(3) = newState
            Dim newTag As String = String.Join("|", parts)

            UserInitSuppressSend = True
            For Each c As Control In pnl.Controls
                Dim bb As Button = TryCast(c, Button)
                If bb IsNot Nothing Then bb.Tag = newTag
            Next
            ApplyToggleDualVisual(pnl, isOnLeft)
            UserInitSuppressSend = False

        Next

    End Sub


    Private Sub ApplyDetermineMultiButtons()

        Dim cache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        For Each ctrl As Control In GetUserDetermineControls()

            Dim pnl As Panel = TryCast(ctrl, Panel)
            If pnl Is Nothing Then Continue For
            If Not pnl.Name.StartsWith("MB_", StringComparison.OrdinalIgnoreCase) Then Continue For

            ' Grab tag from any button inside
            Dim tagStr As String = Nothing
            For Each c As Control In pnl.Controls
                Dim bb As Button = TryCast(c, Button)
                If bb IsNot Nothing Then
                    tagStr = TryCast(bb.Tag, String)
                    If Not String.IsNullOrEmpty(tagStr) Then Exit For
                End If
            Next
            If String.IsNullOrEmpty(tagStr) Then Continue For

            Dim parts() As String = tagStr.Split("|"c)

            Dim deviceName As String = ""
            Dim detQuery As String = ""
            Dim detFmt As String = "respnum"
            Dim detmapPacked As String = ""
            Dim onClr As String = "limegreen"
            Dim offClr As String = "default"

            ' MB|device|prefix|panel|count|state|ONCLR=..|OFFCLR=..|DETQ=..|DETF=..|DETMAP=..
            If parts.Length >= 2 Then deviceName = parts(1).Trim()

            For Each p In parts
                If p.StartsWith("DETQ=", StringComparison.OrdinalIgnoreCase) Then
                    detQuery = p.Substring(5).Trim()
                ElseIf p.StartsWith("DETF=", StringComparison.OrdinalIgnoreCase) Then
                    detFmt = p.Substring(5).Trim().ToLowerInvariant()
                ElseIf p.StartsWith("DETMAP=", StringComparison.OrdinalIgnoreCase) Then
                    detmapPacked = p.Substring(7)
                ElseIf p.StartsWith("ONCLR=", StringComparison.OrdinalIgnoreCase) Then
                    onClr = p.Substring(6).Trim().ToLowerInvariant()
                ElseIf p.StartsWith("OFFCLR=", StringComparison.OrdinalIgnoreCase) Then
                    offClr = p.Substring(7).Trim().ToLowerInvariant()
                End If
            Next

            ' Must have query
            If deviceName = "" OrElse detQuery = "" Then Continue For

            ' detmap optional but needed for multistate determine
            If detmapPacked = "" Then Continue For

            Dim detmap() As String = detmapPacked.Split("§"c)

            Dim cacheKey As String = deviceName & "||" & detQuery & "||" & detFmt
            Dim reply As String = ""
            If Not cache.TryGetValue(cacheKey, reply) Then
                reply = DetermineQuery(deviceName, detQuery, detFmt)
                cache(cacheKey) = reply
            End If

            Dim chosenIdx As Integer = -1
            For i As Integer = 0 To detmap.Length - 1
                If DetermineMatch(reply, detmap(i), detFmt) Then
                    chosenIdx = i
                    Exit For
                End If
            Next

            UserInitSuppressSend = True
            ApplyMultiButtonVisual(pnl, chosenIdx, onClr, offClr)
            UserInitSuppressSend = False

        Next

    End Sub


    Private Sub MultiButton_Click(sender As Object, e As EventArgs)
        If UserInitSuppressSend Then Exit Sub

        Dim b As Button = DirectCast(sender, Button)
        Dim parts = CStr(b.Tag).Split("|"c)

        Dim device = parts(1)
        Dim prefix = parts(2)
        Dim pnlName = parts(3)

        Dim idx As Integer = -1
        Dim item As String = ""
        Dim cmdsPacked As String = ""
        Dim onClr As String = "limegreen"
        Dim offClr As String = "default"

        For Each p In parts
            If p.StartsWith("IDX=") Then Integer.TryParse(p.Substring(4), idx)
            If p.StartsWith("ITEM=") Then item = p.Substring(5)
            If p.StartsWith("CMDS=") Then cmdsPacked = p.Substring(5)
            If p.StartsWith("ONCLR=") Then onClr = p.Substring(6)
            If p.StartsWith("OFFCLR=") Then offClr = p.Substring(7)
        Next

        Dim cmd As String
        If cmdsPacked <> "" Then
            cmd = cmdsPacked.Split("§"c)(idx)
        Else
            cmd = prefix & If(prefix.EndsWith(" "), "", " ") & item
        End If

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

        If useNative Then
            NativeSend(device, cmd & TermStr2())
        Else
            dev.SendAsync(cmd & TermStr2(), True)
        End If

        Dim pnl = DirectCast(GroupBoxCustom.Controls.Find(pnlName, True)(0), Panel)
        ApplyMultiButtonVisual(pnl, idx, onClr, offClr)
    End Sub


    Private Sub ApplyMultiButtonVisual(pnl As Panel, activeIdx As Integer,
                                  Optional onClr As String = "limegreen",
                                  Optional offClr As String = "default")

        For Each c As Control In pnl.Controls
            Dim b As Button = TryCast(c, Button)
            If b Is Nothing Then Continue For

            Dim idx As Integer = -1
            For Each p In CStr(b.Tag).Split("|"c)
                If p.StartsWith("IDX=") Then Integer.TryParse(p.Substring(4), idx)
            Next

            b.UseVisualStyleBackColor = False
            If idx = activeIdx AndAlso activeIdx >= 0 Then
                b.BackColor = ToggleColorFromName(onClr)
                b.ForeColor = Color.Black
            Else
                b.BackColor = ToggleColorFromName(offClr)
                b.ForeColor = SystemColors.ControlText
            End If
        Next
    End Sub


    Private Sub UserKeypad_SetTarget(sender As Object, e As EventArgs)

        Dim c As Control = TryCast(sender, Control)
        If c Is Nothing Then Exit Sub

        ' TEXTBOX ONLY
        If Not TypeOf c Is TextBox Then Exit Sub

        UserKeypadTarget = c
        UpdateKeypadTargetLabel()

        If Not String.Equals(UserKeypadMode, "popup", StringComparison.OrdinalIgnoreCase) Then Exit Sub

        ' Recreate popup if needed (MUST be before Show)
        If UserKeypadPopupForm Is Nothing OrElse UserKeypadPopupForm.IsDisposed Then
            EnsureKeypadPopup(UserKeypadName, UserKeypadCaption, UserKeypadW, UserKeypadH, UserKeypadTargetLabelOn)
        End If

        ' Safety: if still nothing, bail
        If UserKeypadPopupForm Is Nothing OrElse UserKeypadPopupForm.IsDisposed Then Exit Sub

        If Not UserKeypadPopupForm.Visible Then UserKeypadPopupForm.Show(Me)
        UserKeypadPopupForm.BringToFront()

    End Sub


    Private Sub UpdateKeypadTargetLabel()
        If UserKeypadTargetLabel Is Nothing Then Exit Sub
        If UserKeypadTarget Is Nothing Then
            UserKeypadTargetLabel.Text = "Target: (none)"
        Else
            UserKeypadTargetLabel.Text = "Target: " & UserKeypadTarget.Name
        End If
    End Sub

    Private Sub EnsureKeypadPopup(name As String, caption As String, w As Integer, h As Integer, targetLabelOn As Boolean)
        If UserKeypadPopupForm IsNot Nothing AndAlso Not UserKeypadPopupForm.IsDisposed Then Exit Sub

        UserKeypadPopupForm = New Form()
        UserKeypadPopupForm.FormBorderStyle = FormBorderStyle.FixedToolWindow
        UserKeypadPopupForm.Text = caption
        UserKeypadPopupForm.TopMost = True
        UserKeypadPopupForm.ShowInTaskbar = False
        UserKeypadPopupForm.Size = New Size(w + 16, h + 39)

        Dim pnl = CreateUserKeypadPanel(name, caption, w, h, targetLabelOn)
        pnl.Dock = DockStyle.Fill
        UserKeypadPopupForm.Controls.Add(pnl)

        UserKeypadPopupPanel = pnl

        ' Restore last known position (if we have one)
        If UserKeypadPopupPosValid Then
            UserKeypadPopupForm.StartPosition = FormStartPosition.Manual
            UserKeypadPopupForm.Location = UserKeypadPopupPos
        End If

        ' Track moves so we remember position
        AddHandler UserKeypadPopupForm.Move,
    Sub()
        If UserKeypadPopupForm IsNot Nothing AndAlso Not UserKeypadPopupForm.IsDisposed Then
            UserKeypadPopupPos = UserKeypadPopupForm.Location
            UserKeypadPopupPosValid = True
        End If
    End Sub

    End Sub

    Private Function CreateUserKeypadPanel(name As String, caption As String, w As Integer, h As Integer, targetLabelOn As Boolean) As Panel

        Dim pnl As New Panel()
        pnl.Name = "KP_" & name
        pnl.Size = New Size(w, h)

        Dim topY As Integer = 0

        If targetLabelOn Then
            Dim lbl As New Label()
            lbl.AutoSize = False
            lbl.Height = 18
            lbl.Width = w
            lbl.Location = New Point(0, 0)
            lbl.TextAlign = ContentAlignment.MiddleLeft
            lbl.Text = "Target: (none)"
            pnl.Controls.Add(lbl)
            UserKeypadTargetLabel = lbl
            topY = 22
        End If

        ' 4x4 keypad:
        ' 7 8 9 BK
        ' 4 5 6 CL
        ' 1 2 3 +/-
        ' 0 . OK X
        Dim btnW As Integer = (w - 6) \ 4
        Dim btnH As Integer = (h - topY - 6) \ 4

        Dim keys As String(,) = {
        {"7", "8", "9", "BK"},
        {"4", "5", "6", "CL"},
        {"1", "2", "3", "+/-"},
        {"0", ".", "OK", "X"}
    }

        For r As Integer = 0 To 3
            For c As Integer = 0 To 3
                Dim k As String = keys(r, c)

                Dim b As New Button()
                b.Text = k
                b.UseVisualStyleBackColor = True

                Dim xx As Integer = c * (btnW + 2)
                Dim yy As Integer = topY + r * (btnH + 2)

                b.Location = New Point(xx, yy)
                b.Size = New Size(btnW, btnH)
                b.Tag = k

                AddHandler b.Click, AddressOf UserKeypad_ButtonClick
                pnl.Controls.Add(b)
            Next
        Next

        UpdateKeypadTargetLabel()
        Return pnl
    End Function


    Private Sub UserKeypad_ButtonClick(sender As Object, e As EventArgs)

        Dim b As Button = TryCast(sender, Button)
        If b Is Nothing Then Exit Sub

        Dim key As String = TryCast(b.Tag, String)
        If String.IsNullOrEmpty(key) Then Exit Sub

        Dim tb As TextBox = TryCast(UserKeypadTarget, TextBox)
        If tb Is Nothing OrElse tb.IsDisposed Then
            UserKeypadTarget = Nothing
            UpdateKeypadTargetLabel()
            Exit Sub
        End If

        Select Case key
            Case "BK"
                BackspaceAtCaret(tb)

            Case "CL"
                tb.Text = ""
                tb.SelectionStart = 0
                tb.SelectionLength = 0

            Case "+/-"
                ToggleSign(tb)

            Case "OK"
            ' no-op (you can later hook this to "send" if you want)

            Case "X"
                If String.Equals(UserKeypadMode, "popup", StringComparison.OrdinalIgnoreCase) Then
                    If UserKeypadPopupForm IsNot Nothing Then UserKeypadPopupForm.Hide()
                End If

            Case Else
                InsertAtCaret(tb, key)
        End Select

    End Sub

    Private Sub InsertAtCaret(tb As TextBox, s As String)
        Dim pos As Integer = tb.SelectionStart
        If tb.SelectionLength > 0 Then
            tb.Text = tb.Text.Remove(tb.SelectionStart, tb.SelectionLength)
            pos = tb.SelectionStart
        End If
        tb.Text = tb.Text.Insert(pos, s)
        tb.SelectionStart = pos + s.Length
        tb.SelectionLength = 0
    End Sub

    Private Sub BackspaceAtCaret(tb As TextBox)
        If tb.SelectionLength > 0 Then
            Dim start = tb.SelectionStart
            tb.Text = tb.Text.Remove(start, tb.SelectionLength)
            tb.SelectionStart = start
            tb.SelectionLength = 0
            Exit Sub
        End If

        Dim pos As Integer = tb.SelectionStart
        If pos > 0 AndAlso tb.TextLength > 0 Then
            tb.Text = tb.Text.Remove(pos - 1, 1)
            tb.SelectionStart = pos - 1
            tb.SelectionLength = 0
        End If
    End Sub

    Private Sub ToggleSign(tb As TextBox)
        Dim t As String = tb.Text.Trim()
        If t.StartsWith("-", StringComparison.Ordinal) Then
            t = t.Substring(1)
        ElseIf t <> "" Then
            t = "-" & t
        Else
            t = "-"
        End If
        tb.Text = t
        tb.SelectionStart = tb.TextLength
        tb.SelectionLength = 0
    End Sub


    ' Saving user textbox data to settings
    Private Sub SaveUserTextboxState()

        If Not UserConfig_DataSaveEnabled Then Exit Sub

        If String.IsNullOrWhiteSpace(LastUserConfigPath) Then Exit Sub

        Dim dict As Dictionary(Of String, String) = Nothing

        Try
            dict = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(My.Settings.UserGui_TextboxStateJson)
        Catch
            dict = Nothing
        End Try

        If dict Is Nothing Then
            dict = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        End If

        For Each kvp In UiById
            Dim tb = TryCast(kvp.Value, TextBox)
            If tb Is Nothing Then Continue For
            If tb.IsDisposed Then Continue For

            Dim key As String = LastUserConfigPath.ToLowerInvariant() & "|" & tb.Name.ToLowerInvariant()
            dict(key) = tb.Text
        Next

        My.Settings.UserGui_TextboxStateJson = JsonConvert.SerializeObject(dict)
        My.Settings.Save()
    End Sub

    Private Sub RestoreUserTextboxState()

        If Not UserConfig_DataSaveEnabled Then Exit Sub

        If String.IsNullOrWhiteSpace(LastUserConfigPath) Then Exit Sub

        Dim dict As Dictionary(Of String, String) = Nothing

        Try
            dict = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(My.Settings.UserGui_TextboxStateJson)
        Catch
            dict = Nothing
        End Try

        If dict Is Nothing Then Exit Sub

        For Each kvp In UiById
            Dim tb = TryCast(kvp.Value, TextBox)
            If tb Is Nothing Then Continue For

            Dim key As String = LastUserConfigPath.ToLowerInvariant() & "|" & tb.Name.ToLowerInvariant()

            Dim saved As String = Nothing
            If dict.TryGetValue(key, saved) Then
                tb.Text = saved
            End If
        Next
    End Sub


    Private Function ExtractUnitFromCaption(caption As String) As String
        If String.IsNullOrWhiteSpace(caption) Then Return ""

        caption = caption.Trim()

        ' Split on spaces: "10 uA" → {"10","uA"}
        Dim parts = caption.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)

        ' Need at least value + unit
        If parts.Length < 2 Then Return ""

        ' Unit is last token
        Return parts(parts.Length - 1)
    End Function


    Private Sub PopOutChart(chartName As String)
        Dim cfg As ChartConfig = Nothing
        If Not ChartSettings.TryGetValue(chartName, cfg) _
           OrElse cfg Is Nothing _
           OrElse Not cfg.Popup Then
            Exit Sub
        End If

        ' If popup already exists, just bring it to front
        Dim existingForm As Form = Nothing
        If ChartPopupForms.TryGetValue(chartName, existingForm) _
           AndAlso existingForm IsNot Nothing _
           AndAlso Not existingForm.IsDisposed Then

            existingForm.BringToFront()
            existingForm.Activate()
            Exit Sub
        End If

        ' Build popup form
        Dim f As New Form()
        f.Text = If(String.IsNullOrEmpty(cfg.ChartName), chartName, cfg.ChartName)
        f.StartPosition = FormStartPosition.Manual

        ' Size based on CHART w=/h= from config, with a bit of padding for borders/title
        Dim baseW As Integer = If(cfg.Width > 0, cfg.Width, 800)
        Dim baseH As Integer = If(cfg.Height > 0, cfg.Height, 300)
        f.Size = New Size(baseW + 40, baseH + 80)

        ' Try to position near the original chart if it exists
        Dim orig As DataVisualization.Charting.Chart =
            TryCast(GetControlByName(chartName), DataVisualization.Charting.Chart)
        If orig IsNot Nothing Then
            f.Location = orig.PointToScreen(Point.Empty)
        Else
            f.StartPosition = FormStartPosition.CenterParent
        End If

        Dim ch As New DataVisualization.Charting.Chart()
        ch.Name = chartName & "_popup"
        ch.Dock = DockStyle.Fill
        ch.BackColor = Color.Black

        Dim ca As New DataVisualization.Charting.ChartArea("Default")
        ch.ChartAreas.Add(ca)

        ca.Position.Auto = False
        ca.Position = New DataVisualization.Charting.ElementPosition(0, 0, 100, 100)

        ' Use innerX/innerY/innerW/innerH/labelsize from cfg
        ca.InnerPlotPosition.Auto = False
        Dim ix As Single = If(cfg.InnerX > 0, cfg.InnerX, 2.0F)
        Dim iy As Single = If(cfg.InnerY > 0, cfg.InnerY, 2.0F)
        Dim iw As Single = If(cfg.InnerW > 0, cfg.InnerW, 96.0F)
        Dim ih As Single = If(cfg.InnerH > 0, cfg.InnerH, 96.0F)

        Dim ls As Single = If(cfg.labelsize > 0, cfg.labelsize, 7.0F)
        'Dim rawLs As Double = If(cfg.labelsize > 0, cfg.labelsize, 7.0R)
        'Dim ls As Single = CSng(Math.Round(rawLs, 8))

        ca.InnerPlotPosition = New DataVisualization.Charting.ElementPosition(ix, iy, iw, ih)

        ' Y range / autoscale
        If Not cfg.AutoScaleY Then
            If cfg.YMin.HasValue Then ca.AxisY.Minimum = cfg.YMin.Value
            If cfg.YMax.HasValue Then ca.AxisY.Maximum = cfg.YMax.Value
        Else
            ca.AxisY.Minimum = Double.NaN
            ca.AxisY.Maximum = Double.NaN
        End If

        ' Simple grid / labels styling
        ca.BackColor = Color.Black

        ca.AxisX.LabelStyle.Enabled = False         ' disabled x-axis labels
        ca.AxisX.LabelStyle.ForeColor = Color.White
        ca.AxisX.LabelStyle.Format = "0"        ' secs only
        ca.AxisX.MajorTickMark.Enabled = False
        ca.AxisX.MinorTickMark.Enabled = False
        ca.AxisX.MajorGrid.Enabled = True
        ca.AxisX.MajorGrid.LineColor = Color.FromArgb(60, 60, 60)
        ca.AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        ca.AxisX.Title = "Time (s)"
        ca.AxisX.TitleFont = New Font("Segoe UI", 8.0F, FontStyle.Regular)

        ca.AxisX.Interval = 0.5R
        ca.AxisX.MajorGrid.Interval = 0.5R

        ca.AxisY.MajorGrid.Enabled = True
        ca.AxisY.MajorGrid.LineColor = Color.FromArgb(80, 80, 80)
        ca.AxisY.LabelStyle.ForeColor = Color.White

        ca.AxisY.LabelStyle.Format = "0.########"                   ' Limit Y-axis label decimals to max 8 dp

        'ca.AxisY.LabelStyle.Font = New Font("Segoe UI", 7.0F, FontStyle.Regular)
        ca.AxisX.LabelStyle.Font = New Font("Segoe UI", ls, FontStyle.Regular)
        ca.AxisY.MajorTickMark.Enabled = False

        ' Simple grid styling
        ca.AxisX.MajorTickMark.Enabled = False
        ca.AxisX.MinorTickMark.Enabled = False
        ca.AxisX.MajorGrid.Enabled = True
        ca.AxisX.MajorGrid.LineColor = Color.FromArgb(60, 60, 60)
        ca.AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dot

        ca.AxisY.MajorGrid.Enabled = True
        ca.AxisY.MajorGrid.LineColor = Color.FromArgb(80, 80, 80)
        ca.AxisY.LabelStyle.ForeColor = Color.White
        'ca.AxisY.LabelStyle.Font = New Font("Segoe UI", 7.0F, FontStyle.Regular)
        ca.AxisY.LabelStyle.Font = New Font("Segoe UI", ls, FontStyle.Regular)
        ca.AxisY.MajorTickMark.Enabled = False

        ' Let Y-axis choose more gridlines when there's more height
        ca.AxisY.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        ca.AxisY.Interval = Double.NaN
        ca.AxisY.MajorGrid.Interval = Double.NaN

        ' Series
        Dim s As New DataVisualization.Charting.Series("S1")
        s.ChartType = DataVisualization.Charting.SeriesChartType.Line
        s.BorderWidth = 1
        s.Color = Color.Yellow
        s.MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
        s.MarkerSize = 3
        s.MarkerColor = s.Color
        ch.Series.Add(s)

        ' If original chart exists, copy its basic style (color, width)
        If orig IsNot Nothing AndAlso orig.Series.Count > 0 Then
            Dim s0 = orig.Series(0)
            s.BorderWidth = s0.BorderWidth
            s.Color = s0.Color
            s.MarkerColor = s0.MarkerColor
        End If

        f.Controls.Add(ch)

        ChartPopupForms(chartName) = f

        AddHandler f.FormClosing,
            Sub(sender As Object, args As FormClosingEventArgs)
                ChartPopupForms.Remove(chartName)
            End Sub

        f.Show(Me)
    End Sub


    Private Sub FormMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ' No keypad panel? bail.
        If Keypad1Panel Is Nothing OrElse Not Keypad1Panel.Visible Then Return

        Dim keyChar As Char = ChrW(0)

        Select Case e.KeyCode
            Case Keys.NumPad0, Keys.D0 : keyChar = "0"c
            Case Keys.NumPad1, Keys.D1 : keyChar = "1"c
            Case Keys.NumPad2, Keys.D2 : keyChar = "2"c
            Case Keys.NumPad3, Keys.D3 : keyChar = "3"c
            Case Keys.NumPad4, Keys.D4 : keyChar = "4"c
            Case Keys.NumPad5, Keys.D5 : keyChar = "5"c
            Case Keys.NumPad6, Keys.D6 : keyChar = "6"c
            Case Keys.NumPad7, Keys.D7 : keyChar = "7"c
            Case Keys.NumPad8, Keys.D8 : keyChar = "8"c
            Case Keys.NumPad9, Keys.D9 : keyChar = "9"c

            Case Keys.Decimal, Keys.OemPeriod
                keyChar = "."c

            Case Keys.Add
                keyChar = "+"c

            Case Keys.Subtract
                keyChar = "-"c

            Case Keys.Back
                ' If you have a [←] / [DEL] keypad button, call it directly:
                Dim backBtn = Keypad1Panel.Controls.
                            OfType(Of Button)().
                            FirstOrDefault(Function(b) b.Text = "←")
                If backBtn IsNot Nothing Then
                    backBtn.PerformClick()
                    e.Handled = True
                End If
                Return

            Case Keys.Enter, Keys.Return
                ' If you have an [ENTER] / [OK] keypad button:
                Dim enterBtn = Keypad1Panel.Controls.
                             OfType(Of Button)().
                             FirstOrDefault(Function(b) b.Text.Equals("ENTER", StringComparison.OrdinalIgnoreCase) _
                                            OrElse b.Text.Equals("OK", StringComparison.OrdinalIgnoreCase))
                If enterBtn IsNot Nothing Then
                    enterBtn.PerformClick()
                    e.Handled = True
                End If
                Return

            Case Else
                Return   ' ignore other keys
        End Select

        If keyChar = ChrW(0) Then Return

        ' Find the keypad button whose Text matches the key and "click" it
        Dim btn = Keypad1Panel.Controls.
                OfType(Of Button)().
                FirstOrDefault(Function(b) b.Text = keyChar.ToString())

        If btn IsNot Nothing Then
            btn.PerformClick()   ' reuses your existing keypad click logic
            e.Handled = True
        End If
    End Sub


    Private Sub HandleBootCommandsLine(line As String)

        ' Expected config:
        '   BOOTCOMMANDS;
        '      device=dev2;
        '      DelayPerCmd=0.2;
        '      commandlist=CMD1,CMD2,CMD3
        '
        ' The logical-line joiner will already have turned it into:
        '   "BOOTCOMMANDS;device=dev2;DelayPerCmd=0.2;commandlist=CMD1,CMD2,CMD3"

        'Debug.WriteLine($"BOOTDBG LINE = [{line}]")

        Dim deviceName As String = ""
        Dim commandList As String = ""
        Dim delaySeconds As Double = 0.0R

        Dim work As String = line.Trim()
        Dim segments As String() = work.Split(";"c)

        For Each seg As String In segments
            'Debug.WriteLine($"BOOTDBG RAW SEG   = [{seg}]")

            Dim t As String = seg

            ' Strip inline ;; comment from this segment only
            Dim ccIdx As Integer = t.IndexOf(";;", StringComparison.Ordinal)
            If ccIdx >= 0 Then
                t = t.Substring(0, ccIdx)
            End If

            t = t.Trim()
            'Debug.WriteLine($"BOOTDBG CLEAN SEG = [{t}]")

            If t.Length = 0 Then Continue For

            ' First token: BOOTCOMMANDS (ignore any trailing content in that segment)
            If t.StartsWith("BOOTCOMMANDS", StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            ' Normal key=value segments: device=..., DelayPerCmd=..., commandlist=...
            Dim eqIdx As Integer = t.IndexOf("="c)
            If eqIdx <= 0 OrElse eqIdx >= t.Length - 1 Then Continue For

            Dim key As String = t.Substring(0, eqIdx).Trim().ToLowerInvariant()
            Dim val As String = t.Substring(eqIdx + 1).Trim()

            'Debug.WriteLine($"BOOTDBG KV SEG key=[{key}] val=[{val}]")

            Select Case key
                Case "device"
                    deviceName = val

                Case "delaypercmd"
                    Dim d As Double
                    If Double.TryParse(val,
                               Globalization.NumberStyles.Float,
                               Globalization.CultureInfo.InvariantCulture,
                               d) AndAlso d >= 0.0R Then
                        delaySeconds = d
                    End If

                Case "commandlist"
                    commandList = val
            End Select
        Next

        'Debug.WriteLine($"BOOTDBG DEVICE = [{deviceName}]")
        'Debug.WriteLine($"BOOTDBG DELAY  = [{delaySeconds}]")
        'Debug.WriteLine($"BOOTDBG CMDLIST RAW = [{commandList}]")

        If String.IsNullOrWhiteSpace(deviceName) OrElse String.IsNullOrWhiteSpace(commandList) Then
            AppendLog("[BOOT] bootcommands line missing device or commandlist: " & line)
            Return
        End If

        ' Split commandlist on commas → individual commands
        Dim cmds As New List(Of String)

        For Each c As String In commandList.Split(","c)
            Dim token As String = c
            'Debug.WriteLine($"BOOTDBG RAW CMD   = [{token}]")

            ' Strip inline ;; comment from the command token
            Dim ccIdx As Integer = token.IndexOf(";;", StringComparison.Ordinal)
            If ccIdx >= 0 Then
                token = token.Substring(0, ccIdx)
            End If

            Dim cmd As String = token.Trim()
            'Debug.WriteLine($"BOOTDBG CLEAN CMD = [{cmd}]")

            If cmd <> "" Then cmds.Add(cmd)
        Next

        'Debug.WriteLine($"BOOTDBG CMD COUNT = {cmds.Count}")

        If cmds.Count = 0 Then
            AppendLog("[BOOT] bootcommands commandlist empty for device " & deviceName)
            Return
        End If

        RunBootCommands(deviceName, cmds, delaySeconds)
    End Sub



    Private Sub RunBootCommands(deviceName As String,
                            commands As IEnumerable(Of String),
                            Optional delaySeconds As Double = 0.0R)

        If String.IsNullOrWhiteSpace(deviceName) OrElse commands Is Nothing Then Return

        ' For now, support bootcommands only when using native engine,
        ' so behaviour exactly matches CASE "SEND" with NativeSend.
        If Not IsNativeEngine(deviceName) Then
            AppendLog("[BOOT] bootcommands currently only supported with native engine for " & deviceName)
            Return
        End If

        For Each cmdRaw As String In commands
            Dim cmd As String = cmdRaw.Trim()
            If cmd = "" Then Continue For

            Try
                ' EXACTLY the same native path as your CASE "SEND"
                NativeSend(deviceName, cmd)
                AppendLog($"[BOOT] {deviceName} << {cmd}")

                ' NEW: optional delay between commands (seconds → ms)
                If delaySeconds > 0.0R Then
                    Threading.Thread.Sleep(CInt(delaySeconds * 1000.0R))
                End If

            Catch ex As Exception
                AppendLog($"[BOOT] Error sending '{cmd}' to {deviceName}: {ex.Message}")
            End Try
        Next
    End Sub


    Private Sub AddToUserContainer(ctrl As Control)
        Dim parent As Control

        If UserTabControl Is Nothing Then
            ' No sub-tabs in use: everything goes directly on GroupBoxCustom
            parent = GroupBoxCustom
        Else
            ' Sub-tabs active: prefer current tab page, fall back to first tab if needed
            If CurrentUserContainer IsNot Nothing Then
                parent = CurrentUserContainer
            ElseIf UserTabControl.TabPages.Count > 0 Then
                parent = UserTabControl.TabPages(0)
            Else
                parent = GroupBoxCustom
            End If
        End If

        parent.Controls.Add(ctrl)
    End Sub


    ' Return all top-level user controls, including those on sub-tabs
    Private Function GetUserDetermineControls() As List(Of Control)

        Dim result As New List(Of Control)

        For Each ctrl As Control In GroupBoxCustom.Controls
            Dim tc As TabControl = TryCast(ctrl, TabControl)
            If tc Is Nothing Then
                ' Normal child of GroupBoxCustom (old behaviour)
                result.Add(ctrl)
            Else
                ' Include children of each TabPage
                For Each page As TabPage In tc.TabPages
                    For Each child As Control In page.Controls
                        result.Add(child)
                    Next
                Next
            End If
        Next

        Return result

    End Function


    ' Multiline-aware parser for fields like:
    '   commands=a,b,
    '       c,d,
    '       e;
    ' or single-line:
    '   commands=a,b,c,d,e;
    Private Function ParseMultiValueField(parts() As String, key As String) As String
        ' Join the whole MULTIBUTTON block back together so we can
        ' reliably grab: key= ... up to the next semicolon.
        Dim block As String = String.Join(";", parts)

        ' Find "key=" (e.g. "items=" or "commands=" or "detmap=")
        Dim idx As Integer = block.IndexOf(key & "=", StringComparison.OrdinalIgnoreCase)
        If idx < 0 Then
            Return ""
        End If

        ' Start of the value, just after "key="
        idx += key.Length + 1

        ' Value runs up to the next ';' or end of string
        Dim endIdx As Integer = block.IndexOf(";"c, idx)
        Dim raw As String
        If endIdx >= 0 Then
            raw = block.Substring(idx, endIdx - idx)
        Else
            raw = block.Substring(idx)
        End If

        ' Now raw is something like:
        '   "ACV,
        '    DCV,
        '    R2W,
        '    ..."
        ' Split on commas, trim spaces/CRLF, drop empties
        Dim vals =
            raw.Split(","c).
                Select(Function(s) s.Trim()).
                Where(Function(s) s <> "").
                ToArray()

        Return String.Join(",", vals)
    End Function













End Class