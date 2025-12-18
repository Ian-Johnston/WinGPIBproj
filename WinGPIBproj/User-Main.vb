' User customizeable tab

Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms.DataVisualization.Charting
Imports IODevices



Partial Class Formtest

    ' Auto-read state - Timer 5 (dev1)
    Private AutoReadDeviceName As String
    Private AutoReadCommand As String
    Private AutoReadResultControl As String
    'Private AutoReadAction As String
    Private AutoReadValueControl As String
    Private UserAutoBusy As Integer = 0

    ' Auto-read state - Timer 16 (dev2)
    Private AutoReadDeviceName2 As String = ""
    Private AutoReadCommand2 As String = ""
    Private AutoReadResultControl2 As String = ""
    'Private AutoReadAction2 As String
    Private AutoReadValueControl2 As String
    Private UserAutoBusy2 As Integer = 0

    Private UserLayoutGen As Integer = 0

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

    ' Stats panel runtime state =====
    Private StatsState As New Dictionary(Of String, RunningStatsState)(StringComparer.OrdinalIgnoreCase)
    ' Holds the per-panel UI row outputs (label -> value label)
    Private StatsRows As New Dictionary(Of String, List(Of StatsRow))(StringComparer.OrdinalIgnoreCase)

    ' ===== History grid runtime state =====
    Private HistoryTables As New Dictionary(Of String, DataTable)(StringComparer.OrdinalIgnoreCase)
    Private HistoryStates As New Dictionary(Of String, RunningStatsState)(StringComparer.OrdinalIgnoreCase)
    Private HistoryPpmRef As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    Private HistoryMaxRows As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    Private HistoryFormats As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    Private HistoryState As New Dictionary(Of String, RunningStatsState)(StringComparer.OrdinalIgnoreCase)
    Private HistoryLists As New Dictionary(Of String, BindingSource)(StringComparer.OrdinalIgnoreCase)
    Private HistoryCols As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
    Dim colsRaw As String = "Value,Time"   ' default if cols= not provided
    Private HistoryData As New Dictionary(Of String, BindingList(Of HistoryRow))(StringComparer.OrdinalIgnoreCase)
    Private HistorySettings As New Dictionary(Of String, HistoryGridConfig)(StringComparer.OrdinalIgnoreCase)
    Private HistoryGrids As New Dictionary(Of String, DataGridView)(StringComparer.OrdinalIgnoreCase)

    ' Invisibility
    Private UiById As New Dictionary(Of String, Control)(StringComparer.OrdinalIgnoreCase)      ' Invisibility: every created control, by ID (ChartDMM, Hist1, Stats1, ...)
    Private HideFuncToTargetId As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)       ' Invisibility: function-name -> target-control-id mapping (ChartDMMhide -> ChartDMM)
    Private InvisFuncToTargets As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)

    ' Trigger
    Private TriggerEng As TriggerEngine

    ' LUA
    Dim inLuaScript As Boolean = False
    Dim luaScriptName As String = ""
    Dim luaLines As New List(Of String)
    Dim autoY As Integer = 10


    ' History Grid
    Private Class HistoryGridConfig
        Public Property GridName As String
        Public Property ResultTarget As String
        Public Property MaxRows As Integer
        Public Property Format As String
        Public Property Columns As List(Of String)   ' <-- MUST be a list
        Public Property PpmRef As String
    End Class


    ' History Grid
    Private Class HistoryRow
        Public Property Value As Double
        Public Property Time As DateTime
        Public Property Min As Double
        Public Property Max As Double
        Public Property PkPk As Double
        Public Property Mean As Double
        Public Property Std As Double
        Public Property PPM As Double
        Public Property Count As Long
    End Class


    Private Sub ButtonLoadTxt_Click(sender As Object, e As EventArgs) Handles ButtonLoadTxt.Click

        Using dlg As New OpenFileDialog()
            dlg.Title = "Select Custom GUI Layout File"
            dlg.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

            If dlg.ShowDialog() <> DialogResult.OK Then
                Exit Sub   ' user cancelled → leave label alone
            End If

            Try
                LoadCustomGuiFromFile(dlg.FileName)
                LabelUSERtab1.Visible = False   ' only hide if load succeeds
            Catch ex As Exception
                MessageBox.Show("Error loading layout file: " & ex.Message)
                LabelUSERtab1.Visible = True    ' keep it visible on failure
            End Try
        End Using

    End Sub



    Private Sub ButtonResetTxt_Click(sender As Object, e As EventArgs) Handles ButtonResetTxt.Click

        ' Invalidate any in-flight async UI updates
        Threading.Interlocked.Increment(UserLayoutGen)

        ' Stop User tab dev1 and dev2 timers
        Me.Timer5.Stop()
        Me.Timer16.Stop()

        ' Clear slot1
        AutoReadDeviceName = ""
        AutoReadCommand = ""
        AutoReadResultControl = ""
        Threading.Interlocked.Exchange(UserAutoBusy, 0)

        ' Clear slot2
        AutoReadDeviceName2 = ""
        AutoReadCommand2 = ""
        AutoReadResultControl2 = ""
        Threading.Interlocked.Exchange(UserAutoBusy2, 0)

        ' Clear invisibility
        UiById.Clear()
        InvisFuncToTargets.Clear()

        ' Remove all dynamically created controls
        'GroupBoxCustom.Controls.Clear()

        For i As Integer = GroupBoxCustom.Controls.Count - 1 To 0 Step -1
            Dim c = GroupBoxCustom.Controls(i)
            If c Is LabelUSERtab1 Then Continue For
            GroupBoxCustom.Controls.RemoveAt(i)
            c.Dispose()
        Next

        ' Reset the title text
        GroupBoxCustom.Text = "User Defineable"

        If LabelUSERtab1.IsDisposed Then
            MessageBox.Show("LabelUSERtab1 was disposed (it must be outside the dynamic clear).")
        Else
            If LabelUSERtab1.Parent IsNot GroupBoxCustom Then
                LabelUSERtab1.Parent = GroupBoxCustom
            End If

            LabelUSERtab1.Visible = True
            LabelUSERtab1.BringToFront()
        End If

    End Sub


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


    ' Parser loop
    Private Sub BuildCustomGuiFromText(def As String)

        _lockedOriginalSizes.Clear()
        _lastConfigLines.Clear()
        _lastAutoLayoutEnabled = False

        ' Reset per-config runtime state
        Timer5.Enabled = False
        AutoReadDeviceName = ""
        AutoReadCommand = ""
        AutoReadResultControl = ""
        AutoReadValueControl = ""

        Timer16.Enabled = False
        AutoReadDeviceName2 = ""
        AutoReadCommand2 = ""
        AutoReadResultControl2 = ""
        AutoReadValueControl2 = ""

        AutoLayoutEnabled = False

        Threading.Interlocked.Exchange(UserAutoBusy2, 0)

        GroupBoxCustom.Controls.Clear()
        LuaScriptsByName.Clear()

        CurrentUserScale = 1.0   ' reset scale each time we load a layout

        Dim autoY As Integer = 10

        For Each rawLine In def.Split({vbCrLf, vbLf}, StringSplitOptions.None)

            ' Capture raw line exactly (INCLUDING blank lines)
            _lastConfigLines.Add(rawLine)

            Dim line = rawLine.Trim()
            If line = "" OrElse line.StartsWith("'") Then Continue For

            If line.StartsWith("AUTOLAYOUT", StringComparison.OrdinalIgnoreCase) Then
                Dim partsAL = line.Split({"="c}, 2)
                If partsAL.Length = 2 Then
                    AutoLayoutEnabled =
                (partsAL(1).Trim().Equals("TRUE", StringComparison.OrdinalIgnoreCase) OrElse
                 partsAL(1).Trim().Equals("1"))
                    _lastAutoLayoutEnabled = AutoLayoutEnabled
                End If
                Continue For
            End If


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
                        panel.Tag = "LOCKSIZE"

                        If Not _lockedOriginalSizes.ContainsKey(panel.Name) Then
                            _lockedOriginalSizes(panel.Name) = panel.Size
                        End If


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

                        If Not unitsOn Then
                            lbl.Tag = "BIGTEXT_UNITS_OFF"
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

                    If Not _lockedOriginalSizes.ContainsKey(Chart.Name) Then
                        _lockedOriginalSizes(Chart.Name) = Chart.Size
                    End If

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

                                Dim d As Double
                                If Not TryExtractFirstDouble(DirectCast(senderSrc, TextBox).Text, d) Then Exit Sub

                                st.AddSample(d)

                                Dim row As New HistoryRow()
                                row.Time = DateTime.Now
                                row.Value = d
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



            End Select

        Next

        ' apply auto-layout (overrides any x/y from config)
        If AutoLayoutEnabled Then
            ApplyAutoLayout(GroupBoxCustom)
            Dim exported As String = ExportConfigWithCurrentXY(_lastConfigLines, GroupBoxCustom)
            ShowExportPopup(exported)

            If Not AutoLayoutResizeHooked Then
                AddHandler GroupBoxCustom.Resize, AddressOf GroupBoxCustom_ResizeAutoLayout
                AutoLayoutResizeHooked = True
            End If
        End If


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


        ' ================= INVISIBILITY FUNCTION INTERCEPT =================
        For Each p As String In parts
            Dim raw As String = If(p, "").Trim()
            If raw = "" Then Continue For

            ' If the token has config tail like ";x=..;y=..", strip it off
            Dim s As String = raw.Split(";"c)(0).Trim()

            If s <> "" AndAlso InvisFuncToTargets.ContainsKey(s) Then
                TryRunInvisibilityFunction(s)
                Exit Sub
            End If
        Next
        ' ==================================================================




        If parts.Length < 3 Then Exit Sub

        Dim action = parts(0)                 ' SEND / SENDVALUE / QUERY / SENDLINES / CLEARCHART
        Dim deviceName = parts(1)             ' "dev1" / "dev2" / or ChartName for CLEARCHART
        Dim commandOrPrefix = parts(2)
        Dim valueControlName = If(parts.Length > 3, parts(3), "")
        Dim resultControlName = If(parts.Length > 4, parts(4), "")

        ' INVISIBILITY: your hide buttons store the function name in resultControlName
        If Not String.IsNullOrWhiteSpace(resultControlName) Then
            Dim fn = resultControlName.Split(";"c)(0).Trim()
            If TryRunInvisibilityFunction(fn) Then Exit Sub
        End If

        ' --- INVISIBILITY / Function buttons (no device needed) ---
        ' Your config example leaves action/device/command blank and puts the function name in resultControlName
        Dim funcName As String = ""

        If String.IsNullOrWhiteSpace(action) AndAlso String.IsNullOrWhiteSpace(deviceName) AndAlso String.IsNullOrWhiteSpace(commandOrPrefix) Then

            funcName = resultControlName
        ElseIf String.Equals(action, "FUNC", StringComparison.OrdinalIgnoreCase) Then
            ' (optional) if you later choose to use BUTTON;Caption;FUNC;;;;FuncName
            funcName = resultControlName
        End If

        If Not String.IsNullOrWhiteSpace(funcName) Then
            If TryRunInvisibilityFunction(funcName) Then Exit Sub
        End If
        ' ---------------------------------------------------------


        ' Only resolve a GPIB device for actions that actually need one
        Dim dev As IODevices.IODevice = Nothing

        If Not String.IsNullOrWhiteSpace(action) AndAlso
            Not String.Equals(action, "CLEARCHART", StringComparison.OrdinalIgnoreCase) AndAlso
            Not String.Equals(action, "RESETSTATS", StringComparison.OrdinalIgnoreCase) AndAlso
            Not String.Equals(action, "CLEARHISTORY", StringComparison.OrdinalIgnoreCase) AndAlso
            Not String.Equals(action, "RUNLUA", StringComparison.OrdinalIgnoreCase) AndAlso
            Not String.Equals(action, "STOPLUA", StringComparison.OrdinalIgnoreCase) Then

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

                    Const MIN_AUTOREAD_MS As Integer = 1
                    Const MAX_AUTOREAD_MS As Integer = 60000

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
                            intervalMs = CInt(Math.Ceiling(secs * 1000.0R))
                        End If
                    End If

                    If intervalMs < MIN_AUTOREAD_MS Then intervalMs = MIN_AUTOREAD_MS
                    If intervalMs > MAX_AUTOREAD_MS Then intervalMs = MAX_AUTOREAD_MS

                    If String.Equals(deviceName, "dev2", StringComparison.OrdinalIgnoreCase) Then
                        ' --- slot #2 / dev2 ---
                        AutoReadDeviceName2 = deviceName
                        AutoReadCommand2 = commandOrPrefix
                        AutoReadResultControl2 = resultControlName

                        Timer16.Interval = intervalMs
                        Timer16.Enabled = True
                    Else
                        ' --- slot #1 / dev1 ---
                        AutoReadDeviceName = deviceName
                        AutoReadCommand = commandOrPrefix
                        AutoReadResultControl = resultControlName

                        Timer5.Interval = intervalMs
                        Timer5.Enabled = True
                    End If

                Else
                    ' Auto refresh OFF for this specific device
                    If String.Equals(deviceName, "dev2", StringComparison.OrdinalIgnoreCase) Then
                        Timer16.Enabled = False
                        AutoReadDeviceName2 = ""
                        AutoReadCommand2 = ""
                        AutoReadResultControl2 = ""
                    Else
                        Timer5.Enabled = False
                        AutoReadDeviceName = ""
                        AutoReadCommand = ""
                        AutoReadResultControl = ""
                    End If
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


            Case "RESETSTATS"
                ' deviceName holds panelName for RESETSTATS
                Dim panelName As String = deviceName

                If StatsState.ContainsKey(panelName) Then
                    StatsState(panelName).Reset()
                    UpdateStatsPanel(panelName, Double.NaN) ' refresh display
                End If


            Case "CLEARHISTORY"
                Dim gridName As String = deviceName

                If HistoryData.ContainsKey(gridName) Then
                    HistoryData(gridName).Clear()
                End If

                If HistoryState.ContainsKey(gridName) Then
                    HistoryState(gridName).Reset()
                End If

                Exit Sub


            Case "RUNLUA"
                Dim scriptName As String = commandOrPrefix.Trim()

                ' 1) First try config-defined LUASCRIPT
                Dim luaText As String = ""
                If LuaScriptsByName.TryGetValue(scriptName, luaText) Then
                    RunLua(luaText.Replace("|", vbCrLf))
                    Exit Sub
                End If

                ' 2) Fallback: old TEXTAREA method still allowed
                Dim tb = TryCast(GetControlByName(scriptName), TextBoxBase)
                If tb Is Nothing Then
                    MessageBox.Show("Lua script not found: " & scriptName)
                    Exit Sub
                End If

                RunLua(String.Join(Environment.NewLine, tb.Lines))


            Case "STOPLUA"
                RequestLuaStop()
                Exit Sub



        End Select

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

        Dim myGen As Integer = Threading.Interlocked.CompareExchange(UserLayoutGen, 0, 0)

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

                                                        ' If layout was reset/rebuilt while we were running, ignore
                                                        If myGen <> Threading.Interlocked.CompareExchange(UserLayoutGen, 0, 0) Then
                                                            Threading.Interlocked.Exchange(UserAutoBusy, 0)
                                                            Return
                                                        End If

                                                        RunQueryToResult(devName, cmd, resultName, raw)
                                                        Threading.Interlocked.Exchange(UserAutoBusy, 0)
                                                    End Sub)

                                 End Sub)

    End Sub


    Private Sub Timer16_Tick(sender As Object, e As EventArgs) Handles Timer16.Tick

        Dim myGen As Integer = Threading.Interlocked.CompareExchange(UserLayoutGen, 0, 0)

        Dim autoCb As CheckBox = GetCheckboxFor(AutoReadResultControl2, "FuncAuto")
        If autoCb Is Nothing OrElse Not autoCb.Checked Then
            Timer16.Enabled = False
            Exit Sub
        End If

        If String.IsNullOrEmpty(AutoReadDeviceName2) OrElse String.IsNullOrEmpty(AutoReadCommand2) Then
            Timer16.Enabled = False
            Exit Sub
        End If

        ' Prevent overlap (timer ticks while query still running)
        If Threading.Interlocked.Exchange(UserAutoBusy2, 1) = 1 Then Exit Sub

        Dim devName As String = AutoReadDeviceName2
        Dim cmd As String = AutoReadCommand2
        Dim resultName As String = AutoReadResultControl2

        Threading.Tasks.Task.Run(Sub()

                                     Dim raw As String = QueryRawOnly(devName, cmd)

                                     ' If busy/no update, just release lock
                                     If raw = "" Then
                                         Threading.Interlocked.Exchange(UserAutoBusy2, 0)
                                         Return
                                     End If

                                     Me.BeginInvoke(Sub()

                                                        ' If layout was reset/rebuilt while we were running, ignore
                                                        If myGen <> Threading.Interlocked.CompareExchange(UserLayoutGen, 0, 0) Then
                                                            Threading.Interlocked.Exchange(UserAutoBusy, 0)
                                                            Return
                                                        End If

                                                        RunQueryToResult(devName, cmd, resultName, raw)
                                                        Threading.Interlocked.Exchange(UserAutoBusy2, 0)
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


    ' ===== One output row in a stats panel =====
    Private Class StatsRow
        Public LabelText As String
        Public FuncKey As String
        Public RefToken As String
        Public ValueLabel As Label
        Public FormatOverride As String   ' "", "F", or "E"
    End Class


    ' ===== Running stats (fast, stable) =====
    Private Class RunningStatsState
        Public Count As Long
        Public Min As Double = Double.PositiveInfinity
        Public Max As Double = Double.NegativeInfinity
        Public Last As Double
        Public First As Double
        Public HasFirst As Boolean

        ' Welford for mean/std
        Public Mean As Double
        Public M2 As Double


        Public Sub Reset()
            Count = 0
            Min = Double.PositiveInfinity
            Max = Double.NegativeInfinity
            Last = 0
            First = 0
            HasFirst = False
            Mean = 0
            M2 = 0
        End Sub


        Public Sub AddSample(x As Double)
            Last = x
            If Not HasFirst Then
                First = x
                HasFirst = True
            End If

            If x < Min Then Min = x
            If x > Max Then Max = x

            Count += 1
            Dim delta As Double = x - Mean
            Mean += delta / Count
            Dim delta2 As Double = x - Mean
            M2 += delta * delta2
        End Sub


        Public Function StdDevSample() As Double
            If Count < 2 Then Return 0
            Return Math.Sqrt(M2 / (Count - 1))
        End Function
    End Class


    Private Function TryExtractFirstDouble(text As String, ByRef value As Double) As Boolean
        If String.IsNullOrWhiteSpace(text) Then Return False

        ' split on common separators, take first token that parses
        Dim tokens = text.Split({","c, ";"c, " "c, ControlChars.Cr, ControlChars.Lf},
                            StringSplitOptions.RemoveEmptyEntries)

        For Each tok In tokens
            Dim t = tok.Trim()
            If Double.TryParse(t, Globalization.NumberStyles.Float,
                           Globalization.CultureInfo.InvariantCulture, value) Then
                If Not Double.IsNaN(value) AndAlso Not Double.IsInfinity(value) Then Return True
            End If
        Next

        Return False
    End Function


    Private Sub UpdateStatsPanel(panelName As String, sample As Double)

        If Not StatsState.ContainsKey(panelName) Then Exit Sub
        If Not StatsRows.ContainsKey(panelName) Then Exit Sub

        Dim st = StatsState(panelName)

        ' Add new sample (unless refresh-only)
        If Not Double.IsNaN(sample) Then
            st.AddSample(sample)
        End If

        Dim gb = TryCast(GetControlByName(panelName), GroupBox)
        If gb Is Nothing Then Exit Sub

        ' Panel default format (e.g. G6)
        Dim tagStr = TryCast(gb.Tag, String)
        Dim panelFmt As String = "G6"
        If Not String.IsNullOrEmpty(tagStr) Then
            Dim tp = tagStr.Split("|"c)
            If tp.Length >= 3 Then panelFmt = tp(2)
        End If

        For Each r In StatsRows(panelName)

            Dim outVal As String = ""

            Select Case r.FuncKey

                Case "MIN"
                    If st.Count > 0 Then outVal = FormatStatValue(st.Min, panelFmt, r.FormatOverride)

                Case "MAX"
                    If st.Count > 0 Then outVal = FormatStatValue(st.Max, panelFmt, r.FormatOverride)

                Case "PKPK"
                    If st.Count > 0 Then outVal = FormatStatValue(st.Max - st.Min, panelFmt, r.FormatOverride)

                Case "MEAN"
                    If st.Count > 0 Then outVal = FormatStatValue(st.Mean, panelFmt, r.FormatOverride)

                Case "STD"
                    If st.Count > 1 Then outVal = FormatStatValue(st.StdDevSample(), panelFmt, r.FormatOverride)

                Case "LAST"
                    If st.Count > 0 Then outVal = FormatStatValue(st.Last, panelFmt, r.FormatOverride)

                Case "COUNT"
                    outVal = st.Count.ToString(Globalization.CultureInfo.InvariantCulture)

                Case "PPM"
                    outVal = ComputePpmString(st, r.RefToken, panelFmt)

                Case Else
                    outVal = ""
            End Select

            r.ValueLabel.Text = outVal
        Next

        ' Publish stats so TRIGGER expressions can read them
        Dim stdNow As Double = 0
        If st.Count > 1 Then stdNow = st.StdDevSample()

        PublishStats(panelName,
             st.Last,
             st.Mean,
             stdNow,
             If(st.Count > 0, st.Max - st.Min, 0),
             st.Min,
             st.Max,
             st.Count)

    End Sub


    Private Function ComputePpmString(st As RunningStatsState, refTok As String, fmt As String) As String
        If st.Count = 0 Then Return ""

        Dim refVal As Double = 0
        Dim rt As String = If(refTok, "").Trim()

        If rt = "" OrElse rt.Equals("MEAN", StringComparison.OrdinalIgnoreCase) Then
            refVal = st.Mean
        ElseIf rt.Equals("FIRST", StringComparison.OrdinalIgnoreCase) Then
            If Not st.HasFirst Then Return ""
            refVal = st.First
        Else
            Double.TryParse(rt, Globalization.NumberStyles.Float,
                        Globalization.CultureInfo.InvariantCulture, refVal)
        End If

        If refVal = 0 Then Return ""

        Dim ppm As Double = ((st.Last - refVal) / refVal) * 1000000.0R
        Return ppm.ToString(fmt, Globalization.CultureInfo.InvariantCulture)
    End Function


    Private Function FormatStatValue(value As Double,
                                 panelFmt As String,
                                 rowFmt As String) As String
        Dim f As String = panelFmt

        If Not String.IsNullOrEmpty(rowFmt) Then
            If rowFmt = "F" Then
                f = panelFmt   ' respect STATSPANEL format
            ElseIf rowFmt = "E" Then
                f = "E6"        ' forced scientific
            End If
        End If

        Return value.ToString(f, Globalization.CultureInfo.InvariantCulture)
    End Function


    Private Sub UpdateHistoryGrid(gridName As String, sample As Double)

        If Not HistoryState.ContainsKey(gridName) Then Exit Sub
        If Not HistoryLists.ContainsKey(gridName) Then Exit Sub

        Dim st = HistoryState(gridName)
        st.AddSample(sample)

        Dim fmt As String = "G6"
        If HistoryFormats.ContainsKey(gridName) Then fmt = HistoryFormats(gridName)

        Dim ppmRef As String = "MEAN"
        If HistoryPpmRef.ContainsKey(gridName) Then ppmRef = HistoryPpmRef(gridName)

        Dim row As New HistoryRow()
        row.Value = sample.ToString(fmt, Globalization.CultureInfo.InvariantCulture)
        row.Time = DateTime.Now.ToString("HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
        row.Min = If(st.Count > 0, st.Min.ToString(fmt, Globalization.CultureInfo.InvariantCulture), "")
        row.Max = If(st.Count > 0, st.Max.ToString(fmt, Globalization.CultureInfo.InvariantCulture), "")
        row.PkPk = If(st.Count > 0, (st.Max - st.Min).ToString(fmt, Globalization.CultureInfo.InvariantCulture), "")
        row.Mean = If(st.Count > 0, st.Mean.ToString(fmt, Globalization.CultureInfo.InvariantCulture), "")
        row.Std = If(st.Count > 1, st.StdDevSample().ToString(fmt, Globalization.CultureInfo.InvariantCulture), "")
        row.Count = st.Count.ToString(Globalization.CultureInfo.InvariantCulture)
        row.Value = st.Last
        row.PPM = ComputePpmString(st, ppmRef, fmt)

        Dim bs = HistoryLists(gridName)
        Dim lst = TryCast(bs.DataSource, List(Of HistoryRow))
        If lst Is Nothing Then Exit Sub

        ' Newest at TOP
        lst.Insert(0, row)

        Dim maxRows As Integer = HistoryMaxRows(gridName)
        While lst.Count > maxRows
            lst.RemoveAt(lst.Count - 1)
        End While

        bs.ResetBindings(False)

    End Sub


    Private Sub RegisterDynamicControl(id As String, ctrl As Control)
        If String.IsNullOrWhiteSpace(id) Then Return
        If ctrl Is Nothing Then Return
        UiById(id.Trim()) = ctrl
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
            ' ---------- Existing positional behaviour ----------
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

        ' ---------- Named tokens behaviour ----------
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
            End Select

            ' When we have both, commit a mapping and clear for next pair
            If pendingTargets <> "" AndAlso pendingFunc <> "" Then
                Dim targets As New List(Of String)
                For Each t In pendingTargets.Split(","c)
                    Dim tid = t.Trim()
                    If tid <> "" Then targets.Add(tid)
                Next

                If targets.Count > 0 Then
                    InvisFuncToTargets(pendingFunc.Trim()) = targets
                End If

                pendingTargets = ""
                pendingFunc = ""
            End If
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
            End If
        End If
    End Sub



    Private Function TryRunInvisibilityFunction(funcName As String) As Boolean
        If String.IsNullOrWhiteSpace(funcName) Then Return False

        Dim targets As List(Of String) = Nothing
        If Not InvisFuncToTargets.TryGetValue(funcName.Trim(), targets) Then Return False

        ' Toggle visibility of all targets in the list
        For Each id In targets
            Dim c As Control = Nothing
            If UiById.TryGetValue(id, c) AndAlso c IsNot Nothing Then
                c.Visible = Not c.Visible
                If c.Visible Then c.BringToFront() ' helpful for stacked overlays
            End If
        Next

        Return True
    End Function


    Private Sub ClearHistoryGrid(gridName As String)
        If String.IsNullOrWhiteSpace(gridName) Then Exit Sub

        Dim ctrl As Control = Nothing
        If Not ControlByName.TryGetValue(gridName, ctrl) OrElse ctrl Is Nothing Then Exit Sub

        Dim dgv = TryCast(ctrl, DataGridView)
        If dgv Is Nothing Then Exit Sub

        ' If you bound to a DataTable / BindingSource, clear the underlying source instead.
        dgv.Rows.Clear()
    End Sub


    ' ====================================================================================================================================================================================================

    ' VARIABLE REGISTRY
    ' Universal registry of live values (case-insensitive)
    Private ReadOnly Vars As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

    ' Control registries so triggers can invoke actions by name
    Private ReadOnly BtnByName As New Dictionary(Of String, Button)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly TextAreaByName As New Dictionary(Of String, TextBox)(StringComparer.OrdinalIgnoreCase) ' multiline
    Private ReadOnly ControlByName As New Dictionary(Of String, Control)(StringComparer.OrdinalIgnoreCase)

    ' Trigger engine instance
    ' PUBLISH METHODS
    ' Publish any textbox value (including ResultTarget textbox)
    Private Sub PublishTextBox(controlName As String, textValue As String)
        If String.IsNullOrWhiteSpace(controlName) Then Exit Sub

        Vars($"tb:{controlName}") = textValue

        Dim d As Double
        If Double.TryParse(textValue, NumberStyles.Float, CultureInfo.InvariantCulture, d) Then
            Vars($"num:{controlName}") = d
        End If
    End Sub


    ' Publish BIGTEXT (if BIGTEXT has its own name, publish it too)
    Private Sub PublishBigText(controlName As String, textValue As String)
        If String.IsNullOrWhiteSpace(controlName) Then Exit Sub

        Vars($"big:{controlName}") = textValue

        Dim d As Double
        If Double.TryParse(textValue, NumberStyles.Float, CultureInfo.InvariantCulture, d) Then
            Vars($"bignum:{controlName}") = d
        End If
    End Sub


    ' Publish STATSPANEL live metrics (call each time stats update)
    Private Sub PublishStats(panelName As String,
                         lastVal As Double,
                         meanVal As Double,
                         stdVal As Double,
                         pkpkVal As Double,
                         minVal As Double,
                         maxVal As Double,
                         countVal As Integer)

        If String.IsNullOrWhiteSpace(panelName) Then Exit Sub

        Vars($"stats:{panelName}.last") = lastVal
        Vars($"stats:{panelName}.mean") = meanVal
        Vars($"stats:{panelName}.std") = stdVal
        Vars($"stats:{panelName}.pkpk") = pkpkVal
        Vars($"stats:{panelName}.min") = minVal
        Vars($"stats:{panelName}.max") = maxVal
        Vars($"stats:{panelName}.count") = countVal
    End Sub


    ' REGISTRATION (call as you create dynamic controls)
    Private Sub RegisterButton(controlName As String, btn As Button)
        If String.IsNullOrWhiteSpace(controlName) OrElse btn Is Nothing Then Exit Sub
        BtnByName(controlName) = btn
        ControlByName(controlName) = btn
    End Sub


    Private Sub RegisterTextArea(controlName As String, tb As TextBox)
        If String.IsNullOrWhiteSpace(controlName) OrElse tb Is Nothing Then Exit Sub
        TextAreaByName(controlName) = tb
        ControlByName(controlName) = tb
    End Sub


    Private Sub RegisterAnyControl(controlName As String, ctrl As Control)
        If String.IsNullOrWhiteSpace(controlName) OrElse ctrl Is Nothing Then Exit Sub
        ControlByName(controlName) = ctrl
    End Sub


    ' TRIGGER ENGINE STARTUP
    Private Sub EnsureTriggerEngine()
        If TriggerEng Is Nothing Then
            TriggerEng = New TriggerEngine(AddressOf GetVarValue,
                                       AddressOf ExecuteThenChain,
                                       AddressOf AppendLog)

            TriggerEng.Start()
        End If

    End Sub


    Private Function GetVarValue(name As String) As Object
        Dim v As Object = Nothing
        If Vars.TryGetValue(name, v) Then Return v
        Return Nothing
    End Function


    ' CONFIG PARSER ENTRYPOINT
    ' Call this from your config parsing loop when you see "TRIGGER"
    Private Sub ParseTriggerLine(parts() As String)
        EnsureTriggerEngine()

        Dim kv = ParseNamedParams(parts)

        Dim name = GetKV(kv, "name", "")
        If name = "" Then Exit Sub

        Dim period = GetKVD(kv, "period", 0.5)
        Dim ifExpr = GetKV(kv, "if", "")
        Dim thenExpr = GetKV(kv, "then", "")

        Dim cooldown = GetKVD(kv, "cooldown", 0.0)
        Dim need = GetKVI(kv, "need", 1)
        Dim oneshot = GetKVB(kv, "oneshot", False)
        Dim enabled = GetKVB(kv, "enabled", True)

        TriggerEng.AddOrUpdate(New TriggerDef With {
        .Name = name,
        .Enabled = enabled,
        .PeriodSec = period,
        .IfExpr = ifExpr,
        .ThenChain = thenExpr,
        .CooldownSec = cooldown,
        .NeedHits = need,
        .OneShot = oneshot
    })
    End Sub


    ' ACTION DISPATCHER
    ' then=fire:BtnX|run:ScriptBox|resetstats:Stats1|clearchart:ChartDMM|togglevis:ChartDMM,Hist1|runlua:ScriptName
    Private Sub ExecuteThenChain(chain As String)
        If String.IsNullOrWhiteSpace(chain) Then Exit Sub

        For Each raw In chain.Split("|"c)
            Dim a = raw.Trim()
            AppendLog($"    → {a}")
            If a = "" Then Continue For

            Dim idx = a.IndexOf(":"c)
            If idx <= 0 Then Continue For

            Dim key = a.Substring(0, idx).Trim().ToLowerInvariant()
            Dim arg = a.Substring(idx + 1).Trim()

            Select Case key
                Case "fire"
                    Dim b As Button = Nothing
                    If BtnByName.TryGetValue(arg, b) AndAlso b IsNot Nothing Then b.PerformClick()

                Case "run"
                    RunTextAreaAsSendLines(arg)

                Case "resetstats"
                    ResetStatsPanel(arg)

                Case "clearchart"
                    ClearChart(arg)

                Case "togglevis"
                    ToggleVisibility(arg)

                Case "show"
                    SetVisibility(arg, True)

                Case "hide"
                    SetVisibility(arg, False)

                Case "enabletrig"
                    If TriggerEng IsNot Nothing Then TriggerEng.SetEnabled(arg, True)

                Case "disabletrig"
                    If TriggerEng IsNot Nothing Then TriggerEng.SetEnabled(arg, False)

                Case "resettrig"
                    If TriggerEng IsNot Nothing Then TriggerEng.Reset(arg)

                Case "led"
                    SetLedFromSpec(arg)     ' arg format: LedName=ON / OFF / BAD / 1 / 0 / TRUE / FALSE

                Case "runlua"
                    Dim scriptName As String = arg

                    ' 1) Prefer config-defined LUASCRIPT
                    Dim luaText As String = ""
                    If LuaScriptsByName IsNot Nothing AndAlso LuaScriptsByName.TryGetValue(scriptName, luaText) Then
                        If String.IsNullOrWhiteSpace(luaText) Then
                            AppendLog("[LUA] RUNLUA: LUASCRIPT empty: " & scriptName)
                        Else
                            RunLua(luaText.Replace("|", vbCrLf))
                        End If
                        Continue For
                    End If

                    ' 2) Fallback: TEXTAREA by name (old behaviour)
                    luaText = GetTextAreaText(scriptName)
                    If String.IsNullOrWhiteSpace(luaText) Then
                        AppendLog("[LUA] RUNLUA: script not found (LUASCRIPT or TEXTAREA): " & scriptName)
                    Else
                        RunLua(luaText.Replace("|", vbCrLf))
                    End If
            End Select
        Next
    End Sub


    Private Sub ToggleVisibility(csv As String)
        For Each id In csv.Split(","c)
            Dim k = id.Trim()
            If k = "" Then Continue For
            Dim c As Control = Nothing
            If ControlByName.TryGetValue(k, c) AndAlso c IsNot Nothing Then
                c.Visible = Not c.Visible
                If c.Visible Then c.BringToFront()
            End If
        Next
    End Sub


    Private Sub SetVisibility(csv As String, visible As Boolean)
        For Each id In csv.Split(","c)
            Dim k = id.Trim()
            If k = "" Then Continue For
            Dim c As Control = Nothing
            If ControlByName.TryGetValue(k, c) AndAlso c IsNot Nothing Then
                c.Visible = visible
                If c.Visible Then c.BringToFront()
            End If
        Next
    End Sub


    ' REQUIRED: call the same code you already run for Action=RESETSTATS
    Private Sub ResetStatsPanel(panelName As String)
        ' Put your existing RESETSTATS logic here, or call the method you already use.
        ' Example (replace with YOUR structures):
        ' If StatsState.ContainsKey(panelName) Then
        '     StatsState(panelName).Reset()
        '     UpdateStatsPanel(panelName, Double.NaN)
        ' End If
    End Sub

    ' REQUIRED: call the same code you already run for Action=CLEARCHART
    Private Sub ClearChart(chartName As String)
        ' Put your existing CLEARCHART logic here, or call the method you already use.
    End Sub

    ' REQUIRED: run an existing textarea using your SENDLINES execution logic
    Private Sub RunTextAreaAsSendLines(textAreaName As String)
        Dim tb As TextBox = Nothing
        If Not TextAreaByName.TryGetValue(textAreaName, tb) OrElse tb Is Nothing Then Exit Sub

        ' If you already have a SENDLINES executor function, CALL IT HERE and return.
        ' Otherwise this fallback just gives you lines to process:
        For Each raw In tb.Lines
            Dim s = raw.Trim()
            If s = "" Then Continue For
            If s.StartsWith(";"c) Then Continue For

            Dim noReply As Boolean = False
            If s.StartsWith("(noreply)", StringComparison.OrdinalIgnoreCase) Then
                noReply = True
                s = s.Substring(9).Trim()
            End If

            ' Replace these two calls with YOUR existing send/query methods:
            ' If noReply Then SendToDevice(currentDev, s) Else QueryDevice(currentDev, s)
        Next
    End Sub


    ' ============================
    ' 8) MINIMAL NAMED PARAM PARSER HELPERS
    ' ============================
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

    Private Function GetKV(kv As Dictionary(Of String, String), key As String, def As String) As String
        Dim v As String = Nothing
        If kv.TryGetValue(key, v) Then Return v
        Return def
    End Function

    Private Function GetKVD(kv As Dictionary(Of String, String), key As String, def As Double) As Double
        Dim s As String = Nothing
        If kv.TryGetValue(key, s) Then
            Dim d As Double
            If Double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, d) Then Return d
        End If
        Return def
    End Function

    Private Function GetKVI(kv As Dictionary(Of String, String), key As String, def As Integer) As Integer
        Dim s As String = Nothing
        If kv.TryGetValue(key, s) Then
            Dim i As Integer
            If Integer.TryParse(s, i) Then Return i
        End If
        Return def
    End Function

    Private Function GetKVB(kv As Dictionary(Of String, String), key As String, def As Boolean) As Boolean
        Dim s As String = Nothing
        If kv.TryGetValue(key, s) Then
            If s.Equals("1") OrElse s.Equals("true", StringComparison.OrdinalIgnoreCase) OrElse s.Equals("yes", StringComparison.OrdinalIgnoreCase) Then Return True
            If s.Equals("0") OrElse s.Equals("false", StringComparison.OrdinalIgnoreCase) OrElse s.Equals("no", StringComparison.OrdinalIgnoreCase) Then Return False
        End If
        Return def
    End Function


    ' ============================
    ' 9) TRIGGER ENGINE + EXPRESSION EVALUATOR
    ' ============================
    Private Class TriggerDef
        Public Name As String
        Public Enabled As Boolean
        Public PeriodSec As Double
        Public IfExpr As String
        Public ThenChain As String
        Public CooldownSec As Double
        Public NeedHits As Integer
        Public OneShot As Boolean
    End Class


    Private Class TriggerRuntime
        Public Def As TriggerDef
        Public NextEvalUtc As DateTime
        Public Hits As Integer
        Public CooldownUntilUtc As DateTime
        Public FiredOnce As Boolean
        Public Busy As Boolean
    End Class


    Private Class TriggerEngine
        Private ReadOnly _getVar As Func(Of String, Object)
        Private ReadOnly _doActions As Action(Of String)
        Private ReadOnly _trigs As New Dictionary(Of String, TriggerRuntime)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _timer As New System.Windows.Forms.Timer()
        Private ReadOnly _log As Action(Of String)


        Public Sub Reset(triggerName As String)
            If String.IsNullOrWhiteSpace(triggerName) Then Exit Sub

            Dim rt As TriggerRuntime = Nothing
            If _trigs.TryGetValue(triggerName, rt) AndAlso rt IsNot Nothing Then
                rt.Hits = 0
                rt.CooldownUntilUtc = DateTime.MinValue
                rt.NextEvalUtc = DateTime.UtcNow
                If _log IsNot Nothing Then _log($"[TRIGGER] {triggerName} reset")
            End If
        End Sub


        Public Sub SetEnabled(triggerName As String, enabled As Boolean)
            If String.IsNullOrWhiteSpace(triggerName) Then Exit Sub

            Dim rt As TriggerRuntime = Nothing
            If _trigs.TryGetValue(triggerName, rt) AndAlso rt IsNot Nothing AndAlso rt.Def IsNot Nothing Then
                rt.Def.Enabled = enabled
                rt.Hits = 0
                rt.CooldownUntilUtc = DateTime.MinValue
                rt.NextEvalUtc = DateTime.UtcNow
            End If
        End Sub


        Public Sub New(getVar As Func(Of String, Object), doActions As Action(Of String), log As Action(Of String))

            _getVar = getVar
            _doActions = doActions
            _log = log

            _timer.Interval = 100
            AddHandler _timer.Tick, AddressOf Tick

        End Sub


        Public Sub Start()
            _timer.Start()
        End Sub

        Public Sub AddOrUpdate(def As TriggerDef)
            If def Is Nothing OrElse String.IsNullOrWhiteSpace(def.Name) Then Exit Sub

            Dim rt As TriggerRuntime = Nothing
            If Not _trigs.TryGetValue(def.Name, rt) Then
                rt = New TriggerRuntime()
                _trigs(def.Name) = rt
            End If

            rt.Def = def
            If rt.Def.NeedHits < 1 Then rt.Def.NeedHits = 1
            If rt.Def.PeriodSec <= 0 Then rt.Def.PeriodSec = 0.5
            rt.NextEvalUtc = DateTime.UtcNow
        End Sub


        Private Sub Tick(sender As Object, e As EventArgs)
            Dim now = DateTime.UtcNow

            For Each kv In _trigs
                Dim rt = kv.Value
                Dim d = rt.Def
                If d Is Nothing OrElse Not d.Enabled Then Continue For
                If rt.Busy Then Continue For
                If d.OneShot AndAlso rt.FiredOnce Then Continue For
                If now < rt.CooldownUntilUtc Then Continue For
                If now < rt.NextEvalUtc Then Continue For

                rt.NextEvalUtc = now.AddSeconds(d.PeriodSec)

                Dim ok As Boolean
                Try
                    ok = EvalBool(d.IfExpr, _getVar)
                    If _log IsNot Nothing AndAlso d.Name.Equals("TrigSettled", StringComparison.OrdinalIgnoreCase) Then
                        Dim c = _getVar("stats:Stats1.count")
                        Dim s = _getVar("stats:Stats1.std")
                        _log($"[DBG] count={If(c, "NULL")} std={If(s, "NULL")} ok={ok}")
                    End If
                Catch
                    ok = False
                End Try

                If ok Then
                    rt.Hits += 1
                Else
                    rt.Hits = 0
                End If

                If rt.Hits < d.NeedHits Then Continue For

                rt.Hits = 0
                rt.Busy = True
                Try
                    If _log IsNot Nothing Then _log($"[TRIGGER] {d.Name} fired")
                    _doActions(d.ThenChain)
                    rt.FiredOnce = True
                    If d.CooldownSec > 0 Then rt.CooldownUntilUtc = now.AddSeconds(d.CooldownSec)
                Catch ex As Exception
                    If _log IsNot Nothing Then _log($"[ERROR] {d.Name}: {ex.Message}")
                Finally
                    rt.Busy = False
                End Try

            Next
        End Sub


        ' --- Expression evaluator: numbers, variables, AND/OR/NOT, comparisons, parentheses
        Private Shared Function EvalBool(expr As String, getVar As Func(Of String, Object)) As Boolean
            If String.IsNullOrWhiteSpace(expr) Then Return False
            Dim tokens = Tokenize(expr)
            Dim rpn = ToRpn(tokens)
            Dim obj = EvalRpn(rpn, getVar)
            Return ToBool(obj)
        End Function


        Private Shared Function Tokenize(expr As String) As List(Of String)
            Dim t As New List(Of String)()
            Dim i As Integer = 0
            While i < expr.Length
                Dim c = expr(i)
                If Char.IsWhiteSpace(c) Then i += 1 : Continue While

                If c = "("c OrElse c = ")"c Then
                    t.Add(c.ToString()) : i += 1 : Continue While
                End If

                If i + 1 < expr.Length Then
                    Dim two = expr.Substring(i, 2)
                    If two = "<=" OrElse two = ">=" OrElse two = "==" OrElse two = "!=" Then
                        t.Add(two) : i += 2 : Continue While
                    End If
                End If

                If c = "<"c OrElse c = ">"c Then
                    t.Add(c.ToString()) : i += 1 : Continue While
                End If

                Dim sb As New StringBuilder()
                While i < expr.Length
                    c = expr(i)
                    If Char.IsWhiteSpace(c) Then Exit While
                    If c = "("c OrElse c = ")"c Then Exit While
                    If c = "<"c OrElse c = ">"c OrElse c = "="c OrElse c = "!"c Then Exit While
                    sb.Append(c) : i += 1
                End While

                Dim w = sb.ToString().Trim()
                If w <> "" Then
                    If w.Equals("and", StringComparison.OrdinalIgnoreCase) Then w = "AND"
                    If w.Equals("or", StringComparison.OrdinalIgnoreCase) Then w = "OR"
                    If w.Equals("not", StringComparison.OrdinalIgnoreCase) Then w = "NOT"
                    t.Add(w)
                End If
            End While
            Return t
        End Function


        Private Shared Function IsOp(tok As String) As Boolean
            Select Case tok
                Case "NOT", "AND", "OR", "<", "<=", ">", ">=", "==", "!=" : Return True
            End Select
            Return False
        End Function


        Private Shared Function Prec(op As String) As Integer
            Select Case op
                Case "NOT" : Return 3
                Case "<", "<=", ">", ">=", "==", "!=" : Return 2
                Case "AND" : Return 1
                Case "OR" : Return 0
            End Select
            Return -1
        End Function


        Private Shared Function ToRpn(tokens As List(Of String)) As List(Of String)
            Dim out As New List(Of String)()
            Dim ops As New Stack(Of String)()

            For Each tok In tokens
                If tok = "(" Then
                    ops.Push(tok)
                ElseIf tok = ")" Then
                    While ops.Count > 0 AndAlso ops.Peek() <> "("
                        out.Add(ops.Pop())
                    End While
                    If ops.Count > 0 AndAlso ops.Peek() = "(" Then ops.Pop()
                ElseIf IsOp(tok) Then
                    While ops.Count > 0 AndAlso IsOp(ops.Peek()) AndAlso Prec(ops.Peek()) >= Prec(tok)
                        out.Add(ops.Pop())
                    End While
                    ops.Push(tok)
                Else
                    out.Add(tok)
                End If
            Next

            While ops.Count > 0
                out.Add(ops.Pop())
            End While

            Return out
        End Function


        Private Shared Function EvalRpn(rpn As List(Of String), getVar As Func(Of String, Object)) As Object
            Dim st As New Stack(Of Object)()

            For Each tok In rpn
                If Not IsOp(tok) Then
                    st.Push(ParseVal(tok, getVar))
                    Continue For
                End If

                If tok = "NOT" Then
                    Dim a = If(st.Count > 0, st.Pop(), False)
                    st.Push(Not ToBool(a))
                    Continue For
                End If

                Dim b = If(st.Count > 0, st.Pop(), Nothing)
                Dim a2 = If(st.Count > 0, st.Pop(), Nothing)

                Select Case tok
                    Case "AND" : st.Push(ToBool(a2) AndAlso ToBool(b))
                    Case "OR" : st.Push(ToBool(a2) OrElse ToBool(b))
                    Case "<", "<=", ">", ">=", "==", "!=" : st.Push(Cmp(a2, b, tok))
                End Select
            Next

            If st.Count = 0 Then Return False
            Return st.Pop()
        End Function


        Private Shared Function ParseVal(tok As String, getVar As Func(Of String, Object)) As Object
            Dim d As Double
            If Double.TryParse(tok, NumberStyles.Float, CultureInfo.InvariantCulture, d) Then Return d
            Dim v = getVar(tok)
            If v IsNot Nothing Then Return v
            Return tok
        End Function


        Private Shared Function ToBool(v As Object) As Boolean
            If v Is Nothing Then Return False
            If TypeOf v Is Boolean Then Return CBool(v)

            Dim d As Double
            If Double.TryParse(v.ToString().Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, d) Then
                Return d <> 0.0
            End If

            Dim s = v.ToString().Trim()
            If s.Equals("true", StringComparison.OrdinalIgnoreCase) OrElse s.Equals("on", StringComparison.OrdinalIgnoreCase) Then Return True
            If s.Equals("false", StringComparison.OrdinalIgnoreCase) OrElse s.Equals("off", StringComparison.OrdinalIgnoreCase) Then Return False

            Return False
        End Function


        Private Shared Function Cmp(a As Object, b As Object, op As String) As Boolean
            Dim da As Double, db As Double
            If TryNum(a, da) AndAlso TryNum(b, db) Then
                Select Case op
                    Case "<" : Return da < db
                    Case "<=" : Return da <= db
                    Case ">" : Return da > db
                    Case ">=" : Return da >= db
                    Case "==" : Return da = db
                    Case "!=" : Return da <> db
                End Select
            End If

            Dim sa = If(a, "").ToString()
            Dim sb = If(b, "").ToString()
            If op = "==" Then Return sa.Equals(sb, StringComparison.OrdinalIgnoreCase)
            If op = "!=" Then Return Not sa.Equals(sb, StringComparison.OrdinalIgnoreCase)
            Return False
        End Function


        Private Shared Function TryNum(v As Object, ByRef d As Double) As Boolean
            If v Is Nothing Then Return False
            Return Double.TryParse(v.ToString().Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, d)
        End Function

    End Class


    ' UI logger: append a timestamped line to a named TEXTAREA (TextBox)
    Private Sub AppendLog(msg As String, Optional targetName As String = Nothing)
        If String.IsNullOrWhiteSpace(msg) Then Exit Sub

        Dim tb As TextBox = Nothing

        ' 1) If caller specified a target, try it first
        If Not String.IsNullOrWhiteSpace(targetName) Then
            tb = TryCast(GroupBoxCustom.Controls.Find(targetName, True).FirstOrDefault(), TextBox)
        End If

        ' 2) Fallback: first readonly multiline TextBox (log-style TEXTAREA)
        If tb Is Nothing Then
            For Each c As Control In GroupBoxCustom.Controls
                Dim t = TryCast(c, TextBox)
                If t IsNot Nothing AndAlso t.Multiline AndAlso t.ReadOnly Then
                    tb = t
                    Exit For
                End If
            Next
        End If

        If tb Is Nothing Then Exit Sub

        Dim line As String = $"{DateTime.Now:HH:mm:ss.fff}  {msg}{Environment.NewLine}"

        If tb.InvokeRequired Then
            tb.BeginInvoke(Sub() tb.AppendText(line))
        Else
            tb.AppendText(line)
        End If
    End Sub


    Private Sub FuncTrigEnableCheckbox_CheckedChanged(sender As Object, e As EventArgs)
        Dim cb = TryCast(sender, CheckBox)
        If cb Is Nothing Then Exit Sub

        ' Your Tag format is: result|FUNC|param
        Dim tagStr As String = TryCast(cb.Tag, String)
        If String.IsNullOrWhiteSpace(tagStr) Then Exit Sub

        Dim p = tagStr.Split("|"c)
        Dim funcKey As String = If(p.Length >= 2, p(1).Trim().ToUpperInvariant(), "")
        Dim param As String = If(p.Length >= 3, p(2).Trim(), "")   ' <-- trigger name

        If funcKey <> "FUNCTRIGENABLE" Then Exit Sub
        If String.IsNullOrWhiteSpace(param) Then Exit Sub

        EnsureTriggerEngine() ' make sure engine exists + started
        TriggerEng.SetEnabled(param, cb.Checked)

        AppendLog($"[TRIGGER] {param} enabled={cb.Checked}")
    End Sub


    Private Sub SetLedFromSpec(spec As String)
        If String.IsNullOrWhiteSpace(spec) Then Exit Sub

        Dim parts = spec.Split({"="c}, 2)
        If parts.Length <> 2 Then Exit Sub

        Dim ledName = parts(0).Trim()
        Dim stateText = parts(1).Trim()

        ' Find the LED panel control by Name
        Dim ledCtrl As Control = Nothing

        ' Prefer your registry if you have it
        If ControlByName IsNot Nothing AndAlso ControlByName.TryGetValue(ledName, ledCtrl) Then
            ' ok
        Else
            ' fallback: search controls
            ledCtrl = GroupBoxCustom.Controls.Find(ledName, True).FirstOrDefault()
        End If

        If ledCtrl Is Nothing Then Exit Sub

        ' Your LED is a Panel
        Dim pnl = TryCast(ledCtrl, Panel)
        If pnl Is Nothing Then Exit Sub

        SetLedStateFromText(pnl, stateText)
    End Sub






End Class
