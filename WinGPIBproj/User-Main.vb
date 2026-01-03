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
    Private AutoReadValueControl As String
    Private UserAutoBusy As Integer = 0

    ' Auto-read state - Timer 16 (dev2)
    Private AutoReadDeviceName2 As String = ""
    Private AutoReadCommand2 As String = ""
    Private AutoReadResultControl2 As String = ""
    Private AutoReadValueControl2 As String
    Private UserAutoBusy2 As Integer = 0

    Private AutoReadOverloadToken As String = ""
    Private AutoReadOverloadToken2 As String = ""

    Private UserLayoutGen As Integer = 0

    Dim intervalMs As Integer = 2000

    ' Current engineering unit (e.g. Ω, kΩ, MΩ, mA, µA) taken from the selected RADIO
    Private CurrentUserUnit As String = ""
    Private CurrentUserDp As Integer = -1

    ' Current scale factor for USER-tab numeric result (e.g. Ω → kΩ)
    Private CurrentUserScale As Double = 1.0

    ' auto-scale state for USER tab
    Private CurrentUserScaleIsAuto As Boolean = False
    Private CurrentUserRangeQuery As String = ""
    Private CurrentUserAutoMapSpec As String = ""

    ' Per-device GPIB engine selection for USER tab ("standalone" or "native")
    Private GpibEngineDev1 As String = "standalone"   ' default if not specified
    Private GpibEngineDev2 As String = "standalone"   ' default if not specified

    ' Stats panel runtime state =====
    Private StatsState As New Dictionary(Of String, RunningStatsState)(StringComparer.OrdinalIgnoreCase)
    ' Holds the per-panel UI row outputs (label -> value label)
    Private StatsRows As New Dictionary(Of String, List(Of StatsRow))(StringComparer.OrdinalIgnoreCase)
    ' Per-stats-panel font size
    Private StatsFontSize As New Dictionary(Of String, Single)(StringComparer.OrdinalIgnoreCase)
    Private StatsColLeft As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    Private StatsColRight As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

    ' History grid runtime state
    Private HistoryPpmRef As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    Private HistoryMaxRows As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    Private HistoryFormats As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    Private HistoryState As New Dictionary(Of String, RunningStatsState)(StringComparer.OrdinalIgnoreCase)
    Private HistoryLists As New Dictionary(Of String, BindingSource)(StringComparer.OrdinalIgnoreCase)
    Dim colsRaw As String = "Value,Time"   ' default if cols= not provided
    Private HistoryData As New Dictionary(Of String, BindingList(Of HistoryRow))(StringComparer.OrdinalIgnoreCase)
    Private HistorySettings As New Dictionary(Of String, HistoryGridConfig)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly HistoryGridsByTarget As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)

    ' Calc
    Private ReadOnly ResultLastValue As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase) ' resultName -> last numeric value
    Private ReadOnly CalcDefs As New Dictionary(Of String, CalcDef)(StringComparer.OrdinalIgnoreCase)        ' outResult -> calc definition
    Private ReadOnly CalcDeps As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase) ' depName -> list of outResults
    Private UserCalcDepth As Integer = 0

    ' Invisibility
    Private UiById As New Dictionary(Of String, Control)(StringComparer.OrdinalIgnoreCase)      ' Invisibility: every created control, by ID (ChartDMM, Hist1, Stats1, ...)
    Private InvisFuncToTargets As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
    Private InvisFuncDefaultVisible As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase) ' func -> default Visible state

    ' Trigger
    Private TriggerEng As TriggerEngine
    ' Auto naming for unnamed triggers
    Private TriggerAutoIndex As Integer = 0

    ' DATASOURCE registry: resultName -> device|command
    Public DataSources As New Dictionary(Of String, DataSourceDef)(StringComparer.OrdinalIgnoreCase)

    ' Used during "determine" pass so that setting Radio.Checked does NOT send SCPI
    Private UserInitSuppressSend As Boolean = False

    ' KEYPAD STATE
    Private UserKeypadTarget As Control = Nothing
    Private UserKeypadPopupForm As Form = Nothing
    Private UserKeypadPopupPanel As Panel = Nothing
    Private UserKeypadFixedPanel As Panel = Nothing
    Private UserKeypadTargetLabel As Label = Nothing
    Private UserKeypadMode As String = "fixed"  ' "fixed" or "popup"
    Private UserKeypadName As String = "Keypad1"
    Private UserKeypadCaption As String = "Keypad"
    Private UserKeypadW As Integer = 230
    Private UserKeypadH As Integer = 280
    Private UserKeypadTargetLabelOn As Boolean = True
    Private UserKeypadPopupPosValid As Boolean = False
    Private UserKeypadPopupPos As Point

    ' Config file
    Private LastUserConfigPath As String = ""
    Private UserConfig_DataSaveEnabled As Boolean = True

    ' Buffer logs until TempBox exists (because BuildCustomGuiFromText clears/rebuilds controls)
    Private ReadOnly TempBoxLogBuffer As New List(Of String)

    ' Simple popup while user config is initializing
    Private UserInitPopup As Form = Nothing
    Private HistoryGrids As New Dictionary(Of String, DataGridView)(StringComparer.OrdinalIgnoreCase)
    Private HistoryGridPopupForms As New Dictionary(Of String, Form)(StringComparer.OrdinalIgnoreCase)

    ' CHART popup support
    Private ChartSettings As New Dictionary(Of String, ChartConfig)(StringComparer.OrdinalIgnoreCase)
    Private ChartPopupForms As New Dictionary(Of String, Form)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly ChartStartTimes As New Dictionary(Of String, DateTime)









    ' CALC support
    Private Class CalcDef
        Public OutResult As String
        Public Expr As String
        Public Deps As List(Of String)
    End Class


    ' History Grid
    Private Class HistoryGridConfig
        Public Property GridName As String
        Public Property ResultTarget As String
        Public Property MaxRows As Integer
        Public Property Format As String
        Public Property Columns As List(Of String)
        Public Property PpmRef As String
        Public Property Popup As Boolean
        Public Property Width As Integer
        Public Property Height As Integer
    End Class


    Private Class ChartConfig
        Public Property ChartName As String
        Public Property ResultTarget As String
        Public Property YMin As Double?
        Public Property YMax As Double?
        Public Property XStep As Double
        Public Property MaxPoints As Integer
        Public Property AutoScaleY As Boolean
        Public Property Popup As Boolean
        Public Property Width As Integer
        Public Property Height As Integer
        Public Property InnerX As Single
        Public Property InnerY As Single
        Public Property InnerW As Single
        Public Property InnerH As Single
        Public Property labelsize As Single
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

        UserConfig_DataSaveEnabled = True

        Using dlg As New OpenFileDialog()
            dlg.Title = "Select Config File"
            dlg.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

            ' Base folder: prefer CSVfilepath.Text if it is a valid directory, else fall back to Documents\WinGPIBdata
            Dim baseDir As String = CSVfilepath.Text
            If String.IsNullOrWhiteSpace(baseDir) OrElse Not IO.Directory.Exists(baseDir) Then
                Dim documentsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                baseDir = IO.Path.Combine(documentsPath, "WinGPIBdata")
            End If

            ' Force dialog into \WinGPIBdata\Devices
            Dim devicesDir As String = IO.Path.Combine(baseDir, "Devices")
            IO.Directory.CreateDirectory(devicesDir)

            dlg.InitialDirectory = devicesDir

            If dlg.ShowDialog() <> DialogResult.OK Then Exit Sub

            ShowUserInitPopup()

            Try
                LoadCustomGuiFromFile(dlg.FileName)   ' sets LastUserConfigPath + restores textboxes

                ButtonLoadTxtRefresh.Enabled = True
                If GroupBoxCustom.Controls.Count > 1 Then   ' means UI actually built
                    LabelUSERtab1.Visible = False
                End If

            Catch ex As Exception
                MessageBox.Show("Error loading layout file: " & ex.Message)
                LabelUSERtab1.Visible = True

            Finally
                HideUserInitPopup()
            End Try

        End Using

    End Sub


    Private Sub ButtonResetTxt_Click(sender As Object, e As EventArgs) Handles ButtonResetTxt.Click

        ' If a config is currently active, save any edited textbox values first
        If Not String.IsNullOrWhiteSpace(LastUserConfigPath) AndAlso IO.File.Exists(LastUserConfigPath) Then
            SaveUserTextboxState()
        End If

        ResetUsertab()

    End Sub


    Private Sub ResetUsertab()

        ' Invalidate any in-flight async UI updates ASAP
        Threading.Interlocked.Increment(UserLayoutGen)

        ' Stop User tab dev1 and dev2 timers ASAP
        Me.Timer5.Stop()
        Me.Timer16.Stop()

        ' Reset LUA state completely
        inLuaScript = False
        luaLines.Clear()
        luaScriptName = ""
        LuaScriptsByName.Clear()

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

        ' Clear invisibility + UI maps
        UiById.Clear()
        InvisFuncToTargets.Clear()
        InvisFuncDefaultVisible.Clear()

        ' Clear Stats user vars
        StatsFontSize.Clear()
        StatsColLeft.Clear()
        StatsColRight.Clear()

        ' Remove all dynamically created controls
        GroupBoxCustom.SuspendLayout()
        Try
            For i As Integer = GroupBoxCustom.Controls.Count - 1 To 0 Step -1
                Dim c = GroupBoxCustom.Controls(i)
                If c Is LabelUSERtab1 Then Continue For
                GroupBoxCustom.Controls.RemoveAt(i)
                c.Dispose()
            Next
        Finally
            GroupBoxCustom.ResumeLayout()
        End Try

        ' Reset the title text
        GroupBoxCustom.Text = "User Defineable"

        If LabelUSERtab1 IsNot Nothing AndAlso Not LabelUSERtab1.IsDisposed Then

            ' Make absolutely sure it is parented correctly
            If LabelUSERtab1.Parent IsNot GroupBoxCustom Then
                LabelUSERtab1.Parent = GroupBoxCustom
            End If

            LabelUSERtab1.Visible = True      ' <<--- THIS guarantees it shows
            LabelUSERtab1.BringToFront()

        Else
            MessageBox.Show("LabelUSERtab1 is missing or disposed!", "Debug")
        End If

        ' Clear DATASOURCE + linked-control mappings
        DataSources.Clear()
        HistoryGridsByTarget.Clear()

        ' If you implemented the multi-job schedulers
        AutoJobs5.Clear()
        AutoJobs16.Clear()

        ' Calc
        ResultLastValue.Clear()
        CalcDefs.Clear()
        CalcDeps.Clear()
        ' Auto-scale range cache
        UserRangeQueryCache.Clear()

        ' Trigger
        If TriggerEng IsNot Nothing Then TriggerEng.ClearAll()
        TriggerAutoIndex = 0

        ' Keypad reset
        UserKeypadTarget = Nothing
        If UserKeypadPopupForm IsNot Nothing Then
            UserKeypadPopupForm.Close()
            UserKeypadPopupForm.Dispose()
            UserKeypadPopupForm = Nothing
        End If
        UserKeypadPopupPanel = Nothing
        UserKeypadFixedPanel = Nothing
        UserKeypadTargetLabel = Nothing
        UserKeypadMode = "fixed"
        ' UserKeypadPopupPosValid = False

        ' Close any popup HISTORYGRID windows
        If HistoryGridPopupForms IsNot Nothing Then
            Dim forms() As Form = HistoryGridPopupForms.Values.ToArray()
            HistoryGridPopupForms.Clear()

            For Each f As Form In forms
                If f IsNot Nothing AndAlso Not f.IsDisposed Then
                    f.Close()
                    f.Dispose()
                End If
            Next
        End If

        ' Close any popup CHART windows
        If ChartPopupForms IsNot Nothing Then
            Dim forms() As Form = ChartPopupForms.Values.ToArray()
            ChartPopupForms.Clear()

            For Each f As Form In forms
                If f IsNot Nothing AndAlso Not f.IsDisposed Then
                    f.Close()
                    f.Dispose()
                End If
            Next
        End If

        ' After a reset, REFRESH is no longer valid until a load succeeds again
        ButtonLoadTxtRefresh.Enabled = False
        'LastUserConfigPath = ""

    End Sub


    Private Sub ButtonLoadTxtRefresh_Click(sender As Object, e As EventArgs) Handles ButtonLoadTxtRefresh.Click

        SaveUserTextboxState()
        ResetUsertab()
        LoadCustomGuiFromFile(LastUserConfigPath)
        RestoreUserTextboxState()

        If String.IsNullOrWhiteSpace(LastUserConfigPath) OrElse Not IO.File.Exists(LastUserConfigPath) Then
            MessageBox.Show("No previously loaded config file to refresh.")
            ButtonLoadTxtRefresh.Enabled = False
            Exit Sub
        End If

        ButtonLoadTxtRefresh.Enabled = False

        Try
            ' 1) Save textbox state before destroying controls
            SaveUserTextboxState()

            ' 2) Reset user tab (dispose dynamic controls, stop timers, clear dictionaries)
            ResetUsertab()

            ' 3) Reload config
            LoadCustomGuiFromFile(LastUserConfigPath)

            ' 4) Restore textbox values for this config
            RestoreUserTextboxState()

            LabelUSERtab1.Visible = False

        Catch ex As Exception
            MessageBox.Show("REFRESH failed:" & vbCrLf & ex.ToString())
            LabelUSERtab1.Visible = True
            LabelUSERtab1.BringToFront()
        Finally
            ' Re-enable if file still exists
            ButtonLoadTxtRefresh.Enabled = IO.File.Exists(LastUserConfigPath)
        End Try

    End Sub


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


    Private Sub LoadCustomGuiFromFile(path As String)

        UserConfig_DataSaveEnabled = True   ' reset to default for this config

        If Not IO.File.Exists(path) Then
            MessageBox.Show("Custom GUI file not found: " & path)
            Exit Sub
        End If

        LastUserConfigPath = path   ' <-- key fix

        Dim text = IO.File.ReadAllText(path)
        BuildCustomGuiFromText(text)

        RestoreUserTextboxState()   ' <-- auto-populate after build

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

        ' OVERLOADTOKEN
        Dim overloadToken As String = If(parts.Length > 5, parts(5), "")
        If Not String.IsNullOrWhiteSpace(overloadToken) Then
            overloadToken = overloadToken.Split(";"c)(0).Trim()
        End If

        ' INVISIBILITY: your hide buttons store the function name in resultControlName
        If Not String.IsNullOrWhiteSpace(resultControlName) Then
            Dim fn = resultControlName.Split(";"c)(0).Trim()
            If TryRunInvisibilityFunction(fn) Then Exit Sub
        End If

        ' INVISIBILITY / Function buttons (no device needed)
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
                ' Single-shot read via button.
                ' For native engine, explicitly request RAW numeric and feed it into RunQueryToResult
                ' so we bypass any oddities in the generic GET RAW RESPONSE path.

                ' NEW: pick up overload token from DATASOURCE (if any)
                Dim ds As DataSourceDef = Nothing
                Dim olTok As String = Nothing

                If DataSources IsNot Nothing AndAlso
       DataSources.TryGetValue(resultControlName, ds) AndAlso ds IsNot Nothing Then
                    olTok = ds.OverloadToken     ' e.g. "9.90000000E+37|raw"
                Else
                    ' Fallback to whatever was parsed into overloadToken (if you have that in the tag)
                    olTok = overloadToken
                End If

                If IsNativeEngine(deviceName) Then
                    Try
                        ' Always ask native engine for RAW numeric
                        Dim raw As String = NativeQuery(deviceName, commandOrPrefix, requireRaw:=True)
                        ' rawOverride ensures RunQueryToResult does not issue a second query
                        RunQueryToResult(deviceName, commandOrPrefix, resultControlName, raw, olTok)
                    Catch ex As Exception
                        MessageBox.Show("Error in QUERY (native): " & ex.Message)
                    End Try
                Else
                    ' Non-native: fall back to existing pipeline
                    RunQueryToResult(deviceName, commandOrPrefix, resultControlName, Nothing, olTok)
                End If
                Return



            Case "SENDLINES"

                Dim valCtrl = TryCast(GetControlByName(valueControlName), TextBox)
                If valCtrl Is Nothing Then
                    MessageBox.Show("TEXTAREA not found: " & valueControlName)
                    Exit Sub
                End If

                ' ================================
                ' NATIVE engine → use NativeSend()
                ' ================================
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
                ' deviceName is the chart name from the config, e.g. "ChartDMM"
                Dim chartName As String = deviceName

                ' 1) Clear the main in-panel chart (if present)
                Dim chMain As DataVisualization.Charting.Chart =
                TryCast(GetControlByName(chartName), DataVisualization.Charting.Chart)

                If chMain IsNot Nothing AndAlso chMain.Series.Count > 0 Then
                    For Each s In chMain.Series
                        s.Points.Clear()
                    Next
                End If

                ' 2) Clear the popup chart (if a popup window exists)
                Dim popupForm As Form = Nothing
                If ChartPopupForms IsNot Nothing AndAlso
               ChartPopupForms.TryGetValue(chartName, popupForm) AndAlso
               popupForm IsNot Nothing AndAlso Not popupForm.IsDisposed Then

                    Dim chPopup As DataVisualization.Charting.Chart = Nothing
                    For Each ctrl As Control In popupForm.Controls
                        chPopup = TryCast(ctrl, DataVisualization.Charting.Chart)
                        If chPopup IsNot Nothing Then Exit For
                    Next

                    If chPopup IsNot Nothing AndAlso chPopup.Series.Count > 0 Then
                        For Each s In chPopup.Series
                            s.Points.Clear()
                        Next
                    End If
                End If

                Exit Sub


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
        If UserInitSuppressSend Then Exit Sub

        Dim b As Button = DirectCast(sender, Button)

        Dim tagStr As String = TryCast(b.Tag, String)
        If String.IsNullOrEmpty(tagStr) Then Exit Sub

        Dim tagParts = tagStr.Split("|"c)
        If tagParts.Length < 4 Then Exit Sub

        Dim device As String = tagParts(0)
        Dim cmdOn As String = tagParts(1)
        Dim cmdOff As String = tagParts(2)

        Dim state As Integer = 0
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

        Dim newState As Integer = If(state = 0, 1, 0)

        If newState = 1 Then
            If useNative Then
                NativeSend(device, cmdOn & TermStr2())
            Else
                dev.SendAsync(cmdOn & TermStr2(), True)
            End If
        Else
            If useNative Then
                NativeSend(device, cmdOff & TermStr2())
            Else
                dev.SendAsync(cmdOff & TermStr2(), True)
            End If
        End If

        ' Update Tag with new state (PRESERVE any DETQ/DETE/DETF tokens)
        tagParts(3) = newState.ToString(Globalization.CultureInfo.InvariantCulture)
        b.Tag = String.Join("|", tagParts)

        ' Apply illumination
        ApplyToggleVisual(b, newState)
    End Sub


    Private Sub Dropdown_SelectedIndexChanged(sender As Object, e As EventArgs)

        If UserInitSuppressSend Then Exit Sub

        Dim cb = TryCast(sender, ComboBox)
        If cb Is Nothing Then Exit Sub

        Dim meta As String = TryCast(cb.Tag, String)
        If String.IsNullOrEmpty(meta) Then Exit Sub

        Dim parts = meta.Split("|"c)
        If parts.Length < 2 Then Exit Sub

        Dim deviceName As String = parts(0).Trim()

        ' First entry (index 0) is always a placeholder → do nothing
        If cb.SelectedIndex <= 0 Then Exit Sub

        Dim selected As String = ""
        If cb.SelectedItem IsNot Nothing Then selected = cb.SelectedItem.ToString().Trim()
        If selected = "" Then Exit Sub

        Dim cmd As String = ""

        ' =========================================================
        ' commands= list mode
        ' Tag format: dev|CMDLIST|cmd1,cmd2,cmd3|...
        ' =========================================================
        If String.Equals(parts(1).Trim(), "CMDLIST", StringComparison.OrdinalIgnoreCase) Then

            If parts.Length < 3 Then Exit Sub

            Dim listRaw As String = parts(2)

            Dim cmds() As String =
            listRaw.Split(","c).
            Select(Function(s) s.Trim()).
            Where(Function(s) s <> "").
            ToArray()

            Dim idx As Integer = cb.SelectedIndex - 1   ' because index 0 is placeholder
            If idx < 0 OrElse idx >= cmds.Length Then Exit Sub

            cmd = cmds(idx)

        Else
            ' =========================================================
            ' Existing command-prefix mode
            ' Tag format: dev|prefix|...
            ' =========================================================
            Dim commandPrefix As String = parts(1).TrimEnd()

            If commandPrefix.EndsWith(" ") OrElse commandPrefix.EndsWith(vbTab) Then
                cmd = commandPrefix & selected
            Else
                cmd = commandPrefix & " " & selected
            End If
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
                MessageBox.Show("Unknown device in DROPDOWN: " & deviceName)
                Exit Sub
        End Select

        If dev Is Nothing Then
            MessageBox.Show("Device not available: " & deviceName)
            Exit Sub
        End If

        If useNative Then
            NativeSend(deviceName, cmd)
        Else
            dev.SendAsync(cmd, True)
        End If

    End Sub


    Private Sub Radio_CheckedChanged(sender As Object, e As EventArgs)
        Dim rb As RadioButton = TryCast(sender, RadioButton)
        If rb Is Nothing Then Exit Sub
        If Not rb.Checked Then Exit Sub

        If UserInitSuppressSend Then Exit Sub

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
        '   DP + MAP from Tag (all radios)
        ' ===============================
        CurrentUserDp = -1
        CurrentUserAutoMapSpec = ""

        Dim tp() As String = meta.Split("|"c)
        For Each t In tp
            Dim tt = t.Trim()

            If tt.StartsWith("DP=", StringComparison.OrdinalIgnoreCase) Then
                Dim dpn As Integer
                If Integer.TryParse(tt.Substring(3), dpn) Then
                    CurrentUserDp = dpn
                End If

            ElseIf tt.StartsWith("MAP=", StringComparison.OrdinalIgnoreCase) Then
                CurrentUserAutoMapSpec = tt.Substring(4).Trim()
            End If
        Next

        ' ===============================
        '   MODE vs RANGE DETECTION
        ' ===============================
        Dim cap As String = rb.Text.Trim()
        Dim lastSpace As Integer = cap.LastIndexOf(" "c)
        Dim isModeRadio As Boolean = (lastSpace < 0)

        If isModeRadio Then
            ' A mode change (DCV/ACV/etc) – skip next display update
            UserSkipNextUserDisplay = True

            ' When changing mode (DCV/ACV etc), clear unit then
            ' re-sync from whichever RANGE radio is already selected
            CurrentUserUnit = ""
            CurrentUserDp = -1

            ' Uncheck only other radios in the SAME group box as this mode radio
            Dim parentGb As GroupBox = TryCast(rb.Parent, GroupBox)
            If parentGb IsNot Nothing Then
                For Each child As Control In parentGb.Controls
                    Dim r As RadioButton = TryCast(child, RadioButton)
                    If r Is Nothing OrElse r Is rb Then Continue For
                    r.Checked = False
                Next
            End If

            ' Rebuild CurrentUserUnit / scale from the currently checked RANGE radio
            SyncUserUnitsAndScaleFromCheckedRadios()

        Else
            ' A range change – skip next display update
            UserSkipNextUserDisplay = True

            ' RANGE radios: unit comes from the tail of the caption ("10 V" → "V")
            Dim unit As String = ""
            If lastSpace >= 0 AndAlso lastSpace < cap.Length - 1 Then
                unit = cap.Substring(lastSpace + 1).Trim()
            End If
            CurrentUserUnit = unit
        End If

        ' ===============================
        '   SEND THE RADIO COMMAND
        ' ===============================

        If UserInitSuppressSend Then Exit Sub

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

        ' Allow multiple commands separated by commas, e.g.
        '   command=*CLS,:CONF:VOLT:DC
        Dim cmdParts As String() = command.Split(","c)

        For Each part As String In cmdParts
            Dim singleCmd As String = part.Trim()
            If singleCmd = "" Then Continue For

            Try
                If useNative Then
                    ' Same path as SEND buttons for native engine
                    NativeSend(deviceName, singleCmd)
                Else
                    ' Standalone engine path
                    dev.SendAsync(singleCmd, True)
                End If
            Catch ex As Exception
                AppendLog($"[RADIO] Error sending '{singleCmd}' to {deviceName}: {ex.Message}")
            End Try
        Next

    End Sub


    Private Sub Slider_Scroll(sender As Object, e As EventArgs)       ' handler

        If UserInitSuppressSend Then Exit Sub

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


    Private Sub Slider_MouseUp(sender As Object, e As MouseEventArgs)       ' handler

        If UserInitSuppressSend Then Exit Sub

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


    Private Sub Spinner_ValueChanged(sender As Object, e As EventArgs)

        If UserInitSuppressSend Then Exit Sub

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

        ' Publish LED state for triggers:
        ' Vars("led:LedMAV") and Vars("LedMAV") -> 1/0
        Dim ledName As String = led.Name
        Dim isOn As Integer = 0
        If newColor = onColor Then isOn = 1 Else isOn = 0

        Vars($"led:{ledName}") = isOn
        Vars(ledName) = isOn
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







    ' Datasource
    Private Class AutoJob
        Public Device As String
        Public Command As String
        Public Result As String
        Public IntervalMs As Integer
        Public NextDue As Integer
        Public InFlight As Boolean
        Public OverloadToken As String = ""
    End Class


    ' Separate job sets (Timer5 vs Timer16)
    Private ReadOnly AutoJobs5 As New Dictionary(Of String, AutoJob)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly AutoJobs16 As New Dictionary(Of String, AutoJob)(StringComparer.OrdinalIgnoreCase)


    ' TickCount helpers (wrap-safe)
    Private Function NowTick() As Integer
        Return Environment.TickCount
    End Function


    Private Function Due(nowTick As Integer, dueTick As Integer) As Boolean
        Return CInt(nowTick - dueTick) >= 0
    End Function


    Private Sub FuncAutoCheckbox_CheckedChanged(sender As Object, e As EventArgs)
        Dim cb = TryCast(sender, CheckBox)
        If cb Is Nothing Then Exit Sub

        Dim tagStr = TryCast(cb.Tag, String)
        If String.IsNullOrEmpty(tagStr) Then Exit Sub

        Dim parts = tagStr.Split("|"c)
        If parts.Length < 2 Then Exit Sub

        ' Allow multiple results in the first Tag field, comma-separated
        Dim resultNames As String() =
        parts(0).Split(","c).
        Select(Function(s) s.Trim()).
        Where(Function(s) s <> "").
        ToArray()

        If resultNames.Length = 0 Then Exit Sub

        ' Compute intervalMs from FuncAuto param (seconds) in Tag part[2]
        Dim intervalMs As Integer = 2000
        If parts.Length >= 3 Then
            Dim paramStr As String = parts(2).Trim()
            Dim numeric As String = ""
            For Each ch As Char In paramStr
                If Char.IsDigit(ch) OrElse ch = "."c Then numeric &= ch
            Next
            Dim secs As Double
            If Double.TryParse(numeric, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, secs) AndAlso secs > 0.0R Then
                intervalMs = CInt(Math.Ceiling(secs * 1000.0R))
            End If
        End If
        If intervalMs < 1 Then intervalMs = 1
        If intervalMs > 60000 Then intervalMs = 60000

        ' Apply to each result in the list
        For Each resultName As String In resultNames

            ' Determine producer: DATASOURCE preferred, else steal from QUERY button (backwards compat)
            Dim dev As String = ""
            Dim cmd As String = ""

            Dim overloadToken As String = ""
            If DataSources.ContainsKey(resultName) Then
                Dim ds As DataSourceDef = DataSources(resultName)
                dev = ds.Device
                cmd = ds.Command
                overloadToken = If(ds.OverloadToken, "").Trim()
            Else
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
                        dev = deviceName
                        cmd = commandOrPrefix
                        Exit For
                    End If
                Next
            End If

            ' Choose which timer/dictionary this result belongs to
            Dim jobs As Dictionary(Of String, AutoJob) = AutoJobs5
            If dev.Equals("dev2", StringComparison.OrdinalIgnoreCase) Then jobs = AutoJobs16

            ' Unchecked: remove this job from the correct scheduler
            If Not cb.Checked Then
                jobs.Remove(resultName)

                If AutoJobs5.Count = 0 Then Timer5.Enabled = False
                If AutoJobs16.Count = 0 Then Timer16.Enabled = False

                Continue For
            End If

            If String.IsNullOrWhiteSpace(dev) OrElse String.IsNullOrWhiteSpace(cmd) Then Continue For

            ' Immediate one-shot read on enable
            SyncUserUnitsAndScaleFromCheckedRadios()
            RunQueryToResult(dev, cmd, resultName, Nothing, overloadToken)

            ' Add/update job
            Dim j As AutoJob = Nothing
            If Not jobs.TryGetValue(resultName, j) Then
                j = New AutoJob()
                jobs(resultName) = j
            End If

            Dim nowT As Integer = NowTick()
            j.Device = dev
            j.Command = cmd
            j.Result = resultName
            j.IntervalMs = intervalMs
            j.NextDue = nowT + 1
            j.InFlight = False
            j.OverloadToken = overloadToken

            ' Enable correct scheduler timer (fixed scheduler tick)
            If jobs Is AutoJobs16 Then
                Timer16.Interval = 50
                Timer16.Enabled = True
            Else
                Timer5.Interval = 50
                Timer5.Enabled = True
            End If

        Next

    End Sub


    ' Try to apply instrument-specific AUTO range mapping from CurrentUserAutoMapSpec.
    ' mapSpec format: "<rangeReply>:<unit>:<dp>,<rangeReply>:<unit>:<dp>,..."
    ' Returns True if a mapping entry matched and scaledVal/unit/dp were updated.
    Private Function TryApplyAutoRangeMap(rangeVal As Double,
                                          originalVal As Double,
                                          ByRef scaledVal As Double) As Boolean


        'System.Diagnostics.Debug.WriteLine(
        '"USER-AUTO-MAP-CALL: rangeVal=" &
        'rangeVal.ToString("G", Globalization.CultureInfo.InvariantCulture) &
        '" mapSpec='" & CurrentUserAutoMapSpec & "'")


        Dim spec As String = CurrentUserAutoMapSpec
        If String.IsNullOrWhiteSpace(spec) Then
            Return False
        End If

        Dim a As Double = Math.Abs(rangeVal)
        If Double.IsNaN(a) OrElse Double.IsInfinity(a) Then
            Return False
        End If

        Dim entries() As String = spec.Split(","c)
        For Each entry In entries
            Dim e = entry.Trim()
            If e = "" Then Continue For

            Dim parts() As String = e.Split(":"c)
            If parts.Length < 3 Then Continue For

            Dim rangeMatch As Double
            If Not Double.TryParse(parts(0).Trim(),
                                   Globalization.NumberStyles.Float,
                                   Globalization.CultureInfo.InvariantCulture,
                                   rangeMatch) Then
                Continue For
            End If

            ' Tolerant numeric compare
            Dim targetAbs As Double = Math.Abs(rangeMatch)
            Dim tol As Double = Math.Max(0.000000001, targetAbs * 0.000001)
            If Math.Abs(a - targetAbs) > tol Then
                Continue For
            End If

            Dim unit As String = parts(1).Trim()
            Dim dpVal As Integer
            Integer.TryParse(parts(2).Trim(), dpVal)

            If unit <> "" Then
                CurrentUserUnit = unit
            End If
            If dpVal >= 0 Then
                CurrentUserDp = dpVal
            End If

            ' Derive scale factor from engineering prefix in unit
            Dim scaleFactor As Double = 1.0
            If unit.Length > 0 Then
                Dim prefix As Char = unit(0)
                Dim exp As Integer = 0

                Select Case prefix
                    Case "f"c
                        exp = -15
                    Case "p"c
                        exp = -12
                    Case "n"c
                        exp = -9
                    Case "u"c, "µ"c, "μ"c
                        exp = -6
                    Case "m"c
                        exp = -3
                    Case "k"c, "K"c
                        exp = 3
                    Case "M"c
                        exp = 6
                    Case "G"c
                        exp = 9
                    Case Else
                        exp = 0
                End Select

                If exp <> 0 Then
                    scaleFactor = Math.Pow(10.0, -exp)
                End If
            End If

            scaledVal = originalVal * scaleFactor


            'System.Diagnostics.Debug.WriteLine(
            '"USER-AUTO-MAP-HARDTEST-HIT: rangeVal=" &
            'rangeVal.ToString("G", Globalization.CultureInfo.InvariantCulture) &
            '" unit='mV' dp=3 scaleFactor=1000")


            '            System.Diagnostics.Debug.WriteLine(
            '    "USER-AUTO-MAP-HIT: rangeVal=" &
            '            rangeVal.ToString("G", Globalization.CultureInfo.InvariantCulture) &
            '    " unit='" & unit & "'" &
            '    " dp=" & dpVal &
            '    " scaleFactor=" & scaleFactor.ToString("G", Globalization.CultureInfo.InvariantCulture))

            'System.Diagnostics.Debug.WriteLine(
            '"USER-AUTO-MAP-HARDTEST-HIT: rangeVal=" &
            'rangeVal.ToString("G", Globalization.CultureInfo.InvariantCulture) &
            '" unit='mV' dp=3 scaleFactor=1000")

            Return True
        Next

        Return False
    End Function


    ' Compute a display scale factor from a numeric range value
    ' Generic engineering-prefix style:
    '   very small ranges -> n / µ / m
    '   mid ranges        -> base units
    '   large ranges      -> k / M

    '   small ranges      -> m / µ / n
    '   large ranges      -> k / M
    Private Function ComputeAutoScaleFromRange(rangeVal As Double) As Double
        Dim a As Double = Math.Abs(rangeVal)
        Dim factor As Double

        If a <= 0 OrElse Double.IsNaN(a) OrElse Double.IsInfinity(a) Then
            factor = 1.0

            ' Below 1 µ → n-units
        ElseIf a < 0.000001 Then
            factor = 1000000000.0     ' n

            ' 1 µ .. < 1 m → µ-units
        ElseIf a < 0.001 Then
            factor = 1000000.0        ' µ

            ' 1 m .. < 1 → m-units
        ElseIf a < 1.0 Then
            factor = 1000.0           ' m

            ' 1 .. < 1 k → base units
        ElseIf a < 1000.0 Then
            factor = 1.0              ' base

            ' 1 k .. < 1 M → k-units
        ElseIf a < 1000000.0 Then
            factor = 0.001            ' k

            ' >= 1 M → M-units
        Else
            factor = 0.000001         ' M
        End If

        ' DEBUG: see what auto-scale is doing for each range
        'Debug.WriteLine(
        '$"[AutoScale] rangeVal={rangeVal}, abs={a}, factor={factor}")

        Return factor
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
                raw = respNorm
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


    ' One output row in a stats panel
    Private Class StatsRow
        Public LabelText As String
        Public FuncKey As String
        Public RefToken As String
        Public ValueLabel As Label
        Public FormatOverride As String   ' "", "F", or "E"
    End Class


    ' Running stats (fast, stable)
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

        ' First, try the existing token-based method
        Dim tokens = text.Split({","c, ";"c, " "c, ControlChars.Cr, ControlChars.Lf},
                            StringSplitOptions.RemoveEmptyEntries)

        For Each tok In tokens
            Dim t = tok.Trim()
            If Double.TryParse(t,
                           Globalization.NumberStyles.Float,
                           Globalization.CultureInfo.InvariantCulture,
                           value) Then
                If Not Double.IsNaN(value) AndAlso Not Double.IsInfinity(value) Then Return True
            End If
        Next

        ' Fallback: regex – finds first number even if glued to units, e.g. "+0.99998801E+00VDC"
        Dim m = System.Text.RegularExpressions.Regex.Match(
                text,
                "[+-]?\d+(\.\d+)?([Ee][+-]?\d+)?",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant)

        If m.Success Then
            Dim s = m.Value
            If Double.TryParse(s,
                           Globalization.NumberStyles.Float,
                           Globalization.CultureInfo.InvariantCulture,
                           value) Then
                If Not Double.IsNaN(value) AndAlso Not Double.IsInfinity(value) Then Return True
            End If
        End If

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

        ' Local resolver: supports numeric OR "@ControlName"
        Dim ResolveRefToken As Func(Of String, String) =
        Function(tok As String) As String
            If String.IsNullOrWhiteSpace(tok) Then Return ""

            tok = tok.Trim()

            If tok.StartsWith("@") Then
                Dim ctrlName As String = tok.Substring(1).Trim()
                If ctrlName = "" Then Return ""

                Dim tb = TryCast(GetControlByName(ctrlName), TextBox)
                If tb Is Nothing Then Return ""

                Return tb.Text.Trim()
            End If

            Return tok
        End Function

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
                    Dim refResolved As String = ResolveRefToken(r.RefToken)
                    outVal = ComputePpmString(st, refResolved, panelFmt)

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


    Private Sub RegisterDynamicControl(id As String, ctrl As Control)
        If String.IsNullOrWhiteSpace(id) Then Return
        If ctrl Is Nothing Then Return
        UiById(id.Trim()) = ctrl
    End Sub


    Private Function TryRunInvisibilityFunction(funcName As String) As Boolean
        If String.IsNullOrWhiteSpace(funcName) Then Return False

        Dim targets As List(Of String) = Nothing
        If Not InvisFuncToTargets.TryGetValue(funcName.Trim(), targets) Then Return False

        Dim poppedHist As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim poppedCharts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        ' First pass: handle pop-out targets (history grids and charts)
        For Each id In targets
            ' HISTORYGRID popup?
            Dim hCfg As HistoryGridConfig = Nothing
            If HistorySettings.TryGetValue(id, hCfg) _
           AndAlso hCfg IsNot Nothing _
           AndAlso hCfg.Popup Then

                PopOutHistoryGrid(id)
                poppedHist.Add(id)
                Continue For
            End If

            ' CHART popup?
            Dim cCfg As ChartConfig = Nothing
            If ChartSettings.TryGetValue(id, cCfg) _
           AndAlso cCfg IsNot Nothing _
           AndAlso cCfg.Popup Then

                PopOutChart(id)
                poppedCharts.Add(id)
            End If
        Next

        ' Second pass: normal visibility toggle for non-popup targets
        For Each id In targets
            If poppedHist.Contains(id) OrElse poppedCharts.Contains(id) Then Continue For

            Dim c As Control = Nothing
            If UiById.TryGetValue(id, c) AndAlso c IsNot Nothing Then
                c.Visible = Not c.Visible
                If c.Visible Then c.BringToFront()
            End If
        Next

        Return True
    End Function


    Private Sub PopOutHistoryGrid(gridName As String)
        Dim cfg As HistoryGridConfig = Nothing
        If Not HistorySettings.TryGetValue(gridName, cfg) _
       OrElse cfg Is Nothing _
       OrElse Not cfg.Popup Then

            Exit Sub
        End If

        ' If there's already a popup, just bring it to front
        Dim existingForm As Form = Nothing
        If HistoryGridPopupForms.TryGetValue(gridName, existingForm) _
       AndAlso existingForm IsNot Nothing _
       AndAlso Not existingForm.IsDisposed Then

            existingForm.BringToFront()
            existingForm.Activate()
            Exit Sub
        End If

        ' Get (or create) the shared history list
        Dim list As BindingList(Of HistoryRow) = Nothing
        If Not HistoryData.TryGetValue(gridName, list) Then
            list = New BindingList(Of HistoryRow)()
            HistoryData(gridName) = list
        End If

        ' Build popup form
        Dim f As New Form()
        f.Text = If(String.IsNullOrEmpty(cfg.GridName), gridName, cfg.GridName)
        f.StartPosition = FormStartPosition.CenterParent
        Dim baseW As Integer = If(cfg.Width > 0, cfg.Width, 600)
        Dim baseH As Integer = If(cfg.Height > 0, cfg.Height, 300)
        f.Size = New Size(baseW + 40, baseH + 80)   ' a bit of padding for borders/title

        f.FormBorderStyle = FormBorderStyle.FixedDialog
        f.MaximizeBox = False
        f.MinimizeBox = False
        f.ShowInTaskbar = False   ' optional
        f.ControlBox = True       ' keep close button

        Dim dgv As New DataGridView()
        dgv.Dock = DockStyle.Fill
        dgv.Name = gridName & "_popup"

        dgv.AllowUserToAddRows = False
        dgv.AllowUserToDeleteRows = False
        dgv.ReadOnly = True
        dgv.RowHeadersVisible = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.MultiSelect = False
        dgv.AutoGenerateColumns = False
        dgv.ScrollBars = ScrollBars.Vertical
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        ' Double-buffer to reduce flicker
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

        ' Build columns from cfg.Columns
        Dim colNames As List(Of String) = cfg.Columns
        If colNames Is Nothing OrElse colNames.Count = 0 Then
            colNames = New List(Of String) From {"Value", "Time"}
        End If

        dgv.Columns.Clear()

        For Each colName As String In colNames
            Dim c As New DataGridViewTextBoxColumn()
            c.Name = colName
            c.HeaderText = colName
            c.DataPropertyName = colName   ' must match HistoryRow properties
            c.SortMode = DataGridViewColumnSortMode.NotSortable

            c.Width = 70
            If colName.Equals("Value", StringComparison.OrdinalIgnoreCase) Then c.Width = 95
            If colName.Equals("Time", StringComparison.OrdinalIgnoreCase) Then c.Width = 120
            If colName.Equals("PkPk", StringComparison.OrdinalIgnoreCase) Then c.Width = 90
            If colName.Equals("Mean", StringComparison.OrdinalIgnoreCase) Then c.Width = 95

            If colName.Equals("Time", StringComparison.OrdinalIgnoreCase) Then
                c.DefaultCellStyle.Format = "HH:mm:ss"
            ElseIf Not colName.Equals("Count", StringComparison.OrdinalIgnoreCase) Then
                c.DefaultCellStyle.Format = cfg.Format
            End If

            dgv.Columns.Add(c)
        Next

        If dgv.Columns.Count > 0 Then
            dgv.Columns(0).DefaultCellStyle.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        End If

        ' Bind to the SAME data list the history engine uses
        dgv.DataSource = list

        f.Controls.Add(dgv)

        HistoryGridPopupForms(gridName) = f

        AddHandler f.FormClosing,
        Sub(sender As Object, args As FormClosingEventArgs)
            HistoryGridPopupForms.Remove(gridName)
        End Sub

        f.Show(Me)
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
        UiById(controlName) = ctrl      ' makes INVISIBILITY work
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
        If String.IsNullOrWhiteSpace(name) Then Return Nothing

        Dim key = name.Trim()

        Dim v As Object = Nothing

        ' 1) Exact match
        If Vars.TryGetValue(key, v) Then Return v

        ' 2) Common result streams
        If Vars.TryGetValue($"num:{key}", v) Then Return v
        If Vars.TryGetValue($"bignum:{key}", v) Then Return v

        ' 3) LED stream
        If Vars.TryGetValue($"led:{key}", v) Then Return v

        ' 4) Stats short form: Stats1.mean -> stats:Stats1.mean
        If key.Contains("."c) Then
            If Vars.TryGetValue($"stats:{key}", v) Then Return v
        End If

        Return Nothing
    End Function


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

                Case "toggle"
                    ToggleVisibility(arg)

                Case "set"
                    SetResultFromSpec(arg)   ' arg: ResultName=123.4

                Case "send"
                    SendFromSpec(arg)        ' arg: dev1:*CLS

                Case "query"
                    QueryFromSpec(arg)       ' arg: dev1:MEAS?->YourDMM

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


    ' Run an existing textarea using your SENDLINES execution logic
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


    ' TRIGGER ENGINE + EXPRESSION EVALUATOR
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


        ' Expression evaluator: numbers, variables, AND/OR/NOT, comparisons, parentheses
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


        Public Sub ClearAll()
            _trigs.Clear()
        End Sub

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


    ' Called by RunQueryToResult() when DATASOURCE feeds a result name
    Private Sub AddHistorySample(gridName As String, dv As Double)

        Dim list As BindingList(Of HistoryRow) = Nothing
        If Not HistoryData.TryGetValue(gridName, list) Then Exit Sub

        Dim st As RunningStatsState = Nothing
        If Not HistoryState.TryGetValue(gridName, st) Then Exit Sub

        Dim cfg As HistoryGridConfig = Nothing
        If Not HistorySettings.TryGetValue(gridName, cfg) Then Exit Sub

        st.AddSample(dv)

        Dim row As New HistoryRow()
        row.Time = DateTime.Now
        row.Value = dv
        row.Min = If(st.Count > 0, st.Min, 0)
        row.Max = If(st.Count > 0, st.Max, 0)
        row.PkPk = If(st.Count > 0, st.Max - st.Min, 0)
        row.Mean = If(st.Count > 0, st.Mean, 0)
        row.Std = If(st.Count > 1, st.StdDevSample(), 0)

        Dim refVal As Double = st.Mean
        If cfg.PpmRef IsNot Nothing AndAlso cfg.PpmRef.Equals("FIRST", StringComparison.OrdinalIgnoreCase) AndAlso st.HasFirst Then
            refVal = st.First
        ElseIf cfg.PpmRef IsNot Nothing AndAlso Not cfg.PpmRef.Equals("MEAN", StringComparison.OrdinalIgnoreCase) Then
            Dim tmp As Double
            If Double.TryParse(cfg.PpmRef, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, tmp) Then
                refVal = tmp
            End If
        End If

        row.PPM = If(refVal <> 0, ((st.Last - refVal) / refVal) * 1000000.0R, 0)
        row.Count = CInt(st.Count)

        list.Insert(0, row)

        While list.Count > cfg.MaxRows
            list.RemoveAt(list.Count - 1)
        End While

    End Sub


    Private Function ExtractCalcDeps(expr As String) As List(Of String)
        Dim deps As New List(Of String)()
        If String.IsNullOrWhiteSpace(expr) Then Return deps

        Dim i As Integer = 0
        While i < expr.Length
            Dim ch As Char = expr(i)

            ' Identifier starts: letter or underscore
            If Char.IsLetter(ch) OrElse ch = "_"c Then
                Dim j As Integer = i + 1
                While j < expr.Length
                    Dim c2 As Char = expr(j)
                    If Char.IsLetterOrDigit(c2) OrElse c2 = "_"c Then
                        j += 1
                    Else
                        Exit While
                    End If
                End While

                Dim ident As String = expr.Substring(i, j - i).Trim()
                If ident <> "" Then
                    ' Filter out reserved words if you add any later
                    If Not deps.Contains(ident, StringComparer.OrdinalIgnoreCase) Then
                        deps.Add(ident)
                    End If
                End If

                i = j
                Continue While
            End If

            i += 1
        End While

        Return deps
    End Function


    Private Function OpPrec(op As Char) As Integer
        Select Case op
            Case "+"c, "-"c : Return 1
            Case "*"c, "/"c : Return 2
        End Select
        Return 0
    End Function


    Private Iterator Function TokenizeExpr(expr As String) As IEnumerable(Of String)
        Dim i As Integer = 0
        While i < expr.Length
            Dim ch As Char = expr(i)

            If Char.IsWhiteSpace(ch) Then
                i += 1
                Continue While
            End If

            ' Operators / parentheses
            If "+-*/()".IndexOf(ch) >= 0 Then
                Yield ch.ToString()
                i += 1
                Continue While
            End If

            ' Number: 1, 1.23, .5, 1e-6, 1.2E+3   (IMPORTANT: only +/- allowed after e/E)
            If Char.IsDigit(ch) OrElse ch = "."c Then
                Dim j As Integer = i
                Dim sawDot As Boolean = False

                ' integer/decimal part
                While j < expr.Length
                    Dim c2 As Char = expr(j)
                    If Char.IsDigit(c2) Then
                        j += 1
                    ElseIf c2 = "."c AndAlso Not sawDot Then
                        sawDot = True
                        j += 1
                    Else
                        Exit While
                    End If
                End While

                ' optional exponent: e/E then optional sign then digits
                If j < expr.Length AndAlso (expr(j) = "e"c OrElse expr(j) = "E"c) Then
                    Dim k As Integer = j + 1

                    If k < expr.Length AndAlso (expr(k) = "+"c OrElse expr(k) = "-"c) Then
                        k += 1
                    End If

                    Dim expStart As Integer = k
                    While k < expr.Length AndAlso Char.IsDigit(expr(k))
                        k += 1
                    End While

                    If k > expStart Then
                        j = k
                    End If
                End If

                Yield expr.Substring(i, j - i)
                i = j
                Continue While
            End If


            ' Identifier
            If Char.IsLetter(ch) OrElse ch = "_"c Then
                Dim j As Integer = i + 1
                While j < expr.Length
                    Dim c2 As Char = expr(j)
                    If Char.IsLetterOrDigit(c2) OrElse c2 = "_"c Then
                        j += 1
                    Else
                        Exit While
                    End If
                End While
                Yield expr.Substring(i, j - i)
                i = j
                Continue While
            End If

            ' Unknown char: skip
            i += 1
        End While
    End Function


    Private Function TryEvalCalc(expr As String, ByRef result As Double) As Boolean
        result = 0
        If String.IsNullOrWhiteSpace(expr) Then Return False

        Dim output As New List(Of String)()
        Dim ops As New Stack(Of Char)()

        ' Track previous token type so we can detect unary minus
        Dim prevWasValue As Boolean = False  ' value = number/identifier/")"

        ' Shunting-yard: infix -> RPN
        For Each tok In TokenizeExpr(expr)

            ' operator / paren?
            If tok.Length = 1 AndAlso "+-*/()".Contains(tok(0)) Then
                Dim c As Char = tok(0)

                ' Unary minus handling: if "-" appears where a value can't precede it, treat as "0 - ..."
                If c = "-"c AndAlso Not prevWasValue Then
                    output.Add("0")
                    ' proceed as normal binary minus
                End If

                If c = "("c Then
                    ops.Push(c)
                    prevWasValue = False

                ElseIf c = ")"c Then
                    While ops.Count > 0 AndAlso ops.Peek() <> "("c
                        output.Add(ops.Pop().ToString())
                    End While
                    If ops.Count > 0 AndAlso ops.Peek() = "("c Then ops.Pop()
                    prevWasValue = True

                Else
                    While ops.Count > 0 AndAlso ops.Peek() <> "("c AndAlso OpPrec(ops.Peek()) >= OpPrec(c)
                        output.Add(ops.Pop().ToString())
                    End While
                    ops.Push(c)
                    prevWasValue = False
                End If

            Else
                ' number or identifier
                output.Add(tok)
                prevWasValue = True
            End If
        Next

        While ops.Count > 0
            output.Add(ops.Pop().ToString())
        End While

        ' Evaluate RPN
        Dim st As New Stack(Of Double)()
        For Each tok In output
            If tok.Length = 1 AndAlso "+-*/".Contains(tok(0)) Then
                If st.Count < 2 Then Return False
                Dim b As Double = st.Pop()
                Dim a As Double = st.Pop()

                Select Case tok(0)
                    Case "+"c : st.Push(a + b)
                    Case "-"c : st.Push(a - b)
                    Case "*"c : st.Push(a * b)
                    Case "/"c : st.Push(a / b)
                End Select

            Else
                Dim dv As Double
                If Double.TryParse(tok, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, dv) Then
                    st.Push(dv)
                Else
                    ' identifier: lookup last value
                    If Not ResultLastValue.TryGetValue(tok, dv) Then
                        Return False ' missing input => can't compute yet
                    End If
                    st.Push(dv)
                End If
            End If
        Next

        If st.Count <> 1 Then Return False
        result = st.Pop()
        Return True
    End Function


    Private Sub RecalcAllCalcs()
        If CalcDefs.Count = 0 Then Exit Sub

        ' Multi-pass resolves chained calcs (Ref -> PPM etc.)
        ' Keep passes small; chains are short.
        Dim passes As Integer = 4

        For pass As Integer = 1 To passes
            Dim progressed As Boolean = False

            ' NEW local variable: snapshot of current calc defs
            Dim calcList As New List(Of CalcDef)(CalcDefs.Values)

            For Each cd In calcList
                Dim dvOut As Double
                If Not TryEvalCalc(cd.Expr, dvOut) Then Continue For

                ' Avoid endless re-push if unchanged
                Dim prev As Double
                If ResultLastValue.TryGetValue(cd.OutResult, prev) Then
                    If Math.Abs(prev - dvOut) < 0.0R Then
                        ' (leave as-is; you can use an epsilon if you want)
                    End If
                End If

                ResultLastValue(cd.OutResult) = dvOut
                PushNumericResult(cd.OutResult, dvOut)

                progressed = True
            Next

            If Not progressed Then Exit For
        Next
    End Sub


    Private Sub PushNumericResult(resultName As String, dv As Double)
        Dim outText As String = dv.ToString("G", Globalization.CultureInfo.InvariantCulture)
        RunQueryToResult("dev1", "", resultName, rawOverride:=outText)
    End Sub


    Private Sub SetResultFromSpec(spec As String)
        If String.IsNullOrWhiteSpace(spec) Then Exit Sub

        Dim parts = spec.Split({"="c}, 2)
        If parts.Length <> 2 Then Exit Sub

        Dim resultName = parts(0).Trim()
        Dim valueText = parts(1).Trim()

        Dim d As Double
        If Not Double.TryParse(valueText, Globalization.NumberStyles.Float,
                           Globalization.CultureInfo.InvariantCulture, d) Then Exit Sub

        ' Publish into Vars (both friendly + numeric stream)
        Vars(resultName) = d
        Vars($"num:{resultName}") = d

        ' Optional: update a UI control if it exists and is text-based
        Dim c As Control = Nothing
        If ControlByName IsNot Nothing AndAlso ControlByName.TryGetValue(resultName, c) AndAlso c IsNot Nothing Then
            If TypeOf c Is TextBox Then
                DirectCast(c, TextBox).Text = d.ToString(Globalization.CultureInfo.InvariantCulture)
            ElseIf TypeOf c Is Label Then
                DirectCast(c, Label).Text = d.ToString(Globalization.CultureInfo.InvariantCulture)
            End If
        End If
    End Sub


    Private Sub SendFromSpec(spec As String)
        If String.IsNullOrWhiteSpace(spec) Then Exit Sub

        Dim idx = spec.IndexOf(":"c)
        If idx <= 0 Then Exit Sub

        Dim dev = spec.Substring(0, idx).Trim()
        Dim cmd = spec.Substring(idx + 1).Trim()

        If dev = "" OrElse cmd = "" Then Exit Sub

        ' Use your existing native send path
        NativeSend(dev, cmd)
    End Sub


    Private Sub QueryFromSpec(spec As String)
        If String.IsNullOrWhiteSpace(spec) Then Exit Sub

        Dim arrow = spec.IndexOf("->", StringComparison.Ordinal)
        If arrow <= 0 Then Exit Sub

        Dim left = spec.Substring(0, arrow).Trim()
        Dim resultName = spec.Substring(arrow + 2).Trim()
        If resultName = "" Then Exit Sub

        Dim idx = left.IndexOf(":"c)
        If idx <= 0 Then Exit Sub

        Dim dev = left.Substring(0, idx).Trim()
        Dim cmd = left.Substring(idx + 1).Trim()

        If dev = "" OrElse cmd = "" Then Exit Sub

        ' You already have this in User-DeviceIO.vb
        RunQueryToResult(dev, cmd, resultName)
    End Sub


    Private Sub SyncUserUnitsAndScaleFromCheckedRadios()

        Dim modeCaption As String = ""

        ' 1) Find selected MODE radio (caption has no space e.g. "DCV", "DCI", "2WΩ")
        For Each ctrl As Control In GroupBoxCustom.Controls
            Dim gb = TryCast(ctrl, GroupBox)
            If gb Is Nothing Then Continue For

            For Each child As Control In gb.Controls
                Dim rb = TryCast(child, RadioButton)
                If rb Is Nothing OrElse Not rb.Checked Then Continue For

                Dim cap = rb.Text.Trim()
                If cap <> "" AndAlso cap.LastIndexOf(" "c) < 0 Then
                    modeCaption = cap
                    Exit For
                End If
            Next

            If modeCaption <> "" Then Exit For
        Next

        If modeCaption = "" Then Exit Sub

        ' Reset scale state (will be set from selected RANGE radio)
        CurrentUserScaleIsAuto = False
        CurrentUserRangeQuery = ""
        CurrentUserAutoMapSpec = ""
        CurrentUserScale = 1.0
        CurrentUserRangeQueryDone = True
        CurrentUserDp = -1  ' reset DP here as well

        ' 2) Find selected RANGE radio inside the group whose caption matches modeCaption
        For Each ctrl As Control In GroupBoxCustom.Controls
            Dim gb = TryCast(ctrl, GroupBox)
            If gb Is Nothing Then Continue For
            If Not gb.Text.Trim().Equals(modeCaption, StringComparison.OrdinalIgnoreCase) Then Continue For

            For Each child As Control In gb.Controls
                Dim rb = TryCast(child, RadioButton)
                If rb Is Nothing OrElse Not rb.Checked Then Continue For

                ' ---- UNIT from caption tail ("10 V" → "V") ----
                Dim cap = rb.Text.Trim()
                Dim lastSpace = cap.LastIndexOf(" "c)
                If lastSpace >= 0 AndAlso lastSpace < cap.Length - 1 Then
                    CurrentUserUnit = cap.Substring(lastSpace + 1).Trim()
                Else
                    CurrentUserUnit = ""
                End If

                ' ---- SCALE / AUTO from Tag (device|command|scale OR device|command|AUTO|rangeQuery...) ----
                Dim meta As String = TryCast(rb.Tag, String)
                If Not String.IsNullOrEmpty(meta) Then
                    Dim parts() As String = meta.Split("|"c)

                    If parts.Length >= 3 Then
                        If String.Equals(parts(2), "AUTO", StringComparison.OrdinalIgnoreCase) Then
                            CurrentUserScaleIsAuto = True
                            If parts.Length >= 4 Then
                                CurrentUserRangeQuery = parts(3).Trim()
                            End If
                            CurrentUserScale = 1.0
                        Else
                            Dim sc As Double
                            If Double.TryParse(parts(2),
                                           Globalization.NumberStyles.Float,
                                           Globalization.CultureInfo.InvariantCulture,
                                           sc) Then
                                CurrentUserScale = sc
                            Else
                                CurrentUserScale = 1.0
                            End If
                        End If
                    End If

                    ' ---- DP and MAP from Tag (DP=n, MAP=spec) ----
                    CurrentUserDp = -1
                    CurrentUserAutoMapSpec = ""
                    Dim tp() As String = meta.Split("|"c)
                    For Each t In tp
                        Dim tt = t.Trim()

                        If tt.StartsWith("DP=", StringComparison.OrdinalIgnoreCase) Then
                            Dim dpn As Integer
                            If Integer.TryParse(tt.Substring(3), dpn) Then
                                CurrentUserDp = dpn
                            End If

                        ElseIf tt.StartsWith("MAP=", StringComparison.OrdinalIgnoreCase) Then
                            CurrentUserAutoMapSpec = tt.Substring(4).Trim()
                        End If
                    Next

                    'System.Diagnostics.Debug.WriteLine(
                    '"USER-SYNC: mode=" & modeCaption &
                    '" radio='" & rb.Text & "'" &
                    '" meta='" & meta & "'" &
                    '" mapSpec='" & CurrentUserAutoMapSpec & "'" &
                    '" dp=" & CurrentUserDp)


                    ' *** DEBUG: see what we ended up with for this checked RANGE radio ***
                    'System.Diagnostics.Debug.WriteLine(
                    '"USER-AUTO-CONFIG: mode=" & modeCaption &
                    '" group='" & gb.Text & "'" &
                    '" radioText='" & rb.Text & "'" &
                    '" meta='" & meta & "'" &
                    '" scaleIsAuto=" & CurrentUserScaleIsAuto &
                    '" rangeQuery='" & CurrentUserRangeQuery & "'" &
                    '" mapSpec='" & CurrentUserAutoMapSpec & "'" &
                    '" dp=" & CurrentUserDp)

                End If

                ' TEMP: show what we have for K2001 DCV Auto
                'If CurrentUserScaleIsAuto AndAlso
                'modeCaption = "DCV" AndAlso
                'rb.Text.Trim().Equals("Auto", StringComparison.OrdinalIgnoreCase) Then

                'MsgBox("K2001 DCV Auto sync:" & vbCrLf &
                '      "rb.Text = " & rb.Text & vbCrLf &
                '     "Tag = " & CStr(meta) & vbCrLf &
                '     "CurrentUserRangeQuery = " & CurrentUserRangeQuery & vbCrLf &
                '     "CurrentUserAutoMapSpec = " & CurrentUserAutoMapSpec & vbCrLf &
                '     "CurrentUserDp = " & CurrentUserDp.ToString())
                'End If


                ' Determine path also must (re)arm the one-shot range query
                If CurrentUserScaleIsAuto AndAlso Not String.IsNullOrWhiteSpace(CurrentUserRangeQuery) Then
                    CurrentUserRangeQueryDone = False
                Else
                    CurrentUserRangeQueryDone = True
                End If



                Exit Sub
            Next
        Next



    End Sub



    Private Sub ShowUserInitPopup()
        If Not (UserInitPopup Is Nothing) Then Return

        UserInitPopup = New Form()
        With UserInitPopup
            .FormBorderStyle = FormBorderStyle.FixedDialog
            .ControlBox = False
            .MinimizeBox = False
            .MaximizeBox = False
            .ShowInTaskbar = False
            .StartPosition = FormStartPosition.Manual   ' <<< manual positioning
            .Text = "Initializing..."
            .Size = New Size(360, 150)
            .TopMost = True
        End With

        Dim lbl As New Label()
        With lbl
            .Font = New Font("Segoe UI", 11.0F, FontStyle.Regular)
            .Dock = DockStyle.Fill
            .TextAlign = ContentAlignment.MiddleCenter
            .Text = "Loading configuration" & Environment.NewLine &
                "Initializing instruments" & Environment.NewLine &
                "Please wait...."
        End With

        UserInitPopup.Controls.Add(lbl)

        ' --- Center over this form manually ---
        Dim parentForm As Form = Me
        Dim x As Integer = parentForm.Left + (parentForm.Width - UserInitPopup.Width) \ 2
        Dim y As Integer = parentForm.Top + (parentForm.Height - UserInitPopup.Height) \ 2
        UserInitPopup.Location = New Point(x, y)

        ' Disable user area so nothing can be clicked
        If GroupBoxCustom IsNot Nothing Then
            GroupBoxCustom.Enabled = False
        End If

        UserInitPopup.Show(Me)
        UserInitPopup.BringToFront()
        UserInitPopup.Refresh()
    End Sub


    Private Sub ShowConfigWarningPopup(msg As String)

        Dim f As New Form()
        With f
            .FormBorderStyle = FormBorderStyle.FixedDialog
            .ControlBox = False
            .MinimizeBox = False
            .MaximizeBox = False
            .ShowInTaskbar = False
            .StartPosition = FormStartPosition.Manual
            .Text = "WinGPIB – Config Warning"
            .Size = New Size(420, 160)
            .TopMost = True
        End With

        ' Centre on main WinGPIB form (Me)
        Dim owner As Form = Me
        Dim offsetY As Integer = 200
        f.Location = New Point(
    Me.Left + (Me.Width - f.Width) \ 2,
    Me.Top + (Me.Height - f.Height) \ 2 + offsetY)

        Dim lbl As New Label()
        With lbl
            .Dock = DockStyle.Fill
            .TextAlign = ContentAlignment.MiddleCenter
            .Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            .Text = msg
        End With
        f.Controls.Add(lbl)

        ' Auto-close after 5 seconds
        ' Auto-close after 5 seconds  (use existing Timer17)
        Timer17.Stop()
        Timer17.Interval = 10000

        Dim closer As EventHandler = Nothing
        closer =
    Sub(s, e)
        Timer17.Stop()
        RemoveHandler Timer17.Tick, closer   ' prevent stacking handlers
        f.Close()
        f.Dispose()
    End Sub

        AddHandler Timer17.Tick, closer
        Timer17.Start()


        f.Show(owner)   ' modeless, doesn’t block loading
    End Sub



    Private Sub HideUserInitPopup()
        If UserInitPopup IsNot Nothing Then
            Try
                UserInitPopup.Close()
                UserInitPopup.Dispose()
            Catch
            End Try
            UserInitPopup = Nothing
        End If

        If GroupBoxCustom IsNot Nothing Then
            GroupBoxCustom.Enabled = True
        End If
    End Sub


    Private Sub ButtonEditor_Click(sender As Object, e As EventArgs) Handles ButtonEditor.Click

        Dim cfgPath As String = Nothing

        ' 1) Prefer the last loaded user config, if it exists
        If Not String.IsNullOrWhiteSpace(LastUserConfigPath) AndAlso
           IO.File.Exists(LastUserConfigPath) Then

            cfgPath = LastUserConfigPath

        Else
            ' 2) Otherwise, let the user pick a file (same base logic as ButtonLoadTxt_Click)

            Using dlg As New OpenFileDialog()
                dlg.Title = "Select Config File to Edit"
                dlg.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

                ' Base folder: prefer CSVfilepath.Text if it is a valid directory,
                ' else fall back to Documents\WinGPIBdata
                Dim baseDir As String = CSVfilepath.Text
                If String.IsNullOrWhiteSpace(baseDir) OrElse Not IO.Directory.Exists(baseDir) Then
                    Dim documentsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    baseDir = IO.Path.Combine(documentsPath, "WinGPIBdata")
                End If

                ' Force dialog into \WinGPIBdata\Devices
                Dim devicesDir As String = IO.Path.Combine(baseDir, "Devices")
                IO.Directory.CreateDirectory(devicesDir)

                dlg.InitialDirectory = devicesDir

                If dlg.ShowDialog() <> DialogResult.OK Then Exit Sub

                cfgPath = dlg.FileName
            End Using
        End If

        If String.IsNullOrWhiteSpace(cfgPath) Then Exit Sub

        ' 3) Open the editor on that path
        Dim f As New FormUserConfigEditor(cfgPath)

        f.Show(Me)   ' or f.ShowDialog(Me) if you prefer modal
    End Sub



End Class
