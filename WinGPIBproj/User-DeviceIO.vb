' User customizeable tab - Device IO


Partial Class Formtest

    ' Latest DEVICES-tab query outputs for USER tab to read
    Private USERdev1output As String = ""
    Private USERdev2output As String = ""
    Private USERdev1output2 As String = ""
    Private USERdev2output2 As String = ""
    Private OutputReceiveddev1 As Boolean = False
    Private OutputReceiveddev2 As Boolean = False

    ' When TRUE, native queries should store the *raw* instrument response into USERdevXoutput
    ' (used by determine=...|...|resptext)
    Private USERdev1rawoutput As Boolean = False
    Private USERdev2rawoutput As Boolean = False
    Private USERdev1fastquery As Boolean = False
    Private USERdev2fastquery As Boolean = False

    ' Cache for auto-scale range queries (key: "dev||query")
    Private ReadOnly UserRangeQueryCache As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
    ' Has the auto-range query already been sent for this selection?
    Private CurrentUserRangeQueryDone As Boolean = False

    ' Skip the very next User display update after a mode/range change
    Private UserSkipNextUserDisplay As Boolean = False


    Private Sub RunQueryToResult(deviceName As String,
                                commandOrPrefix As String,
                                resultControlName As String,
                                Optional rawOverride As String = Nothing,
                                Optional overloadToken As String = Nothing)

        Dim target = GetControlByName(resultControlName)

        ' test only
        'DetDbgLog($"AUTO DEBUG dev={deviceName}, cmd={commandOrPrefix}, target={resultControlName}")
        'ShowCopyableDebug("Determine Debug", _detDbg.ToString())

        ' If this is a pushed value (CALC / fan-out) and there is no direct UI control,
        ' that's OK — linked controls (CHART/STATS/HISTORYGRID/BIGTEXT etc.) may still exist.
        If target Is Nothing AndAlso rawOverride Is Nothing Then
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

        ' Single READ button sync units
        Try
            SyncUserUnitsAndScaleFromCheckedRadios()
        Catch
            ' Best-effort only – never let a units sync failure kill the read
        End Try

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

                Dim overloadSource As String = "disp"
                Dim wantRaw As Boolean = False

                If Not String.IsNullOrWhiteSpace(overloadToken) AndAlso overloadToken.Contains("|"c) Then
                    Dim sp = overloadToken.Split({"|"c}, 2)
                    overloadSource = sp(1).Trim().ToLowerInvariant()
                    wantRaw = (overloadSource = "raw")
                Else
                    ' No overload info → if this result is driven by a DATASOURCE,
                    ' prefer RAW data so auto-refresh works even without overload=
                    Dim ds As DataSourceDef
                    If DataSources IsNot Nothing AndAlso DataSources.TryGetValue(resultControlName, ds) Then
                        wantRaw = True
                    End If
                End If

                ' NATIVE PATH: use existing UI query buttons/textboxes
                raw = NativeQuery(deviceName, commandOrPrefix, requireRaw:=wantRaw)

            Else

                ' STANDALONE PATH: original IODevices QueryBlocking behaviour
                Dim q As IODevices.IOQuery = Nothing
                Dim status As Integer

                Dim fullCmd As String = commandOrPrefix & TermStr2()
                status = dev.QueryBlocking(fullCmd, q, False)

                If status = 0 AndAlso q IsNot Nothing Then

                    raw = respNorm.Trim()

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


        ' test only
        'MessageBox.Show($"RAW-trim: dev={deviceName}, cmd={commandOrPrefix}, target={resultControlName}, raw=[{raw}]")




        ' =========================================================
        '   OVERLOAD DETECT (overload=token|raw or token|disp)
        '   Default (no |suffix) = disp
        '   Uses CONTAINS match; RAW is trimmed to ignore leading spaces
        ' =========================================================
        If Not String.IsNullOrWhiteSpace(overloadToken) Then

            ' Parse token + optional source selector
            Dim overloadValue As String = overloadToken.Trim()
            Dim overloadSource As String = "disp"   ' default

            If overloadValue.Contains("|"c) Then
                Dim sp = overloadValue.Split({"|"c}, 2)
                overloadValue = sp(0).Trim()
                overloadSource = sp(1).Trim().ToLowerInvariant()
            End If

            ' Choose comparison text
            Dim cmpText As String = ""

            If IsNativeEngine(deviceName) Then
                ' Native: we have BOTH raw and disp available from Formtest.vb capture
                If overloadSource = "raw" Then
                    cmpText =
                If(deviceName.Equals("dev2", StringComparison.OrdinalIgnoreCase),
                   If(USERdev2output2, "").Trim(),
                   If(USERdev1output2, "").Trim())
                Else
                    ' disp
                    cmpText =
                If(deviceName.Equals("dev2", StringComparison.OrdinalIgnoreCase),
                   If(USERdev2output, ""),
                   If(USERdev1output, ""))
                End If
            Else
                ' Standalone: only the query reply exists here -> treat as raw
                cmpText = If(raw, "").Trim()
            End If

            ' CONTAINS match
            If overloadValue <> "" AndAlso
       cmpText.IndexOf(overloadValue, StringComparison.OrdinalIgnoreCase) >= 0 Then

                outText = "OVERLOAD"

                Vars(resultControlName) = Double.NaN
                Vars($"num:{resultControlName}") = Double.NaN
                Vars($"bignum:{resultControlName}") = Double.NaN
                ResultLastValue(resultControlName) = Double.NaN

                GoTo FanOut
            End If

        End If


        ' =========================================================
        '   EXISTING FORMAT / SCALE / UNITS LOGIC (unchanged)
        ' =========================================================
        'Dim decCb As CheckBox = GetCheckboxFor(resultControlName, "FuncDecimal")           ' moved to config

        Dim scale As Double = CurrentUserScale
        Dim d As Double

        If TryExtractFirstDouble(raw, d) Then

            Dim v As Double = d

            ' ============================
            '    AUTO SCALE MODE LOGIC
            ' ============================
            If CurrentUserScaleIsAuto AndAlso
               dev IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(CurrentUserRangeQuery) Then

                Dim cacheKey As String = deviceName & "||" & CurrentUserRangeQuery
                Dim rngVal As Double

                ' Always query current range in AUTO mode – one extra query per reading
                If Not String.IsNullOrEmpty(CurrentUserRangeQuery) Then

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
                                'rngStr = rq.ResponseAsString.Trim()                            ' Original potentially unsafe ResponseAsString
                                rngStr = NormalizeNumericResponse(rq.ResponseAsString).Trim()   ' locale-safe as it uses the helper (Formtest.vb)
                            End If
                        End If

                        '                       System.Diagnostics.Debug.WriteLine(
                        '   "USER-AUTO-RANGE: dev=" & deviceName &
                        '   " query='" & CurrentUserRangeQuery & "'" &
                        '   " replyRaw='" & rngStr & "'")


                        If Double.TryParse(rngStr,
                                           Globalization.NumberStyles.Float,
                                           Globalization.CultureInfo.InvariantCulture,
                                           rngVal) Then

                            ' Optional: keep latest range in cache if you use it elsewhere
                            UserRangeQueryCache(cacheKey) = rngVal

                            ' First try instrument-specific AUTO map; fall back to generic engineering scale
                            If Not TryApplyAutoRangeMap(rngVal, d, v) Then
                                Dim dynScale As Double = ComputeAutoScaleFromRange(rngVal)
                                v = d * dynScale
                            End If
                        Else
                            ' Parse failed – fallback to fixed scale (likely 1.0)
                            v = d * scale
                        End If

                    Catch
                        ' Error – fallback to fixed scale
                        v = d * scale
                    End Try

                Else
                    ' No AUTO range query defined; just use fixed scale
                    v = d * scale
                End If

            Else
                ' FIXED SCALE MODE
                v = d * scale
            End If

            ' ============================
            '      OUTPUT FORMATTING
            ' ============================
            Dim numericStr As String   ' for charts / stats
            Dim displayStr As String   ' for BIGTEXT

            ' Look up DATASOURCE to see if user wants decimal or raw/sci
            Dim useDecimal As Boolean = True
            Dim ds As DataSourceDef = Nothing
            If DataSources.TryGetValue(resultControlName, ds) AndAlso ds IsNot Nothing Then
                useDecimal = ds.ForceDecimal
            End If

            ' Always generate a clean numeric string for internal use
            numericStr = v.ToString("G",
                                    Globalization.CultureInfo.InvariantCulture)

            If useDecimal Then
                ' decimal=1 (or default) → show decimal formatted value in BIGTEXT
                displayStr = v.ToString("0.###############", Globalization.CultureInfo.InvariantCulture)
            Else
                ' decimal=0 → show instrument's raw string if present,
                ' otherwise fall back to the clean numeric.
                If Not String.IsNullOrWhiteSpace(raw) Then
                    displayStr = raw.Trim()
                Else
                    displayStr = numericStr
                End If
            End If

            Debug.WriteLine($"RAW for {resultControlName} = [{raw}]")


            ' Charts / stats:
            ' - When decimal=1, numericText is the clean numeric (BIGTEXT can reformat it).
            ' - When decimal=0, leave numericText empty so BIGTEXT uses outText/raw instead.
            If useDecimal Then
                numericText = numericStr
            Else
                numericText = ""
            End If


            ' BIGTEXT uses displayStr (+ unit)
            If Not String.IsNullOrEmpty(CurrentUserUnit) Then
                numericWithUnitText = displayStr & "   " & CurrentUserUnit
            Else
                numericWithUnitText = displayStr
            End If

            outText = numericWithUnitText

            ' Publish numeric result for TRIGGER engine
            Vars(resultControlName) = v
            Vars($"num:{resultControlName}") = v
            Vars($"bignum:{resultControlName}") = v

            ' Calc - Cache last numeric value for this stream
            ResultLastValue(resultControlName) = v
            If Threading.Interlocked.Increment(UserCalcDepth) = 1 Then
                Try
                    RecalcAllCalcs()
                Finally
                    Threading.Interlocked.Decrement(UserCalcDepth)
                End Try
            Else
                Threading.Interlocked.Decrement(UserCalcDepth)
            End If


        Else
            ' Non-numeric reply: clear numeric fields and pass through raw text
            numericText = ""
            numericWithUnitText = ""
            outText = raw
        End If


FanOut:

        ' ============================
        '      FAN-OUT TO CONTROLS
        ' ============================

        ' Skip the very next display update after a mode/range change
        If UserSkipNextUserDisplay Then
            UserSkipNextUserDisplay = False
            Exit Sub
        End If

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

                Dim unitsOff As Boolean = False
                Dim fmt As String = ""

                If Not String.IsNullOrEmpty(tagStr) Then
                    Dim tp() As String = tagStr.Split("|"c)
                    For Each t In tp
                        Dim tt As String = t.Trim()
                        If tt.Equals("BIGTEXT_UNITS_OFF", StringComparison.OrdinalIgnoreCase) Then
                            unitsOff = True
                        ElseIf tt.StartsWith("BIGTEXT_FMT=", StringComparison.OrdinalIgnoreCase) Then
                            fmt = tt.Substring("BIGTEXT_FMT=".Length).Trim()
                        End If
                    Next
                End If

                ' Choose base numeric string (prefer numericText, then parse from feed)
                Dim baseNumeric As String = numericText
                If String.IsNullOrWhiteSpace(baseNumeric) Then
                    baseNumeric = outText
                End If

                ' Apply DP or format if requested AND we have a numeric value
                Dim displayNumeric As String = baseNumeric
                Dim dv As Double
                If Double.TryParse(baseNumeric,
                                   Globalization.NumberStyles.Float,
                                   Globalization.CultureInfo.InvariantCulture,
                                   dv) Then

                    ' 1) RANGE dp= wins
                    If CurrentUserDp >= 0 Then
                        displayNumeric = dv.ToString("F" & CurrentUserDp,
                                                     Globalization.CultureInfo.InvariantCulture)

                        ' 2) otherwise BIGTEXT_FMT= if present
                    ElseIf fmt <> "" Then
                        displayNumeric = dv.ToString(fmt,
                                                     Globalization.CultureInfo.InvariantCulture)
                    End If
                End If

                If unitsOff AndAlso numericText <> "" Then
                    ' BIGTEXT units=off → numeric only (formatted if DP/FMT was applied)
                    lbl.Text = displayNumeric

                ElseIf numericWithUnitText <> "" Then
                    ' Normal label/BIGTEXT units=on
                    ' Only rebuild with DP/FMT when we also have a unit; otherwise keep existing text
                    If (CurrentUserDp >= 0 OrElse fmt <> "") AndAlso
                       displayNumeric <> "" AndAlso
                       Not String.IsNullOrEmpty(CurrentUserUnit) Then

                        lbl.Text = displayNumeric & "   " & CurrentUserUnit
                    Else
                        lbl.Text = numericWithUnitText
                    End If

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

        ' UPDATED LINKED CONTROLS (no TextBox required)
        Dim feedText As String = If(numericWithUnitText <> "", numericWithUnitText, outText)

        ' Walk all controls in GroupBoxCustom recursively
        Dim stack As New Stack(Of Control)
        stack.Push(GroupBoxCustom)

        While stack.Count > 0
            Dim c As Control = stack.Pop()

            For Each child As Control In c.Controls
                stack.Push(child)
            Next

            ' ---- CHART ----
            If TypeOf c Is DataVisualization.Charting.Chart Then
                Dim tagStr As String = TryCast(c.Tag, String)
                If Not String.IsNullOrEmpty(tagStr) AndAlso tagStr.StartsWith("CHART|", StringComparison.OrdinalIgnoreCase) Then
                    Dim tp() As String = tagStr.Split("|"c)
                    ' CHART|target|ymin|ymax|xstep|maxpoints|autoscale
                    If tp.Length >= 7 AndAlso tp(1).Equals(resultControlName, StringComparison.OrdinalIgnoreCase) Then

                        Dim yMin2 As Double? = Nothing
                        Dim yMax2 As Double? = Nothing
                        Dim xStep2 As Double = 1.0R
                        Dim maxPts2 As Integer = 100
                        Dim auto2 As Boolean = False

                        If tp(2) <> "" AndAlso Double.TryParse(tp(2), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, d) Then
                            yMin2 = d
                        End If

                        If tp(3) <> "" AndAlso Double.TryParse(tp(3), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, d) Then
                            yMax2 = d
                        End If

                        Double.TryParse(tp(4), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, xStep2)
                        Integer.TryParse(tp(5), maxPts2)
                        auto2 = (tp(6) = "1")

                        UpdateChartFromText(DirectCast(c, DataVisualization.Charting.Chart), feedText, yMin2, yMax2, xStep2, maxPts2, auto2)
                    End If
                End If
            End If

            ' ---- STATSPANEL ----
            If TypeOf c Is GroupBox Then
                Dim tagStr As String = TryCast(c.Tag, String)
                If Not String.IsNullOrEmpty(tagStr) AndAlso tagStr.StartsWith("STATSPANEL|", StringComparison.OrdinalIgnoreCase) Then
                    Dim tp() As String = tagStr.Split("|"c)
                    ' STATSPANEL|target|fmt
                    If tp.Length >= 2 AndAlso tp(1).Equals(resultControlName, StringComparison.OrdinalIgnoreCase) Then
                        Dim dv As Double
                        If TryExtractFirstDouble(feedText, dv) Then
                            UpdateStatsPanel(c.Name, dv)
                        End If
                    End If
                End If
            End If

        End While

        ' ---- POPUP CHARTS (in separate windows) ----
        If ChartSettings IsNot Nothing AndAlso ChartPopupForms IsNot Nothing Then
            For Each kv In ChartSettings
                Dim cfg As ChartConfig = kv.Value
                If cfg Is Nothing OrElse Not cfg.Popup Then Continue For

                ' Only charts bound to this resultControlName
                If Not String.Equals(cfg.ResultTarget, resultControlName, StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim popupForm As Form = Nothing
                If Not ChartPopupForms.TryGetValue(kv.Key, popupForm) Then Continue For
                If popupForm Is Nothing OrElse popupForm.IsDisposed Then Continue For

                ' Find the Chart inside the popup form
                Dim popupChart As DataVisualization.Charting.Chart = Nothing
                For Each ctrl As Control In popupForm.Controls
                    popupChart = TryCast(ctrl, DataVisualization.Charting.Chart)
                    If popupChart IsNot Nothing Then Exit For
                Next
                If popupChart Is Nothing Then Continue For

                Dim yMin2 As Double? = cfg.YMin
                Dim yMax2 As Double? = cfg.YMax
                Dim xStep2 As Double = cfg.XStep
                Dim maxPts2 As Integer = cfg.MaxPoints
                Dim auto2 As Boolean = cfg.AutoScaleY

                UpdateChartFromText(popupChart, feedText, yMin2, yMax2, xStep2, maxPts2, auto2)
            Next
        End If

        ' ---- HISTORYGRID (no TextBox required) ----
        Dim grids As List(Of String) = Nothing
        If HistoryGridsByTarget.TryGetValue(resultControlName, grids) Then
            Dim dv As Double
            If TryExtractFirstDouble(feedText, dv) Then
                For Each gridName In grids
                    AddHistorySample(gridName, dv)
                Next
            End If
        End If

    End Sub


    Private Function NativeQuery(deviceName As String, cmd As String, Optional requireRaw As Boolean = False) As String

        Const timeoutMs As Integer = 5000

        If deviceName.Equals("dev1", StringComparison.OrdinalIgnoreCase) Then

            ' Tell Formtest callback to use the fast path (fills USERdev1output2 + USERdev1output)
            USERdev1fastquery = True

            Try
                OutputReceiveddev1 = False
                txtq1a.Text = cmd
                RunBtnq1aCore()

                Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()
                Do While Not OutputReceiveddev1 AndAlso sw.ElapsedMilliseconds < timeoutMs
                    Application.DoEvents()
                    Threading.Thread.Sleep(1)
                Loop

                Dim s As String = If(requireRaw, USERdev1output2, USERdev1output)
                Return If(s, "").Trim()

            Finally
                USERdev1fastquery = False
            End Try

        ElseIf deviceName.Equals("dev2", StringComparison.OrdinalIgnoreCase) Then

            USERdev2fastquery = True

            Try
                OutputReceiveddev2 = False
                txtq2a.Text = cmd
                RunBtnq2aCore()

                Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()
                Do While Not OutputReceiveddev2 AndAlso sw.ElapsedMilliseconds < timeoutMs
                    Application.DoEvents()
                    Threading.Thread.Sleep(1)
                Loop

                Dim s As String = If(requireRaw, USERdev2output2, USERdev2output)
                Return If(s, "").Trim()

            Finally
                USERdev2fastquery = False
            End Try

        End If

        Return ""
    End Function


    ' Used by determine=... on RADIOs. Performs a single query and returns the raw response text.
    ' detFmt:
    '   resptext -> sets USERdevXrawoutput TRUE during query
    '   respnum  -> leaves USERdevXrawoutput FALSE (default)
    Private Function DetermineQuery(deviceName As String, cmd As String, detFmt As String) As String

        Dim requireRaw As Boolean = String.Equals(detFmt, "resptext", StringComparison.OrdinalIgnoreCase)

        If IsNativeEngine(deviceName) Then

            If deviceName.Equals("dev1", StringComparison.OrdinalIgnoreCase) Then
                USERdev1rawoutput = requireRaw
                USERdev1fastquery = True          ' <<< ONLY HERE
            ElseIf deviceName.Equals("dev2", StringComparison.OrdinalIgnoreCase) Then
                USERdev2rawoutput = requireRaw
                USERdev2fastquery = True
            End If

            Try
                NativeQuery(deviceName, cmd, requireRaw)

                If deviceName.Equals("dev1", StringComparison.OrdinalIgnoreCase) Then
                    Return If(USERdev1output2, "").Trim()
                Else
                    Return If(USERdev2output2, "").Trim()
                End If

            Finally
                If deviceName.Equals("dev1", StringComparison.OrdinalIgnoreCase) Then
                    USERdev1rawoutput = False
                    USERdev1fastquery = False     ' <<< RESET HERE
                ElseIf deviceName.Equals("dev2", StringComparison.OrdinalIgnoreCase) Then
                    USERdev2rawoutput = False
                    USERdev2fastquery = False
                End If
            End Try
        End If

        ' Standalone engine path unchanged
        Dim dev As IODevices.IODevice = Nothing
        Select Case deviceName.ToLowerInvariant()
            Case "dev1" : dev = dev1
            Case "dev2" : dev = dev2
        End Select
        If dev Is Nothing Then Return ""

        Dim q As IODevices.IOQuery = Nothing
        Dim status = dev.QueryBlocking(cmd & TermStr2(), q, False)
        If status = 0 AndAlso q IsNot Nothing Then
            Return respNorm
        End If

        Return ""
    End Function


    ' TXT Format:
    ' BUTTON;Caption;Action;Device;CommandOrPrefix;ValueControl;ResultControl
    '
    ' Action:
    '   SEND        = send fixed command
    '   SENDVALUE   = send prefix + value from ValueControl
    '   QUERY       = query -> result stored into ResultControl
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


End Class