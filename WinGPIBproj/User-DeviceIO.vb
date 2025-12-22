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


    Private Sub RunQueryToResult(deviceName As String,
                                commandOrPrefix As String,
                                resultControlName As String,
                                Optional rawOverride As String = Nothing,
                                Optional overloadToken As String = Nothing)

        Dim target = GetControlByName(resultControlName)

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
        '   OVERLOAD DETECT (user-defined token match)
        ' =========================================================
        If Not String.IsNullOrWhiteSpace(overloadToken) Then
            If raw.IndexOf(overloadToken, StringComparison.OrdinalIgnoreCase) >= 0 Then

                outText = "OVERLOAD"

                ' Publish a numeric marker for trigger/calc users (optional but useful)
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

            ' Publish numeric result for TRIGGER engine
            ' v is the scaled numeric value you already computed
            Vars(resultControlName) = v
            Vars($"num:{resultControlName}") = v
            Vars($"bignum:{resultControlName}") = v

            ' Calc - Cache last numeric value for this stream
            ResultLastValue(resultControlName) = v
            ' Prevent recursive calc -> push -> RunQueryToResult -> calc loops
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

                ' Apply format if requested AND we have a numeric value
                Dim displayNumeric As String = baseNumeric
                If fmt <> "" Then
                    Dim dv As Double
                    If Double.TryParse(baseNumeric,
                               Globalization.NumberStyles.Float,
                               Globalization.CultureInfo.InvariantCulture,
                               dv) Then
                        displayNumeric = dv.ToString(fmt, Globalization.CultureInfo.InvariantCulture)
                    End If
                End If

                If unitsOff AndAlso numericText <> "" Then
                    ' BIGTEXT units=off → numeric only (formatted if fmt was applied)
                    lbl.Text = displayNumeric
                ElseIf numericWithUnitText <> "" Then
                    ' Normal label/BIGTEXT units=on
                    ' If we formatted, rebuild with units; otherwise keep original string
                    If fmt <> "" AndAlso displayNumeric <> "" Then
                        If Not String.IsNullOrEmpty(CurrentUserUnit) Then
                            lbl.Text = displayNumeric & "   " & CurrentUserUnit
                        Else
                            lbl.Text = displayNumeric
                        End If
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

        Const timeoutMs As Integer = 5000   ' adjust if needed

        If deviceName.Equals("dev1", StringComparison.OrdinalIgnoreCase) Then

            USERdev1rawoutput = requireRaw
            Try

                OutputReceiveddev1 = False
                txtq1a.Text = cmd
                RunBtnq1aCore()

                Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()
                Do While Not OutputReceiveddev1 AndAlso sw.ElapsedMilliseconds < timeoutMs
                    Application.DoEvents()
                    Threading.Thread.Sleep(1)
                Loop

                Return If(USERdev1output, "").Trim()

            Finally
                USERdev1rawoutput = False
            End Try

        ElseIf deviceName.Equals("dev2", StringComparison.OrdinalIgnoreCase) Then

            USERdev2rawoutput = requireRaw
            Try

                OutputReceiveddev2 = False
                txtq2a.Text = cmd
                RunBtnq2aCore()

                Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()
                Do While Not OutputReceiveddev2 AndAlso sw.ElapsedMilliseconds < timeoutMs
                    Application.DoEvents()
                    Threading.Thread.Sleep(1)
                Loop

                Return If(USERdev2output, "").Trim()

            Finally
                USERdev2rawoutput = False
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
            Return q.ResponseAsString
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