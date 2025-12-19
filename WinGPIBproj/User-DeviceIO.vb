' User customizeable tab - Device IO


Partial Class Formtest

    ' Latest DEVICES-tab query outputs for USER tab to read
    Private USERdev1output As String = ""
    Private USERdev2output As String = ""
    Private USERdev1output2 As String = ""
    Private USERdev2output2 As String = ""
    Private OutputReceiveddev1 As Boolean = False
    Private OutputReceiveddev2 As Boolean = False


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


End Class