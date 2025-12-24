' User customizeable tab - Lua support

Imports System.Globalization
Imports MoonSharp.Interpreter
Imports MoonSharp.Interpreter.Interop


Partial Class Formtest


    ' LUA
    Dim inLuaScript As Boolean = False
    Dim luaScriptName As String = ""
    Dim luaLines As New List(Of String)
    Dim autoY As Integer = 10

    ' Lua engine instance
    Private _lua As Script
    Private _luaReady As Boolean = False

    ' Lua cancel flag
    Private _luaCancelRequested As Boolean = False

    ' Lua scripts loaded from config (no UI control required)
    Private ReadOnly LuaScriptsByName As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)


    ' call this to request stop
    Private Sub RequestLuaStop()
        _luaCancelRequested = True
        AppendLog("[LUA] Stopped by user", "LuaLog")
    End Sub


    ' throw to abort Lua cleanly
    Private Sub ThrowIfLuaCancelled()
        If _luaCancelRequested Then
            Throw New OperationCanceledException("Lua cancelled")
        End If
    End Sub


    Private Sub EnsureLua()
        If _luaReady Then Exit Sub

        ' Sandbox
        UserData.RegistrationPolicy = InteropRegistrationPolicy.Default

        _lua = New Script(CoreModules.Preset_SoftSandbox)

        ' Expose small API to Lua
        _lua.Globals("logto") = CType(Sub(target As String, msg As Object)
                                          If String.IsNullOrWhiteSpace(target) Then
                                              target = "TriggerLog"
                                          End If
                                          AppendLog("[LUA] " & If(msg, "").ToString(), target)
                                      End Sub, Action(Of String, Object))

        _lua.Globals("fire") = CType(Sub(ctrlName As String)
                                         If String.IsNullOrWhiteSpace(ctrlName) Then Exit Sub

                                         ' 1) User-tab fast path
                                         Dim b As Button = Nothing
                                         If BtnByName IsNot Nothing AndAlso BtnByName.TryGetValue(ctrlName, b) AndAlso b IsNot Nothing Then
                                             If b.InvokeRequired Then
                                                 b.BeginInvoke(Sub() b.PerformClick())
                                             Else
                                                 b.PerformClick()
                                             End If
                                             Exit Sub
                                         End If

                                         ' 2) Find control anywhere
                                         Dim c As Control = Me.Controls.Find(ctrlName, True).FirstOrDefault()
                                         Dim btn As Button = TryCast(c, Button)

                                         If btn Is Nothing Then
                                             AppendLog("[LUA] fire(): button not found: " & ctrlName)
                                             Exit Sub
                                         End If

                                         btn.BeginInvoke(Sub()
                                                             ' Find owning TabControl / TabPage
                                                             Dim tp As TabPage = Nothing
                                                             Dim parent As Control = btn.Parent
                                                             While parent IsNot Nothing
                                                                 tp = TryCast(parent, TabPage)
                                                                 If tp IsNot Nothing Then Exit While
                                                                 parent = parent.Parent
                                                             End While

                                                             Dim tc As TabControl = TryCast(tp?.Parent, TabControl)
                                                             Dim oldTab As TabPage = If(tc IsNot Nothing, tc.SelectedTab, Nothing)

                                                             ' Switch tab if required
                                                             If tc IsNot Nothing AndAlso tp IsNot Nothing AndAlso tc.SelectedTab IsNot tp Then
                                                                 tc.SelectedTab = tp
                                                                 Application.DoEvents()
                                                             End If

                                                             btn.PerformClick()

                                                             ' Restore original tab
                                                             If tc IsNot Nothing AndAlso oldTab IsNot Nothing Then
                                                                 tc.SelectedTab = oldTab
                                                             End If
                                                         End Sub)
                                     End Sub, Action(Of String))

        _lua.Globals("getnum") = CType(Function(varName As String) As Double
                                           Dim o = GetVarValue(varName) ' you already have this
                                           Dim d As Double
                                           If o IsNot Nothing AndAlso Double.TryParse(o.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, d) Then
                                               Return d
                                           End If
                                           Return Double.NaN
                                       End Function, Func(Of String, Double))

        _lua.Globals("send") = CType(Sub(devName As String, cmd As String)
                                         devName = If(devName, "").Trim()
                                         cmd = If(cmd, "").Trim()

                                         If String.IsNullOrWhiteSpace(devName) OrElse String.IsNullOrWhiteSpace(cmd) Then Exit Sub

                                         If IsNativeEngine(devName) Then
                                             NativeSend(devName, cmd)
                                         Else
                                             Dim dev As IODevices.IODevice = Nothing
                                             Select Case devName.ToLowerInvariant()
                                                 Case "dev1" : dev = dev1
                                                 Case "dev2" : dev = dev2
                                             End Select

                                             If dev Is Nothing Then
                                                 AppendLog("[LUA] send(): device not available: " & devName)
                                                 Exit Sub
                                             End If

                                             dev.SendAsync(cmd, True)
                                         End If
                                     End Sub, Action(Of String, String))


        _lua.Globals("query") = CType(Function(devName As String, cmd As String) As String
                                          devName = If(devName, "").Trim()
                                          cmd = If(cmd, "").Trim()

                                          If String.IsNullOrWhiteSpace(devName) OrElse String.IsNullOrWhiteSpace(cmd) Then Return ""

                                          If IsNativeEngine(devName) Then
                                              Return NativeQuery(devName, cmd)
                                          Else
                                              Dim dev As IODevices.IODevice = Nothing
                                              Select Case devName.ToLowerInvariant()
                                                  Case "dev1" : dev = dev1
                                                  Case "dev2" : dev = dev2
                                              End Select

                                              If dev Is Nothing Then
                                                  AppendLog("[LUA] query(): device not available: " & devName)
                                                  Return ""
                                              End If

                                              Dim q As IODevices.IOQuery = Nothing
                                              Dim fullCmd As String = cmd & TermStr2()
                                              Dim status As Integer = dev.QueryBlocking(fullCmd, q, False)

                                              If status = 0 AndAlso q IsNot Nothing Then
                                                  Return q.ResponseAsString.Trim()
                                              End If

                                              If q IsNot Nothing Then
                                                  AppendLog("[LUA] query(): ERR " & status & ": " & q.errmsg)
                                              Else
                                                  AppendLog("[LUA] query(): ERR " & status & " (no IOQuery)")
                                              End If

                                              Return ""
                                          End If
                                      End Function, Func(Of String, String, String))

        _lua.Globals("settext") = CType(Sub(ctrlName As String, val As Object)
                                            Dim c = GetControlByName(If(ctrlName, "").Trim())
                                            If c Is Nothing Then
                                                AppendLog("[LUA] settext(): control not found: " & ctrlName)
                                                Exit Sub
                                            End If

                                            Dim s As String = If(val, "").ToString()

                                            If TypeOf c Is TextBoxBase Then
                                                Dim tb = DirectCast(c, TextBoxBase)
                                                If tb.InvokeRequired Then
                                                    tb.BeginInvoke(Sub() tb.Text = s)
                                                Else
                                                    tb.Text = s
                                                End If
                                                Exit Sub
                                            End If

                                            If TypeOf c Is Label Then
                                                Dim lbl = DirectCast(c, Label)
                                                If lbl.InvokeRequired Then
                                                    lbl.BeginInvoke(Sub() lbl.Text = s)
                                                Else
                                                    lbl.Text = s
                                                End If
                                                Exit Sub
                                            End If

                                            AppendLog("[LUA] settext(): unsupported control type: " & c.GetType().Name)
                                        End Sub, Action(Of String, Object))

        ' UI repaint helper (optional but useful)
        _lua.Globals("doui") = CType(Sub()
                                         Application.DoEvents()
                                     End Sub, Action)

        ' Simple sleep (blocks UI thread if Lua runs on UI thread)
        _lua.Globals("sleep") = CType(Sub(ms As Double)
                                          Dim total As Integer = CInt(Math.Max(0, ms))
                                          Dim remaining As Integer = total

                                          While remaining > 0
                                              ThrowIfLuaCancelled()

                                              Dim stepMs As Integer = Math.Min(50, remaining) ' 50ms chunks
                                              Threading.Thread.Sleep(stepMs)
                                              remaining -= stepMs

                                              ' optional: let UI breathe if you’re running on UI thread
                                              Application.DoEvents()
                                          End While
                                      End Sub, Action(Of Double))

        ' Expose setnum() to Lua: publishes numeric values into Vars() so TRIGGERs can see them
        _lua.Globals("setnum") = CType(Sub(varName As String, val As Object)
                                           varName = If(varName, "").Trim()
                                           If varName = "" Then Exit Sub

                                           Dim d As Double
                                           If val Is Nothing OrElse Not Double.TryParse(val.ToString(),
                                       NumberStyles.Float, CultureInfo.InvariantCulture, d) Then
                                               AppendLog("[LUA] setnum(): not numeric: " & varName & "=" & If(val, ""))
                                               Exit Sub
                                           End If

                                           ' Publish for triggers and general use
                                           Vars(varName) = d
                                           Vars($"num:{varName}") = d
                                           Vars($"bignum:{varName}") = d

                                           ' Optional: update a UI control with same name (TextBox/Label)
                                           Dim c = GetControlByName(varName)
                                           If c IsNot Nothing Then
                                               Dim s = d.ToString(CultureInfo.InvariantCulture)

                                               If TypeOf c Is TextBoxBase Then
                                                   Dim tb = DirectCast(c, TextBoxBase)
                                                   If tb.InvokeRequired Then tb.BeginInvoke(Sub() tb.Text = s) Else tb.Text = s
                                                   Exit Sub
                                               End If

                                               If TypeOf c Is Label Then
                                                   Dim lbl = DirectCast(c, Label)
                                                   If lbl.InvokeRequired Then lbl.BeginInvoke(Sub() lbl.Text = s) Else lbl.Text = s
                                                   Exit Sub
                                               End If
                                           End If
                                       End Sub, Action(Of String, Object))

        ' Expose now() to Lua: returns current system time (hour/min/sec)
        _lua.Globals("now") = CType(Function() As Table
                                        Dim t = DateTime.Now
                                        Dim tbl = New Table(_lua)

                                        tbl("hour") = t.Hour       ' 0..23
                                        tbl("min") = t.Minute      ' 0..59
                                        tbl("sec") = t.Second      ' 0..59

                                        Return tbl
                                    End Function, Func(Of Table))

        ' Lua helper to control LED panels by name (state = ON/OFF/BAD)
        _lua.Globals("setled") = CType(Sub(ledName As String, stateText As String)
                                           ledName = If(ledName, "").Trim()
                                           stateText = If(stateText, "").Trim()
                                           If ledName = "" OrElse stateText = "" Then Exit Sub

                                           SetLedFromSpec($"{ledName}={stateText}")
                                       End Sub, Action(Of String, String))


        _luaReady = True
    End Sub

    Private Sub RunLua(luaCode As String)
        EnsureLua()

        _luaCancelRequested = False   ' Clear any previous stop request

        Try
            _lua.DoString(luaCode)

        Catch ex As OperationCanceledException
            AppendLog("[LUA] Script stopped")

        Catch ex As Exception
            AppendLog("[LUA ERROR] " & ex.Message)
        End Try
    End Sub



    Private Function GetTextAreaText(name As String) As String
        Dim tb As TextBoxBase = Nothing
        If TextAreaByName.TryGetValue(name, tb) AndAlso tb IsNot Nothing Then
            Return tb.Text
        End If
        Return Nothing
    End Function


    ' Read "key=value" params from a ; separated config line
    Private Function GetParam(lineText As String, key As String) As String
        For Each part In lineText.Split(";"c)
            Dim p = part.Trim()
            If p.Length = 0 Then Continue For

            Dim idx = p.IndexOf("="c)
            If idx <= 0 Then Continue For

            Dim k = p.Substring(0, idx).Trim()
            If k.Equals(key, StringComparison.OrdinalIgnoreCase) Then
                Return p.Substring(idx + 1).Trim()
            End If
        Next
        Return ""
    End Function



End Class