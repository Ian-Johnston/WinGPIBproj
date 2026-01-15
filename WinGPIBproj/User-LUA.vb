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
                                              ' Use the same raw fast-path we use for DATASOURCE etc.
                                              Return NativeQuery(devName, cmd, requireRaw:=True)
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
                                              Debug.WriteLine("BLOCKING DetermineQuery: " & status)


                                              If status = 0 AndAlso q IsNot Nothing Then
                                                  Return respNorm.Trim()
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

        ' Expose Polynomials
        _lua.Globals("polyeval") = DynValue.NewCallback(AddressOf Lua_PolyEval)
        _lua.Globals("solve2") = DynValue.NewCallback(AddressOf Lua_Solve2)
        _lua.Globals("solve3") = DynValue.NewCallback(AddressOf Lua_Solve3)
        _lua.Globals("solve4") = DynValue.NewCallback(AddressOf Lua_Solve4)
        _lua.Globals("realroots") = DynValue.NewCallback(AddressOf Lua_RealRoots)

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


    ' MoonSharp callback: polyeval(x, a0, a1, a2, ...)
    ' - a0..an are coefficients for: a0 + a1*x + a2*x^2 + ...
    Private Function Lua_PolyEval(ctx As ScriptExecutionContext, args As CallbackArguments) As DynValue
        Try
            If args Is Nothing OrElse args.Count < 2 Then
                Return DynValue.NewNumber(0R)
            End If

            Dim x As Double = LuaArgToDouble(args(0))
            Dim coeffs(args.Count - 2) As Double
            For i As Integer = 1 To args.Count - 1
                coeffs(i - 1) = LuaArgToDouble(args(i))
            Next

            Dim y As Double = PolynomialSolvers.PolyEval(x, coeffs)
            Return DynValue.NewNumber(y)

        Catch ex As Exception
            AppendLog("[LUA] PolyEval ERROR: " & ex.Message)
            Return DynValue.NewNumber(0R)
        End Try
    End Function


    Private Function LuaArgToDouble(v As DynValue) As Double
        If v Is Nothing Then Return Double.NaN
        If v.Type = DataType.Number Then Return v.Number
        If v.Type = DataType.String Then
            Dim d As Double
            If Double.TryParse(v.String, NumberStyles.Float, CultureInfo.InvariantCulture, d) Then
                Return d
            End If
        End If
        Return Double.NaN
    End Function


    Private Function Lua_Solve2(ctx As ScriptExecutionContext, args As CallbackArguments) As DynValue
        Dim a = LuaArgToDouble(args(0)) : Dim b = LuaArgToDouble(args(1)) : Dim c = LuaArgToDouble(args(2))
        Dim r = PolynomialSolvers.SolveQuadratic(a, b, c)
        Return RootsToLuaTable(r)
    End Function

    Private Function Lua_Solve3(ctx As ScriptExecutionContext, args As CallbackArguments) As DynValue
        Dim a = LuaArgToDouble(args(0)) : Dim b = LuaArgToDouble(args(1)) : Dim c = LuaArgToDouble(args(2)) : Dim d = LuaArgToDouble(args(3))
        Dim r = PolynomialSolvers.SolveCubic(a, b, c, d)
        Return RootsToLuaTable(r)
    End Function

    Private Function Lua_Solve4(ctx As ScriptExecutionContext, args As CallbackArguments) As DynValue
        Dim a = LuaArgToDouble(args(0)) : Dim b = LuaArgToDouble(args(1)) : Dim c = LuaArgToDouble(args(2)) : Dim d = LuaArgToDouble(args(3)) : Dim e = LuaArgToDouble(args(4))
        Dim r = PolynomialSolvers.SolveQuartic(a, b, c, d, e)
        Return RootsToLuaTable(r)
    End Function

    Private Function Lua_RealRoots(ctx As ScriptExecutionContext, args As CallbackArguments) As DynValue
        ' realroots( rootsTable [, tol] )
        Dim tol As Double = 0.0000000001
        If args.Count >= 2 AndAlso args(1).Type = DataType.Number Then tol = args(1).Number

        Dim t = args(0).Table
        Dim lst As New List(Of Global.System.Numerics.Complex)

        For Each pair In t.Pairs
            Dim s As String = pair.Value.String
            lst.Add(ParseComplexString(s))
        Next

        Dim rr = PolynomialSolvers.RealRoots(lst.ToArray(), tol)
        Dim outT = New Table(_lua)
        For i As Integer = 0 To rr.Length - 1
            outT.Set(i + 1, DynValue.NewNumber(rr(i)))
        Next
        Return DynValue.NewTable(outT)
    End Function

    Private Function RootsToLuaTable(r() As Global.System.Numerics.Complex) As DynValue
        Dim t As New Table(_lua)
        For i As Integer = 0 To r.Length - 1
            t.Set(i + 1, DynValue.NewString(FormatComplex(r(i))))
        Next
        Return DynValue.NewTable(t)
    End Function

    Private Function FormatComplex(z As Global.System.Numerics.Complex) As String
        ' Fixed-point so ParseComplexString doesn't get confused by E-notation
        Dim a = z.Real.ToString("0.################", CultureInfo.InvariantCulture)
        Dim bi = z.Imaginary

        ' Clamp tiny imag to 0
        If Math.Abs(bi) < 0.000000000001 Then bi = 0R

        Dim b = Math.Abs(bi).ToString("0.################", CultureInfo.InvariantCulture)
        Dim sign = If(bi >= 0, "+", "-")
        Return a & sign & b & "i"
    End Function

    Private Function ParseComplexString(s As String) As Global.System.Numerics.Complex
        ' expects "a+bi" or "a-bi" (created by FormatComplex)
        s = s.Replace(" ", "").Trim()
        Dim idx = s.LastIndexOf("+"c)
        If idx < 0 Then idx = s.LastIndexOf("-"c, s.Length - 2) ' allow leading minus
        If idx <= 0 Then Return New Global.System.Numerics.Complex(Double.Parse(s, CultureInfo.InvariantCulture), 0)

        Dim aStr = s.Substring(0, idx)
        Dim bStr = s.Substring(idx, s.Length - idx).Replace("i", "")
        Dim a = Double.Parse(aStr, CultureInfo.InvariantCulture)
        Dim b = Double.Parse(bStr, CultureInfo.InvariantCulture)
        Return New Global.System.Numerics.Complex(a, b)
    End Function


End Class