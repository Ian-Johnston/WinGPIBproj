' User customizeable tab - Lua support

Imports System.Globalization
Imports MoonSharp.Interpreter
Imports MoonSharp.Interpreter.Interop


Partial Class Formtest

    ' Lua engine instance
    Private _lua As Script
    Private _luaReady As Boolean = False

    Private Sub EnsureLua()
        If _luaReady Then Exit Sub

        ' Sandbox
        UserData.RegistrationPolicy = InteropRegistrationPolicy.Default

        _lua = New Script(CoreModules.Preset_SoftSandbox)

        ' Expose small API to Lua
        _lua.Globals("log") = CType(Sub(msg As Object)
                                        AppendTriggerLog("[LUA] " & If(msg, "").ToString())
                                    End Sub, Action(Of Object))

        _lua.Globals("fire") = CType(Sub(btnName As String)
                                         If String.IsNullOrWhiteSpace(btnName) Then Exit Sub
                                         Dim b As Button = Nothing
                                         If BtnByName.TryGetValue(btnName, b) AndAlso b IsNot Nothing Then
                                             If b.InvokeRequired Then
                                                 b.BeginInvoke(Sub() b.PerformClick())
                                             Else
                                                 b.PerformClick()
                                             End If

                                         Else
                                             AppendTriggerLog("[LUA] fire(): button not found: " & btnName)
                                         End If
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
                                                 AppendTriggerLog("[LUA] send(): device not available: " & devName)
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
                                                  AppendTriggerLog("[LUA] query(): device not available: " & devName)
                                                  Return ""
                                              End If

                                              Dim q As IODevices.IOQuery = Nothing
                                              Dim fullCmd As String = cmd & TermStr2()
                                              Dim status As Integer = dev.QueryBlocking(fullCmd, q, False)

                                              If status = 0 AndAlso q IsNot Nothing Then
                                                  Return q.ResponseAsString.Trim()
                                              End If

                                              If q IsNot Nothing Then
                                                  AppendTriggerLog("[LUA] query(): ERR " & status & ": " & q.errmsg)
                                              Else
                                                  AppendTriggerLog("[LUA] query(): ERR " & status & " (no IOQuery)")
                                              End If

                                              Return ""
                                          End If
                                      End Function, Func(Of String, String, String))



        _luaReady = True
    End Sub

    Private Sub RunLua(luaCode As String)
        EnsureLua()
        Try
            _lua.DoString(luaCode)
        Catch ex As Exception
            AppendTriggerLog("[LUA ERROR] " & ex.Message)
        End Try
    End Sub


    Private Function GetTextAreaText(name As String) As String
        Dim tb As TextBoxBase = Nothing
        If TextAreaByName.TryGetValue(name, tb) AndAlso tb IsNot Nothing Then
            Return tb.Text
        End If
        Return Nothing
    End Function



End Class