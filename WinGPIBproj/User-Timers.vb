' User customizeable tab - Timers


Partial Class Formtest


    ' --- Timer15 (QUERIESTOFILE) state ---
    Private Auto15DeviceName As String = ""
    Private Auto15ScriptBoxName As String = ""
    Private Auto15FilePathControl As String = ""
    Private Auto15ResultControl As String = ""


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


End Class