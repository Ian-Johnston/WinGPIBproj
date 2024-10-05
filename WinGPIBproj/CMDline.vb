
' Command Line Interface

Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Partial Class Formtest

    Dim CMDlineOp As Boolean = False



    Private Sub TextBoxDev1CMD_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxDev1CMD.KeyPress

        ' Detect if ESCAPE pressed - Empty current line being edited
        'If e.KeyChar = Convert.ToChar(27) Then
        'TextBoxDev1CMD.AppendText(Environment.NewLine)
        'End If

        CMDlineOp = True

        Dim currline As Integer = TextBoxDev1CMD.GetLineFromCharIndex(TextBoxDev1CMD.SelectionStart)        ' Get current line number, first line is 0	
        Dim CMDtext As String


        ' User will have typed a new command, so when ENTER is pressed and the user text is forcing Send & Return or just Send then manipulate the radio buttons
        If e.KeyChar = Convert.ToChar(13) Then

            CMDtext = TextBoxDev1CMD.Lines(currline)
            If CMDtext = "#QA" Then
                CheckBoxDev1Query.Checked = True
                CheckBoxDev1Async.Checked = False
                Exit Sub
            End If
            If CMDtext = "#SA" Then
                CheckBoxDev1Query.Checked = False
                CheckBoxDev1Async.Checked = True
                Exit Sub
            End If

        End If


        ' Send & return
        If CheckBoxDev1Query.Checked = True Then

            ' User will have typed a new command, so when ENTER is pressed
            If e.KeyChar = Convert.ToChar(13) Then

                ' delete empty line after ENTER is pressed when Query Async only.
                If e.KeyChar = ChrW(Keys.Enter) And CheckBoxDev1Query.Checked = True Then
                    RemoveCurrentLineDev1()
                    ' Consume the Enter key press to prevent it from being added to the textbox
                    e.Handled = True
                End If

                CMDtext = TextBoxDev1CMD.Lines(currline)
                'MsgBox(CMDtext)

                If CMDtext <> "" Then
                    If Dev1PollingEnable.Checked = True Then
                        dev1.enablepoll = True
                    Else
                        dev1.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
                    End If

                    If Dev1STBMask.Text = "" Then
                        Dev1STBMask.Text = "16"
                    End If
                    dev1.MAVmask = Val(Dev1STBMask.Text)
                    If Dev1STBMask.Text = "0" Then
                        dev1.enablepoll = False
                        Dev1PollingEnable.Checked = False
                    End If

                    dev1.QueryAsync(TextBoxDev1CMD.Lines(currline) & TermStr2(), AddressOf Cbdev1, True)

                End If
            End If

        End If


        ' Send only
        If CheckBoxDev1Async.Checked = True Then

            ' Detect if ENTER pressed
            If e.KeyChar = Convert.ToChar(13) Then

                CMDtext = TextBoxDev1CMD.Lines(currline)
                'MsgBox(CMDtext)

                If CMDtext <> "" Then
                    dev1.SendAsync(TextBoxDev1CMD.Lines(currline), True)
                    txtr1astat.Text = "Send Async '" & txtq1c.Text
                End If

            End If

        End If

        'dev1.showmessages = True

    End Sub

    Private Sub RemoveCurrentLineDev1()
        ' Get the current line index
        Dim currentLineIndex As Integer = TextBoxDev1CMD.GetLineFromCharIndex(TextBoxDev1CMD.SelectionStart)

        ' Get the start index of the current line
        Dim startIndex As Integer = TextBoxDev1CMD.GetFirstCharIndexFromLine(currentLineIndex)

        ' Get the end index of the current line (excluding newline characters)
        Dim endIndex As Integer = TextBoxDev1CMD.GetFirstCharIndexFromLine(currentLineIndex + 1) - 1

        ' If endIndex is -1, it means it's the last line, so we get the index of the last character
        If endIndex = -1 Then
            endIndex = TextBoxDev1CMD.TextLength - 1
        End If

        ' Check if startIndex and endIndex are valid
        If startIndex < 0 OrElse endIndex < 0 Then
            ' Invalid indices, do nothing
            Return
        End If

        ' Calculate the length of the current line
        Dim lineLength As Integer = endIndex - startIndex + 1

        ' Remove the current line
        TextBoxDev1CMD.Text = TextBoxDev1CMD.Text.Remove(startIndex, lineLength)

        ' Set the cursor position to the beginning of the current line
        TextBoxDev1CMD.SelectionStart = startIndex
    End Sub

    Private Sub CheckBoxDev1Query_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxDev1Query.CheckedChanged

        If CheckBoxDev1Query.Checked = True Then
            CheckBoxDev1Async.Checked = False
        End If

    End Sub

    Private Sub CheckBoxDev1Async_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxDev1Async.CheckedChanged

        If CheckBoxDev1Async.Checked = True Then
            CheckBoxDev1Query.Checked = False
        End If

    End Sub



    Private Sub TextBoxDev2CMD_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxDev2CMD.KeyPress

        ' Detect if ESCAPE pressed - Empty current line being edited
        'If e.KeyChar = Convert.ToChar(27) Then
        'TextBoxDev1CMD.AppendText(Environment.NewLine)
        'End If

        CMDlineOp = True

        Dim currline As Integer = TextBoxDev2CMD.GetLineFromCharIndex(TextBoxDev2CMD.SelectionStart)        ' Get current line number, first line is 0	
        Dim CMDtext As String


        ' User will have typed a new command, so when ENTER is pressed and the user text is forcing Send & Return or just Send then manipulate the radio buttons
        If e.KeyChar = Convert.ToChar(13) Then

            CMDtext = TextBoxDev2CMD.Lines(currline)
            If CMDtext = "#QA" Then
                CheckBoxDev2Query.Checked = True
                CheckBoxDev2Async.Checked = False
                Exit Sub
            End If
            If CMDtext = "#SA" Then
                CheckBoxDev2Query.Checked = False
                CheckBoxDev2Async.Checked = True
                Exit Sub
            End If

        End If


        ' Send & return
        If CheckBoxDev2Query.Checked = True Then

            ' User will have typed a new command, so when ENTER is pressed
            If e.KeyChar = Convert.ToChar(13) Then

                ' delete empty line after ENTER is pressed when Query Async only.
                If e.KeyChar = ChrW(Keys.Enter) And CheckBoxDev2Query.Checked = True Then

                    RemoveCurrentLineDev2()
                    ' Consume the Enter key press to prevent it from being added to the textbox
                    e.Handled = True
                End If

                CMDtext = TextBoxDev2CMD.Lines(currline)
                'MsgBox(CMDtext)

                If CMDtext <> "" Then
                    If Dev2PollingEnable.Checked = True Then
                        dev2.enablepoll = True
                    Else
                        dev2.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
                    End If

                    If Dev2STBMask.Text = "" Then
                        Dev2STBMask.Text = "16"
                    End If
                    dev2.MAVmask = Val(Dev2STBMask.Text)
                    If Dev2STBMask.Text = "0" Then
                        dev2.enablepoll = False
                        Dev2PollingEnable.Checked = False
                    End If

                    dev2.QueryAsync(TextBoxDev2CMD.Lines(currline) & TermStr(), AddressOf Cbdev2, True)

                End If
            End If

        End If


        ' Send only
        If CheckBoxDev2Async.Checked = True Then

            ' Detect if ENTER pressed
            If e.KeyChar = Convert.ToChar(13) Then

                CMDtext = TextBoxDev2CMD.Lines(currline)
                'MsgBox(CMDtext)

                If CMDtext <> "" Then
                    dev2.SendAsync(TextBoxDev2CMD.Lines(currline), True)
                    txtr2astat.Text = "Send Async '" & txtq2c.Text
                End If

            End If

        End If

    End Sub

    Private Sub RemoveCurrentLineDev2()

        ' Get the current line index
        Dim currentLineIndex As Integer = TextBoxDev2CMD.GetLineFromCharIndex(TextBoxDev2CMD.SelectionStart)

        ' Get the start index of the current line
        Dim startIndex As Integer = TextBoxDev2CMD.GetFirstCharIndexFromLine(currentLineIndex)

        ' Get the end index of the current line (excluding newline characters)
        Dim endIndex As Integer = TextBoxDev2CMD.GetFirstCharIndexFromLine(currentLineIndex + 1) - 1

        ' If endIndex is -1, it means it's the last line, so we get the index of the last character
        If endIndex = -1 Then
            endIndex = TextBoxDev2CMD.TextLength - 1
        End If

        ' Check if startIndex and endIndex are valid
        If startIndex < 0 OrElse endIndex < 0 Then
            ' Invalid indices, do nothing
            Return
        End If

        ' Calculate the length of the current line
        Dim lineLength As Integer = endIndex - startIndex + 1

        ' Remove the current line
        TextBoxDev2CMD.Text = TextBoxDev2CMD.Text.Remove(startIndex, lineLength)

        ' Set the cursor position to the beginning of the current line
        TextBoxDev2CMD.SelectionStart = startIndex
    End Sub

    Private Sub CheckBoxDev2Query_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxDev2Query.CheckedChanged

        If CheckBoxDev2Query.Checked = True Then
            CheckBoxDev2Async.Checked = False
        End If

    End Sub

    Private Sub CheckBoxDev2Async_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxDev2Async.CheckedChanged

        If CheckBoxDev2Async.Checked = True Then
            CheckBoxDev2Query.Checked = False
        End If

    End Sub

    Private Sub CMD1clear_Click(sender As Object, e As EventArgs) Handles CMD1clear.Click

        TextBoxDev1CMD.Text = "READY!" + Environment.NewLine

    End Sub

    Private Sub CMD2clear_Click(sender As Object, e As EventArgs) Handles CMD2clear.Click

        TextBoxDev2CMD.Text = "READY!" + Environment.NewLine

    End Sub


End Class