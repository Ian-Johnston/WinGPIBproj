' User customizeable tab

Imports IODevices
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.IO

Partial Class Formtest

    ' INI Format:
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

    Private Sub BuildCustomGuiFromText(def As String)

        GroupBoxCustom.Controls.Clear()

        Dim autoY As Integer = 10   ' fallback vertical stacking

        For Each rawLine In def.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            Dim line = rawLine.Trim()
            If line = "" OrElse line.StartsWith(";") Then Continue For

            Dim parts = line.Split(";"c)
            If parts.Length < 2 Then Continue For

            Dim typeStr = parts(0).Trim().ToUpperInvariant()

            Select Case typeStr

            ' ============================
            '   SET GROUPBOX TITLE
            '   GROUPBOX;Some title text
            ' ============================
                Case "GROUPBOX"
                    Dim title = parts(1).Trim()
                    GroupBoxCustom.Text = title   ' <-- your custom GroupBox
                    Continue For

            ' ======================================
            '   CREATE A LABELED TEXTBOX FROM INI
            '   TEXTBOX;Name;Label text;X;Y
            ' ======================================
                Case "TEXTBOX"
                    If parts.Length < 3 Then Continue For

                    Dim tbName = parts(1).Trim()
                    Dim labelText = parts(2).Trim()

                    Dim x As Integer = 10
                    Dim y As Integer = autoY

                    If parts.Length >= 5 Then
                        Integer.TryParse(parts(3).Trim(), x)
                        Integer.TryParse(parts(4).Trim(), y)
                    End If

                    ' Label
                    Dim lbl As New Label()
                    lbl.Text = labelText
                    lbl.AutoSize = True
                    lbl.Location = New Point(x, y)
                    GroupBoxCustom.Controls.Add(lbl)

                    ' TextBox to the right of label
                    Dim tb As New TextBox()
                    tb.Name = tbName
                    tb.Width = 120
                    tb.Location = New Point(lbl.Right + 5, y - 2)
                    GroupBoxCustom.Controls.Add(tb)

                    ' Only advance autoY if coordinates not explicitly given
                    If parts.Length < 5 Then
                        autoY += Math.Max(lbl.Height, tb.Height) + 5
                    End If

                    Continue For

            ' ======================================
            '   BUTTON LINES
            '   BUTTON;Caption;Action;Device;Command;ValueCtl;ResultCtl;X;Y
            ' ======================================
                Case "BUTTON"
                    If parts.Length < 5 Then Continue For

                    Dim caption = parts(1).Trim()
                    Dim action = parts(2).Trim().ToUpperInvariant()
                    Dim deviceName = parts(3).Trim()
                    Dim commandOrPrefix = parts(4).Trim()
                    Dim valueControlName As String = If(parts.Length > 5, parts(5).Trim(), "")
                    Dim resultControlName As String = If(parts.Length > 6, parts(6).Trim(), "")

                    ' --- Position (X,Y) ---
                    Dim x As Integer = 10
                    Dim y As Integer = autoY

                    If parts.Length >= 9 Then
                        Integer.TryParse(parts(7).Trim(), x)
                        Integer.TryParse(parts(8).Trim(), y)
                    End If

                    Dim b As New Button()
                    b.Text = caption
                    b.Tag = $"{action}|{deviceName}|{commandOrPrefix}|{valueControlName}|{resultControlName}"

                    ' --- Size (Width,Height) ---
                    Dim w As Integer = 160
                    Dim h As Integer = b.Height   ' default button height

                    If parts.Length >= 11 Then
                        Integer.TryParse(parts(9).Trim(), w)
                        Integer.TryParse(parts(10).Trim(), h)
                    End If

                    b.Location = New Point(x, y)
                    b.Size = New Size(w, h)

                    AddHandler b.Click, AddressOf CustomButton_Click
                    GroupBoxCustom.Controls.Add(b)

                    If parts.Length < 9 Then
                        autoY += b.Height + 5      ' uses final height (after any override)
                    End If


            End Select

        Next

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

        ' 1) First look inside PanelCustom (INI-generated stuff)
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


    Private Sub CustomButton_Click(sender As Object, e As EventArgs)

        Dim b = DirectCast(sender, Button)
        Dim meta = CStr(b.Tag)
        Dim parts = meta.Split("|"c)
        If parts.Length < 3 Then Exit Sub

        Dim action = parts(0)                 ' SEND / SENDVALUE / QUERY
        Dim deviceName = parts(1)             ' "dev1" / "dev2"
        Dim commandOrPrefix = parts(2)
        Dim valueControlName = If(parts.Length > 3, parts(3), "")
        Dim resultControlName = If(parts.Length > 4, parts(4), "")

        Dim dev = GetDeviceByName(deviceName)
        If dev Is Nothing Then
            MessageBox.Show("Device not available: " & deviceName)
            Exit Sub
        End If



        Select Case action

            Case "SEND"
                dev.SendAsync(commandOrPrefix, True)

            Case "SENDVALUE"
                Dim valCtrl = TryCast(GetControlByName(valueControlName), TextBox)
                If valCtrl Is Nothing Then
                    MessageBox.Show("Value control not found: " & valueControlName)
                    Exit Sub
                End If

                Dim cmd = commandOrPrefix & valCtrl.Text.Trim()
                dev.SendAsync(cmd, True)


            Case "QUERY"

                Dim target = GetControlByName(resultControlName)
                If target Is Nothing Then
                    MessageBox.Show("Result control not found: " & resultControlName)
                    Exit Sub
                End If

                Dim q As IOQuery = Nothing
                Dim status As Integer

                Dim fullCmd As String = commandOrPrefix & TermStr2()   ' e.g. "USE?" + terminator

                status = dev.QueryBlocking(fullCmd, q, False)

                ' Always show RAW response in txtr2b for debugging
                If q IsNot Nothing Then
                    txtr2b.Text = q.ResponseAsString
                Else
                    txtr2b.Text = "[no IOQuery]"
                End If

                Dim outText As String
                If status = 0 AndAlso q IsNot Nothing Then
                    ' Trim so "     0" becomes "0"
                    outText = q.ResponseAsString.Trim()
                ElseIf q IsNot Nothing Then
                    outText = "ERR " & status & ": " & q.errmsg
                Else
                    outText = "ERR " & status & " (no IOQuery)"
                End If

                ' Now write to the INI-created target control
                If TypeOf target Is TextBox Then
                    DirectCast(target, TextBox).Text = outText
                ElseIf TypeOf target Is Label Then
                    DirectCast(target, Label).Text = outText
                End If

        End Select

    End Sub



End Class
