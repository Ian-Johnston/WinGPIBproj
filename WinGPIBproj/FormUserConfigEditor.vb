Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System.Runtime.InteropServices


Public Class FormUserConfigEditor
    Inherits Form

    ' === Master keyword lists for syntax highlighting & navigation ===

    ' Top-level block types (lines like "RADIO;" / "STATSPANEL;" etc.)
    Private Shared ReadOnly BlockKeywords As String() = {
        "BIGTEXT",
        "BOOTCOMMANDS",
        "BUTTON",
        "CALC",
        "CHART",
        "CHECKBOX",
        "DATASOURCE",
        "DROPDOWN",
        "GROUPBOX",
        "HISTORYGRID",
        "INVISIBILITY",
        "KEYPAD",
        "LABEL",
        "LED",
        "LUASCRIPTBEGIN",
        "MULTIBUTTON",
        "RADIO",
        "RADIOGROUP",
        "SLIDER",
        "SPINNER",
        "STAT",
        "STATSPANEL",
        "TAB",
        "TEXTAREA",
        "TEXTBOX",
        "TOGGLE",
        "TOGGLEDUAL",
        "TRIGGER"
    }

    ' Parameter keys of the form key=... used across all configs
    Private Shared ReadOnly ParamKeys As String() = {
        "DATASAVE",
        "DelayPerCmd",
        "GpibEngineDev1",
        "GpibEngineDev2",
        "GpibEnginedev1",
        "action",
        "autoscale",
        "bad",
        "border",
        "caption",
        "captionpos",
        "color",
        "cols",
        "command",
        "commandlist",
        "commands",
        "decimal",
        "default",
        "determine",
        "detmap",
        "device",
        "do",
        "dp",
        "expr",
        "f",
        "format",
        "func",
        "gap",
        "group",
        "h",
        "if",
        "init",
        "innerH",
        "innerW",
        "innerX",
        "innerY",
        "items",
        "label",
        "labelsize",
        "left",
        "linewidth",
        "max",
        "maxpoints",
        "maxrows",
        "min",
        "mode",
        "name",
        "off",
        "on",
        "oncolor",
        "overload",
        "panel",
        "param",
        "period",
        "popup",
        "ppmref",
        "readonly",
        "ref",
        "result",
        "right",
        "scale",
        "sendval",
        "size",
        "step",
        "target",
        "targetlabel",
        "targets",
        "then",
        "units",
        "value",
        "w",
        "when",
        "x",
        "xstep",
        "y",
        "ymax",
        "ymin"
    }


    Private Const WM_SETREDRAW As Integer = &HB
    Private Const EM_GETFIRSTVISIBLELINE As Integer = &HCE
    Private Const EM_LINESCROLL As Integer = &HB6
    Private WithEvents lstBlocks As New ListBox()
    Private ReadOnly _configPath As String
    Private WithEvents txtEditor As New RichTextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnClose As New Button()
    Private WithEvents pnlLineNumbers As New Panel()

    ' Context menu + find bar
    Private ctxEditor As ContextMenuStrip
    Private editorHost As Panel
    Private findPanel As Panel
    Private txtFind As TextBox
    Private btnFindNext As Button
    Private btnFindPrev As Button
    Private btnFindClose As Button
    Private _lastFindText As String = ""
    Private _lastFindIndex As Integer = -1

    ' Optional: let caller know when a save happened (for auto-reload)
    Public Event ConfigSaved(path As String)


    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer,
                                    wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    Private Class BlockItem
        Public Property Display As String
        Public Property CharIndex As Integer

        Public Overrides Function ToString() As String
            Return Display
        End Function
    End Class


    Public Sub New(configPath As String)
        _configPath = configPath

        Me.Text = "User Config Editor - " & Path.GetFileName(configPath)
        Me.StartPosition = FormStartPosition.CenterParent

        ' --- Load size from settings or fall back to default ---
        Dim defaultSize As New Size(900, 700)
        Dim s As String = My.Settings.UserGui_Editor
        Dim w As Integer = defaultSize.Width
        Dim h As Integer = defaultSize.Height

        If Not String.IsNullOrWhiteSpace(s) Then
            Dim parts = s.Split(","c)
            If parts.Length = 2 Then
                Dim tmpW, tmpH As Integer
                If Integer.TryParse(parts(0).Trim(), tmpW) AndAlso
               Integer.TryParse(parts(1).Trim(), tmpH) AndAlso
               tmpW > 400 AndAlso tmpH > 300 Then   ' basic sanity
                    w = tmpW
                    h = tmpH
                End If
            End If
        End If

        Me.Size = New Size(w, h)
        Me.Font = New Font("Segoe UI", 9.0F)
        Me.BackColor = Color.Black

        ' ---- Editor ----
        txtEditor.Dock = DockStyle.Fill
        txtEditor.Font = New Font("Consolas", 9.0F)
        txtEditor.AcceptsTab = True
        txtEditor.WordWrap = False
        txtEditor.HideSelection = False
        txtEditor.BackColor = Color.Black
        txtEditor.ForeColor = Color.Gainsboro

        ' ---- Context menu for editor ----
        ctxEditor = New ContextMenuStrip()

        Dim miUndo = New ToolStripMenuItem("Undo", Nothing, AddressOf CtxUndo)
        Dim miCut = New ToolStripMenuItem("Cut", Nothing, AddressOf CtxCut)
        Dim miCopy = New ToolStripMenuItem("Copy", Nothing, AddressOf CtxCopy)
        Dim miPaste = New ToolStripMenuItem("Paste", Nothing, AddressOf CtxPaste)
        Dim miDelete = New ToolStripMenuItem("Delete", Nothing, AddressOf CtxDelete)
        Dim miSelectAll = New ToolStripMenuItem("Select All", Nothing, AddressOf CtxSelectAll)
        Dim miDupLine = New ToolStripMenuItem("Duplicate Line", Nothing, AddressOf CtxDuplicateLine)
        Dim miSave = New ToolStripMenuItem("Save", Nothing, Sub() SaveConfigFile())

        ctxEditor.Items.AddRange(New ToolStripItem() {
        miUndo,
        miSave,
        New ToolStripSeparator(),
        miCut, miCopy, miPaste, miDelete,
        New ToolStripSeparator(),
        miSelectAll,
        New ToolStripSeparator(),
        miDupLine
    })

        txtEditor.ContextMenuStrip = ctxEditor

        ' ---- Bottom panel with buttons ----
        Dim panelButtons As New Panel()
        panelButtons.Dock = DockStyle.Bottom
        panelButtons.Height = 40
        panelButtons.BackColor = SystemColors.Control

        btnSave.Text = "Save"
        btnSave.Width = 80
        btnSave.Height = 24
        btnSave.Left = 10
        btnSave.Top = 8
        btnSave.FlatStyle = FlatStyle.Standard
        btnSave.UseVisualStyleBackColor = True

        btnClose.Text = "Close"
        btnClose.Width = 80
        btnClose.Height = 24
        btnClose.Left = btnSave.Right + 10
        btnClose.Top = 8
        btnClose.FlatStyle = FlatStyle.Standard
        btnClose.UseVisualStyleBackColor = True

        panelButtons.Controls.Add(btnSave)
        panelButtons.Controls.Add(btnClose)

        ' ---- Right panel with block list ----
        Dim panelRight As New Panel()
        panelRight.Dock = DockStyle.Right
        panelRight.Width = 200
        panelRight.BackColor = Color.Black

        Dim lblBlocks As New Label()
        lblBlocks.Text = "Blocks"
        lblBlocks.AutoSize = False
        lblBlocks.Dock = DockStyle.Top
        lblBlocks.Height = 20
        lblBlocks.TextAlign = ContentAlignment.MiddleLeft
        lblBlocks.ForeColor = Color.Gainsboro
        lblBlocks.BackColor = Color.Black

        lstBlocks.Dock = DockStyle.Fill
        lstBlocks.BackColor = Color.Black
        lstBlocks.ForeColor = Color.Gainsboro
        lstBlocks.BorderStyle = BorderStyle.FixedSingle

        panelRight.Controls.Add(lstBlocks)
        panelRight.Controls.Add(lblBlocks)

        ' ---- Left panel for line numbers ----
        pnlLineNumbers.Dock = DockStyle.Left
        pnlLineNumbers.Width = 55
        pnlLineNumbers.BackColor = Color.FromArgb(25, 25, 25)
        AddHandler pnlLineNumbers.Paint, AddressOf pnlLineNumbers_Paint

        ' ---- Find bar (top) ----
        findPanel = New Panel()
        findPanel.Dock = DockStyle.Top
        findPanel.Height = 28
        findPanel.BackColor = SystemColors.Control    ' << light, normal WinForms bar
        findPanel.Visible = False

        Dim lblFind As New Label()
        lblFind.Text = "Find:"
        lblFind.AutoSize = True
        lblFind.ForeColor = SystemColors.ControlText  ' << normal text colour
        lblFind.BackColor = Color.Transparent
        lblFind.Left = 8
        lblFind.Top = 7

        txtFind = New TextBox()
        txtFind.Left = 50
        txtFind.Top = 4
        txtFind.Width = 220
        txtFind.BackColor = SystemColors.Window       ' << normal textbox colours
        txtFind.ForeColor = SystemColors.WindowText
        AddHandler txtFind.KeyDown,
    Sub(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            DoFind(True)   ' same as Next
        End If
    End Sub

        btnFindNext = New Button()
        btnFindNext.Text = "Next"
        btnFindNext.Width = 60
        btnFindNext.Height = 22
        btnFindNext.Left = txtFind.Right + 8
        btnFindNext.Top = 3

        btnFindPrev = New Button()
        btnFindPrev.Text = "Prev"
        btnFindPrev.Width = 60
        btnFindPrev.Height = 22
        btnFindPrev.Left = btnFindNext.Right + 4
        btnFindPrev.Top = 3

        btnFindClose = New Button()
        btnFindClose.Text = "X"
        btnFindClose.Width = 30
        btnFindClose.Height = 22
        btnFindClose.Left = btnFindPrev.Right + 4
        btnFindClose.Top = 3

        ' NORMALISE FIND BUTTONS TO STANDARD LOOK
        For Each b As Button In {btnFindNext, btnFindPrev, btnFindClose}
            b.FlatStyle = FlatStyle.Standard
            b.UseVisualStyleBackColor = True          ' << let OS theme draw them
            b.BackColor = SystemColors.Control        ' (mostly ignored if visual styles on)
            b.ForeColor = SystemColors.ControlText
        Next


        AddHandler btnFindNext.Click, AddressOf BtnFindNext_Click
        AddHandler btnFindPrev.Click, AddressOf BtnFindPrev_Click
        AddHandler btnFindClose.Click, AddressOf BtnFindClose_Click

        findPanel.Controls.Add(lblFind)
        findPanel.Controls.Add(txtFind)
        findPanel.Controls.Add(btnFindNext)
        findPanel.Controls.Add(btnFindPrev)
        findPanel.Controls.Add(btnFindClose)

        ' ---- Host panel for editor / line numbers / blocks / find ----
        editorHost = New Panel()
        editorHost.Dock = DockStyle.Fill
        editorHost.BackColor = Color.Black

        ' Add in dock-friendly order inside host:
        editorHost.Controls.Add(txtEditor)       ' FILL
        editorHost.Controls.Add(pnlLineNumbers)  ' LEFT
        editorHost.Controls.Add(panelRight)      ' RIGHT
        editorHost.Controls.Add(findPanel)       ' TOP

        ' ---- Add to form (bottom bar last) ----
        Me.Controls.Add(editorHost)
        Me.Controls.Add(panelButtons)

        ' Load file on start
        LoadConfigFile()
    End Sub


    Private Sub LoadConfigFile()
        Try
            If File.Exists(_configPath) Then
                txtEditor.Text = File.ReadAllText(_configPath, Encoding.UTF8)
            Else
                txtEditor.Text = "; New User config file" & Environment.NewLine
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading config file:" & Environment.NewLine &
                        ex.Message,
                        "Load Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
        End Try

        ColorizeConfig()
        RebuildBlockList()
    End Sub


    Private Sub SaveConfigFile()
        Try
            File.WriteAllText(_configPath, txtEditor.Text, Encoding.UTF8)
            RaiseEvent ConfigSaved(_configPath)
        Catch ex As Exception
            MessageBox.Show("Error saving config file:" & Environment.NewLine &
                            ex.Message,
                            "Save Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
        End Try
    End Sub

    ' === Button handlers ===

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        SaveConfigFile()
        ' No recolour / no block-list rebuild here to avoid any jumping
    End Sub


    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ' === Simple syntax highlighting tuned for your config ===

    Private _isColorizing As Boolean = False

    Private Sub ColorizeConfig()
        If _isColorizing Then Return
        _isColorizing = True

        ' Remember caret & selection
        Dim selStart = txtEditor.SelectionStart
        Dim selLength = txtEditor.SelectionLength

        ' Remember current first visible line
        Dim firstVisibleBefore As Integer = SendMessage(txtEditor.Handle,
                                                   EM_GETFIRSTVISIBLELINE,
                                                   IntPtr.Zero,
                                                   IntPtr.Zero).ToInt32()

        ' Freeze redraw
        SendMessage(txtEditor.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero)

        ' ---- reset colours ----
        txtEditor.SelectAll()
        txtEditor.SelectionColor = Color.Gainsboro
        txtEditor.SelectionFont = txtEditor.Font

        Dim text = txtEditor.Text
        Dim length = text.Length

        ' 1) Comments + big section headers + inline ;; comments
        Dim idx As Integer = 0
        While idx < length
            Dim lineStart As Integer = idx
            Dim lineEnd As Integer = text.IndexOf(vbLf, idx)
            If lineEnd = -1 Then lineEnd = length

            Dim lineTextFull = text.Substring(lineStart, lineEnd - lineStart)

            ' --- FULL line comment (starts with ;) ---
            Dim firstNonSpace As Integer = lineStart
            While firstNonSpace < lineEnd AndAlso Char.IsWhiteSpace(text(firstNonSpace))
                firstNonSpace += 1
            End While

            If firstNonSpace < lineEnd AndAlso text(firstNonSpace) = ";"c Then
                If lineTextFull.Contains("########") Then
                    txtEditor.[Select](firstNonSpace, lineEnd - firstNonSpace)
                    txtEditor.SelectionColor = Color.DeepSkyBlue
                    txtEditor.SelectionFont = New Font(txtEditor.Font, FontStyle.Bold)
                Else
                    txtEditor.[Select](firstNonSpace, lineEnd - firstNonSpace)
                    txtEditor.SelectionColor = Color.LimeGreen
                    txtEditor.SelectionFont = txtEditor.Font
                End If

            Else
                ' --- INLINE ;; comment support ---
                Dim inlinePos = lineTextFull.IndexOf(";;", StringComparison.Ordinal)
                If inlinePos >= 0 Then
                    Dim start = lineStart + inlinePos
                    txtEditor.[Select](start, lineEnd - start)
                    txtEditor.SelectionColor = Color.LimeGreen
                    txtEditor.SelectionFont = txtEditor.Font
                End If
            End If

            idx = lineEnd + 1
        End While

        ' 2) Block keywords
        For Each kw In BlockKeywords
            Dim search As String = kw & ";"
            Dim pos As Integer = 0
            While True
                pos = text.IndexOf(search, pos, StringComparison.OrdinalIgnoreCase)
                If pos < 0 Then Exit While

                txtEditor.[Select](pos, kw.Length)
                txtEditor.SelectionColor = Color.Cyan
                txtEditor.SelectionFont = New Font(txtEditor.Font, FontStyle.Bold)

                pos += search.Length
            End While
        Next

        For Each kw In blockKeywords
            Dim search As String = kw & ";"
            Dim pos As Integer = 0
            While True
                pos = text.IndexOf(search, pos, StringComparison.OrdinalIgnoreCase)
                If pos < 0 Then Exit While

                txtEditor.[Select](pos, kw.Length)
                txtEditor.SelectionColor = Color.Cyan
                txtEditor.SelectionFont = New Font(txtEditor.Font, FontStyle.Bold)

                pos += search.Length
            End While
        Next

        ' 3) Parameter keys before '='
        For Each key In ParamKeys
            Dim search As String = key & "="
            Dim pos As Integer = 0
            While True
                pos = text.IndexOf(search, pos, StringComparison.OrdinalIgnoreCase)
                If pos < 0 Then Exit While

                txtEditor.[Select](pos, key.Length)
                txtEditor.SelectionColor = Color.Orange
                txtEditor.SelectionFont = txtEditor.Font

                pos += search.Length
            End While
        Next

        ' Restore caret/selection
        txtEditor.[Select](selStart, selLength)

        ' Restore scroll position (first visible line)
        Dim firstVisibleAfter As Integer = SendMessage(txtEditor.Handle,
                                                   EM_GETFIRSTVISIBLELINE,
                                                   IntPtr.Zero,
                                                   IntPtr.Zero).ToInt32()
        Dim delta As Integer = firstVisibleBefore - firstVisibleAfter
        If delta <> 0 Then
            SendMessage(txtEditor.Handle, EM_LINESCROLL, IntPtr.Zero, New IntPtr(delta))
        End If

        ' Re-enable redraw and repaint once
        SendMessage(txtEditor.Handle, WM_SETREDRAW, New IntPtr(1), IntPtr.Zero)
        txtEditor.Invalidate()

        _isColorizing = False
    End Sub


    Private Sub txtEditor_KeyDown(sender As Object, e As KeyEventArgs) Handles txtEditor.KeyDown

        If e.KeyCode <> Keys.Enter Then Return

        Dim caretPos As Integer = txtEditor.SelectionStart
        Dim currentLineIndex As Integer = txtEditor.GetLineFromCharIndex(caretPos)
        If currentLineIndex < 0 OrElse currentLineIndex >= txtEditor.Lines.Length Then Return

        Dim currentLine As String = txtEditor.Lines(currentLineIndex)
        Dim trimmed As String = currentLine.TrimStart()

        ' 1) If this line IS a block keyword (e.g. "RADIO;") → indent next
        Dim isBlockHeader As Boolean =
            BlockKeywords.Any(Function(k) trimmed.StartsWith(k & ";", StringComparison.OrdinalIgnoreCase))

        If isBlockHeader Then
            e.SuppressKeyPress = True
            Dim indent As String = "   "    ' 3-space default indent
            txtEditor.SelectedText = Environment.NewLine & indent
            Return
        End If

        ' 2) Otherwise: auto-indent based on existing indentation
        Dim leadingWhitespace As String =
            New String(currentLine.TakeWhile(Function(c) Char.IsWhiteSpace(c)).ToArray())

        If String.IsNullOrEmpty(leadingWhitespace) Then Return

        e.SuppressKeyPress = True
        txtEditor.SelectedText = Environment.NewLine & leadingWhitespace

    End Sub


    Private Sub RebuildBlockList()
        lstBlocks.Items.Clear()

        Dim lines = txtEditor.Lines
        If lines Is Nothing OrElse lines.Length = 0 Then Return

        ' If the Blocks list numbers are consistently off, adjust here.
        ' Currently your DATASOURCE; is 2 higher in the Blocks list,
        ' so we subtract 2 to align with what you see in the editor.
        Const BlockLineOffset As Integer = -2   ' tweak to -1 / 0 / etc if needed

        For i As Integer = 0 To lines.Length - 1
            Dim raw = lines(i)
            Dim trimmed = raw.TrimStart()

            ' Skip pure comments and blank lines
            If trimmed = "" OrElse trimmed.StartsWith(";"c) Then Continue For

            ' Is this a block keyword line? (e.g. "RADIO; ...")
            Dim isBlock As Boolean = BlockKeywords.Any(
            Function(k) trimmed.StartsWith(k & ";", StringComparison.OrdinalIgnoreCase))

            ' Also treat lines like "GpibEngineDev1=..." as "block" config lines
            If Not isBlock AndAlso
           (trimmed.StartsWith("GpibEngineDev1", StringComparison.OrdinalIgnoreCase) OrElse
            trimmed.StartsWith("GpibEngineDev2", StringComparison.OrdinalIgnoreCase)) Then
                isBlock = True
            End If

            If Not isBlock Then Continue For

            Dim charIndex As Integer = txtEditor.GetFirstCharIndexFromLine(i)

            ' Line number shown in Blocks panel (1-based, with offset tweak)
            Dim displayLine As Integer = i + 1 + BlockLineOffset
            If displayLine < 1 Then displayLine = 1

            Dim display As String = $"{displayLine:000}: {trimmed}"

            lstBlocks.Items.Add(New BlockItem With {
            .Display = display,
            .CharIndex = charIndex
        })
        Next
    End Sub



    Private Sub lstBlocks_DoubleClick(sender As Object, e As EventArgs) Handles lstBlocks.DoubleClick
        Dim item = TryCast(lstBlocks.SelectedItem, BlockItem)
        If item Is Nothing Then Return

        txtEditor.SelectionStart = item.CharIndex
        txtEditor.SelectionLength = 0
        txtEditor.ScrollToCaret()
        txtEditor.Focus()
    End Sub


    Private Sub pnlLineNumbers_Paint(sender As Object, e As PaintEventArgs)

        e.Graphics.Clear(pnlLineNumbers.BackColor)

        If txtEditor.TextLength = 0 Then Return

        Dim totalLines As Integer = txtEditor.Lines.Length
        If totalLines <= 0 Then Return

        ' First visible char & line
        Dim firstCharIndex As Integer = txtEditor.GetCharIndexFromPosition(New Point(0, 0))
        Dim firstLine As Integer = txtEditor.GetLineFromCharIndex(firstCharIndex)
        If firstLine < 0 Then firstLine = 0

        ' Last visible char & line – use Height-1 so we stay in client area
        Dim lastCharIndex As Integer = txtEditor.GetCharIndexFromPosition(New Point(0, txtEditor.ClientSize.Height - 1))
        Dim lastLine As Integer = txtEditor.GetLineFromCharIndex(lastCharIndex)

        ' Clamp to existing lines
        If lastLine >= totalLines Then lastLine = totalLines - 1
        If lastLine < firstLine Then Return

        ' Measure actual line height
        Dim firstPos As Point = txtEditor.GetPositionFromCharIndex(firstCharIndex)
        Dim lineHeight As Integer

        Dim nextLineFirstChar As Integer = txtEditor.GetFirstCharIndexFromLine(firstLine + 1)
        If nextLineFirstChar >= 0 Then
            Dim nextPos As Point = txtEditor.GetPositionFromCharIndex(nextLineFirstChar)
            lineHeight = nextPos.Y - firstPos.Y
            If lineHeight <= 0 Then lineHeight = txtEditor.Font.Height
        Else
            lineHeight = txtEditor.Font.Height
        End If

        ' Start y aligned to first visible line
        Dim y As Integer = -firstPos.Y

        ' Center text vertically in the line a bit
        Dim textSize = TextRenderer.MeasureText("8", txtEditor.Font)
        y += (lineHeight - textSize.Height) \ 2
        y = y + 5

        Dim brush As Brush = Brushes.Gainsboro

        For i As Integer = firstLine To lastLine
            Dim lineNumberText = (i + 1).ToString()

            Dim textWidth = TextRenderer.MeasureText(lineNumberText, txtEditor.Font).Width
            Dim x = pnlLineNumbers.Width - 6 - textWidth

            e.Graphics.DrawString(lineNumberText, txtEditor.Font, brush, x, y)
            y += lineHeight
        Next
    End Sub


    Private Sub txtEditor_VScroll(sender As Object, e As EventArgs) Handles txtEditor.VScroll
        pnlLineNumbers.Invalidate()
    End Sub

    Private Sub txtEditor_TextChanged(sender As Object, e As EventArgs) Handles txtEditor.TextChanged
        pnlLineNumbers.Invalidate()
    End Sub

    Private Sub txtEditor_Resize(sender As Object, e As EventArgs) Handles txtEditor.Resize
        pnlLineNumbers.Invalidate()
    End Sub

    Private Sub FormUserConfigEditor_FormClosing(sender As Object, e As FormClosingEventArgs) _
    Handles Me.FormClosing

        My.Settings.UserGui_Editor = $"{Me.Width},{Me.Height}"
        My.Settings.Save()
    End Sub

    Private Sub CtxUndo(sender As Object, e As EventArgs)
        If txtEditor.CanUndo Then txtEditor.Undo()
    End Sub

    Private Sub CtxCut(sender As Object, e As EventArgs)
        txtEditor.Cut()
    End Sub

    Private Sub CtxCopy(sender As Object, e As EventArgs)
        txtEditor.Copy()
    End Sub

    Private Sub CtxPaste(sender As Object, e As EventArgs)
        txtEditor.Paste()
    End Sub

    Private Sub CtxDelete(sender As Object, e As EventArgs)
        txtEditor.SelectedText = ""
    End Sub

    Private Sub CtxSelectAll(sender As Object, e As EventArgs)
        txtEditor.SelectAll()
    End Sub

    Private Sub CtxDuplicateLine(sender As Object, e As EventArgs)
        Dim caret As Integer = txtEditor.SelectionStart
        Dim lineIndex As Integer = txtEditor.GetLineFromCharIndex(caret)
        If lineIndex < 0 OrElse lineIndex >= txtEditor.Lines.Length Then Return

        Dim lineStart As Integer = txtEditor.GetFirstCharIndexFromLine(lineIndex)
        Dim lineEnd As Integer

        If lineIndex < txtEditor.Lines.Length - 1 Then
            lineEnd = txtEditor.GetFirstCharIndexFromLine(lineIndex + 1)
        Else
            lineEnd = txtEditor.TextLength
        End If

        Dim lineText As String = txtEditor.Text.Substring(lineStart, lineEnd - lineStart)
        If Not lineText.EndsWith(Environment.NewLine) Then
            lineText &= Environment.NewLine
        End If

        txtEditor.SelectionStart = lineEnd
        txtEditor.SelectionLength = 0
        txtEditor.SelectedText = lineText
    End Sub


    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean

        ' --- Ctrl+S = Save ---
        If keyData = (Keys.Control Or Keys.S) Then
            SaveConfigFile()
            Return True
        End If

        ' --- Ctrl+F = Show Find Bar ---
        If keyData = (Keys.Control Or Keys.F) Then
            ShowFindBar()
            Return True
        End If

        ' --- F3 / Shift+F3 = Find next / prev ---
        If keyData = Keys.F3 Then
            DoFind(True)
            Return True
        End If

        If keyData = (Keys.Shift Or Keys.F3) Then
            DoFind(False)
            Return True
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function


    Private Sub ShowFindBar()
        If findPanel Is Nothing Then Return

        Dim wasHidden = Not findPanel.Visible
        findPanel.Visible = True
        findPanel.BringToFront()

        If wasHidden Then
            Dim selText = txtEditor.SelectedText
            If Not String.IsNullOrWhiteSpace(selText) AndAlso
           Not selText.Contains(vbCr) AndAlso
           selText.Length < 80 Then
                txtFind.Text = selText
            End If
        End If

        txtFind.Focus()
        txtFind.SelectAll()
    End Sub


    Private Sub BtnFindClose_Click(sender As Object, e As EventArgs)
        findPanel.Visible = False
        txtEditor.Focus()
    End Sub

    Private Sub BtnFindNext_Click(sender As Object, e As EventArgs)
        DoFind(True)
    End Sub

    Private Sub BtnFindPrev_Click(sender As Object, e As EventArgs)
        DoFind(False)
    End Sub

    Private Sub DoFind(forward As Boolean)
        If txtEditor.TextLength = 0 Then Return

        Dim term As String = txtFind.Text
        If String.IsNullOrEmpty(term) Then Return

        Dim content As String = txtEditor.Text
        Dim comparison = StringComparison.OrdinalIgnoreCase

        Dim start As Integer

        If Not String.Equals(term, _lastFindText, StringComparison.OrdinalIgnoreCase) Then
            ' new search term → start from caret
            start = txtEditor.SelectionStart
        Else
            If forward Then
                start = txtEditor.SelectionStart + txtEditor.SelectionLength
            Else
                start = txtEditor.SelectionStart - 1
                If start < 0 Then start = 0
            End If
        End If

        Dim index As Integer

        If forward Then
            index = content.IndexOf(term, start, comparison)
            If index = -1 AndAlso start > 0 Then
                ' wrap
                index = content.IndexOf(term, 0, comparison)
            End If
        Else
            If start <= 0 Then start = txtEditor.TextLength - 1
            index = content.LastIndexOf(term, start, comparison)
            If index = -1 AndAlso start < txtEditor.TextLength - 1 Then
                index = content.LastIndexOf(term, txtEditor.TextLength - 1, comparison)
            End If
        End If

        If index >= 0 Then
            txtEditor.SelectionStart = index
            txtEditor.SelectionLength = term.Length
            txtEditor.ScrollToCaret()

            ' Make sure it isn't hiding under the Find bar
            EnsureSelectionVisibleWithMargin()

            txtEditor.Focus()

            _lastFindText = term
            _lastFindIndex = index
        Else
            System.Media.SystemSounds.Beep.Play()
        End If

    End Sub


    Private Sub EnsureSelectionVisibleWithMargin()
        If txtEditor Is Nothing OrElse txtEditor.TextLength = 0 Then Return

        ' Which line is the caret on?
        Dim selIndex As Integer = txtEditor.SelectionStart
        Dim selLine As Integer = txtEditor.GetLineFromCharIndex(selIndex)

        ' Current first visible line
        Dim firstVisible As Integer = SendMessage(txtEditor.Handle,
                                              EM_GETFIRSTVISIBLELINE,
                                              IntPtr.Zero,
                                              IntPtr.Zero).ToInt32()

        ' We want the selected line to be a few lines below the top
        Dim desiredFirstVisible As Integer = Math.Max(0, selLine - 3)   ' 3-line margin

        Dim delta As Integer = desiredFirstVisible - firstVisible
        If delta <> 0 Then
            ' Scroll by delta lines (positive = scroll down, negative = scroll up)
            SendMessage(txtEditor.Handle, EM_LINESCROLL, IntPtr.Zero, New IntPtr(delta))
        End If
    End Sub


End Class
