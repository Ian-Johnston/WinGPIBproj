Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Reflection


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

    Private _configPath As String         ' was ReadOnly
    Private WithEvents lstBlocks As New ListBox()
    Private WithEvents txtEditor As New RichTextBox()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnClose As New Button()
    Private WithEvents pnlLineNumbers As New Panel()

    Private lblLastSave As New Label()    ' NEW
    Private txtFilePath As New TextBox()  ' NEW

    Private suppressBlockSync As Boolean = False

    ' Optional: let caller know when a save happened (for auto-reload)
    Public Event ConfigSaved(path As String)


    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer,
                                    wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll",
           CharSet:=CharSet.Auto,
           EntryPoint:="SendMessageW")>
    Private Shared Function SendMessageTabs(hWnd As IntPtr,
                                        msg As Integer,
                                        wParam As Integer,
                                        <[In]()> lParam As Integer()) As IntPtr
    End Function

    Private Const EM_SETTABSTOPS As Integer = &HCB

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
        Dim miSave = New ToolStripMenuItem("Save", Nothing, AddressOf btnSave_Click)   ' NEW

        ctxEditor.Items.AddRange(New ToolStripItem() {
        miUndo,
        New ToolStripSeparator(),
        miCut, miCopy, miPaste, miDelete,
        New ToolStripSeparator(),
        miSelectAll,
        New ToolStripSeparator(),
        miDupLine,
        New ToolStripSeparator(),
        miSave
    })

        txtEditor.ContextMenuStrip = ctxEditor

        ' ---- Bottom panel with buttons + status ----
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

        ' --- Last save label ---
        lblLastSave.AutoSize = False
        lblLastSave.Left = btnClose.Right + 20
        lblLastSave.Top = 11
        lblLastSave.Width = 220
        lblLastSave.Height = 18
        lblLastSave.TextAlign = ContentAlignment.MiddleLeft
        lblLastSave.Text = "Last save: (not yet)"

        ' --- File path textbox (editable) ---
        txtFilePath.Top = 8
        txtFilePath.Left = lblLastSave.Right + 10
        txtFilePath.Height = 22
        txtFilePath.Width = panelButtons.Width - txtFilePath.Left - 10
        txtFilePath.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
        txtFilePath.Text = _configPath

        panelButtons.Controls.Add(btnSave)
        panelButtons.Controls.Add(btnClose)
        panelButtons.Controls.Add(lblLastSave)
        panelButtons.Controls.Add(txtFilePath)

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
        EnableDoubleBuffer(pnlLineNumbers)

        ' ---- Find bar (top) ----
        findPanel = New Panel()
        findPanel.Dock = DockStyle.Top
        findPanel.Height = 28
        findPanel.BackColor = SystemColors.Control    ' << light, normal WinForms bar
        findPanel.Visible = True

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

        editorHost.Controls.Add(txtEditor)       ' FILL
        editorHost.Controls.Add(pnlLineNumbers)  ' LEFT
        editorHost.Controls.Add(panelRight)      ' RIGHT
        editorHost.Controls.Add(findPanel)       ' TOP

        Me.Controls.Add(editorHost)
        Me.Controls.Add(panelButtons)

        ' Match Notepad++ tab stop spacing: 4-space tabs
        If txtEditor.IsHandleCreated Then
            Dim tabStops() As Integer = {16}   ' 4 chars × 4 dialog units
            SendMessageTabs(txtEditor.Handle, EM_SETTABSTOPS, tabStops.Length, tabStops)
            txtEditor.Invalidate()
        Else
            ' If handle not yet created, set tabs once it is
            AddHandler txtEditor.HandleCreated,
        Sub()
            Dim tabStops() As Integer = {16}
            SendMessageTabs(txtEditor.Handle, EM_SETTABSTOPS, tabStops.Length, tabStops)
            txtEditor.Invalidate()
        End Sub
        End If

        ' Load file on start
        LoadConfigFile()
    End Sub


    Private Sub LoadConfigFile()
        Try
            If File.Exists(_configPath) Then
                txtEditor.Text = File.ReadAllText(_configPath, Encoding.UTF8)

                ' show path and last-write time
                txtFilePath.Text = _configPath
                Dim ts = File.GetLastWriteTime(_configPath)
                lblLastSave.Text = "Last save: " & ts.ToString("yyyy-MM-dd HH:mm:ss")
            Else
                txtEditor.Text = "; New User config file" & Environment.NewLine
                txtFilePath.Text = _configPath
                lblLastSave.Text = "Last save: (not yet)"
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
        SyncBlockSelectionToFirstVisibleLine()
    End Sub



    Private Sub SaveConfigFile()
        Try
            Dim newPath As String = txtFilePath.Text.Trim()

            If String.IsNullOrEmpty(newPath) Then
                newPath = _configPath
            End If

            ' If user typed only a file name, keep original directory
            If Not Path.IsPathRooted(newPath) Then
                Dim baseDir As String = Path.GetDirectoryName(_configPath)
                If String.IsNullOrEmpty(baseDir) Then
                    baseDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                End If
                newPath = Path.Combine(baseDir, newPath)
            End If

            ' Ensure folder exists
            Dim dir As String = Path.GetDirectoryName(newPath)
            If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            ' Write file
            File.WriteAllText(newPath, txtEditor.Text, Encoding.UTF8)

            ' Update current path + UI
            _configPath = newPath
            txtFilePath.Text = _configPath
            lblLastSave.Text = "Last save: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Me.Text = "User Config Editor - " & Path.GetFileName(_configPath)

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
        Dim firstVisibleBefore As Integer = SendMessage(
        txtEditor.Handle,
        EM_GETFIRSTVISIBLELINE,
        IntPtr.Zero,
        IntPtr.Zero
    ).ToInt32()

        ' Freeze redraw
        SendMessage(txtEditor.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero)

        ' ---- reset colours & font (single uniform font!) ----
        txtEditor.SelectAll()
        txtEditor.SelectionColor = Color.Gainsboro
        txtEditor.SelectionFont = txtEditor.Font

        Dim text = txtEditor.Text
        Dim length = text.Length

        ' 1) Full-line comments + inline ;;
        Dim idx As Integer = 0
        While idx < length
            Dim lineStart As Integer = idx
            Dim lineEnd As Integer = text.IndexOf(vbLf, idx)
            If lineEnd = -1 Then lineEnd = length

            Dim lineTextFull = text.Substring(lineStart, lineEnd - lineStart)

            ' First non-space char
            Dim firstNonSpace As Integer = lineStart
            While firstNonSpace < lineEnd AndAlso Char.IsWhiteSpace(text(firstNonSpace))
                firstNonSpace += 1
            End While

            ' FULL LINE COMMENT → ALWAYS GREEN
            If firstNonSpace < lineEnd AndAlso text(firstNonSpace) = ";"c Then
                txtEditor.Select(lineStart, lineEnd - lineStart)
                txtEditor.SelectionColor = Color.LimeGreen
                txtEditor.SelectionFont = txtEditor.Font

            Else
                ' INLINE ;; COMMENT
                Dim inlinePos = lineTextFull.IndexOf(";;", StringComparison.Ordinal)
                If inlinePos >= 0 Then
                    Dim start = lineStart + inlinePos
                    txtEditor.Select(start, lineEnd - start)
                    txtEditor.SelectionColor = Color.LimeGreen
                    txtEditor.SelectionFont = txtEditor.Font
                End If
            End If

            idx = lineEnd + 1
        End While

        ' 2) Block keywords (DATASOURCE;, RADIO;, STATSPANEL; etc.) – line-based to avoid K2001Datasource
        For lineIndex As Integer = 0 To txtEditor.Lines.Length - 1
            Dim rawLine = txtEditor.Lines(lineIndex)
            If String.IsNullOrWhiteSpace(rawLine) Then Continue For

            Dim trimmed = rawLine.TrimStart()
            Dim lineStartCharIndex = txtEditor.GetFirstCharIndexFromLine(lineIndex)
            Dim leadingSpaces = rawLine.Length - rawLine.TrimStart().Length

            For Each kw In BlockKeywords
                Dim token = kw & ";"
                If trimmed.StartsWith(token, StringComparison.OrdinalIgnoreCase) Then
                    Dim paintPos = lineStartCharIndex + leadingSpaces
                    txtEditor.Select(paintPos, kw.Length)
                    txtEditor.SelectionColor = Color.Cyan
                    Exit For
                End If
            Next
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

                pos += search.Length
            End While
        Next

        ' 4) Numbers after "=" or "<" or ">"  (e.g. default=1.234, x=500, ymin=-3.2e-6, > 1.123)
        Dim numberPattern As New Regex("(?<=(<=|>=|=|<|>))\s*([-+]?\d+(\.\d+)?([eE][-+]?\d+)?)")
        For Each m As Match In numberPattern.Matches(text)
            Dim valuePart As Group = m.Groups(2)   ' just numeric text
            txtEditor.[Select](valuePart.Index, valuePart.Length)
            txtEditor.SelectionColor = Color.DeepSkyBlue
        Next

        ' 5) Single-line command-like fields: command= / commands= / determine=
        '    (stop at ";" or end-of-line, so ";;" comments remain green)
        Dim cmdPattern As New Regex("(?i)\b(commands?|command|determine)\s*=\s*([^;\r\n]*)")
        For Each m As Match In cmdPattern.Matches(text)
            Dim rhs As Group = m.Groups(2)
            If rhs.Length > 0 Then
                txtEditor.Select(rhs.Index, rhs.Length)
                txtEditor.SelectionColor = Color.Yellow
            End If
        Next

        ' 6) Device keywords: dev1 / dev2
        Dim deviceWords As String() = {"dev1", "dev2"}
        For Each dw In deviceWords
            Dim pos As Integer = 0
            While True
                pos = text.IndexOf(dw, pos, StringComparison.OrdinalIgnoreCase)
                If pos < 0 Then Exit While

                ' Don't match inside larger words (e.g. dev123)
                Dim leftOk As Boolean =
                (pos = 0 OrElse Not Char.IsLetterOrDigit(text(pos - 1)))
                Dim rightOk As Boolean =
                (pos + dw.Length >= text.Length OrElse
                 Not Char.IsLetterOrDigit(text(pos + dw.Length)))

                If leftOk AndAlso rightOk Then
                    txtEditor.[Select](pos, dw.Length)
                    txtEditor.SelectionColor = Color.Red
                End If

                pos += dw.Length
            End While
        Next

        ' 7) Device keywords: dev1 / dev2
        Dim standnativeWords As String() = {"standalone", "native"}

        For Each dw In standnativeWords
            Dim pos As Integer = 0

            While True
                pos = text.IndexOf(dw, pos, StringComparison.OrdinalIgnoreCase)
                If pos < 0 Then Exit While

                ' Don't match inside larger words (e.g. dev123)
                Dim leftOk As Boolean =
                (pos = 0 OrElse Not Char.IsLetterOrDigit(text(pos - 1)))
                Dim rightOk As Boolean =
                (pos + dw.Length >= text.Length OrElse
                 Not Char.IsLetterOrDigit(text(pos + dw.Length)))

                If leftOk AndAlso rightOk Then
                    txtEditor.[Select](pos, dw.Length)
                    txtEditor.SelectionColor = Color.Violet
                    ' keep same font, no bold if you want it subtle
                End If

                pos += dw.Length
            End While
        Next

        ' 8) Multi-line commandlist=  (with inline ;; comments)
        Dim clPattern As New Regex("(?i)\bcommandlist\s*=")
        For Each m As Match In clPattern.Matches(text)
            ' m.Index is at start of "commandlist"
            Dim equalsPos As Integer = text.IndexOf("="c, m.Index)
            If equalsPos < 0 Then Continue For

            ' First line bounds
            Dim lineStart As Integer = text.LastIndexOf(ChrW(10), equalsPos)
            If lineStart = -1 Then lineStart = 0 Else lineStart += 1
            Dim lineEnd As Integer = text.IndexOf(ChrW(10), equalsPos)
            If lineEnd = -1 Then lineEnd = length

            ' Highlight from just after "=" to either ";;" or end-of-line
            Dim rhsStart As Integer = equalsPos + 1
            While rhsStart < lineEnd AndAlso Char.IsWhiteSpace(text(rhsStart))
                rhsStart += 1
            End While

            Dim firstInlineComment As Integer = text.IndexOf(";;", rhsStart, StringComparison.Ordinal)
            Dim rhsEnd As Integer = If(firstInlineComment >= 0 AndAlso firstInlineComment < lineEnd,
                                   firstInlineComment,
                                   lineEnd)

            If rhsEnd > rhsStart Then
                txtEditor.Select(rhsStart, rhsEnd - rhsStart)
                txtEditor.SelectionColor = Color.Yellow
            End If

            ' Now scan continuation lines until block ends
            Dim scanPos As Integer = lineEnd + 1
            While scanPos < length
                Dim contStart As Integer = scanPos
                Dim contEnd As Integer = text.IndexOf(ChrW(10), contStart)
                If contEnd = -1 Then contEnd = length

                Dim lineText = text.Substring(contStart, contEnd - contStart)
                Dim trimmed = lineText.Trim()

                ' End of commandlist block?
                If trimmed = "" Then Exit While
                If trimmed.StartsWith(";"c) Then Exit While

                ' New block keyword?
                If BlockKeywords.Any(Function(kw) trimmed.StartsWith(kw & ";", StringComparison.OrdinalIgnoreCase)) Then
                    Exit While
                End If

                ' New parameter line (something like "x=..." or "target=...")
                Dim eqPos = lineText.IndexOf("="c)
                If eqPos >= 0 Then
                    Dim keyPart = lineText.Substring(0, eqPos).Trim()
                    If ParamKeys.Any(Function(k) keyPart.Equals(k, StringComparison.OrdinalIgnoreCase)) Then
                        Exit While
                    End If
                End If

                ' Treat as continuation of commandlist: highlight from first non-space to ";;" or EOL
                Dim firstNonSpace As Integer = contStart
                While firstNonSpace < contEnd AndAlso Char.IsWhiteSpace(text(firstNonSpace))
                    firstNonSpace += 1
                End While

                If firstNonSpace < contEnd Then
                    Dim inlinePos2 As Integer = text.IndexOf(";;", firstNonSpace, StringComparison.Ordinal)
                    Dim contRhsEnd As Integer = If(inlinePos2 >= 0 AndAlso inlinePos2 < contEnd,
                                               inlinePos2,
                                               contEnd)

                    If contRhsEnd > firstNonSpace Then
                        txtEditor.Select(firstNonSpace, contRhsEnd - firstNonSpace)
                        txtEditor.SelectionColor = Color.Yellow
                    End If
                End If

                scanPos = contEnd + 1
            End While
        Next


        ' Restore caret/selection
        txtEditor.[Select](selStart, selLength)

        ' Restore scroll position (first visible line)
        Dim firstVisibleAfter As Integer = SendMessage(
        txtEditor.Handle,
        EM_GETFIRSTVISIBLELINE,
        IntPtr.Zero,
        IntPtr.Zero
    ).ToInt32()

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

        ' If every block is off by a constant amount, fix it here.
        ' Currently: Blocks panel shows 53 where editor shows 51 → offset = -2
        Const BlockLineOffset As Integer = 0   ' tweak if you ever change header layout

        For i As Integer = 0 To lines.Length - 1
            Dim raw = lines(i)
            Dim trimmed = raw.TrimStart()

            ' Skip pure comments and blank lines
            If trimmed = "" OrElse trimmed.StartsWith(";"c) Then Continue For

            ' Is this a block keyword line? (e.g. "RADIO;" / "DATASOURCE;" / "TAB;")
            Dim isBlock As Boolean = BlockKeywords.Any(
            Function(k) trimmed.StartsWith(k & ";", StringComparison.OrdinalIgnoreCase))

            ' Also treat lines like "GpibEngineDev1=..." / "GpibEngineDev2=..." as block-style entries
            If Not isBlock AndAlso
           (trimmed.StartsWith("GpibEngineDev1", StringComparison.OrdinalIgnoreCase) OrElse
            trimmed.StartsWith("GpibEngineDev2", StringComparison.OrdinalIgnoreCase)) Then
                isBlock = True
            End If

            If Not isBlock Then Continue For

            Dim charIndex As Integer = txtEditor.GetFirstCharIndexFromLine(i)

            ' Display line number (1-based) with global offset
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

        ' Tell the auto-sync to skip the next scroll update
        suppressBlockSync = True

        Dim blockChar As Integer = item.CharIndex
        Dim blockLine As Integer = txtEditor.GetLineFromCharIndex(blockChar)

        ' How many lines of context above the block?
        Const contextLines As Integer = 3

        Dim firstVisibleLine As Integer = Math.Max(0, blockLine - contextLines)
        Dim firstVisibleChar As Integer = txtEditor.GetFirstCharIndexFromLine(firstVisibleLine)

        ' 1) Move caret to the earlier line and scroll so it’s near the top
        txtEditor.SelectionStart = firstVisibleChar
        txtEditor.SelectionLength = 0
        txtEditor.ScrollToCaret()

        ' 2) Now move caret back to the actual block start without extra scrolling
        txtEditor.SelectionStart = blockChar
        txtEditor.SelectionLength = 0

        txtEditor.Focus()
    End Sub


    Private Sub pnlLineNumbers_Paint(sender As Object, e As PaintEventArgs)

        e.Graphics.Clear(pnlLineNumbers.BackColor)

        If txtEditor.TextLength = 0 Then Return

        Dim totalLines As Integer = txtEditor.Lines.Length
        If totalLines <= 0 Then Return

        ' First visible line (by hit-testing at top of editor)
        Dim firstHitChar As Integer = txtEditor.GetCharIndexFromPosition(New Point(0, 0))
        Dim firstLine As Integer = txtEditor.GetLineFromCharIndex(firstHitChar)
        If firstLine < 0 Then firstLine = 0

        ' Last visible line (by hit-testing at bottom of editor)
        Dim lastHitChar As Integer = txtEditor.GetCharIndexFromPosition(New Point(0, txtEditor.ClientSize.Height - 1))
        Dim lastLine As Integer = txtEditor.GetLineFromCharIndex(lastHitChar)
        If lastLine >= totalLines Then lastLine = totalLines - 1
        If lastLine < firstLine Then Return

        Dim lineFont As Font = txtEditor.Font
        Dim brush As Brush = Brushes.Gainsboro

        ' Single tweak for how far down the number sits on each line
        Const verticalNudge As Integer = 2   ' adjust 0..5 to taste

        For i As Integer = firstLine To lastLine

            Dim lineFirstChar As Integer = txtEditor.GetFirstCharIndexFromLine(i)
            If lineFirstChar < 0 Then Continue For

            ' This Y is in the same coordinate system as the panel (they share a parent)
            Dim linePos As Point = txtEditor.GetPositionFromCharIndex(lineFirstChar)
            Dim y As Integer = linePos.Y + verticalNudge

            Dim lineNumberText As String = (i + 1).ToString()
            Dim textWidth As Integer = TextRenderer.MeasureText(lineNumberText, lineFont).Width
            Dim x As Integer = pnlLineNumbers.Width - 6 - textWidth   ' right-justify

            e.Graphics.DrawString(lineNumberText, lineFont, brush, x, y)
        Next
    End Sub

    Private Sub EnableDoubleBuffer(panel As Panel)
        Dim t = panel.GetType()
        Dim pi = t.GetProperty("DoubleBuffered",
                           BindingFlags.Instance Or BindingFlags.NonPublic)
        If pi IsNot Nothing Then
            pi.SetValue(panel, True, Nothing)
        End If
    End Sub


    Private Sub txtEditor_VScroll(sender As Object, e As EventArgs) Handles txtEditor.VScroll
        pnlLineNumbers.Invalidate()
        SyncBlockSelectionToFirstVisibleLine()
    End Sub

    Private Sub txtEditor_TextChanged(sender As Object, e As EventArgs) Handles txtEditor.TextChanged
        pnlLineNumbers.Invalidate()
        SyncBlockSelectionToFirstVisibleLine()
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


    Private Sub SyncBlockSelectionToFirstVisibleLine()

        ' If we just jumped to a block via double-click, skip one auto-sync
        If suppressBlockSync Then
            suppressBlockSync = False
            Return
        End If

        If lstBlocks.Items.Count = 0 OrElse txtEditor.TextLength = 0 Then Return

        ' First visible character & line in the editor
        Dim firstVisibleChar As Integer =
            txtEditor.GetCharIndexFromPosition(New Point(0, 0))
        Dim firstVisibleLine As Integer =
            txtEditor.GetLineFromCharIndex(firstVisibleChar)

        Dim bestIdx As Integer = -1
        Dim bestLine As Integer = Integer.MaxValue

        ' Find the block whose header line is the first one at or below the top visible line
        For i As Integer = 0 To lstBlocks.Items.Count - 1
            Dim bi = TryCast(lstBlocks.Items(i), BlockItem)
            If bi Is Nothing Then Continue For

            Dim blockLine As Integer = txtEditor.GetLineFromCharIndex(bi.CharIndex)

            ' We only care about blocks that are not above the visible area
            If blockLine >= firstVisibleLine AndAlso blockLine < bestLine Then
                bestLine = blockLine
                bestIdx = i
            End If
        Next

        ' If nothing found below/at the top, fall back to the last block in file
        If bestIdx = -1 Then
            bestIdx = lstBlocks.Items.Count - 1
        End If

        If bestIdx >= 0 AndAlso bestIdx < lstBlocks.Items.Count Then
            If lstBlocks.SelectedIndex <> bestIdx Then
                lstBlocks.SelectedIndex = bestIdx
                ' Keep selection roughly in view in the block list
                lstBlocks.TopIndex = Math.Max(0, bestIdx - 3)
            End If
        End If
    End Sub



End Class
