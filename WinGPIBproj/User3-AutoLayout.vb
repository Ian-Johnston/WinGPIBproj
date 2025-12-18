


Partial Class Formtest

    ' Auto layout
    Private AutoLayoutEnabled As Boolean = False
    Private AutoLayoutResizeHooked As Boolean = False
    ' Rremember original sizes from config for locked-size controls
    Private ReadOnly _lockedOriginalSizes As New Dictionary(Of String, Size)(StringComparer.OrdinalIgnoreCase)
    ' keep last loaded config lines (for export)
    Private _lastConfigLines As New List(Of String)
    ' remember if AUTOLAYOUT was on (so we can suggest removing it in export)
    Private _lastAutoLayoutEnabled As Boolean = False



    ' Export updated config with current x/y from rendered controls
    Private Function ExportConfigWithCurrentXY(lines As List(Of String), host As Control) As String

        Dim sb As New System.Text.StringBuilder()

        For Each line In lines
            Dim t As String = line.Trim()
            ' Preserve blank lines exactly
            If t = "" Then
                sb.AppendLine(line)
                Continue For
            End If

            ' Drop AUTOLAYOUT line from export (user will delete it)
            If t.StartsWith("AUTOLAYOUT", StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            ' Only update lines that look like control definitions: "TYPE,..."
            Dim parts = t.Split(","c)
            If parts.Length < 2 Then
                sb.AppendLine(t)
                Continue For
            End If

            Dim ctrlType As String = parts(0).Trim()

            ' Try to extract the controlName from the line (supports positional OR name=)
            Dim controlName As String = ExtractNameFromConfigParts(parts)
            If String.IsNullOrWhiteSpace(controlName) Then
                sb.AppendLine(t)
                Continue For
            End If

            ' Find the rendered control to read x/y from (handles wrappers like Panel_NAME)
            Dim target As Control = FindLayoutControlForName(host, ctrlType, controlName)
            If target Is Nothing Then
                sb.AppendLine(t)
                Continue For
            End If

            Dim newX As Integer = target.Left
            Dim newY As Integer = target.Top

            Dim updatedLine As String = UpsertXYTokens(t, newX, newY)
            sb.AppendLine(updatedLine)
        Next

        Return sb.ToString()
    End Function


    Private Function ExtractNameFromConfigParts(parts() As String) As String
        ' named form: name=XXX
        For i As Integer = 1 To parts.Length - 1
            Dim p = parts(i).Trim()
            If p.StartsWith("name=", StringComparison.OrdinalIgnoreCase) Then
                Return p.Substring(5).Trim()
            End If
        Next

        ' positional form: parts(1) is name
        If parts.Length >= 2 AndAlso Not parts(1).Contains("="c) Then
            Return parts(1).Trim()
        End If

        Return ""
    End Function


    Private Function FindLayoutControlForName(host As Control, ctrlType As String, controlName As String) As Control

        ' BIGTEXT uses Panel_<name> wrapper
        If ctrlType.Equals("BIGTEXT", StringComparison.OrdinalIgnoreCase) Then
            Dim p = host.Controls.Find("Panel_" & controlName, False)
            If p IsNot Nothing AndAlso p.Length > 0 Then Return p(0)
        End If

        ' CHART might be Chart itself OR wrapped; prefer wrapper if present
        If ctrlType.Equals("CHART", StringComparison.OrdinalIgnoreCase) Then
            Dim p = host.Controls.Find("Panel_" & controlName, False)
            If p IsNot Nothing AndAlso p.Length > 0 Then Return p(0)

            Dim c = host.Controls.Find(controlName, False)
            If c IsNot Nothing AndAlso c.Length > 0 Then Return c(0)
        End If

        ' default: control itself by name
        Dim found = host.Controls.Find(controlName, False)
        If found IsNot Nothing AndAlso found.Length > 0 Then Return found(0)

        ' fallback: wrapper panel if you used that pattern elsewhere
        Dim foundPanel = host.Controls.Find("Panel_" & controlName, False)
        If foundPanel IsNot Nothing AndAlso foundPanel.Length > 0 Then Return foundPanel(0)

        Return Nothing
    End Function


    Private Function UpsertXYTokens(line As String, x As Integer, y As Integer) As String
        Dim hasX As Boolean = System.Text.RegularExpressions.Regex.IsMatch(line, "(?i)(^|,)\s*x\s*=")
        Dim hasY As Boolean = System.Text.RegularExpressions.Regex.IsMatch(line, "(?i)(^|,)\s*y\s*=")

        Dim s As String = line

        If hasX Then
            s = System.Text.RegularExpressions.Regex.Replace(s, "(?i)(^|,)\s*x\s*=\s*[^,]*", "$1x=" & x.ToString())
        Else
            s &= ",x=" & x.ToString()
        End If

        If hasY Then
            s = System.Text.RegularExpressions.Regex.Replace(s, "(?i)(^|,)\s*y\s*=\s*[^,]*", "$1y=" & y.ToString())
        Else
            s &= ",y=" & y.ToString()
        End If

        Return s
    End Function


    Private Sub ShowExportPopup(text As String)

        Dim f As New Form()
        f.Text = "AutoLayout Export (copy/paste into your config)"
        f.StartPosition = FormStartPosition.CenterParent
        f.Width = 900
        f.Height = 700

        Dim tb As New TextBox()
        tb.Multiline = True
        tb.ScrollBars = ScrollBars.Both
        tb.WordWrap = False
        tb.Dock = DockStyle.Fill
        tb.Font = New Font("Consolas", 10.0F, FontStyle.Regular)
        tb.Text = text

        Dim btnCopy As New Button()
        btnCopy.Text = "Copy to Clipboard"
        btnCopy.Dock = DockStyle.Top
        AddHandler btnCopy.Click, Sub()
                                      Clipboard.SetText(tb.Text)
                                  End Sub

        f.Controls.Add(tb)
        f.Controls.Add(btnCopy)

        f.ShowDialog(Me)
    End Sub


    Private Sub GroupBoxCustom_ResizeAutoLayout(sender As Object, e As EventArgs)
        If AutoLayoutEnabled Then
            ApplyAutoLayout(GroupBoxCustom)
        End If
    End Sub


    ' simple auto layout that packs controls left-to-right, wraps, and supports "full width" controls
    Private Sub ApplyAutoLayout(host As Control)

        Dim skip As New HashSet(Of Control) From {LabelUSERtab1}

        Dim pad As Integer = 10
        Dim gapX As Integer = 10
        Dim gapY As Integer = 8

        Dim maxX As Integer = 1030
        Dim maxY As Integer = 547

        Dim limitW As Integer = Math.Min(host.ClientSize.Width, maxX)
        Dim limitH As Integer = Math.Min(host.ClientSize.Height, maxY)

        ' Column settings
        Dim colW As Integer = 500
        Dim colX As Integer = pad
        Dim forceColumnBreakY As Integer = 210

        Dim x As Integer = colX
        Dim y As Integer = pad
        Dim rowH As Integer = 0

        ' shrink radio-group GroupBoxes to fit their radios
        For Each gb As GroupBox In host.Controls.OfType(Of GroupBox)()
            Dim radios = gb.Controls.OfType(Of RadioButton)().ToList()
            If radios.Count > 0 Then
                SizeToChildren(gb, 12)
            End If
        Next

        Dim allItems = host.Controls.Cast(Of Control)().
        Where(Function(c) Not skip.Contains(c) AndAlso c.Visible).
        OrderBy(Function(c) c.TabIndex).
        ToList()

        Dim chartItems = allItems.Where(Function(c) IsChartContainer(c)).ToList()
        Dim items = allItems.Where(Function(c) Not IsChartContainer(c)).ToList()

        host.SuspendLayout()

        ' ---------- Place chart(s) FIRST, fixed at bottom, and reserve space ----------
        Dim chartReserveH As Integer = 0
        For Each ch As Control In chartItems
            ' restore original size if known
            Dim sz As Size = Nothing
            If _lockedOriginalSizes.TryGetValue(ch.Name, sz) Then
                ch.Size = sz
            End If
            chartReserveH += ch.Height + gapY
        Next

        If chartReserveH > 0 Then chartReserveH -= gapY ' remove last gap

        ' Everything else must stay ABOVE this line
        Dim layoutBottom As Integer = (limitH - pad) - chartReserveH - gapY
        If layoutBottom < pad Then layoutBottom = limitH - pad ' fallback

        ' Now actually position charts at the bottom
        If chartItems.Count > 0 Then
            Dim chartY As Integer = layoutBottom + gapY
            For Each ch As Control In chartItems
                ch.Location = New Point(pad, chartY)  ' X fixed left; keep size
                chartY += ch.Height + gapY
            Next
        End If

        ' ---------- Layout everything else within [pad .. layoutBottom] ----------
        For Each c As Control In items

            Dim lockSize As Boolean = IsLockedSize(c)

            If lockSize Then
                Dim sz As Size = Nothing
                If _lockedOriginalSizes.TryGetValue(c.Name, sz) Then
                    c.Size = sz
                End If
            End If

            Dim colRight As Integer = Math.Min(limitW - pad, colX + colW)

            If colX = pad AndAlso y > forceColumnBreakY AndAlso (colX + colW + 80) < (limitW - pad) Then
                colX += colW
                x = colX
                y = pad
                rowH = 0
                colRight = Math.Min(limitW - pad, colX + colW)
            End If

            Dim w As Integer = If(lockSize, c.Width, If(c.AutoSize, c.PreferredSize.Width, c.Width))
            Dim h As Integer = If(lockSize, c.Height, If(c.AutoSize, c.PreferredSize.Height, c.Height))

            Dim maxItemW As Integer = Math.Max(50, colRight - colX)
            If Not lockSize AndAlso w > maxItemW Then w = maxItemW

            Dim fullWidth As Boolean = False
            If Not lockSize Then
                fullWidth =
                TypeOf c Is DataGridView OrElse
                TypeOf c Is Panel OrElse
                (w >= CInt((colRight - colX) * 0.75))
            End If

            ' vertical limit is now layoutBottom (NOT limitH - pad)
            If y + h > layoutBottom Then
                colX += colW
                If colX + 50 > (limitW - pad) Then Exit For
                x = colX
                y = pad
                rowH = 0
                colRight = Math.Min(limitW - pad, colX + colW)
            End If

            If fullWidth Then
                If x <> colX Then
                    x = colX
                    y += rowH + gapY
                    rowH = 0
                End If

                If y + h > layoutBottom Then
                    colX += colW
                    If colX + 50 > (limitW - pad) Then Exit For
                    x = colX
                    y = pad
                    rowH = 0
                    colRight = Math.Min(limitW - pad, colX + colW)
                End If

                c.Location = New Point(colX, y)

                Dim fillW As Integer = Math.Max(50, colRight - colX)
                c.Width = fillW
                c.Height = h

                y += c.Height + gapY
                x = colX
                rowH = 0
                Continue For
            End If

            If x + w > colRight Then
                x = colX
                y += rowH + gapY
                rowH = 0
            End If

            If y + h > layoutBottom Then
                colX += colW
                If colX + 50 > (limitW - pad) Then Exit For
                x = colX
                y = pad
                rowH = 0
                colRight = Math.Min(limitW - pad, colX + colW)
            End If

            c.Location = New Point(x, y)

            If Not lockSize Then
                c.Width = w
                c.Height = h
            End If

            x += w + gapX
            rowH = Math.Max(rowH, h)

        Next

        host.ResumeLayout()

    End Sub


    ' chart may be a Chart control OR a container (Panel) that contains a Chart
    Private Function IsChartContainer(c As Control) As Boolean
        If TypeOf c Is System.Windows.Forms.DataVisualization.Charting.Chart Then Return True

        Dim p = TryCast(c, Panel)
        If p IsNot Nothing Then
            If p.Controls.OfType(Of System.Windows.Forms.DataVisualization.Charting.Chart)().Any() Then Return True
        End If

        Return False
    End Function


    ' treat certain labels as layout headers (own row; don't resize)
    Private Function IsLayoutHeader(c As Control) As Boolean
        Dim lbl = TryCast(c, Label)
        If lbl Is Nothing Then Return False

        ' Header heuristic: autosize + no border + bigger font
        If lbl.AutoSize AndAlso lbl.BorderStyle = BorderStyle.None AndAlso lbl.Font.Size >= 12.0F Then
            Return True
        End If

        Return False
    End Function


    ' shrink a container to fit its children
    Private Sub SizeToChildren(box As Control, innerPad As Integer)
        If box.Controls.Count = 0 Then Exit Sub
        Dim maxR As Integer = 0
        Dim maxB As Integer = 0
        For Each ch As Control In box.Controls
            maxR = Math.Max(maxR, ch.Right)
            maxB = Math.Max(maxB, ch.Bottom)
        Next
        box.Width = maxR + innerPad
        box.Height = maxB + innerPad
    End Sub


    ' controls whose size must NOT be altered by auto-layout
    Private Function IsLockedSize(c As Control) As Boolean

        If IsChartContainer(c) Then Return True

        If TypeOf c Is System.Windows.Forms.DataVisualization.Charting.Chart Then Return True

        If c.Tag IsNot Nothing AndAlso c.Tag.ToString().Equals("LOCKSIZE", StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        Dim n As String = If(c.Name, "")

        If n.StartsWith("BIGTEXT", StringComparison.OrdinalIgnoreCase) Then Return True
        If n.StartsWith("LED", StringComparison.OrdinalIgnoreCase) Then Return True
        If n.StartsWith("TEXTAREA", StringComparison.OrdinalIgnoreCase) Then Return True

        ' NEW: treat actual textarea controls as locked-size regardless of name
        Dim tb = TryCast(c, TextBox)
        If tb IsNot Nothing AndAlso tb.Multiline Then Return True

        If TypeOf c Is RichTextBox Then Return True

        Return False
    End Function


End Class