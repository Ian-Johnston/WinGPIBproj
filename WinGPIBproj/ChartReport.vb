Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Text

' Playback Chart "Export Report": analyses the loaded CSV and writes a self-contained HTML report (text, tables, key findings and
' charts) next to the CSV. Complements Export Results, which only writes the numbers. The data is copied on the UI thread, then the
' analysis runs on a worker thread (so a large CSV doesn't freeze the form) and never touches a control.
Partial Public Class Chart

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Data handed to the report builder
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Class ReportDevice
        Public Slot As Integer
        Public Name As String = ""
        Public FirstIndex As Integer                ' position of the first reading of the scope within the whole run
        Public X() As Double                        ' readings of the scope
        Public T() As Double                        ' temperature of each reading
        Public H() As Double                        ' humidity of each reading
        Public TimeText() As String                 ' DATETIME text of each reading
        Public BaseValue As Double                  ' PPM baseline: Initial Value
        Public BaseTemp As Double                   ' PPM baseline: Initial Temp
        Public BaseSource As String = ""
    End Class

    Private Class ReportInput
        Public Devices As New List(Of ReportDevice)
        Public CsvFile As String = ""
        Public Metadata As String = ""
        Public ScopeText As String = "whole run"
        Public RegionOn As Boolean = False
        Public WarmUpMinutes As Double = 0
        Public SpecPpm As Double = Double.NaN
        Public MinsPerSample As Double = 0
        Public RmsWindow As Integer = 10
        Public Culture As CultureInfo = CultureInfo.CurrentCulture
    End Class

    Private Class ReportResult
        Public Html As String = ""
        Public Summary As String = ""
    End Class

    Private Class ReportScope
        Public Dev As ReportDevice
        Public Count As Integer
        Public Mean As Double
        Public StdDev As Double
        Public Ti() As Double                       ' time of each reading, minutes since the start of the run (sample numbers if the time scale is unknown)
        Public Times() As DateTime                  ' parsed timestamps (MinValue where unreadable)
        Public TimesOk As Boolean
        Public Colour As ScottPlot.Color
        Public Spikes As List(Of Double()) = Nothing    ' found by RpSectionEvents: {index, deviation, readings}
        Public Steps As List(Of Double()) = Nothing     ' found by RpSectionEvents: {index, shift, strength}
    End Class

    Private Class ReportFinding
        Public Level As String                      ' "warn", "note" or "ok"
        Public Text As String
        Public Sub New(findingLevel As String, findingText As String)
            Level = findingLevel
            Text = findingText
        End Sub
    End Class

    ' Figures shared by the sections of one device
    Private Class ReportCtx
        Public Tag As String = ""
        Public AbsMean As Double
        Public Sd As Double
        Public PointNoise As Double
        Public MinGap As Double = Double.MaxValue
        Public Mps As Double
        Public Function Ppm(v As Double) As Double
            If AbsMean > 0 Then Return v / AbsMean * 1000000.0
            Return Double.NaN
        End Function
        Public Function PpmText(v As Double) As String
            Dim p As Double = Ppm(v)
            If Double.IsNaN(p) OrElse Double.IsInfinity(p) Then Return "n/a"
            Return p.ToString("G4", CultureInfo.InvariantCulture) & " ppm"
        End Function
    End Class

    Private Chart2ReportWarmUp As Decimal = 10D
    Private Chart2ReportSpec As String = ""
    Private Chart2ReportSummary As Boolean = True

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Button
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Async Sub ButtonExportReport_Click(sender As Object, e As EventArgs) Handles ButtonExportReport.Click

        If Not (ChartLoaded AndAlso CSVfileok) Then Exit Sub

        Dim regionOn As Boolean = Chart2RegionSpan IsNot Nothing AndAlso Chart2RegionSpan.IsVisible
        Dim regionFirst As Integer = 0
        Dim regionLast As Integer = 0
        If regionOn Then
            regionFirst = Math.Max(0, CInt(Math.Ceiling(Chart2RegionSpan.Left)))
            regionLast = CInt(Math.Floor(Chart2RegionSpan.Right))
        End If

        Dim useDev1 As Boolean = False
        Dim useDev2 As Boolean = False
        Dim useRegion As Boolean = regionOn
        Dim warmUp As Decimal = Chart2ReportWarmUp
        Dim specText As String = Chart2ReportSpec
        Dim writeSummary As Boolean = Chart2ReportSummary

        If Not ShowReportOptions(useDev1, useDev2, useRegion, warmUp, specText, writeSummary, regionOn,
                                 regionFirst.ToString() & " - " & regionLast.ToString()) Then Exit Sub

        Chart2ReportWarmUp = warmUp
        Chart2ReportSpec = specText
        Chart2ReportSummary = writeSummary

        Dim specPpm As Double = Double.NaN
        If specText.Trim() <> "" Then
            Double.TryParse(specText.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, specPpm)
            If Double.IsNaN(specPpm) Then Double.TryParse(specText.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, specPpm)
        End If

        Dim inp As ReportInput = GatherReportInput(useDev1, useDev2, useRegion AndAlso regionOn, regionFirst, regionLast, CDbl(warmUp), specPpm)

        If inp.Devices.Count = 0 Then
            MessageBox.Show("There is nothing to report: no device has enough readings in the chosen scope.", "Export Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        ' Progress window while the worker thread builds the report
        Dim progress As New Form With {
            .Text = "WinGPIB - Export Report",
            .FormBorderStyle = FormBorderStyle.FixedDialog,
            .StartPosition = FormStartPosition.CenterParent,
            .ControlBox = False,
            .ShowInTaskbar = False,
            .ClientSize = New Size(340, 70)
        }
        progress.Controls.Add(New Label With {.Text = "Analysing the CSV and building the report...", .AutoSize = True, .Location = New Point(14, 12)})
        progress.Controls.Add(New ProgressBar With {.Style = ProgressBarStyle.Marquee, .MarqueeAnimationSpeed = 30, .Location = New Point(14, 38), .Size = New Size(312, 16)})

        Dim result As ReportResult = Nothing
        Dim failure As String = ""

        ButtonExportReport.Enabled = False
        progress.Show(Me)
        Me.Cursor = Cursors.WaitCursor

        Try
            result = Await Threading.Tasks.Task.Run(Function() BuildReport(inp))
        Catch ex As Exception
            failure = ex.Message
        End Try

        Me.Cursor = Cursors.Default
        progress.Close()
        progress.Dispose()
        ButtonExportReport.Enabled = True

        If failure <> "" Then
            MessageBox.Show("The report could not be built: " & failure, "Export Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' Save
        Dim baseName As String = Path.GetFileNameWithoutExtension(filePlayback)
        Dim fileSuffix As String = If(useRegion AndAlso regionOn, "_region_" & regionFirst.ToString() & "-" & regionLast.ToString(), "")
        Dim folder As String = ""
        Try
            folder = Path.GetDirectoryName(filePlayback)
        Catch
        End Try

        Dim savedPath As String = ""

        Using dlg As New SaveFileDialog With {
            .Title = "Export Report",
            .Filter = "HTML report (*.html)|*.html|All files (*.*)|*.*",
            .FileName = baseName & "_Report" & fileSuffix & ".html",
            .OverwritePrompt = True
        }
            If folder <> "" AndAlso Directory.Exists(folder) Then dlg.InitialDirectory = folder
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Exit Sub

            Try
                File.WriteAllText(dlg.FileName, result.Html, New UTF8Encoding(False))
                savedPath = dlg.FileName

                ' Plain-text copy of the key findings, for pasting into an email or a forum post
                If writeSummary Then
                    Dim summaryPath As String = Path.Combine(Path.GetDirectoryName(savedPath), Path.GetFileNameWithoutExtension(savedPath) & "_Summary.txt")
                    File.WriteAllText(summaryPath, result.Summary, New UTF8Encoding(False))
                End If
            Catch ex As Exception
                MessageBox.Show("The report could not be saved: " & ex.Message, "Export Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try
        End Using

        If MessageBox.Show("Report saved:" & vbCrLf & savedPath & If(writeSummary, vbCrLf & "(plain-text summary saved beside it)", "") & vbCrLf & vbCrLf & "Open it now?", "Export Report", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                Process.Start(New ProcessStartInfo(savedPath) With {.UseShellExecute = True})
            Catch ex As Exception
                MessageBox.Show("The report could not be opened: " & ex.Message, "Export Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Try
        End If

    End Sub

    ' Small options window. False if cancelled.
    Private Function ShowReportOptions(ByRef useDev1 As Boolean, ByRef useDev2 As Boolean, ByRef useRegion As Boolean, ByRef warmUp As Decimal,
                                       ByRef specText As String, ByRef writeSummary As Boolean, regionOn As Boolean, regionText As String) As Boolean

        Using dlg As New Form With {
            .Text = "WinGPIB - Export Report",
            .FormBorderStyle = FormBorderStyle.FixedDialog,
            .StartPosition = FormStartPosition.CenterParent,
            .MaximizeBox = False,
            .MinimizeBox = False,
            .ShowInTaskbar = False,
            .ClientSize = New Size(400, 366)
        }

            Dim name1 As String = DeviceName1.Text
            Dim name2 As String = DeviceName2.Text

            dlg.Controls.Add(New Label With {.Text = "Devices to include:", .AutoSize = True, .Location = New Point(14, 12)})
            Dim chk1 As New CheckBox With {.Text = "Dev 1" & If(name1 <> "", " - " & name1, " (not in this file)"), .AutoSize = True,
                                           .Location = New Point(30, 34), .Enabled = name1 <> "", .Checked = name1 <> ""}
            Dim chk2 As New CheckBox With {.Text = "Dev 2" & If(name2 <> "", " - " & name2, " (not in this file)"), .AutoSize = True,
                                           .Location = New Point(30, 58), .Enabled = name2 <> "", .Checked = name2 <> ""}
            dlg.Controls.Add(chk1)
            dlg.Controls.Add(chk2)

            dlg.Controls.Add(New Label With {.Text = "Part of the run:", .AutoSize = True, .Location = New Point(14, 92)})
            Dim rbWhole As New RadioButton With {.Text = "Whole run", .AutoSize = True, .Location = New Point(30, 114), .Checked = Not regionOn}
            Dim rbRegion As New RadioButton With {.Text = "Regional Stats band only (samples " & regionText & ")", .AutoSize = True,
                                                  .Location = New Point(30, 138), .Enabled = regionOn, .Checked = regionOn}
            dlg.Controls.Add(rbWhole)
            dlg.Controls.Add(rbRegion)

            dlg.Controls.Add(New Label With {.Text = "Warm-up to skip when measuring drift (minutes):", .AutoSize = True, .Location = New Point(14, 178)})
            Dim nudWarm As New NumericUpDown With {.Minimum = 0, .Maximum = 100000, .DecimalPlaces = 1, .Increment = 1, .Width = 80,
                                                   .Location = New Point(30, 200), .Value = Math.Max(0D, Math.Min(100000D, warmUp))}
            dlg.Controls.Add(nudWarm)

            dlg.Controls.Add(New Label With {.Text = "Meter specification, ppm of reading (optional):", .AutoSize = True, .Location = New Point(14, 238)})
            Dim txtSpec As New TextBox With {.Width = 80, .Location = New Point(30, 260), .Text = specText}
            dlg.Controls.Add(txtSpec)

            Dim chkSummary As New CheckBox With {.Text = "Also save a plain-text summary of the key findings (for pasting)", .AutoSize = True,
                                                 .Location = New Point(14, 296), .Checked = writeSummary}
            dlg.Controls.Add(chkSummary)

            Dim ok As New Button With {.Text = "Create report", .DialogResult = DialogResult.OK, .Size = New Size(104, 28), .Location = New Point(176, 326)}
            Dim cancel As New Button With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Size = New Size(80, 28), .Location = New Point(306, 326)}
            dlg.Controls.Add(ok)
            dlg.Controls.Add(cancel)
            dlg.AcceptButton = ok
            dlg.CancelButton = cancel

            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return False

            If Not chk1.Checked AndAlso Not chk2.Checked Then
                MessageBox.Show("Tick at least one device.", "Export Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If

            useDev1 = chk1.Checked
            useDev2 = chk2.Checked
            useRegion = rbRegion.Checked
            warmUp = nudWarm.Value
            specText = txtSpec.Text
            writeSummary = chkSummary.Checked
            Return True

        End Using

    End Function

    ' Copies what the report needs out of the loaded data (UI thread).
    Private Function GatherReportInput(useDev1 As Boolean, useDev2 As Boolean, useRegion As Boolean, regionFirst As Integer, regionLast As Integer,
                                       warmUp As Double, specPpm As Double) As ReportInput

        Dim inp As New ReportInput With {
            .CsvFile = filePlayback,
            .Metadata = MetadataChart.Text,
            .RegionOn = useRegion,
            .WarmUpMinutes = warmUp,
            .SpecPpm = specPpm,
            .MinsPerSample = Chart2MinsPerSample,
            .RmsWindow = Math.Max(2, CInt(Val(RMSwindow.Text))),
            .Culture = Globalization.CultureInfo.CurrentCulture
        }

        Dim ppmSlot As Integer = If(RadioButtonDev2.Checked, 2, 1)

        If useRegion Then inp.ScopeText = "Regional Stats band, samples " & regionFirst.ToString() & " - " & regionLast.ToString()

        For slot As Integer = 1 To 2

            If slot = 1 AndAlso Not useDev1 Then Continue For
            If slot = 2 AndAlso Not useDev2 Then Continue For

            Dim name As String = If(slot = 1, DeviceName1.Text, DeviceName2.Text)
            If name = "" Then Continue For

            Dim rows() As DataRow = dataTable1.Select("DEVICE ='" & name.Replace("'", "''") & "'")
            Dim n As Integer = rows.Length
            If n = 0 Then Continue For

            Dim first As Integer = 0
            Dim last As Integer = n - 1
            If useRegion Then
                first = Math.Max(0, regionFirst)
                last = Math.Min(n - 1, regionLast)
            End If

            If last - first < 2 Then Continue For

            ' PPM baseline as the chart uses it: the first reading in the file, or the typed Initial Value / Temp for the PPM device
            Dim baseValue As Double = Convert.ToDouble(rows(0)("VALUE"))
            Dim baseTemp As Double = Convert.ToDouble(rows(0)("TEMP"))
            Dim baseSource As String = "first reading in the file"
            If slot = ppmSlot Then
                If Not CheckBoxMedianV.Checked Then
                    baseValue = ParseInvariantDouble(MedianValue.Text)
                    baseSource = "typed Initial Value"
                End If
                If Not CheckBoxMedianT.Checked Then baseTemp = ParseInvariantDouble(MedianTemp.Text)
            End If

            Dim count As Integer = last - first + 1
            Dim dev As New ReportDevice With {
                .Slot = slot,
                .Name = name,
                .FirstIndex = first,
                .BaseValue = baseValue,
                .BaseTemp = baseTemp,
                .BaseSource = baseSource,
                .X = New Double(count - 1) {},
                .T = New Double(count - 1) {},
                .H = New Double(count - 1) {},
                .TimeText = New String(count - 1) {}
            }

            For i As Integer = 0 To count - 1
                Dim r As DataRow = rows(first + i)
                dev.X(i) = Convert.ToDouble(r("VALUE"))
                dev.T(i) = Convert.ToDouble(r("TEMP"))
                dev.H(i) = Convert.ToDouble(r("HUM"))
                dev.TimeText(i) = Convert.ToString(r("DATETIME"))
            Next

            inp.Devices.Add(dev)

        Next

        Return inp

    End Function

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Small helpers
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Function RpHtml(s As String) As String
        Return WebUtility.HtmlEncode(If(s, ""))
    End Function

    Private Function RpNum(x As Double, Optional digits As Integer = 6) As String
        If Double.IsNaN(x) OrElse Double.IsInfinity(x) Then Return "n/a"
        Return x.ToString("G" & digits.ToString(), CultureInfo.InvariantCulture)
    End Function

    Private Function RpPpmNum(x As Double) As String
        If Double.IsNaN(x) OrElse Double.IsInfinity(x) Then Return "n/a"
        Return x.ToString("G4", CultureInfo.InvariantCulture) & " ppm"
    End Function

    ' "45 s", "12.5 min", "3.2 h"
    Private Function RpDur(minutes As Double) As String
        If Double.IsNaN(minutes) OrElse Double.IsInfinity(minutes) Then Return "n/a"
        If minutes < 1 Then Return (minutes * 60).ToString("0") & " s"
        If minutes < 120 Then Return minutes.ToString("0.0", CultureInfo.InvariantCulture) & " min"
        Return (minutes / 60).ToString("0.0", CultureInfo.InvariantCulture) & " h"
    End Function

    Private Sub RpRow(sb As StringBuilder, label As String, value As String, Optional note As String = "")
        sb.Append("<tr><th>").Append(RpHtml(label)).Append("</th><td>").Append(RpHtml(value)).Append("</td><td class=""note"">").Append(RpHtml(note)).AppendLine("</td></tr>")
    End Sub

    Private Sub RpTableStart(sb As StringBuilder, title As String)
        sb.AppendLine("<h3>" & RpHtml(title) & "</h3>")
        sb.AppendLine("<table class=""data"">")
    End Sub

    Private Sub RpTableEnd(sb As StringBuilder)
        sb.AppendLine("</table>")
    End Sub

    Private Function RpParseTime(text As String, culture As CultureInfo) As DateTime
        Try
            Dim parts() As String = text.Split("_"c)
            If parts.Length < 2 Then Return DateTime.MinValue
            Return DateTime.Parse(parts(0).Trim(), culture).Add(DateTime.Parse(parts(1).Trim(), culture).TimeOfDay)
        Catch
            Return DateTime.MinValue
        End Try
    End Function

    Private Function RpMedian(a() As Double) As Double
        If a.Length = 0 Then Return Double.NaN
        Dim s() As Double = DirectCast(a.Clone(), Double())
        Array.Sort(s)
        Dim m As Integer = s.Length \ 2
        If s.Length Mod 2 = 1 Then Return s(m)
        Return (s(m - 1) + s(m)) / 2.0
    End Function

    ' Median absolute deviation
    Private Function RpMad(a() As Double) As Double
        If a.Length = 0 Then Return Double.NaN
        Dim med As Double = RpMedian(a)
        Dim d(a.Length - 1) As Double
        For i As Integer = 0 To a.Length - 1
            d(i) = Math.Abs(a(i) - med)
        Next
        Return RpMedian(d)
    End Function

    ' STDEV of the successive differences / sqrt(2): the noise of a single reading, blind to slow drift.
    Private Function RpPointNoise(a() As Double) As Double
        If a.Length < 3 Then Return Double.NaN
        Dim d(a.Length - 2) As Double
        For i As Integer = 0 To a.Length - 2
            d(i) = a(i + 1) - a(i)
        Next
        Return ExStdev(d) / Math.Sqrt(2.0)
    End Function

    Private Function RpCorrelation(a() As Double, b() As Double) As Double
        Dim n As Integer = Math.Min(a.Length, b.Length)
        If n < 3 Then Return Double.NaN
        Dim ma As Double = 0.0
        Dim mb As Double = 0.0
        For i As Integer = 0 To n - 1
            ma += a(i)
            mb += b(i)
        Next
        ma /= n
        mb /= n
        Dim sab As Double = 0.0
        Dim saa As Double = 0.0
        Dim sbb As Double = 0.0
        For i As Integer = 0 To n - 1
            sab += (a(i) - ma) * (b(i) - mb)
            saa += (a(i) - ma) * (a(i) - ma)
            sbb += (b(i) - mb) * (b(i) - mb)
        Next
        If saa <= 0 OrElse sbb <= 0 Then Return Double.NaN
        Return sab / Math.Sqrt(saa * sbb)
    End Function

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  The report
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Function BuildReport(inp As ReportInput) As ReportResult

        Dim findings As New List(Of ReportFinding)
        Dim body As New StringBuilder()
        Dim scopes As New List(Of ReportScope)
        Dim colours() As ScottPlot.Color = {New ScottPlot.Color(Color.FromArgb(0, 90, 190)), New ScottPlot.Color(Color.FromArgb(215, 95, 0))}

        For Each dev As ReportDevice In inp.Devices

            Dim sc As ReportScope = RpBuildScope(inp, dev)
            If sc Is Nothing Then Continue For
            sc.Colour = colours(If(dev.Slot = 1, 0, 1))
            scopes.Add(sc)

            Try
                RpDeviceSection(inp, sc, body, findings)
            Catch ex As Exception
                body.AppendLine("<p class=""warn"">Part of the analysis of Dev " & dev.Slot.ToString() & " could not be completed: " & RpHtml(ex.Message) & "</p>")
                body.AppendLine("</section>")
            End Try

        Next

        If scopes.Count = 2 Then
            Try
                RpCompareSection(inp, scopes(0), scopes(1), body, findings)
            Catch ex As Exception
                body.AppendLine("<section><h2>Dev 1 compared with Dev 2</h2><p class=""warn"">This part could not be completed: " & RpHtml(ex.Message) & "</p></section>")
            End Try
        End If

        ' A title that tells reports apart: the devices, then any note from the CSV's metadata lines (lines that merely repeat a device name are skipped)
        Dim deviceNames As New List(Of String)
        For Each sc0 As ReportScope In scopes
            deviceNames.Add(sc0.Dev.Name)
        Next

        Dim noteParts As New List(Of String)
        For Each metaLine As String In inp.Metadata.Replace(vbCr, "").Split(vbLf(0))
            Dim line As String = metaLine.Trim()
            If line = "" Then Continue For
            Dim repeatsDevice As Boolean = False
            For Each nm As String In deviceNames
                If nm.IndexOf(line, StringComparison.OrdinalIgnoreCase) >= 0 OrElse line.IndexOf(nm, StringComparison.OrdinalIgnoreCase) >= 0 Then repeatsDevice = True
            Next
            If Not repeatsDevice Then noteParts.Add(line)
        Next

        Dim reportTitle As String = If(deviceNames.Count > 0, String.Join(" and ", deviceNames), Path.GetFileNameWithoutExtension(inp.CsvFile))
        If noteParts.Count > 0 Then reportTitle &= ": " & String.Join(" - ", noteParts)

        Dim html As New StringBuilder()

        html.AppendLine("<!DOCTYPE html>")
        html.AppendLine("<html lang=""en""><head><meta charset=""utf-8"">")
        html.AppendLine("<title>" & RpHtml(reportTitle) & " - " & DateTime.Now.ToString("yyyy-MM-dd") & " - WinGPIB report</title>")
        html.AppendLine("<style>")
        html.AppendLine("body{font-family:'Segoe UI',Arial,sans-serif;margin:24px auto;max-width:1060px;padding:0 16px;color:#1d1d1d;line-height:1.45}")
        html.AppendLine("h1{font-size:26px;margin:0 0 4px}h2{font-size:20px;margin:30px 0 8px;border-bottom:2px solid #c9d3e0;padding-bottom:4px}h3{font-size:15px;margin:20px 0 6px;color:#27406a}")
        html.AppendLine(".sub{color:#555;margin:0 0 14px}.eyebrow{color:#27406a;font-size:13px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;margin-top:6px}section{margin-bottom:10px}")
        html.AppendLine("table.data{border-collapse:collapse;width:100%;font-size:13.5px}table.data th{text-align:left;font-weight:600;width:30%;padding:4px 8px;border-bottom:1px solid #e3e3e3;vertical-align:top}")
        html.AppendLine("table.data td{padding:4px 8px;border-bottom:1px solid #e3e3e3;vertical-align:top}table.data td.note{color:#666;font-size:12.5px}table.data thead th{background:#eef2f8;width:auto}")
        html.AppendLine("table.grid{border-collapse:collapse;font-size:13px}table.grid th,table.grid td{padding:3px 10px;border:1px solid #d5d9e0;text-align:right}table.grid th{background:#eef2f8}")
        html.AppendLine(".findings{background:#f6f8fb;border:1px solid #d5dbe6;border-radius:6px;padding:6px 18px 10px}.findings ul{list-style:none;padding:0;margin:6px 0}")
        html.AppendLine(".findings li{padding:3px 0 3px 26px;position:relative}.findings li:before{position:absolute;left:0;top:3px;font-weight:700}")
        html.AppendLine(".findings li.ok:before{content:'\2714';color:#2a8a3c}.findings li.note:before{content:'\2022';color:#27406a;font-size:20px;top:-1px}.findings li.warn:before{content:'\26A0';color:#c25a00}")
        html.AppendLine(".chart{margin:8px 0 14px;overflow-x:auto}.chart svg{width:100%;max-width:100%;height:auto}.warn{color:#b04a00}pre{background:#f4f4f4;padding:8px 12px;border-radius:4px;white-space:pre-wrap}")
        html.AppendLine("html{scroll-behavior:smooth;scroll-padding-top:60px}")
        html.AppendLine("#topnav{position:sticky;top:0;z-index:10;background:#fff;border-bottom:1px solid #c9d3e0;padding:9px 0;margin:8px 0 14px;display:flex;flex-wrap:wrap;gap:6px 18px;font-size:14px}")
        html.AppendLine("#topnav a,.toc a{color:#1f4f9c;text-decoration:none}#topnav a:hover,.toc a:hover{text-decoration:underline}")
        html.AppendLine(".toc{background:#f6f8fb;border:1px solid #d5dbe6;border-radius:6px;padding:4px 18px 8px}.toc ul{margin:2px 0 6px;padding-left:20px}.toc>ul{list-style:none;padding-left:0}.toc>ul>li{margin-top:6px;font-weight:600}.toc ul ul li{font-weight:400;font-size:14px}")
        html.AppendLine(".foot{color:#555;font-size:13px}@media print{body{max-width:none}section{page-break-inside:auto}#topnav{display:none}}")
        html.AppendLine("</style></head><body>")

        html.AppendLine("<div class=""eyebrow"">WinGPIB Playback Report</div>")
        html.AppendLine("<h1>" & RpHtml(reportTitle) & "</h1>")
        html.AppendLine("<p class=""sub"">" & RpHtml(Path.GetFileName(inp.CsvFile)) & " &nbsp;|&nbsp; " & RpHtml(inp.ScopeText) & " &nbsp;|&nbsp; created " &
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm") & "</p>")

        If inp.Metadata.Trim() <> "" Then
            Dim metaLines As New List(Of String)
            For Each metaLine As String In inp.Metadata.Replace(vbCr, "").Split(vbLf(0))
                If metaLine.Trim() <> "" Then metaLines.Add(metaLine.Trim())
            Next
            html.AppendLine("<pre>" & RpHtml(String.Join(vbLf, metaLines)) & "</pre>")
        End If

        Dim content As New StringBuilder()

        ' Key findings, most important first
        content.AppendLine("<section class=""findings""><h2 style=""border:none;margin-top:8px"">Key findings</h2><ul>")
        For Each level As String In {"warn", "note", "ok"}
            For Each f As ReportFinding In findings
                If f.Level = level Then content.AppendLine("<li class=""" & level & """>" & RpHtml(f.Text) & "</li>")
            Next
        Next
        If findings.Count = 0 Then content.AppendLine("<li class=""note"">No findings could be generated.</li>")
        content.AppendLine("</ul></section>")

        ' Run summary
        content.AppendLine("<section><h2>Run summary</h2>")
        RpTableStart(content, "Details")
        RpRow(content, "CSV file", inp.CsvFile)
        RpRow(content, "Scope", inp.ScopeText)
        For Each sc As ReportScope In scopes
            Dim label As String = "Dev " & sc.Dev.Slot.ToString() & " - " & sc.Dev.Name
            Dim spanText As String = ""
            If inp.MinsPerSample > 0 Then spanText = RpDur((sc.Count - 1) * inp.MinsPerSample)
            RpRow(content, label, sc.Count.ToString() & " readings" & If(spanText <> "", ", " & spanText, ""),
                  If(sc.TimesOk, sc.Dev.TimeText(0).Replace("_", " ") & "  to  " & sc.Dev.TimeText(sc.Count - 1).Replace("_", " "), ""))
        Next
        RpTableEnd(content)
        content.AppendLine("</section>")

        content.Append(body.ToString())

        content.AppendLine("<section class=""foot""><h2>How to read this report</h2><ul>")
        content.AppendLine("<li>All figures use the raw readings in the CSV (the Avg boxes are ignored). ppm values are relative to the mean reading of the scope.</li>")
        content.AppendLine("<li>The key findings are generated automatically from simple tests (thresholds, windows and fits described beside each result). They point to things worth a look; check the charts before drawing conclusions.</li>")
        content.AppendLine("<li>Errors quoted for drift and tempco are the formal least-squares errors. Real instrument data is correlated from reading to reading, so the true uncertainty is larger.</li>")
        content.AppendLine("<li>Allan deviation is calculated from overlapping windows (as in the Allan Deviation pop-up). The tempco is only meaningful when the temperature changed enough, independently of the drift.</li>")
        content.AppendLine("</ul></section>")
        ' Give every heading an anchor and build the menu (a bar that stays at the top, and a contents list)
        Dim entries As New List(Of String())
        Dim anchored As String = RpAddAnchors(content.ToString(), entries)

        html.AppendLine("<nav id=""topnav"">")
        For Each entry As String() In entries
            If entry(0) = "2" Then html.AppendLine("<a href=""#" & entry(1) & """>" & entry(2) & "</a>")
        Next
        html.AppendLine("</nav>")

        html.AppendLine("<section class=""toc""><h2 style=""border:none;margin-top:4px"">Contents</h2><ul>")
        Dim subOpen As Boolean = False
        For Each entry As String() In entries
            If entry(0) = "2" Then
                If subOpen Then html.AppendLine("</ul></li>")
                html.AppendLine("<li><a href=""#" & entry(1) & """>" & entry(2) & "</a>")
                html.AppendLine("<ul>")
                subOpen = True
            Else
                html.AppendLine("<li><a href=""#" & entry(1) & """>" & entry(2) & "</a></li>")
            End If
        Next
        If subOpen Then html.AppendLine("</ul></li>")
        html.AppendLine("</ul></section>")

        html.Append(anchored)
        html.AppendLine("</body></html>")

        ' Plain-text summary
        Dim summary As New StringBuilder()
        summary.AppendLine("WinGPIB Playback Report: " & reportTitle)
        summary.AppendLine("File: " & Path.GetFileName(inp.CsvFile) & "     Scope: " & inp.ScopeText & "     Created: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm"))
        If inp.Metadata.Trim() <> "" Then
            For Each metaLine As String In inp.Metadata.Replace(vbCr, "").Split(vbLf(0))
                If metaLine.Trim() <> "" Then summary.AppendLine(metaLine.Trim())
            Next
        End If
        summary.AppendLine()
        summary.AppendLine("KEY FINDINGS   ([!] check this, [-] for information, [ok] fine)")
        For Each level As String In {"warn", "note", "ok"}
            For Each f As ReportFinding In findings
                If f.Level = level Then summary.AppendLine(If(level = "warn", "[!]  ", If(level = "note", "[-]  ", "[ok] ")) & f.Text)
            Next
        Next

        Return New ReportResult With {.Html = html.ToString(), .Summary = summary.ToString()}

    End Function

    ' Adds an id to every <h2> / <h3> and records {level, id, text} of each, for the menu.
    Private Function RpAddAnchors(content As String, entries As List(Of String())) As String

        Dim counter As Integer = 0

        Return System.Text.RegularExpressions.Regex.Replace(content, "<h([23])([^>]*)>(.*?)</h\1>",
            Function(m As System.Text.RegularExpressions.Match) As String
                counter += 1
                Dim id As String = "sec" & counter.ToString()
                Dim plain As String = System.Text.RegularExpressions.Regex.Replace(m.Groups(3).Value, "<[^>]+>", "")
                entries.Add(New String() {m.Groups(1).Value, id, plain})
                Return "<h" & m.Groups(1).Value & " id=""" & id & """" & m.Groups(2).Value & ">" & m.Groups(3).Value & "</h" & m.Groups(1).Value & ">"
            End Function)

    End Function

    Private Function RpBuildScope(inp As ReportInput, dev As ReportDevice) As ReportScope

        Dim n As Integer = dev.X.Length
        If n < 3 Then Return Nothing

        Dim sc As New ReportScope With {.Dev = dev, .Count = n}
        sc.Mean = dev.X.Average()
        sc.StdDev = ExStdev(dev.X)

        Dim unit As Double = If(inp.MinsPerSample > 0, inp.MinsPerSample, 1.0)
        sc.Ti = New Double(n - 1) {}
        sc.Times = New DateTime(n - 1) {}

        Dim okCount As Integer = 0
        For i As Integer = 0 To n - 1
            sc.Ti(i) = (dev.FirstIndex + i) * unit
            sc.Times(i) = RpParseTime(dev.TimeText(i), inp.Culture)
            If sc.Times(i) > DateTime.MinValue Then okCount += 1
        Next
        sc.TimesOk = okCount > n * 0.98

        Return sc

    End Function

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  One device
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Sub RpDeviceSection(inp As ReportInput, sc As ReportScope, html As StringBuilder, findings As List(Of ReportFinding))

        Dim ctx As New ReportCtx With {
            .Tag = "Dev " & sc.Dev.Slot.ToString(),
            .AbsMean = Math.Abs(sc.Mean),
            .Sd = sc.StdDev,
            .Mps = inp.MinsPerSample
        }

        html.AppendLine("<section>")
        html.AppendLine("<h2>Dev " & sc.Dev.Slot.ToString() & " - " & RpHtml(sc.Dev.Name) & "</h2>")

        html.AppendLine(RpTraceChart(inp, sc))

        RpSectionStats(inp, sc, ctx, html, findings)
        html.AppendLine(RpHistogramChart(sc))
        RpSectionDrift(inp, sc, ctx, html, findings)
        RpSectionQuarters(sc, ctx, html)
        RpSectionTemperature(inp, sc, ctx, html, findings)
        RpSectionPpm(inp, sc, ctx, html)
        RpSectionStability(inp, sc, ctx, html, findings)
        RpSectionPeriodic(inp, sc, ctx, html, findings)
        RpSectionSettling(sc, ctx, html, findings)
        RpSectionEvents(inp, sc, ctx, html, findings)

        html.AppendLine("</section>")

    End Sub

    ' ---- Statistics ----

    Private Sub RpSectionStats(inp As ReportInput, sc As ReportScope, ctx As ReportCtx, html As StringBuilder, findings As List(Of ReportFinding))

        Dim x() As Double = sc.Dev.X
        Dim n As Integer = sc.Count

        Dim sorted() As Double = DirectCast(x.Clone(), Double())
        Array.Sort(sorted)

        Dim median As Double = If(n Mod 2 = 1, sorted(n \ 2), (sorted(n \ 2 - 1) + sorted(n \ 2)) / 2.0)
        Dim minV As Double = sorted(0)
        Dim maxV As Double = sorted(n - 1)
        Dim mean As Double = sc.Mean
        Dim sd As Double = sc.StdDev

        Dim m2 As Double = 0.0
        Dim m3 As Double = 0.0
        Dim m4 As Double = 0.0
        For Each v As Double In x
            Dim d As Double = v - mean
            m2 += d * d
            m3 += d * d * d
            m4 += d * d * d * d
        Next
        m2 /= n
        m3 /= n
        m4 /= n

        Dim skew As Double = If(m2 > 0, m3 / Math.Pow(m2, 1.5), Double.NaN)
        Dim kurtosis As Double = If(m2 > 0, m4 / (m2 * m2) - 3.0, Double.NaN)

        Dim distinct As Integer = 1
        Dim minGap As Double = Double.MaxValue
        For i As Integer = 1 To n - 1
            If sorted(i) <> sorted(i - 1) Then
                distinct += 1
                minGap = Math.Min(minGap, sorted(i) - sorted(i - 1))
            End If
        Next
        ctx.MinGap = minGap
        ctx.PointNoise = RpPointNoise(x)

        RpTableStart(html, "Statistics of the readings")
        RpRow(html, "Mean", RpNum(mean, 10))
        RpRow(html, "Median", RpNum(median, 10), "differs from the mean when the distribution is lop-sided")
        RpRow(html, "STDEV", RpNum(sd, 6), ctx.PpmText(sd) & " of the mean (includes any drift)")
        RpRow(html, "SEM (STDEV / sqrt n)", RpNum(sd / Math.Sqrt(n), 6), ctx.PpmText(sd / Math.Sqrt(n)))
        RpRow(html, "Point-to-point noise", RpNum(ctx.PointNoise, 6), ctx.PpmText(ctx.PointNoise) & ": STDEV of successive differences / sqrt(2), so slow drift does not inflate it")
        RpRow(html, "Minimum", RpNum(minV, 10))
        RpRow(html, "Maximum", RpNum(maxV, 10))
        RpRow(html, "Peak-to-peak", RpNum(maxV - minV, 6), ctx.PpmText(maxV - minV))
        If ctx.PointNoise > 0 AndAlso ctx.AbsMean > 0 Then
            Dim digits As Double = Math.Log10(ctx.AbsMean / (6.6 * ctx.PointNoise))
            RpRow(html, "Effective resolution", digits.ToString("0.0", CultureInfo.InvariantCulture) & " digits",
                  "log10(mean / (6.6 x point-to-point noise)): how many digits of the reading are steady rather than jitter")
        End If
        RpRow(html, "Skew", RpNum(skew, 3), "0 = symmetrical")
        RpRow(html, "Excess kurtosis", RpNum(kurtosis, 3), "0 = bell-shaped, negative = flat or double-humped, positive = long tails")
        RpRow(html, "Distinct values", distinct.ToString(), If(distinct > 1, "smallest step between readings " & RpNum(minGap, 4), ""))

        ' Meter specification comparison
        If Not Double.IsNaN(inp.SpecPpm) AndAlso inp.SpecPpm > 0 AndAlso ctx.AbsMean > 0 Then
            Dim sdPpm As Double = ctx.Ppm(sd)
            Dim ratio As Double = sdPpm / inp.SpecPpm
            RpRow(html, "Meter specification", RpPpmNum(inp.SpecPpm), "STDEV is " & RpNum(ratio, 3) & " x the specification")
            If ratio <= 1 Then
                findings.Add(New ReportFinding("ok", ctx.Tag & ": STDEV " & RpPpmNum(sdPpm) & " is within the " & RpPpmNum(inp.SpecPpm) & " specification."))
            Else
                findings.Add(New ReportFinding("warn", ctx.Tag & ": STDEV " & RpPpmNum(sdPpm) & " is " & RpNum(ratio, 3) & " x the " & RpPpmNum(inp.SpecPpm) & " specification (it includes any drift; point-to-point noise is " & ctx.PpmText(ctx.PointNoise) & ")."))
            End If
        End If

        ' Readings that never occur (missing codes / coarse resolution): look at the grid of the smallest step
        If distinct > 20 AndAlso minGap > 0 AndAlso minGap < Double.MaxValue Then

            Dim range As Double = maxV - minV
            Dim slots As Double = range / minGap

            If slots < 200000 Then

                Dim occupied As New HashSet(Of Long)
                Dim onGrid As Integer = 0

                For Each v As Double In x
                    Dim q As Double = (v - minV) / minGap
                    Dim r As Double = Math.Round(q)
                    If Math.Abs(q - r) < 0.02 Then
                        onGrid += 1
                        occupied.Add(CLng(r))
                    End If
                Next

                If onGrid >= 0.98 * n AndAlso sd / minGap >= 3 Then

                    Dim lowSlot As Long = Math.Max(0L, CLng(Math.Floor((mean - 2 * sd - minV) / minGap)))
                    Dim highSlot As Long = Math.Min(CLng(Math.Round(range / minGap)), CLng(Math.Ceiling((mean + 2 * sd - minV) / minGap)))
                    Dim possible As Long = highSlot - lowSlot + 1
                    Dim present As Long = 0

                    For s As Long = lowSlot To highSlot
                        If occupied.Contains(s) Then present += 1
                    Next

                    If possible >= 8 Then
                        Dim missing As Long = possible - present
                        RpRow(html, "Readings that never occurred", missing.ToString() & " of " & possible.ToString(),
                              "possible values within +/-2 STDEV on the grid of the smallest step (" & RpNum(minGap, 4) & ")")
                        If missing / possible > 0.25 Then
                            findings.Add(New ReportFinding("note", ctx.Tag & ": " & missing.ToString() & " of the " & possible.ToString() &
                                         " possible reading values near the mean never occurred, which suggests missing codes or a coarser real resolution than the smallest step."))
                        End If
                    End If

                End If

            End If

        End If

        RpTableEnd(html)

        findings.Add(New ReportFinding("note", ctx.Tag & " (" & sc.Dev.Name & "): " & n.ToString() & " readings, mean " & RpNum(mean, 8) & ", STDEV " & ctx.PpmText(sd) &
                                       ", point-to-point noise " & ctx.PpmText(ctx.PointNoise) & "."))

        ' Shape of the distribution
        If Not Double.IsNaN(skew) AndAlso Not Double.IsNaN(kurtosis) AndAlso n >= 200 Then
            Dim bimodality As Double = (skew * skew + 1.0) / (kurtosis + 3.0)      ' Sarle's coefficient: above 0.555 hints at more than one hump
            If bimodality > 0.555 Then
                findings.Add(New ReportFinding("note", ctx.Tag & ": the histogram looks lop-sided or double-humped (bimodality coefficient " & RpNum(bimodality, 3) &
                                               "), which usually means drift, a step or interference rather than random noise."))
            End If
        End If

    End Sub

    ' ---- Drift ----

    Private Sub RpSectionDrift(inp As ReportInput, sc As ReportScope, ctx As ReportCtx, html As StringBuilder, findings As List(Of ReportFinding))

        Dim x() As Double = sc.Dev.X
        Dim n As Integer = sc.Count
        Dim mps As Double = inp.MinsPerSample

        Dim slope, r2, slopeErr, meanTi, meanX As Double
        If Not ExRegress(sc.Ti, x, slope, r2, slopeErr, meanTi, meanX) Then Exit Sub

        RpTableStart(html, "Drift")

        If mps > 0 AndAlso ctx.AbsMean > 0 Then

            Dim perHour As Double = slope * 60.0
            Dim ppmHour As Double = ctx.Ppm(perHour)
            Dim durationHours As Double = (n - 1) * mps / 60.0
            Dim totalPpm As Double = ppmHour * durationHours
            Dim ppmHourErr As Double = Math.Abs(ctx.Ppm(slopeErr * 60.0))
            Dim tStat As Double = If(slopeErr > 0, Math.Abs(slope / slopeErr), Double.PositiveInfinity)

            RpRow(html, "Linear trend", RpNum(ppmHour, 4) & " ppm/hour", "+/- " & RpNum(ppmHourErr, 2) & " (formal), R2 = " & RpNum(r2, 3))
            RpRow(html, "Trend, in readings", RpNum(perHour, 4) & " per hour")
            RpRow(html, "Total drift over the scope", RpPpmNum(totalPpm), "trend x duration (" & RpDur((n - 1) * mps) & ")")

            ' Noise left after removing the trend
            Dim sumSq As Double = 0.0
            For i As Integer = 0 To n - 1
                Dim resid As Double = x(i) - (meanX + slope * (sc.Ti(i) - meanTi))
                sumSq += resid * resid
            Next
            Dim residual As Double = Math.Sqrt(sumSq / (n - 2))
            RpRow(html, "STDEV after removing the trend", RpNum(residual, 6), ctx.PpmText(residual))

            ' Beginning against end
            Dim k As Integer = Math.Max(1, n \ 10)
            Dim firstMean As Double = 0.0
            Dim lastMean As Double = 0.0
            For i As Integer = 0 To k - 1
                firstMean += x(i)
                lastMean += x(n - 1 - i)
            Next
            firstMean /= k
            lastMean /= k
            RpRow(html, "Last 10% against first 10%", ctx.PpmText(lastMean - firstMean), "difference of the two averages (" & k.ToString() & " readings each)")

            ' After the warm-up
            Dim afterPpm As Double = Double.NaN
            If inp.WarmUpMinutes > 0 Then
                Dim startIdx As Integer = CInt(Math.Ceiling(inp.WarmUpMinutes / mps))
                If n - startIdx >= 10 Then
                    Dim m As Integer = n - startIdx
                    Dim ti2(m - 1) As Double
                    Dim x2(m - 1) As Double
                    Array.Copy(sc.Ti, startIdx, ti2, 0, m)
                    Array.Copy(x, startIdx, x2, 0, m)
                    Dim s2, q2, e2, mt2, mx2 As Double
                    If ExRegress(ti2, x2, s2, q2, e2, mt2, mx2) Then
                        afterPpm = ctx.Ppm(s2 * 60.0)
                        RpRow(html, "Trend after skipping the first " & RpDur(inp.WarmUpMinutes), RpNum(afterPpm, 4) & " ppm/hour",
                              "+/- " & RpNum(Math.Abs(ctx.Ppm(e2 * 60.0)), 2) & " (formal), R2 = " & RpNum(q2, 3))
                    End If
                Else
                    RpRow(html, "Trend after the warm-up", "n/a", "the run is shorter than the warm-up period")
                End If
            End If

            Dim noiseTotal As Double = ctx.Ppm(ctx.Sd)
            If tStat > 3 AndAlso Math.Abs(totalPpm) > 2 * ctx.Ppm(ctx.PointNoise) Then
                findings.Add(New ReportFinding("note", ctx.Tag & ": drift of " & RpPpmNum(totalPpm) & " over the scope (" & RpNum(ppmHour, 3) & " ppm/hour" &
                             If(Double.IsNaN(afterPpm), "", ", " & RpNum(afterPpm, 3) & " ppm/hour after the warm-up") & ")."))
            Else
                findings.Add(New ReportFinding("ok", ctx.Tag & ": no significant linear drift (" & RpNum(ppmHour, 3) & " ppm/hour)."))
            End If

        Else

            RpRow(html, "Linear trend", RpNum(slope, 4) & " per sample", "R2 = " & RpNum(r2, 3) & " (the time scale of the file is not known, so no per-hour figure)")

        End If

        RpTableEnd(html)

    End Sub

    ' ---- Through the run: four equal parts ----

    Private Sub RpSectionQuarters(sc As ReportScope, ctx As ReportCtx, html As StringBuilder)

        Dim n As Integer = sc.Count
        If n < 40 OrElse ctx.AbsMean <= 0 Then Exit Sub

        html.AppendLine("<h3>Through the run (four equal parts)</h3>")
        html.AppendLine("<table class=""grid""><tr><th>Part</th><th>Time</th><th>Mean</th><th>Mean against the whole run (ppm)</th><th>STDEV (ppm)</th><th>Point-to-point noise (ppm)</th></tr>")

        For q As Integer = 0 To 3
            Dim i0 As Integer = (q * n) \ 4
            Dim i1 As Integer = ((q + 1) * n) \ 4 - 1
            Dim m As Integer = i1 - i0 + 1
            Dim seg(m - 1) As Double
            Array.Copy(sc.Dev.X, i0, seg, 0, m)
            Dim segMean As Double = seg.Average()
            html.AppendLine("<tr><td>" & (q + 1).ToString() & "</td><td>" & RpHtml(RpPosText(sc, i0, ctx) & " to " & RpPosText(sc, i1, ctx)) & "</td><td>" & RpNum(segMean, 10) &
                            "</td><td>" & RpNum(ctx.Ppm(segMean - sc.Mean), 3) & "</td><td>" & RpNum(ctx.Ppm(ExStdev(seg)), 3) & "</td><td>" & RpNum(ctx.Ppm(RpPointNoise(seg)), 3) & "</td></tr>")
        Next

        html.AppendLine("</table>")

    End Sub

    ' ---- PPM deviation and PPM/DegC, as the chart's own PPM box works them out ----

    Private Sub RpSectionPpm(inp As ReportInput, sc As ReportScope, ctx As ReportCtx, html As StringBuilder)

        Dim dev As ReportDevice = sc.Dev
        Dim x() As Double = dev.X
        Dim t() As Double = dev.T
        Dim n As Integer = sc.Count
        Dim v0 As Double = dev.BaseValue
        Dim t0 As Double = dev.BaseTemp

        If v0 = 0 OrElse Double.IsNaN(v0) OrElse Double.IsInfinity(v0) Then Exit Sub

        RpTableStart(html, "PPM deviation and PPM/DegC (as on the chart)")
        RpRow(html, "Baseline", "Initial Value " & RpNum(v0, 10) & ", Initial Temp " & RpNum(t0, 6), dev.BaseSource)

        Dim devMax As Double = Double.MinValue
        Dim devMin As Double = Double.MaxValue
        Dim ptMax As Double = Double.MinValue
        Dim ptMin As Double = Double.MaxValue

        For i As Integer = 0 To n - 1
            Dim d As Double = Math.Max(-ExportPpmClamp, Math.Min(ExportPpmClamp, (x(i) - v0) / v0 * 1000000.0))
            devMax = Math.Max(devMax, d)
            devMin = Math.Min(devMin, d)
            Dim pt As Double = ExInstantPpmDegC(x(i), t(i), v0, t0)
            ptMax = Math.Max(ptMax, pt)
            ptMin = Math.Min(ptMin, pt)
        Next

        Dim lastDev As Double = Math.Max(-ExportPpmClamp, Math.Min(ExportPpmClamp, (x(n - 1) - v0) / v0 * 1000000.0))
        RpRow(html, "PPM Deviation", "last " & RpNum(lastDev, 5) & " ppm", "max " & RpNum(devMax, 5) & ", min " & RpNum(devMin, 5) & " (clamped to +/-" & ExportPpmClamp.ToString("0") & ")")
        RpRow(html, "PPM/DegC (point)", "last " & RpNum(ExInstantPpmDegC(x(n - 1), t(n - 1), v0, t0), 5), "max " & RpNum(ptMax, 5) & ", min " & RpNum(ptMin, 5) & " (0.00000001 where Temp = Initial Temp)")

        Dim fSlope, fR2, fErr, fMeanT, fMeanV As Double
        If ExRegress(t, x, fSlope, fR2, fErr, fMeanT, fMeanV) Then
            RpRow(html, "PPM/DegC (Fit)", RpNum(fSlope / v0 * 1000000.0, 5), "+/- " & RpNum(Math.Abs(fErr / v0 * 1000000.0), 3))
        Else
            RpRow(html, "PPM/DegC (Fit)", "n/a", "the temperature does not change enough to fit")
        End If

        ' Trend: the same fit over the last RMS-window readings
        Dim rollWindow As Integer = Math.Max(2, inp.RmsWindow)
        Dim rs As Integer = Math.Max(0, n - rollWindow)
        Dim rn As Integer = n - rs
        If rn >= 2 Then
            Dim sT As Double = 0.0
            Dim sV As Double = 0.0
            Dim sTV As Double = 0.0
            Dim sTT As Double = 0.0
            For i As Integer = rs To n - 1
                sT += t(i)
                sV += x(i)
                sTV += t(i) * x(i)
                sTT += t(i) * t(i)
            Next
            Dim den As Double = rn * sTT - sT * sT
            If den <> 0 Then
                Dim rSlope As Double = (rn * sTV - sT * sV) / den
                RpRow(html, "PPM/DegC (Trend)", RpNum(Math.Max(-ExportPpmClamp, Math.Min(ExportPpmClamp, rSlope / v0 * 1000000.0)), 5), "last " & rn.ToString() & " readings")
            Else
                RpRow(html, "PPM/DegC (Trend)", "0.00000001", "the temperature is constant over the last " & rn.ToString() & " readings")
            End If
        End If

        RpTableEnd(html)

    End Sub

    ' ---- Temperature ----

    Private Sub RpSectionTemperature(inp As ReportInput, sc As ReportScope, ctx As ReportCtx, html As StringBuilder, findings As List(Of ReportFinding))

        Dim x() As Double = sc.Dev.X
        Dim t() As Double = sc.Dev.T
        Dim h() As Double = sc.Dev.H
        Dim n As Integer = sc.Count

        Dim tMin As Double = t.Min()
        Dim tMax As Double = t.Max()
        Dim swing As Double = tMax - tMin

        RpTableStart(html, "Temperature and humidity")
        RpRow(html, "Temperature", RpNum(t.Average(), 5) & " mean", "min " & RpNum(tMin, 5) & ", max " & RpNum(tMax, 5) & ", swing " & RpNum(swing, 3))

        Dim hMin As Double = h.Min()
        Dim hMax As Double = h.Max()
        If hMax - hMin > 0.0000001 Then
            RpRow(html, "Humidity", RpNum(h.Average(), 4) & " mean", "min " & RpNum(hMin, 4) & ", max " & RpNum(hMax, 4))

            If hMax - hMin >= 1.0 AndAlso ctx.AbsMean > 0 Then
                Dim hSlope, hR2, hErr, hMeanH, hMeanX As Double
                If ExRegress(h, x, hSlope, hR2, hErr, hMeanH, hMeanX) Then
                    Dim weak As Boolean = hR2 < 0.3 OrElse hMax - hMin < 5.0
                    RpRow(html, "Humidity sensitivity", RpNum(ctx.Ppm(hSlope), 3) & " ppm per %RH", "R2 = " & RpNum(hR2, 3) &
                          If(weak, " - weak evidence (small humidity swing or poor fit); humidity and temperature often move together", ""))
                End If
            End If
        End If

        If swing < 0.05 Then

            RpRow(html, "Tempco", "not measurable", "the temperature did not change (or was not logged)")

        ElseIf ctx.AbsMean > 0 Then

            Dim slope, r2, slopeErr, meanT, meanX As Double
            If ExRegress(t, x, slope, r2, slopeErr, meanT, meanX) Then

                Dim tempco As Double = ctx.Ppm(slope)
                Dim tempcoErr As Double = Math.Abs(ctx.Ppm(slopeErr))
                RpRow(html, "Tempco (reading against temperature)", RpNum(tempco, 4) & " ppm/degC", "+/- " & RpNum(tempcoErr, 2) & " (formal), R2 = " & RpNum(r2, 3) &
                      " (temperature explains " & (r2 * 100).ToString("0") & "% of the variation)")

                Dim reasons As New List(Of String)
                If swing < 1.0 Then reasons.Add("the temperature only changed by " & RpNum(swing, 3) & " degC (a few degrees are needed)")
                If r2 < 0.5 Then reasons.Add("temperature explains only " & (r2 * 100).ToString("0") & "% of the variation")
                If n < 30 Then reasons.Add("fewer than 30 readings")

                Dim timeCorr As Double = RpCorrelation(t, sc.Ti)
                If Not Double.IsNaN(timeCorr) AndAlso Math.Abs(timeCorr) > 0.95 Then
                    reasons.Add("the temperature rises or falls steadily with time, so tempco cannot be told apart from drift")
                End If

                If reasons.Count = 0 Then
                    RpRow(html, "Tempco reliability", "usable", "enough temperature swing and a reasonable fit")
                    findings.Add(New ReportFinding("note", ctx.Tag & ": tempco " & RpNum(tempco, 3) & " ppm/degC over a " & RpNum(swing, 3) & " degC swing (R2 " & RpNum(r2, 2) & ")."))
                Else
                    RpRow(html, "Tempco reliability", "not reliable", String.Join("; ", reasons))
                    findings.Add(New ReportFinding("warn", ctx.Tag & ": the tempco figure (" & RpNum(tempco, 3) & " ppm/degC) is not reliable: " & String.Join("; ", reasons) & "."))
                End If

            Else
                RpRow(html, "Tempco", "not measurable", "no usable fit")
            End If

        End If

        RpTableEnd(html)

        If swing >= 0.05 Then
            html.AppendLine(RpLineChart(inp, "Dev " & sc.Dev.Slot.ToString() & ": temperature over the run", sc.Ti, t, "Temperature (degC)", New ScottPlot.Color(Color.FromArgb(190, 50, 50)), False))
        End If
        If hMax - hMin >= 1.0 Then
            html.AppendLine(RpLineChart(inp, "Dev " & sc.Dev.Slot.ToString() & ": humidity over the run", sc.Ti, h, "Humidity (%RH)", New ScottPlot.Color(Color.FromArgb(40, 130, 150)), False))
        End If
        If swing >= 0.05 AndAlso ctx.AbsMean > 0 Then html.AppendLine(RpTempcoChart(sc, ctx))
        If swing >= 0.05 AndAlso ctx.AbsMean > 0 Then RpTemperatureLag(sc, ctx, html, findings)

    End Sub

    ' ---- Stability (Allan deviation) ----

    Private Sub RpSectionStability(inp As ReportInput, sc As ReportScope, ctx As ReportCtx, html As StringBuilder, findings As List(Of ReportFinding))

        Dim n As Integer = sc.Count
        If n < 8 OrElse ctx.AbsMean <= 0 Then Exit Sub

        Dim list As New List(Of Double)(sc.Dev.X)
        Dim adev As List(Of KeyValuePair(Of Integer, Double)) = ComputeAllanDeviation(list, True)
        Dim mdev As List(Of KeyValuePair(Of Integer, Double)) = ComputeModifiedAllanDeviation(list)

        If adev.Count = 0 Then Exit Sub

        Dim mdevByTau As New Dictionary(Of Integer, Double)
        For Each kvp As KeyValuePair(Of Integer, Double) In mdev
            mdevByTau(kvp.Key) = kvp.Value
        Next

        Dim mps As Double = inp.MinsPerSample

        html.AppendLine("<h3>Stability (Allan deviation)</h3>")
        html.AppendLine("<p class=""foot"">How the noise changes with averaging time. White noise falls as 1/sqrt(tau); where the curve flattens or rises, drift or flicker noise has taken over.</p>")
        html.AppendLine("<table class=""grid""><tr><th>tau (readings)</th><th>tau (time)</th><th>ADEV</th><th>ADEV (ppm)</th><th>MDEV (ppm)</th></tr>")

        Dim bestTau As Integer = adev(0).Key
        Dim bestPpm As Double = Double.MaxValue

        For Each kvp As KeyValuePair(Of Integer, Double) In adev
            Dim p As Double = ctx.Ppm(kvp.Value)
            If p < bestPpm Then
                bestPpm = p
                bestTau = kvp.Key
            End If
            html.AppendLine("<tr><td>" & kvp.Key.ToString() & "</td><td>" & If(mps > 0, RpDur(kvp.Key * mps), "n/a") & "</td><td>" & RpNum(kvp.Value, 4) & "</td><td>" &
                            RpNum(p, 4) & "</td><td>" & If(mdevByTau.ContainsKey(kvp.Key), RpNum(ctx.Ppm(mdevByTau(kvp.Key)), 4), "-") & "</td></tr>")
        Next
        html.AppendLine("</table>")

        html.AppendLine(RpAllanChart(sc, ctx, adev))

        Dim lastTau As Integer = adev(adev.Count - 1).Key
        Dim firstPpm As Double = ctx.Ppm(adev(0).Value)

        Dim bestText As String = If(mps > 0, RpDur(bestTau * mps), bestTau.ToString() & " readings")

        If bestTau >= lastTau Then
            findings.Add(New ReportFinding("note", ctx.Tag & ": the Allan deviation is still falling at the longest averaging time (" & If(mps > 0, RpDur(lastTau * mps), lastTau.ToString() & " readings") &
                                           ", " & RpPpmNum(bestPpm) & "), so a longer run or longer averaging would give a lower noise figure."))
        ElseIf bestTau > 1 Then
            findings.Add(New ReportFinding("note", ctx.Tag & ": the best averaging time is about " & bestText & ", where the Allan deviation reaches " & RpPpmNum(bestPpm) &
                                           " (" & RpPpmNum(firstPpm) & " for single readings). Averaging for longer than this does not help."))
        Else
            findings.Add(New ReportFinding("note", ctx.Tag & ": the Allan deviation is lowest for single readings (" & RpPpmNum(bestPpm) & "), so averaging does not reduce the noise (drift or correlated noise dominates)."))
        End If

    End Sub

    ' ---- Settling ----

    Private Sub RpSectionSettling(sc As ReportScope, ctx As ReportCtx, html As StringBuilder, findings As List(Of ReportFinding))

        Dim x() As Double = sc.Dev.X
        Dim n As Integer = sc.Count

        If ctx.Mps <= 0 OrElse n < 40 OrElse ctx.AbsMean <= 0 Then Exit Sub

        Dim w As Integer = Math.Max(5, n \ 200)
        Dim prefix(n) As Double
        For i As Integer = 0 To n - 1
            prefix(i + 1) = prefix(i) + x(i)
        Next

        Dim finalLen As Integer = Math.Max(10, n \ 10)
        Dim finalMean As Double = (prefix(n) - prefix(n - finalLen)) / finalLen
        Dim rollNoisePpm As Double = ctx.Ppm(ctx.PointNoise) / Math.Sqrt(w)

        RpTableStart(html, "Settling")

        Dim durationMins As Double = (n - 1) * ctx.Mps
        Dim measured As New List(Of Double())       ' {tolerance, minutes to settle; -1 = from the start, -2 = not settled before the end}

        For Each tol As Double In New Double() {10.0, 5.0, 2.0, 1.0, 0.5}

            Dim rowLabel As String = "Within +/-" & tol.ToString("0.0#", CultureInfo.InvariantCulture) & " ppm of the final value"

            If tol < 3 * rollNoisePpm Then
                RpRow(html, rowLabel, "not measurable", "below the noise of the " & w.ToString() & "-reading averages used")
                Continue For
            End If

            Dim lastBad As Integer = -1
            For i As Integer = w - 1 To n - 1
                Dim roll As Double = (prefix(i + 1) - prefix(i + 1 - w)) / w
                If Math.Abs(roll - finalMean) / ctx.AbsMean * 1000000.0 > tol Then lastBad = i
            Next

            If lastBad < 0 Then
                RpRow(html, rowLabel, "from the start")
                measured.Add(New Double() {tol, -1})
            Else
                Dim settleMins As Double = sc.Ti(Math.Min(n - 1, lastBad + 1)) - sc.Ti(0)
                If settleMins >= 0.9 * durationMins Then
                    RpRow(html, rowLabel, "not settled before the end of the run", "still moving by more than this late in the run (drift or level shifts)")
                    measured.Add(New Double() {tol, -2})
                Else
                    RpRow(html, rowLabel, "after " & RpDur(settleMins))
                    measured.Add(New Double() {tol, settleMins})
                End If
            End If

        Next

        RpRow(html, "Method", "", "final value = mean of the last " & finalLen.ToString() & " readings; compared with a " & w.ToString() & "-reading moving average")
        RpTableEnd(html)

        ' Finding: the tightest tolerance that was reached, and the loosest one that was not
        Dim reached As Double() = Nothing
        Dim notReachedTol As Double = Double.NaN
        For Each m As Double() In measured
            If m(1) <> -2 Then
                reached = m
            ElseIf Double.IsNaN(notReachedTol) Then
                notReachedTol = m(0)
            End If
        Next

        Dim notReachedText As String = If(Double.IsNaN(notReachedTol), "",
            "; it was still moving by more than +/-" & notReachedTol.ToString("0.0#", CultureInfo.InvariantCulture) & " ppm late in the run (drift or level shifts)")

        If reached Is Nothing Then
            If Not Double.IsNaN(notReachedTol) Then
                findings.Add(New ReportFinding("warn", ctx.Tag & ": did not settle: it was still moving by more than +/-" & notReachedTol.ToString("0.0#", CultureInfo.InvariantCulture) & " ppm of its final value late in the run."))
            End If
        ElseIf reached(1) = -1 Then
            findings.Add(New ReportFinding(If(notReachedText = "", "ok", "note"), ctx.Tag & ": within +/-" & reached(0).ToString("0.0#", CultureInfo.InvariantCulture) & " ppm of its final value from the start of the scope" & notReachedText & "."))
        Else
            findings.Add(New ReportFinding(If(reached(1) > 0.25 * durationMins OrElse notReachedText <> "", "note", "ok"),
                ctx.Tag & ": settled to within +/-" & reached(0).ToString("0.0#", CultureInfo.InvariantCulture) & " ppm of its final value after " & RpDur(reached(1)) & " (of " & RpDur(durationMins) & ")" & notReachedText & "."))
        End If

    End Sub

    ' ---- Events: spikes, steps, gaps, noisy periods ----

    Private Sub RpSectionEvents(inp As ReportInput, sc As ReportScope, ctx As ReportCtx, html As StringBuilder, findings As List(Of ReportFinding))

        Dim x() As Double = sc.Dev.X
        Dim n As Integer = sc.Count
        Dim unitText As String = If(ctx.Mps > 0, "time", "reading")

        html.AppendLine("<h3>Events</h3>")

        ' --- Spikes ---
        Dim sigma As Double = 0.0
        Dim spikes As List(Of Double()) = RpFindSpikes(x, ctx.MinGap, sigma)
        sc.Spikes = New List(Of Double())(spikes)

        If sigma > 0 Then
            If spikes.Count = 0 Then
                html.AppendLine("<p>Spikes: none found (a spike is a reading more than 6 x " & RpHtml(RpNum(sigma, 3)) & " away from the median of its neighbours).</p>")
                findings.Add(New ReportFinding("ok", ctx.Tag & ": no spikes or dropouts."))
            Else
                spikes.Sort(Function(lhs As Double(), rhs As Double()) Math.Abs(rhs(1)).CompareTo(Math.Abs(lhs(1))))
                html.AppendLine("<p>Spikes: <b>" & spikes.Count.ToString() & "</b> found (a reading more than 6 x " & RpHtml(RpNum(sigma, 3)) & " away from the median of its 11 neighbours). The largest:</p>")
                html.AppendLine("<table class=""grid""><tr><th>" & unitText & "</th><th>size</th><th>ppm</th><th>readings affected</th></tr>")
                For i As Integer = 0 To Math.Min(9, spikes.Count - 1)
                    Dim sp As Double() = spikes(i)
                    html.AppendLine("<tr><td>" & RpPosText(sc, CInt(sp(0)), ctx) & "</td><td>" & RpNum(sp(1), 4) & "</td><td>" & RpNum(ctx.Ppm(sp(1)), 3) & "</td><td>" & CInt(sp(2)).ToString() & "</td></tr>")
                Next
                html.AppendLine("</table>")
                Dim biggest As Double() = spikes(0)
                findings.Add(New ReportFinding("warn", ctx.Tag & ": " & spikes.Count.ToString() & " spike(s); the largest is " & RpPpmNum(Math.Abs(ctx.Ppm(biggest(1)))) & " at " & RpPosText(sc, CInt(biggest(0)), ctx) & "."))
            End If
        End If

        ' --- Steps ---
        Dim steps As List(Of Double()) = RpFindSteps(x, ctx.PointNoise)
        sc.Steps = New List(Of Double())(steps)
        If steps.Count > 0 Then
            html.AppendLine("<p>Steps or sudden shifts in level (found by comparing the averages just before and after each point, after removing the overall trend):</p>")
            html.AppendLine("<table class=""grid""><tr><th>" & unitText & "</th><th>shift</th><th>ppm</th><th>strength</th></tr>")
            For Each st As Double() In steps
                html.AppendLine("<tr><td>" & RpPosText(sc, CInt(st(0)), ctx) & "</td><td>" & RpNum(st(1), 4) & "</td><td>" & RpNum(ctx.Ppm(st(1)), 3) & "</td><td>" & RpNum(Math.Abs(st(2)), 3) & "</td></tr>")
            Next
            html.AppendLine("</table>")
            For i As Integer = 0 To Math.Min(2, steps.Count - 1)
                findings.Add(New ReportFinding("warn", ctx.Tag & ": level shift of " & RpPpmNum(ctx.Ppm(steps(i)(1))) & " at " & RpPosText(sc, CInt(steps(i)(0)), ctx) & "."))
            Next
        ElseIf n >= 160 Then
            html.AppendLine("<p>Steps: no sudden shifts in level found.</p>")
        End If

        ' --- Gaps in the timestamps ---
        If sc.TimesOk Then
            Dim intervals As New List(Of Double)
            For i As Integer = 1 To n - 1
                If sc.Times(i) > DateTime.MinValue AndAlso sc.Times(i - 1) > DateTime.MinValue Then
                    intervals.Add((sc.Times(i) - sc.Times(i - 1)).TotalSeconds)
                Else
                    intervals.Add(Double.NaN)
                End If
            Next
            Dim valid As Double() = intervals.Where(Function(v) Not Double.IsNaN(v)).ToArray()
            If valid.Length > 10 Then
                Dim medianStep As Double = RpMedian(valid)
                Dim backwards As Integer = valid.Count(Function(v) v < 0)
                If medianStep >= 1 Then
                    Dim threshold As Double = Math.Max(3 * medianStep, medianStep + 2)
                    Dim gaps As New List(Of Double())
                    For i As Integer = 0 To intervals.Count - 1
                        If Not Double.IsNaN(intervals(i)) AndAlso intervals(i) > threshold Then gaps.Add(New Double() {i + 1, intervals(i)})
                    Next
                    html.AppendLine("<p>Sampling: median interval " & RpHtml(RpNum(medianStep, 3)) & " s" & If(backwards > 0, "; <span class=""warn"">" & backwards.ToString() & " timestamp(s) go backwards</span>", "") & ".")
                    If gaps.Count = 0 Then
                        html.AppendLine(" No gaps found.</p>")
                    Else
                        gaps.Sort(Function(lhs As Double(), rhs As Double()) rhs(1).CompareTo(lhs(1)))
                        html.AppendLine(" <b>" & gaps.Count.ToString() & " gap(s)</b> longer than " & RpHtml(RpNum(threshold, 3)) & " s; the longest:</p>")
                        html.AppendLine("<table class=""grid""><tr><th>" & unitText & "</th><th>gap</th></tr>")
                        For i As Integer = 0 To Math.Min(4, gaps.Count - 1)
                            html.AppendLine("<tr><td>" & RpPosText(sc, CInt(gaps(i)(0)), ctx) & "</td><td>" & RpHtml(RpDur(gaps(i)(1) / 60.0)) & "</td></tr>")
                        Next
                        html.AppendLine("</table>")
                        findings.Add(New ReportFinding("warn", ctx.Tag & ": " & gaps.Count.ToString() & " gap(s) in the readings, the longest " & RpDur(gaps(0)(1) / 60.0) & " at " & RpPosText(sc, CInt(gaps(0)(0)), ctx) & "."))
                    End If
                    If backwards > 0 Then findings.Add(New ReportFinding("warn", ctx.Tag & ": " & backwards.ToString() & " timestamp(s) run backwards (clock change or a damaged file)."))
                End If
            End If
        End If

        ' --- Noisier periods ---
        Dim blockSize As Integer = Math.Max(30, n \ 50)
        If n >= blockSize * 5 Then
            Dim blockCount As Integer = n \ blockSize
            Dim blockNoise(blockCount - 1) As Double
            For b As Integer = 0 To blockCount - 1
                Dim seg(blockSize - 1) As Double
                Array.Copy(x, b * blockSize, seg, 0, blockSize)
                blockNoise(b) = RpPointNoise(seg)
            Next
            Dim typical As Double = RpMedian(blockNoise)
            If typical > 0 Then
                Dim chartX(blockCount - 1) As Double
                Dim chartY(blockCount - 1) As Double
                For b As Integer = 0 To blockCount - 1
                    chartX(b) = sc.Ti(b * blockSize + blockSize \ 2)
                    chartY(b) = ctx.Ppm(blockNoise(b))
                Next
                html.AppendLine(RpLineChart(inp, "Dev " & sc.Dev.Slot.ToString() & ": noise through the run (point-to-point noise of each block of " & blockSize.ToString() & " readings)",
                                            chartX, chartY, "Noise (ppm of mean)", sc.Colour, False))

                Dim noisy As New List(Of Double())
                For b As Integer = 0 To blockCount - 1
                    If blockNoise(b) > 2 * typical Then noisy.Add(New Double() {b, blockNoise(b) / typical})
                Next
                If noisy.Count > 0 Then
                    noisy.Sort(Function(lhs As Double(), rhs As Double()) rhs(1).CompareTo(lhs(1)))
                    html.AppendLine("<p>Noisier periods (point-to-point noise more than twice the typical " & RpHtml(ctx.PpmText(typical)) & ", in blocks of " & blockSize.ToString() & " readings):</p>")
                    html.AppendLine("<table class=""grid""><tr><th>" & unitText & "</th><th>noise</th><th>x typical</th></tr>")
                    For i As Integer = 0 To Math.Min(4, noisy.Count - 1)
                        Dim b As Integer = CInt(noisy(i)(0))
                        html.AppendLine("<tr><td>" & RpPosText(sc, b * blockSize, ctx) & " to " & RpPosText(sc, b * blockSize + blockSize - 1, ctx) & "</td><td>" & RpHtml(ctx.PpmText(blockNoise(b))) &
                                        "</td><td>" & RpNum(noisy(i)(1), 3) & "</td></tr>")
                    Next
                    html.AppendLine("</table>")
                    Dim worst As Integer = CInt(noisy(0)(0))
                    findings.Add(New ReportFinding("note", ctx.Tag & ": noisiest period around " & RpPosText(sc, worst * blockSize, ctx) & " (" & RpNum(noisy(0)(1), 3) & " x the typical noise)."))
                End If
            End If
        End If

    End Sub

    ' Position text for a reading index within the scope: time since the start of the run, or the reading number.
    Private Function RpPosText(sc As ReportScope, index As Integer, ctx As ReportCtx) As String

        index = Math.Max(0, Math.Min(sc.Count - 1, index))
        If ctx.Mps > 0 Then Return RpDur(sc.Ti(index))
        Return (sc.Dev.FirstIndex + index).ToString()

    End Function

    ' Readings far from the median of their 11 neighbours. Each result: {index of the biggest reading, its deviation, readings in the event}.
    Private Function RpFindSpikes(x() As Double, minGap As Double, ByRef sigma As Double) As List(Of Double())

        Dim events As New List(Of Double())
        Dim n As Integer = x.Length
        If n < 20 Then
            sigma = 0
            Return events
        End If

        Dim resid(n - 1) As Double
        For i As Integer = 0 To n - 1
            Dim a As Integer = Math.Max(0, i - 5)
            Dim b As Integer = Math.Min(n - 1, i + 5)
            Dim seg(b - a) As Double
            Array.Copy(x, a, seg, 0, b - a + 1)
            Array.Sort(seg)
            resid(i) = x(i) - seg((b - a + 1) \ 2)
        Next

        sigma = 1.4826 * RpMad(resid)
        If minGap < Double.MaxValue Then sigma = Math.Max(sigma, 0.5 * minGap)
        If sigma <= 0 Then Return events

        Dim limit As Double = 6.0 * sigma
        Dim i2 As Integer = 0

        Do While i2 < n
            If Math.Abs(resid(i2)) > limit Then
                Dim peak As Integer = i2
                Dim count As Integer = 0
                Dim j As Integer = i2
                Dim lastHit As Integer = i2
                Do While j < n AndAlso j <= lastHit + 1
                    If Math.Abs(resid(j)) > limit Then
                        lastHit = j
                        count += 1
                        If Math.Abs(resid(j)) > Math.Abs(resid(peak)) Then peak = j
                    End If
                    j += 1
                Loop
                events.Add(New Double() {peak, resid(peak), count})
                i2 = lastHit + 1
            Else
                i2 += 1
            End If
        Loop

        Return events

    End Function

    ' Sudden shifts in level. Each result: {index where the level changes, size of the shift, strength}.
    Private Function RpFindSteps(x() As Double, noise As Double) As List(Of Double())

        Dim results As New List(Of Double())
        Dim n As Integer = x.Length
        Dim w As Integer = Math.Max(20, Math.Min(500, n \ 40))

        If n < 4 * w OrElse Double.IsNaN(noise) OrElse noise <= 0 Then Return results

        ' Remove the overall linear trend first, so a steady drift is not mistaken for steps
        Dim idx(n - 1) As Double
        For i As Integer = 0 To n - 1
            idx(i) = i
        Next
        Dim slope, r2, slopeErr, mx, my As Double
        If Not ExRegress(idx, x, slope, r2, slopeErr, mx, my) Then Return results

        Dim prefix(n) As Double
        For i As Integer = 0 To n - 1
            prefix(i + 1) = prefix(i) + (x(i) - (my + slope * (i - mx)))
        Next

        Dim candidates As New List(Of Double())
        For i As Integer = w To n - w
            Dim diff As Double = (prefix(i + w) - prefix(i)) / w - (prefix(i) - prefix(i - w)) / w
            Dim z As Double = diff / (noise * Math.Sqrt(2.0 / w))
            If Math.Abs(z) > 12 AndAlso Math.Abs(diff) > 2 * noise Then candidates.Add(New Double() {i, diff, z})
        Next

        candidates.Sort(Function(lhs As Double(), rhs As Double()) Math.Abs(rhs(2)).CompareTo(Math.Abs(lhs(2))))

        For Each c As Double() In candidates
            Dim tooClose As Boolean = False
            For Each accepted As Double() In results
                If Math.Abs(accepted(0) - c(0)) < w Then
                    tooClose = True
                    Exit For
                End If
            Next
            If Not tooClose Then results.Add(c)
            If results.Count >= 6 Then Exit For
        Next

        Return results

    End Function

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Periodic variation (spectrum and autocorrelation)
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Class ReportSpectrum
        Public Period() As Double                       ' minutes, for each plotted bin (longest period last)
        Public Amp() As Double                          ' amplitude, ppm of the mean reading
        Public Peaks As New List(Of Double())           ' {period (min), amplitude (ppm), strength (x local noise level)}
        Public ShortestPeriod As Double                 ' minutes (twice the sample interval)
        Public LongestPeriod As Double                  ' minutes (a third of the run)
        Public AcfLagMinutes As Double = Double.NaN     ' first clear repeat found by the autocorrelation
        Public AcfValue As Double = Double.NaN
    End Class

    ' In-place radix-2 FFT (array length must be a power of 2).
    Private Sub RpFft(re() As Double, im() As Double, inverse As Boolean)

        Dim n As Integer = re.Length

        Dim j As Integer = 0
        For i As Integer = 1 To n - 1
            Dim bit As Integer = n >> 1
            Do While (j And bit) <> 0
                j = j Xor bit
                bit = bit >> 1
            Loop
            j = j Xor bit
            If i < j Then
                Dim tr As Double = re(i)
                re(i) = re(j)
                re(j) = tr
                Dim ti As Double = im(i)
                im(i) = im(j)
                im(j) = ti
            End If
        Next

        Dim blockLen As Integer = 2
        Do While blockLen <= n

            Dim angle As Double = 2.0 * Math.PI / blockLen * If(inverse, 1.0, -1.0)
            Dim wStepRe As Double = Math.Cos(angle)
            Dim wStepIm As Double = Math.Sin(angle)
            Dim half As Integer = blockLen \ 2

            For start As Integer = 0 To n - 1 Step blockLen
                Dim wRe As Double = 1.0
                Dim wIm As Double = 0.0
                For k As Integer = 0 To half - 1
                    Dim a As Integer = start + k
                    Dim b As Integer = a + half
                    Dim vRe As Double = re(b) * wRe - im(b) * wIm
                    Dim vIm As Double = re(b) * wIm + im(b) * wRe
                    re(b) = re(a) - vRe
                    im(b) = im(a) - vIm
                    re(a) = re(a) + vRe
                    im(a) = im(a) + vIm
                    Dim nextRe As Double = wRe * wStepRe - wIm * wStepIm
                    wIm = wRe * wStepIm + wIm * wStepRe
                    wRe = nextRe
                Next
            Next

            blockLen = blockLen << 1

        Loop

        If inverse Then
            For i As Integer = 0 To n - 1
                re(i) /= n
                im(i) /= n
            Next
        End If

    End Sub

    ' Straight-line fit removed from the readings (so the spectrum and correlations are about the wiggles, not the drift).
    Private Function RpDetrend(a() As Double) As Double()

        Dim n As Integer = a.Length
        Dim result(n - 1) As Double
        If n < 3 Then
            Array.Copy(a, result, n)
            Return result
        End If

        Dim meanI As Double = (n - 1) / 2.0
        Dim meanA As Double = a.Average()
        Dim sxy As Double = 0.0
        Dim sxx As Double = 0.0
        For i As Integer = 0 To n - 1
            sxy += (i - meanI) * (a(i) - meanA)
            sxx += (i - meanI) * (i - meanI)
        Next
        Dim slope As Double = If(sxx > 0, sxy / sxx, 0.0)

        For i As Integer = 0 To n - 1
            result(i) = a(i) - (meanA + slope * (i - meanI))
        Next

        Return result

    End Function

    ' Averages the readings in blocks of 'factor' (to keep the FFT size sensible on very long runs).
    Private Function RpBlockMeans(a() As Double, factor As Integer) As Double()

        If factor <= 1 Then Return DirectCast(a.Clone(), Double())

        Dim m As Integer = a.Length \ factor
        Dim result(m - 1) As Double
        For k As Integer = 0 To m - 1
            Dim sum As Double = 0.0
            For i As Integer = k * factor To k * factor + factor - 1
                sum += a(i)
            Next
            result(k) = sum / factor
        Next

        Return result

    End Function

    Private Function RpComputeSpectrum(x() As Double, minsPerSample As Double, absMean As Double) As ReportSpectrum

        If minsPerSample <= 0 OrElse absMean <= 0 OrElse x.Length < 64 Then Return Nothing

        Const maxFft As Integer = 2097152

        Dim factor As Integer = 1
        Do While x.Length \ factor > maxFft
            factor += 1
        Loop

        Dim y() As Double = RpDetrend(RpBlockMeans(x, factor))
        Dim m As Integer = y.Length
        Dim dt As Double = minsPerSample * factor          ' minutes between the values in y
        Dim duration As Double = m * dt

        Dim size As Integer = 1
        Do While size < m
            size = size << 1
        Loop

        ' Hann window; amplitude scaled so a sine of amplitude A shows as A
        Dim re(size - 1) As Double
        Dim im(size - 1) As Double
        Dim windowSum As Double = 0.0
        For i As Integer = 0 To m - 1
            Dim w As Double = 0.5 * (1.0 - Math.Cos(2.0 * Math.PI * i / (m - 1)))
            windowSum += w
            re(i) = y(i) * w
        Next

        RpFft(re, im, False)

        Dim bins As Integer = size \ 2
        Dim amp(bins) As Double                            ' ppm
        For k As Integer = 1 To bins - 1
            amp(k) = 2.0 * Math.Sqrt(re(k) * re(k) + im(k) * im(k)) / windowSum / absMean * 1000000.0
        Next

        Dim spec As New ReportSpectrum With {.ShortestPeriod = 2.0 * dt, .LongestPeriod = duration / 3.0}

        ' Local noise level of each bin: the average of ln(amplitude) over the neighbouring frequencies (a Rayleigh noise amplitude has a mean ln of
        ' ln(scale) + 0.058). Peaks must also stand above it by a wide margin, because many bins are tested.
        Dim logPrefix(bins) As Double
        For k As Integer = 1 To bins - 1
            logPrefix(k + 1) = logPrefix(k) + Math.Log(Math.Max(amp(k), 1.0E-300))
        Next

        Dim firstBin As Integer = CInt(Math.Ceiling(3.0 * size / m))      ' at least three cycles in the run
        Dim candidates As New List(Of Double())

        For k As Integer = Math.Max(2, firstBin) To bins - 2
            If amp(k) > amp(k - 1) AndAlso amp(k) >= amp(k + 1) Then
                Dim lo As Integer = Math.Max(1, (k * 2) \ 3)
                Dim hi As Integer = Math.Min(bins - 1, (k * 3) \ 2 + 2)
                Dim count As Integer = hi - lo + 1
                If count >= 8 Then
                    Dim scale As Double = Math.Exp((logPrefix(hi + 1) - logPrefix(lo)) / count - 0.058)
                    If scale > 0 AndAlso amp(k) > 6.0 * scale Then
                        candidates.Add(New Double() {k, amp(k), amp(k) / scale})
                    End If
                End If
            End If
        Next

        candidates.Sort(Function(lhs As Double(), rhs As Double()) rhs(1).CompareTo(lhs(1)))
        Dim spacing As Double = Math.Max(3.0, 2.0 * size / m)

        For Each c As Double() In candidates
            Dim tooClose As Boolean = False
            For Each accepted As Double() In spec.Peaks
                ' accepted peaks are stored by period; compare in bins
                If Math.Abs(size * dt / accepted(0) - c(0)) < spacing Then
                    tooClose = True
                    Exit For
                End If
            Next
            If Not tooClose Then spec.Peaks.Add(New Double() {size * dt / c(0), c(1), c(2)})
            If spec.Peaks.Count >= 6 Then Exit For
        Next

        ' Curve for the chart: the largest amplitude in each of up to 400 log-spaced period bins
        Dim plotCount As Integer = 400
        Dim logMin As Double = Math.Log10(size * dt / (bins - 1))
        Dim logMax As Double = Math.Log10(size * dt / Math.Max(2, firstBin))
        If logMax > logMin Then
            Dim bestAmp(plotCount - 1) As Double
            For k As Integer = Math.Max(2, firstBin) To bins - 1
                Dim pos As Integer = CInt(Math.Floor((Math.Log10(size * dt / k) - logMin) / (logMax - logMin) * (plotCount - 1)))
                If pos >= 0 AndAlso pos < plotCount AndAlso amp(k) > bestAmp(pos) Then bestAmp(pos) = amp(k)
            Next
            Dim periods As New List(Of Double)
            Dim amps As New List(Of Double)
            For p As Integer = 0 To plotCount - 1
                If bestAmp(p) > 0 Then
                    periods.Add(Math.Pow(10.0, logMin + (logMax - logMin) * p / (plotCount - 1)))
                    amps.Add(bestAmp(p))
                End If
            Next
            spec.Period = periods.ToArray()
            spec.Amp = amps.ToArray()
        End If

        ' Autocorrelation (through the FFT): the first clear repeat after the correlation has dropped away
        Try
            Dim acfSize As Integer = 1
            Do While acfSize < 2 * m
                acfSize = acfSize << 1
            Loop
            Dim ar(acfSize - 1) As Double
            Dim ai(acfSize - 1) As Double
            Array.Copy(y, ar, m)
            RpFft(ar, ai, False)
            For i As Integer = 0 To acfSize - 1
                ar(i) = ar(i) * ar(i) + ai(i) * ai(i)
                ai(i) = 0.0
            Next
            RpFft(ar, ai, True)

            If ar(0) > 0 Then
                Dim maxLag As Integer = m \ 3
                Dim crossed As Boolean = False
                Dim bestLag As Integer = 0
                Dim bestVal As Double = 0.0
                For lag As Integer = 1 To maxLag - 1
                    Dim r As Double = ar(lag) / ar(0)
                    If Not crossed Then
                        If r <= 0 Then crossed = True
                    ElseIf r > bestVal AndAlso r > ar(lag - 1) / ar(0) AndAlso r >= ar(lag + 1) / ar(0) Then
                        bestVal = r
                        bestLag = lag
                    End If
                Next
                If bestLag > 0 AndAlso bestVal >= 0.3 Then
                    spec.AcfLagMinutes = bestLag * dt
                    spec.AcfValue = bestVal
                End If
            End If
        Catch
        End Try

        Return spec

    End Function

    Private Sub RpSectionPeriodic(inp As ReportInput, sc As ReportScope, ctx As ReportCtx, html As StringBuilder, findings As List(Of ReportFinding))

        If ctx.Mps <= 0 OrElse ctx.AbsMean <= 0 OrElse sc.Count < 64 Then Exit Sub

        Dim spec As ReportSpectrum = RpComputeSpectrum(sc.Dev.X, ctx.Mps, ctx.AbsMean)
        If spec Is Nothing Then Exit Sub

        html.AppendLine("<h3>Periodic variation (spectrum)</h3>")
        html.AppendLine("<p class=""foot"">The drift is removed, then the readings are split into their cyclic components: a regular cycle (air-conditioning, a heater, mains pick-up) shows as a peak. " &
                        "Periods from " & RpHtml(RpDur(spec.ShortestPeriod)) & " to " & RpHtml(RpDur(spec.LongestPeriod)) & " (at least three cycles in the run) are searched. Assumes evenly spaced readings.</p>")

        If spec.Peaks.Count = 0 Then
            html.AppendLine("<p>No clear periodic component found: nothing stands well above the surrounding noise level.</p>")
            findings.Add(New ReportFinding("ok", ctx.Tag & ": no regular cycles in the readings (spectrum)."))
        Else
            html.AppendLine("<table class=""grid""><tr><th>Period</th><th>Cycles per hour</th><th>Amplitude</th><th>Strength (x local noise level)</th></tr>")
            For Each pk As Double() In spec.Peaks
                html.AppendLine("<tr><td>" & RpHtml(RpDur(pk(0))) & "</td><td>" & RpNum(60.0 / pk(0), 4) & "</td><td>" & RpNum(pk(1), 3) & " ppm</td><td>" & RpNum(pk(2), 3) & "</td></tr>")
            Next
            html.AppendLine("</table>")

            Dim biggest As Double() = spec.Peaks(0)
            For i As Integer = 0 To Math.Min(2, spec.Peaks.Count - 1)
                Dim pk As Double() = spec.Peaks(i)
                findings.Add(New ReportFinding(If(pk(1) > 2.0 * ctx.Ppm(ctx.PointNoise), "warn", "note"),
                    ctx.Tag & ": regular cycle with a period of " & RpDur(pk(0)) & " and amplitude " & RpPpmNum(pk(1)) & " (" & RpNum(pk(2), 3) & " x the local noise level)."))
            Next
        End If

        If Not Double.IsNaN(spec.AcfLagMinutes) Then
            html.AppendLine("<p>Autocorrelation: the readings repeat after about <b>" & RpHtml(RpDur(spec.AcfLagMinutes)) & "</b> (correlation " & RpHtml(RpNum(spec.AcfValue, 3)) & " at that lag).</p>")
            findings.Add(New ReportFinding("note", ctx.Tag & ": the readings tend to repeat every " & RpDur(spec.AcfLagMinutes) & " (autocorrelation " & RpNum(spec.AcfValue, 2) & ")."))
        End If

        html.AppendLine(RpSpectrumChart(sc, spec))

    End Sub

    Private Function RpSpectrumChart(sc As ReportScope, spec As ReportSpectrum) As String

        Try
            If spec.Period Is Nothing OrElse spec.Period.Length < 3 Then Return ""

            Dim lx(spec.Period.Length - 1) As Double
            Dim ly(spec.Period.Length - 1) As Double
            For i As Integer = 0 To lx.Length - 1
                lx(i) = Math.Log10(spec.Period(i))
                ly(i) = Math.Log10(spec.Amp(i))
            Next

            Dim plot As New ScottPlot.Plot()

            Dim curve As ScottPlot.Plottables.Scatter = plot.Add.Scatter(lx, ly)
            curve.Color = sc.Colour
            curve.LineWidth = 1.5F
            curve.MarkerStyle.IsVisible = False

            If spec.Peaks.Count > 0 Then
                Dim px(spec.Peaks.Count - 1) As Double
                Dim py(spec.Peaks.Count - 1) As Double
                For i As Integer = 0 To spec.Peaks.Count - 1
                    px(i) = Math.Log10(spec.Peaks(i)(0))
                    py(i) = Math.Log10(spec.Peaks(i)(1))
                Next
                Dim marks As ScottPlot.Plottables.Scatter = plot.Add.Scatter(px, py)
                marks.LineWidth = 0
                marks.MarkerSize = 10
                marks.Color = New ScottPlot.Color(Color.FromArgb(200, 30, 30))
            End If

            Dim xTicks As New ScottPlot.TickGenerators.NumericManual()
            For e As Integer = CInt(Math.Floor(lx.Min())) To CInt(Math.Ceiling(lx.Max()))
                xTicks.AddMajor(e, Math.Pow(10, e).ToString("G", CultureInfo.InvariantCulture))
            Next
            plot.Axes.Bottom.TickGenerator = xTicks

            Dim yTicks As New ScottPlot.TickGenerators.NumericManual()
            For e As Integer = CInt(Math.Floor(ly.Min())) To CInt(Math.Ceiling(ly.Max()))
                yTicks.AddMajor(e, Math.Pow(10, e).ToString("G", CultureInfo.InvariantCulture))
            Next
            plot.Axes.Left.TickGenerator = yTicks

            plot.Title("Dev " & sc.Dev.Slot.ToString() & ": spectrum of the readings (drift removed; red = clear cycles)", 14)
            plot.XLabel("Period of the cycle (minutes)", 12)
            plot.YLabel("Amplitude (ppm of mean)", 12)

            Return RpSvg(plot, 1000, 340)
        Catch
            Return ""
        End Try

    End Function

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Temperature lag
    ' ---------------------------------------------------------------------------------------------------------------------------

    ' Correlation of a(i) with b(i + lag) over the readings both have.
    Private Function RpLagCorrelation(a() As Double, b() As Double, lag As Integer) As Double

        Dim n As Integer = a.Length
        Dim i0 As Integer = Math.Max(0, -lag)
        Dim i1 As Integer = Math.Min(n - 1, n - 1 - lag)
        If i1 - i0 < 10 Then Return Double.NaN

        Dim sa As Double = 0.0
        Dim sb As Double = 0.0
        Dim saa As Double = 0.0
        Dim sbb As Double = 0.0
        Dim sab As Double = 0.0
        Dim count As Integer = i1 - i0 + 1

        For i As Integer = i0 To i1
            Dim va As Double = a(i)
            Dim vb As Double = b(i + lag)
            sa += va
            sb += vb
            saa += va * va
            sbb += vb * vb
            sab += va * vb
        Next

        Dim varA As Double = saa - sa * sa / count
        Dim varB As Double = sbb - sb * sb / count
        If varA <= 0 OrElse varB <= 0 Then Return Double.NaN

        Return (sab - sa * sb / count) / Math.Sqrt(varA * varB)

    End Function

    Private Sub RpTemperatureLag(sc As ReportScope, ctx As ReportCtx, html As StringBuilder, findings As List(Of ReportFinding))

        Dim n As Integer = sc.Count
        If ctx.Mps <= 0 OrElse ctx.AbsMean <= 0 OrElse n < 100 Then Exit Sub

        Dim factor As Integer = Math.Max(1, n \ 12000)
        Dim xd() As Double = RpDetrend(RpBlockMeans(sc.Dev.X, factor))
        Dim td() As Double = RpDetrend(RpBlockMeans(sc.Dev.T, factor))
        Dim m As Integer = xd.Length
        Dim maxLag As Integer = Math.Min(m \ 4, 400)
        If maxLag < 3 Then Exit Sub

        ' reading(i + lag) against temperature(i): a positive lag means the reading follows the temperature
        Dim bestLag As Integer = 0
        Dim bestR As Double = 0.0
        Dim zeroR As Double = RpLagCorrelation(td, xd, 0)

        For lag As Integer = -maxLag To maxLag
            Dim r As Double = RpLagCorrelation(td, xd, lag)
            If Not Double.IsNaN(r) AndAlso Math.Abs(r) > Math.Abs(bestR) Then
                bestR = r
                bestLag = lag
            End If
        Next

        html.AppendLine("<h3>Temperature lag</h3>")

        If Double.IsNaN(zeroR) OrElse Math.Abs(bestR) < 0.3 Then
            html.AppendLine("<p>Once the overall trends are removed, the readings show no clear relation to the temperature at any delay up to " & RpHtml(RpDur(maxLag * factor * ctx.Mps)) & " (strongest correlation " & RpHtml(RpNum(bestR, 2)) & ").</p>")
            Exit Sub
        End If

        Dim lagMinutes As Double = bestLag * factor * ctx.Mps

        html.AppendLine("<table class=""data"">")
        Dim meaningfulLag As Boolean = bestLag <> 0 AndAlso Math.Abs(bestR) > Math.Abs(zeroR) + 0.05
        RpRow(html, "Best match", If(Not meaningfulLag, "no meaningful delay", RpDur(Math.Abs(lagMinutes)) & If(bestLag > 0, ": the reading follows the temperature", ": the temperature sensor follows the reading")),
              "correlation " & RpNum(bestR, 3) & " at the best delay (" & RpDur(Math.Abs(lagMinutes)) & "), " & RpNum(zeroR, 3) & " with no delay (both with the straight-line trend removed)")

        ' Tempco at that delay (full resolution): reading(i + lag) against temperature(i)
        Dim fullLag As Integer = bestLag * factor
        Dim i0 As Integer = Math.Max(0, -fullLag)
        Dim i1 As Integer = Math.Min(n - 1, n - 1 - fullLag)
        If meaningfulLag AndAlso i1 - i0 >= 30 Then
            Dim cnt As Integer = i1 - i0 + 1
            Dim tl(cnt - 1) As Double
            Dim xl(cnt - 1) As Double
            For i As Integer = 0 To cnt - 1
                tl(i) = sc.Dev.T(i0 + i)
                xl(i) = sc.Dev.X(i0 + i + fullLag)
            Next
            Dim slope, r2, slopeErr, mT, mX As Double
            Dim slope0, r20, err0, mT0, mX0 As Double
            If ExRegress(tl, xl, slope, r2, slopeErr, mT, mX) Then
                Dim noLagText As String = ""
                If ExRegress(sc.Dev.T, sc.Dev.X, slope0, r20, err0, mT0, mX0) Then noLagText = "; with no delay R2 = " & RpNum(r20, 3)
                RpRow(html, "Tempco allowing for that delay", RpNum(ctx.Ppm(slope), 4) & " ppm/degC", "+/- " & RpNum(Math.Abs(ctx.Ppm(slopeErr)), 2) & " (formal), R2 = " & RpNum(r2, 3) & noLagText)
            End If
        End If
        html.AppendLine("</table>")

        If bestLag <> 0 AndAlso Math.Abs(bestR) > Math.Abs(zeroR) + 0.05 Then
            findings.Add(New ReportFinding("note", ctx.Tag & If(bestLag > 0, ": the reading follows the temperature with a delay of about " & RpDur(Math.Abs(lagMinutes)),
                         ": the temperature sensor follows the reading by about " & RpDur(Math.Abs(lagMinutes))) & " (correlation " & RpNum(bestR, 2) & " at that delay against " & RpNum(zeroR, 2) & " with none), so a tempco fitted with no delay will understate the real figure."))
        End If

    End Sub

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Allan deviation of the difference between two devices
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Sub RpCompareStability(inp As ReportInput, a As ReportScope, b As ReportScope, d() As Double, html As StringBuilder, findings As List(Of ReportFinding))

        Dim n As Integer = d.Length
        Dim meanA As Double = Math.Abs(a.Mean)
        Dim meanB As Double = Math.Abs(b.Mean)
        If n < 8 OrElse meanA <= 0 OrElse meanB <= 0 Then Exit Sub

        Dim xa As New List(Of Double)(a.Dev.X.Take(n))
        Dim xb As New List(Of Double)(b.Dev.X.Take(n))
        Dim xd As New List(Of Double)(d)

        Dim adevA As List(Of KeyValuePair(Of Integer, Double)) = ComputeAllanDeviation(xa, True)
        Dim adevB As List(Of KeyValuePair(Of Integer, Double)) = ComputeAllanDeviation(xb, True)
        Dim adevD As List(Of KeyValuePair(Of Integer, Double)) = ComputeAllanDeviation(xd, True)

        If adevD.Count = 0 Then Exit Sub

        Dim byTauA As New Dictionary(Of Integer, Double)
        Dim byTauB As New Dictionary(Of Integer, Double)
        For Each kvp As KeyValuePair(Of Integer, Double) In adevA
            byTauA(kvp.Key) = kvp.Value
        Next
        For Each kvp As KeyValuePair(Of Integer, Double) In adevB
            byTauB(kvp.Key) = kvp.Value
        Next

        Dim mps As Double = inp.MinsPerSample

        html.AppendLine("<h3>Stability of the difference (Allan deviation)</h3>")
        html.AppendLine("<p class=""foot"">How well the two meters agree at each averaging time, in ppm of Dev 1's reading. If their noise were independent the difference would be the two added in quadrature; a lower figure means they share noise that cancels out of the difference.</p>")
        html.AppendLine("<table class=""grid""><tr><th>tau (readings)</th><th>tau (time)</th><th>Dev 1 (ppm)</th><th>Dev 2 (ppm)</th><th>Difference (ppm)</th><th>If independent (ppm)</th></tr>")

        Dim lx As New List(Of Double)
        Dim ya As New List(Of Double)
        Dim yb As New List(Of Double)
        Dim yd As New List(Of Double)
        Dim bestTau As Integer = adevD(0).Key
        Dim bestPpm As Double = Double.MaxValue

        For Each kvp As KeyValuePair(Of Integer, Double) In adevD

            If Not byTauA.ContainsKey(kvp.Key) OrElse Not byTauB.ContainsKey(kvp.Key) Then Continue For

            Dim pa As Double = byTauA(kvp.Key) / meanA * 1000000.0
            Dim pb As Double = byTauB(kvp.Key) / meanB * 1000000.0
            Dim pd As Double = kvp.Value / meanA * 1000000.0
            Dim expected As Double = Math.Sqrt(pa * pa + pb * pb)

            If pd < bestPpm Then
                bestPpm = pd
                bestTau = kvp.Key
            End If

            html.AppendLine("<tr><td>" & kvp.Key.ToString() & "</td><td>" & If(mps > 0, RpDur(kvp.Key * mps), "n/a") & "</td><td>" & RpNum(pa, 4) & "</td><td>" & RpNum(pb, 4) &
                            "</td><td>" & RpNum(pd, 4) & "</td><td>" & RpNum(expected, 4) & "</td></tr>")

            If pa > 0 AndAlso pb > 0 AndAlso pd > 0 Then
                lx.Add(Math.Log10(kvp.Key))
                ya.Add(Math.Log10(pa))
                yb.Add(Math.Log10(pb))
                yd.Add(Math.Log10(pd))
            End If

        Next
        html.AppendLine("</table>")

        html.AppendLine(RpCompareAllanChart(lx, ya, yb, yd, a, b))

        Dim lastTau As Integer = adevD(adevD.Count - 1).Key
        Dim bestText As String = If(mps > 0, RpDur(bestTau * mps), bestTau.ToString() & " readings")

        If bestTau >= lastTau Then
            findings.Add(New ReportFinding("note", "Dev 1 - Dev 2: the two meters agree to " & RpPpmNum(bestPpm) & " at the longest averaging time (" & bestText & "), and the difference is still getting smaller."))
        Else
            findings.Add(New ReportFinding("note", "Dev 1 - Dev 2: the two meters agree best with averaging of about " & bestText & ", where the Allan deviation of their difference is " & RpPpmNum(bestPpm) &
                                           ". Beyond that, drift between them dominates."))
        End If

        ' At the best averaging time: how much of the noise is common to both meters?
        Dim bestA As Double = Double.NaN
        Dim bestB As Double = Double.NaN
        If byTauA.ContainsKey(bestTau) Then bestA = byTauA(bestTau) / meanA * 1000000.0
        If byTauB.ContainsKey(bestTau) Then bestB = byTauB(bestTau) / meanB * 1000000.0
        If Not Double.IsNaN(bestA) AndAlso Not Double.IsNaN(bestB) Then
            Dim independent As Double = Math.Sqrt(bestA * bestA + bestB * bestB)
            If independent > 0 AndAlso bestPpm / independent < 0.6 Then
                findings.Add(New ReportFinding("note", "Dev 1 - Dev 2: at that averaging time the difference (" & RpPpmNum(bestPpm) & ") is only " & RpNum(bestPpm / independent, 2) &
                             " x what independent meters would give (" & RpPpmNum(independent) & "), so most of what each meter sees at that timescale is common to both (the shared source or the environment) and cancels out of the difference."))
            End If
        End If

        ' Single readings: is the difference quieter than independent noise would give?
        If byTauA.ContainsKey(1) AndAlso byTauB.ContainsKey(1) Then
            Dim firstD As Double = Double.NaN
            For Each kvp As KeyValuePair(Of Integer, Double) In adevD
                If kvp.Key = 1 Then firstD = kvp.Value / meanA * 1000000.0
            Next
            Dim pa1 As Double = byTauA(1) / meanA * 1000000.0
            Dim pb1 As Double = byTauB(1) / meanB * 1000000.0
            Dim expected1 As Double = Math.Sqrt(pa1 * pa1 + pb1 * pb1)
            If Not Double.IsNaN(firstD) AndAlso expected1 > 0 Then
                Dim ratio As Double = firstD / expected1
                If ratio < 0.8 Then
                    findings.Add(New ReportFinding("note", "Dev 1 - Dev 2: for single readings the difference is " & RpNum(ratio, 2) & " x what independent noise would give, so part of the noise is common to both meters."))
                End If
            End If
        End If

    End Sub

    Private Function RpCompareAllanChart(lx As List(Of Double), ya As List(Of Double), yb As List(Of Double), yd As List(Of Double), a As ReportScope, b As ReportScope) As String

        Try
            If lx.Count < 2 Then Return ""

            Dim plot As New ScottPlot.Plot()

            Dim xs() As Double = lx.ToArray()

            Dim sa As ScottPlot.Plottables.Scatter = plot.Add.Scatter(xs, ya.ToArray())
            sa.Color = a.Colour
            sa.LineWidth = 1.5F
            sa.MarkerSize = 4
            sa.LegendText = "Dev 1"

            Dim sb As ScottPlot.Plottables.Scatter = plot.Add.Scatter(xs, yb.ToArray())
            sb.Color = b.Colour
            sb.LineWidth = 1.5F
            sb.MarkerSize = 4
            sb.LegendText = "Dev 2"

            Dim sd As ScottPlot.Plottables.Scatter = plot.Add.Scatter(xs, yd.ToArray())
            sd.Color = New ScottPlot.Color(Color.FromArgb(90, 60, 150))
            sd.LineWidth = 2.5F
            sd.MarkerSize = 5
            sd.LegendText = "Dev 1 - Dev 2"

            Dim yMin As Double = Math.Min(ya.Min(), Math.Min(yb.Min(), yd.Min()))
            Dim yMax As Double = Math.Max(ya.Max(), Math.Max(yb.Max(), yd.Max()))

            Dim xTicks As New ScottPlot.TickGenerators.NumericManual()
            For e As Integer = CInt(Math.Floor(xs.Min())) To CInt(Math.Ceiling(xs.Max()))
                xTicks.AddMajor(e, Math.Pow(10, e).ToString("G", CultureInfo.InvariantCulture))
            Next
            plot.Axes.Bottom.TickGenerator = xTicks

            Dim yTicks As New ScottPlot.TickGenerators.NumericManual()
            For e As Integer = CInt(Math.Floor(yMin)) To CInt(Math.Ceiling(yMax))
                yTicks.AddMajor(e, Math.Pow(10, e).ToString("G", CultureInfo.InvariantCulture))
            Next
            plot.Axes.Left.TickGenerator = yTicks

            plot.Title("Allan deviation: each meter and their difference", 14)
            plot.XLabel("Averaging time tau (readings)", 12)
            plot.YLabel("Allan deviation (ppm of reading)", 12)
            plot.ShowLegend()

            Return RpSvg(plot, 1000, 380)
        Catch
            Return ""
        End Try

    End Function

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Two devices together
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Sub RpCompareSection(inp As ReportInput, a As ReportScope, b As ReportScope, html As StringBuilder, findings As List(Of ReportFinding))

        Dim n As Integer = Math.Min(a.Count, b.Count)
        If n < 10 OrElse a.Dev.FirstIndex <> b.Dev.FirstIndex Then Exit Sub

        html.AppendLine("<section><h2>Dev 1 compared with Dev 2</h2>")
        html.AppendLine("<p class=""foot"">Readings are paired by position (reading k of Dev 1 with reading k of Dev 2).</p>")

        Dim d(n - 1) As Double
        Dim d1(n - 2) As Double
        Dim d2(n - 2) As Double
        For i As Integer = 0 To n - 1
            d(i) = a.Dev.X(i) - b.Dev.X(i)
        Next
        For i As Integer = 0 To n - 2
            d1(i) = a.Dev.X(i + 1) - a.Dev.X(i)
            d2(i) = b.Dev.X(i + 1) - b.Dev.X(i)
        Next

        Dim meanDiff As Double = d.Average()
        Dim sdDiff As Double = ExStdev(d)
        Dim noiseDiff As Double = RpPointNoise(d)
        Dim noiseA As Double = RpPointNoise(a.Dev.X)
        Dim noiseB As Double = RpPointNoise(b.Dev.X)
        Dim expected As Double = Math.Sqrt(noiseA * noiseA + noiseB * noiseB)
        Dim correlation As Double = RpCorrelation(d1, d2)

        RpTableStart(html, "Difference (Dev 1 - Dev 2)")
        RpRow(html, "Mean difference", RpNum(meanDiff, 8))
        RpRow(html, "STDEV of the difference", RpNum(sdDiff, 5))
        RpRow(html, "Point-to-point noise of the difference", RpNum(noiseDiff, 5), "if the two meters' noise were independent this would be about " & RpNum(expected, 5))
        RpRow(html, "Correlation of the readings' changes", RpNum(correlation, 3), "close to +1: both see the same disturbances (common-mode); close to 0: independent noise")

        ' Ratio of the two devices (useful when they measure the same quantity)
        Dim ratio(n - 1) As Double
        Dim ratioOk As Boolean = True
        For i As Integer = 0 To n - 1
            If b.Dev.X(i) = 0 Then
                ratioOk = False
                Exit For
            End If
            ratio(i) = a.Dev.X(i) / b.Dev.X(i)
        Next
        If ratioOk Then
            Dim meanRatio As Double = ratio.Average()
            RpRow(html, "Ratio Dev 1 / Dev 2", RpNum(meanRatio, 10), "mean over the scope")
            If inp.MinsPerSample > 0 AndAlso meanRatio <> 0 Then
                Dim tiRatio(n - 1) As Double
                Array.Copy(a.Ti, tiRatio, n)
                Dim rs, rq, re, rmt, rmr As Double
                If ExRegress(tiRatio, ratio, rs, rq, re, rmt, rmr) Then
                    RpRow(html, "Drift of the ratio", RpNum(rs * 60.0 / Math.Abs(meanRatio) * 1000000.0, 4) & " ppm/hour", "R2 = " & RpNum(rq, 3))
                End If
            End If
        End If

        If inp.MinsPerSample > 0 Then
            Dim tiArr(n - 1) As Double
            Array.Copy(a.Ti, tiArr, n)
            Dim slope, r2, slopeErr, mt, md As Double
            If ExRegress(tiArr, d, slope, r2, slopeErr, mt, md) Then
                Dim meanA As Double = Math.Abs(a.Mean)
                If meanA > 0 Then
                    RpRow(html, "Drift of the difference", RpNum(slope * 60.0 / meanA * 1000000.0, 4) & " ppm/hour", "of Dev 1's mean reading, R2 = " & RpNum(r2, 3))
                    findings.Add(New ReportFinding("note", "Dev 1 - Dev 2: mean difference " & RpNum(meanDiff, 6) & " (" & RpPpmNum(meanDiff / meanA * 1000000.0) & " of Dev 1), drifting at " & RpNum(slope * 60.0 / meanA * 1000000.0, 3) & " ppm/hour relative to each other."))
                End If
            End If
        End If
        RpTableEnd(html)

        Dim tiAll(n - 1) As Double
        Array.Copy(a.Ti, tiAll, n)
        html.AppendLine(RpLineChart(inp, "Dev 1 - Dev 2 over the run", tiAll, d, "Difference", New ScottPlot.Color(Color.FromArgb(90, 60, 150)), True))

        RpCompareStability(inp, a, b, d, html, findings)

        ' Spikes and level shifts at the same time in both devices point to something outside the meters
        If a.Spikes IsNot Nothing AndAlso b.Spikes IsNot Nothing Then
            Dim commonSpikes As Integer = 0
            For Each sa As Double() In a.Spikes
                For Each sb As Double() In b.Spikes
                    If Math.Abs(sa(0) - sb(0)) <= 2 Then
                        commonSpikes += 1
                        Exit For
                    End If
                Next
            Next

            Dim commonSteps As New List(Of Double())
            If a.Steps IsNot Nothing AndAlso b.Steps IsNot Nothing Then
                For Each stA As Double() In a.Steps
                    For Each stB As Double() In b.Steps
                        If Math.Abs(stA(0) - stB(0)) <= 30 Then
                            commonSteps.Add(stA)
                            Exit For
                        End If
                    Next
                Next
            End If

            If commonSpikes > 0 OrElse commonSteps.Count > 0 Then
                Dim ctxA As New ReportCtx With {.Mps = inp.MinsPerSample}
                Dim whereText As String = ""
                For i As Integer = 0 To Math.Min(2, commonSteps.Count - 1)
                    whereText &= If(whereText = "", "", ", ") & RpPosText(a, CInt(commonSteps(i)(0)), ctxA)
                Next
                Dim eventText As String = ""
                If commonSpikes > 0 Then eventText = commonSpikes.ToString() & " spike(s)"
                If commonSteps.Count > 0 Then eventText &= If(eventText = "", "", " and ") & commonSteps.Count.ToString() & " level shift(s) (at " & whereText & ")"
                html.AppendLine("<p><b>Events in common:</b> " & RpHtml(eventText) & " occur at the same time in both devices.</p>")
                findings.Add(New ReportFinding("warn", "Dev 1 and Dev 2 both show " & eventText & " at the same times: that points to something outside the meters (the shared source or reference, the mains, or the environment) rather than a fault in either meter."))
            End If
        End If

        If Not Double.IsNaN(correlation) AndAlso correlation > 0.5 Then
            findings.Add(New ReportFinding("note", "Dev 1 and Dev 2 move together (correlation " & RpNum(correlation, 2) & "): much of the noise is common to both, for example the reference or the environment."))
        End If

        html.AppendLine("</section>")

    End Sub

    ' ---------------------------------------------------------------------------------------------------------------------------
    '  Charts (SVG, drawn off-screen by ScottPlot and embedded in the page)
    ' ---------------------------------------------------------------------------------------------------------------------------

    Private Function RpSvg(plot As ScottPlot.Plot, width As Integer, height As Integer) As String

        Try
            Dim svg As String = plot.GetSvgHtml(width, height)

            ' ScottPlot writes a fixed width and height and no viewBox, so the chart could not shrink to fit a narrower window or a PDF page
            ' (it was cut off on the right). A viewBox lets it scale to the page.
            svg = New System.Text.RegularExpressions.Regex("<svg(?![^>]*viewBox)").Replace(svg, "<svg viewBox=""0 0 " & width.ToString() & " " & height.ToString() & """", 1)

            ' ScottPlot also positions every single character of every label (x="0, 7.4, 10.5, ..." with a matching comma-ended y). Some SVG
            ' viewers ignore those lists and pile the characters of each label on top of one another (garbled text). Keep just the first
            ' position of each label and let the viewer lay the characters out itself, and give the fonts fallbacks for machines without them.
            svg = New System.Text.RegularExpressions.Regex("\b([xy])=""(-?[0-9.]+(?:[eE][-+]?[0-9]+)?)\s*,[^""]*""").Replace(svg, "$1=""$2""")
            svg = New System.Text.RegularExpressions.Regex("font-family=""([^""]*)""").Replace(svg, "font-family=""$1, Segoe UI, Arial, Helvetica, sans-serif""")

            Return "<div class=""chart"">" & svg & "</div>"
        Catch
            Return ""
        End Try

    End Function

    ' Fixed-decimal tick labels sized to the span (the default drops decimals on values like 9.99999).
    Private Sub RpFixLeftTicks(plot As ScottPlot.Plot, span As Double)

        If span <= 0 Then Exit Sub

        Dim gen = TryCast(plot.Axes.Left.TickGenerator, ScottPlot.TickGenerators.NumericAutomatic)
        If gen IsNot Nothing Then
            Dim decimals As Integer = Math.Min(10, Math.Max(0, CInt(Math.Ceiling(-Math.Log10(span / 8))) + 1))
            Dim numberFormat As String = "F" & decimals.ToString()
            gen.LabelFormatter = Function(v As Double) v.ToString(numberFormat)
        End If

    End Sub

    Private Function RpTraceChart(inp As ReportInput, sc As ReportScope) As String

        Try
            Dim n As Integer = sc.Count
            Dim b As Integer = Math.Min(800, n)

            Dim xs(b - 1) As Double
            Dim lows(b - 1) As Double
            Dim highs(b - 1) As Double
            Dim means(b - 1) As Double
            Dim yLo As Double = Double.MaxValue
            Dim yHi As Double = Double.MinValue

            For k As Integer = 0 To b - 1
                Dim i0 As Integer = CInt((CLng(k) * n) \ b)
                Dim i1 As Integer = CInt((CLng(k + 1) * n) \ b) - 1
                If i1 < i0 Then i1 = i0

                Dim lo As Double = Double.MaxValue
                Dim hi As Double = Double.MinValue
                Dim sum As Double = 0.0
                For i As Integer = i0 To i1
                    Dim v As Double = sc.Dev.X(i)
                    If v < lo Then lo = v
                    If v > hi Then hi = v
                    sum += v
                Next

                xs(k) = sc.Ti((i0 + i1) \ 2)
                lows(k) = lo
                highs(k) = hi
                means(k) = sum / (i1 - i0 + 1)
                If lo < yLo Then yLo = lo
                If hi > yHi Then yHi = hi
            Next

            Dim plot As New ScottPlot.Plot()

            Dim envelope As ScottPlot.Plottables.FillY = plot.Add.FillY(xs, lows, highs)
            envelope.FillColor = sc.Colour.WithAlpha(0.35)
            envelope.LineColor = sc.Colour.WithAlpha(0.7)
            envelope.LineWidth = 1
            envelope.MarkerSize = 0

            Dim meanLine As ScottPlot.Plottables.Scatter = plot.Add.Scatter(xs, means)
            meanLine.Color = sc.Colour
            meanLine.LineWidth = 1
            meanLine.MarkerStyle.IsVisible = False

            plot.Title("Dev " & sc.Dev.Slot.ToString() & " - " & sc.Dev.Name & ": readings over the run (min/max envelope and mean)", 14)
            plot.XLabel(If(inp.MinsPerSample > 0, "Time (minutes)", "Reading number"), 12)
            plot.YLabel("Reading", 12)

            Dim pad As Double = (yHi - yLo) * 0.06
            If pad = 0 Then pad = If(yHi = 0, 0.001, Math.Abs(yHi) / 1000)
            plot.Axes.SetLimits(xs(0), xs(b - 1), yLo - pad, yHi + pad)
            RpFixLeftTicks(plot, yHi - yLo + 2 * pad)

            Return RpSvg(plot, 1000, 340)
        Catch
            Return ""
        End Try

    End Function

    ' A simple line chart (averaged into at most 1200 points) with sensible Y limits.
    Private Function RpLineChart(inp As ReportInput, title As String, xs() As Double, ys() As Double, yLabel As String, colour As ScottPlot.Color, fixTicks As Boolean) As String

        Try
            Dim n As Integer = Math.Min(xs.Length, ys.Length)
            If n < 2 Then Return ""

            Dim b As Integer = Math.Min(1200, n)
            Dim dx(b - 1) As Double
            Dim dy(b - 1) As Double
            Dim yLo As Double = Double.MaxValue
            Dim yHi As Double = Double.MinValue

            For k As Integer = 0 To b - 1
                Dim i0 As Integer = CInt((CLng(k) * n) \ b)
                Dim i1 As Integer = CInt((CLng(k + 1) * n) \ b) - 1
                If i1 < i0 Then i1 = i0
                Dim sum As Double = 0.0
                For i As Integer = i0 To i1
                    sum += ys(i)
                Next
                dx(k) = xs((i0 + i1) \ 2)
                dy(k) = sum / (i1 - i0 + 1)
                If dy(k) < yLo Then yLo = dy(k)
                If dy(k) > yHi Then yHi = dy(k)
            Next

            Dim plot As New ScottPlot.Plot()
            Dim line As ScottPlot.Plottables.Scatter = plot.Add.Scatter(dx, dy)
            line.Color = colour
            line.LineWidth = 1.5F
            line.MarkerStyle.IsVisible = False

            plot.Title(title, 14)
            plot.XLabel(If(inp.MinsPerSample > 0, "Time (minutes)", "Reading number"), 12)
            plot.YLabel(yLabel, 12)

            Dim pad As Double = (yHi - yLo) * 0.08
            If pad = 0 Then pad = If(yHi = 0, 0.001, Math.Abs(yHi) / 1000)
            plot.Axes.SetLimits(dx(0), dx(b - 1), yLo - pad, yHi + pad)
            If fixTicks Then RpFixLeftTicks(plot, yHi - yLo + 2 * pad)

            Return RpSvg(plot, 1000, 280)
        Catch
            Return ""
        End Try

    End Function

    Private Function RpHistogramChart(sc As ReportScope) As String

        Try
            Dim plot As New ScottPlot.Plot()
            Chart2BuildHistogram(plot, "Dev " & sc.Dev.Slot.ToString() & " - " & sc.Dev.Name & " (" & sc.Count.ToString() & " readings)", sc.Dev.X, sc.Colour, True)
            Return RpSvg(plot, 1000, 400)
        Catch
            Return ""
        End Try

    End Function

    Private Function RpTempcoChart(sc As ReportScope, ctx As ReportCtx) As String

        Try
            Dim n As Integer = sc.Count
            Dim stride As Integer = Math.Max(1, n \ 3000)
            Dim count As Integer = (n + stride - 1) \ stride

            Dim txs(count - 1) As Double
            Dim vys(count - 1) As Double
            For k As Integer = 0 To count - 1
                txs(k) = sc.Dev.T(k * stride)
                vys(k) = sc.Dev.X(k * stride)
            Next

            Dim plot As New ScottPlot.Plot()

            Dim points As ScottPlot.Plottables.Scatter = plot.Add.Scatter(txs, vys)
            points.LineWidth = 0
            points.MarkerSize = 3
            points.Color = sc.Colour.WithAlpha(0.45)

            Dim slope, r2, slopeErr, mt, mv As Double
            If ExRegress(sc.Dev.T, sc.Dev.X, slope, r2, slopeErr, mt, mv) Then
                Dim t0 As Double = sc.Dev.T.Min()
                Dim t1 As Double = sc.Dev.T.Max()
                Dim fit As ScottPlot.Plottables.Scatter = plot.Add.Scatter(New Double() {t0, t1}, New Double() {mv + slope * (t0 - mt), mv + slope * (t1 - mt)})
                fit.Color = New ScottPlot.Color(Color.FromArgb(200, 30, 30))
                fit.LineWidth = 2
                fit.MarkerStyle.IsVisible = False
            End If

            plot.Title("Dev " & sc.Dev.Slot.ToString() & ": reading against temperature (red = least-squares fit)", 14)
            plot.XLabel("Temperature (degC)", 12)
            plot.YLabel("Reading", 12)

            Dim span As Double = sc.Dev.X.Max() - sc.Dev.X.Min()
            RpFixLeftTicks(plot, span * 1.1)

            Return RpSvg(plot, 1000, 340)
        Catch
            Return ""
        End Try

    End Function

    ' Log-log Allan deviation curve with the 1/sqrt(tau) white-noise line through the first point.
    Private Function RpAllanChart(sc As ReportScope, ctx As ReportCtx, adev As List(Of KeyValuePair(Of Integer, Double))) As String

        Try
            Dim pts As New List(Of Double())
            For Each kvp As KeyValuePair(Of Integer, Double) In adev
                Dim p As Double = ctx.Ppm(kvp.Value)
                If p > 0 AndAlso kvp.Key > 0 Then pts.Add(New Double() {Math.Log10(kvp.Key), Math.Log10(p)})
            Next
            If pts.Count < 2 Then Return ""

            Dim lx(pts.Count - 1) As Double
            Dim ly(pts.Count - 1) As Double
            For i As Integer = 0 To pts.Count - 1
                lx(i) = pts(i)(0)
                ly(i) = pts(i)(1)
            Next

            Dim plot As New ScottPlot.Plot()

            Dim ideal As ScottPlot.Plottables.Scatter = plot.Add.Scatter(New Double() {lx(0), lx(lx.Length - 1)}, New Double() {ly(0), ly(0) - 0.5 * (lx(lx.Length - 1) - lx(0))})
            ideal.Color = New ScottPlot.Color(Color.FromArgb(120, 120, 120))
            ideal.LineWidth = 1.5F
            ideal.LinePattern = ScottPlot.LinePattern.Dotted
            ideal.MarkerStyle.IsVisible = False

            Dim curve As ScottPlot.Plottables.Scatter = plot.Add.Scatter(lx, ly)
            curve.Color = sc.Colour
            curve.LineWidth = 2
            curve.MarkerSize = 5

            Dim xTicks As New ScottPlot.TickGenerators.NumericManual()
            For e As Integer = CInt(Math.Floor(lx.Min())) To CInt(Math.Ceiling(lx.Max()))
                xTicks.AddMajor(e, Math.Pow(10, e).ToString("G", CultureInfo.InvariantCulture))
            Next
            plot.Axes.Bottom.TickGenerator = xTicks

            Dim yTicks As New ScottPlot.TickGenerators.NumericManual()
            For e As Integer = CInt(Math.Floor(Math.Min(ly.Min(), ly(0) - 0.5 * (lx(lx.Length - 1) - lx(0))))) To CInt(Math.Ceiling(ly.Max()))
                yTicks.AddMajor(e, Math.Pow(10, e).ToString("G", CultureInfo.InvariantCulture))
            Next
            plot.Axes.Left.TickGenerator = yTicks

            plot.Title("Dev " & sc.Dev.Slot.ToString() & ": Allan deviation (dotted = ideal white noise, falling as 1/sqrt(tau))", 14)
            plot.XLabel("Averaging time tau (readings)", 12)
            plot.YLabel("Allan deviation (ppm of mean)", 12)

            Return RpSvg(plot, 1000, 380)
        Catch
            Return ""
        End Try

    End Function

End Class
