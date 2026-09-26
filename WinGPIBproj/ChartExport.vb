Imports System.Globalization
Imports System.IO
Imports System.Text

' Playback Chart "Export Results": writes every statistic and analysis result for the visible Dev 1 / Dev 2 traces to a
' text file, over the Regional Stats band if it is showing, otherwise the whole run. The analysis checkboxes
' (Noise Band, Trend Line, ...) only control drawing, so all results are always written.
Partial Public Class Chart

    Private Const ExportPpmClamp As Double = 99.0

    Private Function ExNum(x As Double) As String

        If Double.IsNaN(x) OrElse Double.IsInfinity(x) Then Return "n/a"
        Return x.ToString("G10", CultureInfo.InvariantCulture)

    End Function

    Private Function ExStdev(a() As Double) As Double

        If a.Length < 2 Then Return 0.0
        Dim m As Double = a.Average()
        Dim s As Double = 0.0
        For Each v As Double In a
            s += (v - m) * (v - m)
        Next
        Return Math.Sqrt(s / (a.Length - 1))

    End Function

    ' Least-squares line through (xs, ys). False if there are fewer than 3 points or no spread in x.
    Private Function ExRegress(xs() As Double, ys() As Double, ByRef slope As Double, ByRef r2 As Double, ByRef slopeErr As Double,
                               ByRef meanX As Double, ByRef meanY As Double) As Boolean

        Dim n As Integer = xs.Length
        If n < 3 Then Return False

        meanX = xs.Average()
        meanY = ys.Average()

        Dim sxx As Double = 0.0
        Dim sxy As Double = 0.0
        Dim syy As Double = 0.0
        For i As Integer = 0 To n - 1
            Dim dx As Double = xs(i) - meanX
            Dim dy As Double = ys(i) - meanY
            sxx += dx * dx
            sxy += dx * dy
            syy += dy * dy
        Next

        If sxx <= 0 Then Return False

        slope = sxy / sxx
        r2 = If(syy > 0, (sxy * sxy) / (sxx * syy), 0.0)
        slopeErr = Math.Sqrt(Math.Max(0.0, syy - slope * sxy) / (n - 2) / sxx)
        Return True

    End Function

    Private Function ExCell(dr As DataRow, columnName As String) As Double

        Dim result As Double = Double.NaN
        If dataTable1.Columns.Contains(columnName) AndAlso Not dr.IsNull(columnName) Then
            Double.TryParse(Convert.ToString(dr(columnName)), NumberStyles.Float, CultureInfo.InvariantCulture, result)
        End If
        Return result

    End Function

    Private Function ExInstantPpmDegC(v As Double, t As Double, v0 As Double, t0 As Double) As Double

        If t - t0 = 0 Then Return 0.00000001
        Dim ppm As Double = ((v - v0) / (v0 * (t - t0))) * 1000000.0
        Return Math.Max(-ExportPpmClamp, Math.Min(ExportPpmClamp, ppm))

    End Function

    Private Sub ButtonExportResults_Click(sender As Object, e As EventArgs) Handles ButtonExportResults.Click

        If Not (ChartLoaded AndAlso CSVfileok) Then Exit Sub

        Dim regionOn As Boolean = Chart2RegionSpan IsNot Nothing AndAlso Chart2RegionSpan.IsVisible
        Dim regionLeft As Double = If(regionOn, Chart2RegionSpan.Left, 0.0)
        Dim regionRight As Double = If(regionOn, Chart2RegionSpan.Right, 0.0)

        ' Devices to include: those whose Data checkbox is ticked.
        Dim slots As New List(Of Integer)
        If DeviceName1.Text <> "" AndAlso Chart2Dev1Series IsNot Nothing AndAlso Chart2Dev1Series.IsVisible Then slots.Add(1)
        If DeviceName2.Text <> "" AndAlso Chart2Dev2Series IsNot Nothing AndAlso Chart2Dev2Series.IsVisible Then slots.Add(2)

        If slots.Count = 0 Then
            MessageBox.Show("Tick the Data checkbox of at least one device to include it in the export.", "Export Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim sb As New StringBuilder()
        Dim scopeText As String = "whole run"
        Dim fileSuffix As String = ""
        If regionOn Then
            Dim regionFirst As Integer = Math.Max(0, CInt(Math.Ceiling(regionLeft)))
            Dim regionLast As Integer = CInt(Math.Floor(regionRight))
            scopeText = "region " & regionFirst.ToString() & " - " & regionLast.ToString()
            fileSuffix = "_region_" & regionFirst.ToString() & "-" & regionLast.ToString()
        End If

        sb.AppendLine("WinGPIB Playback Chart - exported results")
        sb.AppendLine("Exported: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        sb.AppendLine("CSV file: " & filePlayback)
        sb.AppendLine("Scope: " & scopeText)
        sb.AppendLine()

        For Each slot As Integer In slots

            Dim name As String = If(slot = 1, DeviceName1.Text, DeviceName2.Text)
            Dim rows() As DataRow = dataTable1.Select("DEVICE ='" & name & "'")
            Dim n As Integer = rows.Length
            If n = 0 Then Continue For

            Dim allValues(n - 1) As Double
            Dim allTemps(n - 1) As Double
            For i As Integer = 0 To n - 1
                allValues(i) = Convert.ToDouble(rows(i)("VALUE"))
                allTemps(i) = Convert.ToDouble(rows(i)("TEMP"))
            Next

            Dim first As Integer = 0
            Dim last As Integer = n - 1
            If regionOn Then
                first = Math.Max(0, CInt(Math.Ceiling(regionLeft)))
                last = Math.Min(n - 1, CInt(Math.Floor(regionRight)))
            End If

            sb.AppendLine(New String("="c, 78))
            sb.AppendLine("DEVICE " & slot.ToString() & ": " & name)
            sb.AppendLine(New String("="c, 78))

            If last - first < 1 Then
                sb.AppendLine("Scope too small (" & scopeText & ") - at least 2 samples are needed.")
                sb.AppendLine()
                Continue For
            End If

            Dim count As Integer = last - first + 1
            Dim x(count - 1) As Double
            Dim t(count - 1) As Double
            Array.Copy(allValues, first, x, 0, count)
            Array.Copy(allTemps, first, t, 0, count)

            Dim mps As Double = Chart2MinsPerSample
            Dim meanX As Double = x.Average()
            Dim minX As Double = x.Min()
            Dim maxX As Double = x.Max()
            Dim avgBox As String = If(slot = 1, DEV1avg.Text, DEV2avg.Text)

            sb.AppendLine("Scope: " & If(regionOn, "region", "whole run") & "  (samples " & first.ToString() & " - " & last.ToString() & " on the chart's 0-based X axis, " & count.ToString() & " samples" &
                          If(mps > 0, ", " & ExNum((count - 1) * mps) & " mins", "") & ")")
            sb.AppendLine("From " & Convert.ToString(rows(first)("DATETIME")) & " to " & Convert.ToString(rows(last)("DATETIME")))
            sb.AppendLine("Avg box = " & avgBox & " (all figures below use the raw readings)   RMS window = " & RMSwindow.Text)
            sb.AppendLine()

            ' 1. Recorded statistics (columns in the CSV, at the last sample of the scope)
            Dim prefix As String = "DEV" & slot.ToString() & "_"
            Dim lastRow As DataRow = rows(last)
            sb.AppendLine("1. RECORDED STATISTICS (from the CSV, at the last sample of the scope)")
            sb.AppendLine("   n=" & ExNum(ExCell(lastRow, prefix & "SAMPLES")) & "  Mean=" & ExNum(ExCell(lastRow, prefix & "MEAN")) &
                          "  STDEV=" & ExNum(ExCell(lastRow, prefix & "STDEV")) & "  SEM=" & ExNum(ExCell(lastRow, prefix & "SEM")) &
                          "  Gain=" & ExNum(ExCell(lastRow, prefix & "GAIN")))
            sb.AppendLine("   MaxDiff=" & ExNum(ExCell(lastRow, prefix & "MAXDIFF")) & "  PPM Deviation=" & ExNum(ExCell(lastRow, prefix & "DEVIATION")))
            sb.AppendLine()

            ' 2. Data panel
            Dim noiseWindow As Integer = CInt(Val(RMSwindow.Text))
            If noiseWindow <= 0 Then noiseWindow = 10
            If noiseWindow > count Then noiseWindow = count
            Dim noiseSum As Double = 0.0
            For i As Integer = 0 To count - 1
                Dim s As Integer = Math.Max(0, i - noiseWindow + 1)
                Dim baseline As Double = 0.0
                For j As Integer = s To i
                    baseline += x(j)
                Next
                baseline /= (i - s + 1)
                noiseSum += (x(i) - baseline) * (x(i) - baseline)
            Next
            sb.AppendLine("2. DATA PANEL")
            sb.AppendLine("   Max=" & ExNum(maxX) & "  Min=" & ExNum(minX) & "  Max-Min=" & ExNum(maxX - minX))
            sb.AppendLine("   RMS Noise (drift removed by a trailing mean of " & noiseWindow.ToString() & ")=" & ExNum(Math.Sqrt(noiseSum / count)))

            Dim stmStart As Integer = Math.Max(0, last - ShortTermMeanWindow + 1)
            Dim stmSum As Double = 0.0
            For i As Integer = stmStart To last
                stmSum += allValues(i)
            Next
            sb.AppendLine("   Short Term Mean (last " & ShortTermMeanWindow.ToString() & " readings up to the last sample)=" & ExNum(stmSum / (last - stmStart + 1)))
            sb.AppendLine()

            ' 3. PPM baseline
            Dim ppmSlot As Integer = If(RadioButtonDev2.Checked, 2, 1)
            Dim v0 As Double = allValues(0)
            Dim t0 As Double = allTemps(0)
            Dim baselineSource As String = "first reading in the file"
            If slot = ppmSlot Then
                If Not CheckBoxMedianV.Checked Then
                    v0 = ParseInvariantDouble(MedianValue.Text)
                    baselineSource = "typed Initial Value"
                End If
                If Not CheckBoxMedianT.Checked Then t0 = ParseInvariantDouble(MedianTemp.Text)
            End If

            sb.AppendLine("3. PPM DEVIATION / PPM/DegC  (Initial Value " & ExNum(v0) & ", Initial Temp " & ExNum(t0) & ";  Initial Value = " & baselineSource & ")")

            If v0 = 0 Then
                sb.AppendLine("   Initial Value is zero - PPM figures not available.")
            Else
                Dim devMax As Double = Double.MinValue
                Dim devMin As Double = Double.MaxValue
                Dim ptMax As Double = Double.MinValue
                Dim ptMin As Double = Double.MaxValue
                For i As Integer = 0 To count - 1
                    Dim dev As Double = Math.Max(-ExportPpmClamp, Math.Min(ExportPpmClamp, (x(i) - v0) / v0 * 1000000.0))
                    devMax = Math.Max(devMax, dev)
                    devMin = Math.Min(devMin, dev)
                    Dim pt As Double = ExInstantPpmDegC(x(i), t(i), v0, t0)
                    ptMax = Math.Max(ptMax, pt)
                    ptMin = Math.Min(ptMin, pt)
                Next
                Dim lastDev As Double = Math.Max(-ExportPpmClamp, Math.Min(ExportPpmClamp, (x(count - 1) - v0) / v0 * 1000000.0))
                sb.AppendLine("   PPM Deviation (ppm):  last=" & ExNum(lastDev) & "  max=" & ExNum(devMax) & "  min=" & ExNum(devMin) & "   (clamped to +/-" & ExportPpmClamp.ToString("0") & ")")
                sb.AppendLine("   PPM/DegC (point):  last=" & ExNum(ExInstantPpmDegC(x(count - 1), t(count - 1), v0, t0)) & "  max=" & ExNum(ptMax) & "  min=" & ExNum(ptMin) &
                              "   (0.00000001 where Temp = Initial Temp)")

                ' Fit: least-squares slope of Value against Temp over the scope, divided by the Initial Value.
                Dim fSlope, fR2, fErr, fMeanT, fMeanV As Double
                If ExRegress(t, x, fSlope, fR2, fErr, fMeanT, fMeanV) Then
                    sb.AppendLine("   PPM/DegC (Fit):  " & ExNum(fSlope / v0 * 1000000.0) & "  +/- " & ExNum(Math.Abs(fErr / v0 * 1000000.0)))
                Else
                    sb.AppendLine("   PPM/DegC (Fit):  n/a (the temperature does not change enough to fit)")
                End If

                ' Trend: the same fit over the last RMS-window points up to the last sample.
                Dim rollWindow As Integer = Math.Max(2, CInt(Val(RMSwindow.Text)))
                Dim rs As Integer = Math.Max(0, last - rollWindow + 1)
                Dim rn As Integer = last - rs + 1
                Dim rt(rn - 1) As Double
                Dim rv(rn - 1) As Double
                Array.Copy(allTemps, rs, rt, 0, rn)
                Array.Copy(allValues, rs, rv, 0, rn)
                Dim rSlope As Double
                If rn >= 2 Then
                    Dim sT As Double = rt.Sum()
                    Dim sV As Double = rv.Sum()
                    Dim sTV As Double = 0.0
                    Dim sTT As Double = 0.0
                    For i As Integer = 0 To rn - 1
                        sTV += rt(i) * rv(i)
                        sTT += rt(i) * rt(i)
                    Next
                    Dim den As Double = rn * sTT - sT * sT
                    If den <> 0 Then
                        rSlope = (rn * sTV - sT * sV) / den
                        sb.AppendLine("   PPM/DegC (Trend, last " & rn.ToString() & " points):  " & ExNum(Math.Max(-ExportPpmClamp, Math.Min(ExportPpmClamp, rSlope / v0 * 1000000.0))))
                    Else
                        sb.AppendLine("   PPM/DegC (Trend, last " & rn.ToString() & " points):  0.00000001 (the temperature is constant over the window)")
                    End If
                End If
            End If
            sb.AppendLine()

            ' 4. Tempco curve: reading against temperature, normalised by the mean reading.
            sb.AppendLine("4. TEMPCO CURVE (Value against Temp over the scope, divided by the mean reading)")
            Dim tSlope, tR2, tErr, tMT, tMV As Double
            If ExRegress(t, x, tSlope, tR2, tErr, tMT, tMV) AndAlso tMV <> 0 Then
                sb.AppendLine("   tempco=" & ExNum(tSlope / Math.Abs(tMV) * 1000000.0) & " ppm/DegC  +/- " & ExNum(tErr / Math.Abs(tMV) * 1000000.0) &
                              "  R2=" & ExNum(tR2) & "  (temperature explains " & (tR2 * 100).ToString("0") & "% of the variation)")
            Else
                sb.AppendLine("   n/a (the temperature does not change enough to fit)")
            End If
            sb.AppendLine()

            ' 5. Trend line: reading against sample index.
            sb.AppendLine("5. TREND LINE (Value against sample index over the scope)")
            Dim idx(count - 1) As Double
            For i As Integer = 0 To count - 1
                idx(i) = first + i
            Next
            Dim lSlope, lR2, lErr, lMX, lMY As Double
            If ExRegress(idx, x, lSlope, lR2, lErr, lMX, lMY) Then
                Dim line As String = "   slope=" & ExNum(lSlope) & " per sample"
                If mps > 0 Then
                    Dim perHour As Double = lSlope / mps * 60.0
                    line &= "  drift=" & ExNum(perHour) & " per hour"
                    If lMY <> 0 Then line &= " (" & ExNum(perHour / Math.Abs(lMY) * 1000000.0) & " ppm/hour)"
                End If
                sb.AppendLine(line & "  R2=" & ExNum(lR2))
            Else
                sb.AppendLine("   n/a (needs at least 3 samples)")
            End If
            sb.AppendLine()

            ' 6. Noise band: centred rolling window, cut short at the ends of the scope.
            Dim bandWindow As Integer = CInt(Val(RMSwindow.Text))
            If bandWindow < 2 Then bandWindow = 100
            bandWindow = Math.Min(bandWindow, count)
            sb.AppendLine("6. NOISE BAND (rolling STDEV, window " & bandWindow.ToString() & ")")
            If count >= 3 Then
                Dim half As Integer = bandWindow \ 2
                Dim sdSum As Double = 0.0
                Dim sdMin As Double = Double.MaxValue
                Dim sdMax As Double = Double.MinValue
                Dim meanSum As Double = 0.0
                For i As Integer = 0 To count - 1
                    Dim a As Integer = Math.Max(0, i - half)
                    Dim b As Integer = Math.Min(count - 1, i + half)
                    Dim seg(b - a) As Double
                    Array.Copy(x, a, seg, 0, b - a + 1)
                    Dim sd As Double = ExStdev(seg)
                    sdSum += sd
                    meanSum += seg.Average()
                    sdMin = Math.Min(sdMin, sd)
                    sdMax = Math.Max(sdMax, sd)
                Next
                Dim typical As Double = sdSum / count
                Dim typicalMean As Double = meanSum / count
                sb.AppendLine("   typical STDEV=" & ExNum(typical) & If(typicalMean <> 0, " (" & ExNum(typical / Math.Abs(typicalMean) * 1000000.0) & " ppm)", "") &
                              "  quietest=" & ExNum(sdMin) & "  noisiest=" & ExNum(sdMax))
            Else
                sb.AppendLine("   n/a (needs at least 3 samples)")
            End If
            sb.AppendLine()

            ' 7. Region-style statistics of the scope.
            sb.AppendLine("7. STATISTICS OF THE SCOPE (as in the Regional Stats box)")
            Dim drift As Double = x(count - 1) - x(0)
            sb.AppendLine("   mean=" & ExNum(meanX) & "  STDEV=" & ExNum(ExStdev(x)) & "  min=" & ExNum(minX) & "  max=" & ExNum(maxX) & "  p-p=" & ExNum(maxX - minX))
            sb.AppendLine("   drift (last - first reading)=" & ExNum(drift) & If(meanX <> 0, " (" & ExNum(drift / Math.Abs(meanX) * 1000000.0) & " ppm)", ""))
            sb.AppendLine()

            ' 8. Histogram statistics.
            sb.AppendLine("8. HISTOGRAM OF READINGS (statistics)")
            Dim m2 As Double = 0.0
            Dim m3 As Double = 0.0
            Dim m4 As Double = 0.0
            For Each v As Double In x
                Dim d As Double = v - meanX
                m2 += d * d
                m3 += d * d * d
                m4 += d * d * d * d
            Next
            m2 /= count
            m3 /= count
            m4 /= count
            Dim sorted() As Double = DirectCast(x.Clone(), Double())
            Array.Sort(sorted)
            Dim distinctCount As Integer = 1
            Dim minGap As Double = Double.MaxValue
            For i As Integer = 1 To count - 1
                If sorted(i) <> sorted(i - 1) Then
                    distinctCount += 1
                    minGap = Math.Min(minGap, sorted(i) - sorted(i - 1))
                End If
            Next
            sb.AppendLine("   n=" & count.ToString() & "  skew=" & If(m2 > 0, ExNum(m3 / Math.Pow(m2, 1.5)), "n/a") & "  excess kurtosis=" & If(m2 > 0, ExNum(m4 / (m2 * m2) - 3), "n/a") &
                          "  distinct values=" & distinctCount.ToString() & If(distinctCount > 1, "  smallest step=" & ExNum(minGap), ""))
            sb.AppendLine()

            ' 9. Allan deviation and MDEV (ppm of the scope's mean reading).
            sb.AppendLine("9. ALLAN DEVIATION / MDEV (ppm of the mean reading; tau in samples)")
            If count < 4 OrElse meanX = 0 Then
                sb.AppendLine("   n/a (needs at least 4 samples)")
            Else
                Dim list As New List(Of Double)(x)
                Dim table As New SortedDictionary(Of Integer, Double())
                For Each kvp As KeyValuePair(Of Integer, Double) In ComputeAllanDeviation(list, False)
                    table(kvp.Key) = New Double() {kvp.Value, Double.NaN, Double.NaN}
                Next
                For Each kvp As KeyValuePair(Of Integer, Double) In ComputeAllanDeviation(list, True)
                    If Not table.ContainsKey(kvp.Key) Then table(kvp.Key) = New Double() {Double.NaN, Double.NaN, Double.NaN}
                    table(kvp.Key)(1) = kvp.Value
                Next
                For Each kvp As KeyValuePair(Of Integer, Double) In ComputeModifiedAllanDeviation(list)
                    If Not table.ContainsKey(kvp.Key) Then table(kvp.Key) = New Double() {Double.NaN, Double.NaN, Double.NaN}
                    table(kvp.Key)(2) = kvp.Value
                Next
                sb.AppendLine("   tau    ADEV (non-overlapping)   ADEV (overlapping)   MDEV")
                For Each kvp As KeyValuePair(Of Integer, Double()) In table
                    sb.AppendLine("   " & kvp.Key.ToString().PadRight(6) &
                                  ExNumPpm(kvp.Value(0), meanX).PadRight(25) & ExNumPpm(kvp.Value(1), meanX).PadRight(21) & ExNumPpm(kvp.Value(2), meanX))
                Next
            End If
            sb.AppendLine()

        Next

        ' Save
        Dim baseName As String = Path.GetFileNameWithoutExtension(filePlayback)
        Dim folder As String = ""
        Try
            folder = Path.GetDirectoryName(filePlayback)
        Catch
        End Try

        Using dlg As New SaveFileDialog With {
            .Title = "Export Results",
            .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
            .FileName = baseName & "_Results" & fileSuffix & ".txt",
            .OverwritePrompt = True
        }
            If folder <> "" AndAlso Directory.Exists(folder) Then dlg.InitialDirectory = folder
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Exit Sub

            Try
                File.WriteAllText(dlg.FileName, sb.ToString(), New UTF8Encoding(False))
            Catch ex As Exception
                MessageBox.Show("The results could not be saved: " & ex.Message, "Export Results", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using

    End Sub

    Private Function ExNumPpm(sigma As Double, meanValue As Double) As String

        If Double.IsNaN(sigma) Then Return "-"
        Return ExNum(sigma / meanValue * 1000000.0)

    End Function

End Class
