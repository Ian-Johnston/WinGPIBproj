Imports IODevices
Imports System.Globalization
Imports System.Text

Partial Class Formtest

    Private Cal72Table As New DataTable()
    Private Cal72CsvFile As String = ""

    Private Const Cal72Command As String = "CAL? 72"
    Private Const Cal72TempCommand As String = "TEMP?"

    Private Cal72Initialised As Boolean = False

    Private Sub InitCal72DriftTab()

        'Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift.csv"
        SelectCal72CsvFile()

        Cal72Table.Columns.Clear()

        Cal72Table.Columns.Add("Day", GetType(Double))
        Cal72Table.Columns.Add("Date", GetType(String))
        Cal72Table.Columns.Add("Time", GetType(String))
        Cal72Table.Columns.Add("Temp", GetType(Double))
        Cal72Table.Columns.Add("CAL? 72", GetType(Double))
        Cal72Table.Columns.Add("Days From Day 1", GetType(Double))
        Cal72Table.Columns.Add("Days From Last", GetType(Double))
        Cal72Table.Columns.Add("Drift ppm Day 1", GetType(Double))
        Cal72Table.Columns.Add("Avg ppm/day", GetType(Double))
        Cal72Table.Columns.Add("Drift ppm Last", GetType(Double))
        Cal72Table.Columns.Add("Notes", GetType(String))

        DataGridViewCal72.DataSource = Cal72Table

        ' mod headings
        DataGridViewCal72.Columns("Temp").HeaderText = "Temp?"
        DataGridViewCal72.Columns("Drift ppm Day 1").HeaderText = "Drift ppm Ref. Day 1"
        DataGridViewCal72.Columns("Drift ppm Last").HeaderText = "Drift ppm Ref. Last"
        DataGridViewCal72.Columns("Avg ppm/day").HeaderText = "Drift Avg ppm/Day"

        DataGridViewCal72.RowHeadersVisible = False
        DataGridViewCal72.ColumnHeadersVisible = True
        DataGridViewCal72.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCal72.RowTemplate.Height = 20

        DataGridViewCal72.DefaultCellStyle.Font = New Font("Segoe UI", 9)
        DataGridViewCal72.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)

        DataGridViewCal72.AllowUserToAddRows = False
        DataGridViewCal72.AllowUserToDeleteRows = False
        DataGridViewCal72.MultiSelect = False
        DataGridViewCal72.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridViewCal72.ReadOnly = False

        DataGridViewCal72.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        DataGridViewCal72.Columns("Day").Width = 45
        DataGridViewCal72.Columns("Date").Width = 75
        DataGridViewCal72.Columns("Time").Width = 50
        DataGridViewCal72.Columns("Temp").Width = 50
        DataGridViewCal72.Columns("CAL? 72").Width = 90
        DataGridViewCal72.Columns("Days From Day 1").Width = 50
        DataGridViewCal72.Columns("Days From Last").Width = 50
        DataGridViewCal72.Columns("Drift ppm Day 1").Width = 85
        DataGridViewCal72.Columns("Avg ppm/day").Width = 80
        DataGridViewCal72.Columns("Drift ppm Last").Width = 88
        DataGridViewCal72.Columns("Notes").Width = 364

        DataGridViewCal72.ScrollBars = ScrollBars.Vertical

        LoadCal72Csv()
        RecalculateCal72Table()
        SetCal72ColumnReadOnly()

        LabelCal72Status.Text = "READY"

        BeginInvoke(New MethodInvoker(AddressOf ScrollCal72GridToBottom))

        Cal72Initialised = True

    End Sub


    Private Sub RadioButton3458_CheckedChanged(sender As Object, e As EventArgs) _
Handles RadioButton34581.CheckedChanged,
        RadioButton34582.CheckedChanged,
        RadioButton34583.CheckedChanged,
        RadioButton34584.CheckedChanged,
        RadioButton34585.CheckedChanged

        If Cal72Initialised = False Then Exit Sub

        If CType(sender, RadioButton).Checked = False Then Exit Sub

        Try

            If Cal72CsvFile <> "" Then

                RecalculateCal72Table()
                SaveCal72Csv()

            End If

            SelectCal72CsvFile()
            LoadCal72Csv()
            RecalculateCal72Table()
            SetCal72ColumnReadOnly()
            ScrollCal72GridToBottom()

            LabelCal72Status.Text =
            "LOADED " & IO.Path.GetFileName(Cal72CsvFile)

        Catch ex As Exception

            MessageBox.Show(ex.Message,
                        "CAL72 File Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

        End Try

    End Sub


    Private Sub SelectCal72CsvFile()

        If RadioButton34581.Checked = True Then
            Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift.csv"

        ElseIf RadioButton34582.Checked = True Then
            Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift2.csv"

        ElseIf RadioButton34583.Checked = True Then
            Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift3.csv"

        ElseIf RadioButton34584.Checked = True Then
            Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift4.csv"

        ElseIf RadioButton34585.Checked = True Then
            Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift5.csv"

        Else
            RadioButton34581.Checked = True
            Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift.csv"
        End If

    End Sub


    Private Sub ScrollCal72GridToBottom()

        If DataGridViewCal72.Rows.Count = 0 Then Exit Sub

        DataGridViewCal72.ClearSelection()

        Dim lastRow As Integer = DataGridViewCal72.Rows.Count - 1

        DataGridViewCal72.FirstDisplayedScrollingRowIndex = lastRow
        DataGridViewCal72.Rows(lastRow).Selected = True
        DataGridViewCal72.CurrentCell = DataGridViewCal72.Rows(lastRow).Cells("Day")

    End Sub


    Private Sub SetCal72ColumnReadOnly()

        For Each col As DataGridViewColumn In DataGridViewCal72.Columns

            col.ReadOnly = True

            col.DefaultCellStyle.BackColor = Color.Gainsboro

        Next

        ' Editable columns
        DataGridViewCal72.Columns("Day").ReadOnly = False
        DataGridViewCal72.Columns("Date").ReadOnly = False
        DataGridViewCal72.Columns("Time").ReadOnly = False
        DataGridViewCal72.Columns("Temp").ReadOnly = False
        DataGridViewCal72.Columns("CAL? 72").ReadOnly = False
        DataGridViewCal72.Columns("Notes").ReadOnly = False

        ' Editable column colours
        DataGridViewCal72.Columns("Day").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("Date").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("Time").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("Temp").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("CAL? 72").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("Notes").DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub ButtonCal72Read_Click(sender As Object, e As EventArgs) Handles ButtonCal72Read.Click

        If ButtonDev1Run.Enabled = False Then
            MessageBox.Show("Device 1 is not started", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        LabelCal72Status.Text = "READING CAL? 72"
        Me.Refresh()

        TextBoxCal72Value.Text = Query3458AValue(Cal72Command)

        LabelCal72Status.Text = "READING TEMP"
        Me.Refresh()

        TextBoxCal72Temp.Text = Query3458AValue(Cal72TempCommand)

        LabelCal72Status.Text = "READ COMPLETE"

    End Sub

    Private Function Query3458AValue(command As String) As String

        txtr1a.Text = ""

        Dim q As IOQuery = Nothing
        dev1.QueryBlocking(command, q, False)
        Cbdev1(q)

        Return txtr1a.Text.Trim()

    End Function

    Private Sub ButtonCal72Add_Click(sender As Object, e As EventArgs) Handles ButtonCal72Add.Click

        Dim cal72Value As Double
        Dim tempValue As Double

        If Not Double.TryParse(TextBoxCal72Value.Text, NumberStyles.Float, CultureInfo.InvariantCulture, cal72Value) Then
            MessageBox.Show("Invalid CAL? 72 value.", "CAL? 72", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Not Double.TryParse(TextBoxCal72Temp.Text, NumberStyles.Float, CultureInfo.InvariantCulture, tempValue) Then
            MessageBox.Show("Invalid temperature value.", "CAL? 72", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim r As DataRow = Cal72Table.NewRow()

        r("Day") = GetSuggestedNextCal72Day()
        r("Date") = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        r("Time") = DateTime.Now.ToString("HH:mm", CultureInfo.InvariantCulture)
        r("Temp") = tempValue
        r("CAL? 72") = cal72Value
        r("Notes") = TextBoxCal72Notes.Text

        Cal72Table.Rows.Add(r)

        RecalculateCal72Table()
        SaveCal72Csv()

        ScrollCal72GridToBottom()

        LabelCal72Status.Text = "ENTRY ADDED"

    End Sub

    Private Function GetSuggestedNextCal72Day() As Double

        If Cal72Table.Rows.Count = 0 Then Return 1

        Dim firstDate As DateTime
        Dim firstDay As Double

        If Not GetRowDateTime(Cal72Table.Rows(0), firstDate) Then Return Cal72Table.Rows.Count + 1

        If Not Double.TryParse(Cal72Table.Rows(0)("Day").ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, firstDay) Then
            firstDay = 1
        End If

        Dim elapsedDays As Double = (DateTime.Now - firstDate).TotalDays

        Return Math.Round(firstDay + elapsedDays, 0)

    End Function

    Private Sub ButtonCal72Delete_Click(sender As Object, e As EventArgs) Handles ButtonCal72Delete.Click

        If DataGridViewCal72.SelectedRows.Count = 0 Then Exit Sub

        If MessageBox.Show("Delete selected CAL72 entry?", "CAL? 72", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Exit Sub
        End If

        DataGridViewCal72.Rows.Remove(DataGridViewCal72.SelectedRows(0))

        RecalculateCal72Table()
        SaveCal72Csv()

        LabelCal72Status.Text = "ENTRY DELETED"

    End Sub

    Private Sub ButtonCal72Save_Click(sender As Object, e As EventArgs) Handles ButtonCal72Save.Click

        RecalculateCal72Table()
        SaveCal72Csv()

        LabelCal72Status.Text = "SAVED"

    End Sub

    Private Sub DataGridViewCal72_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewCal72.CellEndEdit

        RecalculateCal72Table()
        SaveCal72Csv()

        LabelCal72Status.Text = "EDITED AND RECALCULATED"

    End Sub

    Private Sub RecalculateCal72Table()

        If Cal72Table.Rows.Count = 0 Then Exit Sub

        Dim day1Cal72 As Double
        Dim firstDay As Double

        If Not Double.TryParse(Cal72Table.Rows(0)("CAL? 72").ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, day1Cal72) Then Exit Sub
        If Not Double.TryParse(Cal72Table.Rows(0)("Day").ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, firstDay) Then firstDay = 1

        If day1Cal72 = 0 Then Exit Sub

        Dim previousDay As Double = firstDay
        Dim previousCal72 As Double = day1Cal72

        For i As Integer = 0 To Cal72Table.Rows.Count - 1

            Dim r As DataRow = Cal72Table.Rows(i)

            Dim thisDay As Double
            Dim thisCal72 As Double

            If Not Double.TryParse(r("Day").ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, thisDay) Then Continue For
            If Not Double.TryParse(r("CAL? 72").ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, thisCal72) Then Continue For

            Dim daysFromDay1 As Double = thisDay - firstDay
            Dim daysFromLast As Double = If(i = 0, 0, thisDay - previousDay)

            Dim driftPpmDay1 As Double = -(((day1Cal72 - thisCal72) * 1000000.0R) / day1Cal72)

            Dim avgPpmPerDay As Double = 0
            If thisDay <> 0 Then
                avgPpmPerDay = Math.Abs(driftPpmDay1 / thisDay)
            End If

            Dim driftPpmLast As Double = 0
            If i > 0 AndAlso previousCal72 <> 0 AndAlso daysFromLast <> 0 Then
                driftPpmLast = ((thisCal72 - previousCal72) * 1000000.0R) / (previousCal72 * daysFromLast)
            End If

            r("Days From Day 1") = Math.Round(daysFromDay1, 4)
            r("Days From Last") = Math.Round(daysFromLast, 4)
            r("Drift ppm Day 1") = Math.Round(driftPpmDay1, 6)
            r("Avg ppm/day") = Math.Round(avgPpmPerDay, 6)
            r("Drift ppm Last") = Math.Round(driftPpmLast, 6)

            previousDay = thisDay
            previousCal72 = thisCal72

        Next

    End Sub

    Private Function GetRowDateTime(r As DataRow, ByRef dt As DateTime) As Boolean

        Dim dateText As String = r("Date").ToString().Trim()
        Dim timeText As String = r("Time").ToString().Trim()

        Return DateTime.TryParseExact(dateText & " " & timeText,
                                      "yyyy-MM-dd HH:mm",
                                      CultureInfo.InvariantCulture,
                                      DateTimeStyles.None,
                                      dt)

    End Function

    Private Sub SaveCal72Csv()

        Try

            If Not IO.File.Exists(Cal72CsvFile) Then
                IO.File.Create(Cal72CsvFile).Dispose()
            End If

            IO.File.WriteAllText(Cal72CsvFile, "")

            IO.File.AppendAllText(Cal72CsvFile, "Day,Date,Time,Temp,CAL? 72,Days From Day 1,Days From Last,Drift ppm Day 1,Avg ppm/day,Drift ppm Last,Notes" & Environment.NewLine)

            For Each r As DataRow In Cal72Table.Rows

                Dim lineOut As String =
                    Csv(r("Day").ToString()) & "," &
                    Csv(r("Date").ToString()) & "," &
                    Csv(r("Time").ToString()) & "," &
                    Csv(Convert.ToDouble(r("Temp")).ToString("G17", CultureInfo.InvariantCulture)) & "," &
                    Csv(Convert.ToDouble(r("CAL? 72")).ToString("G17", CultureInfo.InvariantCulture)) & "," &
                    Csv(r("Days From Day 1").ToString()) & "," &
                    Csv(r("Days From Last").ToString()) & "," &
                    Csv(r("Drift ppm Day 1").ToString()) & "," &
                    Csv(r("Avg ppm/day").ToString()) & "," &
                    Csv(r("Drift ppm Last").ToString()) & "," &
                    Csv(r("Notes").ToString())

                IO.File.AppendAllText(Cal72CsvFile, lineOut & Environment.NewLine)

            Next

        Catch ex As Exception
            MessageBox.Show("Unable to save CAL72 file: " & ex.Message, "CAL72 File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub LoadCal72Csv()

        Cal72Table.Rows.Clear()

        Try

            If Not IO.File.Exists(Cal72CsvFile) Then
                IO.File.Create(Cal72CsvFile).Dispose()
                Exit Sub
            End If

            Dim lines() As String = IO.File.ReadAllLines(Cal72CsvFile)

            For i As Integer = 1 To lines.Length - 1

                If String.IsNullOrWhiteSpace(lines(i)) Then Continue For

                Dim p As List(Of String) = ParseCsvLine(lines(i))

                If p.Count < 11 Then Continue For

                Dim r As DataRow = Cal72Table.NewRow()

                r("Day") = ToDoubleSafe(p(0))
                r("Date") = p(1)
                r("Time") = p(2)
                r("Temp") = ToDoubleSafe(p(3))
                r("CAL? 72") = ToDoubleSafe(p(4))
                r("Days From Day 1") = ToDoubleSafe(p(5))
                r("Days From Last") = ToDoubleSafe(p(6))
                r("Drift ppm Day 1") = ToDoubleSafe(p(7))
                r("Avg ppm/day") = ToDoubleSafe(p(8))
                r("Drift ppm Last") = ToDoubleSafe(p(9))
                r("Notes") = p(10)

                Cal72Table.Rows.Add(r)

            Next

        Catch ex As Exception
            MessageBox.Show("Unable to load CAL72 file: " & ex.Message, "CAL72 File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Function ToDoubleSafe(s As String) As Double

        Dim v As Double

        If Double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, v) Then
            Return v
        End If

        Return 0

    End Function

    Private Function Csv(s As String) As String

        If s Is Nothing Then s = ""

        s = s.Replace("""", """""")

        If s.Contains(",") OrElse s.Contains("""") OrElse s.Contains(vbCr) OrElse s.Contains(vbLf) Then
            Return """" & s & """"
        End If

        Return s

    End Function

    Private Function ParseCsvLine(line As String) As List(Of String)

        Dim result As New List(Of String)
        Dim sb As New StringBuilder()
        Dim inQuotes As Boolean = False

        For i As Integer = 0 To line.Length - 1

            Dim ch As Char = line(i)

            If ch = """"c Then

                If inQuotes AndAlso i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                    sb.Append(""""c)
                    i += 1
                Else
                    inQuotes = Not inQuotes
                End If

            ElseIf ch = ","c AndAlso Not inQuotes Then

                result.Add(sb.ToString())
                sb.Clear()

            Else

                sb.Append(ch)

            End If

        Next

        result.Add(sb.ToString())

        Return result

    End Function

End Class