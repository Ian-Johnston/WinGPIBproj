Imports IODevices
Imports System.Globalization
Imports System.Text

Partial Class Formtest

    Private Cal72Table As New DataTable()
    Private Cal72CsvFile As String = ""

    Private Const Cal72Command As String = "CAL? 72"
    Private Const Cal72TempCommand As String = "TEMP?"

    Private Cal72Initialised As Boolean = False

    Private Const Cal72Cal11Command As String = "CAL? 1,1"
    Private Const Cal72Cal21Command As String = "CAL? 2,1"

    Private Sub InitCal72DriftTab()

        'Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift.csv"
        SelectCal72CsvFile()

        Cal72Table.Columns.Clear()

        Cal72Table.Columns.Add("Day", GetType(Double))
        Cal72Table.Columns.Add("Date", GetType(String))
        Cal72Table.Columns.Add("Time", GetType(String))
        Cal72Table.Columns.Add("Temp", GetType(Double))
        Cal72Table.Columns.Add("CAL? 72", GetType(Double))

        Cal72Table.Columns.Add("CAL? 1,1 40k", GetType(String))
        Cal72Table.Columns.Add("CAL? 2,1 Vref", GetType(String))

        Cal72Table.Columns.Add("CAL? 1,1 Dev", GetType(Double))
        Cal72Table.Columns.Add("CAL? 2,1 Dev", GetType(Double))

        Cal72Table.Columns.Add("Days From Day 1", GetType(Double))
        Cal72Table.Columns.Add("Days From Last", GetType(Double))
        Cal72Table.Columns.Add("Drift ppm Day 1", GetType(Double))
        Cal72Table.Columns.Add("Avg ppm/day", GetType(Double))
        Cal72Table.Columns.Add("Drift ppm Last", GetType(Double))
        Cal72Table.Columns.Add("Notes", GetType(String))

        DataGridViewCal72.DataSource = Cal72Table

        ' mod headings
        DataGridViewCal72.Columns("Drift ppm Day 1").HeaderText = "Drift ppm Ref. Day 1"
        DataGridViewCal72.Columns("Drift ppm Last").HeaderText = "Drift ppm Ref. Last"
        DataGridViewCal72.Columns("Avg ppm/day").HeaderText = "Drift Avg ppm/Day"
        DataGridViewCal72.Columns("CAL? 1,1 Dev").HeaderText = "CAL? 1,1 40k" & vbCrLf & "Deviation"
        DataGridViewCal72.Columns("CAL? 2,1 Dev").HeaderText = "CAL? 2,1 Vref" & vbCrLf & "Deviation"
        DataGridViewCal72.Columns("Days From Day 1").HeaderText = "Days Since Day 1"
        DataGridViewCal72.Columns("Days From Last").HeaderText = "Days Since Last"
        DataGridViewCal72.Columns("CAL? 1,1 40k").HeaderText = "CAL? 1,1" & vbCrLf & "40k"
        DataGridViewCal72.Columns("CAL? 2,1 Vref").HeaderText = "CAL? 2,1" & vbCrLf & "Vref"
        DataGridViewCal72.Columns("Temp").HeaderText = "3458A" & vbCrLf & "Temp?"
        DataGridViewCal72.Columns("CAL? 72").HeaderText = "CAL? 72" & vbCrLf & "Value"

        DataGridViewCal72.RowHeadersVisible = False
        DataGridViewCal72.ColumnHeadersVisible = True
        DataGridViewCal72.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCal72.RowTemplate.Height = 18

        DataGridViewCal72.DefaultCellStyle.Font = New Font("Segoe UI", 9)
        DataGridViewCal72.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)

        DataGridViewCal72.AllowUserToAddRows = False
        DataGridViewCal72.AllowUserToDeleteRows = False
        DataGridViewCal72.MultiSelect = False
        DataGridViewCal72.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridViewCal72.ReadOnly = False

        DataGridViewCal72.AllowUserToResizeRows = False
        DataGridViewCal72.AllowUserToResizeColumns = False

        DataGridViewCal72.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        DataGridViewCal72.Columns("Day").Width = 50
        DataGridViewCal72.Columns("Date").Width = 75
        DataGridViewCal72.Columns("Time").Width = 50
        DataGridViewCal72.Columns("Temp").Width = 55

        DataGridViewCal72.Columns("CAL? 72").Width = 80
        DataGridViewCal72.Columns("CAL? 1,1 40k").Width = 100
        DataGridViewCal72.Columns("CAL? 2,1 Vref").Width = 100

        DataGridViewCal72.Columns("Days From Day 1").Width = 80
        DataGridViewCal72.Columns("Days From Last").Width = 80

        DataGridViewCal72.Columns("Drift ppm Day 1").Width = 85
        DataGridViewCal72.Columns("Avg ppm/day").Width = 80
        DataGridViewCal72.Columns("Drift ppm Last").Width = 88

        DataGridViewCal72.Columns("Notes").Width = 600

        DataGridViewCal72.Columns("CAL? 1,1 Dev").Width = 90
        DataGridViewCal72.Columns("CAL? 2,1 Dev").Width = 90

        DataGridViewCal72.Columns("CAL? 1,1 Dev").DefaultCellStyle.Format = "0.0000000000"
        DataGridViewCal72.Columns("CAL? 2,1 Dev").DefaultCellStyle.Format = "0.0000000000"

        For Each col As DataGridViewColumn In DataGridViewCal72.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        DataGridViewCal72.ScrollBars = ScrollBars.Both

        LoadCal72Csv()
        RecalculateCal72Table()
        UpdateCal72SummaryPanel()
        SetCal72ColumnReadOnly()

        LabelCal72Status.Text = "READY"

        'BeginInvoke(New MethodInvoker(AddressOf ScrollCal72GridToBottom))

        BeginInvoke(New MethodInvoker(AddressOf ScrollCal72GridToBottom))
        BeginInvoke(New MethodInvoker(AddressOf ClearCal72Selection))

        'DataGridViewCal72.ClearSelection()

        Cal72Initialised = True

    End Sub


    Private Sub DataGridViewCal72_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridViewCal72.CellFormatting

        If e.RowIndex <> 0 Then Exit Sub

        Dim colName As String =
        DataGridViewCal72.Columns(e.ColumnIndex).Name

        Select Case colName

            Case "CAL? 1,1 Dev",
             "CAL? 2,1 Dev",
             "Days From Day 1",
             "Days From Last",
             "Drift ppm Day 1",
             "Avg ppm/day",
             "Drift ppm Last"

                e.Value = "-"
                e.FormattingApplied = True

        End Select

    End Sub


    Private Sub ClearCal72Selection()

        DataGridViewCal72.ClearSelection()
        DataGridViewCal72.CurrentCell = Nothing

    End Sub


    Private Sub DataGridViewCal72_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles DataGridViewCal72.DataBindingComplete

        DataGridViewCal72.ClearSelection()
        DataGridViewCal72.CurrentCell = Nothing

    End Sub


    Private Sub RadioButton3458_CheckedChanged(sender As Object, e As EventArgs) _
Handles RadioButton34581.CheckedChanged,
        RadioButton34582.CheckedChanged,
        RadioButton34583.CheckedChanged,
        RadioButton34584.CheckedChanged,
        RadioButton34585.CheckedChanged,
        RadioButton34586.CheckedChanged,
        RadioButton34587.CheckedChanged,
        RadioButton34588.CheckedChanged

        If Cal72Initialised = False Then Exit Sub

        If CType(sender, RadioButton).Checked = False Then Exit Sub

        Try

            If Cal72CsvFile <> "" Then

                RecalculateCal72Table()
                UpdateCal72SummaryPanel()
                SaveCal72Csv()

            End If

            SelectCal72CsvFile()
            LoadCal72Csv()
            RecalculateCal72Table()
            UpdateCal72SummaryPanel()
            SetCal72ColumnReadOnly()
            ScrollCal72GridToBottom()

            DataGridViewCal72.ClearSelection()

            LabelCal72Status.Text = "LOADED " & IO.Path.GetFileName(Cal72CsvFile)

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

        ElseIf RadioButton34586.Checked = True Then
            Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift6.csv"

        ElseIf RadioButton34587.Checked = True Then
            Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift7.csv"

        ElseIf RadioButton34588.Checked = True Then
            Cal72CsvFile = strPath & "\" & "3458A_CAL72_Drift8.csv"

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

        DataGridViewCal72.ClearSelection()

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
        DataGridViewCal72.Columns("CAL? 1,1 40k").ReadOnly = False
        DataGridViewCal72.Columns("CAL? 2,1 Vref").ReadOnly = False
        DataGridViewCal72.Columns("Notes").ReadOnly = False

        ' Editable column colours
        DataGridViewCal72.Columns("Day").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("Date").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("Time").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("Temp").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("CAL? 72").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("CAL? 1,1 40k").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("CAL? 2,1 Vref").DefaultCellStyle.BackColor = Color.White
        DataGridViewCal72.Columns("Notes").DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub ButtonCal72Read_Click(sender As Object, e As EventArgs) Handles ButtonCal72Read.Click

        If ButtonDev1Run.Enabled = False Then
            MessageBox.Show("Device 1 is not started",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)
            Exit Sub
        End If

        LabelCal72Status.Text = "READING CAL? 72"
        Me.Refresh()

        TextBoxCal72Value.Text =
        Query3458AValue(Cal72Command)

        LabelCal72Status.Text = "READING TEMP?"
        Me.Refresh()

        TextBoxCal72Temp.Text =
        Query3458AValue(Cal72TempCommand)

        LabelCal72Status.Text = "READING CAL? 1,1"
        Me.Refresh()

        TextBoxCal1140k.Text =
        Query3458AValue(Cal72Cal11Command)

        LabelCal72Status.Text = "READING CAL? 2,1"
        Me.Refresh()

        TextBoxCal21Vref.Text =
        Query3458AValue(Cal72Cal21Command)

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
        Dim cal11Value As Double = 0
        Dim cal21Value As Double = 0

        If Not Double.TryParse(TextBoxCal72Value.Text, NumberStyles.Float, CultureInfo.InvariantCulture, cal72Value) Then
            MessageBox.Show("Invalid CAL? 72 value.", "CAL? 72", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Not Double.TryParse(TextBoxCal72Temp.Text, NumberStyles.Float, CultureInfo.InvariantCulture, tempValue) Then
            MessageBox.Show("Invalid temperature value.", "CAL? 72", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Optional extra constants.
        ' These are only read when Device 1 is running.
        If ButtonDev1Run.Enabled = True Then

            LabelCal72Status.Text = "READING CAL? 1,1"
            Me.Refresh()

            Double.TryParse(Query3458AValue(Cal72Cal11Command),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        cal11Value)

            LabelCal72Status.Text = "READING CAL? 2,1"
            Me.Refresh()

            Double.TryParse(Query3458AValue(Cal72Cal21Command),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        cal21Value)

        End If

        Dim r As DataRow = Cal72Table.NewRow()

        r("Day") = GetSuggestedNextCal72Day()
        r("Date") = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        r("Time") = DateTime.Now.ToString("HH:mm", CultureInfo.InvariantCulture)
        r("Temp") = tempValue
        r("CAL? 72") = cal72Value
        r("CAL? 1,1 40k") = TextBoxCal1140k.Text.Trim()
        r("CAL? 2,1 Vref") = TextBoxCal21Vref.Text.Trim()
        r("Notes") = TextBoxCal72Notes.Text

        Cal72Table.Rows.Add(r)

        RecalculateCal72Table()
        UpdateCal72SummaryPanel()
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

        Dim deletingFirstRow As Boolean =
        DataGridViewCal72.SelectedRows(0).Index = 0

        If MessageBox.Show("Delete selected entry?", "CAL? 72", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Exit Sub
        End If

        DataGridViewCal72.Rows.Remove(DataGridViewCal72.SelectedRows(0))

        If deletingFirstRow Then
            RenumberCal72DaysFromFirstRow()
        End If

        RecalculateCal72Table()
        UpdateCal72SummaryPanel()
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
        UpdateCal72SummaryPanel()
        SaveCal72Csv()

        LabelCal72Status.Text = "EDITED AND RECALCULATED"

    End Sub

    Private Sub RecalculateCal72Table()

        If Cal72Table.Rows.Count = 0 Then Exit Sub

        Dim day1Cal72 As Double
        Dim firstDay As Double

        Dim day1Cal11 As Double
        Dim day1Cal21 As Double

        If Not Double.TryParse(Cal72Table.Rows(0)("CAL? 72").ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, day1Cal72) Then Exit Sub
        If Not Double.TryParse(Cal72Table.Rows(0)("Day").ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, firstDay) Then firstDay = 1

        Double.TryParse(Cal72Table.Rows(0)("CAL? 1,1 40k").ToString(),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    day1Cal11)

        Double.TryParse(Cal72Table.Rows(0)("CAL? 2,1 Vref").ToString(),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    day1Cal21)

        If day1Cal72 = 0 Then Exit Sub

        Dim previousDay As Double = firstDay
        Dim previousCal72 As Double = day1Cal72

        For i As Integer = 0 To Cal72Table.Rows.Count - 1

            Dim r As DataRow = Cal72Table.Rows(i)

            Dim thisDay As Double
            Dim thisCal72 As Double

            Dim thisCal11 As Double
            Dim thisCal21 As Double

            If Not Double.TryParse(r("Day").ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, thisDay) Then Continue For
            If Not Double.TryParse(r("CAL? 72").ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, thisCal72) Then Continue For

            Double.TryParse(r("CAL? 1,1 40k").ToString(),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        thisCal11)

            Double.TryParse(r("CAL? 2,1 Vref").ToString(),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        thisCal21)

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

            Dim cal11Text As String =
    r("CAL? 1,1 40k").ToString().Trim()

            Dim cal21Text As String =
    r("CAL? 2,1 Vref").ToString().Trim()

            If String.IsNullOrWhiteSpace(cal11Text) OrElse
   String.IsNullOrWhiteSpace(Cal72Table.Rows(0)("CAL? 1,1 40k").ToString()) Then

                r("CAL? 1,1 Dev") = DBNull.Value

            Else

                r("CAL? 1,1 Dev") =
        Math.Round(thisCal11 - day1Cal11, 10)

            End If

            If String.IsNullOrWhiteSpace(cal21Text) OrElse
   String.IsNullOrWhiteSpace(Cal72Table.Rows(0)("CAL? 2,1 Vref").ToString()) Then

                r("CAL? 2,1 Dev") = DBNull.Value

            Else

                r("CAL? 2,1 Dev") =
        Math.Round(thisCal21 - day1Cal21, 10)

            End If

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

            IO.File.AppendAllText(Cal72CsvFile, "SerialNumber," & Csv(TextBoxCal72SerialNumber.Text.Trim()) & Environment.NewLine & Environment.NewLine)

            IO.File.AppendAllText(Cal72CsvFile, "Day,Date,Time,Temp,CAL? 72,CAL? 1,1 40k,CAL? 2,1 Vref,CAL? 1,1 Dev,CAL? 2,1 Dev,Days From Day 1,Days From Last,Drift ppm Day 1,Avg ppm/day,Drift ppm Last,Notes" & Environment.NewLine)

            For Each r As DataRow In Cal72Table.Rows

                Dim lineOut As String =
                Csv(r("Day").ToString()) & "," &
                Csv(r("Date").ToString()) & "," &
                Csv(r("Time").ToString()) & "," &
                Csv(Convert.ToDouble(r("Temp")).ToString("G17", CultureInfo.InvariantCulture)) & "," &
                Csv(Convert.ToDouble(r("CAL? 72")).ToString("G17", CultureInfo.InvariantCulture)) & "," &
                Csv(r("CAL? 1,1 40k").ToString()) & "," &
                Csv(r("CAL? 2,1 Vref").ToString()) & "," &
                Csv(r("CAL? 1,1 Dev").ToString()) & "," &
                Csv(r("CAL? 2,1 Dev").ToString()) & "," &
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
                TextBoxCal72SerialNumber.Text = ""
                Exit Sub
            End If

            Dim lines() As String = IO.File.ReadAllLines(Cal72CsvFile)

            TextBoxCal72SerialNumber.Text = ""

            Dim startLine As Integer = 1

            If lines.Length > 0 AndAlso lines(0).StartsWith("SerialNumber,", StringComparison.OrdinalIgnoreCase) Then

                Dim snParts As List(Of String) = ParseCsvLine(lines(0))

                If snParts.Count >= 2 Then
                    TextBoxCal72SerialNumber.Text = snParts(1)
                End If

                startLine = 3

            End If

            For i As Integer = startLine To lines.Length - 1

                If String.IsNullOrWhiteSpace(lines(i)) Then Continue For

                If lines(i).StartsWith("Day,", StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim p As List(Of String) = ParseCsvLine(lines(i))

                Dim r As DataRow = Cal72Table.NewRow()

                If p.Count >= 15 Then

                    ' New CSV format with CAL? 1,1, CAL? 2,1 and deviation columns
                    r("Day") = ToDoubleSafe(p(0))
                    r("Date") = p(1)
                    r("Time") = p(2)
                    r("Temp") = ToDoubleSafe(p(3))
                    r("CAL? 72") = ToDoubleSafe(p(4))
                    r("CAL? 1,1 40k") = p(5)
                    r("CAL? 2,1 Vref") = p(6)
                    r("CAL? 1,1 Dev") = ToDoubleSafe(p(7))
                    r("CAL? 2,1 Dev") = ToDoubleSafe(p(8))
                    r("Days From Day 1") = ToDoubleSafe(p(9))
                    r("Days From Last") = ToDoubleSafe(p(10))
                    r("Drift ppm Day 1") = ToDoubleSafe(p(11))
                    r("Avg ppm/day") = ToDoubleSafe(p(12))
                    r("Drift ppm Last") = ToDoubleSafe(p(13))
                    r("Notes") = p(14)

                ElseIf p.Count >= 13 Then

                    ' Previous new format with CAL? 1,1 and CAL? 2,1 but no deviation columns
                    r("Day") = ToDoubleSafe(p(0))
                    r("Date") = p(1)
                    r("Time") = p(2)
                    r("Temp") = ToDoubleSafe(p(3))
                    r("CAL? 72") = ToDoubleSafe(p(4))
                    r("CAL? 1,1 40k") = p(5)
                    r("CAL? 2,1 Vref") = p(6)
                    r("CAL? 1,1 Dev") = DBNull.Value
                    r("CAL? 2,1 Dev") = DBNull.Value
                    r("Days From Day 1") = ToDoubleSafe(p(7))
                    r("Days From Last") = ToDoubleSafe(p(8))
                    r("Drift ppm Day 1") = ToDoubleSafe(p(9))
                    r("Avg ppm/day") = ToDoubleSafe(p(10))
                    r("Drift ppm Last") = ToDoubleSafe(p(11))
                    r("Notes") = p(12)

                ElseIf p.Count >= 11 Then

                    ' Old CSV format without CAL? 1,1 and CAL? 2,1 columns
                    r("Day") = ToDoubleSafe(p(0))
                    r("Date") = p(1)
                    r("Time") = p(2)
                    r("Temp") = ToDoubleSafe(p(3))
                    r("CAL? 72") = ToDoubleSafe(p(4))
                    r("CAL? 1,1 40k") = ""
                    r("CAL? 2,1 Vref") = ""
                    r("CAL? 1,1 Dev") = DBNull.Value
                    r("CAL? 2,1 Dev") = DBNull.Value
                    r("Days From Day 1") = ToDoubleSafe(p(5))
                    r("Days From Last") = ToDoubleSafe(p(6))
                    r("Drift ppm Day 1") = ToDoubleSafe(p(7))
                    r("Avg ppm/day") = ToDoubleSafe(p(8))
                    r("Drift ppm Last") = ToDoubleSafe(p(9))
                    r("Notes") = p(10)

                Else

                    Continue For

                End If

                Cal72Table.Rows.Add(r)

            Next

        Catch ex As Exception

            MessageBox.Show("Unable to load CAL72 file: " & ex.Message,
                "CAL72 File Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)

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


    Private Sub ButtonCal72Help_Click(sender As Object, e As EventArgs) Handles ButtonCal72Help.Click

        MessageBox.Show(
"3458A U180 A/D CURRENT STEERING HYBRID DRIFT MONITOR" & vbCrLf & vbCrLf &
"- Connect Device 1 to your 3458A and leave in STOP position" & vbCrLf &
"- Allow 3458A to thermally stabilise before recording data" & vbCrLf &
"- Perform ACAL DCV before recording values where possible" & vbCrLf &
"- First entry becomes the permanent Day 1 reference baseline" & vbCrLf &
"- CAL? 72 drift can be temp. related - monitor temp. carefully" & vbCrLf &
"- Lower ppm/day values indicate better U180 stability" & vbCrLf &
"- Values can be entered manually or read directly via GPIB" & vbCrLf &
"- CAL? 1,1 and CAL? 2,1 recorded for historical reference" & vbCrLf &
"- Deviation cols change from Day 1 CAL? 1,1 and CAL? 2,1" & vbCrLf &
"- Estimated Annual Drift = Average Drift ppm/day × 365.25" & vbCrLf &
"- Total Drift Ref Day1 shows drift rel. to the 1st recorded entry" & vbCrLf &
"- Worst Drift Ref Day1 shows the largest drift recorded" & vbCrLf &
"- Days Since Last Entry shows elapsed time since last" & vbCrLf &
"- Deleting Day 1 renumbers and creates a new Day 1 baseline",
"3458A U180 Drift Monitor Help",
MessageBoxButtons.OK,
MessageBoxIcon.Information)

    End Sub


    Private Sub UpdateCal72SummaryPanel()

        If Cal72Table.Rows.Count = 0 Then

            LabelCal72LatestDay.Text = "-"
            LabelCal72LatestValue.Text = "-"
            LabelCal72LatestDrift.Text = "-"
            LabelCal72LatestAvg.Text = "-"
            LabelCal72LatestLast.Text = "-"
            LabelCal72LatestEntries.Text = "0"
            LabelCal72AnnualDrift.Text = "-"
            LabelCal72WorstDrift.Text = "-"

            Exit Sub

        End If

        Dim r As DataRow =
        Cal72Table.Rows(Cal72Table.Rows.Count - 1)

        LabelCal72LatestDay.Text = r("Day").ToString()
        LabelCal72LatestValue.Text = r("CAL? 72").ToString()
        LabelCal72LatestDrift.Text = r("Drift ppm Day 1").ToString()
        LabelCal72LatestAvg.Text = r("Avg ppm/day").ToString()
        LabelCal72LatestLast.Text = r("Drift ppm Last").ToString()
        LabelCal72LatestEntries.Text = Cal72Table.Rows.Count.ToString()

        Dim avgPpmPerDay As Double

        If Double.TryParse(r("Avg ppm/day").ToString(),
                       NumberStyles.Float,
                       CultureInfo.InvariantCulture,
                       avgPpmPerDay) Then

            LabelCal72AnnualDrift.Text =
            (avgPpmPerDay * 365.25).ToString("0.000")

        Else

            LabelCal72AnnualDrift.Text = "-"

        End If

        Dim worstDrift As Double = 0

        For Each row As DataRow In Cal72Table.Rows

            Dim drift As Double

            If Double.TryParse(row("Drift ppm Day 1").ToString(),
                           NumberStyles.Float,
                           CultureInfo.InvariantCulture,
                           drift) Then

                If Math.Abs(drift) > Math.Abs(worstDrift) Then
                    worstDrift = drift
                End If

            End If

        Next

        LabelCal72WorstDrift.Text =
        worstDrift.ToString("0.000000")

    End Sub


    Private Sub RenumberCal72DaysFromFirstRow()

        If Cal72Table.Rows.Count = 0 Then Exit Sub

        Dim firstDay As Double

        If Not Double.TryParse(Cal72Table.Rows(0)("Day").ToString(),
                               NumberStyles.Float,
                               CultureInfo.InvariantCulture,
                               firstDay) Then
            Exit Sub
        End If

        For Each r As DataRow In Cal72Table.Rows

            Dim oldDay As Double

            If Double.TryParse(r("Day").ToString(),
                               NumberStyles.Float,
                               CultureInfo.InvariantCulture,
                               oldDay) Then

                r("Day") = oldDay - firstDay + 1

            End If

        Next

    End Sub


    Private Sub ButtonCal72Backup_Click(sender As Object, e As EventArgs) Handles ButtonCal72Backup.Click

        Try

            Dim files() As String = {
                IO.Path.Combine(strPath, "3458A_CAL72_Drift.csv"),
                IO.Path.Combine(strPath, "3458A_CAL72_Drift2.csv"),
                IO.Path.Combine(strPath, "3458A_CAL72_Drift3.csv"),
                IO.Path.Combine(strPath, "3458A_CAL72_Drift4.csv"),
                IO.Path.Combine(strPath, "3458A_CAL72_Drift5.csv"),
                IO.Path.Combine(strPath, "3458A_CAL72_Drift6.csv"),
                IO.Path.Combine(strPath, "3458A_CAL72_Drift7.csv"),
                IO.Path.Combine(strPath, "3458A_CAL72_Drift8.csv")
            }

            Dim backupCount As Integer = 0

            For Each fileName As String In files

                If IO.File.Exists(fileName) Then

                    IO.File.Copy(fileName,
                                 fileName & ".bak",
                                 True)

                    backupCount += 1

                End If

            Next

            MessageBox.Show(
                backupCount.ToString() & " CSV file(s) backed up successfully.",
                "CAL72 Backup",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

        Catch ex As Exception

            MessageBox.Show(
                "Backup failed: " & ex.Message,
                "CAL72 Backup",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)

        End Try

    End Sub


End Class