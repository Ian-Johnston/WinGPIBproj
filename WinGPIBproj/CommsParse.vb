Partial Class Formtest

    Dim NoCommsVars As Integer

    Private Sub Timer6_Tick(sender As Object, e As EventArgs) Handles Timer6.Tick

        ' If serial port has closed
        If Not SerialPort1.IsOpen Then
            SerialPort1.Close()
            connect_BTN.Text = "Connect"
            CommsStatus = False
            Timer1.Enabled = False
        End If

        ' If serial port is open
        If SerialPort1.IsOpen Then

            ' Set number of vars depending if WryTech PDVS2mini or not
            NoCommsVars = If(WryTech.Checked, 23, 21)

            receivedData = ReceiveSerialData()

            intPos = InStr(receivedData, vbCr)
            If intPos > 0 Then

                ' Parse received data
                If InStr(receivedData, vbCrLf) Then
                    Dim tempBuffer() As String = Split(receivedData, vbCrLf)  ' Split out each line received
                    If UBound(tempBuffer) > NoCommsVars Then  ' number of substrings

                        ' Flag that comms is running
                        CommsStatus = True
                        Call RxLEDon()

                        Dim delimiter As Char = ","

                        ' Create a dictionary to map keys to actions
                        Dim keyActions As New Dictionary(Of String, Action(Of String)) From {
                            {"KV", Sub(value) UpdateLabel(LabelKeyVoltage, value, KeyV)},
                            {"BV", Sub(value) UpdateTextBox(TextBoxLabelBatteryV, value, BattV)},
                            {"BVFM", Sub(value) LabelBatteryVFeedMult.Text = value},
                            {"OVF", Sub(value) UpdateTextBox(TextBoxLabelOutputVFeedback, value)},
                            {"OVFM", Sub(value) LabelOutputVFeedMult.Text = value},
                            {"BC", Sub(value) UpdateTextBox(TextBoxLabelBatteryCharge, value)},
                            {"BLI", Sub(value) UpdateTextBox(TextBoxLabelBatteryLowInd, value)},
                            {"Mode", Sub(value) UpdateTextBox(TextBoxLabelMode, value)},
                            {"BMIM", Sub(value) LabelBatteryMonICMult.Text = value},
                            {"BI", Sub(value) UpdateTextBox(TextBoxLabelBatteryI, value)},
                            {"CMS", Sub(value) ' Do nothing
                                    End Sub},
                            {"dacZ0", Sub(value) LabeldacZero0Cal.Text = value},
                            {"dacS0", Sub(value) LabeldacSpan0Cal.Text = value},
                            {"dacS1", Sub(value) LabeldacSpan1Cal.Text = value},
                            {"dacS2", Sub(value) LabeldacSpan2Cal.Text = value},
                            {"dacS3", Sub(value) LabeldacSpan3Cal.Text = value},
                            {"dacS4", Sub(value) LabeldacSpan4Cal.Text = value},
                            {"dacS5", Sub(value) LabeldacSpan5Cal.Text = value},
                            {"dacS6", Sub(value) LabeldacSpan6Cal.Text = value},
                            {"dacS7", Sub(value) LabeldacSpan7Cal.Text = value},
                            {"dacS8", Sub(value) LabeldacSpan8Cal.Text = value},
                            {"dacS9", Sub(value) LabeldacSpan9Cal.Text = value}
                        }

                        ' Handle the extra parameters for WryTech PDVS2mini
                        If WryTech.Checked Then
                            keyActions.Add("dacS10", Sub(value) LabeldacSpan10Cal.Text = value)
                            keyActions.Add("TEMP", Sub(value) LabelTemperature3.Text = value)
                        End If

                        ' Parse each line
                        For i As Integer = 0 To NoCommsVars
                            Dim substrings() As String = tempBuffer(i).Split(delimiter)
                            If keyActions.ContainsKey(substrings(0)) Then
                                keyActions(substrings(0)).Invoke(substrings(2))
                            Else
                                Debug.WriteLine($"Unhandled key: {substrings(0)}")
                            End If
                        Next

                        ' Retain any extra data
                        receivedData = String.Join("", tempBuffer.Skip(NoCommsVars + 1).Take(UBound(tempBuffer) - NoCommsVars - 1))
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub UpdateLabel(label As Label, value As String, ByRef field As Double)
        label.Text = Replace(value, ",", ".")
        field = CDbl(Val(label.Text))
    End Sub

    Private Sub UpdateTextBox(textBox As TextBox, value As String, Optional ByRef field As Double = 0)
        textBox.Text = value
        If field <> 0 Then
            field = CDbl(Val(value))
        End If
    End Sub

End Class
