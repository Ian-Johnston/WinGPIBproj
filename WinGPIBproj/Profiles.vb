Imports System.Configuration
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
'Imports System.Threading
'Imports System.Xml.Serialization
'Imports Microsoft.Office.Interop.Word
'Imports System.Collections.Generic

Partial Class Formtest

    ' Profile selection vars
    Dim ProfDev1checked_1 As Boolean = True
    Dim ProfDev1checked_2 As Boolean = False
    Dim ProfDev1checked_3 As Boolean = False
    Dim ProfDev2checked_1 As Boolean = True
    Dim ProfDev2checked_2 As Boolean = False
    Dim ProfDev2checked_3 As Boolean = False
    Dim TextEditorPath As String

    ' 3458A Cal Pre-Run
    Private Sub BtnSave3458A_Click(sender As Object, e As EventArgs) Handles BtnSave3458A.Click

        My.Settings.data508 = CalRam3458APreRun.Text
        My.Settings.Save()

    End Sub

    ' Settings page
    Private Sub CheckBoxAllowSaveAnytime_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxAllowSaveAnytime.CheckedChanged

        My.Settings.data505 = CheckBoxAllowSaveAnytime.Checked          ' save off immediately
        My.Settings.Save()

    End Sub

    Private Sub CheckBoxEnableTooltips_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxEnableTooltips.CheckedChanged

        If CheckBoxEnableTooltips.Checked = True Then
            ToolTip1.Active = True
        Else
            ToolTip1.Active = False
        End If

        My.Settings.data507 = CheckBoxEnableTooltips.Checked          ' save off immediately
        My.Settings.Save()

    End Sub

    Private Sub TextBoxTextEditor_TextChanged(sender As Object, e As EventArgs) Handles TextBoxTextEditor.TextChanged

        TextEditorPath = TextBoxTextEditor.Text ' Adjust the path as needed
        My.Settings.data506 = TextBoxTextEditor.Text          ' save off immediately
        My.Settings.Save()
        'Process.Start(notepadPlusPlusPath, strPath & "\" & "GPIBchannels.txt")

    End Sub


    Private Sub ButtonSaveLiveSettings_Click(sender As Object, e As EventArgs) Handles ButtonSaveLiveSettings.Click

        ' Save Live Chart numerical settings
        My.Settings.data255 = XaxisPoints.Text
        My.Settings.data256 = Dev1Max.Text
        My.Settings.data257 = Dev1Min.Text
        My.Settings.data258 = LCTempMax.Text
        My.Settings.data259 = LCTempMin.Text
        My.Settings.data525 = TextBoxAvgWindow.Text

    End Sub

    Private Sub ButtonSaveTempHumSettings_Click(sender As Object, e As EventArgs) Handles ButtonSaveTempHumSettings.Click

        ' Temperature & Humidity Sensor
        My.Settings.data16 = txtname3.Text
        My.Settings.data319 = TextBoxProtocolInput.Text
        My.Settings.data320 = TextBoxParseLeft.Text
        My.Settings.data321 = TextBoxParseRight.Text
        My.Settings.data322 = TextBoxRegex.Text
        My.Settings.data323 = TextBoxTempArithmentic.Text
        My.Settings.data324 = TextBoxTempUnits.Text
        My.Settings.data325 = TextBoxHumUnits.Text
        My.Settings.data326 = lstIntf3.SelectedItem
        My.Settings.data327 = TextBoxSerialPortBaud.Text
        My.Settings.data328 = TextBoxSerialPortBits.Text
        My.Settings.data329 = TextBoxSerialPortParity.Text
        My.Settings.data330 = TextBoxSerialPortStop.Text
        My.Settings.data331 = TextBoxSerialPortHand.Text
        My.Settings.data332 = ComboBoxPort.SelectedItem
        My.Settings.data333 = comPort_ComboBox.SelectedItem
        My.Settings.data334 = CheckBoxParseLeftRight.Checked     ' temp/hum checkbox
        My.Settings.data335 = CheckBoxRegex.Checked              ' temp/hum checkbox
        My.Settings.data336 = CheckBoxArithmetic.Checked         ' temp/hum checkbox
        My.Settings.data504 = TextBoxTempHumSample.Text

    End Sub

    Private Sub ButtonSaveSettings_Click(sender As Object, e As EventArgs) Handles ButtonSaveSettings.Click

        ' Profile checkboxes Dev 1 & 2
        My.Settings.Dev1Prof1 = ProfDev1_1.Checked
        My.Settings.Dev1Prof2 = ProfDev1_2.Checked
        My.Settings.Dev1Prof3 = ProfDev1_3.Checked
        My.Settings.Dev1Prof4 = ProfDev1_4.Checked
        My.Settings.Dev1Prof5 = ProfDev1_5.Checked
        My.Settings.Dev1Prof6 = ProfDev1_6.Checked
        My.Settings.Dev1Prof7 = ProfDev1_7.Checked
        My.Settings.Dev1Prof8 = ProfDev1_8.Checked
        My.Settings.Dev1Prof9 = ProfDev1_9.Checked
        My.Settings.Dev1Prof10 = ProfDev1_10.Checked
        My.Settings.Dev1Prof11 = ProfDev1_11.Checked
        My.Settings.Dev1Prof12 = ProfDev1_12.Checked
        My.Settings.Dev2Prof1 = ProfDev2_1.Checked
        My.Settings.Dev2Prof2 = ProfDev2_2.Checked
        My.Settings.Dev2Prof3 = ProfDev2_3.Checked
        My.Settings.Dev2Prof4 = ProfDev2_4.Checked
        My.Settings.Dev2Prof5 = ProfDev2_5.Checked
        My.Settings.Dev2Prof6 = ProfDev2_6.Checked
        My.Settings.Dev2Prof7 = ProfDev2_7.Checked
        My.Settings.Dev2Prof8 = ProfDev2_8.Checked
        My.Settings.Dev2Prof9 = ProfDev2_9.Checked
        My.Settings.Dev2Prof10 = ProfDev2_10.Checked
        My.Settings.Dev2Prof11 = ProfDev2_11.Checked
        My.Settings.Dev2Prof12 = ProfDev2_12.Checked

        ' Save general settings
        My.Settings.data11 = CSVfilename.Text
        My.Settings.data12 = CSVfilepath.Text
        My.Settings.data13 = XaxisPoints.Text
        My.Settings.data14 = Dev1Max.Text
        My.Settings.data15 = Dev1Min.Text
        My.Settings.data16 = txtname3.Text
        My.Settings.data19 = Dev12SampleRate.Text   ' Both devices Sample Rate
        My.Settings.data78 = TempOffset.Text

        My.Settings.data500 = Dev1Units.Text
        My.Settings.data501 = Dev2Units.Text


        ' Save Dev1 Profile 1
        If (ProfDev1_1.Checked = True) Then
            My.Settings.data1 = txtname1.Text
            My.Settings.data3 = txtaddr1.Text
            My.Settings.data5 = CommandStart1.Text
            My.Settings.data17 = CommandStart1run.Text
            My.Settings.data7 = CommandStop1.Text
            My.Settings.data33 = lstIntf1.SelectedIndex
            My.Settings.data9 = Dev1SampleRate.Text
            My.Settings.data36 = Dev1PollingEnable.Checked
            My.Settings.data37 = Dev1removeletters.Checked
            My.Settings.data38 = IgnoreErrors1.Checked
            My.Settings.data39 = Dev1TerminatorEnable.Checked
            My.Settings.data40 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data66 = Dev1STBMask.Text
            My.Settings.data72 = Div1000Dev1.Checked
            My.Settings.data79 = Dev13457Aseven.Checked
            My.Settings.data85 = Dev1TerminatorEnable2.Checked
            My.Settings.data207 = Dev1K2001isolatedata.Checked
            My.Settings.data208 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data225 = Mult1000Dev1.Checked
            My.Settings.data231 = Val(Dev1Timeout.Text)
            My.Settings.data243 = Val(Dev1delayop.Text)
            My.Settings.data271 = txtq1d.Text
            My.Settings.data272 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data273 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data274 = Dev1IntEnable.Checked
            My.Settings.data337 = Dev1Regex.Checked
            My.Settings.data453 = Dev1DecimalNumDPs.Text
            My.Settings.data484 = Dev1IntEnable.Checked
            My.Settings.data509 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 2
        If (ProfDev1_2.Checked = True) Then
            My.Settings.data1b = txtname1.Text
            My.Settings.data3b = txtaddr1.Text
            My.Settings.data5b = CommandStart1.Text
            My.Settings.data17b = CommandStart1run.Text
            My.Settings.data7b = CommandStop1.Text
            My.Settings.data34 = lstIntf1.SelectedIndex
            My.Settings.data9b = Dev1SampleRate.Text
            My.Settings.data41 = Dev1PollingEnable.Checked
            My.Settings.data42 = Dev1removeletters.Checked
            My.Settings.data43 = IgnoreErrors1.Checked
            My.Settings.data44 = Dev1TerminatorEnable.Checked
            My.Settings.data45 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data68 = Dev1STBMask.Text
            My.Settings.data73 = Div1000Dev1.Checked
            My.Settings.data81 = Dev13457Aseven.Checked
            My.Settings.data87 = Dev1TerminatorEnable2.Checked
            My.Settings.data209 = Dev1K2001isolatedata.Checked
            My.Settings.data210 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data226 = Mult1000Dev1.Checked
            My.Settings.data232 = Val(Dev1Timeout.Text)
            My.Settings.data244 = Val(Dev1delayop.Text)
            My.Settings.data275 = txtq1d.Text
            My.Settings.data276 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data277 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data278 = Dev1IntEnable.Checked
            My.Settings.data338 = Dev1Regex.Checked
            My.Settings.data454 = Dev1DecimalNumDPs.Text
            My.Settings.data486 = Dev1IntEnable.Checked
            My.Settings.data510 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 3
        If (ProfDev1_3.Checked = True) Then
            My.Settings.data1c = txtname1.Text
            My.Settings.data3c = txtaddr1.Text
            My.Settings.data5c = CommandStart1.Text
            My.Settings.data17c = CommandStart1run.Text
            My.Settings.data7c = CommandStop1.Text
            My.Settings.data35 = lstIntf1.SelectedIndex
            My.Settings.data9c = Dev1SampleRate.Text
            My.Settings.data46 = Dev1PollingEnable.Checked
            My.Settings.data47 = Dev1removeletters.Checked
            My.Settings.data48 = IgnoreErrors1.Checked
            My.Settings.data49 = Dev1TerminatorEnable.Checked
            My.Settings.data50 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data70 = Dev1STBMask.Text
            My.Settings.data74 = Div1000Dev1.Checked
            My.Settings.data83 = Dev13457Aseven.Checked
            My.Settings.data89 = Dev1TerminatorEnable2.Checked
            My.Settings.data211 = Dev1K2001isolatedata.Checked
            My.Settings.data212 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data227 = Mult1000Dev1.Checked
            My.Settings.data233 = Val(Dev1Timeout.Text)
            My.Settings.data245 = Val(Dev1delayop.Text)
            My.Settings.data279 = txtq1d.Text
            My.Settings.data280 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data281 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data282 = Dev1IntEnable.Checked
            My.Settings.data339 = Dev1Regex.Checked
            My.Settings.data455 = Dev1DecimalNumDPs.Text
            My.Settings.data488 = Dev1IntEnable.Checked
            My.Settings.data511 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 4
        If (ProfDev1_4.Checked = True) Then
            My.Settings.data139 = txtname1.Text
            My.Settings.data140 = CommandStart1.Text
            My.Settings.data141 = CommandStart1run.Text
            My.Settings.data142 = CommandStop1.Text
            My.Settings.data143 = lstIntf1.SelectedIndex
            My.Settings.data144 = txtaddr1.Text
            My.Settings.data145 = Dev1SampleRate.Text
            My.Settings.data146 = Dev1PollingEnable.Checked
            My.Settings.data147 = Dev1removeletters.Checked
            My.Settings.data148 = IgnoreErrors1.Checked
            My.Settings.data149 = Dev1TerminatorEnable.Checked
            My.Settings.data150 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data151 = Dev1STBMask.Text
            My.Settings.data152 = Div1000Dev1.Checked
            My.Settings.data153 = Dev13457Aseven.Checked
            My.Settings.data154 = Dev1TerminatorEnable2.Checked
            My.Settings.data213 = Dev1K2001isolatedata.Checked
            My.Settings.data214 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data228 = Mult1000Dev1.Checked
            My.Settings.data234 = Val(Dev1Timeout.Text)
            My.Settings.data246 = Val(Dev1delayop.Text)
            My.Settings.data283 = txtq1d.Text
            My.Settings.data284 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data285 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data286 = Dev1IntEnable.Checked
            My.Settings.data340 = Dev1Regex.Checked
            My.Settings.data456 = Dev1DecimalNumDPs.Text
            My.Settings.data490 = Dev1IntEnable.Checked
            My.Settings.data512 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 5
        If (ProfDev1_5.Checked = True) Then
            My.Settings.data155 = txtname1.Text
            My.Settings.data156 = CommandStart1.Text
            My.Settings.data157 = CommandStart1run.Text
            My.Settings.data158 = CommandStop1.Text
            My.Settings.data159 = lstIntf1.SelectedIndex
            My.Settings.data160 = txtaddr1.Text
            My.Settings.data161 = Dev1SampleRate.Text
            My.Settings.data162 = Dev1PollingEnable.Checked
            My.Settings.data163 = Dev1removeletters.Checked
            My.Settings.data164 = IgnoreErrors1.Checked
            My.Settings.data165 = Dev1TerminatorEnable.Checked
            My.Settings.data166 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data167 = Dev1STBMask.Text
            My.Settings.data168 = Div1000Dev1.Checked
            My.Settings.data169 = Dev13457Aseven.Checked
            My.Settings.data170 = Dev1TerminatorEnable2.Checked
            My.Settings.data215 = Dev1K2001isolatedata.Checked
            My.Settings.data216 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data229 = Mult1000Dev1.Checked
            My.Settings.data235 = Val(Dev1Timeout.Text)
            My.Settings.data247 = Val(Dev1delayop.Text)
            My.Settings.data287 = txtq1d.Text
            My.Settings.data288 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data289 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data290 = Dev1IntEnable.Checked
            My.Settings.data341 = Dev1Regex.Checked
            My.Settings.data457 = Dev1DecimalNumDPs.Text
            My.Settings.data492 = Dev1IntEnable.Checked
            My.Settings.data513 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 6
        If (ProfDev1_6.Checked = True) Then
            My.Settings.data171 = txtname1.Text
            My.Settings.data172 = CommandStart1.Text
            My.Settings.data173 = CommandStart1run.Text
            My.Settings.data174 = CommandStop1.Text
            My.Settings.data175 = lstIntf1.SelectedIndex
            My.Settings.data176 = txtaddr1.Text
            My.Settings.data177 = Dev1SampleRate.Text
            My.Settings.data178 = Dev1PollingEnable.Checked
            My.Settings.data179 = Dev1removeletters.Checked
            My.Settings.data180 = IgnoreErrors1.Checked
            My.Settings.data181 = Dev1TerminatorEnable.Checked
            My.Settings.data182 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data183 = Dev1STBMask.Text
            My.Settings.data184 = Div1000Dev1.Checked
            My.Settings.data185 = Dev13457Aseven.Checked
            My.Settings.data186 = Dev1TerminatorEnable2.Checked
            My.Settings.data217 = Dev1K2001isolatedata.Checked
            My.Settings.data218 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data230 = Mult1000Dev1.Checked
            My.Settings.data236 = Val(Dev1Timeout.Text)
            My.Settings.data248 = Val(Dev1delayop.Text)
            My.Settings.data291 = txtq1d.Text
            My.Settings.data292 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data293 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data294 = Dev1IntEnable.Checked
            My.Settings.data342 = Dev1Regex.Checked
            My.Settings.data458 = Dev1DecimalNumDPs.Text
            My.Settings.data494 = Dev1IntEnable.Checked
            My.Settings.data514 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 7
        If (ProfDev1_7.Checked = True) Then
            My.Settings.data349 = txtname1.Text
            My.Settings.data350 = CommandStart1.Text
            My.Settings.data351 = CommandStart1run.Text
            My.Settings.data352 = CommandStop1.Text
            My.Settings.data353 = lstIntf1.SelectedIndex
            My.Settings.data354 = txtaddr1.Text
            My.Settings.data355 = Dev1SampleRate.Text
            My.Settings.data356 = Dev1PollingEnable.Checked
            My.Settings.data357 = Dev1removeletters.Checked
            My.Settings.data358 = IgnoreErrors1.Checked
            My.Settings.data359 = Dev1TerminatorEnable.Checked
            My.Settings.data360 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data361 = Dev1STBMask.Text
            My.Settings.data362 = Div1000Dev1.Checked
            My.Settings.data363 = Dev13457Aseven.Checked
            My.Settings.data364 = Dev1TerminatorEnable2.Checked
            My.Settings.data365 = Dev1K2001isolatedata.Checked
            My.Settings.data366 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data367 = Mult1000Dev1.Checked
            My.Settings.data368 = Val(Dev1Timeout.Text)
            My.Settings.data369 = Val(Dev1delayop.Text)
            My.Settings.data370 = txtq1d.Text
            My.Settings.data371 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data372 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data373 = Dev1IntEnable.Checked
            My.Settings.data374 = Dev1Regex.Checked
            My.Settings.data459 = Dev1DecimalNumDPs.Text
            My.Settings.data496 = Dev1IntEnable.Checked
            My.Settings.data515 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 8
        If (ProfDev1_8.Checked = True) Then
            My.Settings.data375 = txtname1.Text
            My.Settings.data376 = CommandStart1.Text
            My.Settings.data377 = CommandStart1run.Text
            My.Settings.data378 = CommandStop1.Text
            My.Settings.data379 = lstIntf1.SelectedIndex
            My.Settings.data380 = txtaddr1.Text
            My.Settings.data381 = Dev1SampleRate.Text
            My.Settings.data382 = Dev1PollingEnable.Checked
            My.Settings.data383 = Dev1removeletters.Checked
            My.Settings.data384 = IgnoreErrors1.Checked
            My.Settings.data385 = Dev1TerminatorEnable.Checked
            My.Settings.data386 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data387 = Dev1STBMask.Text
            My.Settings.data388 = Div1000Dev1.Checked
            My.Settings.data389 = Dev13457Aseven.Checked
            My.Settings.data390 = Dev1TerminatorEnable2.Checked
            My.Settings.data391 = Dev1K2001isolatedata.Checked
            My.Settings.data392 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data393 = Mult1000Dev1.Checked
            My.Settings.data394 = Val(Dev1Timeout.Text)
            My.Settings.data395 = Val(Dev1delayop.Text)
            My.Settings.data396 = txtq1d.Text
            My.Settings.data397 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data398 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data399 = Dev1IntEnable.Checked
            My.Settings.data400 = Dev1Regex.Checked
            My.Settings.data460 = Dev1DecimalNumDPs.Text
            My.Settings.data498 = Dev1IntEnable.Checked
            My.Settings.data516 = txtOperationDev1.Text
        End If


        ' Save Dev1 Profile 9
        If (ProfDev1_9.Checked = True) Then
            My.Settings.data526 = txtname1.Text
            My.Settings.data527 = CommandStart1.Text
            My.Settings.data528 = CommandStart1run.Text
            My.Settings.data529 = CommandStop1.Text
            My.Settings.data530 = lstIntf1.SelectedIndex
            My.Settings.data531 = txtaddr1.Text
            My.Settings.data532 = Dev1SampleRate.Text
            My.Settings.data533 = Dev1PollingEnable.Checked
            My.Settings.data534 = Dev1removeletters.Checked
            My.Settings.data535 = IgnoreErrors1.Checked
            My.Settings.data536 = Dev1TerminatorEnable.Checked
            My.Settings.data537 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data538 = Dev1STBMask.Text
            My.Settings.data539 = Div1000Dev1.Checked
            My.Settings.data540 = Dev13457Aseven.Checked
            My.Settings.data541 = Dev1TerminatorEnable2.Checked
            My.Settings.data542 = Dev1K2001isolatedata.Checked
            My.Settings.data543 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data544 = Mult1000Dev1.Checked
            My.Settings.data545 = Val(Dev1Timeout.Text)
            My.Settings.data546 = Val(Dev1delayop.Text)
            My.Settings.data547 = txtq1d.Text
            My.Settings.data548 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data549 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data550 = Dev1IntEnable.Checked
            My.Settings.data551 = Dev1Regex.Checked
            My.Settings.data552 = Dev1DecimalNumDPs.Text
            My.Settings.data553 = Dev1IntEnable.Checked
            My.Settings.data554 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 10
        If (ProfDev1_10.Checked = True) Then
            My.Settings.data555 = txtname1.Text
            My.Settings.data556 = CommandStart1.Text
            My.Settings.data557 = CommandStart1run.Text
            My.Settings.data558 = CommandStop1.Text
            My.Settings.data559 = lstIntf1.SelectedIndex
            My.Settings.data560 = txtaddr1.Text
            My.Settings.data561 = Dev1SampleRate.Text
            My.Settings.data562 = Dev1PollingEnable.Checked
            My.Settings.data563 = Dev1removeletters.Checked
            My.Settings.data564 = IgnoreErrors1.Checked
            My.Settings.data565 = Dev1TerminatorEnable.Checked
            My.Settings.data566 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data567 = Dev1STBMask.Text
            My.Settings.data568 = Div1000Dev1.Checked
            My.Settings.data569 = Dev13457Aseven.Checked
            My.Settings.data570 = Dev1TerminatorEnable2.Checked
            My.Settings.data571 = Dev1K2001isolatedata.Checked
            My.Settings.data572 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data573 = Mult1000Dev1.Checked
            My.Settings.data574 = Val(Dev1Timeout.Text)
            My.Settings.data575 = Val(Dev1delayop.Text)
            My.Settings.data576 = txtq1d.Text
            My.Settings.data577 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data578 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data579 = Dev1IntEnable.Checked
            My.Settings.data580 = Dev1Regex.Checked
            My.Settings.data581 = Dev1DecimalNumDPs.Text
            My.Settings.data582 = Dev1IntEnable.Checked
            My.Settings.data583 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 11
        If (ProfDev1_11.Checked = True) Then
            My.Settings.data584 = txtname1.Text
            My.Settings.data585 = CommandStart1.Text
            My.Settings.data586 = CommandStart1run.Text
            My.Settings.data587 = CommandStop1.Text
            My.Settings.data588 = lstIntf1.SelectedIndex
            My.Settings.data589 = txtaddr1.Text
            My.Settings.data590 = Dev1SampleRate.Text
            My.Settings.data591 = Dev1PollingEnable.Checked
            My.Settings.data592 = Dev1removeletters.Checked
            My.Settings.data593 = IgnoreErrors1.Checked
            My.Settings.data594 = Dev1TerminatorEnable.Checked
            My.Settings.data595 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data596 = Dev1STBMask.Text
            My.Settings.data597 = Div1000Dev1.Checked
            My.Settings.data598 = Dev13457Aseven.Checked
            My.Settings.data599 = Dev1TerminatorEnable2.Checked
            My.Settings.data600 = Dev1K2001isolatedata.Checked
            My.Settings.data601 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data602 = Mult1000Dev1.Checked
            My.Settings.data603 = Val(Dev1Timeout.Text)
            My.Settings.data604 = Val(Dev1delayop.Text)
            My.Settings.data605 = txtq1d.Text
            My.Settings.data606 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data607 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data608 = Dev1IntEnable.Checked
            My.Settings.data609 = Dev1Regex.Checked
            My.Settings.data610 = Dev1DecimalNumDPs.Text
            My.Settings.data611 = Dev1IntEnable.Checked
            My.Settings.data612 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 12
        If (ProfDev1_12.Checked = True) Then
            My.Settings.data613 = txtname1.Text
            My.Settings.data614 = CommandStart1.Text
            My.Settings.data615 = CommandStart1run.Text
            My.Settings.data616 = CommandStop1.Text
            My.Settings.data617 = lstIntf1.SelectedIndex
            My.Settings.data618 = txtaddr1.Text
            My.Settings.data619 = Dev1SampleRate.Text
            My.Settings.data620 = Dev1PollingEnable.Checked
            My.Settings.data621 = Dev1removeletters.Checked
            My.Settings.data622 = IgnoreErrors1.Checked
            My.Settings.data623 = Dev1TerminatorEnable.Checked
            My.Settings.data624 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data625 = Dev1STBMask.Text
            My.Settings.data626 = Div1000Dev1.Checked
            My.Settings.data627 = Dev13457Aseven.Checked
            My.Settings.data628 = Dev1TerminatorEnable2.Checked
            My.Settings.data629 = Dev1K2001isolatedata.Checked
            My.Settings.data630 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data631 = Mult1000Dev1.Checked
            My.Settings.data632 = Val(Dev1Timeout.Text)
            My.Settings.data633 = Val(Dev1delayop.Text)
            My.Settings.data634 = txtq1d.Text
            My.Settings.data635 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data636 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data637 = Dev1IntEnable.Checked
            My.Settings.data638 = Dev1Regex.Checked
            My.Settings.data639 = Dev1DecimalNumDPs.Text
            My.Settings.data640 = Dev1IntEnable.Checked
            My.Settings.data641 = txtOperationDev1.Text
        End If

        ' Save Dev2 Profile 1
        If (ProfDev2_1.Checked = True) Then
            My.Settings.data2 = txtname2.Text
            My.Settings.data4 = txtaddr2.Text
            My.Settings.data6 = CommandStart2.Text
            My.Settings.data18 = CommandStart2run.Text
            My.Settings.data8 = CommandStop2.Text
            My.Settings.data30 = lstIntf2.SelectedIndex
            My.Settings.data10 = Dev2SampleRate.Text
            My.Settings.data51 = Dev2PollingEnable.Checked
            My.Settings.data52 = Dev2removeletters.Checked
            My.Settings.data53 = IgnoreErrors2.Checked
            My.Settings.data54 = Dev2TerminatorEnable.Checked
            My.Settings.data55 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data67 = Dev2STBMask.Text
            My.Settings.data75 = Div1000Dev2.Checked
            My.Settings.data80 = Dev23457Aseven.Checked
            My.Settings.data86 = Dev2TerminatorEnable2.Checked
            My.Settings.data195 = Dev2K2001isolatedata.Checked
            My.Settings.data196 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data219 = Mult1000Dev2.Checked
            My.Settings.data237 = Val(Dev2Timeout.Text)
            My.Settings.data249 = Val(Dev2delayop.Text)
            My.Settings.data295 = txtq2d.Text
            My.Settings.data296 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data297 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data298 = Dev2IntEnable.Checked
            My.Settings.data343 = Dev2Regex.Checked
            My.Settings.data461 = Dev2DecimalNumDPs.Text
            My.Settings.data485 = Dev2SendQuery.Checked
            My.Settings.data517 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 2
        If (ProfDev2_2.Checked = True) Then
            My.Settings.data2b = txtname2.Text
            My.Settings.data4b = txtaddr2.Text
            My.Settings.data6b = CommandStart2.Text
            My.Settings.data18b = CommandStart2run.Text
            My.Settings.data8b = CommandStop2.Text
            My.Settings.data31 = lstIntf2.SelectedIndex
            My.Settings.data10b = Dev2SampleRate.Text
            My.Settings.data56 = Dev2PollingEnable.Checked
            My.Settings.data57 = Dev2removeletters.Checked
            My.Settings.data58 = IgnoreErrors2.Checked
            My.Settings.data59 = Dev2TerminatorEnable.Checked
            My.Settings.data60 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data69 = Dev2STBMask.Text
            My.Settings.data76 = Div1000Dev2.Checked
            My.Settings.data82 = Dev23457Aseven.Checked
            My.Settings.data88 = Dev2TerminatorEnable2.Checked
            My.Settings.data197 = Dev2K2001isolatedata.Checked
            My.Settings.data198 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data220 = Mult1000Dev2.Checked
            My.Settings.data238 = Val(Dev2Timeout.Text)
            My.Settings.data250 = Val(Dev2delayop.Text)
            My.Settings.data299 = txtq2d.Text
            My.Settings.data300 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data301 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data302 = Dev2IntEnable.Checked
            My.Settings.data344 = Dev2Regex.Checked
            My.Settings.data462 = Dev2DecimalNumDPs.Text
            My.Settings.data487 = Dev2SendQuery.Checked
            My.Settings.data518 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 3
        If (ProfDev2_3.Checked = True) Then
            My.Settings.data2c = txtname2.Text
            My.Settings.data4c = txtaddr2.Text
            My.Settings.data6c = CommandStart2.Text
            My.Settings.data18c = CommandStart2run.Text
            My.Settings.data8c = CommandStop2.Text
            My.Settings.data32 = lstIntf2.SelectedIndex
            My.Settings.data10c = Dev2SampleRate.Text
            My.Settings.data61 = Dev2PollingEnable.Checked
            My.Settings.data62 = Dev2removeletters.Checked
            My.Settings.data63 = IgnoreErrors2.Checked
            My.Settings.data64 = Dev2TerminatorEnable.Checked
            My.Settings.data65 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data71 = Dev2STBMask.Text
            My.Settings.data77 = Div1000Dev2.Checked
            My.Settings.data84 = Dev23457Aseven.Checked
            My.Settings.data90 = Dev2TerminatorEnable2.Checked
            My.Settings.data199 = Dev2K2001isolatedata.Checked
            My.Settings.data200 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data221 = Mult1000Dev2.Checked
            My.Settings.data239 = Val(Dev2Timeout.Text)
            My.Settings.data251 = Val(Dev2delayop.Text)
            My.Settings.data303 = txtq2d.Text
            My.Settings.data304 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data305 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data306 = Dev2IntEnable.Checked
            My.Settings.data345 = Dev2Regex.Checked
            My.Settings.data463 = Dev2DecimalNumDPs.Text
            My.Settings.data489 = Dev2SendQuery.Checked
            My.Settings.data519 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 4
        If (ProfDev2_4.Checked = True) Then
            My.Settings.data91 = txtname2.Text
            My.Settings.data92 = CommandStart2.Text
            My.Settings.data93 = CommandStart2run.Text
            My.Settings.data94 = CommandStop2.Text
            My.Settings.data95 = lstIntf2.SelectedIndex
            My.Settings.data96 = txtaddr2.Text
            My.Settings.data97 = Dev2SampleRate.Text
            My.Settings.data98 = Dev2PollingEnable.Checked
            My.Settings.data99 = Dev2removeletters.Checked
            My.Settings.data100 = IgnoreErrors2.Checked
            My.Settings.data101 = Dev2TerminatorEnable.Checked
            My.Settings.data102 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data103 = Dev2STBMask.Text
            My.Settings.data104 = Div1000Dev2.Checked
            My.Settings.data105 = Dev23457Aseven.Checked
            My.Settings.data106 = Dev2TerminatorEnable2.Checked
            My.Settings.data201 = Dev2K2001isolatedata.Checked
            My.Settings.data202 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data222 = Mult1000Dev2.Checked
            My.Settings.data240 = Val(Dev2Timeout.Text)
            My.Settings.data252 = Val(Dev2delayop.Text)
            My.Settings.data307 = txtq2d.Text
            My.Settings.data308 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data309 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data310 = Dev2IntEnable.Checked
            My.Settings.data346 = Dev2Regex.Checked
            My.Settings.data464 = Dev2DecimalNumDPs.Text
            My.Settings.data491 = Dev2SendQuery.Checked
            My.Settings.data520 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 5
        If (ProfDev2_5.Checked = True) Then
            My.Settings.data107 = txtname2.Text
            My.Settings.data108 = CommandStart2.Text
            My.Settings.data109 = CommandStart2run.Text
            My.Settings.data110 = CommandStop2.Text
            My.Settings.data111 = lstIntf2.SelectedIndex
            My.Settings.data112 = txtaddr2.Text
            My.Settings.data113 = Dev2SampleRate.Text
            My.Settings.data114 = Dev2PollingEnable.Checked
            My.Settings.data115 = Dev2removeletters.Checked
            My.Settings.data116 = IgnoreErrors2.Checked
            My.Settings.data117 = Dev2TerminatorEnable.Checked
            My.Settings.data118 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data119 = Dev2STBMask.Text
            My.Settings.data120 = Div1000Dev2.Checked
            My.Settings.data121 = Dev23457Aseven.Checked
            My.Settings.data122 = Dev2TerminatorEnable2.Checked
            My.Settings.data203 = Dev2K2001isolatedata.Checked
            My.Settings.data204 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data223 = Mult1000Dev2.Checked
            My.Settings.data241 = Val(Dev2Timeout.Text)
            My.Settings.data253 = Val(Dev2delayop.Text)
            My.Settings.data311 = txtq2d.Text
            My.Settings.data312 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data313 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data314 = Dev2IntEnable.Checked
            My.Settings.data347 = Dev2Regex.Checked
            My.Settings.data465 = Dev2DecimalNumDPs.Text
            My.Settings.data493 = Dev2SendQuery.Checked
            My.Settings.data521 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 6
        If (ProfDev2_6.Checked = True) Then
            My.Settings.data123 = txtname2.Text
            My.Settings.data124 = CommandStart2.Text
            My.Settings.data125 = CommandStart2run.Text
            My.Settings.data126 = CommandStop2.Text
            My.Settings.data127 = lstIntf2.SelectedIndex
            My.Settings.data128 = txtaddr2.Text
            My.Settings.data129 = Dev2SampleRate.Text
            My.Settings.data130 = Dev2PollingEnable.Checked
            My.Settings.data131 = Dev2removeletters.Checked
            My.Settings.data132 = IgnoreErrors2.Checked
            My.Settings.data133 = Dev2TerminatorEnable.Checked
            My.Settings.data134 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data135 = Dev2STBMask.Text
            My.Settings.data136 = Div1000Dev2.Checked
            My.Settings.data137 = Dev23457Aseven.Checked
            My.Settings.data138 = Dev2TerminatorEnable2.Checked
            My.Settings.data205 = Dev2K2001isolatedata.Checked
            My.Settings.data206 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data224 = Mult1000Dev2.Checked
            My.Settings.data242 = Val(Dev2Timeout.Text)
            My.Settings.data254 = Val(Dev2delayop.Text)
            My.Settings.data315 = txtq2d.Text
            My.Settings.data316 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data317 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data318 = Dev2IntEnable.Checked
            My.Settings.data348 = Dev2Regex.Checked
            My.Settings.data466 = Dev2DecimalNumDPs.Text
            My.Settings.data495 = Dev2SendQuery.Checked
            My.Settings.data522 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 7
        If (ProfDev2_7.Checked = True) Then
            My.Settings.data401 = txtname2.Text
            My.Settings.data402 = CommandStart2.Text
            My.Settings.data403 = CommandStart2run.Text
            My.Settings.data404 = CommandStop2.Text
            My.Settings.data405 = lstIntf2.SelectedIndex
            My.Settings.data406 = txtaddr2.Text
            My.Settings.data407 = Dev2SampleRate.Text
            My.Settings.data408 = Dev2PollingEnable.Checked
            My.Settings.data409 = Dev2removeletters.Checked
            My.Settings.data410 = IgnoreErrors2.Checked
            My.Settings.data411 = Dev2TerminatorEnable.Checked
            My.Settings.data412 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data413 = Dev2STBMask.Text
            My.Settings.data414 = Div1000Dev2.Checked
            My.Settings.data415 = Dev23457Aseven.Checked
            My.Settings.data416 = Dev2TerminatorEnable2.Checked
            My.Settings.data417 = Dev2K2001isolatedata.Checked
            My.Settings.data418 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data419 = Mult1000Dev2.Checked
            My.Settings.data420 = Val(Dev2Timeout.Text)
            My.Settings.data421 = Val(Dev2delayop.Text)
            My.Settings.data422 = txtq2d.Text
            My.Settings.data423 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data424 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data425 = Dev2IntEnable.Checked
            My.Settings.data426 = Dev2Regex.Checked
            My.Settings.data467 = Dev2DecimalNumDPs.Text
            My.Settings.data497 = Dev2SendQuery.Checked
            My.Settings.data523 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 8
        If (ProfDev2_8.Checked = True) Then
            My.Settings.data427 = txtname2.Text
            My.Settings.data428 = CommandStart2.Text
            My.Settings.data429 = CommandStart2run.Text
            My.Settings.data430 = CommandStop2.Text
            My.Settings.data431 = lstIntf2.SelectedIndex
            My.Settings.data432 = txtaddr2.Text
            My.Settings.data433 = Dev2SampleRate.Text
            My.Settings.data434 = Dev2PollingEnable.Checked
            My.Settings.data435 = Dev2removeletters.Checked
            My.Settings.data436 = IgnoreErrors2.Checked
            My.Settings.data437 = Dev2TerminatorEnable.Checked
            My.Settings.data438 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data439 = Dev2STBMask.Text
            My.Settings.data440 = Div1000Dev2.Checked
            My.Settings.data441 = Dev23457Aseven.Checked
            My.Settings.data442 = Dev2TerminatorEnable2.Checked
            My.Settings.data443 = Dev2K2001isolatedata.Checked
            My.Settings.data444 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data445 = Mult1000Dev2.Checked
            My.Settings.data446 = Val(Dev2Timeout.Text)
            My.Settings.data447 = Val(Dev2delayop.Text)
            My.Settings.data448 = txtq2d.Text
            My.Settings.data449 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data450 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data451 = Dev2IntEnable.Checked
            My.Settings.data452 = Dev2Regex.Checked
            My.Settings.data468 = Dev2DecimalNumDPs.Text
            My.Settings.data499 = Dev2SendQuery.Checked
            My.Settings.data524 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 9
        If (ProfDev2_9.Checked = True) Then
            My.Settings.data642 = txtname2.Text
            My.Settings.data643 = CommandStart2.Text
            My.Settings.data644 = CommandStart2run.Text
            My.Settings.data645 = CommandStop2.Text
            My.Settings.data646 = lstIntf2.SelectedIndex
            My.Settings.data647 = txtaddr2.Text
            My.Settings.data648 = Dev2SampleRate.Text
            My.Settings.data649 = Dev2PollingEnable.Checked
            My.Settings.data650 = Dev2removeletters.Checked
            My.Settings.data651 = IgnoreErrors2.Checked
            My.Settings.data652 = Dev2TerminatorEnable.Checked
            My.Settings.data653 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data654 = Dev2STBMask.Text
            My.Settings.data655 = Div1000Dev2.Checked
            My.Settings.data656 = Dev23457Aseven.Checked
            My.Settings.data657 = Dev2TerminatorEnable2.Checked
            My.Settings.data658 = Dev2K2001isolatedata.Checked
            My.Settings.data659 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data660 = Mult1000Dev2.Checked
            My.Settings.data661 = Val(Dev2Timeout.Text)
            My.Settings.data662 = Val(Dev2delayop.Text)
            My.Settings.data663 = txtq2d.Text
            My.Settings.data664 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data665 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data666 = Dev2IntEnable.Checked
            My.Settings.data667 = Dev2Regex.Checked
            My.Settings.data668 = Dev2DecimalNumDPs.Text
            My.Settings.data669 = Dev2IntEnable.Checked
            My.Settings.data670 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 10
        If (ProfDev2_10.Checked = True) Then
            My.Settings.data671 = txtname2.Text
            My.Settings.data672 = CommandStart2.Text
            My.Settings.data673 = CommandStart2run.Text
            My.Settings.data674 = CommandStop2.Text
            My.Settings.data675 = lstIntf2.SelectedIndex
            My.Settings.data676 = txtaddr2.Text
            My.Settings.data677 = Dev2SampleRate.Text
            My.Settings.data678 = Dev2PollingEnable.Checked
            My.Settings.data679 = Dev2removeletters.Checked
            My.Settings.data680 = IgnoreErrors2.Checked
            My.Settings.data681 = Dev2TerminatorEnable.Checked
            My.Settings.data682 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data683 = Dev2STBMask.Text
            My.Settings.data684 = Div1000Dev2.Checked
            My.Settings.data685 = Dev23457Aseven.Checked
            My.Settings.data686 = Dev2TerminatorEnable2.Checked
            My.Settings.data687 = Dev2K2001isolatedata.Checked
            My.Settings.data688 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data689 = Mult1000Dev2.Checked
            My.Settings.data690 = Val(Dev2Timeout.Text)
            My.Settings.data691 = Val(Dev2delayop.Text)
            My.Settings.data692 = txtq2d.Text
            My.Settings.data693 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data694 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data695 = Dev2IntEnable.Checked
            My.Settings.data696 = Dev2Regex.Checked
            My.Settings.data697 = Dev2DecimalNumDPs.Text
            My.Settings.data698 = Dev2IntEnable.Checked
            My.Settings.data699 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 11
        If (ProfDev2_11.Checked = True) Then
            My.Settings.data700 = txtname2.Text
            My.Settings.data701 = CommandStart2.Text
            My.Settings.data702 = CommandStart2run.Text
            My.Settings.data703 = CommandStop2.Text
            My.Settings.data704 = lstIntf2.SelectedIndex
            My.Settings.data705 = txtaddr2.Text
            My.Settings.data706 = Dev2SampleRate.Text
            My.Settings.data707 = Dev2PollingEnable.Checked
            My.Settings.data708 = Dev2removeletters.Checked
            My.Settings.data709 = IgnoreErrors2.Checked
            My.Settings.data710 = Dev2TerminatorEnable.Checked
            My.Settings.data711 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data712 = Dev2STBMask.Text
            My.Settings.data713 = Div1000Dev2.Checked
            My.Settings.data714 = Dev23457Aseven.Checked
            My.Settings.data715 = Dev2TerminatorEnable2.Checked
            My.Settings.data716 = Dev2K2001isolatedata.Checked
            My.Settings.data717 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data718 = Mult1000Dev2.Checked
            My.Settings.data719 = Val(Dev2Timeout.Text)
            My.Settings.data720 = Val(Dev2delayop.Text)
            My.Settings.data721 = txtq2d.Text
            My.Settings.data722 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data723 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data724 = Dev2IntEnable.Checked
            My.Settings.data725 = Dev2Regex.Checked
            My.Settings.data726 = Dev2DecimalNumDPs.Text
            My.Settings.data727 = Dev2IntEnable.Checked
            My.Settings.data728 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 12
        If (ProfDev2_12.Checked = True) Then
            My.Settings.data729 = txtname2.Text
            My.Settings.data730 = CommandStart2.Text
            My.Settings.data731 = CommandStart2run.Text
            My.Settings.data732 = CommandStop2.Text
            My.Settings.data733 = lstIntf2.SelectedIndex
            My.Settings.data734 = txtaddr2.Text
            My.Settings.data735 = Dev2SampleRate.Text
            My.Settings.data736 = Dev2PollingEnable.Checked
            My.Settings.data737 = Dev2removeletters.Checked
            My.Settings.data738 = IgnoreErrors2.Checked
            My.Settings.data739 = Dev2TerminatorEnable.Checked
            My.Settings.data740 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data741 = Dev2STBMask.Text
            My.Settings.data742 = Div1000Dev2.Checked
            My.Settings.data743 = Dev23457Aseven.Checked
            My.Settings.data744 = Dev2TerminatorEnable2.Checked
            My.Settings.data745 = Dev2K2001isolatedata.Checked
            My.Settings.data746 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data747 = Mult1000Dev2.Checked
            My.Settings.data748 = Val(Dev2Timeout.Text)
            My.Settings.data749 = Val(Dev2delayop.Text)
            My.Settings.data750 = txtq2d.Text
            My.Settings.data751 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data752 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data753 = Dev2IntEnable.Checked
            My.Settings.data754 = Dev2Regex.Checked
            My.Settings.data755 = Dev2DecimalNumDPs.Text
            My.Settings.data756 = Dev2IntEnable.Checked
            My.Settings.data757 = txtOperationDev2.Text
        End If


        ' Save the settings to persist the changes
        My.Settings.Save()

        If CSVdelimiterComma.Checked = True Then
            CSVdelimit = ","
            My.Settings.data29 = ","
        Else
            CSVdelimit = ";"
            My.Settings.data29 = ";"
        End If
    End Sub

    ' Device 1 profile 1
    Private Sub ProfDev1_1_Click(sender As Object, e As EventArgs) Handles ProfDev1_1.Click

        If (ProfDev1_1.Checked = True) Then
            'ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = True
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        End If

        txtname1.Text = My.Settings.data1
        CommandStart1.Text = My.Settings.data5
        CommandStart1run.Text = My.Settings.data17
        CommandStop1.Text = My.Settings.data7
        lstIntf1.SelectedIndex = My.Settings.data33
        txtaddr1.Text = My.Settings.data3
        Dev1SampleRate.Text = My.Settings.data9

        Dev1PollingEnable.Checked = My.Settings.data36
        Dev1removeletters.Checked = My.Settings.data37
        IgnoreErrors1.Checked = My.Settings.data38
        Dev1TerminatorEnable.Checked = My.Settings.data39
        CheckBoxSendBlockingDev1.Checked = My.Settings.data40

        Dev1STBMask.Text = My.Settings.data66

        Div1000Dev1.Checked = My.Settings.data72

        Dev13457Aseven.Checked = My.Settings.data79

        Dev1TerminatorEnable2.Checked = My.Settings.data85

        Dev1K2001isolatedata.Checked = My.Settings.data207
        Dev1K2001isolatedataCHAR.Text = My.Settings.data208

        Mult1000Dev1.Checked = My.Settings.data225

        Dev1Timeout.Text = My.Settings.data231

        Dev1delayop.Text = My.Settings.data243

        txtq1d.Text = My.Settings.data271
        Dev1pauseDurationInSeconds.Text = My.Settings.data272
        Dev1runStopwatchEveryInMins.Text = My.Settings.data273
        Dev1IntEnable.Checked = My.Settings.data274
        Dev1Regex.Checked = My.Settings.data337

        Dev1DecimalNumDPs.Text = My.Settings.data453

        Dev1IntEnable.Checked = My.Settings.data484

        txtOperationDev1.Text = My.Settings.data509


    End Sub

    ' Device 1 profile 2
    Private Sub ProfDev1_2_Click(sender As Object, e As EventArgs) Handles ProfDev1_2.Click

        If (ProfDev1_2.Checked = True) Then
            ProfDev1_1.Checked = False
            'ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = True
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        End If

        txtname1.Text = My.Settings.data1b
        CommandStart1.Text = My.Settings.data5b
        CommandStart1run.Text = My.Settings.data17b
        CommandStop1.Text = My.Settings.data7b
        lstIntf1.SelectedIndex = My.Settings.data34
        txtaddr1.Text = My.Settings.data3b
        Dev1SampleRate.Text = My.Settings.data9b

        Dev1PollingEnable.Checked = My.Settings.data41
        Dev1removeletters.Checked = My.Settings.data42
        IgnoreErrors1.Checked = My.Settings.data43
        Dev1TerminatorEnable.Checked = My.Settings.data44
        CheckBoxSendBlockingDev1.Checked = My.Settings.data45

        Dev1STBMask.Text = My.Settings.data68

        Div1000Dev1.Checked = My.Settings.data73

        Dev13457Aseven.Checked = My.Settings.data81

        Dev1TerminatorEnable2.Checked = My.Settings.data87

        Dev1K2001isolatedata.Checked = My.Settings.data209
        Dev1K2001isolatedataCHAR.Text = My.Settings.data210

        Mult1000Dev1.Checked = My.Settings.data226

        Dev1Timeout.Text = My.Settings.data232

        Dev1delayop.Text = My.Settings.data244

        txtq1d.Text = My.Settings.data275
        Dev1pauseDurationInSeconds.Text = My.Settings.data276
        Dev1runStopwatchEveryInMins.Text = My.Settings.data277
        Dev1IntEnable.Checked = My.Settings.data278
        Dev1Regex.Checked = My.Settings.data338

        Dev1DecimalNumDPs.Text = My.Settings.data454

        Dev1IntEnable.Checked = My.Settings.data486

        txtOperationDev1.Text = My.Settings.data510

    End Sub

    ' Device 1 profile 3
    Private Sub ProfDev1_3_Click(sender As Object, e As EventArgs) Handles ProfDev1_3.Click

        If (ProfDev1_3.Checked = True) Then
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            'ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = True
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        End If

        txtname1.Text = My.Settings.data1c
        CommandStart1.Text = My.Settings.data5c
        CommandStart1run.Text = My.Settings.data17c
        CommandStop1.Text = My.Settings.data7c
        lstIntf1.SelectedIndex = My.Settings.data35
        txtaddr1.Text = My.Settings.data3c
        Dev1SampleRate.Text = My.Settings.data9c

        Dev1PollingEnable.Checked = My.Settings.data46
        Dev1removeletters.Checked = My.Settings.data47
        IgnoreErrors1.Checked = My.Settings.data48
        Dev1TerminatorEnable.Checked = My.Settings.data49
        CheckBoxSendBlockingDev1.Checked = My.Settings.data50

        Dev1STBMask.Text = My.Settings.data70

        Div1000Dev1.Checked = My.Settings.data74

        Dev13457Aseven.Checked = My.Settings.data83

        Dev1TerminatorEnable2.Checked = My.Settings.data89

        Dev1K2001isolatedata.Checked = My.Settings.data211
        Dev1K2001isolatedataCHAR.Text = My.Settings.data212

        Mult1000Dev1.Checked = My.Settings.data227

        Dev1Timeout.Text = My.Settings.data233

        Dev1delayop.Text = My.Settings.data245

        txtq1d.Text = My.Settings.data279
        Dev1pauseDurationInSeconds.Text = My.Settings.data280
        Dev1runStopwatchEveryInMins.Text = My.Settings.data281
        Dev1IntEnable.Checked = My.Settings.data282
        Dev1Regex.Checked = My.Settings.data339

        Dev1DecimalNumDPs.Text = My.Settings.data455

        Dev1IntEnable.Checked = My.Settings.data488

        txtOperationDev1.Text = My.Settings.data511

    End Sub

    ' Device 1 profile 4
    Private Sub ProfDev1_4_Click(sender As Object, e As EventArgs) Handles ProfDev1_4.Click

        If (ProfDev1_4.Checked = True) Then
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            'ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = True
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        End If

        txtname1.Text = My.Settings.data139
        CommandStart1.Text = My.Settings.data140
        CommandStart1run.Text = My.Settings.data141
        CommandStop1.Text = My.Settings.data142
        lstIntf1.SelectedIndex = My.Settings.data143
        txtaddr1.Text = My.Settings.data144
        Dev1SampleRate.Text = My.Settings.data145
        Dev1PollingEnable.Checked = My.Settings.data146
        Dev1removeletters.Checked = My.Settings.data147
        IgnoreErrors1.Checked = My.Settings.data148
        Dev1TerminatorEnable.Checked = My.Settings.data149
        CheckBoxSendBlockingDev1.Checked = My.Settings.data150
        Dev1STBMask.Text = My.Settings.data151
        Div1000Dev1.Checked = My.Settings.data152
        Dev13457Aseven.Checked = My.Settings.data153
        Dev1TerminatorEnable2.Checked = My.Settings.data154

        Dev1K2001isolatedata.Checked = My.Settings.data213
        Dev1K2001isolatedataCHAR.Text = My.Settings.data214

        Mult1000Dev1.Checked = My.Settings.data228

        Dev1Timeout.Text = My.Settings.data234

        Dev1delayop.Text = My.Settings.data246

        txtq1d.Text = My.Settings.data283
        Dev1pauseDurationInSeconds.Text = My.Settings.data284
        Dev1runStopwatchEveryInMins.Text = My.Settings.data285
        Dev1IntEnable.Checked = My.Settings.data286
        Dev1Regex.Checked = My.Settings.data340

        Dev1DecimalNumDPs.Text = My.Settings.data456

        Dev1IntEnable.Checked = My.Settings.data490

        txtOperationDev1.Text = My.Settings.data512

    End Sub

    ' Device 1 profile 5
    Private Sub ProfDev1_5_Click(sender As Object, e As EventArgs) Handles ProfDev1_5.Click

        If (ProfDev1_5.Checked = True) Then
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            'ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = True
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        End If

        txtname1.Text = My.Settings.data155
        CommandStart1.Text = My.Settings.data156
        CommandStart1run.Text = My.Settings.data157
        CommandStop1.Text = My.Settings.data158
        lstIntf1.SelectedIndex = My.Settings.data159
        txtaddr1.Text = My.Settings.data160
        Dev1SampleRate.Text = My.Settings.data161
        Dev1PollingEnable.Checked = My.Settings.data162
        Dev1removeletters.Checked = My.Settings.data163
        IgnoreErrors1.Checked = My.Settings.data164
        Dev1TerminatorEnable.Checked = My.Settings.data165
        CheckBoxSendBlockingDev1.Checked = My.Settings.data166
        Dev1STBMask.Text = My.Settings.data167
        Div1000Dev1.Checked = My.Settings.data168
        Dev13457Aseven.Checked = My.Settings.data169
        Dev1TerminatorEnable2.Checked = My.Settings.data170

        Dev1K2001isolatedata.Checked = My.Settings.data215
        Dev1K2001isolatedataCHAR.Text = My.Settings.data216

        Mult1000Dev1.Checked = My.Settings.data229

        Dev1Timeout.Text = My.Settings.data235

        Dev1delayop.Text = My.Settings.data247

        txtq1d.Text = My.Settings.data287
        Dev1pauseDurationInSeconds.Text = My.Settings.data288
        Dev1runStopwatchEveryInMins.Text = My.Settings.data289
        Dev1IntEnable.Checked = My.Settings.data290
        Dev1Regex.Checked = My.Settings.data341

        Dev1DecimalNumDPs.Text = My.Settings.data457

        Dev1IntEnable.Checked = My.Settings.data492

        txtOperationDev1.Text = My.Settings.data513

    End Sub

    ' Device 1 profile 6
    Private Sub ProfDev1_6_Click(sender As Object, e As EventArgs) Handles ProfDev1_6.Click

        If (ProfDev1_6.Checked = True) Then
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            'ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = True
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        End If

        txtname1.Text = My.Settings.data171
        CommandStart1.Text = My.Settings.data172
        CommandStart1run.Text = My.Settings.data173
        CommandStop1.Text = My.Settings.data174
        lstIntf1.SelectedIndex = My.Settings.data175
        txtaddr1.Text = My.Settings.data176
        Dev1SampleRate.Text = My.Settings.data177
        Dev1PollingEnable.Checked = My.Settings.data178
        Dev1removeletters.Checked = My.Settings.data179
        IgnoreErrors1.Checked = My.Settings.data180
        Dev1TerminatorEnable.Checked = My.Settings.data181
        CheckBoxSendBlockingDev1.Checked = My.Settings.data182
        Dev1STBMask.Text = My.Settings.data183
        Div1000Dev1.Checked = My.Settings.data184
        Dev13457Aseven.Checked = My.Settings.data185
        Dev1TerminatorEnable2.Checked = My.Settings.data186

        Dev1K2001isolatedata.Checked = My.Settings.data217
        Dev1K2001isolatedataCHAR.Text = My.Settings.data218

        Mult1000Dev1.Checked = My.Settings.data230

        Dev1Timeout.Text = My.Settings.data236

        Dev1delayop.Text = My.Settings.data248

        txtq1d.Text = My.Settings.data291
        Dev1pauseDurationInSeconds.Text = My.Settings.data292
        Dev1runStopwatchEveryInMins.Text = My.Settings.data293
        Dev1IntEnable.Checked = My.Settings.data294
        Dev1Regex.Checked = My.Settings.data342

        Dev1DecimalNumDPs.Text = My.Settings.data458

        Dev1IntEnable.Checked = My.Settings.data494

        txtOperationDev1.Text = My.Settings.data514

    End Sub

    ' Device 1 profile 7
    Private Sub ProfDev1_7_Click(sender As Object, e As EventArgs) Handles ProfDev1_7.Click

        If (ProfDev1_7.Checked = True) Then
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            'ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = True
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        End If

        txtname1.Text = My.Settings.data349
        CommandStart1.Text = My.Settings.data350
        CommandStart1run.Text = My.Settings.data351
        CommandStop1.Text = My.Settings.data352
        lstIntf1.SelectedIndex = My.Settings.data353
        txtaddr1.Text = My.Settings.data354
        Dev1SampleRate.Text = My.Settings.data355
        Dev1PollingEnable.Checked = My.Settings.data356
        Dev1removeletters.Checked = My.Settings.data357
        IgnoreErrors1.Checked = My.Settings.data358
        Dev1TerminatorEnable.Checked = My.Settings.data359
        CheckBoxSendBlockingDev1.Checked = My.Settings.data360
        Dev1STBMask.Text = My.Settings.data361
        Div1000Dev1.Checked = My.Settings.data362
        Dev13457Aseven.Checked = My.Settings.data363
        Dev1TerminatorEnable2.Checked = My.Settings.data364

        Dev1K2001isolatedata.Checked = My.Settings.data365
        Dev1K2001isolatedataCHAR.Text = My.Settings.data366

        Mult1000Dev1.Checked = My.Settings.data367

        Dev1Timeout.Text = My.Settings.data368

        Dev1delayop.Text = My.Settings.data369

        txtq1d.Text = My.Settings.data370
        Dev1pauseDurationInSeconds.Text = My.Settings.data371
        Dev1runStopwatchEveryInMins.Text = My.Settings.data372
        Dev1IntEnable.Checked = My.Settings.data373
        Dev1Regex.Checked = My.Settings.data374
        Div1000Dev2.Checked = My.Settings.data362

        Dev1DecimalNumDPs.Text = My.Settings.data459

        Dev1IntEnable.Checked = My.Settings.data496

        txtOperationDev1.Text = My.Settings.data515

    End Sub


    ' Device 1 profile 8
    Private Sub ProfDev1_8_Click(sender As Object, e As EventArgs) Handles ProfDev1_8.Click

        If (ProfDev1_8.Checked = True) Then
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            'ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = True
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        End If

        txtname1.Text = My.Settings.data375
        CommandStart1.Text = My.Settings.data376
        CommandStart1run.Text = My.Settings.data377
        CommandStop1.Text = My.Settings.data378
        lstIntf1.SelectedIndex = My.Settings.data379
        txtaddr1.Text = My.Settings.data380
        Dev1SampleRate.Text = My.Settings.data381
        Dev1PollingEnable.Checked = My.Settings.data382
        Dev1removeletters.Checked = My.Settings.data383
        IgnoreErrors1.Checked = My.Settings.data384
        Dev1TerminatorEnable.Checked = My.Settings.data385
        CheckBoxSendBlockingDev1.Checked = My.Settings.data386
        Dev1STBMask.Text = My.Settings.data387
        Div1000Dev1.Checked = My.Settings.data388
        Dev13457Aseven.Checked = My.Settings.data389
        Dev1TerminatorEnable2.Checked = My.Settings.data390

        Dev1K2001isolatedata.Checked = My.Settings.data391
        Dev1K2001isolatedataCHAR.Text = My.Settings.data392

        Mult1000Dev1.Checked = My.Settings.data393

        Dev1Timeout.Text = My.Settings.data394

        Dev1delayop.Text = My.Settings.data395

        txtq1d.Text = My.Settings.data396
        Dev1pauseDurationInSeconds.Text = My.Settings.data397
        Dev1runStopwatchEveryInMins.Text = My.Settings.data398
        Dev1IntEnable.Checked = My.Settings.data399
        Dev1Regex.Checked = My.Settings.data400

        Dev1DecimalNumDPs.Text = My.Settings.data460

        Dev1IntEnable.Checked = My.Settings.data498

        txtOperationDev1.Text = My.Settings.data516

    End Sub

    Private Sub ProfDev1_9_Click(sender As Object, e As EventArgs) Handles ProfDev1_9.Click

        If (ProfDev1_9.Checked = True) Then
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False

            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = True
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        End If

        txtname1.Text = My.Settings.data526
        CommandStart1.Text = My.Settings.data527
        CommandStart1run.Text = My.Settings.data528
        CommandStop1.Text = My.Settings.data529
        lstIntf1.SelectedIndex = My.Settings.data530
        txtaddr1.Text = My.Settings.data531
        Dev1SampleRate.Text = My.Settings.data532
        Dev1PollingEnable.Checked = My.Settings.data533
        Dev1removeletters.Checked = My.Settings.data534
        IgnoreErrors1.Checked = My.Settings.data535
        Dev1TerminatorEnable.Checked = My.Settings.data536
        CheckBoxSendBlockingDev1.Checked = My.Settings.data537
        Dev1STBMask.Text = My.Settings.data538
        Div1000Dev1.Checked = My.Settings.data539
        Dev13457Aseven.Checked = My.Settings.data540
        Dev1TerminatorEnable2.Checked = My.Settings.data541
        Dev1K2001isolatedata.Checked = My.Settings.data542
        Dev1K2001isolatedataCHAR.Text = My.Settings.data543
        Mult1000Dev1.Checked = My.Settings.data544
        Dev1Timeout.Text = My.Settings.data545
        Dev1delayop.Text = My.Settings.data546
        txtq1d.Text = My.Settings.data547
        Dev1pauseDurationInSeconds.Text = My.Settings.data548
        Dev1runStopwatchEveryInMins.Text = My.Settings.data549
        Dev1IntEnable.Checked = My.Settings.data550
        Dev1Regex.Checked = My.Settings.data551
        Dev1DecimalNumDPs.Text = My.Settings.data552
        txtOperationDev1.Text = My.Settings.data554

    End Sub

    Private Sub ProfDev1_10_Click(sender As Object, e As EventArgs) Handles ProfDev1_10.Click

        If (ProfDev1_10.Checked = True) Then
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False

            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = True
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = False
        End If

        txtname1.Text = My.Settings.data555
        CommandStart1.Text = My.Settings.data556
        CommandStart1run.Text = My.Settings.data557
        CommandStop1.Text = My.Settings.data558
        lstIntf1.SelectedIndex = My.Settings.data559
        txtaddr1.Text = My.Settings.data560
        Dev1SampleRate.Text = My.Settings.data561
        Dev1PollingEnable.Checked = My.Settings.data562
        Dev1removeletters.Checked = My.Settings.data563
        IgnoreErrors1.Checked = My.Settings.data564
        Dev1TerminatorEnable.Checked = My.Settings.data565
        CheckBoxSendBlockingDev1.Checked = My.Settings.data566
        Dev1STBMask.Text = My.Settings.data567
        Div1000Dev1.Checked = My.Settings.data568
        Dev13457Aseven.Checked = My.Settings.data569
        Dev1TerminatorEnable2.Checked = My.Settings.data570
        Dev1K2001isolatedata.Checked = My.Settings.data571
        Dev1K2001isolatedataCHAR.Text = My.Settings.data572
        Mult1000Dev1.Checked = My.Settings.data573
        Dev1Timeout.Text = My.Settings.data574
        Dev1delayop.Text = My.Settings.data575
        txtq1d.Text = My.Settings.data576
        Dev1pauseDurationInSeconds.Text = My.Settings.data577
        Dev1runStopwatchEveryInMins.Text = My.Settings.data578
        Dev1IntEnable.Checked = My.Settings.data579
        Dev1Regex.Checked = My.Settings.data580
        Dev1DecimalNumDPs.Text = My.Settings.data581
        txtOperationDev1.Text = My.Settings.data583

    End Sub

    Private Sub ProfDev1_11_Click(sender As Object, e As EventArgs) Handles ProfDev1_11.Click

        If (ProfDev1_11.Checked = True) Then
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False

            ProfDev1_12.Checked = False
        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_12.Checked = False
            ProfDev1_11.Checked = True
        End If

        txtname1.Text = My.Settings.data584
        CommandStart1.Text = My.Settings.data585
        CommandStart1run.Text = My.Settings.data586
        CommandStop1.Text = My.Settings.data587
        lstIntf1.SelectedIndex = My.Settings.data588
        txtaddr1.Text = My.Settings.data589
        Dev1SampleRate.Text = My.Settings.data590
        Dev1PollingEnable.Checked = My.Settings.data591
        Dev1removeletters.Checked = My.Settings.data592
        IgnoreErrors1.Checked = My.Settings.data593
        Dev1TerminatorEnable.Checked = My.Settings.data594
        CheckBoxSendBlockingDev1.Checked = My.Settings.data595
        Dev1STBMask.Text = My.Settings.data596
        Div1000Dev1.Checked = My.Settings.data597
        Dev13457Aseven.Checked = My.Settings.data598
        Dev1TerminatorEnable2.Checked = My.Settings.data599
        Dev1K2001isolatedata.Checked = My.Settings.data600
        Dev1K2001isolatedataCHAR.Text = My.Settings.data601
        Mult1000Dev1.Checked = My.Settings.data602
        Dev1Timeout.Text = My.Settings.data603
        Dev1delayop.Text = My.Settings.data604
        txtq1d.Text = My.Settings.data605
        Dev1pauseDurationInSeconds.Text = My.Settings.data606
        Dev1runStopwatchEveryInMins.Text = My.Settings.data607
        Dev1IntEnable.Checked = My.Settings.data608
        Dev1Regex.Checked = My.Settings.data609
        Dev1DecimalNumDPs.Text = My.Settings.data610
        txtOperationDev1.Text = My.Settings.data612

    End Sub

    Private Sub ProfDev1_12_Click(sender As Object, e As EventArgs) Handles ProfDev1_12.Click

        If (ProfDev1_12.Checked = True) Then
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False

        Else
            ProfDev1_1.Checked = False
            ProfDev1_2.Checked = False
            ProfDev1_3.Checked = False
            ProfDev1_4.Checked = False
            ProfDev1_5.Checked = False
            ProfDev1_6.Checked = False
            ProfDev1_7.Checked = False
            ProfDev1_8.Checked = False
            ProfDev1_9.Checked = False
            ProfDev1_10.Checked = False
            ProfDev1_11.Checked = False
            ProfDev1_12.Checked = True
        End If

        txtname1.Text = My.Settings.data613
        CommandStart1.Text = My.Settings.data614
        CommandStart1run.Text = My.Settings.data615
        CommandStop1.Text = My.Settings.data616
        lstIntf1.SelectedIndex = My.Settings.data617
        txtaddr1.Text = My.Settings.data618
        Dev1SampleRate.Text = My.Settings.data619
        Dev1PollingEnable.Checked = My.Settings.data620
        Dev1removeletters.Checked = My.Settings.data621
        IgnoreErrors1.Checked = My.Settings.data622
        Dev1TerminatorEnable.Checked = My.Settings.data623
        CheckBoxSendBlockingDev1.Checked = My.Settings.data624
        Dev1STBMask.Text = My.Settings.data625
        Div1000Dev1.Checked = My.Settings.data626
        Dev13457Aseven.Checked = My.Settings.data627
        Dev1TerminatorEnable2.Checked = My.Settings.data628
        Dev1K2001isolatedata.Checked = My.Settings.data629
        Dev1K2001isolatedataCHAR.Text = My.Settings.data630
        Mult1000Dev1.Checked = My.Settings.data631
        Dev1Timeout.Text = My.Settings.data632
        Dev1delayop.Text = My.Settings.data633
        txtq1d.Text = My.Settings.data634
        Dev1pauseDurationInSeconds.Text = My.Settings.data635
        Dev1runStopwatchEveryInMins.Text = My.Settings.data636
        Dev1IntEnable.Checked = My.Settings.data637
        Dev1Regex.Checked = My.Settings.data638
        Dev1DecimalNumDPs.Text = My.Settings.data639
        txtOperationDev1.Text = My.Settings.data641

    End Sub


    ' Device 2 profile 1
    Private Sub ProfDev2_1_Click(sender As Object, e As EventArgs) Handles ProfDev2_1.Click

        If (ProfDev2_1.Checked = True) Then

            'ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = True
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data2
        CommandStart2.Text = My.Settings.data6
        CommandStart2run.Text = My.Settings.data18
        CommandStop2.Text = My.Settings.data8
        lstIntf2.SelectedIndex = My.Settings.data30
        txtaddr2.Text = My.Settings.data4
        Dev2SampleRate.Text = My.Settings.data10

        Dev2PollingEnable.Checked = My.Settings.data51
        Dev2removeletters.Checked = My.Settings.data52
        IgnoreErrors2.Checked = My.Settings.data53
        Dev2TerminatorEnable.Checked = My.Settings.data54
        CheckBoxSendBlockingDev2.Checked = My.Settings.data55

        Dev2STBMask.Text = My.Settings.data67

        Div1000Dev2.Checked = My.Settings.data75

        Dev23457Aseven.Checked = My.Settings.data80

        Dev2TerminatorEnable2.Checked = My.Settings.data86

        Dev2K2001isolatedata.Checked = My.Settings.data195
        Dev2K2001isolatedataCHAR.Text = My.Settings.data196

        Mult1000Dev2.Checked = My.Settings.data219

        Dev2Timeout.Text = My.Settings.data237

        Dev2delayop.Text = My.Settings.data249

        txtq2d.Text = My.Settings.data295
        Dev2pauseDurationInSeconds.Text = My.Settings.data296
        Dev2runStopwatchEveryInMins.Text = My.Settings.data297
        Dev2IntEnable.Checked = My.Settings.data298
        Dev2Regex.Checked = My.Settings.data343

        Dev2DecimalNumDPs.Text = My.Settings.data461

        Dev2SendQuery.Checked = My.Settings.data485

        txtOperationDev2.Text = My.Settings.data517

    End Sub

    ' Device 2 profile 2
    Private Sub ProfDev2_2_Click(sender As Object, e As EventArgs) Handles ProfDev2_2.Click

        If (ProfDev2_2.Checked = True) Then
            ProfDev2_1.Checked = False
            'ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = True
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data2b
        CommandStart2.Text = My.Settings.data6b
        CommandStart2run.Text = My.Settings.data18b
        CommandStop2.Text = My.Settings.data8b
        lstIntf2.SelectedIndex = My.Settings.data31
        txtaddr2.Text = My.Settings.data4b
        Dev2SampleRate.Text = My.Settings.data10b

        Dev2PollingEnable.Checked = My.Settings.data56
        Dev2removeletters.Checked = My.Settings.data57
        IgnoreErrors2.Checked = My.Settings.data58
        Dev2TerminatorEnable.Checked = My.Settings.data59
        CheckBoxSendBlockingDev2.Checked = My.Settings.data60

        Dev2STBMask.Text = My.Settings.data69

        Div1000Dev2.Checked = My.Settings.data76

        Dev23457Aseven.Checked = My.Settings.data82

        Dev2TerminatorEnable2.Checked = My.Settings.data88

        Dev2K2001isolatedata.Checked = My.Settings.data197
        Dev2K2001isolatedataCHAR.Text = My.Settings.data198

        Mult1000Dev2.Checked = My.Settings.data220

        Dev2Timeout.Text = My.Settings.data238

        Dev2delayop.Text = My.Settings.data250

        txtq2d.Text = My.Settings.data299
        Dev2pauseDurationInSeconds.Text = My.Settings.data300
        Dev2runStopwatchEveryInMins.Text = My.Settings.data301
        Dev2IntEnable.Checked = My.Settings.data302
        Dev2Regex.Checked = My.Settings.data344

        Dev2DecimalNumDPs.Text = My.Settings.data462

        Dev2SendQuery.Checked = My.Settings.data487

        txtOperationDev2.Text = My.Settings.data518

    End Sub

    ' Device 2 profile 3
    Private Sub ProfDev2_3_Click(sender As Object, e As EventArgs) Handles ProfDev2_3.Click

        If (ProfDev2_3.Checked = True) Then
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            'ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = True
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data2c
        CommandStart2.Text = My.Settings.data6c
        CommandStart2run.Text = My.Settings.data18c
        CommandStop2.Text = My.Settings.data8c
        lstIntf2.SelectedIndex = My.Settings.data32
        txtaddr2.Text = My.Settings.data4c
        Dev2SampleRate.Text = My.Settings.data10c

        Dev2PollingEnable.Checked = My.Settings.data61
        Dev2removeletters.Checked = My.Settings.data62
        IgnoreErrors2.Checked = My.Settings.data63
        Dev2TerminatorEnable.Checked = My.Settings.data64
        CheckBoxSendBlockingDev2.Checked = My.Settings.data65

        Dev2STBMask.Text = My.Settings.data71

        Div1000Dev2.Checked = My.Settings.data77
        Dev23457Aseven.Checked = My.Settings.data84

        Dev2TerminatorEnable2.Checked = My.Settings.data90

        Dev2K2001isolatedata.Checked = My.Settings.data199
        Dev2K2001isolatedataCHAR.Text = My.Settings.data200

        Mult1000Dev2.Checked = My.Settings.data221

        Dev2Timeout.Text = My.Settings.data239

        Dev2delayop.Text = My.Settings.data251

        txtq2d.Text = My.Settings.data303
        Dev2pauseDurationInSeconds.Text = My.Settings.data304
        Dev2runStopwatchEveryInMins.Text = My.Settings.data305
        Dev2IntEnable.Checked = My.Settings.data306
        Dev2Regex.Checked = My.Settings.data345

        Dev2DecimalNumDPs.Text = My.Settings.data463

        Dev2SendQuery.Checked = My.Settings.data489

        txtOperationDev2.Text = My.Settings.data519

    End Sub

    ' Device 2 profile 4
    Private Sub ProfDev2_4_Click(sender As Object, e As EventArgs) Handles ProfDev2_4.Click

        If (ProfDev2_4.Checked = True) Then
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            'ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = True
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data91
        CommandStart2.Text = My.Settings.data92
        CommandStart2run.Text = My.Settings.data93
        CommandStop2.Text = My.Settings.data94
        lstIntf2.SelectedIndex = My.Settings.data95
        txtaddr2.Text = My.Settings.data96
        Dev2SampleRate.Text = My.Settings.data97

        Dev2PollingEnable.Checked = My.Settings.data98
        Dev2removeletters.Checked = My.Settings.data99
        IgnoreErrors2.Checked = My.Settings.data100
        Dev2TerminatorEnable.Checked = My.Settings.data101
        CheckBoxSendBlockingDev2.Checked = My.Settings.data102
        Dev2STBMask.Text = My.Settings.data103
        Div1000Dev2.Checked = My.Settings.data104
        Dev23457Aseven.Checked = My.Settings.data105
        Dev2TerminatorEnable2.Checked = My.Settings.data106

        Dev2K2001isolatedata.Checked = My.Settings.data201
        Dev2K2001isolatedataCHAR.Text = My.Settings.data202

        Mult1000Dev2.Checked = My.Settings.data222

        Dev2Timeout.Text = My.Settings.data240

        Dev2delayop.Text = My.Settings.data252

        txtq2d.Text = My.Settings.data307
        Dev2pauseDurationInSeconds.Text = My.Settings.data308
        Dev2runStopwatchEveryInMins.Text = My.Settings.data309
        Dev2IntEnable.Checked = My.Settings.data310
        Dev2Regex.Checked = My.Settings.data346

        Dev2DecimalNumDPs.Text = My.Settings.data464

        Dev2SendQuery.Checked = My.Settings.data491

        txtOperationDev2.Text = My.Settings.data520

    End Sub

    ' Device 2 profile 5
    Private Sub ProfDev2_5_Click(sender As Object, e As EventArgs) Handles ProfDev2_5.Click

        If (ProfDev2_5.Checked = True) Then
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            'ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = True
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data107
        CommandStart2.Text = My.Settings.data108
        CommandStart2run.Text = My.Settings.data109
        CommandStop2.Text = My.Settings.data110
        lstIntf2.SelectedIndex = My.Settings.data111
        txtaddr2.Text = My.Settings.data112
        Dev2SampleRate.Text = My.Settings.data113
        Dev2PollingEnable.Checked = My.Settings.data114
        Dev2removeletters.Checked = My.Settings.data115
        IgnoreErrors2.Checked = My.Settings.data116
        Dev2TerminatorEnable.Checked = My.Settings.data117
        CheckBoxSendBlockingDev2.Checked = My.Settings.data118
        Dev2STBMask.Text = My.Settings.data119
        Div1000Dev2.Checked = My.Settings.data120
        Dev23457Aseven.Checked = My.Settings.data121
        Dev2TerminatorEnable2.Checked = My.Settings.data122

        Dev2K2001isolatedata.Checked = My.Settings.data203
        Dev2K2001isolatedataCHAR.Text = My.Settings.data204

        Mult1000Dev2.Checked = My.Settings.data223

        Dev2Timeout.Text = My.Settings.data241

        Dev2delayop.Text = My.Settings.data253

        txtq2d.Text = My.Settings.data311
        Dev2pauseDurationInSeconds.Text = My.Settings.data312
        Dev2runStopwatchEveryInMins.Text = My.Settings.data313
        Dev2IntEnable.Checked = My.Settings.data314
        Dev2Regex.Checked = My.Settings.data347

        Dev2DecimalNumDPs.Text = My.Settings.data465

        Dev2SendQuery.Checked = My.Settings.data493

        txtOperationDev2.Text = My.Settings.data521

    End Sub

    ' Device 2 profile 6
    Private Sub ProfDev2_6_Click(sender As Object, e As EventArgs) Handles ProfDev2_6.Click

        If (ProfDev2_6.Checked = True) Then
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            'ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = True
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data123
        CommandStart2.Text = My.Settings.data124
        CommandStart2run.Text = My.Settings.data125
        CommandStop2.Text = My.Settings.data126
        lstIntf2.SelectedIndex = My.Settings.data127
        txtaddr2.Text = My.Settings.data128
        Dev2SampleRate.Text = My.Settings.data129
        Dev2PollingEnable.Checked = My.Settings.data130
        Dev2removeletters.Checked = My.Settings.data131
        IgnoreErrors2.Checked = My.Settings.data132
        Dev2TerminatorEnable.Checked = My.Settings.data133
        CheckBoxSendBlockingDev2.Checked = My.Settings.data134
        Dev2STBMask.Text = My.Settings.data135
        Div1000Dev2.Checked = My.Settings.data136
        Dev23457Aseven.Checked = My.Settings.data137
        Dev2TerminatorEnable2.Checked = My.Settings.data138

        Dev2K2001isolatedata.Checked = My.Settings.data205
        Dev2K2001isolatedataCHAR.Text = My.Settings.data206

        Mult1000Dev2.Checked = My.Settings.data224

        Dev2Timeout.Text = My.Settings.data242

        Dev2delayop.Text = My.Settings.data254

        txtq2d.Text = My.Settings.data315
        Dev2pauseDurationInSeconds.Text = My.Settings.data316
        Dev2runStopwatchEveryInMins.Text = My.Settings.data317
        Dev2IntEnable.Checked = My.Settings.data318
        Dev2Regex.Checked = My.Settings.data348

        Dev2DecimalNumDPs.Text = My.Settings.data466

        Dev2SendQuery.Checked = My.Settings.data495

        txtOperationDev2.Text = My.Settings.data522

    End Sub

    ' Device 2 profile 7
    Private Sub ProfDev2_7_Click(sender As Object, e As EventArgs) Handles ProfDev2_7.Click

        If (ProfDev2_7.Checked = True) Then
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            'ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = True
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data401
        CommandStart2.Text = My.Settings.data402
        CommandStart2run.Text = My.Settings.data403
        CommandStop2.Text = My.Settings.data404
        lstIntf2.SelectedIndex = My.Settings.data405
        txtaddr2.Text = My.Settings.data406
        Dev2SampleRate.Text = My.Settings.data407
        Dev2PollingEnable.Checked = My.Settings.data408
        Dev2removeletters.Checked = My.Settings.data409
        IgnoreErrors2.Checked = My.Settings.data410
        Dev2TerminatorEnable.Checked = My.Settings.data411
        CheckBoxSendBlockingDev2.Checked = My.Settings.data412
        Dev2STBMask.Text = My.Settings.data413
        Div1000Dev2.Checked = My.Settings.data414
        Dev23457Aseven.Checked = My.Settings.data415
        Dev2TerminatorEnable2.Checked = My.Settings.data416

        Dev2K2001isolatedata.Checked = My.Settings.data417
        Dev2K2001isolatedataCHAR.Text = My.Settings.data418

        Mult1000Dev2.Checked = My.Settings.data419

        Dev2Timeout.Text = My.Settings.data420

        Dev2delayop.Text = My.Settings.data421

        txtq2d.Text = My.Settings.data422
        Dev2pauseDurationInSeconds.Text = My.Settings.data423
        Dev2runStopwatchEveryInMins.Text = My.Settings.data424
        Dev2IntEnable.Checked = My.Settings.data425
        Dev2Regex.Checked = My.Settings.data426

        Dev2DecimalNumDPs.Text = My.Settings.data467

        Dev2SendQuery.Checked = My.Settings.data497

        txtOperationDev2.Text = My.Settings.data523

    End Sub


    ' Device 2 profile 8
    Private Sub ProfDev2_8_Click(sender As Object, e As EventArgs) Handles ProfDev2_8.Click

        If (ProfDev2_8.Checked = True) Then
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            'ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = True
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data427
        CommandStart2.Text = My.Settings.data428
        CommandStart2run.Text = My.Settings.data429
        CommandStop2.Text = My.Settings.data430
        lstIntf2.SelectedIndex = My.Settings.data431
        txtaddr2.Text = My.Settings.data432
        Dev2SampleRate.Text = My.Settings.data433
        Dev2PollingEnable.Checked = My.Settings.data434
        Dev2removeletters.Checked = My.Settings.data435
        IgnoreErrors2.Checked = My.Settings.data436
        Dev2TerminatorEnable.Checked = My.Settings.data437
        CheckBoxSendBlockingDev2.Checked = My.Settings.data438
        Dev2STBMask.Text = My.Settings.data439
        Div1000Dev2.Checked = My.Settings.data440
        Dev23457Aseven.Checked = My.Settings.data441
        Dev2TerminatorEnable2.Checked = My.Settings.data442

        Dev2K2001isolatedata.Checked = My.Settings.data443
        Dev2K2001isolatedataCHAR.Text = My.Settings.data444

        Mult1000Dev2.Checked = My.Settings.data445

        Dev2Timeout.Text = My.Settings.data446

        Dev2delayop.Text = My.Settings.data447

        txtq2d.Text = My.Settings.data448
        Dev2pauseDurationInSeconds.Text = My.Settings.data449
        Dev2runStopwatchEveryInMins.Text = My.Settings.data450
        Dev2IntEnable.Checked = My.Settings.data451
        Dev2Regex.Checked = My.Settings.data452

        Dev2DecimalNumDPs.Text = My.Settings.data468

        Dev2SendQuery.Checked = My.Settings.data499

        txtOperationDev2.Text = My.Settings.data524

    End Sub



    Private Sub ProfDev2_9_Click(sender As Object, e As EventArgs) Handles ProfDev2_9.Click

        If (ProfDev2_9.Checked = True) Then
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False

            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = True
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data642
        CommandStart2.Text = My.Settings.data643
        CommandStart2run.Text = My.Settings.data644
        CommandStop2.Text = My.Settings.data645
        lstIntf2.SelectedIndex = My.Settings.data646
        txtaddr2.Text = My.Settings.data647
        Dev2SampleRate.Text = My.Settings.data648
        Dev2PollingEnable.Checked = My.Settings.data649
        Dev2removeletters.Checked = My.Settings.data650
        IgnoreErrors2.Checked = My.Settings.data651
        Dev2TerminatorEnable.Checked = My.Settings.data652
        CheckBoxSendBlockingDev2.Checked = My.Settings.data653
        Dev2STBMask.Text = My.Settings.data654
        Div1000Dev2.Checked = My.Settings.data655
        Dev23457Aseven.Checked = My.Settings.data656
        Dev2TerminatorEnable2.Checked = My.Settings.data657
        Dev2K2001isolatedata.Checked = My.Settings.data658
        Dev2K2001isolatedataCHAR.Text = My.Settings.data659
        Mult1000Dev2.Checked = My.Settings.data660
        Dev2Timeout.Text = My.Settings.data661
        Dev2delayop.Text = My.Settings.data662
        txtq2d.Text = My.Settings.data663
        Dev2pauseDurationInSeconds.Text = My.Settings.data664
        Dev2runStopwatchEveryInMins.Text = My.Settings.data665
        Dev2IntEnable.Checked = My.Settings.data666
        Dev2Regex.Checked = My.Settings.data667
        Dev2DecimalNumDPs.Text = My.Settings.data668
        txtOperationDev2.Text = My.Settings.data670

    End Sub

    Private Sub ProfDev2_10_Click(sender As Object, e As EventArgs) Handles ProfDev2_10.Click

        If (ProfDev2_10.Checked = True) Then
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False

            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = True
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data671
        CommandStart2.Text = My.Settings.data672
        CommandStart2run.Text = My.Settings.data673
        CommandStop2.Text = My.Settings.data674
        lstIntf2.SelectedIndex = My.Settings.data675
        txtaddr2.Text = My.Settings.data676
        Dev2SampleRate.Text = My.Settings.data677
        Dev2PollingEnable.Checked = My.Settings.data678
        Dev2removeletters.Checked = My.Settings.data679
        IgnoreErrors2.Checked = My.Settings.data680
        Dev2TerminatorEnable.Checked = My.Settings.data681
        CheckBoxSendBlockingDev2.Checked = My.Settings.data682
        Dev2STBMask.Text = My.Settings.data683
        Div1000Dev2.Checked = My.Settings.data684
        Dev23457Aseven.Checked = My.Settings.data685
        Dev2TerminatorEnable2.Checked = My.Settings.data686
        Dev2K2001isolatedata.Checked = My.Settings.data687
        Dev2K2001isolatedataCHAR.Text = My.Settings.data688
        Mult1000Dev2.Checked = My.Settings.data689
        Dev2Timeout.Text = My.Settings.data690
        Dev2delayop.Text = My.Settings.data691
        txtq2d.Text = My.Settings.data692
        Dev2pauseDurationInSeconds.Text = My.Settings.data693
        Dev2runStopwatchEveryInMins.Text = My.Settings.data694
        Dev2IntEnable.Checked = My.Settings.data695
        Dev2Regex.Checked = My.Settings.data696
        Dev2DecimalNumDPs.Text = My.Settings.data697
        txtOperationDev2.Text = My.Settings.data699

    End Sub

    Private Sub ProfDev2_11_Click(sender As Object, e As EventArgs) Handles ProfDev2_11.Click

        If (ProfDev2_11.Checked = True) Then
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False

            ProfDev2_12.Checked = False
        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = True
            ProfDev2_12.Checked = False
        End If

        txtname2.Text = My.Settings.data700
        CommandStart2.Text = My.Settings.data701
        CommandStart2run.Text = My.Settings.data702
        CommandStop2.Text = My.Settings.data703
        lstIntf2.SelectedIndex = My.Settings.data704
        txtaddr2.Text = My.Settings.data705
        Dev2SampleRate.Text = My.Settings.data706
        Dev2PollingEnable.Checked = My.Settings.data707
        Dev2removeletters.Checked = My.Settings.data708
        IgnoreErrors2.Checked = My.Settings.data709
        Dev2TerminatorEnable.Checked = My.Settings.data710
        CheckBoxSendBlockingDev2.Checked = My.Settings.data711
        Dev2STBMask.Text = My.Settings.data712
        Div1000Dev2.Checked = My.Settings.data713
        Dev23457Aseven.Checked = My.Settings.data714
        Dev2TerminatorEnable2.Checked = My.Settings.data715
        Dev2K2001isolatedata.Checked = My.Settings.data716
        Dev2K2001isolatedataCHAR.Text = My.Settings.data717
        Mult1000Dev2.Checked = My.Settings.data718
        Dev2Timeout.Text = My.Settings.data719
        Dev2delayop.Text = My.Settings.data720
        txtq2d.Text = My.Settings.data721
        Dev2pauseDurationInSeconds.Text = My.Settings.data722
        Dev2runStopwatchEveryInMins.Text = My.Settings.data723
        Dev2IntEnable.Checked = My.Settings.data724
        Dev2Regex.Checked = My.Settings.data725
        Dev2DecimalNumDPs.Text = My.Settings.data726
        txtOperationDev2.Text = My.Settings.data728

    End Sub

    Private Sub ProfDev2_12_Click(sender As Object, e As EventArgs) Handles ProfDev2_12.Click

        If (ProfDev2_12.Checked = True) Then
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False

        Else
            ProfDev2_1.Checked = False
            ProfDev2_2.Checked = False
            ProfDev2_3.Checked = False
            ProfDev2_4.Checked = False
            ProfDev2_5.Checked = False
            ProfDev2_6.Checked = False
            ProfDev2_7.Checked = False
            ProfDev2_8.Checked = False
            ProfDev2_9.Checked = False
            ProfDev2_10.Checked = False
            ProfDev2_11.Checked = False
            ProfDev2_12.Checked = True
        End If

        txtname2.Text = My.Settings.data729
        CommandStart2.Text = My.Settings.data730
        CommandStart2run.Text = My.Settings.data731
        CommandStop2.Text = My.Settings.data732
        lstIntf2.SelectedIndex = My.Settings.data733
        txtaddr2.Text = My.Settings.data734
        Dev2SampleRate.Text = My.Settings.data735
        Dev2PollingEnable.Checked = My.Settings.data736
        Dev2removeletters.Checked = My.Settings.data737
        IgnoreErrors2.Checked = My.Settings.data738
        Dev2TerminatorEnable.Checked = My.Settings.data739
        CheckBoxSendBlockingDev2.Checked = My.Settings.data740
        Dev2STBMask.Text = My.Settings.data741
        Div1000Dev2.Checked = My.Settings.data742
        Dev23457Aseven.Checked = My.Settings.data743
        Dev2TerminatorEnable2.Checked = My.Settings.data744
        Dev2K2001isolatedata.Checked = My.Settings.data745
        Dev2K2001isolatedataCHAR.Text = My.Settings.data746
        Mult1000Dev2.Checked = My.Settings.data747
        Dev2Timeout.Text = My.Settings.data748
        Dev2delayop.Text = My.Settings.data749
        txtq2d.Text = My.Settings.data750
        Dev2pauseDurationInSeconds.Text = My.Settings.data751
        Dev2runStopwatchEveryInMins.Text = My.Settings.data752
        Dev2IntEnable.Checked = My.Settings.data753
        Dev2Regex.Checked = My.Settings.data754
        Dev2DecimalNumDPs.Text = My.Settings.data755
        txtOperationDev2.Text = My.Settings.data757

    End Sub



    ' Backup application settings to a file
    Public Sub BackupSettings(filePath As String)
        Try
            Dim settings As New Dictionary(Of String, Object)

            ' Iterate over all the settings
            For Each key As String In My.Settings.Properties.Cast(Of SettingsProperty)().Select(Function(p) p.Name)
                settings.Add(key, My.Settings(key))
            Next

            ' Serialize the dictionary to a file
            Using fs As New FileStream(filePath, FileMode.Create)
                Dim formatter As New BinaryFormatter()
                formatter.Serialize(fs, settings)
            End Using

            Dim entryCount As Integer = settings.Count
            MessageBox.Show($"Profiles/settings export completed successfully{vbCrLf}{vbCrLf}File = ProfilesData.dat{vbCrLf}{vbCrLf}{entryCount} entries were exported", "WinGPIB Export", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"An error occurred while exporting profiles/settings: {ex.Message}", "WinGPIB Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnBackup_Click(sender As Object, e As EventArgs) Handles btnBackup.Click
        Dim backupFilePath As String = Path.Combine(documentsPath, "WinGPIBdata\ProfilesData.dat")
        BackupSettings(backupFilePath)
    End Sub


    ' Restore application settings from a file
    Public Sub RestoreSettings(filePath As String)
        Try
            ' Deserialize the dictionary from the file
            Dim settings As Dictionary(Of String, Object)

            Using fs As New FileStream(filePath, FileMode.Open)
                Dim formatter As New BinaryFormatter()
                settings = CType(formatter.Deserialize(fs), Dictionary(Of String, Object))
            End Using

            ' Restore the settings
            For Each kvp As KeyValuePair(Of String, Object) In settings
                My.Settings(kvp.Key) = kvp.Value
            Next

            ' Save the settings to make them persistent
            My.Settings.Save()

            ' Count the number of entries and display it with a carriage return
            Dim entryCount As Integer = settings.Count
            MessageBox.Show($"Profiles/settings imported successfully.{vbCrLf}{vbCrLf}{entryCount} entries were loaded", "WinGPIB Import", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"An error occurred while importing profiles/settings: {ex.Message}", "WinGPIB import Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnRestore_Click(sender As Object, e As EventArgs) Handles btnRestore.Click
        Using openFileDialog As New OpenFileDialog()
            openFileDialog.InitialDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WinGPIBdata")
            openFileDialog.Filter = "Data Files (*.dat)|*.dat|All Files (*.*)|*.*"
            openFileDialog.Title = "Select a Profiles Data File"

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Dim selectedFilePath As String = openFileDialog.FileName

                ' Check if the selected file has valid content
                If IsValidProfileFile(selectedFilePath) Then
                    RestoreSettings(selectedFilePath)
                Else
                    MessageBox.Show("The selected file is not a valid profile/settings data file.", "Invalid File", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Using
    End Sub


    ' Function to check if the file has valid contents using deserialization
    Private Function IsValidProfileFile(filePath As String) As Boolean
        Try
            ' Attempt to open and deserialize the file
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                Dim formatter As New BinaryFormatter()
                Dim data As Object = formatter.Deserialize(fs)

                ' Check if the deserialized data is of the expected type (e.g., Dictionary)
                If TypeOf data Is Dictionary(Of String, Object) Then
                    ' Perform additional checks on contents if necessary
                    Return True
                End If
            End Using
        Catch ex As Exception
            ' If deserialization fails, return False
        End Try
        Return False
    End Function

End Class