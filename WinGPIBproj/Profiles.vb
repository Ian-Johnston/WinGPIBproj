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

    End Sub

    Private Sub CheckBoxEnableTooltips_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxEnableTooltips.CheckedChanged

        If CheckBoxEnableTooltips.Checked = True Then
            ToolTip1.Active = True
        Else
            ToolTip1.Active = False
        End If

        My.Settings.data507 = CheckBoxEnableTooltips.Checked          ' save off immediately

    End Sub

    Private Sub TextBoxTextEditor_TextChanged(sender As Object, e As EventArgs) Handles TextBoxTextEditor.TextChanged

        TextEditorPath = TextBoxTextEditor.Text ' Adjust the path as needed
        My.Settings.data506 = TextBoxTextEditor.Text          ' save off immediately
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

        SaveSettings()

    End Sub


    Private Sub ButtonSaveSettings2_Click(sender As Object, e As EventArgs) Handles ButtonSaveSettings2.Click

        SaveSettings()

    End Sub


    Private Sub SaveSettings()

        ' Profile selection flags Dev 1 & 2 (1..20)
        My.Settings.Dev1Prof1 = (Dev1ProfileNumber() = 1)
        My.Settings.Dev1Prof2 = (Dev1ProfileNumber() = 2)
        My.Settings.Dev1Prof3 = (Dev1ProfileNumber() = 3)
        My.Settings.Dev1Prof4 = (Dev1ProfileNumber() = 4)
        My.Settings.Dev1Prof5 = (Dev1ProfileNumber() = 5)
        My.Settings.Dev1Prof6 = (Dev1ProfileNumber() = 6)
        My.Settings.Dev1Prof7 = (Dev1ProfileNumber() = 7)
        My.Settings.Dev1Prof8 = (Dev1ProfileNumber() = 8)
        My.Settings.Dev1Prof9 = (Dev1ProfileNumber() = 9)
        My.Settings.Dev1Prof10 = (Dev1ProfileNumber() = 10)
        My.Settings.Dev1Prof11 = (Dev1ProfileNumber() = 11)
        My.Settings.Dev1Prof12 = (Dev1ProfileNumber() = 12)
        My.Settings.Dev1Prof13 = (Dev1ProfileNumber() = 13)
        My.Settings.Dev1Prof14 = (Dev1ProfileNumber() = 14)
        My.Settings.Dev1Prof15 = (Dev1ProfileNumber() = 15)
        My.Settings.Dev1Prof16 = (Dev1ProfileNumber() = 16)
        My.Settings.Dev1Prof17 = (Dev1ProfileNumber() = 17)
        My.Settings.Dev1Prof18 = (Dev1ProfileNumber() = 18)
        My.Settings.Dev1Prof19 = (Dev1ProfileNumber() = 19)
        My.Settings.Dev1Prof20 = (Dev1ProfileNumber() = 20)
        My.Settings.Dev2Prof1 = (Dev2ProfileNumber() = 1)
        My.Settings.Dev2Prof2 = (Dev2ProfileNumber() = 2)
        My.Settings.Dev2Prof3 = (Dev2ProfileNumber() = 3)
        My.Settings.Dev2Prof4 = (Dev2ProfileNumber() = 4)
        My.Settings.Dev2Prof5 = (Dev2ProfileNumber() = 5)
        My.Settings.Dev2Prof6 = (Dev2ProfileNumber() = 6)
        My.Settings.Dev2Prof7 = (Dev2ProfileNumber() = 7)
        My.Settings.Dev2Prof8 = (Dev2ProfileNumber() = 8)
        My.Settings.Dev2Prof9 = (Dev2ProfileNumber() = 9)
        My.Settings.Dev2Prof10 = (Dev2ProfileNumber() = 10)
        My.Settings.Dev2Prof11 = (Dev2ProfileNumber() = 11)
        My.Settings.Dev2Prof12 = (Dev2ProfileNumber() = 12)
        My.Settings.Dev2Prof13 = (Dev2ProfileNumber() = 13)
        My.Settings.Dev2Prof14 = (Dev2ProfileNumber() = 14)
        My.Settings.Dev2Prof15 = (Dev2ProfileNumber() = 15)
        My.Settings.Dev2Prof16 = (Dev2ProfileNumber() = 16)
        My.Settings.Dev2Prof17 = (Dev2ProfileNumber() = 17)
        My.Settings.Dev2Prof18 = (Dev2ProfileNumber() = 18)
        My.Settings.Dev2Prof19 = (Dev2ProfileNumber() = 19)
        My.Settings.Dev2Prof20 = (Dev2ProfileNumber() = 20)

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
        If Dev1ProfileNumber() = 1 Then
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
        If Dev1ProfileNumber() = 2 Then
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
        If Dev1ProfileNumber() = 3 Then
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
        If Dev1ProfileNumber() = 4 Then
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
        If Dev1ProfileNumber() = 5 Then
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
        If Dev1ProfileNumber() = 6 Then
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
        If Dev1ProfileNumber() = 7 Then
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
        If Dev1ProfileNumber() = 8 Then
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
        If Dev1ProfileNumber() = 9 Then
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
        If Dev1ProfileNumber() = 10 Then
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
        If Dev1ProfileNumber() = 11 Then
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
        If Dev1ProfileNumber() = 12 Then
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

        ' Save Dev1 Profile 13
        If Dev1ProfileNumber() = 13 Then
            My.Settings.data758 = txtname1.Text
            My.Settings.data759 = CommandStart1.Text
            My.Settings.data760 = CommandStart1run.Text
            My.Settings.data761 = CommandStop1.Text
            My.Settings.data762 = lstIntf1.SelectedIndex
            My.Settings.data763 = txtaddr1.Text
            My.Settings.data764 = Dev1SampleRate.Text
            My.Settings.data765 = Dev1PollingEnable.Checked
            My.Settings.data766 = Dev1removeletters.Checked
            My.Settings.data767 = IgnoreErrors1.Checked
            My.Settings.data768 = Dev1TerminatorEnable.Checked
            My.Settings.data769 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data770 = Dev1STBMask.Text
            My.Settings.data771 = Div1000Dev1.Checked
            My.Settings.data772 = Dev13457Aseven.Checked
            My.Settings.data773 = Dev1TerminatorEnable2.Checked
            My.Settings.data774 = Dev1K2001isolatedata.Checked
            My.Settings.data775 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data776 = Mult1000Dev1.Checked
            My.Settings.data777 = Val(Dev1Timeout.Text)
            My.Settings.data778 = Val(Dev1delayop.Text)
            My.Settings.data779 = txtq1d.Text
            My.Settings.data780 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data781 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data782 = Dev1IntEnable.Checked
            My.Settings.data783 = Dev1Regex.Checked
            My.Settings.data784 = Dev1DecimalNumDPs.Text
            My.Settings.data785 = Dev1IntEnable.Checked
            My.Settings.data786 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 14
        If Dev1ProfileNumber() = 14 Then
            My.Settings.data787 = txtname1.Text
            My.Settings.data788 = CommandStart1.Text
            My.Settings.data789 = CommandStart1run.Text
            My.Settings.data790 = CommandStop1.Text
            My.Settings.data791 = lstIntf1.SelectedIndex
            My.Settings.data792 = txtaddr1.Text
            My.Settings.data793 = Dev1SampleRate.Text
            My.Settings.data794 = Dev1PollingEnable.Checked
            My.Settings.data795 = Dev1removeletters.Checked
            My.Settings.data796 = IgnoreErrors1.Checked
            My.Settings.data797 = Dev1TerminatorEnable.Checked
            My.Settings.data798 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data799 = Dev1STBMask.Text
            My.Settings.data800 = Div1000Dev1.Checked
            My.Settings.data801 = Dev13457Aseven.Checked
            My.Settings.data802 = Dev1TerminatorEnable2.Checked
            My.Settings.data803 = Dev1K2001isolatedata.Checked
            My.Settings.data804 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data805 = Mult1000Dev1.Checked
            My.Settings.data806 = Val(Dev1Timeout.Text)
            My.Settings.data807 = Val(Dev1delayop.Text)
            My.Settings.data808 = txtq1d.Text
            My.Settings.data809 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data810 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data811 = Dev1IntEnable.Checked
            My.Settings.data812 = Dev1Regex.Checked
            My.Settings.data813 = Dev1DecimalNumDPs.Text
            My.Settings.data814 = Dev1IntEnable.Checked
            My.Settings.data815 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 15
        If Dev1ProfileNumber() = 15 Then
            My.Settings.data816 = txtname1.Text
            My.Settings.data817 = CommandStart1.Text
            My.Settings.data818 = CommandStart1run.Text
            My.Settings.data819 = CommandStop1.Text
            My.Settings.data820 = lstIntf1.SelectedIndex
            My.Settings.data821 = txtaddr1.Text
            My.Settings.data822 = Dev1SampleRate.Text
            My.Settings.data823 = Dev1PollingEnable.Checked
            My.Settings.data824 = Dev1removeletters.Checked
            My.Settings.data825 = IgnoreErrors1.Checked
            My.Settings.data826 = Dev1TerminatorEnable.Checked
            My.Settings.data827 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data828 = Dev1STBMask.Text
            My.Settings.data829 = Div1000Dev1.Checked
            My.Settings.data830 = Dev13457Aseven.Checked
            My.Settings.data831 = Dev1TerminatorEnable2.Checked
            My.Settings.data832 = Dev1K2001isolatedata.Checked
            My.Settings.data833 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data834 = Mult1000Dev1.Checked
            My.Settings.data835 = Val(Dev1Timeout.Text)
            My.Settings.data836 = Val(Dev1delayop.Text)
            My.Settings.data837 = txtq1d.Text
            My.Settings.data838 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data839 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data840 = Dev1IntEnable.Checked
            My.Settings.data841 = Dev1Regex.Checked
            My.Settings.data842 = Dev1DecimalNumDPs.Text
            My.Settings.data843 = Dev1IntEnable.Checked
            My.Settings.data844 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 16
        If Dev1ProfileNumber() = 16 Then
            My.Settings.data845 = txtname1.Text
            My.Settings.data846 = CommandStart1.Text
            My.Settings.data847 = CommandStart1run.Text
            My.Settings.data848 = CommandStop1.Text
            My.Settings.data849 = lstIntf1.SelectedIndex
            My.Settings.data850 = txtaddr1.Text
            My.Settings.data851 = Dev1SampleRate.Text
            My.Settings.data852 = Dev1PollingEnable.Checked
            My.Settings.data853 = Dev1removeletters.Checked
            My.Settings.data854 = IgnoreErrors1.Checked
            My.Settings.data855 = Dev1TerminatorEnable.Checked
            My.Settings.data856 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data857 = Dev1STBMask.Text
            My.Settings.data858 = Div1000Dev1.Checked
            My.Settings.data859 = Dev13457Aseven.Checked
            My.Settings.data860 = Dev1TerminatorEnable2.Checked
            My.Settings.data861 = Dev1K2001isolatedata.Checked
            My.Settings.data862 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data863 = Mult1000Dev1.Checked
            My.Settings.data864 = Val(Dev1Timeout.Text)
            My.Settings.data865 = Val(Dev1delayop.Text)
            My.Settings.data866 = txtq1d.Text
            My.Settings.data867 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data868 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data869 = Dev1IntEnable.Checked
            My.Settings.data870 = Dev1Regex.Checked
            My.Settings.data871 = Dev1DecimalNumDPs.Text
            My.Settings.data872 = Dev1IntEnable.Checked
            My.Settings.data873 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 17
        If Dev1ProfileNumber() = 17 Then
            My.Settings.data874 = txtname1.Text
            My.Settings.data875 = CommandStart1.Text
            My.Settings.data876 = CommandStart1run.Text
            My.Settings.data877 = CommandStop1.Text
            My.Settings.data878 = lstIntf1.SelectedIndex
            My.Settings.data879 = txtaddr1.Text
            My.Settings.data880 = Dev1SampleRate.Text
            My.Settings.data881 = Dev1PollingEnable.Checked
            My.Settings.data882 = Dev1removeletters.Checked
            My.Settings.data883 = IgnoreErrors1.Checked
            My.Settings.data884 = Dev1TerminatorEnable.Checked
            My.Settings.data885 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data886 = Dev1STBMask.Text
            My.Settings.data887 = Div1000Dev1.Checked
            My.Settings.data888 = Dev13457Aseven.Checked
            My.Settings.data889 = Dev1TerminatorEnable2.Checked
            My.Settings.data890 = Dev1K2001isolatedata.Checked
            My.Settings.data891 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data892 = Mult1000Dev1.Checked
            My.Settings.data893 = Val(Dev1Timeout.Text)
            My.Settings.data894 = Val(Dev1delayop.Text)
            My.Settings.data895 = txtq1d.Text
            My.Settings.data896 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data897 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data898 = Dev1IntEnable.Checked
            My.Settings.data899 = Dev1Regex.Checked
            My.Settings.data900 = Dev1DecimalNumDPs.Text
            My.Settings.data901 = Dev1IntEnable.Checked
            My.Settings.data902 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 18
        If Dev1ProfileNumber() = 18 Then
            My.Settings.data903 = txtname1.Text
            My.Settings.data904 = CommandStart1.Text
            My.Settings.data905 = CommandStart1run.Text
            My.Settings.data906 = CommandStop1.Text
            My.Settings.data907 = lstIntf1.SelectedIndex
            My.Settings.data908 = txtaddr1.Text
            My.Settings.data909 = Dev1SampleRate.Text
            My.Settings.data910 = Dev1PollingEnable.Checked
            My.Settings.data911 = Dev1removeletters.Checked
            My.Settings.data912 = IgnoreErrors1.Checked
            My.Settings.data913 = Dev1TerminatorEnable.Checked
            My.Settings.data914 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data915 = Dev1STBMask.Text
            My.Settings.data916 = Div1000Dev1.Checked
            My.Settings.data917 = Dev13457Aseven.Checked
            My.Settings.data918 = Dev1TerminatorEnable2.Checked
            My.Settings.data919 = Dev1K2001isolatedata.Checked
            My.Settings.data920 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data921 = Mult1000Dev1.Checked
            My.Settings.data922 = Val(Dev1Timeout.Text)
            My.Settings.data923 = Val(Dev1delayop.Text)
            My.Settings.data924 = txtq1d.Text
            My.Settings.data925 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data926 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data927 = Dev1IntEnable.Checked
            My.Settings.data928 = Dev1Regex.Checked
            My.Settings.data929 = Dev1DecimalNumDPs.Text
            My.Settings.data930 = Dev1IntEnable.Checked
            My.Settings.data931 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 19
        If Dev1ProfileNumber() = 19 Then
            My.Settings.data932 = txtname1.Text
            My.Settings.data933 = CommandStart1.Text
            My.Settings.data934 = CommandStart1run.Text
            My.Settings.data935 = CommandStop1.Text
            My.Settings.data936 = lstIntf1.SelectedIndex
            My.Settings.data937 = txtaddr1.Text
            My.Settings.data938 = Dev1SampleRate.Text
            My.Settings.data939 = Dev1PollingEnable.Checked
            My.Settings.data940 = Dev1removeletters.Checked
            My.Settings.data941 = IgnoreErrors1.Checked
            My.Settings.data942 = Dev1TerminatorEnable.Checked
            My.Settings.data943 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data944 = Dev1STBMask.Text
            My.Settings.data945 = Div1000Dev1.Checked
            My.Settings.data946 = Dev13457Aseven.Checked
            My.Settings.data947 = Dev1TerminatorEnable2.Checked
            My.Settings.data948 = Dev1K2001isolatedata.Checked
            My.Settings.data949 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data950 = Mult1000Dev1.Checked
            My.Settings.data951 = Val(Dev1Timeout.Text)
            My.Settings.data952 = Val(Dev1delayop.Text)
            My.Settings.data953 = txtq1d.Text
            My.Settings.data954 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data955 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data956 = Dev1IntEnable.Checked
            My.Settings.data957 = Dev1Regex.Checked
            My.Settings.data958 = Dev1DecimalNumDPs.Text
            My.Settings.data959 = Dev1IntEnable.Checked
            My.Settings.data960 = txtOperationDev1.Text
        End If

        ' Save Dev1 Profile 20
        If Dev1ProfileNumber() = 20 Then
            My.Settings.data961 = txtname1.Text
            My.Settings.data962 = CommandStart1.Text
            My.Settings.data963 = CommandStart1run.Text
            My.Settings.data964 = CommandStop1.Text
            My.Settings.data965 = lstIntf1.SelectedIndex
            My.Settings.data966 = txtaddr1.Text
            My.Settings.data967 = Dev1SampleRate.Text
            My.Settings.data968 = Dev1PollingEnable.Checked
            My.Settings.data969 = Dev1removeletters.Checked
            My.Settings.data970 = IgnoreErrors1.Checked
            My.Settings.data971 = Dev1TerminatorEnable.Checked
            My.Settings.data972 = CheckBoxSendBlockingDev1.Checked
            My.Settings.data973 = Dev1STBMask.Text
            My.Settings.data974 = Div1000Dev1.Checked
            My.Settings.data975 = Dev13457Aseven.Checked
            My.Settings.data976 = Dev1TerminatorEnable2.Checked
            My.Settings.data977 = Dev1K2001isolatedata.Checked
            My.Settings.data978 = Dev1K2001isolatedataCHAR.Text
            My.Settings.data979 = Mult1000Dev1.Checked
            My.Settings.data980 = Val(Dev1Timeout.Text)
            My.Settings.data981 = Val(Dev1delayop.Text)
            My.Settings.data982 = txtq1d.Text
            My.Settings.data983 = Val(Dev1pauseDurationInSeconds.Text)
            My.Settings.data984 = Val(Dev1runStopwatchEveryInMins.Text)
            My.Settings.data985 = Dev1IntEnable.Checked
            My.Settings.data986 = Dev1Regex.Checked
            My.Settings.data987 = Dev1DecimalNumDPs.Text
            My.Settings.data988 = Dev1IntEnable.Checked
            My.Settings.data989 = txtOperationDev1.Text
        End If


        ' Save Dev2 Profile 1
        If Dev2ProfileNumber() = 1 Then
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
        If Dev2ProfileNumber() = 2 Then
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
        If Dev2ProfileNumber() = 3 Then
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
        If Dev2ProfileNumber() = 4 Then
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
        If Dev2ProfileNumber() = 5 Then
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
        If Dev2ProfileNumber() = 6 Then
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
        If Dev2ProfileNumber() = 7 Then
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
        If Dev2ProfileNumber() = 8 Then
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
        If Dev2ProfileNumber() = 9 Then
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
        If Dev2ProfileNumber() = 10 Then
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
        If Dev2ProfileNumber() = 11 Then
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
        If Dev2ProfileNumber() = 12 Then
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

        ' Save Dev2 Profile 13
        If Dev2ProfileNumber() = 13 Then
            My.Settings.data990 = txtname2.Text
            My.Settings.data991 = CommandStart2.Text
            My.Settings.data992 = CommandStart2run.Text
            My.Settings.data993 = CommandStop2.Text
            My.Settings.data994 = lstIntf2.SelectedIndex
            My.Settings.data995 = txtaddr2.Text
            My.Settings.data996 = Dev2SampleRate.Text
            My.Settings.data997 = Dev2PollingEnable.Checked
            My.Settings.data998 = Dev2removeletters.Checked
            My.Settings.data999 = IgnoreErrors2.Checked
            My.Settings.data1000 = Dev2TerminatorEnable.Checked
            My.Settings.data1001 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data1002 = Dev2STBMask.Text
            My.Settings.data1003 = Div1000Dev2.Checked
            My.Settings.data1004 = Dev23457Aseven.Checked
            My.Settings.data1005 = Dev2TerminatorEnable2.Checked
            My.Settings.data1006 = Dev2K2001isolatedata.Checked
            My.Settings.data1007 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data1008 = Mult1000Dev2.Checked
            My.Settings.data1009 = Val(Dev2Timeout.Text)
            My.Settings.data1010 = Val(Dev2delayop.Text)
            My.Settings.data1011 = txtq2d.Text
            My.Settings.data1012 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data1013 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data1014 = Dev2IntEnable.Checked
            My.Settings.data1015 = Dev2Regex.Checked
            My.Settings.data1016 = Dev2DecimalNumDPs.Text
            My.Settings.data1017 = Dev2SendQuery.Checked
            My.Settings.data1018 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 14
        If Dev2ProfileNumber() = 14 Then
            My.Settings.data1019 = txtname2.Text
            My.Settings.data1020 = CommandStart2.Text
            My.Settings.data1021 = CommandStart2run.Text
            My.Settings.data1022 = CommandStop2.Text
            My.Settings.data1023 = lstIntf2.SelectedIndex
            My.Settings.data1024 = txtaddr2.Text
            My.Settings.data1025 = Dev2SampleRate.Text
            My.Settings.data1026 = Dev2PollingEnable.Checked
            My.Settings.data1027 = Dev2removeletters.Checked
            My.Settings.data1028 = IgnoreErrors2.Checked
            My.Settings.data1029 = Dev2TerminatorEnable.Checked
            My.Settings.data1030 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data1031 = Dev2STBMask.Text
            My.Settings.data1032 = Div1000Dev2.Checked
            My.Settings.data1033 = Dev23457Aseven.Checked
            My.Settings.data1034 = Dev2TerminatorEnable2.Checked
            My.Settings.data1035 = Dev2K2001isolatedata.Checked
            My.Settings.data1036 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data1037 = Mult1000Dev2.Checked
            My.Settings.data1038 = Val(Dev2Timeout.Text)
            My.Settings.data1039 = Val(Dev2delayop.Text)
            My.Settings.data1040 = txtq2d.Text
            My.Settings.data1041 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data1042 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data1043 = Dev2IntEnable.Checked
            My.Settings.data1044 = Dev2Regex.Checked
            My.Settings.data1045 = Dev2DecimalNumDPs.Text
            My.Settings.data1046 = Dev2SendQuery.Checked
            My.Settings.data1047 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 15
        If Dev2ProfileNumber() = 15 Then
            My.Settings.data1048 = txtname2.Text
            My.Settings.data1049 = CommandStart2.Text
            My.Settings.data1050 = CommandStart2run.Text
            My.Settings.data1051 = CommandStop2.Text
            My.Settings.data1052 = lstIntf2.SelectedIndex
            My.Settings.data1053 = txtaddr2.Text
            My.Settings.data1054 = Dev2SampleRate.Text
            My.Settings.data1055 = Dev2PollingEnable.Checked
            My.Settings.data1056 = Dev2removeletters.Checked
            My.Settings.data1057 = IgnoreErrors2.Checked
            My.Settings.data1058 = Dev2TerminatorEnable.Checked
            My.Settings.data1059 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data1060 = Dev2STBMask.Text
            My.Settings.data1061 = Div1000Dev2.Checked
            My.Settings.data1062 = Dev23457Aseven.Checked
            My.Settings.data1063 = Dev2TerminatorEnable2.Checked
            My.Settings.data1064 = Dev2K2001isolatedata.Checked
            My.Settings.data1065 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data1066 = Mult1000Dev2.Checked
            My.Settings.data1067 = Val(Dev2Timeout.Text)
            My.Settings.data1068 = Val(Dev2delayop.Text)
            My.Settings.data1069 = txtq2d.Text
            My.Settings.data1070 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data1071 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data1072 = Dev2IntEnable.Checked
            My.Settings.data1073 = Dev2Regex.Checked
            My.Settings.data1074 = Dev2DecimalNumDPs.Text
            My.Settings.data1075 = Dev2SendQuery.Checked
            My.Settings.data1076 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 16
        If Dev2ProfileNumber() = 16 Then
            My.Settings.data1077 = txtname2.Text
            My.Settings.data1078 = CommandStart2.Text
            My.Settings.data1079 = CommandStart2run.Text
            My.Settings.data1080 = CommandStop2.Text
            My.Settings.data1081 = lstIntf2.SelectedIndex
            My.Settings.data1082 = txtaddr2.Text
            My.Settings.data1083 = Dev2SampleRate.Text
            My.Settings.data1084 = Dev2PollingEnable.Checked
            My.Settings.data1085 = Dev2removeletters.Checked
            My.Settings.data1086 = IgnoreErrors2.Checked
            My.Settings.data1087 = Dev2TerminatorEnable.Checked
            My.Settings.data1088 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data1089 = Dev2STBMask.Text
            My.Settings.data1090 = Div1000Dev2.Checked
            My.Settings.data1091 = Dev23457Aseven.Checked
            My.Settings.data1092 = Dev2TerminatorEnable2.Checked
            My.Settings.data1093 = Dev2K2001isolatedata.Checked
            My.Settings.data1094 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data1095 = Mult1000Dev2.Checked
            My.Settings.data1096 = Val(Dev2Timeout.Text)
            My.Settings.data1097 = Val(Dev2delayop.Text)
            My.Settings.data1098 = txtq2d.Text
            My.Settings.data1099 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data1100 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data1101 = Dev2IntEnable.Checked
            My.Settings.data1102 = Dev2Regex.Checked
            My.Settings.data1103 = Dev2DecimalNumDPs.Text
            My.Settings.data1104 = Dev2SendQuery.Checked
            My.Settings.data1105 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 17
        If Dev2ProfileNumber() = 17 Then
            My.Settings.data1106 = txtname2.Text
            My.Settings.data1107 = CommandStart2.Text
            My.Settings.data1108 = CommandStart2run.Text
            My.Settings.data1109 = CommandStop2.Text
            My.Settings.data1110 = lstIntf2.SelectedIndex
            My.Settings.data1111 = txtaddr2.Text
            My.Settings.data1112 = Dev2SampleRate.Text
            My.Settings.data1113 = Dev2PollingEnable.Checked
            My.Settings.data1114 = Dev2removeletters.Checked
            My.Settings.data1115 = IgnoreErrors2.Checked
            My.Settings.data1116 = Dev2TerminatorEnable.Checked
            My.Settings.data1117 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data1118 = Dev2STBMask.Text
            My.Settings.data1119 = Div1000Dev2.Checked
            My.Settings.data1120 = Dev23457Aseven.Checked
            My.Settings.data1121 = Dev2TerminatorEnable2.Checked
            My.Settings.data1122 = Dev2K2001isolatedata.Checked
            My.Settings.data1123 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data1124 = Mult1000Dev2.Checked
            My.Settings.data1125 = Val(Dev2Timeout.Text)
            My.Settings.data1126 = Val(Dev2delayop.Text)
            My.Settings.data1127 = txtq2d.Text
            My.Settings.data1128 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data1129 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data1130 = Dev2IntEnable.Checked
            My.Settings.data1131 = Dev2Regex.Checked
            My.Settings.data1132 = Dev2DecimalNumDPs.Text
            My.Settings.data1133 = Dev2SendQuery.Checked
            My.Settings.data1134 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 18
        If Dev2ProfileNumber() = 18 Then
            My.Settings.data1135 = txtname2.Text
            My.Settings.data1136 = CommandStart2.Text
            My.Settings.data1137 = CommandStart2run.Text
            My.Settings.data1138 = CommandStop2.Text
            My.Settings.data1139 = lstIntf2.SelectedIndex
            My.Settings.data1140 = txtaddr2.Text
            My.Settings.data1141 = Dev2SampleRate.Text
            My.Settings.data1142 = Dev2PollingEnable.Checked
            My.Settings.data1143 = Dev2removeletters.Checked
            My.Settings.data1144 = IgnoreErrors2.Checked
            My.Settings.data1145 = Dev2TerminatorEnable.Checked
            My.Settings.data1146 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data1147 = Dev2STBMask.Text
            My.Settings.data1148 = Div1000Dev2.Checked
            My.Settings.data1149 = Dev23457Aseven.Checked
            My.Settings.data1150 = Dev2TerminatorEnable2.Checked
            My.Settings.data1151 = Dev2K2001isolatedata.Checked
            My.Settings.data1152 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data1153 = Mult1000Dev2.Checked
            My.Settings.data1154 = Val(Dev2Timeout.Text)
            My.Settings.data1155 = Val(Dev2delayop.Text)
            My.Settings.data1156 = txtq2d.Text
            My.Settings.data1157 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data1158 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data1159 = Dev2IntEnable.Checked
            My.Settings.data1160 = Dev2Regex.Checked
            My.Settings.data1161 = Dev2DecimalNumDPs.Text
            My.Settings.data1162 = Dev2SendQuery.Checked
            My.Settings.data1163 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 19
        If Dev2ProfileNumber() = 19 Then
            My.Settings.data1164 = txtname2.Text
            My.Settings.data1165 = CommandStart2.Text
            My.Settings.data1166 = CommandStart2run.Text
            My.Settings.data1167 = CommandStop2.Text
            My.Settings.data1168 = lstIntf2.SelectedIndex
            My.Settings.data1169 = txtaddr2.Text
            My.Settings.data1170 = Dev2SampleRate.Text
            My.Settings.data1171 = Dev2PollingEnable.Checked
            My.Settings.data1172 = Dev2removeletters.Checked
            My.Settings.data1173 = IgnoreErrors2.Checked
            My.Settings.data1174 = Dev2TerminatorEnable.Checked
            My.Settings.data1175 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data1176 = Dev2STBMask.Text
            My.Settings.data1177 = Div1000Dev2.Checked
            My.Settings.data1178 = Dev23457Aseven.Checked
            My.Settings.data1179 = Dev2TerminatorEnable2.Checked
            My.Settings.data1180 = Dev2K2001isolatedata.Checked
            My.Settings.data1181 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data1182 = Mult1000Dev2.Checked
            My.Settings.data1183 = Val(Dev2Timeout.Text)
            My.Settings.data1184 = Val(Dev2delayop.Text)
            My.Settings.data1185 = txtq2d.Text
            My.Settings.data1186 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data1187 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data1188 = Dev2IntEnable.Checked
            My.Settings.data1189 = Dev2Regex.Checked
            My.Settings.data1190 = Dev2DecimalNumDPs.Text
            My.Settings.data1191 = Dev2SendQuery.Checked
            My.Settings.data1192 = txtOperationDev2.Text
        End If

        ' Save Dev2 Profile 20
        If Dev2ProfileNumber() = 20 Then
            My.Settings.data1193 = txtname2.Text
            My.Settings.data1194 = CommandStart2.Text
            My.Settings.data1195 = CommandStart2run.Text
            My.Settings.data1196 = CommandStop2.Text
            My.Settings.data1197 = lstIntf2.SelectedIndex
            My.Settings.data1198 = txtaddr2.Text
            My.Settings.data1199 = Dev2SampleRate.Text
            My.Settings.data1200 = Dev2PollingEnable.Checked
            My.Settings.data1201 = Dev2removeletters.Checked
            My.Settings.data1202 = IgnoreErrors2.Checked
            My.Settings.data1203 = Dev2TerminatorEnable.Checked
            My.Settings.data1204 = CheckBoxSendBlockingDev2.Checked
            My.Settings.data1205 = Dev2STBMask.Text
            My.Settings.data1206 = Div1000Dev2.Checked
            My.Settings.data1207 = Dev23457Aseven.Checked
            My.Settings.data1208 = Dev2TerminatorEnable2.Checked
            My.Settings.data1209 = Dev2K2001isolatedata.Checked
            My.Settings.data1210 = Dev2K2001isolatedataCHAR.Text
            My.Settings.data1211 = Mult1000Dev2.Checked
            My.Settings.data1212 = Val(Dev2Timeout.Text)
            My.Settings.data1213 = Val(Dev2delayop.Text)
            My.Settings.data1214 = txtq2d.Text
            My.Settings.data1215 = Val(Dev2pauseDurationInSeconds.Text)
            My.Settings.data1216 = Val(Dev2runStopwatchEveryInMins.Text)
            My.Settings.data1217 = Dev2IntEnable.Checked
            My.Settings.data1218 = Dev2Regex.Checked
            My.Settings.data1219 = Dev2DecimalNumDPs.Text
            My.Settings.data1220 = Dev2SendQuery.Checked
            My.Settings.data1221 = txtOperationDev2.Text
        End If


        ' Refresh dropdown names after saving
        Dim dev1Sel As Integer = cboDev1Device.SelectedIndex
        Dim dev2Sel As Integer = cboDev2Device.SelectedIndex

        _suppressDev1Sync = True
        _suppressDev2Sync = True

        PopulateDeviceDropdownsFromNames()

        ' restore selection (only if it was valid)
        If dev1Sel >= 0 AndAlso dev1Sel < cboDev1Device.Items.Count Then cboDev1Device.SelectedIndex = dev1Sel
        If dev2Sel >= 0 AndAlso dev2Sel < cboDev2Device.Items.Count Then cboDev2Device.SelectedIndex = dev2Sel

        _suppressDev1Sync = False
        _suppressDev2Sync = False

        If CSVdelimiterComma.Checked = True Then
            CSVdelimit = ","
            My.Settings.data29 = ","
        Else
            CSVdelimit = ";"
            My.Settings.data29 = ";"
        End If

        ' Save the settings to persist the changes
        My.Settings.Save()

    End Sub

    ' Device 1 profile 1
    Private Sub LoadDev1Profile_1()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 0


    End Sub

    ' Device 1 profile 2
    Private Sub LoadDev1Profile_2()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 1

    End Sub

    ' Device 1 profile 3
    Private Sub LoadDev1Profile_3()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 2

    End Sub

    ' Device 1 profile 4
    Private Sub LoadDev1Profile_4()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 3

    End Sub

    ' Device 1 profile 5
    Private Sub LoadDev1Profile_5()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 4

    End Sub

    ' Device 1 profile 6
    Private Sub LoadDev1Profile_6()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 5

    End Sub

    ' Device 1 profile 7
    Private Sub LoadDev1Profile_7()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 6

    End Sub


    ' Device 1 profile 8
    Private Sub LoadDev1Profile_8()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 7

    End Sub

    Private Sub LoadDev1Profile_9()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 8

    End Sub

    Private Sub LoadDev1Profile_10()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 9

    End Sub

    Private Sub LoadDev1Profile_11()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 10

    End Sub

    Private Sub LoadDev1Profile_12()


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

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 11

    End Sub

    Private Sub LoadDev1Profile_13()


        txtname1.Text = My.Settings.data758
        CommandStart1.Text = My.Settings.data759
        CommandStart1run.Text = My.Settings.data760
        CommandStop1.Text = My.Settings.data761
        lstIntf1.SelectedIndex = My.Settings.data762
        txtaddr1.Text = My.Settings.data763
        Dev1SampleRate.Text = My.Settings.data764
        Dev1PollingEnable.Checked = My.Settings.data765
        Dev1removeletters.Checked = My.Settings.data766
        IgnoreErrors1.Checked = My.Settings.data767
        Dev1TerminatorEnable.Checked = My.Settings.data768
        CheckBoxSendBlockingDev1.Checked = My.Settings.data769
        Dev1STBMask.Text = My.Settings.data770
        Div1000Dev1.Checked = My.Settings.data771
        Dev13457Aseven.Checked = My.Settings.data772
        Dev1TerminatorEnable2.Checked = My.Settings.data773
        Dev1K2001isolatedata.Checked = My.Settings.data774
        Dev1K2001isolatedataCHAR.Text = My.Settings.data775
        Mult1000Dev1.Checked = My.Settings.data776
        Dev1Timeout.Text = My.Settings.data777
        Dev1delayop.Text = My.Settings.data778
        txtq1d.Text = My.Settings.data779
        Dev1pauseDurationInSeconds.Text = My.Settings.data780
        Dev1runStopwatchEveryInMins.Text = My.Settings.data781
        Dev1IntEnable.Checked = My.Settings.data782
        Dev1Regex.Checked = My.Settings.data783
        Dev1DecimalNumDPs.Text = My.Settings.data784
        txtOperationDev1.Text = My.Settings.data786

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 12

    End Sub

    Private Sub LoadDev1Profile_14()


        txtname1.Text = My.Settings.data787
        CommandStart1.Text = My.Settings.data788
        CommandStart1run.Text = My.Settings.data789
        CommandStop1.Text = My.Settings.data790
        lstIntf1.SelectedIndex = My.Settings.data791
        txtaddr1.Text = My.Settings.data792
        Dev1SampleRate.Text = My.Settings.data793
        Dev1PollingEnable.Checked = My.Settings.data794
        Dev1removeletters.Checked = My.Settings.data795
        IgnoreErrors1.Checked = My.Settings.data796
        Dev1TerminatorEnable.Checked = My.Settings.data797
        CheckBoxSendBlockingDev1.Checked = My.Settings.data798
        Dev1STBMask.Text = My.Settings.data799
        Div1000Dev1.Checked = My.Settings.data800
        Dev13457Aseven.Checked = My.Settings.data801
        Dev1TerminatorEnable2.Checked = My.Settings.data802
        Dev1K2001isolatedata.Checked = My.Settings.data803
        Dev1K2001isolatedataCHAR.Text = My.Settings.data804
        Mult1000Dev1.Checked = My.Settings.data805
        Dev1Timeout.Text = My.Settings.data806
        Dev1delayop.Text = My.Settings.data807
        txtq1d.Text = My.Settings.data808
        Dev1pauseDurationInSeconds.Text = My.Settings.data809
        Dev1runStopwatchEveryInMins.Text = My.Settings.data810
        Dev1IntEnable.Checked = My.Settings.data811
        Dev1Regex.Checked = My.Settings.data812
        Dev1DecimalNumDPs.Text = My.Settings.data813
        txtOperationDev1.Text = My.Settings.data815

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 13

    End Sub

    Private Sub LoadDev1Profile_15()


        txtname1.Text = My.Settings.data816
        CommandStart1.Text = My.Settings.data817
        CommandStart1run.Text = My.Settings.data818
        CommandStop1.Text = My.Settings.data819
        lstIntf1.SelectedIndex = My.Settings.data820
        txtaddr1.Text = My.Settings.data821
        Dev1SampleRate.Text = My.Settings.data822
        Dev1PollingEnable.Checked = My.Settings.data823
        Dev1removeletters.Checked = My.Settings.data824
        IgnoreErrors1.Checked = My.Settings.data825
        Dev1TerminatorEnable.Checked = My.Settings.data826
        CheckBoxSendBlockingDev1.Checked = My.Settings.data827
        Dev1STBMask.Text = My.Settings.data828
        Div1000Dev1.Checked = My.Settings.data829
        Dev13457Aseven.Checked = My.Settings.data830
        Dev1TerminatorEnable2.Checked = My.Settings.data831
        Dev1K2001isolatedata.Checked = My.Settings.data832
        Dev1K2001isolatedataCHAR.Text = My.Settings.data833
        Mult1000Dev1.Checked = My.Settings.data834
        Dev1Timeout.Text = My.Settings.data835
        Dev1delayop.Text = My.Settings.data836
        txtq1d.Text = My.Settings.data837
        Dev1pauseDurationInSeconds.Text = My.Settings.data838
        Dev1runStopwatchEveryInMins.Text = My.Settings.data839
        Dev1IntEnable.Checked = My.Settings.data840
        Dev1Regex.Checked = My.Settings.data841
        Dev1DecimalNumDPs.Text = My.Settings.data842
        txtOperationDev1.Text = My.Settings.data844

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 14

    End Sub

    Private Sub LoadDev1Profile_16()


        txtname1.Text = My.Settings.data845
        CommandStart1.Text = My.Settings.data846
        CommandStart1run.Text = My.Settings.data847
        CommandStop1.Text = My.Settings.data848
        lstIntf1.SelectedIndex = My.Settings.data849
        txtaddr1.Text = My.Settings.data850
        Dev1SampleRate.Text = My.Settings.data851
        Dev1PollingEnable.Checked = My.Settings.data852
        Dev1removeletters.Checked = My.Settings.data853
        IgnoreErrors1.Checked = My.Settings.data854
        Dev1TerminatorEnable.Checked = My.Settings.data855
        CheckBoxSendBlockingDev1.Checked = My.Settings.data856
        Dev1STBMask.Text = My.Settings.data857
        Div1000Dev1.Checked = My.Settings.data858
        Dev13457Aseven.Checked = My.Settings.data859
        Dev1TerminatorEnable2.Checked = My.Settings.data860
        Dev1K2001isolatedata.Checked = My.Settings.data861
        Dev1K2001isolatedataCHAR.Text = My.Settings.data862
        Mult1000Dev1.Checked = My.Settings.data863
        Dev1Timeout.Text = My.Settings.data864
        Dev1delayop.Text = My.Settings.data865
        txtq1d.Text = My.Settings.data866
        Dev1pauseDurationInSeconds.Text = My.Settings.data867
        Dev1runStopwatchEveryInMins.Text = My.Settings.data868
        Dev1IntEnable.Checked = My.Settings.data869
        Dev1Regex.Checked = My.Settings.data870
        Dev1DecimalNumDPs.Text = My.Settings.data871
        txtOperationDev1.Text = My.Settings.data873

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 15

    End Sub

    Private Sub LoadDev1Profile_17()


        txtname1.Text = My.Settings.data874
        CommandStart1.Text = My.Settings.data875
        CommandStart1run.Text = My.Settings.data876
        CommandStop1.Text = My.Settings.data877
        lstIntf1.SelectedIndex = My.Settings.data878
        txtaddr1.Text = My.Settings.data879
        Dev1SampleRate.Text = My.Settings.data880
        Dev1PollingEnable.Checked = My.Settings.data881
        Dev1removeletters.Checked = My.Settings.data882
        IgnoreErrors1.Checked = My.Settings.data883
        Dev1TerminatorEnable.Checked = My.Settings.data884
        CheckBoxSendBlockingDev1.Checked = My.Settings.data885
        Dev1STBMask.Text = My.Settings.data886
        Div1000Dev1.Checked = My.Settings.data887
        Dev13457Aseven.Checked = My.Settings.data888
        Dev1TerminatorEnable2.Checked = My.Settings.data889
        Dev1K2001isolatedata.Checked = My.Settings.data890
        Dev1K2001isolatedataCHAR.Text = My.Settings.data891
        Mult1000Dev1.Checked = My.Settings.data892
        Dev1Timeout.Text = My.Settings.data893
        Dev1delayop.Text = My.Settings.data894
        txtq1d.Text = My.Settings.data895
        Dev1pauseDurationInSeconds.Text = My.Settings.data896
        Dev1runStopwatchEveryInMins.Text = My.Settings.data897
        Dev1IntEnable.Checked = My.Settings.data898
        Dev1Regex.Checked = My.Settings.data899
        Dev1DecimalNumDPs.Text = My.Settings.data900
        txtOperationDev1.Text = My.Settings.data902

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 16

    End Sub

    Private Sub LoadDev1Profile_18()


        txtname1.Text = My.Settings.data903
        CommandStart1.Text = My.Settings.data904
        CommandStart1run.Text = My.Settings.data905
        CommandStop1.Text = My.Settings.data906
        lstIntf1.SelectedIndex = My.Settings.data907
        txtaddr1.Text = My.Settings.data908
        Dev1SampleRate.Text = My.Settings.data909
        Dev1PollingEnable.Checked = My.Settings.data910
        Dev1removeletters.Checked = My.Settings.data911
        IgnoreErrors1.Checked = My.Settings.data912
        Dev1TerminatorEnable.Checked = My.Settings.data913
        CheckBoxSendBlockingDev1.Checked = My.Settings.data914
        Dev1STBMask.Text = My.Settings.data915
        Div1000Dev1.Checked = My.Settings.data916
        Dev13457Aseven.Checked = My.Settings.data917
        Dev1TerminatorEnable2.Checked = My.Settings.data918
        Dev1K2001isolatedata.Checked = My.Settings.data919
        Dev1K2001isolatedataCHAR.Text = My.Settings.data920
        Mult1000Dev1.Checked = My.Settings.data921
        Dev1Timeout.Text = My.Settings.data922
        Dev1delayop.Text = My.Settings.data923
        txtq1d.Text = My.Settings.data924
        Dev1pauseDurationInSeconds.Text = My.Settings.data925
        Dev1runStopwatchEveryInMins.Text = My.Settings.data926
        Dev1IntEnable.Checked = My.Settings.data927
        Dev1Regex.Checked = My.Settings.data928
        Dev1DecimalNumDPs.Text = My.Settings.data929
        txtOperationDev1.Text = My.Settings.data931

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 17

    End Sub

    Private Sub LoadDev1Profile_19()


        txtname1.Text = My.Settings.data932
        CommandStart1.Text = My.Settings.data933
        CommandStart1run.Text = My.Settings.data934
        CommandStop1.Text = My.Settings.data935
        lstIntf1.SelectedIndex = My.Settings.data936
        txtaddr1.Text = My.Settings.data937
        Dev1SampleRate.Text = My.Settings.data938
        Dev1PollingEnable.Checked = My.Settings.data939
        Dev1removeletters.Checked = My.Settings.data940
        IgnoreErrors1.Checked = My.Settings.data941
        Dev1TerminatorEnable.Checked = My.Settings.data942
        CheckBoxSendBlockingDev1.Checked = My.Settings.data943
        Dev1STBMask.Text = My.Settings.data944
        Div1000Dev1.Checked = My.Settings.data945
        Dev13457Aseven.Checked = My.Settings.data946
        Dev1TerminatorEnable2.Checked = My.Settings.data947
        Dev1K2001isolatedata.Checked = My.Settings.data948
        Dev1K2001isolatedataCHAR.Text = My.Settings.data949
        Mult1000Dev1.Checked = My.Settings.data950
        Dev1Timeout.Text = My.Settings.data951
        Dev1delayop.Text = My.Settings.data952
        txtq1d.Text = My.Settings.data953
        Dev1pauseDurationInSeconds.Text = My.Settings.data954
        Dev1runStopwatchEveryInMins.Text = My.Settings.data955
        Dev1IntEnable.Checked = My.Settings.data956
        Dev1Regex.Checked = My.Settings.data957
        Dev1DecimalNumDPs.Text = My.Settings.data958
        txtOperationDev1.Text = My.Settings.data960

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 18

    End Sub

    Private Sub LoadDev1Profile_20()


        txtname1.Text = My.Settings.data961
        CommandStart1.Text = My.Settings.data962
        CommandStart1run.Text = My.Settings.data963
        CommandStop1.Text = My.Settings.data964
        lstIntf1.SelectedIndex = My.Settings.data965
        txtaddr1.Text = My.Settings.data966
        Dev1SampleRate.Text = My.Settings.data967
        Dev1PollingEnable.Checked = My.Settings.data968
        Dev1removeletters.Checked = My.Settings.data969
        IgnoreErrors1.Checked = My.Settings.data970
        Dev1TerminatorEnable.Checked = My.Settings.data971
        CheckBoxSendBlockingDev1.Checked = My.Settings.data972
        Dev1STBMask.Text = My.Settings.data973
        Div1000Dev1.Checked = My.Settings.data974
        Dev13457Aseven.Checked = My.Settings.data975
        Dev1TerminatorEnable2.Checked = My.Settings.data976
        Dev1K2001isolatedata.Checked = My.Settings.data977
        Dev1K2001isolatedataCHAR.Text = My.Settings.data978
        Mult1000Dev1.Checked = My.Settings.data979
        Dev1Timeout.Text = My.Settings.data980
        Dev1delayop.Text = My.Settings.data981
        txtq1d.Text = My.Settings.data982
        Dev1pauseDurationInSeconds.Text = My.Settings.data983
        Dev1runStopwatchEveryInMins.Text = My.Settings.data984
        Dev1IntEnable.Checked = My.Settings.data985
        Dev1Regex.Checked = My.Settings.data986
        Dev1DecimalNumDPs.Text = My.Settings.data987
        txtOperationDev1.Text = My.Settings.data989

        If Not _suppressDev1Sync Then cboDev1Device.SelectedIndex = 19

    End Sub




    ' Device 2 profile 1
    Private Sub LoadDev2Profile_1()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 0

    End Sub

    ' Device 2 profile 2
    Private Sub LoadDev2Profile_2()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 1

    End Sub

    ' Device 2 profile 3
    Private Sub LoadDev2Profile_3()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 2

    End Sub

    ' Device 2 profile 4
    Private Sub LoadDev2Profile_4()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 3

    End Sub

    ' Device 2 profile 5
    Private Sub LoadDev2Profile_5()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 4

    End Sub

    ' Device 2 profile 6
    Private Sub LoadDev2Profile_6()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 5

    End Sub

    ' Device 2 profile 7
    Private Sub LoadDev2Profile_7()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 6

    End Sub


    ' Device 2 profile 8
    Private Sub LoadDev2Profile_8()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 7

    End Sub



    Private Sub LoadDev2Profile_9()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 8

    End Sub

    Private Sub LoadDev2Profile_10()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 9

    End Sub

    Private Sub LoadDev2Profile_11()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 10

    End Sub

    Private Sub LoadDev2Profile_12()


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

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 11

    End Sub

    Private Sub LoadDev2Profile_13()


        txtname2.Text = My.Settings.data990
        CommandStart2.Text = My.Settings.data991
        CommandStart2run.Text = My.Settings.data992
        CommandStop2.Text = My.Settings.data993
        lstIntf2.SelectedIndex = My.Settings.data994
        txtaddr2.Text = My.Settings.data995
        Dev2SampleRate.Text = My.Settings.data996
        Dev2PollingEnable.Checked = My.Settings.data997
        Dev2removeletters.Checked = My.Settings.data998
        IgnoreErrors2.Checked = My.Settings.data999
        Dev2TerminatorEnable.Checked = My.Settings.data1000
        CheckBoxSendBlockingDev2.Checked = My.Settings.data1001
        Dev2STBMask.Text = My.Settings.data1002
        Div1000Dev2.Checked = My.Settings.data1003
        Dev23457Aseven.Checked = My.Settings.data1004
        Dev2TerminatorEnable2.Checked = My.Settings.data1005
        Dev2K2001isolatedata.Checked = My.Settings.data1006
        Dev2K2001isolatedataCHAR.Text = My.Settings.data1007
        Mult1000Dev2.Checked = My.Settings.data1008
        Dev2Timeout.Text = My.Settings.data1009
        Dev2delayop.Text = My.Settings.data1010
        txtq2d.Text = My.Settings.data1011
        Dev2pauseDurationInSeconds.Text = My.Settings.data1012
        Dev2runStopwatchEveryInMins.Text = My.Settings.data1013
        Dev2IntEnable.Checked = My.Settings.data1014
        Dev2Regex.Checked = My.Settings.data1015
        Dev2DecimalNumDPs.Text = My.Settings.data1016
        txtOperationDev2.Text = My.Settings.data1018

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 12

    End Sub

    Private Sub LoadDev2Profile_14()


        txtname2.Text = My.Settings.data1019
        CommandStart2.Text = My.Settings.data1020
        CommandStart2run.Text = My.Settings.data1021
        CommandStop2.Text = My.Settings.data1022
        lstIntf2.SelectedIndex = My.Settings.data1023
        txtaddr2.Text = My.Settings.data1024
        Dev2SampleRate.Text = My.Settings.data1025
        Dev2PollingEnable.Checked = My.Settings.data1026
        Dev2removeletters.Checked = My.Settings.data1027
        IgnoreErrors2.Checked = My.Settings.data1028
        Dev2TerminatorEnable.Checked = My.Settings.data1029
        CheckBoxSendBlockingDev2.Checked = My.Settings.data1030
        Dev2STBMask.Text = My.Settings.data1031
        Div1000Dev2.Checked = My.Settings.data1032
        Dev23457Aseven.Checked = My.Settings.data1033
        Dev2TerminatorEnable2.Checked = My.Settings.data1034
        Dev2K2001isolatedata.Checked = My.Settings.data1035
        Dev2K2001isolatedataCHAR.Text = My.Settings.data1036
        Mult1000Dev2.Checked = My.Settings.data1037
        Dev2Timeout.Text = My.Settings.data1038
        Dev2delayop.Text = My.Settings.data1039
        txtq2d.Text = My.Settings.data1040
        Dev2pauseDurationInSeconds.Text = My.Settings.data1041
        Dev2runStopwatchEveryInMins.Text = My.Settings.data1042
        Dev2IntEnable.Checked = My.Settings.data1043
        Dev2Regex.Checked = My.Settings.data1044
        Dev2DecimalNumDPs.Text = My.Settings.data1045
        txtOperationDev2.Text = My.Settings.data1047

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 13

    End Sub

    Private Sub LoadDev2Profile_15()


        txtname2.Text = My.Settings.data1048
        CommandStart2.Text = My.Settings.data1049
        CommandStart2run.Text = My.Settings.data1050
        CommandStop2.Text = My.Settings.data1051
        lstIntf2.SelectedIndex = My.Settings.data1052
        txtaddr2.Text = My.Settings.data1053
        Dev2SampleRate.Text = My.Settings.data1054
        Dev2PollingEnable.Checked = My.Settings.data1055
        Dev2removeletters.Checked = My.Settings.data1056
        IgnoreErrors2.Checked = My.Settings.data1057
        Dev2TerminatorEnable.Checked = My.Settings.data1058
        CheckBoxSendBlockingDev2.Checked = My.Settings.data1059
        Dev2STBMask.Text = My.Settings.data1060
        Div1000Dev2.Checked = My.Settings.data1061
        Dev23457Aseven.Checked = My.Settings.data1062
        Dev2TerminatorEnable2.Checked = My.Settings.data1063
        Dev2K2001isolatedata.Checked = My.Settings.data1064
        Dev2K2001isolatedataCHAR.Text = My.Settings.data1065
        Mult1000Dev2.Checked = My.Settings.data1066
        Dev2Timeout.Text = My.Settings.data1067
        Dev2delayop.Text = My.Settings.data1068
        txtq2d.Text = My.Settings.data1069
        Dev2pauseDurationInSeconds.Text = My.Settings.data1070
        Dev2runStopwatchEveryInMins.Text = My.Settings.data1071
        Dev2IntEnable.Checked = My.Settings.data1072
        Dev2Regex.Checked = My.Settings.data1073
        Dev2DecimalNumDPs.Text = My.Settings.data1074
        txtOperationDev2.Text = My.Settings.data1076

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 14

    End Sub

    Private Sub LoadDev2Profile_16()


        txtname2.Text = My.Settings.data1077
        CommandStart2.Text = My.Settings.data1078
        CommandStart2run.Text = My.Settings.data1079
        CommandStop2.Text = My.Settings.data1080
        lstIntf2.SelectedIndex = My.Settings.data1081
        txtaddr2.Text = My.Settings.data1082
        Dev2SampleRate.Text = My.Settings.data1083
        Dev2PollingEnable.Checked = My.Settings.data1084
        Dev2removeletters.Checked = My.Settings.data1085
        IgnoreErrors2.Checked = My.Settings.data1086
        Dev2TerminatorEnable.Checked = My.Settings.data1087
        CheckBoxSendBlockingDev2.Checked = My.Settings.data1088
        Dev2STBMask.Text = My.Settings.data1089
        Div1000Dev2.Checked = My.Settings.data1090
        Dev23457Aseven.Checked = My.Settings.data1091
        Dev2TerminatorEnable2.Checked = My.Settings.data1092
        Dev2K2001isolatedata.Checked = My.Settings.data1093
        Dev2K2001isolatedataCHAR.Text = My.Settings.data1094
        Mult1000Dev2.Checked = My.Settings.data1095
        Dev2Timeout.Text = My.Settings.data1096
        Dev2delayop.Text = My.Settings.data1097
        txtq2d.Text = My.Settings.data1098
        Dev2pauseDurationInSeconds.Text = My.Settings.data1099
        Dev2runStopwatchEveryInMins.Text = My.Settings.data1100
        Dev2IntEnable.Checked = My.Settings.data1101
        Dev2Regex.Checked = My.Settings.data1102
        Dev2DecimalNumDPs.Text = My.Settings.data1103
        txtOperationDev2.Text = My.Settings.data1105

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 15

    End Sub

    Private Sub LoadDev2Profile_17()


        txtname2.Text = My.Settings.data1106
        CommandStart2.Text = My.Settings.data1107
        CommandStart2run.Text = My.Settings.data1108
        CommandStop2.Text = My.Settings.data1109
        lstIntf2.SelectedIndex = My.Settings.data1110
        txtaddr2.Text = My.Settings.data1111
        Dev2SampleRate.Text = My.Settings.data1112
        Dev2PollingEnable.Checked = My.Settings.data1113
        Dev2removeletters.Checked = My.Settings.data1114
        IgnoreErrors2.Checked = My.Settings.data1115
        Dev2TerminatorEnable.Checked = My.Settings.data1116
        CheckBoxSendBlockingDev2.Checked = My.Settings.data1117
        Dev2STBMask.Text = My.Settings.data1118
        Div1000Dev2.Checked = My.Settings.data1119
        Dev23457Aseven.Checked = My.Settings.data1120
        Dev2TerminatorEnable2.Checked = My.Settings.data1121
        Dev2K2001isolatedata.Checked = My.Settings.data1122
        Dev2K2001isolatedataCHAR.Text = My.Settings.data1123
        Mult1000Dev2.Checked = My.Settings.data1124
        Dev2Timeout.Text = My.Settings.data1125
        Dev2delayop.Text = My.Settings.data1126
        txtq2d.Text = My.Settings.data1127
        Dev2pauseDurationInSeconds.Text = My.Settings.data1128
        Dev2runStopwatchEveryInMins.Text = My.Settings.data1129
        Dev2IntEnable.Checked = My.Settings.data1130
        Dev2Regex.Checked = My.Settings.data1131
        Dev2DecimalNumDPs.Text = My.Settings.data1132
        txtOperationDev2.Text = My.Settings.data1134

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 16

    End Sub

    Private Sub LoadDev2Profile_18()


        txtname2.Text = My.Settings.data1135
        CommandStart2.Text = My.Settings.data1136
        CommandStart2run.Text = My.Settings.data1137
        CommandStop2.Text = My.Settings.data1138
        lstIntf2.SelectedIndex = My.Settings.data1139
        txtaddr2.Text = My.Settings.data1140
        Dev2SampleRate.Text = My.Settings.data1141
        Dev2PollingEnable.Checked = My.Settings.data1142
        Dev2removeletters.Checked = My.Settings.data1143
        IgnoreErrors2.Checked = My.Settings.data1144
        Dev2TerminatorEnable.Checked = My.Settings.data1145
        CheckBoxSendBlockingDev2.Checked = My.Settings.data1146
        Dev2STBMask.Text = My.Settings.data1147
        Div1000Dev2.Checked = My.Settings.data1148
        Dev23457Aseven.Checked = My.Settings.data1149
        Dev2TerminatorEnable2.Checked = My.Settings.data1150
        Dev2K2001isolatedata.Checked = My.Settings.data1151
        Dev2K2001isolatedataCHAR.Text = My.Settings.data1152
        Mult1000Dev2.Checked = My.Settings.data1153
        Dev2Timeout.Text = My.Settings.data1154
        Dev2delayop.Text = My.Settings.data1155
        txtq2d.Text = My.Settings.data1156
        Dev2pauseDurationInSeconds.Text = My.Settings.data1157
        Dev2runStopwatchEveryInMins.Text = My.Settings.data1158
        Dev2IntEnable.Checked = My.Settings.data1159
        Dev2Regex.Checked = My.Settings.data1160
        Dev2DecimalNumDPs.Text = My.Settings.data1161
        txtOperationDev2.Text = My.Settings.data1163

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 17

    End Sub

    Private Sub LoadDev2Profile_19()


        txtname2.Text = My.Settings.data1164
        CommandStart2.Text = My.Settings.data1165
        CommandStart2run.Text = My.Settings.data1166
        CommandStop2.Text = My.Settings.data1167
        lstIntf2.SelectedIndex = My.Settings.data1168
        txtaddr2.Text = My.Settings.data1169
        Dev2SampleRate.Text = My.Settings.data1170
        Dev2PollingEnable.Checked = My.Settings.data1171
        Dev2removeletters.Checked = My.Settings.data1172
        IgnoreErrors2.Checked = My.Settings.data1173
        Dev2TerminatorEnable.Checked = My.Settings.data1174
        CheckBoxSendBlockingDev2.Checked = My.Settings.data1175
        Dev2STBMask.Text = My.Settings.data1176
        Div1000Dev2.Checked = My.Settings.data1177
        Dev23457Aseven.Checked = My.Settings.data1178
        Dev2TerminatorEnable2.Checked = My.Settings.data1179
        Dev2K2001isolatedata.Checked = My.Settings.data1180
        Dev2K2001isolatedataCHAR.Text = My.Settings.data1181
        Mult1000Dev2.Checked = My.Settings.data1182
        Dev2Timeout.Text = My.Settings.data1183
        Dev2delayop.Text = My.Settings.data1184
        txtq2d.Text = My.Settings.data1185
        Dev2pauseDurationInSeconds.Text = My.Settings.data1186
        Dev2runStopwatchEveryInMins.Text = My.Settings.data1187
        Dev2IntEnable.Checked = My.Settings.data1188
        Dev2Regex.Checked = My.Settings.data1189
        Dev2DecimalNumDPs.Text = My.Settings.data1190
        txtOperationDev2.Text = My.Settings.data1192

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 18

    End Sub

    Private Sub LoadDev2Profile_20()


        txtname2.Text = My.Settings.data1193
        CommandStart2.Text = My.Settings.data1194
        CommandStart2run.Text = My.Settings.data1195
        CommandStop2.Text = My.Settings.data1196
        lstIntf2.SelectedIndex = My.Settings.data1197
        txtaddr2.Text = My.Settings.data1198
        Dev2SampleRate.Text = My.Settings.data1199
        Dev2PollingEnable.Checked = My.Settings.data1200
        Dev2removeletters.Checked = My.Settings.data1201
        IgnoreErrors2.Checked = My.Settings.data1202
        Dev2TerminatorEnable.Checked = My.Settings.data1203
        CheckBoxSendBlockingDev2.Checked = My.Settings.data1204
        Dev2STBMask.Text = My.Settings.data1205
        Div1000Dev2.Checked = My.Settings.data1206
        Dev23457Aseven.Checked = My.Settings.data1207
        Dev2TerminatorEnable2.Checked = My.Settings.data1208
        Dev2K2001isolatedata.Checked = My.Settings.data1209
        Dev2K2001isolatedataCHAR.Text = My.Settings.data1210
        Mult1000Dev2.Checked = My.Settings.data1211
        Dev2Timeout.Text = My.Settings.data1212
        Dev2delayop.Text = My.Settings.data1213
        txtq2d.Text = My.Settings.data1214
        Dev2pauseDurationInSeconds.Text = My.Settings.data1215
        Dev2runStopwatchEveryInMins.Text = My.Settings.data1216
        Dev2IntEnable.Checked = My.Settings.data1217
        Dev2Regex.Checked = My.Settings.data1218
        Dev2DecimalNumDPs.Text = My.Settings.data1219
        txtOperationDev2.Text = My.Settings.data1221

        If Not _suppressDev2Sync Then cboDev2Device.SelectedIndex = 19

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