
' PDVS2mini Auto Calibration
' SerialPort1

' Yes yes yes.....some rather looooong SUBs and repeated ones at that!, I need to get around to re-writing those and make it all more efficient and an easier read and modify!

Imports IODevices
Imports System.Threading
Imports System.Runtime.InteropServices

' PDFsharp must be installed.
' To use PDFsharp, you need to install its NuGet package. Right-click on your project in the Solution Explorer, select "Manage NuGet Packages," search
' for "PDFsharp" and "PDFsharp-MigraDoc" (for document layout), then click "Install."
' I believe I installed V1.50.5147.0 as the later version wouldn't install.
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Imports PdfSharp.Drawing.Layout
#If WORD_INSTALLED Then
Imports Microsoft.Office.Interop.Word
#End If
Imports System.Drawing.Printing
Imports System.Net.Mime.MediaTypeNames
Imports System.IO.Ports
Imports System.IO

Imports System.Management


Partial Class Formtest

    Dim Calpass1 As Boolean = False
    Dim Calpass2 As Boolean = False
    Dim Pcounter As Integer = 0
    Dim Tweak As Boolean = True
    Dim TweakOneshot As Boolean = True
    Dim TweakVar As Integer = 999
    Dim VoltageSet As Double
    Dim VoltageCalAccuracy As Double

    ' Auto-Cal vars
    Dim myString As String
    Dim DACcalparam As Integer = 0

    'Auto-Cal vars - change these to suit batch of mini's so that cal is faster........not used now as there are now textboxes on the tab for this.
    'Dim DacAutoCal0 As Integer = 2434
    'Dim DacAutoCal1 As Integer = 101837
    'Dim DacAutoCal2 As Integer = 201168
    'Dim DacAutoCal3 As Integer = 300524
    'Dim DacAutoCal4 As Integer = 399880
    'Dim DacAutoCal5 As Integer = 499241
    'Dim DacAutoCal6 As Integer = 598597
    'Dim DacAutoCal7 As Integer = 697953
    'Dim DacAutoCal8 As Integer = 797314
    'Dim DacAutoCal9 As Integer = 896670
    'Dim DacAutoCal10 As Integer = 996026
    Dim DacAutoCal0 As Integer
    Dim DacAutoCal1 As Integer
    Dim DacAutoCal2 As Integer
    Dim DacAutoCal3 As Integer
    Dim DacAutoCal4 As Integer
    Dim DacAutoCal5 As Integer
    Dim DacAutoCal6 As Integer
    Dim DacAutoCal7 As Integer
    Dim DacAutoCal8 As Integer
    Dim DacAutoCal9 As Integer
    Dim DacAutoCal10 As Integer
    Dim DacAutoCal11 As Integer

    Private Sub PDVS2miniSave_Click(sender As Object, e As EventArgs) Handles PDVS2miniSave.Click

        My.Settings.data469 = TextBox3458Asn.Text       ' 3458A serial number
        My.Settings.data470 = TextBoxUser.Text          ' User/Company
        My.Settings.data471 = TextBoxLOWSHUT.Text
        My.Settings.data472 = TextBoxCENABLE.Text
        My.Settings.data473 = TextBoxOLMA.Text
        My.Settings.data474 = TextBoxFULLMA.Text
        My.Settings.data475 = TextBoxSERIAL.Text
        My.Settings.data476 = TextBoxSOAK.Text
        My.Settings.data477 = TextBoxDC.Text
        My.Settings.data478 = TextBoxCD.Text
        My.Settings.data479 = CalStep.Text
        My.Settings.data480 = CalAccuracy.Text
        My.Settings.data481 = CalStepFinal.Text
        My.Settings.data482 = CalAccuracyFinal.Text
        My.Settings.data483 = PDVS2delay.Text
        My.Settings.data333 = comPort_ComboBox.SelectedItem
        My.Settings.data503 = WryTech.Checked

        LabelCalCount.Text = "Settings saved"

    End Sub

    Private Sub ButtonLoadeDefs_Click(sender As Object, e As EventArgs) Handles ButtonLoadDefs.Click

        ' Load default counts from app
        Default0.Text = My.Settings.data260
        Default1.Text = My.Settings.data261
        Default2.Text = My.Settings.data262
        Default3.Text = My.Settings.data263
        Default4.Text = My.Settings.data264
        Default5.Text = My.Settings.data265
        Default6.Text = My.Settings.data266
        Default7.Text = My.Settings.data267
        Default8.Text = My.Settings.data268
        Default9.Text = My.Settings.data269
        Default10.Text = My.Settings.data270

        If WryTech.Checked = True Then
            Default11.Text = My.Settings.data502
        End If

        DacAutoCal0 = Val(Default0.Text)
        DacAutoCal1 = Val(Default1.Text)
        DacAutoCal2 = Val(Default2.Text)
        DacAutoCal3 = Val(Default3.Text)
        DacAutoCal4 = Val(Default4.Text)
        DacAutoCal5 = Val(Default5.Text)
        DacAutoCal6 = Val(Default6.Text)
        DacAutoCal7 = Val(Default7.Text)
        DacAutoCal8 = Val(Default8.Text)
        DacAutoCal9 = Val(Default9.Text)
        DacAutoCal10 = Val(Default10.Text)

        If WryTech.Checked = True Then
            DacAutoCal11 = Val(Default11.Text)
        End If

    End Sub

    Private Sub ButtonSaveDefs_Click(sender As Object, e As EventArgs) Handles ButtonSaveDefs.Click

        ' Save default counts to app
        My.Settings.data260 = Default0.Text
        My.Settings.data261 = Default1.Text
        My.Settings.data262 = Default2.Text
        My.Settings.data263 = Default3.Text
        My.Settings.data264 = Default4.Text
        My.Settings.data265 = Default5.Text
        My.Settings.data266 = Default6.Text
        My.Settings.data267 = Default7.Text
        My.Settings.data268 = Default8.Text
        My.Settings.data269 = Default9.Text
        My.Settings.data270 = Default10.Text

        If WryTech.Checked = True Then
            My.Settings.data502 = Default11.Text
        End If

        ' Overwrite the settings as read from the PDVS2mini
        'LabeldacZero0Cal.Text = Default0.Text
        'LabeldacSpan0Cal.Text = Default1.Text
        'LabeldacSpan1Cal.Text = Default2.Text
        'LabeldacSpan2Cal.Text = Default3.Text
        'LabeldacSpan3Cal.Text = Default4.Text
        'LabeldacSpan4Cal.Text = Default5.Text
        'LabeldacSpan5Cal.Text = Default6.Text
        'LabeldacSpan6Cal.Text = Default7.Text
        'LabeldacSpan7Cal.Text = Default8.Text
        'LabeldacSpan8Cal.Text = Default9.Text
        'LabeldacSpan9Cal.Text = Default10.Text

    End Sub

    Private Sub precal_BTN_Click(sender As Object, e As EventArgs) Handles precal_BTN.Click

        ' Set Pre-Cal

        Me.Timer2.Start()    ' start GPIB activity if not already
        ButtonDev1Run.Text = "Stop"

        If Me.SerialPort1.IsOpen = True Then

            If dev1 IsNot Nothing Then      ' Device 1 is open

                LabelCalCount.Text = "Comms ready, GPIB ready"
                System.Threading.Thread.Sleep(Val(1000))

                getcal_BTN.Enabled = False ' disable button
                precal_BTN.Enabled = False ' disable Full Auto-Cal button
                CalOnExisting.Enabled = False ' disable Auto-Cal on existing button
                ButtonSetPrecalvars.Enabled = False ' disable Pre-Cal button
                ButtonSetXYcalvars.Enabled = False ' disable button
                SavePDVS2Eprom.Enabled = False ' disable button
                ButtonSetRetrievedVars.Enabled = False
                ButtonAutoSET.Enabled = False
                ButtonAutomV.Enabled = False

                Me.Timer2.Stop()    ' stop GPIB activity
                ButtonDev1Run.Text = "Run"

                DACcalparam = 0
                Tweak = True
                TweakVar = 999
                TweakOneshot = True

                OutVdc.Text = LabelOutputVFeedMult.Text
                BattVdc.Text = LabelBatteryVFeedMult.Text
                ChargeI.Text = LabelBatteryMonICMult.Text

                'LabelDeltaV.Text = "####"

                LabelCalCount.Text = "Setting pre-cal values"

                Me.Refresh()

                ENotationDecimal.Checked = True
                CheckboxEnableLOG.Checked = True

                ' pre-load saved ballpark cal counts
                'DacAutoCal0 = Val(Default0.Text)
                'DacAutoCal1 = Val(Default1.Text)
                'DacAutoCal2 = Val(Default2.Text)
                'DacAutoCal3 = Val(Default3.Text)
                'DacAutoCal4 = Val(Default4.Text)
                'DacAutoCal5 = Val(Default5.Text)
                'DacAutoCal6 = Val(Default6.Text)
                'DacAutoCal7 = Val(Default7.Text)
                'DacAutoCal8 = Val(Default8.Text)
                'DacAutoCal9 = Val(Default9.Text)
                'DacAutoCal10 = Val(Default10.Text)

                ' reset delatV
                LabeldacZero0Delta.Text = ""
                LabeldacSpan0Delta.Text = ""
                LabeldacSpan1Delta.Text = ""
                LabeldacSpan2Delta.Text = ""
                LabeldacSpan3Delta.Text = ""
                LabeldacSpan4Delta.Text = ""
                LabeldacSpan5Delta.Text = ""
                LabeldacSpan6Delta.Text = ""
                LabeldacSpan7Delta.Text = ""
                LabeldacSpan8Delta.Text = ""
                LabeldacSpan9Delta.Text = ""

                If WryTech.Checked = True Then
                    LabeldacSpan10Delta.Text = ""
                End If

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish

                myString = DacAutoCal0.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacZero0," + myString + ">")    ' Write data to the send buffer
                DacZero0.Text = DacAutoCal0
                Me.Refresh()

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish

                myString = DacAutoCal1.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan0," + myString + ">")    ' Write data to the send buffer
                DacSpan0.Text = DacAutoCal1
                Me.Refresh()

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                myString = DacAutoCal2.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan1," + myString + ">")    ' Write data to the send buffer
                DacSpan1.Text = DacAutoCal2
                Me.Refresh()

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                myString = DacAutoCal3.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan2," + myString + ">")    ' Write data to the send buffer
                DacSpan2.Text = DacAutoCal3
                Me.Refresh()

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                myString = DacAutoCal4.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan3," + myString + ">")    ' Write data to the send buffer
                DacSpan3.Text = DacAutoCal4
                Me.Refresh()

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                myString = DacAutoCal5.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan4," + myString + ">")    ' Write data to the send buffer
                DacSpan4.Text = DacAutoCal5
                Me.Refresh()

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                myString = DacAutoCal6.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan5," + myString + ">")    ' Write data to the send buffer
                DacSpan5.Text = DacAutoCal6
                Me.Refresh()

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                myString = DacAutoCal7.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan6," + myString + ">")    ' Write data to the send buffer
                DacSpan6.Text = DacAutoCal7
                Me.Refresh()

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                myString = DacAutoCal8.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan7," + myString + ">")    ' Write data to the send buffer
                DacSpan7.Text = DacAutoCal8
                Me.Refresh()

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                myString = DacAutoCal9.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan8," + myString + ">")    ' Write data to the send buffer
                DacSpan8.Text = DacAutoCal9
                Me.Refresh()

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                myString = DacAutoCal10.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan9," + myString + ">")    ' Write data to the send buffer
                DacSpan9.Text = DacAutoCal10
                Call TxLEDon()
                Me.Refresh()

                If WryTech.Checked = True Then
                    myString = DacAutoCal11.ToString()
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<DacSpan10," + myString + ">")    ' Write data to the send buffer
                    DacSpan10.Text = DacAutoCal11
                    Call TxLEDon()
                    Me.Refresh()
                End If

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                ' preset 0.0000Vdc
                KeyVoltage.Text = "0.00000"
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                DACcalparam = 0
                LabelCalCount.Text = "Done"
                KeyVoltage.Text = ""

                Me.Timer2.Start()   ' restart GPIB activity
                ButtonDev1Run.Text = "Stop"
                ' set GPIB to NPLC 5
                dev1.SendAsync("NPLC 5", True)
                NPLC.Text = "5"
                Me.Timer2.Stop()
                ButtonDev1Run.Text = "Run"
                Dev1SampleRate.Text = "0.4"
                Dev1TimerDuration = Dev1SampleRate.Text
                Me.Timer2.Interval = Dev1TimerDuration * 1000
                Me.Timer2.Start()
                ButtonDev1Run.Text = "Stop"

                Me.Refresh()
                ' Pre-cal end


                ' Now set up & run the actual calibration
                ENotationDecimal.Checked = True
                CheckboxEnableLOG.Checked = True

                DACcalparam = 0
                Tweak = True
                TweakVar = 999
                TweakOneshot = True

                DacAutoCal0 = Val(LabeldacZero0Cal.Text)
                DacAutoCal1 = Val(LabeldacSpan0Cal.Text)
                DacAutoCal2 = Val(LabeldacSpan1Cal.Text)
                DacAutoCal3 = Val(LabeldacSpan2Cal.Text)
                DacAutoCal4 = Val(LabeldacSpan3Cal.Text)
                DacAutoCal5 = Val(LabeldacSpan4Cal.Text)
                DacAutoCal6 = Val(LabeldacSpan5Cal.Text)
                DacAutoCal7 = Val(LabeldacSpan6Cal.Text)
                DacAutoCal8 = Val(LabeldacSpan7Cal.Text)
                DacAutoCal9 = Val(LabeldacSpan8Cal.Text)
                DacAutoCal10 = Val(LabeldacSpan9Cal.Text)

                If WryTech.Checked = True Then
                    DacAutoCal11 = Val(LabeldacSpan10Cal.Text)
                End If

                ' Set DeltaV to ####
                LabeldacZero0Delta.Text = ""
                LabeldacSpan0Delta.Text = ""
                LabeldacSpan1Delta.Text = ""
                LabeldacSpan2Delta.Text = ""
                LabeldacSpan3Delta.Text = ""
                LabeldacSpan4Delta.Text = ""
                LabeldacSpan5Delta.Text = ""
                LabeldacSpan6Delta.Text = ""
                LabeldacSpan7Delta.Text = ""
                LabeldacSpan8Delta.Text = ""
                LabeldacSpan9Delta.Text = ""

                If WryTech.Checked = True Then
                    LabeldacSpan10Delta.Text = ""
                End If

                ' preset 0.0000Vdc
                KeyVoltage.Text = "0.00000"
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                KeyVoltage.Text = ""

                Me.Timer2.Start()   ' restart GPIB activity
                ButtonDev1Run.Text = "Stop"
                ' set GPIB to NPLC 5
                dev1.SendAsync("NPLC 5", True)
                NPLC.Text = "5"
                Me.Timer2.Stop()
                ButtonDev1Run.Text = "Run"
                Dev1SampleRate.Text = "0.4"
                Dev1TimerDuration = Dev1SampleRate.Text
                Me.Timer2.Interval = Dev1TimerDuration * 1000
                Me.Timer2.Start()
                ButtonDev1Run.Text = "Stop"

                Pcounter = 0

                Tweak = True

                Me.Timer7.Stop()
                Me.Timer7.Interval = 400
                Me.Timer7.Start()       ' This runs PDVS2mini()

            Else

                LabelCalCount.Text = "Comms ready, GPIB not ready"
                Me.Timer2.Stop()    ' stop GPIB activity
                ButtonDev1Run.Text = "Run"

            End If

        Else

            LabelCalCount.Text = "Comms not ready"
            Me.Timer2.Stop()    ' stop GPIB activity
            ButtonDev1Run.Text = "Run"

        End If

    End Sub

    Private Sub CalOnExisting_Click(sender As Object, e As EventArgs) Handles CalOnExisting.Click

        ' Run auto cal on existing data

        If Me.SerialPort1.IsOpen = True And DacZero0.Text <> "" And DacSpan9.Text <> "" Then       ' only allow if serial port open and we have some data retrieved

            If dev1 IsNot Nothing Then      ' Device 1 is open

                LabelCalCount.Text = "Comms ready, GPIB ready"
                System.Threading.Thread.Sleep(Val(1000))

                getcal_BTN.Enabled = False ' disable button
                precal_BTN.Enabled = False ' disable Full Auto-Cal button
                CalOnExisting.Enabled = False ' disable Auto-Cal on existing button
                ButtonSetPrecalvars.Enabled = False ' disable Pre-Cal button
                ButtonSetXYcalvars.Enabled = False ' disable button
                SavePDVS2Eprom.Enabled = False ' disable button
                ButtonSetRetrievedVars.Enabled = False
                ButtonAutoSET.Enabled = False
                ButtonAutomV.Enabled = False

                DACcalparam = 0
                Tweak = True
                TweakVar = 999
                TweakOneshot = True

                DacAutoCal0 = Val(LabeldacZero0Cal.Text)
                DacAutoCal1 = Val(LabeldacSpan0Cal.Text)
                DacAutoCal2 = Val(LabeldacSpan1Cal.Text)
                DacAutoCal3 = Val(LabeldacSpan2Cal.Text)
                DacAutoCal4 = Val(LabeldacSpan3Cal.Text)
                DacAutoCal5 = Val(LabeldacSpan4Cal.Text)
                DacAutoCal6 = Val(LabeldacSpan5Cal.Text)
                DacAutoCal7 = Val(LabeldacSpan6Cal.Text)
                DacAutoCal8 = Val(LabeldacSpan7Cal.Text)
                DacAutoCal9 = Val(LabeldacSpan8Cal.Text)
                DacAutoCal10 = Val(LabeldacSpan9Cal.Text)

                If WryTech.Checked = True Then
                    DacAutoCal11 = Val(LabeldacSpan10Cal.Text)
                End If

                ' Set DeltaV to ####
                LabeldacZero0Delta.Text = ""
                LabeldacSpan0Delta.Text = ""
                LabeldacSpan1Delta.Text = ""
                LabeldacSpan2Delta.Text = ""
                LabeldacSpan3Delta.Text = ""
                LabeldacSpan4Delta.Text = ""
                LabeldacSpan5Delta.Text = ""
                LabeldacSpan6Delta.Text = ""
                LabeldacSpan7Delta.Text = ""
                LabeldacSpan8Delta.Text = ""
                LabeldacSpan9Delta.Text = ""

                If WryTech.Checked = True Then
                    LabeldacSpan10Delta.Text = ""
                End If

                ' preset 0.0000Vdc
                KeyVoltage.Text = "0.00000"
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer

                System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

                KeyVoltage.Text = ""

                Me.Timer2.Start()   ' restart GPIB activity
                ButtonDev1Run.Text = "Stop"

                Pcounter = 0

                ' set params so can jump straight to final tweak
                DACcalparam = 15
                Tweak = True
                TweakOneshot = True
                'txtr1aBIG.Text = "0.0"
                'VoltageSet = 0.0
                KeyVoltage.Text = "0.00000"

                Me.Timer2.Start()    ' start GPIB activity if not already
                ButtonDev1Run.Text = "Stop"

                Me.Timer7.Stop()
                Me.Timer7.Interval = 400
                Me.Timer7.Start()       ' This runs PDVS2mini()

            Else

                LabelCalCount.Text = "Comms ready, GPIB not ready"

            End If

        Else

            LabelCalCount.Text = "Comms not ready or data not received"

        End If

    End Sub

    Private Sub ButtonSetPrecalvars_Click(sender As Object, e As EventArgs) Handles ButtonSetPrecalvars.Click

        If Me.SerialPort1.IsOpen = True Then       ' only allow if serial port open and we have some data retrieved

            ButtonSetPrecalvars.Enabled = False ' disable Pre-Cal button

            'LabelDeltaV.Text = "####"

            LabelCalCount.Text = "Sending pre-cal calibration"

            ' pre-load ballpark cal counts from app
            DacAutoCal0 = Val(Default0.Text)
            DacAutoCal1 = Val(Default1.Text)
            DacAutoCal2 = Val(Default2.Text)
            DacAutoCal3 = Val(Default3.Text)
            DacAutoCal4 = Val(Default4.Text)
            DacAutoCal5 = Val(Default5.Text)
            DacAutoCal6 = Val(Default6.Text)
            DacAutoCal7 = Val(Default7.Text)
            DacAutoCal8 = Val(Default8.Text)
            DacAutoCal9 = Val(Default9.Text)
            DacAutoCal10 = Val(Default10.Text)

            If WryTech.Checked = True Then
                DacAutoCal11 = Val(Default11.Text)
            End If

            'DacAutoCal0 = Val(LabeldacZero0Cal.Text)       ' this used to base pre-cal on current values in the PDVS2mini
            'DacAutoCal1 = Val(LabeldacSpan0Cal.Text)
            'DacAutoCal2 = Val(LabeldacSpan1Cal.Text)
            'DacAutoCal3 = Val(LabeldacSpan2Cal.Text)
            'DacAutoCal4 = Val(LabeldacSpan3Cal.Text)
            'DacAutoCal5 = Val(LabeldacSpan4Cal.Text)
            'DacAutoCal6 = Val(LabeldacSpan5Cal.Text)
            'DacAutoCal7 = Val(LabeldacSpan6Cal.Text)
            'DacAutoCal8 = Val(LabeldacSpan7Cal.Text)
            'DacAutoCal9 = Val(LabeldacSpan8Cal.Text)
            'DacAutoCal10 = Val(LabeldacSpan9Cal.Text)

            ' Set DeltaV to #######
            LabeldacZero0Delta.Text = ""
            LabeldacSpan0Delta.Text = ""
            LabeldacSpan1Delta.Text = ""
            LabeldacSpan2Delta.Text = ""
            LabeldacSpan3Delta.Text = ""
            LabeldacSpan4Delta.Text = ""
            LabeldacSpan5Delta.Text = ""
            LabeldacSpan6Delta.Text = ""
            LabeldacSpan7Delta.Text = ""
            LabeldacSpan8Delta.Text = ""
            LabeldacSpan9Delta.Text = ""

            If WryTech.Checked = True Then
                LabeldacSpan10Delta.Text = ""
            End If

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish

            myString = DacAutoCal0.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacZero0," + myString + ">")    ' Write data to the send buffer
            DacZero0.Text = DacAutoCal0
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish

            myString = DacAutoCal1.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan0," + myString + ">")    ' Write data to the send buffer
            DacSpan0.Text = DacAutoCal1
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal2.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan1," + myString + ">")    ' Write data to the send buffer
            DacSpan1.Text = DacAutoCal2
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal3.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan2," + myString + ">")    ' Write data to the send buffer
            DacSpan2.Text = DacAutoCal3
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal4.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan3," + myString + ">")    ' Write data to the send buffer
            DacSpan3.Text = DacAutoCal4
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal5.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan4," + myString + ">")    ' Write data to the send buffer
            DacSpan4.Text = DacAutoCal5
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal6.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan5," + myString + ">")    ' Write data to the send buffer
            DacSpan5.Text = DacAutoCal6
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal7.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan6," + myString + ">")    ' Write data to the send buffer
            DacSpan6.Text = DacAutoCal7
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal8.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan7," + myString + ">")    ' Write data to the send buffer
            DacSpan7.Text = DacAutoCal8
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal9.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan8," + myString + ">")    ' Write data to the send buffer
            DacSpan8.Text = DacAutoCal9
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal10.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan9," + myString + ">")    ' Write data to the send buffer
            DacSpan9.Text = DacAutoCal10
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            If WryTech.Checked = True Then
                myString = DacAutoCal11.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan10," + myString + ">")    ' Write data to the send buffer
                DacSpan10.Text = DacAutoCal11
                Me.Refresh()
            End If

            LabelCalCount.Text = "Pre-cal calibration sent"

            ButtonSetPrecalvars.Enabled = True ' enable Pre-Cal button

            ' Play sound to indicate finished
            My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Asterisk)

        Else

            LabelCalCount.Text = "Comms not ready"

        End If



    End Sub

    Private Sub ButtonSetRetrievedVars_Click(sender As Object, e As EventArgs) Handles ButtonSetRetrievedVars.Click

        If Me.SerialPort1.IsOpen = True And DacZero0.Text <> "" Then       ' only allow if serial port open and we have some data retrieved

            ButtonSetRetrievedVars.Enabled = False ' disable button

            'LabelDeltaV.Text = "####"

            LabelCalCount.Text = "Sending current DAC counts (11off)"

            ' pre-load ballpark cal counts
            DacAutoCal0 = DacZero0.Text
            DacAutoCal1 = DacSpan0.Text
            DacAutoCal2 = DacSpan1.Text
            DacAutoCal3 = DacSpan2.Text
            DacAutoCal4 = DacSpan3.Text
            DacAutoCal5 = DacSpan4.Text
            DacAutoCal6 = DacSpan5.Text
            DacAutoCal7 = DacSpan6.Text
            DacAutoCal8 = DacSpan7.Text
            DacAutoCal9 = DacSpan8.Text
            DacAutoCal10 = DacSpan9.Text

            If WryTech.Checked = True Then
                DacAutoCal11 = DacSpan10.Text
            End If


            ' Set DeltaV to #######
            LabeldacZero0Delta.Text = ""
            LabeldacSpan0Delta.Text = ""
            LabeldacSpan1Delta.Text = ""
            LabeldacSpan2Delta.Text = ""
            LabeldacSpan3Delta.Text = ""
            LabeldacSpan4Delta.Text = ""
            LabeldacSpan5Delta.Text = ""
            LabeldacSpan6Delta.Text = ""
            LabeldacSpan7Delta.Text = ""
            LabeldacSpan8Delta.Text = ""
            LabeldacSpan9Delta.Text = ""

            If WryTech.Checked = True Then
                LabeldacSpan10Delta.Text = ""
            End If

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish

            myString = DacAutoCal0.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacZero0," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish

            myString = DacAutoCal1.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan0," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal2.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan1," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal3.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan2," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal4.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan3," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal5.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan4," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal6.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan5," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal7.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan6," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal8.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan7," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal9.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan8," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacAutoCal10.ToString()
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan9," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            If WryTech.Checked = True Then
                myString = DacAutoCal11.ToString()
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan10," + myString + ">")    ' Write data to the send buffer
                Me.Refresh()
            End If


            LabelCalCount.Text = "Done"

            ButtonSetRetrievedVars.Enabled = True ' enable button

        Else

            LabelCalCount.Text = "Comms not ready / No data"

        End If



    End Sub

    Private Sub ButtonSetXYcalvars_Click(sender As Object, e As EventArgs) Handles ButtonSetXYcalvars.Click

        ButtonSetXYcalvars.Enabled = False ' disable button

        ' DacSpan10....?
        If DacZero0.Text <> "" And DacSpan0.Text <> "" And DacSpan1.Text <> "" And DacSpan2.Text <> "" And DacSpan3.Text <> "" And DacSpan4.Text <> "" And DacSpan5.Text <> "" And DacSpan6.Text <> "" And DacSpan7.Text <> "" And DacSpan8.Text <> "" And DacSpan9.Text <> "" Then

            myString = DacZero0.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacZero0," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish

            myString = DacSpan0.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan0," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacSpan1.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan1," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacSpan2.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan2," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacSpan3.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan3," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacSpan4.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan4," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacSpan5.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan5," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacSpan6.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan6," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacSpan7.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan7," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacSpan8.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan8," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            myString = DacSpan9.Text
            Call TxLEDon()
            Me.SerialPort1.WriteLine("<DacSpan9," + myString + ">")    ' Write data to the send buffer
            Me.Refresh()

            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))

            If WryTech.Checked = True Then
                myString = DacSpan10.Text
                Call TxLEDon()
                Me.SerialPort1.WriteLine("<DacSpan10," + myString + ">")    ' Write data to the send buffer
                Me.Refresh()
            End If

            LabelCalCount.Text = "Done"

        Else

            LabelCalCount.Text = "XY data empty?"

        End If

        ButtonSetXYcalvars.Enabled = True ' enable button


    End Sub

    Private Sub PDVS2mini()

        ' Auto-Cal, this runs after Pre-Cal on Timer 7
        ' This sub runs when Timer7 running

        If Me.SerialPort1.IsOpen = True Then

            If dev1 IsNot Nothing Then      ' Device 1 is open

                'Pcounter = Pcounter + 1
                Pcounter += 1
                PDVS2counter.Text = Pcounter
                DACcalparamVAL.Text = DACcalparam

                'Console.WriteLine("ValF = " & inst_value1F)

                If (DACcalparam = 0) Then

                    'autocal_BTN.Enabled = False ' disable Auto-Cal button

                    ' Set 0Vdc
                    LabelCalCount.Text = "Initial calibrating 0Vdc"
                    KeyVoltage.Text = "0.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal0 = DacAutoCal0 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal0 = DacAutoCal0 + Val(CalStep.Text)
                        End If

                        myString = DacAutoCal0.ToString()
                        DacZero0.Text = "<DacZero0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacZero0.Text)    ' Write data to the send buffer
                        DacZero0.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal0 = DacAutoCal0 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal0 = DacAutoCal0 - Val(CalStep.Text)
                        End If

                        'DacAutoCal0 = DacAutoCal0 - Val(CalStep.Text)
                        myString = DacAutoCal0.ToString()
                        DacZero0.Text = "<DacZero0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacZero0.Text)    ' Write data to the send buffer
                        DacZero0.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 1) Then
                    ' Set 1Vdc
                    LabelCalCount.Text = "Initial calibrating 1Vdc"
                    KeyVoltage.Text = "1.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal1 = DacAutoCal1 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal1 = DacAutoCal1 + Val(CalStep.Text)
                        End If

                        myString = DacAutoCal1.ToString()
                        DacSpan0.Text = "<DacSpan0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan0.Text)    ' Write data to the send buffer
                        DacSpan0.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal1 = DacAutoCal1 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal1 = DacAutoCal1 - Val(CalStep.Text)
                        End If

                        myString = DacAutoCal1.ToString()
                        DacSpan0.Text = "<DacSpan0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan0.Text)    ' Write data to the send buffer
                        DacSpan0.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                ' Auto-Calculate the rest
                If (DACcalparam = 2) Then

                    Me.Timer2.Stop()    ' stop GPIB activity
                    ButtonDev1Run.Text = "Run"

                    DacSpan1.Text = ""
                    DacSpan2.Text = ""
                    DacSpan3.Text = ""
                    DacSpan4.Text = ""
                    DacSpan5.Text = ""
                    DacSpan6.Text = ""
                    DacSpan7.Text = ""
                    DacSpan8.Text = ""
                    DacSpan9.Text = ""
                    If WryTech.Checked = True Then
                        DacSpan10.Text = ""
                    End If
                    LabelCalCount.Text = "Initial calibrating 2V to 10V"

                    Me.Refresh()

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    ' 3rd
                    DacAutoCal2 = (DacAutoCal1 - DacAutoCal0) + DacAutoCal1
                    myString = DacAutoCal2.ToString()
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<DacSpan1," + myString + ">")    ' Write data to the send buffer 
                    DacSpan1.Text = myString
                    'txtr1aBIG.Text = Format(inst_value1F, "#0.0000000000")

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    ' 4th
                    DacAutoCal3 = (DacAutoCal2 - DacAutoCal1) + DacAutoCal2
                    myString = DacAutoCal3.ToString()
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<DacSpan2," + myString + ">")    ' Write data to the send buffer 
                    DacSpan2.Text = myString
                    'txtr1aBIG.Text = Format(inst_value1F, "#0.0000000000")

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    ' 5th
                    DacAutoCal4 = (DacAutoCal3 - DacAutoCal2) + DacAutoCal3
                    myString = DacAutoCal4.ToString()
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<DacSpan3," + myString + ">")    ' Write data to the send buffer 
                    DacSpan3.Text = myString
                    'txtr1aBIG.Text = Format(inst_value1F, "#0.0000000000")

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    ' 6th
                    DacAutoCal5 = (DacAutoCal4 - DacAutoCal3) + DacAutoCal4
                    myString = DacAutoCal5.ToString()
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<DacSpan4," + myString + ">")    ' Write data to the send buffer 
                    DacSpan4.Text = myString
                    'txtr1aBIG.Text = Format(inst_value1F, "#0.0000000000")

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    ' 7th
                    DacAutoCal6 = (DacAutoCal5 - DacAutoCal4) + DacAutoCal5
                    myString = DacAutoCal6.ToString()
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<DacSpan5," + myString + ">")    ' Write data to the send buffer 
                    DacSpan5.Text = myString
                    'txtr1aBIG.Text = Format(inst_value1F, "#0.0000000000")

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    ' 8th
                    DacAutoCal7 = (DacAutoCal6 - DacAutoCal5) + DacAutoCal6
                    myString = DacAutoCal7.ToString()
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<DacSpan6," + myString + ">")    ' Write data to the send buffer 
                    DacSpan6.Text = myString
                    'txtr1aBIG.Text = Format(inst_value1F, "#0.0000000000")

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    ' 9th
                    DacAutoCal8 = (DacAutoCal7 - DacAutoCal6) + DacAutoCal7
                    myString = DacAutoCal8.ToString()
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<DacSpan7," + myString + ">")    ' Write data to the send buffer 
                    DacSpan7.Text = myString
                    'txtr1aBIG.Text = Format(inst_value1F, "#0.0000000000")

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    ' 10th
                    DacAutoCal9 = (DacAutoCal8 - DacAutoCal7) + DacAutoCal8
                    myString = DacAutoCal9.ToString()
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<DacSpan8," + myString + ">")    ' Write data to the send buffer 
                    DacSpan8.Text = myString
                    'txtr1aBIG.Text = Format(inst_value1F, "#0.0000000000")

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    ' 11th
                    DacAutoCal10 = (DacAutoCal9 - DacAutoCal8) + DacAutoCal9
                    myString = DacAutoCal10.ToString()
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<DacSpan9," + myString + ">")    ' Write data to the send buffer 
                    DacSpan9.Text = myString
                    'txtr1aBIG.Text = Format(inst_value1F, "#0.0000000000")

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    If WryTech.Checked = True Then
                        ' 12th
                        'DacAutoCal11 = (DacAutoCal10 - DacAutoCal9) + DacAutoCal10
                        DacAutoCal11 = ((DacAutoCal10 - DacAutoCal9) * 0.22222) + DacAutoCal10      ' calc compensated for 10.22222 Vdc
                        myString = DacAutoCal11.ToString()
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine("<DacSpan10," + myString + ">")    ' Write data to the send buffer 
                        DacSpan10.Text = myString
                        'txtr1aBIG.Text = Format(inst_value1F, "#0.0000000000")
                    End If

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    ' Tidy up, send 0.00000Vdc to PDVS2mini
                    KeyVoltage.Text = "0.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    LabelKeyVoltage.Text = "0.00000"

                    DACcalparam = DACcalparam + 1

                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text) + 100)  ' give PDVS2mini time to finish

                    Me.Timer2.Start()
                    ButtonDev1Run.Text = "Stop"

                    'autocal_BTN.Enabled = True ' enable Auto-Cal button - Allow Auto-Cal

                End If




                ' Final Auto-Calculate all setpoints (0-10Vdc)
                If (DACcalparam = 3) Then

                    Me.Refresh()

                    ' Set 0Vdc
                    LabelCalCount.Text = "Calibrating 0Vdc"
                    KeyVoltage.Text = "0.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal0 = DacAutoCal0 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal0 = DacAutoCal0 + Val(CalStep.Text)
                        End If

                        'DacAutoCal0 = DacAutoCal0 + Val(CalStep.Text)
                        myString = DacAutoCal0.ToString()
                        DacZero0.Text = "<DacZero0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacZero0.Text)    ' Write data to the send buffer
                        DacZero0.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal0 = DacAutoCal0 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal0 = DacAutoCal0 - Val(CalStep.Text)
                        End If

                        'DacAutoCal0 = DacAutoCal0 - Val(CalStep.Text)
                        myString = DacAutoCal0.ToString()
                        DacZero0.Text = "<DacZero0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacZero0.Text)    ' Write data to the send buffer
                        DacZero0.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 4) Then
                    ' Set 1Vdc
                    LabelCalCount.Text = "Calibrating 1Vdc"
                    KeyVoltage.Text = "1.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal1 = DacAutoCal1 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal1 = DacAutoCal1 + Val(CalStep.Text)
                        End If

                        'DacAutoCal1 = DacAutoCal1 + Val(CalStep.Text)
                        myString = DacAutoCal1.ToString()
                        DacSpan0.Text = "<DacSpan0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan0.Text)    ' Write data to the send buffer
                        DacSpan0.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal1 = DacAutoCal1 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal1 = DacAutoCal1 - Val(CalStep.Text)
                        End If

                        'DacAutoCal1 = DacAutoCal1 - Val(CalStep.Text)
                        myString = DacAutoCal1.ToString()
                        DacSpan0.Text = "<DacSpan0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan0.Text)    ' Write data to the send buffer
                        DacSpan0.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 5) Then
                    ' Set 2Vdc
                    LabelCalCount.Text = "Calibrating 2Vdc"
                    KeyVoltage.Text = "2.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal2 = DacAutoCal2 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal2 = DacAutoCal2 + Val(CalStep.Text)
                        End If

                        'DacAutoCal2 = DacAutoCal2 + Val(CalStep.Text)
                        myString = DacAutoCal2.ToString()
                        DacSpan1.Text = "<DacSpan1," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan1.Text)    ' Write data to the send buffer
                        DacSpan1.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal2 = DacAutoCal2 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal2 = DacAutoCal2 - Val(CalStep.Text)
                        End If

                        'DacAutoCal2 = DacAutoCal2 - Val(CalStep.Text)
                        myString = DacAutoCal2.ToString()
                        DacSpan1.Text = "<DacSpan1," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan1.Text)    ' Write data to the send buffer
                        DacSpan1.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 6) Then
                    ' Set 3Vdc
                    LabelCalCount.Text = "Calibrating 3Vdc"
                    KeyVoltage.Text = "3.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal3 = DacAutoCal3 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal3 = DacAutoCal3 + Val(CalStep.Text)
                        End If

                        'DacAutoCal3 = DacAutoCal3 + Val(CalStep.Text)
                        myString = DacAutoCal3.ToString()
                        DacSpan2.Text = "<DacSpan2," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan2.Text)    ' Write data to the send buffer
                        DacSpan2.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal3 = DacAutoCal3 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal3 = DacAutoCal3 - Val(CalStep.Text)
                        End If

                        'DacAutoCal3 = DacAutoCal3 - Val(CalStep.Text)
                        myString = DacAutoCal3.ToString()
                        DacSpan2.Text = "<DacSpan2," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan2.Text)    ' Write data to the send buffer
                        DacSpan2.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 7) Then
                    ' Set 4Vdc
                    LabelCalCount.Text = "Calibrating 4Vdc"
                    KeyVoltage.Text = "4.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal4 = DacAutoCal4 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal4 = DacAutoCal4 + Val(CalStep.Text)
                        End If

                        'DacAutoCal4 = DacAutoCal4 + Val(CalStep.Text)
                        myString = DacAutoCal4.ToString()
                        DacSpan3.Text = "<DacSpan3," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan3.Text)    ' Write data to the send buffer
                        DacSpan3.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal4 = DacAutoCal4 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal4 = DacAutoCal4 - Val(CalStep.Text)
                        End If

                        'DacAutoCal4 = DacAutoCal4 - Val(CalStep.Text)
                        myString = DacAutoCal4.ToString()
                        DacSpan3.Text = "<DacSpan3," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan3.Text)    ' Write data to the send buffer
                        DacSpan3.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 8) Then
                    ' Set 5Vdc
                    LabelCalCount.Text = "Calibrating 5Vdc"
                    KeyVoltage.Text = "5.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal5 = DacAutoCal5 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal5 = DacAutoCal5 + Val(CalStep.Text)
                        End If

                        'DacAutoCal5 = DacAutoCal5 + Val(CalStep.Text)
                        myString = DacAutoCal5.ToString()
                        DacSpan4.Text = "<DacSpan4," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan4.Text)    ' Write data to the send buffer
                        DacSpan4.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal5 = DacAutoCal5 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal5 = DacAutoCal5 - Val(CalStep.Text)
                        End If

                        'DacAutoCal5 = DacAutoCal5 - Val(CalStep.Text)
                        myString = DacAutoCal5.ToString()
                        DacSpan4.Text = "<DacSpan4," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan4.Text)    ' Write data to the send buffer
                        DacSpan4.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 9) Then
                    ' Set 6Vdc
                    LabelCalCount.Text = "Calibrating 6Vdc"
                    KeyVoltage.Text = "6.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal6 = DacAutoCal6 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal6 = DacAutoCal6 + Val(CalStep.Text)
                        End If

                        'DacAutoCal6 = DacAutoCal6 + Val(CalStep.Text)
                        myString = DacAutoCal6.ToString()
                        DacSpan5.Text = "<DacSpan5," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan5.Text)    ' Write data to the send buffer
                        DacSpan5.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal6 = DacAutoCal6 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal6 = DacAutoCal6 - Val(CalStep.Text)
                        End If

                        'DacAutoCal6 = DacAutoCal6 - Val(CalStep.Text)
                        myString = DacAutoCal6.ToString()
                        DacSpan5.Text = "<DacSpan5," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan5.Text)    ' Write data to the send buffer
                        DacSpan5.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 10) Then
                    ' Set 7Vdc
                    LabelCalCount.Text = "Calibrating 7Vdc"
                    KeyVoltage.Text = "7.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal7 = DacAutoCal7 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal7 = DacAutoCal7 + Val(CalStep.Text)
                        End If

                        'DacAutoCal7 = DacAutoCal7 + Val(CalStep.Text)
                        myString = DacAutoCal7.ToString()
                        DacSpan6.Text = "<DacSpan6," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan6.Text)    ' Write data to the send buffer
                        DacSpan6.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal7 = DacAutoCal7 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal7 = DacAutoCal7 - Val(CalStep.Text)
                        End If

                        'DacAutoCal7 = DacAutoCal7 - Val(CalStep.Text)
                        myString = DacAutoCal7.ToString()
                        DacSpan6.Text = "<DacSpan6," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan6.Text)    ' Write data to the send buffer
                        DacSpan6.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 11) Then
                    ' Set 8Vdc
                    LabelCalCount.Text = "Calibrating 8Vdc"
                    KeyVoltage.Text = "8.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal8 = DacAutoCal8 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal8 = DacAutoCal8 + Val(CalStep.Text)
                        End If

                        'DacAutoCal8 = DacAutoCal8 + Val(CalStep.Text)
                        myString = DacAutoCal8.ToString()
                        DacSpan7.Text = "<DacSpan7," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan7.Text)    ' Write data to the send buffer
                        DacSpan7.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal8 = DacAutoCal8 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal8 = DacAutoCal8 - Val(CalStep.Text)
                        End If

                        'DacAutoCal8 = DacAutoCal8 - Val(CalStep.Text)
                        myString = DacAutoCal8.ToString()
                        DacSpan7.Text = "<DacSpan7," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan7.Text)    ' Write data to the send buffer
                        DacSpan7.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 12) Then
                    ' Set 9Vdc
                    LabelCalCount.Text = "Calibrating 9Vdc"
                    KeyVoltage.Text = "9.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal9 = DacAutoCal9 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal9 = DacAutoCal9 + Val(CalStep.Text)
                        End If

                        'DacAutoCal9 = DacAutoCal9 + Val(CalStep.Text)
                        myString = DacAutoCal9.ToString()
                        DacSpan8.Text = "<DacSpan8," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan8.Text)    ' Write data to the send buffer
                        DacSpan8.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal9 = DacAutoCal9 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal9 = DacAutoCal9 - Val(CalStep.Text)
                        End If

                        'DacAutoCal9 = DacAutoCal9 - Val(CalStep.Text)
                        myString = DacAutoCal9.ToString()
                        DacSpan8.Text = "<DacSpan8," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan8.Text)    ' Write data to the send buffer
                        DacSpan8.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                If (DACcalparam = 13) Then
                    ' Set 10Vdc
                    LabelCalCount.Text = "Calibrating 10Vdc"
                    KeyVoltage.Text = "10.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal10 = DacAutoCal10 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal10 = DacAutoCal10 + Val(CalStep.Text)
                        End If

                        'DacAutoCal10 = DacAutoCal10 + Val(CalStep.Text)
                        myString = DacAutoCal10.ToString()
                        DacSpan9.Text = "<DacSpan9," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan9.Text)    ' Write data to the send buffer
                        DacSpan9.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then


                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal10 = DacAutoCal10 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal10 = DacAutoCal10 - Val(CalStep.Text)
                        End If

                        'DacAutoCal10 = DacAutoCal10 - Val(CalStep.Text)
                        myString = DacAutoCal10.ToString()
                        DacSpan9.Text = "<DacSpan9," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan9.Text)    ' Write data to the send buffer
                        DacSpan9.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If


                If DACcalparam = 14 And WryTech.Checked = True Then
                    ' Set 10.22222Vdc
                    LabelCalCount.Text = "Calibrating 10.22222Vdc"
                    KeyVoltage.Text = "10.22222"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    Dim VoltageSet As Double = Val(KeyVoltage.Text)
                    Dim VoltageCalAccuracy As Double = Val(CalAccuracy.Text)
                    If (inst_value1F < VoltageSet) Then

                        ' makes steps larger if far away from final value
                        If (VoltageSet - inst_value1F > 0.0003) Then
                            DacAutoCal11 = DacAutoCal11 + (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal11 = DacAutoCal11 + Val(CalStep.Text)
                        End If

                        'DacAutoCal11 = DacAutoCal11 + Val(CalStep.Text)
                        myString = DacAutoCal11.ToString()
                        DacSpan10.Text = "<DacSpan10," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan10.Text)    ' Write data to the send buffer
                        DacSpan10.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If (inst_value1F > VoltageSet) Then


                        ' makes steps larger if far away from final value
                        If (inst_value1F - VoltageSet > 0.0003) Then
                            DacAutoCal11 = DacAutoCal11 - (Val(CalStep.Text) * 2)
                        Else
                            DacAutoCal11 = DacAutoCal11 - Val(CalStep.Text)
                        End If

                        'DacAutoCal11 = DacAutoCal11 - Val(CalStep.Text)
                        myString = DacAutoCal11.ToString()
                        DacSpan10.Text = "<DacSpan10," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan10.Text)    ' Write data to the send buffer
                        DacSpan10.Text = myString
                        txtr1aBIG.Text = Format(inst_value1F, "#0.0########")
                    End If
                    If ((VoltageSet - inst_value1F < VoltageCalAccuracy) And (VoltageSet - inst_value1F > -VoltageCalAccuracy)) Then      ' finish
                        DACcalparam = DACcalparam + 1
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    End If
                End If

                ' If Wrytech PDVS2mini is not used then need to increment DACcalparam counter
                If DACcalparam = 14 And WryTech.Checked = False Then
                    DACcalparam = 15    ' jump past 10.22222
                End If

                ' Final Auto-Cal all setpoints (0-10Vdc)
                ' When this sub is run it alternates between sending to the PDVS2mini and reading from the 3458A via GPIB.
                ' This allows the time for multiple GPIB reads and also to allow the PDVS2mini output to settle.
                ' Tweak - boolean, True to start with and will alternate with False
                ' TweakVar - 0 = do nothing
                '            1 = increase
                '            2 = decrease
                ' The timer for this sub is changed to 2.5secs for final calibration (see Timer7 adjustment in Formtest.vb)

                ' Final Auto-Cal
                ' Run cal on existing jumps in here also
                If (DACcalparam = 15 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    ' set GPIB to NPLC 20
                    dev1.SendAsync("NPLC 20", True)
                    NPLC.Text = "20"
                    Me.Timer2.Stop()
                    ButtonDev1Run.Text = "Run"
                    Dev1SampleRate.Text = "1.0" ' was 1.5
                    Dev1TimerDuration = Dev1SampleRate.Text
                    Me.Timer2.Interval = Dev1TimerDuration * 1000
                    Me.Timer2.Start()
                    ButtonDev1Run.Text = "Stop"

                    KeyVoltage.Text = "0.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish

                    ' and again for luck, because sometimes setting 0Vdc fails here not sure why
                    KeyVoltage.Text = "0.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish

                    TweakOneshot = False
                End If


                ' Set 0Vdc - Final Tweak
                If (DACcalparam = 15 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 0Vdc"
                    KeyVoltage.Text = "0.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal0 = DacAutoCal0 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal0 = DacAutoCal0 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal0 = DacAutoCal0 + Val(CalStepFinal.Text)
                        myString = DacAutoCal0.ToString()
                        DacZero0.Text = "<DacZero0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacZero0.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacZero0.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal0 = DacAutoCal0 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal0 = DacAutoCal0 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal0 = DacAutoCal0 - Val(CalStepFinal.Text)
                        myString = DacAutoCal0.ToString()
                        DacZero0.Text = "<DacZero0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacZero0.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacZero0.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts0.Text = txtr1aBIG.Text
                        volts0.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 15 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If


                ' Set 1Vdc - Final Tweak
                If (DACcalparam = 16 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    KeyVoltage.Text = "1.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    TweakOneshot = False
                End If
                If (DACcalparam = 16 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 1Vdc"
                    KeyVoltage.Text = "1.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        ' makes steps larger if far away from final value
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal1 = DacAutoCal1 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal1 = DacAutoCal1 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal1 = DacAutoCal1 + Val(CalStepFinal.Text)
                        myString = DacAutoCal1.ToString()
                        DacSpan0.Text = "<DacSpan0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan0.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan0.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal1 = DacAutoCal1 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal1 = DacAutoCal1 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal1 = DacAutoCal1 - Val(CalStepFinal.Text)
                        myString = DacAutoCal1.ToString()
                        DacSpan0.Text = "<DacSpan0," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan0.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan0.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts1.Text = txtr1aBIG.Text
                        volts1.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 16 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If

                ' Set 2Vdc - Final Tweak
                If (DACcalparam = 17 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    KeyVoltage.Text = "2.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    TweakOneshot = False
                End If
                If (DACcalparam = 17 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 2Vdc"
                    KeyVoltage.Text = "2.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        ' makes steps larger if far away from final value
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal2 = DacAutoCal2 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal2 = DacAutoCal2 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal2 = DacAutoCal2 + Val(CalStepFinal.Text)
                        myString = DacAutoCal2.ToString()
                        DacSpan1.Text = "<DacSpan1," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan1.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan1.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal2 = DacAutoCal2 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal2 = DacAutoCal2 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal2 = DacAutoCal2 - Val(CalStepFinal.Text)
                        myString = DacAutoCal2.ToString()
                        DacSpan1.Text = "<DacSpan1," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan1.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan1.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts2.Text = txtr1aBIG.Text
                        volts2.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 17 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If

                ' Set 3Vdc - Final Tweak
                If (DACcalparam = 18 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    KeyVoltage.Text = "3.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    TweakOneshot = False
                End If
                If (DACcalparam = 18 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 3Vdc"
                    KeyVoltage.Text = "3.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        ' makes steps larger if far away from final value
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal3 = DacAutoCal3 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal3 = DacAutoCal3 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal3 = DacAutoCal3 + Val(CalStepFinal.Text)
                        myString = DacAutoCal3.ToString()
                        DacSpan2.Text = "<DacSpan2," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan2.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan2.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal3 = DacAutoCal3 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal3 = DacAutoCal3 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal3 = DacAutoCal3 - Val(CalStepFinal.Text)
                        myString = DacAutoCal3.ToString()
                        DacSpan2.Text = "<DacSpan2," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan2.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan2.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts3.Text = txtr1aBIG.Text
                        volts3.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 18 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If

                ' Set 4Vdc - Final Tweak
                If (DACcalparam = 19 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    KeyVoltage.Text = "4.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    TweakOneshot = False
                End If
                If (DACcalparam = 19 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 4Vdc"
                    KeyVoltage.Text = "4.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        ' makes steps larger if far away from final value
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal4 = DacAutoCal4 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal4 = DacAutoCal4 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal4 = DacAutoCal4 + Val(CalStepFinal.Text)
                        myString = DacAutoCal4.ToString()
                        DacSpan3.Text = "<DacSpan3," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan3.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan3.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal4 = DacAutoCal4 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal4 = DacAutoCal4 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal4 = DacAutoCal4 - Val(CalStepFinal.Text)
                        myString = DacAutoCal4.ToString()
                        DacSpan3.Text = "<DacSpan3," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan3.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan3.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts4.Text = txtr1aBIG.Text
                        volts4.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 19 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If

                ' Set 5Vdc - Final Tweak
                If (DACcalparam = 20 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    KeyVoltage.Text = "5.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    TweakOneshot = False
                End If
                If (DACcalparam = 20 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 5Vdc"
                    KeyVoltage.Text = "5.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        ' makes steps larger if far away from final value
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal5 = DacAutoCal5 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal5 = DacAutoCal5 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal5 = DacAutoCal5 + Val(CalStepFinal.Text)
                        myString = DacAutoCal5.ToString()
                        DacSpan4.Text = "<DacSpan4," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan4.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan4.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal5 = DacAutoCal5 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal5 = DacAutoCal5 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal5 = DacAutoCal5 - Val(CalStepFinal.Text)
                        myString = DacAutoCal5.ToString()
                        DacSpan4.Text = "<DacSpan4," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan4.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan4.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts5.Text = txtr1aBIG.Text
                        volts5.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 20 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If

                ' Set 6Vdc - Final Tweak
                If (DACcalparam = 21 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    KeyVoltage.Text = "6.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    TweakOneshot = False
                End If
                If (DACcalparam = 21 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 6Vdc"
                    KeyVoltage.Text = "6.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        ' makes steps larger if far away from final value
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal6 = DacAutoCal6 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal6 = DacAutoCal6 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal6 = DacAutoCal6 + Val(CalStepFinal.Text)
                        myString = DacAutoCal6.ToString()
                        DacSpan5.Text = "<DacSpan5," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan5.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan5.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal6 = DacAutoCal6 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal6 = DacAutoCal6 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal6 = DacAutoCal6 - Val(CalStepFinal.Text)
                        myString = DacAutoCal6.ToString()
                        DacSpan5.Text = "<DacSpan5," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan5.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan5.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts6.Text = txtr1aBIG.Text
                        volts6.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 21 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If

                ' Set 7Vdc - Final Tweak
                If (DACcalparam = 22 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    KeyVoltage.Text = "7.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    TweakOneshot = False
                End If
                If (DACcalparam = 22 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 7Vdc"
                    KeyVoltage.Text = "7.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        ' makes steps larger if far away from final value
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal7 = DacAutoCal7 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal7 = DacAutoCal7 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal7 = DacAutoCal7 + Val(CalStepFinal.Text)
                        myString = DacAutoCal7.ToString()
                        DacSpan6.Text = "<DacSpan6," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan6.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan6.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal7 = DacAutoCal7 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal7 = DacAutoCal7 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal7 = DacAutoCal7 - Val(CalStepFinal.Text)
                        myString = DacAutoCal7.ToString()
                        DacSpan6.Text = "<DacSpan6," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan6.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan6.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts7.Text = txtr1aBIG.Text
                        volts7.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 22 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If

                ' Set 8Vdc - Final Tweak
                If (DACcalparam = 23 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    KeyVoltage.Text = "8.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    TweakOneshot = False
                End If
                If (DACcalparam = 23 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 8Vdc"
                    KeyVoltage.Text = "8.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        ' makes steps larger if far away from final value
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal8 = DacAutoCal8 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal8 = DacAutoCal8 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal8 = DacAutoCal8 + Val(CalStepFinal.Text)
                        myString = DacAutoCal8.ToString()
                        DacSpan7.Text = "<DacSpan7," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan7.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan7.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal8 = DacAutoCal8 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal8 = DacAutoCal8 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal8 = DacAutoCal8 - Val(CalStepFinal.Text)
                        myString = DacAutoCal8.ToString()
                        DacSpan7.Text = "<DacSpan7," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan7.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan7.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts8.Text = txtr1aBIG.Text
                        volts8.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 23 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If

                ' Set 9Vdc - Final Tweak
                If (DACcalparam = 24 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    KeyVoltage.Text = "9.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    TweakOneshot = False
                End If
                If (DACcalparam = 24 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 9Vdc"
                    KeyVoltage.Text = "9.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        ' makes steps larger if far away from final value
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal9 = DacAutoCal9 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal9 = DacAutoCal9 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal9 = DacAutoCal9 + Val(CalStepFinal.Text)
                        myString = DacAutoCal9.ToString()
                        DacSpan8.Text = "<DacSpan8," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan8.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan8.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal9 = DacAutoCal9 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal9 = DacAutoCal9 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal9 = DacAutoCal9 - Val(CalStepFinal.Text)
                        myString = DacAutoCal9.ToString()
                        DacSpan8.Text = "<DacSpan8," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan8.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan8.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts9.Text = txtr1aBIG.Text
                        volts9.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 24 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If

                ' Set 10Vdc - Final Tweak
                If (DACcalparam = 25 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                    KeyVoltage.Text = "10.00000"
                    Call TxLEDon()
                    Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                    System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                    TweakOneshot = False
                End If
                If (DACcalparam = 25 And Tweak = False) Then
                    LabelCalCount.Text = "Final calibrating 10Vdc"
                    KeyVoltage.Text = "10.00000"
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    ' increase req'd
                    If (TweakVar = 1) Then
                        ' makes steps larger if far away from final value
                        If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                            DacAutoCal10 = DacAutoCal10 + (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal10 = DacAutoCal10 + Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal10 = DacAutoCal10 + Val(CalStepFinal.Text)
                        myString = DacAutoCal10.ToString()
                        DacSpan9.Text = "<DacSpan9," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan9.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan9.Text = myString
                        Tweak = True
                    End If

                    ' decrease req'd
                    If (TweakVar = 2) Then
                        ' makes steps larger if far away from final value
                        If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                            DacAutoCal10 = DacAutoCal10 - (Val(CalStepFinal.Text) * 2)
                        Else
                            DacAutoCal10 = DacAutoCal10 - Val(CalStepFinal.Text)
                        End If
                        'DacAutoCal10 = DacAutoCal10 - Val(CalStepFinal.Text)
                        myString = DacAutoCal10.ToString()
                        DacSpan9.Text = "<DacSpan9," + myString + ">"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine(DacSpan9.Text)    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        DacSpan9.Text = myString
                        Tweak = True
                    End If

                    'Tweak = True

                    ' nothing req'd - finished
                    If (TweakVar = 0) Then
                        'volts10.Text = txtr1aBIG.Text
                        volts10.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                        DACcalparam = DACcalparam + 1
                        Tweak = True   ' reset back to start condition for next voltage
                        TweakVar = 999
                        TweakOneshot = True
                    End If
                End If
                If (DACcalparam = 25 And Tweak = True) Then
                    VoltageSet = Val(KeyVoltage.Text)
                    VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                    If (Val(txtr1aBIG.Text) < VoltageSet) Then
                        TweakVar = 1
                    End If
                    If (Val(txtr1aBIG.Text) > VoltageSet) Then
                        TweakVar = 2
                    End If
                    If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                        TweakVar = 0
                    End If
                    Tweak = False
                End If


                If WryTech.Checked = True Then

                    ' Set 10.22222Vdc - Final Tweak
                    If (DACcalparam = 26 And Tweak = True And TweakOneshot = True) Then     ' oneshot setting of PDVS2 output to start off
                        KeyVoltage.Text = "10.22222"
                        Call TxLEDon()
                        Me.SerialPort1.WriteLine("<KeyVoltage,0," + KeyVoltage.Text + ">")    ' Write data to the send buffer
                        System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                        TweakOneshot = False
                    End If
                    If (DACcalparam = 26 And Tweak = False) Then
                        LabelCalCount.Text = "Final calibrating 10.22222Vdc"
                        KeyVoltage.Text = "10.22222"
                        VoltageSet = Val(KeyVoltage.Text)
                        VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                        ' increase req'd
                        If (TweakVar = 1) Then
                            ' makes steps larger if far away from final value
                            If (VoltageSet - Val(txtr1aBIG.Text) > 0.00005) Then
                                DacAutoCal11 = DacAutoCal11 + (Val(CalStepFinal.Text) * 2)
                            Else
                                DacAutoCal11 = DacAutoCal11 + Val(CalStepFinal.Text)
                            End If
                            'DacAutoCal11 = DacAutoCal11 + Val(CalStepFinal.Text)
                            myString = DacAutoCal11.ToString()
                            DacSpan10.Text = "<DacSpan10," + myString + ">"
                            Call TxLEDon()
                            Me.SerialPort1.WriteLine(DacSpan10.Text)    ' Write data to the send buffer
                            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                            DacSpan10.Text = myString
                            Tweak = True
                        End If

                        ' decrease req'd
                        If (TweakVar = 2) Then
                            ' makes steps larger if far away from final value
                            If (Val(txtr1aBIG.Text) - VoltageSet > 0.00005) Then
                                DacAutoCal11 = DacAutoCal11 - (Val(CalStepFinal.Text) * 2)
                            Else
                                DacAutoCal11 = DacAutoCal11 - Val(CalStepFinal.Text)
                            End If
                            'DacAutoCal11 = DacAutoCal11 - Val(CalStepFinal.Text)
                            myString = DacAutoCal11.ToString()
                            DacSpan10.Text = "<DacSpan10," + myString + ">"
                            Call TxLEDon()
                            Me.SerialPort1.WriteLine(DacSpan10.Text)    ' Write data to the send buffer
                            System.Threading.Thread.Sleep(Val(PDVS2delay.Text))  ' give PDVS2mini time to finish
                            DacSpan10.Text = myString
                            Tweak = True
                        End If

                        'Tweak = True

                        ' nothing req'd - finished
                        If (TweakVar = 0) Then
                            'volts10.Text = txtr1aBIG.Text
                            volts11.Text = Format(Val(txtr1aBIG.Text), "#00.0000000")
                            DACcalparam = DACcalparam + 1
                            Tweak = True   ' reset back to start condition for next voltage
                            TweakVar = 999
                            TweakOneshot = True
                        End If
                    End If
                    If (DACcalparam = 26 And Tweak = True) Then
                        VoltageSet = Val(KeyVoltage.Text)
                        VoltageCalAccuracy = Val(CalAccuracyFinal.Text)

                        If (Val(txtr1aBIG.Text) < VoltageSet) Then
                            TweakVar = 1
                        End If
                        If (Val(txtr1aBIG.Text) > VoltageSet) Then
                            TweakVar = 2
                        End If
                        If ((VoltageSet - Val(txtr1aBIG.Text) < VoltageCalAccuracy) And (VoltageSet - Val(txtr1aBIG.Text) > -VoltageCalAccuracy)) Then
                            TweakVar = 0
                        End If
                        Tweak = False
                    End If
                End If

                ' If Wrytech PDVS2mini is not used then need to increment DACcalparam counter
                If DACcalparam = 26 And WryTech.Checked = False Then
                    DACcalparam = 27    ' jump past 10.22222
                End If


                ' Finalize
                If (DACcalparam = 27) Then
                    DACcalparam = 0
                    LabelCalCount.Text = "Done"
                    KeyVoltage.Text = ""
                    Tweak = True
                    TweakVar = 999
                    TweakOneshot = True
                    Me.Timer7.Stop()    ' stop the PDVS2 sub running
                    Me.Timer2.Stop()    ' stop Dev1 GPIB running
                    ButtonDev1Run.Text = "Run"

                    ' set GPIB to NPLC 50
                    'Me.Timer7.Stop()    ' stop the PDVS2 sub running
                    'dev1.SendAsync("NPLC 50", True)
                    'NPLC.Text = "50"
                    'Me.Timer2.Stop()    ' stop Dev1 GPIB running
                    'ButtonDev1Run.Text = "Run"

                    ' set sample rate of GPIB to 3 secs (covers the 50 NPLC)
                    'Dev1SampleRate.Text = "3.0"
                    'Dev1TimerDuration = Dev1SampleRate.Text
                    'Me.Timer2.Interval = Dev1TimerDuration * 1000
                    'Me.Timer2.Start()    ' restart Dev1 GPIB running
                    'ButtonDev1Run.Text = "Stop"

                    'autocal_BTN.Enabled = False ' disable Auto-Cal button
                    getcal_BTN.Enabled = True ' enable button
                    precal_BTN.Enabled = True ' enable Full Auto-Cal button
                    CalOnExisting.Enabled = True ' enable Auto-Cal on existing button
                    ButtonSetPrecalvars.Enabled = True ' enable Pre-Cal button
                    ButtonSetXYcalvars.Enabled = True ' enable button
                    SavePDVS2Eprom.Enabled = True ' enable button
                    ButtonSetRetrievedVars.Enabled = True
                    ButtonAutoSET.Enabled = True
                    ButtonAutomV.Enabled = True

                    ' Play sound to indicate finished
                    My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Asterisk)
                    'System.Threading.Thread.Sleep(100) ' delay 100mS
                    'My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Asterisk)

                End If
            End If
        End If
    End Sub

    Private Sub nplc_BTN_Click(sender As Object, e As EventArgs) Handles nplc_BTN.Click

        dev1.readtimeout = 15000

        'If dev1 IsNot Nothing Then      ' Device 1 is open

        ' set GPIB to NPLC 100
        Me.Timer7.Stop()    ' stop the PDVS2 sub running
        dev1.SendAsync("NPLC 100", True)
        NPLC.Text = "100"
        'Me.Timer2.Stop()    ' stop Dev1 GPIB running
        'ButtonDev1Run.Text = "Run"

        ' set sample rate of GPIB to 10 secs (covers the 100 NPLC)
        Dev1SampleRate.Text = "10.0"
        Dev1TimerDuration = Dev1SampleRate.Text
        'Me.Timer2.Interval = Dev1TimerDuration * 1000
        'Me.Timer2.Start()    ' restart Dev1 GPIB running
        'ButtonDev1Run.Text = "Stop"

        LabelCalCount.Text = "NPLC set to 100"

        'End If

    End Sub

    Private Sub nplc2_BTN_Click(sender As Object, e As EventArgs) Handles nplc2_BTN.Click

        dev1.readtimeout = 5000

        'If dev1 IsNot Nothing Then      ' Device 1 is open

        ' set GPIB to NPLC 25
        Me.Timer7.Stop()    ' stop the PDVS2 sub running
        dev1.SendAsync("NPLC 25", True)
        NPLC.Text = "25"
        'Me.Timer2.Stop()    ' stop Dev1 GPIB running
        'ButtonDev1Run.Text = "Run"

        ' set sample rate of GPIB to 5 secs (covers the 25 NPLC)
        Dev1SampleRate.Text = "5.0"
        Dev1TimerDuration = Dev1SampleRate.Text
        'Me.Timer2.Interval = Dev1TimerDuration * 1000
        'Me.Timer2.Start()    ' restart Dev1 GPIB running
        'ButtonDev1Run.Text = "Stop"

        LabelCalCount.Text = "NPLC set to 25"

        'End If

    End Sub

    Private Sub nplc3_BTN_Click(sender As Object, e As EventArgs) Handles nplc3_BTN.Click

        dev1.readtimeout = 5000

        'If dev1 IsNot Nothing Then      ' Device 1 is open

        ' set GPIB to NPLC 50
        Me.Timer7.Stop()    ' stop the PDVS2 sub running
        dev1.SendAsync("NPLC 50", True)
        NPLC.Text = "50"
        'Me.Timer2.Stop()    ' stop Dev1 GPIB running
        'ButtonDev1Run.Text = "Run"

        ' set sample rate of GPIB to 5 secs (covers the 25 NPLC)
        Dev1SampleRate.Text = "5.0"
        Dev1TimerDuration = Dev1SampleRate.Text
        'Me.Timer2.Interval = Dev1TimerDuration * 1000
        'Me.Timer2.Start()    ' restart Dev1 GPIB running
        'ButtonDev1Run.Text = "Stop"

        LabelCalCount.Text = "NPLC set to 50"

        'End If

    End Sub

    Private Sub nplc4_BTN_Click(sender As Object, e As EventArgs) Handles nplc4_BTN.Click

        dev1.readtimeout = 25000

        'If dev1 IsNot Nothing Then      ' Device 1 is open

        ' set GPIB to NPLC 200
        Me.Timer7.Stop()    ' stop the PDVS2 sub running
        dev1.SendAsync("NPLC 200", True)
        NPLC.Text = "200"
        'Me.Timer2.Stop()    ' stop Dev1 GPIB running
        'ButtonDev1Run.Text = "Run"

        ' set sample rate of GPIB to 5 secs (covers the 200 NPLC)
        Dev1SampleRate.Text = "20.0"
        Dev1TimerDuration = Dev1SampleRate.Text
        'Me.Timer2.Interval = Dev1TimerDuration * 1000
        'Me.Timer2.Start()    ' restart Dev1 GPIB running
        'ButtonDev1Run.Text = "Stop"

        LabelCalCount.Text = "NPLC set to 200"

        'End If

    End Sub

    Private Sub getcal_BTN_Click(sender As Object, e As EventArgs) Handles getcal_BTN.Click

        If Me.SerialPort1.IsOpen = True Then

            LabelCalCount.Text = "Cal data receiving..."

            Me.Refresh()

            'LabelDeltaV.Text = "####"

            getcal_BTN.Enabled = False ' disable button

            DacZero0.Text = LabeldacZero0Cal.Text       ' this is not on a form....it is the data from the actual PDVS2mini

            DacSpan0.Text = LabeldacSpan0Cal.Text
            DacSpan1.Text = LabeldacSpan1Cal.Text
            DacSpan2.Text = LabeldacSpan2Cal.Text
            DacSpan3.Text = LabeldacSpan3Cal.Text
            DacSpan4.Text = LabeldacSpan4Cal.Text
            DacSpan5.Text = LabeldacSpan5Cal.Text
            DacSpan6.Text = LabeldacSpan6Cal.Text
            DacSpan7.Text = LabeldacSpan7Cal.Text
            DacSpan8.Text = LabeldacSpan8Cal.Text
            DacSpan9.Text = LabeldacSpan9Cal.Text

            If WryTech.Checked = True Then
                DacSpan10.Text = LabeldacSpan10Cal.Text
            End If

            OutVdc.Text = LabelOutputVFeedMult.Text
            BattVdc.Text = LabelBatteryVFeedMult.Text
            ChargeI.Text = LabelBatteryMonICMult.Text

            DacAutoCal0 = LabeldacZero0Cal.Text

            DacAutoCal1 = LabeldacSpan0Cal.Text
            DacAutoCal2 = LabeldacSpan1Cal.Text
            DacAutoCal3 = LabeldacSpan2Cal.Text
            DacAutoCal4 = LabeldacSpan3Cal.Text
            DacAutoCal5 = LabeldacSpan4Cal.Text
            DacAutoCal6 = LabeldacSpan5Cal.Text
            DacAutoCal7 = LabeldacSpan6Cal.Text
            DacAutoCal8 = LabeldacSpan7Cal.Text
            DacAutoCal9 = LabeldacSpan8Cal.Text
            DacAutoCal10 = LabeldacSpan9Cal.Text

            If WryTech.Checked = True Then
                DacAutoCal11 = LabeldacSpan10Cal.Text
            End If

            ' Set DeltaV to #######
            LabeldacZero0Delta.Text = ""
            LabeldacSpan0Delta.Text = ""
            LabeldacSpan1Delta.Text = ""
            LabeldacSpan2Delta.Text = ""
            LabeldacSpan3Delta.Text = ""
            LabeldacSpan4Delta.Text = ""
            LabeldacSpan5Delta.Text = ""
            LabeldacSpan6Delta.Text = ""
            LabeldacSpan7Delta.Text = ""
            LabeldacSpan8Delta.Text = ""
            LabeldacSpan9Delta.Text = ""

            If WryTech.Checked = True Then
                LabeldacSpan10Delta.Text = ""
            End If

            volts0.Text = ""
            volts1.Text = ""
            volts2.Text = ""
            volts3.Text = ""
            volts4.Text = ""
            volts5.Text = ""
            volts6.Text = ""
            volts7.Text = ""
            volts8.Text = ""
            volts9.Text = ""
            volts10.Text = ""

            If WryTech.Checked = True Then
                volts11.Text = ""
            End If

            volts000010.Text = ""
            volts000100.Text = ""
            volts001000.Text = ""
            volts010000.Text = ""
            volts020000.Text = ""
            volts030000.Text = ""
            volts050000.Text = ""

            System.Threading.Thread.Sleep(1000) ' delay for 1sec

            LabelCalCount.Text = "Cal data received complete"

            getcal_BTN.Enabled = True ' enable button

            ' Play sound to indicate finished
            My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Asterisk)

        Else

            LabelCalCount.Text = "Comms not ready"


        End If


    End Sub

    ' Abort button

    Private Sub Abort_BTN_Click(sender As Object, e As EventArgs) Handles Abort_BTN.Click

        getcal_BTN.Enabled = True ' enable button
        precal_BTN.Enabled = True ' enable Full Auto-Cal button
        CalOnExisting.Enabled = True ' enable Auto-Cal on existing button
        ButtonSetPrecalvars.Enabled = True ' enable Pre-Cal button
        ButtonSetXYcalvars.Enabled = True ' enable button
        SavePDVS2Eprom.Enabled = True ' enable button
        ButtonSetRetrievedVars.Enabled = True
        ButtonAutoSET.Enabled = True
        ButtonAutomV.Enabled = True

        DACcalparam = 0
        'LabelCalCount.Text = ""
        'DacZero0.Text = ""
        'DacSpan0.Text = ""
        'DacSpan1.Text = ""
        'DacSpan2.Text = ""
        'DacSpan3.Text = ""
        'DacSpan4.Text = ""
        'DacSpan5.Text = ""
        'DacSpan6.Text = ""
        'DacSpan7.Text = ""
        'DacSpan8.Text = ""
        'DacSpan9.Text = ""
        'volts0.Text = ""
        'volts1.Text = ""
        'volts2.Text = ""
        'volts3.Text = ""
        'volts4.Text = ""
        'volts5.Text = ""
        'volts6.Text = ""
        'volts7.Text = ""
        'volts8.Text = ""
        'volts9.Text = ""
        'volts10.Text = ""
        'OutVdc.Text = ""
        'BattVdc.Text = ""
        'ChargeI.Text = ""
        Tweak = True
        TweakVar = 999
        TweakOneshot = True

        LabelCalCount.Text = "Aborted"

        Me.Timer2.Stop()    ' stop GPIB activity
        ButtonDev1Run.Text = "Run"

        Me.Timer7.Stop()
        Me.Refresh()

    End Sub

    Private Sub exportBTN_Click(sender As Object, e As EventArgs) Handles exportBTN.Click

        ' I added this to avoid having a warning on compilation related to fact I don't have Word installed on this devPC. Here are the instructions if you DO have Word installed:
        ' Define the symbol in your project: 
        ' In Visual Studio, go to Project > Properties > Compile tab.
        ' In the Conditional Compilation Symbols textbox, add WORD_INSTALLED (if Word Is installed on the target machine).
        ' Add the conditional defines in the code, per below and the header define.


#If WORD_INSTALLED Then
    ' MS Word code block

        ' MS Word
        If RadioMSWord.Checked = True Then

            Dim oWord As Microsoft.Office.Interop.Word.Application
            Dim oDoc As Microsoft.Office.Interop.Word.Document
            Dim oTable As Microsoft.Office.Interop.Word.Table
            Dim oPara1 As Microsoft.Office.Interop.Word.Paragraph ', oPara2 As Microsoft.Office.Interop.Word.Paragraph

            LabelCalCount.Text = "Generating cal. cert"

            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            oWord.Visible = False   ' this hides Word so it all happens in the background
            oDoc = oWord.Documents.Add

            'Insert a paragraph at the beginning of the document.

            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Font.Size = 14

            If WryTech.Checked = True Then
                oPara1.Range.Text = "PDVS2mini (WryTech) - TEST / CALIBRATION RECORD    " & TextBoxUser.Text
            Else
                oPara1.Range.Text = "PDVS2mini - TEST / CALIBRATION RECORD              " & TextBoxUser.Text
            End If

            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Underline = True
            oPara1.Format.SpaceAfter = 20    '20 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()
            oPara1.Range.Font.Size = 11
            oPara1.Range.Font.Underline = False

            oTable = oDoc.Tables.Add(oDoc.Bookmarks("\endofdoc").Range, 39, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 2

            oTable.Range.Font.Bold = False
            oTable.Rows(1).Range.Font.Bold = True
            oTable.Rows(1).Range.Font.Size = 12

            oTable.Columns(1).Width = oWord.InchesToPoints(2.75)   'Change width of columns 1, 2 & 3
            oTable.Columns(2).Width = oWord.InchesToPoints(2)
            oTable.Columns(3).Width = oWord.InchesToPoints(2)

            ' Column 1
            oTable.Cell(1, 1).Range.Text = "FUNCTION"
            oTable.Cell(1, 2).Range.Text = "COUNTS/SETTING"
            oTable.Cell(1, 3).Range.Text = "VOLTAGE OUTPUT"

            oTable.Cell(2, 1).Range.Text = "0.00000 V"
            oTable.Cell(3, 1).Range.Text = "1.00000 V"
            oTable.Cell(4, 1).Range.Text = "2.00000 V"
            oTable.Cell(5, 1).Range.Text = "3.00000 V"
            oTable.Cell(6, 1).Range.Text = "4.00000 V"
            oTable.Cell(7, 1).Range.Text = "5.00000 V"
            oTable.Cell(8, 1).Range.Text = "6.00000 V"
            oTable.Cell(9, 1).Range.Text = "7.00000 V"
            oTable.Cell(10, 1).Range.Text = "8.00000 V"
            oTable.Cell(11, 1).Range.Text = "9.00000 V"
            oTable.Cell(12, 1).Range.Text = "10.00000 V"

            If WryTech.Checked = True Then
                oTable.Cell(13, 1).Range.Text = "10.22222 V"
            Else
                oTable.Cell(13, 1).Range.Text = ""
            End If

            oTable.Cell(14, 1).Range.Text = ""

            oTable.Cell(15, 1).Range.Text = "0.00010 (0.1 mV)"
            oTable.Cell(16, 1).Range.Text = "0.00100 (1.0 mV)"
            oTable.Cell(17, 1).Range.Text = "0.01000 (10 mV)"
            oTable.Cell(18, 1).Range.Text = "0.10000 (100 mV)"
            oTable.Cell(19, 1).Range.Text = "0.20000 (200 mV)"
            oTable.Cell(20, 1).Range.Text = "0.30000 (300 mV)"
            oTable.Cell(21, 1).Range.Text = "0.50000 (500 mV)"

            oTable.Cell(22, 1).Range.Text = ""

            oTable.Cell(23, 1).Range.Text = "Out.Vdc"
            oTable.Cell(24, 1).Range.Text = "Batt.Vdc"
            oTable.Cell(25, 1).Range.Text = "Charge mA"
            oTable.Cell(26, 1).Range.Text = "Battery Low/Shutdown"
            oTable.Cell(27, 1).Range.Text = "Charge Enable"
            oTable.Cell(28, 1).Range.Text = "Charge Overload mA"
            oTable.Cell(29, 1).Range.Text = "Charge Full mA"
            oTable.Cell(30, 1).Range.Text = "Serial Comms"
            oTable.Cell(31, 1).Range.Text = "Soak Test Duration"
            oTable.Cell(32, 1).Range.Text = "External DC Input"
            oTable.Cell(33, 1).Range.Text = "Battery Charge/Discharge Test"

            oTable.Cell(34, 1).Range.Text = ""

            oTable.Cell(35, 1).Range.Text = "Serial Number"
            oTable.Cell(36, 1).Range.Text = "Ambient Temperature " & TextBoxTempUnits.Text
            oTable.Cell(37, 1).Range.Text = "Date"

            oTable.Cell(38, 1).Range.Text = ""

            oTable.Cell(39, 1).Range.Font.Italic = True
            oTable.Cell(39, 1).Range.Text = "Cal = HP 3458A, S/N=" & TextBox3458Asn.Text
            oTable.Cell(39, 1).Range.Font.Italic = True
            oTable.Cell(39, 1).Range.Text = "WinGPIB App"


            ' Column 2
            oTable.Cell(2, 2).Range.Text = DacZero0.Text
            oTable.Cell(3, 2).Range.Text = DacSpan0.Text
            oTable.Cell(4, 2).Range.Text = DacSpan1.Text
            oTable.Cell(5, 2).Range.Text = DacSpan2.Text
            oTable.Cell(6, 2).Range.Text = DacSpan3.Text
            oTable.Cell(7, 2).Range.Text = DacSpan4.Text
            oTable.Cell(8, 2).Range.Text = DacSpan5.Text
            oTable.Cell(9, 2).Range.Text = DacSpan6.Text
            oTable.Cell(10, 2).Range.Text = DacSpan7.Text
            oTable.Cell(11, 2).Range.Text = DacSpan8.Text
            oTable.Cell(12, 2).Range.Text = DacSpan9.Text

            If WryTech.Checked = True Then
                oTable.Cell(13, 2).Range.Text = DacSpan10.Text
            Else
                oTable.Cell(13, 2).Range.Text = ""
            End If

            oTable.Cell(23, 2).Range.Text = OutVdc.Text
            oTable.Cell(24, 2).Range.Text = BattVdc.Text
            oTable.Cell(25, 2).Range.Text = ChargeI.Text
            oTable.Cell(26, 2).Range.Text = TextBoxLOWSHUT.Text
            oTable.Cell(27, 2).Range.Text = TextBoxCENABLE.Text
            oTable.Cell(28, 2).Range.Text = TextBoxOLMA.Text
            oTable.Cell(29, 2).Range.Text = TextBoxFULLMA.Text
            oTable.Cell(30, 2).Range.Text = TextBoxSERIAL.Text
            oTable.Cell(31, 2).Range.Text = TextBoxSOAK.Text
            oTable.Cell(32, 2).Range.Text = TextBoxDC.Text
            oTable.Cell(33, 2).Range.Text = TextBoxCD.Text
            oTable.Cell(34, 2).Range.Text = ""
            oTable.Cell(35, 2).Range.Text = TextBoxSer.Text
            oTable.Cell(36, 2).Range.Text = TextBoxdegC.Text
            oTable.Cell(37, 2).Range.Text = DateTime.Now.ToString("dd MMM. yyyy")  ' dd/MM/yyyy HH:mm:ss
            oTable.Cell(38, 2).Range.Text = ""


            ' Column 3
            'oTable.Cell(2, 3).Range.Text = "0" & volts0.Text
            oTable.Cell(2, 3).Range.Text = volts0.Text
            oTable.Cell(3, 3).Range.Text = volts1.Text
            oTable.Cell(4, 3).Range.Text = volts2.Text
            oTable.Cell(5, 3).Range.Text = volts3.Text
            oTable.Cell(6, 3).Range.Text = volts4.Text
            oTable.Cell(7, 3).Range.Text = volts5.Text
            oTable.Cell(8, 3).Range.Text = volts6.Text
            oTable.Cell(9, 3).Range.Text = volts7.Text
            oTable.Cell(10, 3).Range.Text = volts8.Text
            oTable.Cell(11, 3).Range.Text = volts9.Text
            oTable.Cell(12, 3).Range.Text = volts10.Text

            If WryTech.Checked = True Then
                oTable.Cell(13, 3).Range.Text = volts11.Text
            Else
                oTable.Cell(13, 3).Range.Text = ""
            End If

            oTable.Cell(14, 3).Range.Text = ""

            oTable.Cell(15, 3).Range.Text = volts000010.Text
            oTable.Cell(16, 3).Range.Text = volts000100.Text
            oTable.Cell(17, 3).Range.Text = volts001000.Text
            oTable.Cell(18, 3).Range.Text = volts010000.Text
            oTable.Cell(19, 3).Range.Text = volts020000.Text
            oTable.Cell(20, 3).Range.Text = volts030000.Text
            oTable.Cell(21, 3).Range.Text = volts050000.Text

            oTable.Cell(21, 3).Range.Text = ""

            ' Save MS Word doc with serial number as name, by default to the Documents folder
            If TextBoxSer.Text = "" Then
                LabelCalCount.Text = "ERROR - Serial Number is req'd"
                'oDoc.Close()
            Else
                'oDoc.SaveAs(TextBoxSer.Text)
                oDoc.SaveAs(CSVfilepath.Text & "\" & "Certificates" & "\" & TextBoxSer.Text & ".docx")
                oDoc.Close()
                LabelCalCount.Text = "Cal. cert saved"    ' - " & TextBoxSer.Text & ".docx"
            End If

            Dialog2.Warning1 = "Word Doc file has been created:"
            Dialog2.Warning2 = CSVfilepath.Text & "\" & "Certificates" & "\" & TextBoxSer.Text & ".docx"
            Dialog2.Warning3 = ""
            'Dialog2.Show() ' this method positions anywhere!
            Dialog2.ShowDialog(Me)  ' this method positions centre of parent form, and requires to hit OK to return back to parent

        End If

#End If


        ' PDF
        If RadioPDF.Checked = True Then

            ' This uses PDFSharp, see top of this page for install instructions.

            ' Create a new PDF document
            Dim document As New PdfDocument()

            Dim page As PdfPage = document.AddPage()
            page.Width = XUnit.FromMillimeter(210)
            page.Height = XUnit.FromMillimeter(297)

            ' Create a graphics object for drawing on the page
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)

            ' Define fonts
            Dim fontHeader As New XFont("Arial", 14, XFontStyle.Bold)
            Dim fontBold As New XFont("Arial", 11, XFontStyle.Bold)
            Dim fontNormal As New XFont("Arial", 11)

            ' Draw the header
            If WryTech.Checked = True Then
                gfx.DrawString("PDVS2mini (WryTech) - TEST / CALIBRATION RECORD    " & TextBoxUser.Text, fontHeader, XBrushes.Black, New XRect(40, 40, 500, 20), XStringFormats.TopLeft)
            Else
                gfx.DrawString("PDVS2mini - TEST / CALIBRATION RECORD              " & TextBoxUser.Text, fontHeader, XBrushes.Black, New XRect(40, 40, 500, 20), XStringFormats.TopLeft)
            End If

            ' Create the table
            Dim tableRect As New XRect(50, 70, 600, 780)
            Dim columnWidth As Double = 190
            Dim rowHeight As Double = 25

            ' Draw column headers
            gfx.DrawString("FUNCTION", fontBold, XBrushes.Black, New XRect(50, 90, columnWidth, rowHeight), XStringFormats.TopLeft)
            gfx.DrawString("COUNTS/SETTING", fontBold, XBrushes.Black, New XRect(250, 90, columnWidth, rowHeight), XStringFormats.TopLeft)
            gfx.DrawString("VOLTAGE OUTPUT", fontBold, XBrushes.Black, New XRect(450, 90, columnWidth, rowHeight), XStringFormats.TopLeft)


            rowHeight = 17

            ' Draw the table content
            Dim startY As Double = 120


            ' Column 1
            gfx.DrawString("0.00000 V", fontNormal, XBrushes.Black, New XRect(50, 120, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("1.00000 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("2.00000 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("3.00000 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("4.00000 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("5.00000 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("6.00000 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("7.00000 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("8.00000 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("9.00000 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("10.00000 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            If WryTech.Checked = True Then
                gfx.DrawString("10.22222 V", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
                startY += rowHeight
            Else
                gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
                startY += rowHeight
            End If

            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString("0.00010 (0.1mV)", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("0.00100 (1.0mV)", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("0.01000 (10mV)", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("0.10000 (100mV)", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("0.20000 (200mV)", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("0.30000 (300mV)", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("0.40000 (400mV)", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Out.Vdc", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Batt.Vdc", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Charge mA", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Battery Low/Shutdown", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Charge Enable", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Charge Overload mA", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Charge Full mA", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Serial Comms", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Soak Test Duration", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("External DC Input", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Battery Charge/Discharge Test", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString("Serial Number", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Ambient Temperature " & TextBoxTempUnits.Text, fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("Date", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString("Cal = HP3458A, S/N=" & TextBox3458Asn.Text, fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("WinGPIB", fontNormal, XBrushes.Black, New XRect(50, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight


            startY = 120


            ' Column 2
            gfx.DrawString(DacZero0.Text, fontNormal, XBrushes.Black, New XRect(250, 120, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DacSpan0.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DacSpan1.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DacSpan2.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DacSpan3.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DacSpan4.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DacSpan5.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DacSpan6.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DacSpan7.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DacSpan8.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DacSpan9.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            If WryTech.Checked = True Then
                gfx.DrawString(DacSpan10.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
                startY += rowHeight
            Else
                gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
                startY += rowHeight
            End If

            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            'gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            'startY += rowHeight

            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString(OutVdc.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(BattVdc.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(ChargeI.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(TextBoxLOWSHUT.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(TextBoxCENABLE.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(TextBoxOLMA.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(TextBoxFULLMA.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(TextBoxSERIAL.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(TextBoxSOAK.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(TextBoxDC.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(TextBoxCD.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString(TextBoxSer.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(TextBoxdegC.Text, fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(DateTime.Now.ToString("dd MMM. yyyy"), fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)    ' dd/MM/yyyy HH:mm:ss
            startY += rowHeight

            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(250, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight


            startY = 120


            ' Column 3
            gfx.DrawString(volts0.Text, fontNormal, XBrushes.Black, New XRect(450, 120, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts1.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts2.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts3.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts4.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts5.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts6.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts7.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts8.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts9.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts10.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            If WryTech.Checked = True Then
                gfx.DrawString(volts11.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
                startY += rowHeight
            Else
                gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
                startY += rowHeight
            End If

            gfx.DrawString("", fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight

            gfx.DrawString(volts000010.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts000100.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts001000.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts010000.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts020000.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts030000.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight
            gfx.DrawString(volts050000.Text, fontNormal, XBrushes.Black, New XRect(450, startY, columnWidth, rowHeight), XStringFormats.TopLeft)
            startY += rowHeight



            ' Save the PDF to a file
            Dim outputPath As String = CSVfilepath.Text & "\" & "Certificates" & "\" & TextBoxSer.Text & ".pdf" ' Replace with your desired output file path
            document.Save(outputPath)

            LabelCalCount.Text = "Cal. cert saved"    ' - " & TextBoxSer.Text & ".pdf"

            Dialog2.Warning1 = "PDF file has been created:"
            Dialog2.Warning2 = CSVfilepath.Text & "\" & "Certificates" & "\" & TextBoxSer.Text & ".pdf"
            Dialog2.Warning3 = ""
            'Dialog2.Show() ' this method positions anywhere!
            Dialog2.ShowDialog(Me)  ' this method positions centre of parent form, and requires to hit OK to return back to parent

        End If


    End Sub


    Private Sub comPort_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comPort_ComboBox.SelectedIndexChanged
        'If (comPort_ComboBox.SelectedItem <> "") Then
        comPORT = comPort_ComboBox.SelectedItem
        'Label230.Text = comPORT
        'End If
    End Sub


    Private Sub connect_BTN_Click(sender As Object, e As EventArgs) Handles connect_BTN.Click
        If (connect_BTN.Text = "Connect") Then

            ' disable Wrytech checkbox from being changed when comms running
            'WryTech.Enabled = False

            ' Check if current com port selection contains something, then set that com port
            If comPort_ComboBox.SelectedItem <> "" Then
                comPORT = comPort_ComboBox.SelectedItem
            Else
                Exit Sub
            End If

            ' Check if selected item is blank in pull down
            If comPort_ComboBox.SelectedItem = "" Then
                MessageBox.Show($"No port is selected, please select one.{vbCrLf}Perhaps the previously saved port is not available.")
                Exit Sub
            End If

            'Label230.Text = comPORT

            If (comPORT <> "") Then

                Try
                    Me.SerialPort1.Close()
                    Me.SerialPort1.PortName = comPORT
                    Me.SerialPort1.BaudRate = 250000
                    Me.SerialPort1.DataBits = 8
                    Me.SerialPort1.Parity = Parity.None
                    Me.SerialPort1.StopBits = StopBits.One
                    Me.SerialPort1.Handshake = Handshake.None
                    Me.SerialPort1.Encoding = System.Text.Encoding.Default
                    Me.SerialPort1.ReadTimeout = 1000
                    Me.SerialPort1.WriteTimeout = 1000
                    Me.SerialPort1.ReadBufferSize = 4096
                    Me.SerialPort1.WriteBufferSize = 4096

                    Me.SerialPort1.Open()

                Catch ex As UnauthorizedAccessException
                    MessageBox.Show("Access to the COM port is denied.")
                    Me.SerialPort1.Close()
                    connect_BTN.Text = "Connect"
                    CommsStatus = False
                    Timer6.Enabled = False
                    Me.comPort_ComboBox.Enabled = True
                    Exit Sub
                Catch ex As IOException
                    MessageBox.Show("The COM port is in an invalid state.")
                    Me.SerialPort1.Close()
                    connect_BTN.Text = "Connect"
                    CommsStatus = False
                    Timer6.Enabled = False
                    Me.comPort_ComboBox.Enabled = True
                    Exit Sub
                Catch ex As InvalidOperationException
                    MessageBox.Show("The specified COM port is already open.")
                    Me.SerialPort1.Close()
                    connect_BTN.Text = "Connect"
                    CommsStatus = False
                    Timer6.Enabled = False
                    Me.comPort_ComboBox.Enabled = True
                    Exit Sub
                Catch ex As TimeoutException
                    MessageBox.Show("The operation has timed out.")
                    Me.SerialPort1.Close()
                    connect_BTN.Text = "Connect"
                    CommsStatus = False
                    Timer6.Enabled = False
                    Me.comPort_ComboBox.Enabled = True
                    Exit Sub
                Catch ex As Exception
                    MessageBox.Show("An error occurred: " & ex.Message)
                    Me.SerialPort1.Close()
                    connect_BTN.Text = "Connect"
                    CommsStatus = False
                    Timer6.Enabled = False
                    Me.comPort_ComboBox.Enabled = True
                    Exit Sub
                End Try

                connect_BTN.Text = "Disconnect"
                CommsStatus = True
                Timer6.Enabled = True
                'Timer_LBL.Text = "Timer: ON"

                Me.comPort_ComboBox.Enabled = False

            Else
                MsgBox("Select a COM port first")
            End If
        Else

            Me.SerialPort1.Close()
            connect_BTN.Text = "Connect"
            CommsStatus = False
            Timer6.Enabled = False
            Me.comPort_ComboBox.Enabled = True

            ' re-enable Wrytech checkbox
            'WryTech.Enabled = True

        End If


    End Sub


    Function ReceiveSerialData() As String
        Dim Incoming As String
        Try
            Incoming = Me.SerialPort1.ReadExisting()
            If Incoming Is Nothing Then
                Return "nothing" & vbCrLf
            Else
                Return Incoming
            End If
        Catch ex As TimeoutException
            Return "Error: Serial Port read timed out."
        End Try

    End Function


    Private Sub Line_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles GroupBox2.Paint

        ' Specify the coordinates for the vertical line
        Dim x As Integer = 882 ' Adjust this according to where you want the line
        Dim startY As Integer = 6 ' Start Y-coordinate
        Dim endY As Integer = 485 ' End Y-coordinate (bottom of the panel)

        ' Create a Pen for drawing the line
        Using pen As New System.Drawing.Pen(System.Drawing.Color.LightGray)
            ' Draw the line
            e.Graphics.DrawLine(pen, x, startY, x, endY)
        End Using
    End Sub


    Private Sub Timer13_Tick(sender As Object, e As EventArgs)

        ' Stop Timer13 and turn off the indicator after 0.1 seconds
        Timer13.Stop()
        RemoveHandler Timer13.Tick, AddressOf Timer13_Tick
        OnOffLed3.State = OnOffLed.LedState.OffSmallBlack
        OnOffLed3.Invalidate()
        OnOffLed3.Refresh()

    End Sub

    Private Sub Timer14_Tick(sender As Object, e As EventArgs)

        ' Stop Timer14 and turn off the indicator after 0.1 seconds
        Timer14.Stop()
        RemoveHandler Timer14.Tick, AddressOf Timer14_Tick
        OnOffLed4.State = OnOffLed.LedState.OffSmallBlack
        OnOffLed4.Invalidate()
        OnOffLed4.Refresh()

    End Sub


    Private Sub TxLEDon()

        ' Turn LED on
        OnOffLed3.State = OnOffLed.LedState.OnSmall
        OnOffLed3.Invalidate()
        OnOffLed3.Refresh()
        Timer13.Interval = 40 ' Set the interval
        AddHandler Timer13.Tick, AddressOf Timer13_Tick
        Timer13.Start()

    End Sub


    Private Sub RxLEDon()

        ' Turn LED on
        OnOffLed4.State = OnOffLed.LedState.On
        OnOffLed4.Invalidate()
        OnOffLed4.Refresh()
        Timer14.Interval = 40 ' Set the interval
        AddHandler Timer14.Tick, AddressOf Timer14_Tick
        Timer14.Start()

    End Sub


    Private Sub DisableAllButtonsInGroupBox2ExceptPDVS2miniSave()
        ' Loop through all controls in GroupBox2
        For Each ctrl As Control In GroupBox2.Controls
            ' Check if the control is a Button and is not PDVS2miniSave
            If TypeOf ctrl Is Button AndAlso ctrl.Name <> "PDVS2miniSave" AndAlso ctrl.Name <> "ButtonLoadDefs" AndAlso ctrl.Name <> "ButtonSaveDefs" AndAlso ctrl.Name <> "getcal_BTN" AndAlso ctrl.Name <> "ButtonSetPrecalvars" AndAlso ctrl.Name <> "ButtonSetRetrievedVars" AndAlso ctrl.Name <> "SavePDVS2Eprom" AndAlso ctrl.Name <> "exportBTN" Then
                ' Disable the button
                ctrl.Enabled = False
            End If
        Next
        PDVS2miniCalAvailable = False
    End Sub


    Private Sub EnableAllButtonsInGroupBox2()
        ' Loop through all controls in GroupBox2
        For Each ctrl As Control In GroupBox2.Controls
            ' Check if the control is a Button
            If TypeOf ctrl Is Button Then
                ' Enable the button
                ctrl.Enabled = True
            End If
        Next
        PDVS2miniCalAvailable = True
    End Sub


    Private Sub ButtonRefreshPorts1_Click(sender As Object, e As EventArgs) Handles ButtonRefreshPorts1.Click

        'SerialPort1.Close()

        ' Set the selected item to Nothing (empty) or an empty string
        Me.comPort_ComboBox.SelectedItem = Nothing

        ' Close and dispose of the current serial port if it's open
        If SerialPort1.IsOpen Then
            Try
                SerialPort1.Close()
                comPort_ComboBox.Enabled = True
            Catch ex As Exception
                ' Handle the exception as needed (log or display an error message)
            End Try
        End If

        SerialPort1.Dispose()

        ' Clear existing items in the ComboBox
        Me.comPort_ComboBox.Items.Clear()

        ' Use WMI to query the list of serial ports
        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'")
        Dim collection As ManagementObjectCollection = searcher.Get()

        ' Add unique ports to the ComboBox
        For Each device As ManagementObject In collection
            Dim name As String = TryCast(device("Name"), String)
            If name IsNot Nothing Then
                Dim portNameStart As Integer = name.IndexOf("(COM", StringComparison.OrdinalIgnoreCase)
                If portNameStart >= 0 Then
                    Dim portName As String = name.Substring(portNameStart)
                    portName = portName.Replace("(", "").Replace(")", "")               ' Remove parentheses
                    If Not Me.comPort_ComboBox.Items.Contains(portName) Then
                        Me.comPort_ComboBox.Items.Add(portName)
                    End If
                End If
            End If
        Next

        ' Set the selected item to Nothing (empty) or an empty string
        Me.comPort_ComboBox.SelectedItem = Nothing

    End Sub


    Private Sub WryTech_CheckedChanged(sender As Object, e As EventArgs) Handles WryTech.CheckedChanged

        If WryTech.Checked = True Then
            Label229.Enabled = True
            LabelTemperature3.Enabled = True
            volts11.Enabled = True

            If PDVS2miniCalAvailable = True Then
                ButtonDacSpan10down.Enabled = True
                ButtonDacSpan10.Enabled = True
                ButtonDacSpan10Up.Enabled = True
            End If

            LabeldacSpan10Cal.Enabled = True
            Label150.Enabled = True
            DacSpan10.Enabled = True
            LabeldacSpan10Delta.Enabled = True
            Default11.Enabled = True
            Label149.Enabled = True
        Else
            Label229.Enabled = False
            LabelTemperature3.Enabled = False
            volts11.Enabled = False
            ButtonDacSpan10down.Enabled = False
            ButtonDacSpan10.Enabled = False
            ButtonDacSpan10Up.Enabled = False
            LabeldacSpan10Cal.Enabled = False
            Label150.Enabled = False
            DacSpan10.Enabled = False
            LabeldacSpan10Delta.Enabled = False
            Default11.Enabled = False
            Label149.Enabled = False
        End If

    End Sub



End Class


