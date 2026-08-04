
' CalRam HP 3458A

' NOTE:
' CalRAM upload code idea taken from kvez's Python code (with permission) here:
' https://github.com/kvez/HP-3458A-NVRAM-tool


'Imports System.Threading
'Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports IODevices
Imports PdfSharp.Pdf.Content.Objects

Partial Class Formtest

    Dim fs As System.IO.FileStream
    Dim fs2 As System.IO.FileStream
    Dim CalRamPathfile As System.IO.BinaryWriter
    Dim CalRamPathfile2 As System.IO.BinaryWriter
    Dim c As Char

    '3458A
    Dim Abort3458A As Boolean = False
    Dim Stepsize As Integer = 2
    Dim RamType As String = ""
    Dim RamType2 As String = ""
    Dim RAMfilename As String
    Dim RAMfilename2 As String
    Dim CalramAddress As Integer
    Dim Calrambytefordisplay As String
    Dim CalramStore(32768) As String
    Dim CalramStoreTemp1 As String = ""
    Dim CalramStoreTemp2 As String = ""
    Dim Counter As Integer = 1
    Dim Counter2 As Integer = 1
    Dim CalramValue As String = ""
    Dim CalAddrStart As Integer = 393216
    Dim CalAddrEnd As Integer = 397311
    Dim lineCountCalRam As Integer

    ' 3458A - Extract Cal data from a user selected .bin file and write to a .txt file
    Private Const BASE As Integer = &H60000

    Private Function ReadFloat64BE(bytes() As Byte, addr As Integer) As Double
        Dim off As Integer = addr - BASE
        If off < 0 OrElse off + 7 >= bytes.Length Then Return Double.NaN
        Dim b() As Byte = {bytes(off + 7), bytes(off + 6), bytes(off + 5), bytes(off + 4), bytes(off + 3), bytes(off + 2), bytes(off + 1), bytes(off + 0)}
        Return BitConverter.ToDouble(b, 0)
    End Function

    Private Function ReadUInt16BE(bytes() As Byte, addr As Integer) As UInteger
        Dim off As Integer = addr - BASE
        If off < 0 OrElse off + 1 >= bytes.Length Then Return 0UI
        Return (CUInt(bytes(off)) << 8) Or CUInt(bytes(off + 1))
    End Function

    Private Function ReadUInt32BE(bytes() As Byte, addr As Integer) As UInteger
        Dim off As Integer = addr - BASE
        If off < 0 OrElse off + 3 >= bytes.Length Then Return 0UI
        Return (CUInt(bytes(off)) << 24) Or (CUInt(bytes(off + 1)) << 16) Or (CUInt(bytes(off + 2)) << 8) Or CUInt(bytes(off + 3))
    End Function

    Private Function ReadUInt8(bytes() As Byte, addr As Integer) As UInteger
        Dim off As Integer = addr - BASE
        If off < 0 OrElse off >= bytes.Length Then Return 0UI
        Return CUInt(bytes(off))
    End Function

    ' Read zero-terminated ASCII at an address
    Private Function ReadAsciiAt(bytes() As Byte, addr As Integer, maxLen As Integer) As String
        Dim off As Integer = addr - BASE
        If off < 0 OrElse off >= bytes.Length Then Return String.Empty
        Dim n As Integer = Math.Min(maxLen, bytes.Length - off)
        Dim raw As String = Encoding.ASCII.GetString(bytes, off, n)
        Dim zero As Integer = raw.IndexOf(ChrW(0))
        If zero >= 0 Then raw = raw.Substring(0, zero)
        ' Keep only printable ASCII to be safe
        Dim sb As New StringBuilder()
        For Each ch As Char In raw
            If AscW(ch) >= 32 AndAlso AscW(ch) <= 126 Then sb.Append(ch)
        Next
        Return sb.ToString().Trim()
    End Function

    ' Parse Calstr (timestamp, temp, serial) and write a header line
    Private Sub ParseAndEmitCalStr(bytes() As Byte, sb As StringBuilder)
        ' In many dumps Calstr shows up at 0x605CA (example you shared).
        ' We’ll read ~64 bytes starting there.
        Dim raw As String = ReadAsciiAt(bytes, &H605CA, 64)

        If String.IsNullOrWhiteSpace(raw) Then
            ' Try a nearby offset (some variants place the string a tad earlier)
            raw = ReadAsciiAt(bytes, &H605C8, 64)
        End If

        If String.IsNullOrWhiteSpace(raw) Then Exit Sub

        ' Strip any surrounding single quotes the meter might have stored (e.g., '36.7')
        If raw.StartsWith("'") AndAlso raw.EndsWith("'") AndAlso raw.Length >= 2 Then
            raw = raw.Substring(1, raw.Length - 2)
        End If

        Dim tsOut As String = ""
        Dim tempOut As String = ""
        Dim serialOut As String = ""

        If raw.Contains("~"c) Then
            ' Expected form: YYYYMMDDHHMMSS~TEMP~SERIAL
            Dim parts = raw.Split("~"c)
            If parts.Length >= 3 Then
                Dim dt = parts(0).Trim()
                If dt.Length = 14 AndAlso dt.All(AddressOf Char.IsDigit) Then
                    tsOut = $"{dt.Substring(0, 4)}-{dt.Substring(4, 2)}-{dt.Substring(6, 2)} {dt.Substring(8, 2)}:{dt.Substring(10, 2)}:{dt.Substring(12, 2)}"
                End If
                Dim tstr = parts(1).Trim()
                Dim tv As Double
                If Double.TryParse(tstr, NumberStyles.Float, CultureInfo.InvariantCulture, tv) Then
                    tempOut = tv.ToString("0.0", CultureInfo.InvariantCulture) & " °C"
                Else
                    tempOut = tstr
                End If
                serialOut = parts(2).Trim()
            End If
        Else
            ' Could be just a temperature like 36.7
            Dim tv As Double
            If Double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, tv) Then
                tempOut = tv.ToString("0.0", CultureInfo.InvariantCulture) & " °C"
            Else
                ' Unknown text, show raw
                tempOut = raw
            End If
        End If

        Dim header As New StringBuilder()
        header.Append("Cal record")
        Dim first As Boolean = True
        If tsOut <> "" Then
            header.Append(": ").Append(tsOut)
            first = False
        End If
        If tempOut <> "" Then
            header.Append(If(first, ": ", " @ ")).Append(tempOut)
            first = False
        End If
        If serialOut <> "" Then
            header.Append(If(first, ": ", " | ")).Append("Serial: ").Append(serialOut)
        End If

        sb.AppendLine(header.ToString())
        sb.AppendLine()     ' blank line
    End Sub

    Private Function ReadInt32BE(bytes() As Byte, addr As Integer) As Integer
        Dim off As Integer = addr - BASE
        If off < 0 OrElse off + 3 >= bytes.Length Then Return 0

        Dim value As UInteger =
        (CUInt(bytes(off)) << 24) Or
        (CUInt(bytes(off + 1)) << 16) Or
        (CUInt(bytes(off + 2)) << 8) Or
        CUInt(bytes(off + 3))

        Return CInt(value)
    End Function

    Private Sub ButtonReadCalBin_Click(sender As Object, e As EventArgs) Handles ButtonReadCalBin.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Binary files (*.bin)|*.bin|All files (*.*)|*.*"
        ofd.Title = "Select HP 3458A CalRAM .bin File"
        ofd.InitialDirectory = strPath

        If ofd.ShowDialog() <> DialogResult.OK Then Exit Sub

        Dim binPath As String = ofd.FileName
        Dim txtPath As String = Path.Combine(Path.GetDirectoryName(binPath), Path.GetFileNameWithoutExtension(binPath) & "_decoded.txt")
        Dim bytes() As Byte = System.IO.File.ReadAllBytes(binPath)

        If bytes.Length <> 2048 Then
            MessageBox.Show(
            "This is not a valid 2048-byte HP 3458A CalRAM file." & vbCrLf & vbCrLf &
            "File size: " & bytes.Length.ToString(CultureInfo.InvariantCulture) & " bytes",
            "HP 3458A CalRAM",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning)

            Exit Sub
        End If

        Dim entries As New List(Of Tuple(Of Integer, String, String)) From {
        Tuple.Create(&H60000, "40Kohm reference", "dbl"),
        Tuple.Create(&H60008, "7Vdc reference", "dbl"),
        Tuple.Create(&H60010, "dcv zero front 100mV", "dbl"),
        Tuple.Create(&H60018, "dcv zero rear  100mV", "dbl"),
        Tuple.Create(&H60020, "dcv zero front 1V", "dbl"),
        Tuple.Create(&H60028, "dcv zero rear  1V", "dbl"),
        Tuple.Create(&H60030, "dcv zero front 10V", "dbl"),
        Tuple.Create(&H60038, "dcv zero rear  10V", "dbl"),
        Tuple.Create(&H60040, "dcv zero front 100V", "dbl"),
        Tuple.Create(&H60048, "dcv zero rear  100V", "dbl"),
        Tuple.Create(&H60050, "dcv zero front 1KV", "dbl"),
        Tuple.Create(&H60058, "dcv zero rear  1KV", "dbl"),
        Tuple.Create(&H60060, "ohm zero front 10", "dbl"),
        Tuple.Create(&H60068, "ohm zero front 100", "dbl"),
        Tuple.Create(&H60070, "ohm zero front 1K", "dbl"),
        Tuple.Create(&H60078, "ohm zero front 10K", "dbl"),
        Tuple.Create(&H60080, "ohm zero front 100K", "dbl"),
        Tuple.Create(&H60088, "ohm zero front 1M", "dbl"),
        Tuple.Create(&H60090, "ohm zero front 10M", "dbl"),
        Tuple.Create(&H60098, "ohm zero front 100M", "dbl"),
        Tuple.Create(&H600A0, "ohm zero front 1G", "dbl"),
        Tuple.Create(&H600A8, "ohm zero rear 10", "dbl"),
        Tuple.Create(&H600B0, "ohm zero rear 100", "dbl"),
        Tuple.Create(&H600B8, "ohm zero rear 1K", "dbl"),
        Tuple.Create(&H600C0, "ohm zero rear 10K", "dbl"),
        Tuple.Create(&H600C8, "ohm zero rear 100K", "dbl"),
        Tuple.Create(&H600D0, "ohm zero rear 1M", "dbl"),
        Tuple.Create(&H600D8, "ohm zero rear 10M", "dbl"),
        Tuple.Create(&H600E0, "ohm zero rear 100M", "dbl"),
        Tuple.Create(&H600E8, "ohm zero rear 1G", "dbl"),
        Tuple.Create(&H600F0, "ohmf zero front 10", "dbl"),
        Tuple.Create(&H600F8, "ohmf zero front 100", "dbl"),
        Tuple.Create(&H60100, "ohmf zero front 1K", "dbl"),
        Tuple.Create(&H60108, "ohmf zero front 10K", "dbl"),
        Tuple.Create(&H60110, "ohmf zero front 100K", "dbl"),
        Tuple.Create(&H60118, "ohmf zero front 1M", "dbl"),
        Tuple.Create(&H60120, "ohmf zero front 10M", "dbl"),
        Tuple.Create(&H60128, "ohmf zero front 100M", "dbl"),
        Tuple.Create(&H60130, "ohmf zero front 1G", "dbl"),
        Tuple.Create(&H60138, "ohmf zero rear 10", "dbl"),
        Tuple.Create(&H60140, "ohmf zero rear 100", "dbl"),
        Tuple.Create(&H60148, "ohmf zero rear 1K", "dbl"),
        Tuple.Create(&H60150, "ohmf zero rear 10K", "dbl"),
        Tuple.Create(&H60158, "ohmf zero rear 100K", "dbl"),
        Tuple.Create(&H60160, "ohmf zero rear 1M", "dbl"),
        Tuple.Create(&H60168, "ohmf zero rear 10M", "dbl"),
        Tuple.Create(&H60170, "ohmf zero rear 100M", "dbl"),
        Tuple.Create(&H60178, "ohmf zero rear 1G", "dbl"),
        Tuple.Create(&H60180, "autorange offset ohm 10", "i32"),
        Tuple.Create(&H60184, "autorange offset ohm 100", "i32"),
        Tuple.Create(&H60188, "autorange offset ohm 1K", "i32"),
        Tuple.Create(&H6018C, "autorange offset ohm 10K", "i32"),
        Tuple.Create(&H60190, "autorange offset ohm 100K", "i32"),
        Tuple.Create(&H60194, "autorange offset ohm 1M", "i32"),
        Tuple.Create(&H60198, "autorange offset ohm 10M", "i32"),
        Tuple.Create(&H6019C, "autorange offset ohm 100M", "i32"),
        Tuple.Create(&H601A0, "autorange offset ohm 1G", "i32"),
        Tuple.Create(&H601A4, "cal 0 temperature", "dbl"),
        Tuple.Create(&H601AC, "cal 10 temperature", "dbl"),
        Tuple.Create(&H601B4, "cal 10k temperature", "dbl"),
        Tuple.Create(&H601BC, "Cal_Sum0", "u16"),
        Tuple.Create(&H601BE, "vos dac", "u16"),
        Tuple.Create(&H601C0, "dci zero rear 100nA", "dbl"),
        Tuple.Create(&H601C8, "dci zero rear 1uA", "dbl"),
        Tuple.Create(&H601D0, "dci zero rear 10uA", "dbl"),
        Tuple.Create(&H601D8, "dci zero rear 100uA", "dbl"),
        Tuple.Create(&H601E0, "dci zero rear 1mA", "dbl"),
        Tuple.Create(&H601E8, "dci zero rear 10mA", "dbl"),
        Tuple.Create(&H601F0, "dci zero rear 100mA", "dbl"),
        Tuple.Create(&H601F8, "dci zero rear 1A", "dbl"),
        Tuple.Create(&H60200, "dcv gain 100mV", "dbl"),
        Tuple.Create(&H60208, "dcv gain 1V", "dbl"),
        Tuple.Create(&H60210, "dcv gain 10V", "dbl"),
        Tuple.Create(&H60218, "dcv gain 100V", "dbl"),
        Tuple.Create(&H60220, "dcv gain 1KV", "dbl"),
        Tuple.Create(&H60228, "ohm gain 10", "dbl"),
        Tuple.Create(&H60230, "ohm gain 100", "dbl"),
        Tuple.Create(&H60238, "ohm gain 1K", "dbl"),
        Tuple.Create(&H60240, "ohm gain 10K", "dbl"),
        Tuple.Create(&H60248, "ohm gain 100K", "dbl"),
        Tuple.Create(&H60250, "ohm gain 1M", "dbl"),
        Tuple.Create(&H60258, "ohm gain 10M", "dbl"),
        Tuple.Create(&H60260, "ohm gain 100M", "dbl"),
        Tuple.Create(&H60268, "ohm gain 1G", "dbl"),
        Tuple.Create(&H60270, "ohm ocomp gain 10", "dbl"),
        Tuple.Create(&H60278, "ohm ocomp gain 100", "dbl"),
        Tuple.Create(&H60280, "ohm ocomp gain 1K", "dbl"),
        Tuple.Create(&H60288, "ohm ocomp gain 10K", "dbl"),
        Tuple.Create(&H60290, "ohm ocomp gain 100K", "dbl"),
        Tuple.Create(&H60298, "ohm ocomp gain 1M", "dbl"),
        Tuple.Create(&H602A0, "ohm ocomp gain 10M", "dbl"),
        Tuple.Create(&H602A8, "ohm ocomp gain 100M", "dbl"),
        Tuple.Create(&H602B0, "ohm ocomp gain 1G", "dbl"),
        Tuple.Create(&H602B8, "dci gain 100nA", "dbl"),
        Tuple.Create(&H602C0, "dci gain 1uA", "dbl"),
        Tuple.Create(&H602C8, "dci gain 10uA", "dbl"),
        Tuple.Create(&H602D0, "dci gain 100uA", "dbl"),
        Tuple.Create(&H602D8, "dci gain 1mA", "dbl"),
        Tuple.Create(&H602E0, "dci gain 10mA", "dbl"),
        Tuple.Create(&H602E8, "dci gain 100mA", "dbl"),
        Tuple.Create(&H602F0, "dci gain 1A", "dbl"),
        Tuple.Create(&H602F8, "precharge dac", "u8"),
        Tuple.Create(&H602F9, "mc dac", "u8"),
        Tuple.Create(&H602FA, "high speed gain", "dbl"),
        Tuple.Create(&H60302, "il", "dbl"),
        Tuple.Create(&H6030A, "il2", "dbl"),
        Tuple.Create(&H60312, "rin", "dbl"),
        Tuple.Create(&H6031A, "low aperture", "dbl"),
        Tuple.Create(&H60322, "high aperture", "dbl"),
        Tuple.Create(&H6032A, "high aperture slope .01 PLC", "dbl"),
        Tuple.Create(&H60332, "high aperture slope .1 PLC", "dbl"),
        Tuple.Create(&H6033A, "high aperture null .01 PLC", "dbl"),
        Tuple.Create(&H60342, "high aperture null .1 PLC", "dbl"),
        Tuple.Create(&H6034A, "underload dcv 100mV", "u16"),
        Tuple.Create(&H6034E, "underload dcv 1V", "u16"),
        Tuple.Create(&H60352, "underload dcv 10V", "u16"),
        Tuple.Create(&H60356, "underload dcv 100V", "u16"),
        Tuple.Create(&H6035A, "underload dcv 1000V", "u16"),
        Tuple.Create(&H6035E, "overload dcv 100mV", "u16"),
        Tuple.Create(&H60362, "overload dcv 1V", "u16"),
        Tuple.Create(&H60366, "overload dcv 10V", "u16"),
        Tuple.Create(&H6036A, "overload dcv 100V", "u16"),
        Tuple.Create(&H6036E, "overload dcv 1000V", "u16"),
        Tuple.Create(&H60372, "underload ohm 10", "u16"),
        Tuple.Create(&H60376, "underload ohm 100", "u16"),
        Tuple.Create(&H6037A, "underload ohm 1K", "u16"),
        Tuple.Create(&H6037E, "underload ohm 10K", "u16"),
        Tuple.Create(&H60382, "underload ohm 100K", "u16"),
        Tuple.Create(&H60386, "underload ohm 1M", "u16"),
        Tuple.Create(&H6038A, "underload ohm 10M", "u16"),
        Tuple.Create(&H6038E, "underload ohm 100M", "u16"),
        Tuple.Create(&H60392, "underload ohm 1G", "u16"),
        Tuple.Create(&H60396, "overload ohm 10", "u16"),
        Tuple.Create(&H6039A, "overload ohm 100", "u16"),
        Tuple.Create(&H6039E, "overload ohm 1K", "u16"),
        Tuple.Create(&H603A2, "overload ohm 10K", "u16"),
        Tuple.Create(&H603A6, "overload ohm 100K", "u16"),
        Tuple.Create(&H603AA, "overload ohm 1M", "u16"),
        Tuple.Create(&H603AE, "overload ohm 10M", "u16"),
        Tuple.Create(&H603B2, "overload ohm 100M", "u16"),
        Tuple.Create(&H603B6, "overload ohm 1G", "u16"),
        Tuple.Create(&H603BA, "underload ohm ocomp 10", "u16"),
        Tuple.Create(&H603BE, "underload ohm ocomp 100", "u16"),
        Tuple.Create(&H603C2, "underload ohm ocomp 1K", "u16"),
        Tuple.Create(&H603C6, "underload ohm ocomp 10K", "u16"),
        Tuple.Create(&H603CA, "underload ohm ocomp 100K", "u16"),
        Tuple.Create(&H603CE, "underload ohm ocomp 1M", "u16"),
        Tuple.Create(&H603D2, "underload ohm ocomp 10M", "u16"),
        Tuple.Create(&H603D6, "underload ohm ocomp 100M", "u16"),
        Tuple.Create(&H603DA, "underload ohm ocomp 1G", "u16"),
        Tuple.Create(&H603DE, "overload ohm ocomp 10", "u16"),
        Tuple.Create(&H603E2, "overload ohm ocomp 100", "u16"),
        Tuple.Create(&H603E6, "overload ohm ocomp 1K", "u16"),
        Tuple.Create(&H603EA, "overload ohm ocomp 10K", "u16"),
        Tuple.Create(&H603EE, "overload ohm ocomp 100K", "u16"),
        Tuple.Create(&H603F2, "overload ohm ocomp 1M", "u16"),
        Tuple.Create(&H603F6, "overload ohm ocomp 10M", "u16"),
        Tuple.Create(&H603FA, "overload ohm ocomp 100M", "u16"),
        Tuple.Create(&H603FE, "overload ohm ocomp 1G", "u16"),
        Tuple.Create(&H60402, "underload dci 100nA", "u16"),
        Tuple.Create(&H60406, "Cal_406", "u16"),
        Tuple.Create(&H6040A, "Cal_40a", "u16"),
        Tuple.Create(&H6040E, "Cal_40e", "u16"),
        Tuple.Create(&H60412, "Cal_412", "u16"),
        Tuple.Create(&H60416, "Cal_416", "u16"),
        Tuple.Create(&H6041A, "Cal_41a", "u16"),
        Tuple.Create(&H6041E, "Cal_41e", "u16"),
        Tuple.Create(&H60422, "overload dci 100nA", "u16"),
        Tuple.Create(&H60426, "Cal_426", "u16"),
        Tuple.Create(&H6042A, "Cal_42a", "u16"),
        Tuple.Create(&H6042E, "Cal_42e", "u16"),
        Tuple.Create(&H60432, "Cal_432", "u16"),
        Tuple.Create(&H60436, "Cal_436", "u16"),
        Tuple.Create(&H6043A, "Cal_43a", "u16"),
        Tuple.Create(&H6043E, "Cal_43e", "u16"),
        Tuple.Create(&H60442, "acal dcv temperature", "dbl"),
        Tuple.Create(&H6044A, "acal ohm temperature", "dbl"),
        Tuple.Create(&H60452, "acal acv temperature", "dbl"),
        Tuple.Create(&H6045A, "ac offset dac 10mV", "u8"),
        Tuple.Create(&H6045B, "ac offset dac 100mV", "u8"),
        Tuple.Create(&H6045C, "ac offset dac 1V", "u8"),
        Tuple.Create(&H6045D, "ac offset dac 10V", "u8"),
        Tuple.Create(&H6045E, "ac offset dac 100V", "u8"),
        Tuple.Create(&H6045F, "ac offset dac 1KV", "u8"),
        Tuple.Create(&H60460, "acdc offset dac 10mV", "u8"),
        Tuple.Create(&H60461, "acdc offset dac 100mV", "u8"),
        Tuple.Create(&H60462, "acdc offset dac 1V", "u8"),
        Tuple.Create(&H60463, "acdc offset dac 10V", "u8"),
        Tuple.Create(&H60464, "acdc offset dac 100V", "u8"),
        Tuple.Create(&H60465, "acdc offset dac 1KV", "u8"),
        Tuple.Create(&H60466, "acdci offset dac 100uA", "u8"),
        Tuple.Create(&H60467, "acdci offset dac 1mA", "u8"),
        Tuple.Create(&H60468, "acdci offset dac 10mA", "u8"),
        Tuple.Create(&H60469, "acdci offset dac 100mA", "u8"),
        Tuple.Create(&H6046A, "acdci offset dac 1A", "u8"),
        Tuple.Create(&H6046C, "flatness dac 10mV", "u16"),
        Tuple.Create(&H6046E, "flatness dac 100mV", "u16"),
        Tuple.Create(&H60470, "flatness dac 1V", "u16"),
        Tuple.Create(&H60472, "flatness dac 10V", "u16"),
        Tuple.Create(&H60474, "flatness dac 100V", "u16"),
        Tuple.Create(&H60476, "flatness dac 1KV", "u16"),
        Tuple.Create(&H60478, "level dac dc 1.2V", "u8"),
        Tuple.Create(&H60479, "level dac dc 12V", "u8"),
        Tuple.Create(&H6047C, "level dac ac 1.2V", "u8"),
        Tuple.Create(&H6047D, "level dac ac 12V", "u8"),
        Tuple.Create(&H6047E, "dcv trigger offset 100mV", "u8"),
        Tuple.Create(&H6047F, "dcv trigger offset 1V", "u8"),
        Tuple.Create(&H60480, "dcv trigger offset 10V", "u8"),
        Tuple.Create(&H60481, "dcv trigger offset 100V", "u8"),
        Tuple.Create(&H60482, "dcv trigger offset 1000V", "u8"),
        Tuple.Create(&H60484, "acdcv sync offset 10mV", "dbl"),
        Tuple.Create(&H6048C, "acdcv sync offset 100mV", "dbl"),
        Tuple.Create(&H60494, "acdcv sync offset 1V", "dbl"),
        Tuple.Create(&H6049C, "acdcv sync offset 10V", "dbl"),
        Tuple.Create(&H604A4, "acdcv sync offset 100V", "dbl"),
        Tuple.Create(&H604AC, "acdcv sync offset 1KV", "dbl"),
        Tuple.Create(&H604B4, "acv sync offset 10mV", "dbl"),
        Tuple.Create(&H604BC, "acv sync offset 100mV", "dbl"),
        Tuple.Create(&H604C4, "acv sync offset 1V", "dbl"),
        Tuple.Create(&H604CC, "acv sync offset 10V", "dbl"),
        Tuple.Create(&H604D4, "acv sync offset 100V", "dbl"),
        Tuple.Create(&H604DC, "acv sync offset 1KV", "dbl"),
        Tuple.Create(&H604E4, "acv sync gain 10mV", "dbl"),
        Tuple.Create(&H604EC, "acv sync gain 100mV", "dbl"),
        Tuple.Create(&H604F4, "acv sync gain 1V", "dbl"),
        Tuple.Create(&H604FC, "acv sync gain 10V", "dbl"),
        Tuple.Create(&H60504, "acv sync gain 100V", "dbl"),
        Tuple.Create(&H6050C, "acv sync gain 1KV", "dbl"),
        Tuple.Create(&H60514, "ab ratio", "dbl"),
        Tuple.Create(&H6051C, "gain ratio", "dbl"),
        Tuple.Create(&H60524, "acv ana gain 10mV", "dbl"),
        Tuple.Create(&H6052C, "acv ana gain 100mV", "dbl"),
        Tuple.Create(&H60534, "acv ana gain 1V", "dbl"),
        Tuple.Create(&H6053C, "acv ana gain 10V", "dbl"),
        Tuple.Create(&H60544, "acv ana gain 100V", "dbl"),
        Tuple.Create(&H6054C, "acv ana gain 1KV", "dbl"),
        Tuple.Create(&H60554, "acv ana offset 10mV", "dbl"),
        Tuple.Create(&H6055C, "acv ana offset 100mV", "dbl"),
        Tuple.Create(&H60564, "acv ana offset 1V", "dbl"),
        Tuple.Create(&H6056C, "acv ana offset 10V", "dbl"),
        Tuple.Create(&H60574, "acv ana offset 100V", "dbl"),
        Tuple.Create(&H6057C, "acv ana offset 1KV", "dbl"),
        Tuple.Create(&H60584, "rmsdc ratio", "dbl"),
        Tuple.Create(&H6058C, "sampdc ratio", "dbl"),
        Tuple.Create(&H60594, "aci gain", "dbl"),
        Tuple.Create(&H6059C, "Cal_Sum1", "u16"),
        Tuple.Create(&H6059E, "Cal_59e", "dbl"),
        Tuple.Create(&H605A6, "Cal_5a6", "dbl"),
        Tuple.Create(&H605AE, "Cal_5ae", "dbl"),
        Tuple.Create(&H605B6, "freq gain", "dbl"),
        Tuple.Create(&H605BE, "attenuator high frequency dac", "u8"),
        Tuple.Create(&H605C0, "amplifier high frequency dac 10mV", "u8"),
        Tuple.Create(&H605C1, "amplifier high frequency dac 100mV", "u8"),
        Tuple.Create(&H605C2, "amplifier high frequency dac 1V", "u8"),
        Tuple.Create(&H605C3, "amplifier high frequency dac 10V", "u8"),
        Tuple.Create(&H605C4, "amplifier high frequency dac 100V", "u8"),
        Tuple.Create(&H605C5, "amplifier high frequency dac 1KV", "u8"),
        Tuple.Create(&H605C6, "interpolator", "u8"),
        Tuple.Create(&H605C8, "Cal_Sum2", "u16"),
        Tuple.Create(&H605CA, "Calstr", "str"),
        Tuple.Create(&H6061A, "Calnum", "u32"),
        Tuple.Create(&H6061E, "Cal_SecureCode", "u32"),
        Tuple.Create(&H60622, "Cal_AcalSecure", "u8"),
        Tuple.Create(&H60624, "Cal_Sum3", "u16"),
        Tuple.Create(&H60626, "Destructive Overloads", "u32"),
        Tuple.Create(&H6062A, "Defeats", "u32")
    }

        Dim sb As New StringBuilder()
        ParseAndEmitCalStr(bytes, sb)

        Dim ts As String = Date.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture)

        For Each entry In entries
            Dim addr As Integer = entry.Item1
            Dim label As String = entry.Item2
            Dim kind As String = entry.Item3
            Dim off As Integer = addr - BASE

            If off < 0 OrElse off >= bytes.Length Then Continue For

            Dim valueText As String = ""
            Dim entryByteCount As Integer = 0

            Select Case kind
                Case "dbl"
                    Dim value As Double = ReadFloat64BE(bytes, addr)
                    valueText = value.ToString("E13", CultureInfo.InvariantCulture)
                    entryByteCount = 8

                Case "i32"
                    Dim value As Integer = ReadInt32BE(bytes, addr)
                    valueText = value.ToString(CultureInfo.InvariantCulture)
                    entryByteCount = 4

                Case "u32"
                    Dim value As UInteger = ReadUInt32BE(bytes, addr)

                    If label = "Cal_SecureCode" Then
                        valueText = "0x" & value.ToString("X8", CultureInfo.InvariantCulture)
                    Else
                        valueText = value.ToString(CultureInfo.InvariantCulture)
                    End If

                    entryByteCount = 4

                Case "u16"
                    Dim value As UInteger = ReadUInt16BE(bytes, addr)

                    If label.StartsWith("Cal_Sum", StringComparison.OrdinalIgnoreCase) Then
                        valueText = "0x" & value.ToString("X4", CultureInfo.InvariantCulture)
                    Else
                        valueText = value.ToString(CultureInfo.InvariantCulture)
                    End If

                    entryByteCount = 2

                Case "u8"
                    Dim value As UInteger = ReadUInt8(bytes, addr)
                    valueText = value.ToString(CultureInfo.InvariantCulture)
                    entryByteCount = 1

                Case "str"
                    Dim stringLength As Integer = Math.Min(80, bytes.Length - off)
                    Dim raw(stringLength - 1) As Byte
                    Array.Copy(bytes, off, raw, 0, stringLength)

                    For index As Integer = 0 To raw.Length - 1
                        If raw(index) = &HA0 Then raw(index) = &H20
                    Next

                    valueText = Encoding.ASCII.GetString(raw).TrimEnd(ChrW(0), " "c)
                    entryByteCount = stringLength

                Case Else
                    valueText = "UNKNOWN TYPE"
            End Select

            Dim availableBytes As Integer = Math.Max(0, Math.Min(entryByteCount, bytes.Length - off))
            Dim hexData As String = ""

            If availableBytes > 0 Then
                hexData = BitConverter.ToString(bytes, off, availableBytes).Replace("-", " ")
            End If

            sb.AppendLine($"{ts} {addr:X5} [{hexData}] - {valueText} - {label}")
        Next

        System.IO.File.WriteAllText(txtPath, sb.ToString())

        MessageBox.Show("Decoded calibration data written to:" & vbCrLf & txtPath, "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ButtonCalramDump_Click(sender As Object, e As EventArgs) Handles ButtonCalramDump3458A.Click

        '3458A
        respUSERTABonly = False

        ' run appropriate routine

        If AddressRangeC.Checked = True Then    ' Calram
            ' 0x60000...0x60fff, so issuing 2048 GPIB commands
            CalAddrStart = 393216
            CalAddrEnd = 397311
            Stepsize = 2
            RamType = "3458A_Cal_ram_"
            TextBoxCalRamFile.Text = ""
            TextBoxCalRamFile2.Text = ""
            Calramextract3458A()
        End If

        If AddressRangeD.Checked = True Then    ' Settings ram 1 & 2
            ' 0x120000...0x12ffff (1179648...1245182 decimal)
            CalAddrStart = 1179648
            CalAddrEnd = 1245183        '1245183 1212415
            Stepsize = 2
            RamType = "3458A_Settings_ram_L_U121_"
            RamType2 = "3458A_Settings_ram_U_U122_"
            TextBoxCalRamFile.Text = ""
            TextBoxCalRamFile2.Text = ""
            Settingsramextract3458A()
        End If

    End Sub

    Private Sub Calramextract3458A()

        ' 3458A

        Abort3458A = False

        CalramStatus.Text = "CHECKING SETUP"

        Me.Refresh()

        If ButtonDev1Run.Enabled = True Then      ' Device 1 is started

            'RAMfilename = CSVfilepath.Text & "\" & RamType & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            RAMfilename = strPath & "\" & RamType & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            c = Chr(9)
            fs = New System.IO.FileStream(RAMfilename, IO.FileMode.OpenOrCreate)
            'fs = New System.IO.FileStream(RAMfilename, IO.FileMode.Append)
            CalRamPathfile = New System.IO.BinaryWriter(fs)
            CalRamPathfile.Seek(0, System.IO.SeekOrigin.Begin)

            LabelCounter.Text = "0"
            Counter = 0
            Counter2 = 0

            TextBoxCalRamFile.Text = RAMfilename

            CalramStatus.Text = "STB MASK, POLLING, CALRAM PRE-RUN"
            System.Threading.Thread.Sleep(500)     ' 500mS delay
            Me.Refresh()

            ' Checkbox options
            If Dev1PollingEnable.Checked = True Then
                dev1.enablepoll = True
            Else
                dev1.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
            End If

            System.Threading.Thread.Sleep(250)     ' 250mS delay

            If Dev1STBMask.Text = "" Then
                Dev1STBMask.Text = "16"
            End If
            dev1.MAVmask = Val(Dev1STBMask.Text)
            If Dev1STBMask.Text = "0" Then
                dev1.enablepoll = False
                Dev1PollingEnable.Checked = False
            End If

            System.Threading.Thread.Sleep(250)     ' 250mS delay

            ' Send all lines from command CalRam PRE-RUN text box
            lineCountCalRam = CalRam3458APreRun.Lines.Count
            For i = 0 To (lineCountCalRam - 1)
                If IgnoreErrors1.Checked = False Then
                    dev1.SendAsync(CalRam3458APreRun.Lines(i), True)
                Else
                    dev1.SendAsync(CalRam3458APreRun.Lines(i), False)
                End If
                System.Threading.Thread.Sleep(250)     ' 250mS delay
            Next i

            txtr1a.Text = ""                       ' Prepare reply as empty

            ' 10 dummy reads to set the interface up (some take a read or two to start getting valid data, buffer flush maybe)
            CalramStatus.Text = "DUMMY READ - BUFFER FLUSH"
            For CalAddrtemp As Integer = 1 To 10 Step 1
                Dim r As IOQuery = Nothing
                dev1.QueryBlocking("MREAD " & CalAddrStart, r, False)
                Debug.WriteLine("BLOCKING DetermineQuery: ")

                Cbdev1(r)
                System.Threading.Thread.Sleep(50)     ' 50mS delay
            Next

            System.Threading.Thread.Sleep(250)     ' 250mS delay

            ' Retrieve the data
            For CalAddr As Integer = CalAddrStart To CalAddrEnd Step Stepsize

                If Abort3458A Then Exit For

                ' Update status
                CalramStatus.Text = "READING 2048 BYTES (1024 16bit)"

                ' Send MREAD command and process reply
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("MREAD " & CalAddr, q, False)
                Debug.WriteLine("BLOCKING DetermineQuery: ")

                Cbdev1(q)

                ' Store reply as hexadecimal and pad to 4 characters
                Dim hexValue As String = Hex(Val(txtr1a.Text)).PadLeft(4, "0"c)

                ' Strip first 4 characters if the value is longer
                If hexValue.Length > 4 Then
                    hexValue = hexValue.Substring(hexValue.Length - 4, 4)
                End If

                ' Extract high byte and ensure it's valid
                Dim highByte As String = hexValue.Substring(0, 2)

                ' Write high byte to binary file
                fs.WriteByte(Convert.ToByte(highByte, 16))

                ' Store value in array
                CalramStore(Counter) = highByte

                ' Update display
                'LabelCalRamAddress.Text = CalAddr.ToString()
                LabelCalRamAddressHex.Text = Convert.ToInt32(CalAddr).ToString("X") & "  (TARGET = 60FFF)"
                'LabelCalRamByte.Text = highByte
                CalramStatus.Text = $"{CalAddr} = {Val(txtr1a.Text)}"
                LabelCounter.Text = Counter.ToString()

                ' Increment counters
                Counter += 1
                Counter2 += 2

            Next

            ' Close file
            fs.Close()

            ' Tidy up
            LabelCounter.Text = "2048"                  ' fudged
            LabelCalRamAddressHex.Text = "60FFF"        ' fudged
            txtr1a.Text = ""
            txtr1a_disp.Text = ""

            ' Abort display update
            If Abort3458A = True Then
                Abort3458A = False
                CalramStatus.Text = "ABORTED!"
                TextBoxCalRamFile.Text = ""
                TextBoxCalRamFile2.Text = ""
                fs.Close()
            Else
                ' Finished
                CalramStatus.Text = "DONE!"
            End If

        Else

            ' GPIB Dev 1 has not been started
            CalramStatus.Text = "DEVICE 1 IS NOT STARTED"

        End If

    End Sub

    Private Sub Settingsramextract3458A()

        ' 3458A

        Abort3458A = False

        CalramStatus.Text = "CHECKING SETUP"

        Me.Refresh()

        If ButtonDev1Run.Enabled = True Then      ' Device 1 is started

            ' RAM0L Lower (U121)
            RAMfilename = strPath & "\" & RamType & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            c = Chr(9)
            fs = New System.IO.FileStream(RAMfilename, IO.FileMode.OpenOrCreate)
            'fs = New System.IO.FileStream(RAMfilename, IO.FileMode.Append)
            CalRamPathfile = New System.IO.BinaryWriter(fs)
            CalRamPathfile.Seek(0, System.IO.SeekOrigin.Begin)

            ' RAM0H Upper (U122)
            RAMfilename2 = strPath & "\" & RamType2 & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            c = Chr(9)
            fs2 = New System.IO.FileStream(RAMfilename2, IO.FileMode.OpenOrCreate)
            'fs = New System.IO.FileStream(RAMfilename, IO.FileMode.Append)
            CalRamPathfile2 = New System.IO.BinaryWriter(fs2)
            CalRamPathfile2.Seek(0, System.IO.SeekOrigin.Begin)

            LabelCounter.Text = "0"
            Counter = 0
            Counter2 = 0

            TextBoxCalRamFile.Text = RAMfilename    ' L
            TextBoxCalRamFile2.Text = RAMfilename2  ' U

            CalramStatus.Text = "STB MASK, POLLING, SSETTINGS PRE-RUN"
            System.Threading.Thread.Sleep(500)     ' 500mS delay
            Me.Refresh()

            ' Checkbox options
            If Dev1PollingEnable.Checked = True Then
                dev1.enablepoll = True
            Else
                dev1.enablepoll = False     'set to FALSE this if a device does not support polling ("poll timeout" is signalled)
            End If

            System.Threading.Thread.Sleep(250)     ' 250mS delay

            If Dev1STBMask.Text = "" Then
                Dev1STBMask.Text = "16"
            End If
            dev1.MAVmask = Val(Dev1STBMask.Text)
            If Dev1STBMask.Text = "0" Then
                dev1.enablepoll = False
                Dev1PollingEnable.Checked = False
            End If

            System.Threading.Thread.Sleep(250)     ' 250mS delay

            ' Send all lines from command CalRam PRE-RUN text box
            lineCountCalRam = CalRam3458APreRun.Lines.Count
            For i = 0 To (lineCountCalRam - 1)
                If IgnoreErrors1.Checked = False Then
                    dev1.SendAsync(CalRam3458APreRun.Lines(i), True)
                Else
                    dev1.SendAsync(CalRam3458APreRun.Lines(i), False)
                End If
                System.Threading.Thread.Sleep(250)     ' 250mS delay
            Next i

            txtr1a.Text = ""                       ' Prepare reply as empty

            ' 10 dummy reads to set the interface up (some take a read or two to start getting valid data, buffer flush maybe)
            CalramStatus.Text = "DUMMY READ - BUFFER FLUSH"
            For CalAddrtemp As Integer = 1 To 10 Step 1
                Dim r As IOQuery = Nothing
                dev1.QueryBlocking("MREAD " & CalAddrStart, r, False)
                Debug.WriteLine("BLOCKING DetermineQuery: ")

                Cbdev1(r)
                System.Threading.Thread.Sleep(50)     ' 50mS delay
            Next

            System.Threading.Thread.Sleep(250)     ' 250mS delay

            ' Retrieve the data
            For CalAddr As Integer = CalAddrStart To CalAddrEnd Step Stepsize

                If Abort3458A Then Exit For

                CalramStatus.Text = "READING 2 LOTS 32768 BYTES (2 LOTS 16384 16-bit)"

                ' Send MREAD command and process reply
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("MREAD " & CalAddr, q, False)
                Debug.WriteLine("BLOCKING DetermineQuery: ")

                Cbdev1(q)

                ' Store reply as hexadecimal
                Dim hexValue As String = Hex(Val(txtr1a.Text))

                ' If value is negative, strip leading 'FFFF'
                If hexValue.Length > 4 Then
                    hexValue = hexValue.Remove(0, 4)
                End If

                ' Pad to 4 characters
                hexValue = hexValue.PadLeft(4, "0"c)

                ' Split into high and low bytes
                Dim highByte As String = hexValue.Remove(2, 2)
                Dim lowByte As String = hexValue.Substring(2, 2)

                ' Write bytes to files
                fs.WriteByte(Convert.ToByte(lowByte, 16))
                fs2.WriteByte(Convert.ToByte(highByte, 16))

                ' Update array
                CalramStore(Counter) = hexValue

                ' Update display
                'LabelCalRamAddress.Text = CalAddr.ToString()
                LabelCalRamAddressHex.Text = Convert.ToInt32(CalAddr).ToString("X") & "  (TARGET = 12FFFF)"
                'LabelCalRamByte.Text = highByte & " " & lowByte
                CalramStatus.Text = $"{CalAddr} = {Val(txtr1a.Text)}"
                LabelCounter.Text = (Counter * 2).ToString()

                ' Increment counters
                Counter += 1
                Counter2 += 2

            Next

            ' Close both file
            fs.Close()
            fs2.Close()

            ' Tidy up
            LabelCounter.Text = "65536"                  ' fudged
            LabelCalRamAddressHex.Text = "12FFFF"        ' fudged
            txtr1a.Text = ""
            txtr1a_disp.Text = ""

            ' QFORMAT NORM, TRIG AUTO - set back to 3458A defaults
            'dev1.SendAsync("QFORMAT NORM", True)
            'dev1.SendAsync("TRIG AUTO", True)

            ' Abort display update
            If Abort3458A = True Then
                Abort3458A = False
                CalramStatus.Text = "ABORTED!"
                TextBoxCalRamFile.Text = ""
                TextBoxCalRamFile2.Text = ""
                fs.Close()
                fs2.Close()
            Else
                ' Finished
                CalramStatus.Text = "DONE!"
            End If

        Else

            ' GPIB Dev 1 has not been started
            CalramStatus.Text = "DEVICE 1 IS NOT STARTED"

        End If

    End Sub

    Private Sub ShowFilesCalRam_Click(sender As Object, e As EventArgs) Handles ShowFilesCalRam.Click
        'Process.Start("explorer.exe", String.Format("/n, /e, {0}", CSVfilepath.Text))
        Process.Start("explorer.exe", String.Format("/n, /e, {0}", strPath))
    End Sub

    Private Sub Button3458Aabort_Click(sender As Object, e As EventArgs) Handles Button3458Aabort.Click

        Abort3458A = True
        TextBoxCalRamFile.Text = ""
        TextBoxCalRamFile2.Text = ""
        respUSERTABonly = False

    End Sub





    ' CalRAM upload to 3458A
    ' 30/07/26
    ' Idea taken from kvez's Python code (with permission) here:
    ' https://github.com/kvez/HP-3458A-NVRAM-tool


    Private Const CalRam3458ABase As Integer = &H60000
    Private Const CalRam3458ACodeBase As Integer = &H12A000
    Private Const CalRam3458ACallbackBase As Integer = &H12A100
    Private Const CalRam3458ANmiTriggerPort As Integer = &HC0001


    Private Class CalRam3458AFirmwareConfig

        Public CallbackPointerAddress As Integer
        Public MagicDeafAddress As Integer
        Public MagicBad1Address As Integer


        Public ReadOnly Property SuccessFlagAddress As Integer
            Get
                Return CallbackPointerAddress + 4
            End Get
        End Property


        Public ReadOnly Property WeCloseValueAddress As Integer
            Get
                Return CallbackPointerAddress + 8
            End Get
        End Property

    End Class


    Private CalRam3458AFirmwareMajor As Integer = 0
    Private CalRam3458AFirmwareResponse As String = ""
    Private CalRam3458AInstrumentID As String = ""
    Private Const CalRamConfirmText As String = "I WISH TO OVERWRITE MY CALRAM"

    Private Const CalRam3458ADataBase As Integer = &H127200
    Private Const CalRam3458ABlockWords As Integer = 128

    Private CalRam3458AWriteInProgress As Boolean = False
    Private Abort3458ACalRamWrite As Boolean = False

    Private CalRam3458AVerifyInProgress As Boolean = False


    Private Sub Button3458ACalRamBrowse_Click(sender As Object, e As EventArgs) Handles Button3458ACalRamBrowse.Click

        Dim ofd As New OpenFileDialog

        ofd.Title = "Select HP 3458A CalRAM File"
        ofd.Filter = "CalRAM Files (*.bin)|*.bin|All Files (*.*)|*.*"
        ofd.InitialDirectory = strPath

        If ofd.ShowDialog() <> DialogResult.OK Then Exit Sub

        TextBox3458ACalRamWriteFile.Text = ofd.FileName

        Label3458ACalRamFileInfo.Text = "CHECKING FILE"
        Label3458AFirmware.Text = "NOT DETECTED"
        Label3458ACalRamStatus.Text = "CHECKING CALRAM FILE"

        ProgressBar3458ACalRam.Value = 0

        Button3458ACalRamTestWrite.Enabled = False
        Button3458ACalRamWrite.Enabled = False
        Button3458ACalRamVerify.Enabled = False

        Try

            Dim fi As New FileInfo(ofd.FileName)

            ' The HP 3458A CalRAM image must contain exactly 2048 bytes.
            If fi.Length <> 2048 Then

                Label3458ACalRamFileInfo.Text = "INVALID FILE SIZE (" & fi.Length.ToString(CultureInfo.InvariantCulture) & " BYTES)"
                Label3458ACalRamStatus.Text = "INVALID FILE"

                MessageBox.Show(
                "CALRAM FILE CHECK FAILED" &
                vbCrLf & vbCrLf &
                "File: " & Path.GetFileName(ofd.FileName) &
                vbCrLf &
                "File size: " & fi.Length.ToString(CultureInfo.InvariantCulture) & " bytes" &
                vbCrLf &
                "Required size: 2048 bytes" &
                vbCrLf & vbCrLf &
                "The selected file is not a valid 2048-byte HP 3458A CalRAM image.",
                "HP 3458A CalRAM File Check",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)

                Exit Sub

            End If

            Dim bytes() As Byte = System.IO.File.ReadAllBytes(ofd.FileName)

            ' Additional safety check in case the file changed between
            ' checking its length and reading it.
            If bytes.Length <> 2048 Then

                Label3458ACalRamFileInfo.Text = "INVALID FILE SIZE (" & bytes.Length.ToString(CultureInfo.InvariantCulture) & " BYTES)"
                Label3458ACalRamStatus.Text = "INVALID FILE"

                MessageBox.Show(
                "CALRAM FILE CHECK FAILED" &
                vbCrLf & vbCrLf &
                "File: " & Path.GetFileName(ofd.FileName) &
                vbCrLf &
                "File size changed while the file was being read." &
                vbCrLf &
                "Current size: " & bytes.Length.ToString(CultureInfo.InvariantCulture) & " bytes" &
                vbCrLf &
                "Required size: 2048 bytes",
                "HP 3458A CalRAM File Check",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)

                Exit Sub

            End If

            ' Reject a completely blank or erased file.
            Dim allZero As Boolean = True
            Dim allFF As Boolean = True

            For Each value As Byte In bytes

                If value <> &H0 Then allZero = False
                If value <> &HFF Then allFF = False

                If Not allZero AndAlso Not allFF Then Exit For

            Next

            If allZero Then

                Label3458ACalRamFileInfo.Text = "INVALID CALRAM FILE - ALL BYTES ARE 00"
                Label3458ACalRamStatus.Text = "INVALID FILE"

                MessageBox.Show(
                "CALRAM FILE CHECK FAILED" &
                vbCrLf & vbCrLf &
                "File: " & Path.GetFileName(ofd.FileName) &
                vbCrLf &
                "File size: 2048 bytes" &
                vbCrLf &
                "Data check: All bytes are 00" &
                vbCrLf & vbCrLf &
                "The selected file contains no valid calibration data.",
                "HP 3458A CalRAM File Check",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)

                Exit Sub

            End If

            If allFF Then

                Label3458ACalRamFileInfo.Text = "INVALID CALRAM FILE - ALL BYTES ARE FF"
                Label3458ACalRamStatus.Text = "INVALID FILE"

                MessageBox.Show(
                "CALRAM FILE CHECK FAILED" &
                vbCrLf & vbCrLf &
                "File: " & Path.GetFileName(ofd.FileName) &
                vbCrLf &
                "File size: 2048 bytes" &
                vbCrLf &
                "Data check: All bytes are FF" &
                vbCrLf & vbCrLf &
                "The selected file appears to be blank or erased.",
                "HP 3458A CalRAM File Check",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)

                Exit Sub

            End If

            ' The first 16 bytes should contain:
            ' 0x000 to 0x007 = 40K reference
            ' 0x008 to 0x00F = 7V reference
            Dim reference40K As Double = ReadFloat64BE(bytes, &H60000)
            Dim reference7V As Double = ReadFloat64BE(bytes, &H60008)

            ' Reject invalid IEEE-754 values.
            If Double.IsNaN(reference40K) OrElse Double.IsInfinity(reference40K) Then

                Label3458ACalRamFileInfo.Text = "INVALID CALRAM FILE - INVALID 40K REFERENCE"
                Label3458ACalRamStatus.Text = "INVALID FILE"

                MessageBox.Show(
                "CALRAM FILE CHECK FAILED" &
                vbCrLf & vbCrLf &
                "File: " & Path.GetFileName(ofd.FileName) &
                vbCrLf &
                "File size: 2048 bytes" &
                vbCrLf &
                "40K reference: Invalid floating-point value" &
                vbCrLf & vbCrLf &
                "The selected file does not appear to contain valid HP 3458A calibration data.",
                "HP 3458A CalRAM File Check",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)

                Exit Sub

            End If

            If Double.IsNaN(reference7V) OrElse Double.IsInfinity(reference7V) Then

                Label3458ACalRamFileInfo.Text = "INVALID CALRAM FILE - INVALID 7V REFERENCE"
                Label3458ACalRamStatus.Text = "INVALID FILE"

                MessageBox.Show(
                "CALRAM FILE CHECK FAILED" &
                vbCrLf & vbCrLf &
                "File: " & Path.GetFileName(ofd.FileName) &
                vbCrLf &
                "File size: 2048 bytes" &
                vbCrLf &
                "40K reference: " & reference40K.ToString("0.000000", CultureInfo.InvariantCulture) &
                vbCrLf &
                "7V reference: Invalid floating-point value" &
                vbCrLf & vbCrLf &
                "The selected file does not appear to contain valid HP 3458A calibration data.",
                "HP 3458A CalRAM File Check",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)

                Exit Sub

            End If

            ' Limits taken from the supplied HP 3458A CalRAM program.
            If reference40K < 39921.5 OrElse reference40K > 40079.2 Then

                Label3458ACalRamFileInfo.Text = "INVALID CALRAM FILE - 40K REFERENCE OUT OF RANGE"
                Label3458ACalRamStatus.Text = "INVALID FILE"

                MessageBox.Show(
                "CALRAM FILE CHECK FAILED" &
                vbCrLf & vbCrLf &
                "File: " & Path.GetFileName(ofd.FileName) &
                vbCrLf &
                "File size: 2048 bytes" &
                vbCrLf &
                "40K reference: " & reference40K.ToString("0.000000", CultureInfo.InvariantCulture) &
                vbCrLf &
                "Expected range: 39921.5 to 40079.2" &
                vbCrLf &
                "7V reference: " & reference7V.ToString("0.000000000", CultureInfo.InvariantCulture) &
                vbCrLf & vbCrLf &
                "The 40K reference value is outside the expected range.",
                "HP 3458A CalRAM File Check",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)

                Exit Sub

            End If

            If reference7V < 6.5 OrElse reference7V > 7.5 Then

                Label3458ACalRamFileInfo.Text = "INVALID CALRAM FILE - 7V REFERENCE OUT OF RANGE"
                Label3458ACalRamStatus.Text = "INVALID FILE"

                MessageBox.Show(
                "CALRAM FILE CHECK FAILED" &
                vbCrLf & vbCrLf &
                "File: " & Path.GetFileName(ofd.FileName) &
                vbCrLf &
                "File size: 2048 bytes" &
                vbCrLf &
                "40K reference: " & reference40K.ToString("0.000000", CultureInfo.InvariantCulture) &
                vbCrLf &
                "7V reference: " & reference7V.ToString("0.000000000", CultureInfo.InvariantCulture) &
                vbCrLf &
                "Expected range: 6.5 to 7.5" &
                vbCrLf & vbCrLf &
                "The 7V reference value is outside the expected range.",
                "HP 3458A CalRAM File Check",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)

                Exit Sub

            End If

            ' All file checks passed.
            Label3458ACalRamFileInfo.Text = "VALID 2048 BYTE CALRAM FILE"
            Label3458AFirmware.Text = "NOT DETECTED"
            Label3458ACalRamStatus.Text = "READY FOR TEST WRITE"

            ProgressBar3458ACalRam.Value = 0

            ' Check Device 1 is active before allowing TEST WRITE.
            If ButtonDev1Run.Enabled = False Then
                Button3458ACalRamTestWrite.Enabled = False
            Else
                Button3458ACalRamTestWrite.Enabled = True
            End If

            Button3458ACalRamWrite.Enabled = False
            Button3458ACalRamVerify.Enabled = False

            MessageBox.Show(
            "CALRAM FILE CHECK PASSED" &
            vbCrLf & vbCrLf &
            "File: " & Path.GetFileName(ofd.FileName) &
            vbCrLf &
            "File size: 2048 bytes" &
            vbCrLf &
            "Blank data check: Passed" &
            vbCrLf &
            "40K reference: " & reference40K.ToString("0.000000", CultureInfo.InvariantCulture) &
            vbCrLf &
            "7V reference: " & reference7V.ToString("0.000000000", CultureInfo.InvariantCulture) &
            vbCrLf & vbCrLf &
            "The file appears to be a valid HP 3458A CalRAM image and is ready for TEST WRITE.",
            "HP 3458A CalRAM File Check",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information)

        Catch ex As Exception

            Label3458ACalRamFileInfo.Text = "FILE ERROR"
            Label3458AFirmware.Text = "NOT DETECTED"
            Label3458ACalRamStatus.Text = "FILE ERROR"

            ProgressBar3458ACalRam.Value = 0

            Button3458ACalRamTestWrite.Enabled = False
            Button3458ACalRamWrite.Enabled = False
            Button3458ACalRamVerify.Enabled = False

            MessageBox.Show(
            "CALRAM FILE CHECK FAILED" &
            vbCrLf & vbCrLf &
            "File: " & Path.GetFileName(ofd.FileName) &
            vbCrLf &
            "The selected file could not be read or checked." &
            vbCrLf & vbCrLf &
            ex.Message,
            "HP 3458A CalRAM File Check",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error)

        End Try

    End Sub


    Private Function Query3458ACalRam(command As String) As String

        txtr1a.Text = ""

        Dim q As IOQuery = Nothing

        dev1.QueryBlocking(command, q, False)

        Cbdev1(q)

        Return txtr1a.Text.Trim()

    End Function


    Private Function Get3458AFirmwareMajor(firmwareReply As String, ByRef firmwareMajor As Integer) As Boolean

        firmwareMajor = 0

        If String.IsNullOrWhiteSpace(firmwareReply) Then
            Return False
        End If

        Try

            ' Examples accepted:
            ' REV 9,0
            ' 9,0
            ' REV 5,3
            ' REV 5.3,0

            Dim cleanReply As String = firmwareReply.ToUpperInvariant().
                Replace("REV", "").
                Trim()

            Dim firstPart As String = cleanReply.Split(","c)(0).Trim()

            Dim revisionValue As Double

            If Not Double.TryParse(firstPart, NumberStyles.Float, CultureInfo.InvariantCulture, revisionValue) Then

                Return False

            End If

            firmwareMajor = CInt(Math.Truncate(revisionValue))

            Return firmwareMajor >= 2 AndAlso firmwareMajor <= 9

        Catch

            Return False

        End Try

    End Function


    Private Sub Button3458ACalRamTestWrite_Click(sender As Object, e As EventArgs) Handles Button3458ACalRamTestWrite.Click

        Button3458ACalRamTestWrite.Enabled = False
        Button3458ACalRamWrite.Enabled = False
        Button3458ACalRamVerify.Enabled = False

        ProgressBar3458ACalRam.Value = 0

        Try

            If ButtonDev1Run.Enabled = False Then

                Label3458ACalRamStatus.Text = "DEVICE 1 IS NOT STARTED"

                Label3458AFirmware.Text = "NOT DETECTED"

                Exit Sub

            End If

            Label3458ACalRamStatus.Text = "STB MASK, POLLING, CALRAM PRE-RUN"

            Me.Refresh()

            If Dev1PollingEnable.Checked Then
                dev1.enablepoll = True
            Else
                dev1.enablepoll = False
            End If

            If String.IsNullOrWhiteSpace(Dev1STBMask.Text) Then
                Dev1STBMask.Text = "16"
            End If

            dev1.MAVmask = Val(Dev1STBMask.Text)

            If Dev1STBMask.Text = "0" Then

                dev1.enablepoll = False
                Dev1PollingEnable.Checked = False

            End If

            System.Threading.Thread.Sleep(250)

            lineCountCalRam = CalRam3458APreRun.Lines.Count

            For index As Integer = 0 To lineCountCalRam - 1

                Dim command As String = CalRam3458APreRun.Lines(index).Trim()

                If command <> "" Then

                    If IgnoreErrors1.Checked = False Then
                        dev1.SendAsync(command, True)
                    Else
                        dev1.SendAsync(command, False)
                    End If

                    System.Threading.Thread.Sleep(250)

                End If

            Next

            Label3458ACalRamStatus.Text = "CHECKING INSTRUMENT"

            Label3458AFirmware.Text = "CHECKING"

            ProgressBar3458ACalRam.Value = 10

            Me.Refresh()

            dev1.SendAsync("PRESET NORM", False)
            System.Threading.Thread.Sleep(800)

            dev1.SendAsync("END ALWAYS", False)
            System.Threading.Thread.Sleep(200)

            dev1.SendAsync("BEEP 1", False)
            System.Threading.Thread.Sleep(200)

            CalRam3458AInstrumentID = Query3458ACalRam("ID?")

            If CalRam3458AInstrumentID.IndexOf("HP3458A", StringComparison.OrdinalIgnoreCase) < 0 Then

                Throw New InvalidOperationException("The connected instrument did not identify as an HP 3458A." & vbCrLf & "ID response: " & CalRam3458AInstrumentID)

            End If

            ProgressBar3458ACalRam.Value = 25

            Label3458ACalRamStatus.Text = "CHECKING FIRMWARE"

            Me.Refresh()

            CalRam3458AFirmwareResponse = Query3458ACalRam("REV?")

            If Not Get3458AFirmwareMajor(CalRam3458AFirmwareResponse, CalRam3458AFirmwareMajor) Then

                Label3458AFirmware.Text = "UNSUPPORTED"

                Throw New InvalidOperationException("Unsupported HP 3458A firmware." & vbCrLf & "REV response: " & CalRam3458AFirmwareResponse)

            End If

            Label3458AFirmware.Text = CalRam3458AFirmwareResponse.ToUpperInvariant()

            ProgressBar3458ACalRam.Value = 40

            ' Read the original first CalRAM word.
            Label3458ACalRamStatus.Text = "READING ORIGINAL CALRAM WORD"

            Me.Refresh()

            Dim originalHighByte As Integer = MRead3458AHighByte(CalRam3458ABase)

            Dim originalLowByte As Integer = MRead3458AHighByte(CalRam3458ABase + 2)

            Dim originalWord As Integer = (originalHighByte << 8) Or originalLowByte

            ProgressBar3458ACalRam.Value = 55

            ' Write exactly the same value back.
            Label3458ACalRamStatus.Text = "PERFORMING UNCHANGED TEST WRITE"

            Me.Refresh()

            Dim jsrOK As Boolean = Write3458ACalRamWord(CalRam3458ABase, originalWord)

            ProgressBar3458ACalRam.Value = 80

            ' Read the word back again.
            Label3458ACalRamStatus.Text = "VERIFYING TEST WRITE"

            Me.Refresh()

            Dim readbackHighByte As Integer = MRead3458AHighByte(CalRam3458ABase)

            Dim readbackLowByte As Integer = MRead3458AHighByte(CalRam3458ABase + 2)

            Dim readbackWord As Integer = (readbackHighByte << 8) Or readbackLowByte

            If Not jsrOK Then

                Throw New InvalidOperationException("The NMI write routine returned an instrument error.")

            End If

            If readbackWord <> originalWord Then

                Throw New InvalidOperationException("Test-write verification failed." & vbCrLf & vbCrLf & "Original: 0x" & originalWord.ToString("X4") & vbCrLf & "Readback: 0x" &
                readbackWord.ToString("X4"))

            End If

            Label3458ACalRamStatus.Text = "TEST WRITE PASSED"

            ProgressBar3458ACalRam.Value = 100

            Button3458ACalRamVerify.Enabled = False

            ' Enable WRITE immediately if the user has already completed
            ' the confirmation requirements.
            If CheckBox3458ACalRamWriteConfirm.Checked AndAlso TextBox3458ACalRamConfirm.Text.Trim().ToUpperInvariant() = CalRamConfirmText Then

                Button3458ACalRamWrite.Enabled = True

            Else

                Button3458ACalRamWrite.Enabled = False

            End If

            MessageBox.Show("The unchanged CalRAM test write passed." & vbCrLf & vbCrLf & "Original word: 0x" & originalWord.ToString("X4") & vbCrLf & "Readback word: 0x" & readbackWord.ToString("X4"), "HP 3458A CalRAM", MessageBoxButtons.OK,
            MessageBoxIcon.Information)

        Catch ex As Exception

            Label3458ACalRamStatus.Text = "TEST WRITE FAILED"

            ProgressBar3458ACalRam.Value = 0

            Button3458ACalRamWrite.Enabled = False
            Button3458ACalRamVerify.Enabled = False

            MessageBox.Show("The HP 3458A CalRAM test write failed." & vbCrLf & vbCrLf & ex.Message, "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally

            Button3458ACalRamTestWrite.Enabled = TextBox3458ACalRamWriteFile.Text <> ""

        End Try

    End Sub


    Private Sub Initialise3458ACalRamControls()

        TextBox3458ACalRamWriteFile.Text = ""

        Label3458ACalRamFileInfo.Text = "NO FILE SELECTED"
        Label3458AFirmware.Text = "NOT DETECTED"
        Label3458ACalRamStatus.Text = "READY"

        ProgressBar3458ACalRam.Value = 0

        CheckBox3458ACalRamWriteConfirm.Checked = False

        Button3458ACalRamTestWrite.Enabled = False
        Button3458ACalRamWrite.Enabled = False
        Button3458ACalRamVerify.Enabled = False

        'Button3458ACalRamAbort.Enabled = True

        CalRam3458AFirmwareMajor = 0
        CalRam3458AFirmwareResponse = ""
        CalRam3458AInstrumentID = ""

        TextBox3458ACalRamConfirm.Text = ""

        TextBox3458ACalRamConfirm.ContextMenuStrip = New ContextMenuStrip()
        TextBox3458ACalRamConfirm.AllowDrop = False

    End Sub


    Private Sub Button3458ACalRamAbort_Click(sender As Object, e As EventArgs) Handles Button3458ACalRamAbort.Click

        If CalRam3458AWriteInProgress OrElse CalRam3458AVerifyInProgress Then

            Abort3458ACalRamWrite = True

            If CalRam3458AVerifyInProgress Then

                Label3458ACalRamStatus.Text = "VERIFY ABORT REQUESTED"

            Else

                Label3458ACalRamStatus.Text = "WRITE ABORT REQUESTED"

            End If

            Button3458ACalRamAbort.Enabled = False

            Application.DoEvents()

        Else

            Initialise3458ACalRamControls()

            Button3458ACalRamBrowse.Enabled = True
            Button3458ACalRamAbort.Enabled = True

            CheckBox3458ACalRamWriteConfirm.Enabled = True
            TextBox3458ACalRamConfirm.Enabled = True

        End If

    End Sub


    Private Sub CheckBox3458ACalRamWriteConfirm_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3458ACalRamWriteConfirm.CheckedChanged

        Button3458ACalRamWrite.Enabled = Label3458ACalRamStatus.Text = "TEST WRITE PASSED" AndAlso CheckBox3458ACalRamWriteConfirm.Checked AndAlso TextBox3458ACalRamConfirm.Text.
            Trim().
            ToUpperInvariant() = CalRamConfirmText

    End Sub


    Private Sub TextBox3458ACalRamConfirm_TextChanged(sender As Object, e As EventArgs) Handles TextBox3458ACalRamConfirm.TextChanged

        Button3458ACalRamWrite.Enabled = Label3458ACalRamStatus.Text = "TEST WRITE PASSED" AndAlso CheckBox3458ACalRamWriteConfirm.Checked AndAlso TextBox3458ACalRamConfirm.Text.
            Trim().
            ToUpperInvariant() = CalRamConfirmText

    End Sub


    Private Sub TextBox3458ACalRamConfirm_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3458ACalRamConfirm.KeyDown

        If (e.Control AndAlso e.KeyCode = Keys.V) OrElse (e.Shift AndAlso e.KeyCode = Keys.Insert) Then

            e.SuppressKeyPress = True
            e.Handled = True

        End If

    End Sub


    Private Function Get3458ACalRamFirmwareConfig(firmwareMajor As Integer) As CalRam3458AFirmwareConfig

        Dim config As New CalRam3458AFirmwareConfig

        Select Case firmwareMajor

            Case 7, 8, 9

                config.CallbackPointerAddress = &H121852
                config.MagicDeafAddress = &H121780
                config.MagicBad1Address = &H120C90

            Case 5, 6

                config.CallbackPointerAddress = &H121A4E
                config.MagicDeafAddress = &H12197C
                config.MagicBad1Address = &H120C62

            Case 4

                config.CallbackPointerAddress = &H121A38
                config.MagicDeafAddress = &H12196E
                config.MagicBad1Address = &H120C62

            Case 2, 3

                config.CallbackPointerAddress = &H1211E8
                config.MagicDeafAddress = &H120AFE
                config.MagicBad1Address = &H1211FC

            Case Else

                Throw New InvalidOperationException("Unsupported HP 3458A firmware revision.")

        End Select

        Return config

    End Function


    Private Sub Send3458ACalRamCommand(command As String)

        dev1.SendAsync(command, False)
        System.Threading.Thread.Sleep(50)

    End Sub


    Private Function MRead3458AWord(address As Integer) As Integer

        Dim response As String = Query3458ACalRam("MREAD " & address.ToString(CultureInfo.InvariantCulture))

        Dim signedValue As Integer

        If Not Integer.TryParse(response, NumberStyles.Integer, CultureInfo.InvariantCulture, signedValue) Then

            Throw New InvalidOperationException("Invalid MREAD response: " & response)

        End If

        If signedValue < -32768 OrElse signedValue > 32767 Then

            Throw New InvalidOperationException("MREAD response outside the 16-bit range: " & response)

        End If

        Return signedValue And &HFFFF

    End Function


    Private Function MRead3458AHighByte(address As Integer) As Integer

        Return (MRead3458AWord(address) >> 8) And &HFF

    End Function


    Private Sub MWrite3458AWord(address As Integer, value As Integer)

        Dim unsignedValue As Integer = value And &HFFFF
        Dim signedValue As Integer

        If unsignedValue <= 32767 Then
            signedValue = unsignedValue
        Else
            signedValue = unsignedValue - 65536
        End If

        Send3458ACalRamCommand("MWRITE " & address.ToString(CultureInfo.InvariantCulture) & "," & signedValue.ToString(CultureInfo.InvariantCulture))

    End Sub


    Private Sub MWrite3458AWords(startAddress As Integer, words As IEnumerable(Of Integer))

        Dim address As Integer = startAddress

        For Each word As Integer In words

            MWrite3458AWord(address, word)
            address += 2

        Next

    End Sub


    Private Function Flush3458AErrors() As String

        Dim response As String = ""

        For index As Integer = 1 To 8

            response = Query3458ACalRam("ERRSTR?")

            If response.IndexOf("NO ERROR", StringComparison.OrdinalIgnoreCase) >= 0 Then

                Return response

            End If

            System.Threading.Thread.Sleep(100)

        Next

        Return response

    End Function


    Private Function Run3458AJSR(address As Integer, waitMilliseconds As Integer) As String

        Send3458ACalRamCommand("JSR " & address.ToString(CultureInfo.InvariantCulture))

        System.Threading.Thread.Sleep(waitMilliseconds)

        Return Query3458ACalRam("ERRSTR?")

    End Function


    Private Sub Set3458ANmiMagicWords(config As CalRam3458AFirmwareConfig)

        MWrite3458AWord(config.MagicDeafAddress, &HDEAF)
        MWrite3458AWord(config.MagicBad1Address, &HBAD1)

        MWrite3458AWord(config.MagicDeafAddress + 2, &HACE)
        MWrite3458AWord(config.MagicBad1Address + 2, &HBEAD)

    End Sub


    Private Sub Write3458ASingleWordCallback()

        ' 68000 callback:
        ' MOVEP.W D2,0(A0)
        ' RTS

        Dim callbackWords() As Integer = {
            &H588, &H0, &H4E75
        }

        MWrite3458AWords(CalRam3458ACallbackBase, callbackWords)

    End Sub


    Private Function Build3458ACallbackSetupWords(config As CalRam3458AFirmwareConfig) As List(Of Integer)

        Dim words As New List(Of Integer)

        ' MOVEA.L #CallbackBase,A4
        words.Add(&H287C)
        words.Add((CalRam3458ACallbackBase >> 16) And &HFFFF)
        words.Add(CalRam3458ACallbackBase And &HFFFF)

        ' MOVE.L A4,CallbackPointerAddress
        words.Add(&H23CC)
        words.Add((config.CallbackPointerAddress >> 16) And &HFFFF)
        words.Add(config.CallbackPointerAddress And &HFFFF)

        ' MOVE.W #0,SuccessFlagAddress
        words.Add(&H33FC)
        words.Add(&H0)
        words.Add((config.SuccessFlagAddress >> 16) And &HFFFF)
        words.Add(config.SuccessFlagAddress And &HFFFF)

        Return words

    End Function


    Private Function Build3458ANmiTriggerWords(config As CalRam3458AFirmwareConfig) As List(Of Integer)

        Dim words As New List(Of Integer)

        ' MOVE.B WeCloseValue,D6
        words.Add(&H1C39)
        words.Add((config.WeCloseValueAddress >> 16) And &HFFFF)
        words.Add(config.WeCloseValueAddress And &HFFFF)

        ' ORI.W #$80,D6
        words.Add(&H46)
        words.Add(&H80)

        ' MOVE.B D6,NmiTriggerPort
        words.Add(&H13C6)
        words.Add((CalRam3458ANmiTriggerPort >> 16) And &HFFFF)
        words.Add(CalRam3458ANmiTriggerPort And &HFFFF)

        ' MOVE.W #9,D6
        words.Add(&H3C3C)
        words.Add(&H9)

        ' DBEQ D6,*
        words.Add(&H57CE)
        words.Add(&HFFFE)

        Return words

    End Function


    Private Function Write3458ACalRamWord(physicalAddress As Integer, wordValue As Integer) As Boolean

        Dim config As CalRam3458AFirmwareConfig = Get3458ACalRamFirmwareConfig(CalRam3458AFirmwareMajor)

        Flush3458AErrors()

        Set3458ANmiMagicWords(config)
        Write3458ASingleWordCallback()

        Dim mainCode As New List(Of Integer)

        ' TRAP #5
        mainCode.Add(&H4E45)

        ' MOVE.W #word,D2
        mainCode.Add(&H343C)
        mainCode.Add(wordValue And &HFFFF)

        ' MOVEA.L #physicalAddress,A0
        mainCode.Add(&H207C)
        mainCode.Add((physicalAddress >> 16) And &HFFFF)
        mainCode.Add(physicalAddress And &HFFFF)

        mainCode.AddRange(Build3458ACallbackSetupWords(config))

        mainCode.AddRange(Build3458ANmiTriggerWords(config))

        ' RTS
        mainCode.Add(&H4E75)

        MWrite3458AWords(CalRam3458ACodeBase, mainCode)

        Dim errorResponse As String = Run3458AJSR(CalRam3458ACodeBase, 1500)

        Return errorResponse.IndexOf("NO ERROR", StringComparison.OrdinalIgnoreCase) >= 0

    End Function


    Private Function Write3458ACalRamBlock(config As CalRam3458AFirmwareConfig, startWord As Integer, wordCount As Integer) As Boolean

        If wordCount < 1 OrElse wordCount > CalRam3458ABlockWords Then

            Throw New ArgumentOutOfRangeException(NameOf(wordCount), "CalRAM block size must be between 1 and 128 words.")

        End If

        Dim stagingAddress As Integer = CalRam3458ADataBase + (startWord * 2)

        Dim calRamAddress As Integer = CalRam3458ABase + (startWord * 4)

        ' Refresh the four firmware-specific NMI safety words
        ' immediately before each JSR block.
        Set3458ANmiMagicWords(config)

        Dim code As New List(Of Integer)

        ' TRAP #5
        code.Add(&H4E45)

        ' MOVEA.L #stagingAddress,A3
        code.Add(&H267C)
        code.Add((stagingAddress >> 16) And &HFFFF)
        code.Add(stagingAddress And &HFFFF)

        ' MOVEA.L #calRamAddress,A0
        code.Add(&H207C)
        code.Add((calRamAddress >> 16) And &HFFFF)
        code.Add(calRamAddress And &HFFFF)

        code.AddRange(Build3458ACallbackSetupWords(config))

        ' MOVE.W #wordCount-1,D0
        code.Add(&H303C)
        code.Add((wordCount - 1) And &HFFFF)

        ' Start of the loop.
        Dim loopStartWord As Integer = code.Count

        ' MOVE.W (A3)+,D2
        code.Add(&H341B)

        ' Trigger the Level-7 NMI.
        code.AddRange(Build3458ANmiTriggerWords(config))

        ' ADDQ.L #4,A0
        code.Add(&H5888)

        Dim dbraWord As Integer = code.Count

        ' 68000 DBRA displacement is relative to the address
        ' immediately after the DBRA opcode word.
        Dim displacement As Integer = (loopStartWord * 2) - ((dbraWord * 2) + 2)

        If displacement < Short.MinValue OrElse displacement > Short.MaxValue Then

            Throw New InvalidOperationException("Calculated DBRA displacement is outside the 16-bit range.")

        End If

        ' DBRA D0,loop
        code.Add(&H51C8)
        code.Add(displacement And &HFFFF)

        ' RTS
        code.Add(&H4E75)

        MWrite3458AWords(CalRam3458ACodeBase, code)

        Flush3458AErrors()

        Dim response As String = Run3458AJSR(CalRam3458ACodeBase, 1500)

        Return response.IndexOf("NO ERROR", StringComparison.OrdinalIgnoreCase) >= 0

    End Function


    Private Sub Set3458ACalRamWriteProgress(value As Integer, statusText As String)

        ProgressBar3458ACalRam.Value = Math.Max(ProgressBar3458ACalRam.Minimum, Math.Min(ProgressBar3458ACalRam.Maximum, value))

        Label3458ACalRamStatus.Text = statusText.ToUpperInvariant()

        Me.Refresh()
        Application.DoEvents()

    End Sub


    Private Sub Button3458ACalRamWrite_Click(sender As Object, e As EventArgs) Handles Button3458ACalRamWrite.Click

        If CalRam3458AWriteInProgress Then Exit Sub

        If Label3458ACalRamStatus.Text <> "TEST WRITE PASSED" Then

            MessageBox.Show("A successful TEST WRITE is required before writing CalRAM.", "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Exit Sub

        End If

        If Not CheckBox3458ACalRamWriteConfirm.Checked Then

            MessageBox.Show("You must tick the CalRAM overwrite confirmation box.", "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Exit Sub

        End If

        If TextBox3458ACalRamConfirm.Text.
            Trim().
            ToUpperInvariant() <> CalRamConfirmText Then

            MessageBox.Show("The CalRAM confirmation text is not correct.", "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Exit Sub

        End If

        Dim fileName As String = TextBox3458ACalRamWriteFile.Text.Trim()

        If fileName = "" OrElse Not System.IO.File.Exists(fileName) Then

            MessageBox.Show("The selected CalRAM file cannot be found.", "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Exit Sub

        End If

        Dim calRamData() As Byte

        Try

            calRamData = System.IO.File.ReadAllBytes(fileName)

        Catch ex As Exception

            MessageBox.Show("The CalRAM file could not be read." & vbCrLf & vbCrLf & ex.Message, "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Exit Sub

        End Try

        If calRamData.Length <> 2048 Then

            MessageBox.Show("The selected CalRAM file is no longer exactly 2048 bytes.", "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Exit Sub

        End If

        Dim confirmation As DialogResult = MessageBox.Show("This will overwrite the complete calibration RAM in the " & "connected HP 3458A." & vbCrLf & vbCrLf & "Do not switch off the meter or interrupt the GPIB connection." & vbCrLf & vbCrLf & "Continue with the CalRAM write?", "CONFIRM HP 3458A CALRAM WRITE", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2)

        If confirmation <> DialogResult.Yes Then Exit Sub

        CalRam3458AWriteInProgress = True
        Abort3458ACalRamWrite = False

        Button3458ACalRamBrowse.Enabled = False
        Button3458ACalRamTestWrite.Enabled = False
        Button3458ACalRamWrite.Enabled = False
        Button3458ACalRamVerify.Enabled = False

        CheckBox3458ACalRamWriteConfirm.Enabled = False
        TextBox3458ACalRamConfirm.Enabled = False

        ' Abort remains active.
        Button3458ACalRamAbort.Enabled = True

        Try

            Set3458ACalRamWriteProgress(0, "CHECKING INSTRUMENT")

            If ButtonDev1Run.Enabled = False Then

                Throw New InvalidOperationException("Device 1 is not started.")

            End If

            ' Confirm that the instrument is still present immediately
            ' before beginning the destructive operation.
            Dim idResponse As String = Query3458ACalRam("ID?")

            If idResponse.IndexOf("HP3458A", StringComparison.OrdinalIgnoreCase) < 0 Then

                Throw New InvalidOperationException("The connected instrument did not identify as an HP 3458A." & vbCrLf & "ID response: " & idResponse)

            End If

            Dim config As CalRam3458AFirmwareConfig = Get3458ACalRamFirmwareConfig(CalRam3458AFirmwareMajor)

            ' These four CalRAM word indexes contain the stored checksums.
            ' They are deliberately excluded from the main transfer and
            ' written individually at the very end.
            Dim checksumWordIndexes As New List(Of Integer) From {
                &H1BC \ 2, &H59C \ 2, &H5C8 \ 2, &H624 \ 2
            }

            checksumWordIndexes.Sort()

            ' ------------------------------------------------------------
            ' PHASE 1: Stage all 1024 words in Settings RAM.
            ' This does not yet write CalRAM.
            ' ------------------------------------------------------------

            For wordIndex As Integer = 0 To 1023

                If Abort3458ACalRamWrite Then
                    Throw New OperationCanceledException()
                End If

                Dim wordValue As Integer = (CInt(calRamData(wordIndex * 2)) << 8) Or CInt(calRamData((wordIndex * 2) + 1))

                MWrite3458AWord(CalRam3458ADataBase + (wordIndex * 2), wordValue)

                If wordIndex Mod 8 = 0 Then

                    Dim progress As Integer = CInt((wordIndex / 1023.0) * 50.0)

                    Set3458ACalRamWriteProgress(progress, "STAGING CALRAM DATA " & wordIndex.ToString(CultureInfo.InvariantCulture) & " / 1024")

                End If

            Next

            ' ------------------------------------------------------------
            ' PHASE 2: Restore three staging words known to be disturbed
            ' by a 3458A background process during the long staging period.
            ' ------------------------------------------------------------

            Set3458ACalRamWriteProgress(50, "CORRECTING STAGING DATA")

            Dim correctionIndexes() As Integer = {
                473, 509, 511
            }

            For Each wordIndex As Integer In correctionIndexes

                If Abort3458ACalRamWrite Then
                    Throw New OperationCanceledException()
                End If

                Dim wordValue As Integer = (CInt(calRamData(wordIndex * 2)) << 8) Or CInt(calRamData((wordIndex * 2) + 1))

                MWrite3458AWord(CalRam3458ADataBase + (wordIndex * 2), wordValue)

            Next

            ' The callback writes one staged word into CalRAM.
            Write3458ASingleWordCallback()

            ' ------------------------------------------------------------
            ' Build continuous ranges with the four checksum words removed.
            ' ------------------------------------------------------------

            Dim ranges As New List(Of Tuple(Of Integer, Integer))

            Dim rangeStart As Integer = 0

            For Each checksumIndex As Integer In checksumWordIndexes

                If rangeStart < checksumIndex Then

                    ranges.Add(Tuple.Create(rangeStart, checksumIndex - rangeStart))

                End If

                rangeStart = checksumIndex + 1

            Next

            If rangeStart < 1024 Then

                ranges.Add(Tuple.Create(rangeStart, 1024 - rangeStart))

            End If

            ' ------------------------------------------------------------
            ' PHASE 3: Write all normal CalRAM words in blocks of up to 128.
            ' ------------------------------------------------------------

            Dim completedWords As Integer = 0
            Dim blockNumber As Integer = 0

            For Each range As Tuple(Of Integer, Integer) In ranges

                Dim position As Integer = range.Item1

                Dim remaining As Integer = range.Item2

                While remaining > 0

                    If Abort3458ACalRamWrite Then
                        Throw New OperationCanceledException()
                    End If

                    Dim wordsThisBlock As Integer = Math.Min(CalRam3458ABlockWords, remaining)

                    blockNumber += 1

                    Dim progress As Integer = 50 + CInt((completedWords / 1024.0) * 45.0)

                    Set3458ACalRamWriteProgress(progress, "WRITING CALRAM BLOCK " & blockNumber.ToString(CultureInfo.InvariantCulture) & " - " & wordsThisBlock.ToString(CultureInfo.InvariantCulture) &
                        " WORDS")

                    If Not Write3458ACalRamBlock(config, position, wordsThisBlock) Then

                        Throw New InvalidOperationException("The HP 3458A reported an error while writing " & "CalRAM block " & blockNumber.ToString(CultureInfo.InvariantCulture) & ".")

                    End If

                    position += wordsThisBlock
                    remaining -= wordsThisBlock
                    completedWords += wordsThisBlock

                End While

            Next

            ' ------------------------------------------------------------
            ' PHASE 4: Write the four checksum words last.
            ' ------------------------------------------------------------

            For checksumNumber As Integer = 0 To checksumWordIndexes.Count - 1

                If Abort3458ACalRamWrite Then
                    Throw New OperationCanceledException()
                End If

                Dim checksumIndex As Integer = checksumWordIndexes(checksumNumber)

                Set3458ACalRamWriteProgress(96 + checksumNumber, "WRITING CALRAM CHECKSUM " & (checksumNumber + 1).ToString(CultureInfo.InvariantCulture) & " / 4")

                If Not Write3458ACalRamBlock(config, checksumIndex, 1) Then

                    Throw New InvalidOperationException("The HP 3458A reported an error while writing " & "checksum word " & (checksumNumber + 1).ToString(CultureInfo.InvariantCulture) & ".")

                End If

                completedWords += 1

            Next

            Set3458ACalRamWriteProgress(100, "CALRAM WRITE COMPLETE - READY TO VERIFY")

            Button3458ACalRamVerify.Enabled = True

            MessageBox.Show("The complete 2048-byte CalRAM image was written." & vbCrLf & vbCrLf & "Press VERIFY to read the CalRAM back and compare every byte.", "HP 3458A CalRAM", MessageBoxButtons.OK,
                MessageBoxIcon.Information)

        Catch ex As OperationCanceledException

            MessageBox.Show("The CalRAM write was aborted.", "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Initialise3458ACalRamControls()

        Catch ex As Exception

            Label3458ACalRamStatus.Text = "CALRAM WRITE FAILED"

            ProgressBar3458ACalRam.Value = 0

            Button3458ACalRamVerify.Enabled = False

            MessageBox.Show("The HP 3458A CalRAM write failed." & vbCrLf & vbCrLf & ex.Message, "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally

            CalRam3458AWriteInProgress = False
            Abort3458ACalRamWrite = False

            Button3458ACalRamAbort.Enabled = True

            If Label3458ACalRamStatus.Text = "CALRAM WRITE COMPLETE - READY TO VERIFY" Then

                Button3458ACalRamBrowse.Enabled = False
                Button3458ACalRamTestWrite.Enabled = False
                Button3458ACalRamWrite.Enabled = False
                Button3458ACalRamVerify.Enabled = True

                CheckBox3458ACalRamWriteConfirm.Enabled = False
                TextBox3458ACalRamConfirm.Enabled = False

            ElseIf TextBox3458ACalRamWriteFile.Text <> "" Then

                Button3458ACalRamBrowse.Enabled = True
                Button3458ACalRamTestWrite.Enabled = True

                CheckBox3458ACalRamWriteConfirm.Enabled = True
                TextBox3458ACalRamConfirm.Enabled = True

            Else

                Button3458ACalRamBrowse.Enabled = True

            End If

        End Try

    End Sub


    Private Sub Button3458ACalRamVerify_Click(sender As Object, e As EventArgs) Handles Button3458ACalRamVerify.Click

        If CalRam3458AVerifyInProgress Then Exit Sub

        Dim fileName As String = TextBox3458ACalRamWriteFile.Text.Trim()

        If String.IsNullOrWhiteSpace(fileName) OrElse Not System.IO.File.Exists(fileName) Then

            MessageBox.Show("The selected CalRAM file cannot be found.", "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Exit Sub

        End If

        Dim expectedData() As Byte

        Try

            expectedData = System.IO.File.ReadAllBytes(fileName)

        Catch ex As Exception

            MessageBox.Show("The selected CalRAM file could not be read." & vbCrLf & vbCrLf & ex.Message, "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Exit Sub

        End Try

        If expectedData.Length <> 2048 Then

            MessageBox.Show("The selected CalRAM file is no longer exactly 2048 bytes.", "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Exit Sub

        End If

        If ButtonDev1Run.Enabled = False Then

            Label3458ACalRamStatus.Text = "DEVICE 1 IS NOT STARTED"

            Exit Sub

        End If

        CalRam3458AVerifyInProgress = True
        Abort3458ACalRamWrite = False

        Button3458ACalRamBrowse.Enabled = False
        Button3458ACalRamTestWrite.Enabled = False
        Button3458ACalRamWrite.Enabled = False
        Button3458ACalRamVerify.Enabled = False

        CheckBox3458ACalRamWriteConfirm.Enabled = False
        TextBox3458ACalRamConfirm.Enabled = False

        Button3458ACalRamAbort.Enabled = True

        ProgressBar3458ACalRam.Value = 0
        Label3458ACalRamStatus.Text = "PREPARING CALRAM VERIFY"

        Me.Refresh()
        Application.DoEvents()

        Try

            ' Ensure the meter is still connected before starting.
            Dim idResponse As String = Query3458ACalRam("ID?")

            If idResponse.IndexOf("HP3458A", StringComparison.OrdinalIgnoreCase) < 0 Then

                Throw New InvalidOperationException("The connected instrument did not identify as an HP 3458A." & vbCrLf & "ID response: " & idResponse)

            End If

            ' Required for reliable MREAD operation.
            dev1.SendAsync("END ALWAYS", False)
            System.Threading.Thread.Sleep(200)

            dev1.SendAsync("BEEP 1", False)
            System.Threading.Thread.Sleep(200)

            Dim differences As New List(Of Tuple(Of Integer, Byte, Byte))

            For byteIndex As Integer = 0 To 2047

                If Abort3458ACalRamWrite Then
                    Throw New OperationCanceledException()
                End If

                ' Each CalRAM byte is returned in the high byte of an
                ' MREAD word. Physical addresses therefore advance by 2.
                Dim physicalAddress As Integer = CalRam3458ABase + (byteIndex * 2)

                Dim actualValue As Byte = CByte(MRead3458AHighByte(physicalAddress))

                Dim expectedValue As Byte = expectedData(byteIndex)

                If actualValue <> expectedValue Then

                    differences.Add(Tuple.Create(byteIndex, expectedValue, actualValue))

                End If

                ' Updating every 16 bytes keeps the display responsive
                ' without slowing down every individual read.
                If byteIndex Mod 16 = 0 OrElse byteIndex = 2047 Then

                    Dim progress As Integer = CInt(((byteIndex + 1) / 2048.0) * 100.0)

                    ProgressBar3458ACalRam.Value = Math.Min(100, progress)

                    Label3458ACalRamStatus.Text = "VERIFYING CALRAM " & (byteIndex + 1).ToString(CultureInfo.InvariantCulture) & " / 2048"

                    Me.Refresh()
                    Application.DoEvents()

                End If

            Next

            ProgressBar3458ACalRam.Value = 100

            If differences.Count = 0 Then

                Label3458ACalRamStatus.Text = "CALRAM VERIFIED - 2048 BYTES MATCH"

                Button3458ACalRamVerify.Enabled = False
                Button3458ACalRamWrite.Enabled = False
                Button3458ACalRamTestWrite.Enabled = False

                MessageBox.Show("CALRAM VERIFY PASSED" & vbCrLf & vbCrLf & "All 2048 bytes in the HP 3458A match the selected file.", "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Initialise3458ACalRamControls()

                Button3458ACalRamBrowse.Enabled = True
                Button3458ACalRamAbort.Enabled = True

                CheckBox3458ACalRamWriteConfirm.Checked = False
                TextBox3458ACalRamConfirm.Clear()

            Else

                Label3458ACalRamStatus.Text = "CALRAM VERIFY FAILED - " & differences.Count.ToString(CultureInfo.InvariantCulture) & " BYTE DIFFERENCES"

                Dim report As New StringBuilder

                report.AppendLine(differences.Count.ToString(CultureInfo.InvariantCulture) & " byte differences were found.")

                report.AppendLine()
                report.AppendLine("The first differences are:")

                report.AppendLine()

                Dim reportCount As Integer = Math.Min(20, differences.Count)

                For index As Integer = 0 To reportCount - 1

                    Dim difference = differences(index)

                    Dim offset As Integer = difference.Item1

                    Dim physicalAddress As Integer = CalRam3458ABase + (offset * 2)

                    report.AppendLine("Offset 0x" & offset.ToString("X4") & "  Address 0x" & physicalAddress.ToString("X6") & "  File 0x" & difference.Item2.ToString("X2") & "  Meter 0x" &
                        difference.Item3.ToString("X2"))

                Next

                If differences.Count > reportCount Then

                    report.AppendLine()
                    report.AppendLine("...and " & (differences.Count - reportCount).ToString(CultureInfo.InvariantCulture) & " more differences.")

                End If

                Button3458ACalRamVerify.Enabled = True
                Button3458ACalRamWrite.Enabled = False
                Button3458ACalRamTestWrite.Enabled = False

                MessageBox.Show(report.ToString(), "HP 3458A CalRAM Verify Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)

            End If

        Catch ex As OperationCanceledException

            Label3458ACalRamStatus.Text = "CALRAM VERIFY ABORTED"

            ProgressBar3458ACalRam.Value = 0

            MessageBox.Show("The CalRAM verification was aborted.", "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As Exception

            Label3458ACalRamStatus.Text = "CALRAM VERIFY FAILED"

            ProgressBar3458ACalRam.Value = 0

            MessageBox.Show("The HP 3458A CalRAM verification failed." & vbCrLf & vbCrLf & ex.Message, "HP 3458A CalRAM", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally

            CalRam3458AVerifyInProgress = False
            Abort3458ACalRamWrite = False

            Button3458ACalRamAbort.Enabled = True

            If Label3458ACalRamStatus.Text = "CALRAM VERIFIED - 2048 BYTES MATCH" Then

                Button3458ACalRamBrowse.Enabled = False
                Button3458ACalRamTestWrite.Enabled = False
                Button3458ACalRamWrite.Enabled = False
                Button3458ACalRamVerify.Enabled = False

                CheckBox3458ACalRamWriteConfirm.Enabled = False
                TextBox3458ACalRamConfirm.Enabled = False

            Else

                Button3458ACalRamBrowse.Enabled = True
                Button3458ACalRamVerify.Enabled = True

            End If

        End Try

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim frm As New Form With {
        .Text = "HP 3458A CalRAM Write Help / Info",
        .StartPosition = FormStartPosition.CenterParent,
        .FormBorderStyle = FormBorderStyle.FixedDialog,
        .ShowIcon = False,
        .ShowInTaskbar = False,
        .Width = 700,
        .Height = 600,
        .MinimizeBox = False,
        .MaximizeBox = False
    }

        Dim txt As New TextBox With {
        .Multiline = True,
        .ReadOnly = True,
        .WordWrap = True,
        .Dock = DockStyle.Fill,
        .Font = New Font("Segoe UI", 9),
        .BackColor = Color.White,
        .Text =
"HP 3458A CALRAM WRITE PROCEDURE" & vbCrLf & vbCrLf &
"1. Browse and select a 2048-byte HP 3458A CalRAM (.bin) file. WinGPIB verifies the file size and performs basic checks on the calibration data to help detect an invalid or corrupted file before writing." & vbCrLf & vbCrLf &
"2. Press TEST WRITE. This safely verifies the CalRAM write mechanism by reading the first calibration word, writing the identical value back using the Level-7 NMI routine, and confirming that it was written correctly. No calibration data is changed during this test." & vbCrLf & vbCrLf &
"3. Tick the confirmation box and type the confirmation phrase exactly as shown to enable WRITE CALRAM." & vbCrLf & vbCrLf &
"4. Press WRITE CALRAM. The selected CalRAM image is first loaded into temporary Settings RAM inside the meter, then copied into the protected CalRAM. The four checksum words are written last." & vbCrLf & vbCrLf &
"5. Press VERIFY. WinGPIB reads back all 2048 bytes from the HP 3458A and compares them with the selected file to confirm that every byte was written correctly." & vbCrLf & vbCrLf &
"HOW IT WORKS:" & vbCrLf &
"The HP 3458A CalRAM cannot be written directly using the normal MWRITE command. WinGPIB therefore uploads a small Motorola 68000 machine-code routine into the meter and uses the processor's Level-7 Non-Maskable Interrupt mechanism to open the protected CalRAM write window." & vbCrLf & vbCrLf &
"The complete CalRAM file is staged in unused Settings RAM before being transferred into CalRAM in blocks. Firmware-specific memory addresses are selected automatically after WinGPIB reads the instrument revision using REV?." & vbCrLf & vbCrLf &
"The checksum words are deliberately written last so that a partially completed transfer does not leave the CalRAM appearing valid. VERIFY is read-only and does not alter any calibration data." & vbCrLf & vbCrLf &
"IMPORTANT" & vbCrLf &
"• Only use a known-good CalRAM backup for the correct instrument." & vbCrLf &
"• Do not switch off the HP 3458A or disconnect the GPIB cable during WRITE or VERIFY." & vbCrLf &
"• The complete WRITE and VERIFY procedure typically takes several minutes."
    }

        Dim btn As New Button With {
        .Text = "OK",
        .DialogResult = DialogResult.OK,
        .Width = 100,
        .Height = 30,
        .Anchor = AnchorStyles.Bottom
    }

        Dim panel As New Panel With {
        .Dock = DockStyle.Bottom,
        .Height = 45
    }

        panel.Controls.Add(btn)

        AddHandler panel.Resize,
        Sub()
            btn.Left = (panel.ClientSize.Width - btn.Width) \ 2
            btn.Top = 7
        End Sub

        frm.Controls.Add(txt)
        frm.Controls.Add(panel)

        frm.AcceptButton = btn

        txt.SelectionStart = 0
        txt.SelectionLength = 0

        frm.ShowDialog(Me)

    End Sub

End Class
