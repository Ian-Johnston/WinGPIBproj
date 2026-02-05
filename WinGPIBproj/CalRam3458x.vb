
' CalRam HP 3458A
' CalRam HP 3457A

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

    '3457A
    Dim Abort3457A As Boolean = False
    Dim RAMfilename3457A As String
    Dim CalramAddress3457A As Integer
    'Dim CalramAddress3457AHex As String
    Dim CalramStore3457A(32768) As String
    Dim CalramStore3457Abyte1(32768) As String
    Dim CalramStore3457Abyte2(32768) As String
    Dim Counter3457A As Integer = 1
    Dim CalramValue3457A As String = ""
    Dim CalAddrStart3457A As Integer = 64
    Dim CalAddrEnd3457A As Integer = 511


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

    ' --- NEW: read zero-terminated ASCII at an address ---
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

    ' --- NEW: parse Calstr (timestamp, temp, serial) and write a header line ---
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

    Private Sub ButtonReadCalBin_Click(sender As Object, e As EventArgs) Handles ButtonReadCalBin.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Binary files (*.bin)|*.bin|All files (*.*)|*.*"
        ofd.Title = "Select HP3458A CALRAM .bin File"
        If ofd.ShowDialog() <> DialogResult.OK Then Exit Sub

        Dim binPath As String = ofd.FileName
        Dim txtPath As String = Path.Combine(Path.GetDirectoryName(binPath), Path.GetFileNameWithoutExtension(binPath) & "_decoded.txt")
        Dim bytes() As Byte = System.IO.File.ReadAllBytes(binPath)

        Dim entries As New List(Of Tuple(Of Integer, String, String)) From {
        Tuple.Create(&H60000, "40Kohm reference", "pair_float64"),
        Tuple.Create(&H60008, "7Vdc reference", "pair_float64"),
        Tuple.Create(&H60010, "dcv zero front 100mV", "pair_float64"),
        Tuple.Create(&H60018, "dcv zero rear 100mV", "pair_float64"),
        Tuple.Create(&H60020, "dcv zero front 1V", "pair_float64"),
        Tuple.Create(&H60028, "dcv zero rear 1V", "pair_float64"),
        Tuple.Create(&H60030, "dcv zero front 10V", "pair_float64"),
        Tuple.Create(&H60038, "dcv zero rear 10V", "pair_float64"),
        Tuple.Create(&H60040, "dcv zero front 100V", "pair_float64"),
        Tuple.Create(&H60048, "dcv zero rear 100V", "pair_float64"),
        Tuple.Create(&H60050, "dcv zero front 1KV", "pair_float64"),
        Tuple.Create(&H60058, "dcv zero rear 1KV", "pair_float64"),
        Tuple.Create(&H60060, "ohm zero front 10", "pair_float64"),
        Tuple.Create(&H60068, "ohm zero front 100", "pair_float64"),
        Tuple.Create(&H60070, "ohm zero front 1K", "pair_float64"),
        Tuple.Create(&H60078, "ohm zero front 10K", "pair_float64"),
        Tuple.Create(&H60080, "ohm zero front 100K", "pair_float64"),
        Tuple.Create(&H60088, "ohm zero front 1M", "pair_float64"),
        Tuple.Create(&H60090, "ohm zero front 10M", "pair_float64"),
        Tuple.Create(&H60098, "ohm zero front 100M", "pair_float64"),
        Tuple.Create(&H600A0, "ohm zero front 1G", "pair_float64"),
        Tuple.Create(&H600A8, "ohm zero rear 10", "pair_float64"),
        Tuple.Create(&H600B0, "ohm zero rear 100", "pair_float64"),
        Tuple.Create(&H600B8, "ohm zero rear 1K", "pair_float64"),
        Tuple.Create(&H600C0, "ohm zero rear 10K", "pair_float64"),
        Tuple.Create(&H600C8, "ohm zero rear 100K", "pair_float64"),
        Tuple.Create(&H600D0, "ohm zero rear 1M", "pair_float64"),
        Tuple.Create(&H600D8, "ohm zero rear 10M", "pair_float64"),
        Tuple.Create(&H600E0, "ohm zero rear 100M", "pair_float64"),
        Tuple.Create(&H600E8, "ohm zero rear 1G", "pair_float64"),
        Tuple.Create(&H600F0, "ohmf zero front 10", "pair_float64"),
        Tuple.Create(&H600F8, "ohmf zero front 100", "pair_float64"),
        Tuple.Create(&H60100, "ohmf zero front 1K", "pair_float64"),
        Tuple.Create(&H60108, "ohmf zero front 10K", "pair_float64"),
        Tuple.Create(&H60110, "ohmf zero front 100K", "pair_float64"),
        Tuple.Create(&H60118, "ohmf zero front 1M", "pair_float64"),
        Tuple.Create(&H60120, "ohmf zero front 10M", "pair_float64"),
        Tuple.Create(&H60128, "ohmf zero front 100M", "pair_float64"),
        Tuple.Create(&H60130, "ohmf zero front 1G", "pair_float64"),
        Tuple.Create(&H60138, "ohmf zero rear 10", "pair_float64"),
        Tuple.Create(&H60140, "ohmf zero rear 100", "pair_float64"),
        Tuple.Create(&H60148, "ohmf zero rear 1K", "pair_float64"),
        Tuple.Create(&H60150, "ohmf zero rear 10K", "pair_float64"),
        Tuple.Create(&H60158, "ohmf zero rear 100K", "pair_float64"),
        Tuple.Create(&H60160, "ohmf zero rear 1M", "pair_float64"),
        Tuple.Create(&H60168, "ohmf zero rear 10M", "pair_float64"),
        Tuple.Create(&H60170, "ohmf zero rear 100M", "pair_float64"),
        Tuple.Create(&H60178, "ohmf zero rear 1G", "pair_float64"),
        Tuple.Create(&H60180, "autorange offset ohm 10", "pair_uint32"),
        Tuple.Create(&H60184, "autorange offset ohm 100", "pair_uint32"),
        Tuple.Create(&H60188, "autorange offset ohm 1k", "pair_uint32"),
        Tuple.Create(&H6018C, "autorange offset ohm 10k", "pair_uint32"),
        Tuple.Create(&H60190, "autorange offset ohm 100k", "pair_uint32"),
        Tuple.Create(&H60194, "autorange offset ohm 1M", "pair_uint32"),
        Tuple.Create(&H60198, "autorange offset ohm 10M", "pair_uint32"),
        Tuple.Create(&H6019C, "autorange offset ohm 100M", "pair_uint32"),
        Tuple.Create(&H601A0, "autorange offset ohm 1G", "pair_uint32"),
        Tuple.Create(&H601A4, "cal 0 temperature", "pair_float64"),
        Tuple.Create(&H601AC, "cal 10 temperature", "pair_float64"),
        Tuple.Create(&H601B4, "cal 10k temperature", "pair_float64"),
        Tuple.Create(&H601BC, "Cal_Sum0", "pair_uint16"),
        Tuple.Create(&H601BE, "vos dac", "pair_uint16"),
        Tuple.Create(&H601C0, "dci zero rear 100nA", "pair_float64"),
        Tuple.Create(&H601C8, "dci zero rear 1uA", "pair_float64"),
        Tuple.Create(&H601D0, "dci zero rear 10uA", "pair_float64"),
        Tuple.Create(&H601D8, "dci zero rear 100uA", "pair_float64"),
        Tuple.Create(&H601E0, "dci zero rear 1mA", "pair_float64"),
        Tuple.Create(&H601E8, "dci zero rear 10mA", "pair_float64"),
        Tuple.Create(&H601F0, "dci zero rear 100mA", "pair_float64"),
        Tuple.Create(&H601F8, "dci zero rear 1A", "pair_float64"),
        Tuple.Create(&H60200, "dcv gain 100mV", "pair_float64"),
        Tuple.Create(&H60208, "dcv gain 1V", "pair_float64"),
        Tuple.Create(&H60210, "dcv gain 10V", "pair_float64"),
        Tuple.Create(&H60218, "dcv gain 100V", "pair_float64"),
        Tuple.Create(&H60220, "dcv gain 1KV", "pair_float64"),
        Tuple.Create(&H60228, "ohm gain 10", "pair_float64"),
        Tuple.Create(&H60230, "ohm gain 100", "pair_float64"),
        Tuple.Create(&H60238, "ohm gain 1K", "pair_float64"),
        Tuple.Create(&H60240, "ohm gain 10K", "pair_float64"),
        Tuple.Create(&H60248, "ohm gain 100K", "pair_float64"),
        Tuple.Create(&H60250, "ohm gain 1M", "pair_float64"),
        Tuple.Create(&H60258, "ohm gain 10M", "pair_float64"),
        Tuple.Create(&H60260, "ohm gain 100M", "pair_float64"),
        Tuple.Create(&H60268, "ohm gain 1G", "pair_float64"),
        Tuple.Create(&H60270, "ohm ocomp gain 10", "pair_float64"),
        Tuple.Create(&H60278, "ohm ocomp gain 100", "pair_float64"),
        Tuple.Create(&H60280, "ohm ocomp gain 1k", "pair_float64"),
        Tuple.Create(&H60288, "ohm ocomp gain 10k", "pair_float64"),
        Tuple.Create(&H60290, "ohm ocomp gain 100k", "pair_float64"),
        Tuple.Create(&H60298, "ohm ocomp gain 1M", "pair_float64"),
        Tuple.Create(&H602A0, "ohm ocomp gain 10M", "pair_float64"),
        Tuple.Create(&H602A8, "ohm ocomp gain 100M", "pair_float64"),
        Tuple.Create(&H602B0, "ohm ocomp gain 1G", "pair_float64"),
        Tuple.Create(&H602B8, "dci gain 100nA", "pair_float64"),
        Tuple.Create(&H602C0, "dci gain 1uA", "pair_float64"),
        Tuple.Create(&H602C8, "dci gain 10uA", "pair_float64"),
        Tuple.Create(&H602D0, "dci gain 100uA", "pair_float64"),
        Tuple.Create(&H602D8, "dci gain 1mA", "pair_float64"),
        Tuple.Create(&H602E0, "dci gain 10mA", "pair_float64"),
        Tuple.Create(&H602E8, "dci gain 100mA", "pair_float64"),
        Tuple.Create(&H602F0, "dci gain 1A", "pair_float64"),
        Tuple.Create(&H602F8, "precharge dac", "pair_uint8"),
        Tuple.Create(&H602F9, "mc dac", "pair_uint8"),
        Tuple.Create(&H602FA, "high speed gain", "pair_float64"),
        Tuple.Create(&H60302, "il", "pair_float64"),
        Tuple.Create(&H6030A, "il2", "pair_float64"),
        Tuple.Create(&H60312, "rin", "pair_float64"),
        Tuple.Create(&H6031A, "low aperture", "pair_float64"),
        Tuple.Create(&H60322, "high aperture", "pair_float64"),
        Tuple.Create(&H6032A, "high aperture slope .01 PLC", "pair_float64"),
        Tuple.Create(&H60332, "high aperture slope .1 PLC", "pair_float64"),
        Tuple.Create(&H6033A, "high aperture null .01 PLC", "pair_float64"),
        Tuple.Create(&H60342, "high aperture null .1 PLC", "pair_float64"),
        Tuple.Create(&H6034A, "underload dcv 100mV", "pair_uint32"),
        Tuple.Create(&H6034E, "underload dcv 1V", "pair_uint32"),
        Tuple.Create(&H60352, "underload dcv 10V", "pair_uint32"),
        Tuple.Create(&H60356, "underload dcv 100V", "pair_uint32"),
        Tuple.Create(&H6035A, "underload dcv 1000V", "pair_uint32"),
        Tuple.Create(&H6035E, "overload dcv 100mV", "pair_uint32"),
        Tuple.Create(&H60362, "overload dcv 1V", "pair_uint32"),
        Tuple.Create(&H60366, "overload dcv 10V", "pair_uint32"),
        Tuple.Create(&H6036A, "overload dcv 100V", "pair_uint32"),
        Tuple.Create(&H6036E, "overtoad dcv 1000V", "pair_uint32"),
        Tuple.Create(&H60372, "underload ohm 10", "pair_uint32"),
        Tuple.Create(&H60376, "underload ohm 100", "pair_uint32"),
        Tuple.Create(&H6037A, "underload ohm 1k", "pair_uint32"),
        Tuple.Create(&H6037E, "underload ohm 10k", "pair_uint32"),
        Tuple.Create(&H60382, "underload ohm 100k", "pair_uint32"),
        Tuple.Create(&H60386, "underload ohm 1M", "pair_uint32"),
        Tuple.Create(&H6038A, "underload ohm 10M", "pair_uint32"),
        Tuple.Create(&H6038E, "underload ohm 100M", "pair_uint32"),
        Tuple.Create(&H60392, "underload ohm 1G", "pair_uint32"),
        Tuple.Create(&H60396, "overload ohm 10", "pair_uint32"),
        Tuple.Create(&H6039A, "overload ohm 100", "pair_uint32"),
        Tuple.Create(&H6039E, "overload ohm 1k", "pair_uint32"),
        Tuple.Create(&H603A2, "overload ohm 10k", "pair_uint32"),
        Tuple.Create(&H603A6, "overload ohm 100k", "pair_uint32"),
        Tuple.Create(&H603AA, "overload ohm 1M", "pair_uint32"),
        Tuple.Create(&H603AE, "overload ohm 10M", "pair_uint32"),
        Tuple.Create(&H603B2, "overload ohm 100M", "pair_uint32"),
        Tuple.Create(&H603B6, "overload ohm 1G", "pair_uint32"),
        Tuple.Create(&H603BA, "underload ohm ocomp 10", "pair_uint32"),
        Tuple.Create(&H603BE, "underload ohm ocomp 100", "pair_uint32"),
        Tuple.Create(&H603C2, "underload ohm ocomp 1k", "pair_uint32"),
        Tuple.Create(&H603C6, "underload ohm ocomp 10k", "pair_uint32"),
        Tuple.Create(&H603CA, "underload ohm ocomp 100k", "pair_uint32"),
        Tuple.Create(&H603CE, "underload ohm ocomp 1M", "pair_uint32"),
        Tuple.Create(&H603D2, "underload ohm ocomp 10M", "pair_uint32"),
        Tuple.Create(&H603D6, "underload ohm ocomp 100M", "pair_uint32"),
        Tuple.Create(&H603DA, "underload ohm ocomp 1G", "pair_uint32"),
        Tuple.Create(&H603DE, "overload ohm ocomp 10", "pair_uint32"),
        Tuple.Create(&H603E2, "overload ohm ocomp 100", "pair_uint32"),
        Tuple.Create(&H603E6, "overload ohm ocomp 1k", "pair_uint32"),
        Tuple.Create(&H603EA, "overload ohm ocomp 10k", "pair_uint32"),
        Tuple.Create(&H603EE, "overload ohm ocomp 100k", "pair_uint32"),
        Tuple.Create(&H603F2, "overload ohm ocomp 1M", "pair_uint32"),
        Tuple.Create(&H603F6, "overload ohm ocomp 10M", "pair_uint32"),
        Tuple.Create(&H603FA, "overload ohm ocomp 100M", "pair_uint32"),
        Tuple.Create(&H603FE, "overload ohm ocomp 1G", "pair_uint32"),
        Tuple.Create(&H60402, "underload dci 100nA", "pair_uint32"),
        Tuple.Create(&H60406, "Cal_406", "pair_uint32"),
        Tuple.Create(&H6040A, "Cal_40a", "pair_uint32"),
        Tuple.Create(&H6040E, "Cal_40e", "pair_uint32"),
        Tuple.Create(&H60412, "Cal_412", "pair_uint32"),
        Tuple.Create(&H60416, "Cal_416", "pair_uint32"),
        Tuple.Create(&H6041A, "Cal_41a", "pair_uint32"),
        Tuple.Create(&H6041E, "Cal_41e", "pair_uint32"),
        Tuple.Create(&H60422, "overload dci 100nA", "pair_uint32"),
        Tuple.Create(&H60426, "Cal_426", "pair_uint32"),
        Tuple.Create(&H6042A, "Cal_42a", "pair_uint32"),
        Tuple.Create(&H6042E, "Cal_42e", "pair_uint32"),
        Tuple.Create(&H60432, "Cal_432", "pair_uint32"),
        Tuple.Create(&H60436, "Cal_436", "pair_uint32"),
        Tuple.Create(&H6043A, "Cal_43a", "pair_uint32"),
        Tuple.Create(&H6043E, "Cal_43e", "pair_uint32"),
        Tuple.Create(&H60442, "acal dcv temperature", "pair_float64"),
        Tuple.Create(&H6044A, "acal ohm temperature", "pair_float64"),
        Tuple.Create(&H60452, "acal acv temperature", "pair_float64"),
        Tuple.Create(&H6045A, "ac offset dac 10mV", "pair_uint8"),
        Tuple.Create(&H6045B, "ac offset dac 100mV", "pair_uint8"),
        Tuple.Create(&H6045C, "ac offset dac 1V", "pair_uint8"),
        Tuple.Create(&H6045D, "ac offset dac 10V", "pair_uint8"),
        Tuple.Create(&H6045E, "ac offset dac 100V", "pair_uint8"),
        Tuple.Create(&H6045F, "ac offset dac 1KV", "pair_uint8"),
        Tuple.Create(&H60460, "acdc offset dac 10mV", "pair_uint8"),
        Tuple.Create(&H60461, "acdc offset dac 100mV", "pair_uint8"),
        Tuple.Create(&H60462, "acdc offset dac 1V", "pair_uint8"),
        Tuple.Create(&H60463, "acdc offset dac 10V", "pair_uint8"),
        Tuple.Create(&H60464, "acdc offset dac 100V", "pair_uint8"),
        Tuple.Create(&H60465, "acdc offset dac 1KV", "pair_uint8"),
        Tuple.Create(&H60466, "acdci offset dac 100uA", "pair_uint8"),
        Tuple.Create(&H60467, "acdci offset dac 1mA", "pair_uint8"),
        Tuple.Create(&H60468, "acdci offset dac 10mA", "pair_uint8"),
        Tuple.Create(&H60469, "acdci offset dac 100mA", "pair_uint8"),
        Tuple.Create(&H6046A, "acdci offset dac 1A", "pair_uint8"),
        Tuple.Create(&H6046C, "flatness dac 10mV", "pair_uint16"),
        Tuple.Create(&H6046E, "flatness dac 100mV", "pair_uint16"),
        Tuple.Create(&H60470, "flatness dac 1V", "pair_uint16"),
        Tuple.Create(&H60472, "flatness dac 10V", "pair_uint16"),
        Tuple.Create(&H60474, "flatness dac 100V", "pair_uint16"),
        Tuple.Create(&H60476, "flatness dac 1KV", "pair_uint16"),
        Tuple.Create(&H60478, "level dac dc 1.2V", "pair_uint8"),
        Tuple.Create(&H6047A, "level dac dc 12V", "pair_uint8"),
        Tuple.Create(&H6047C, "level dac ac 1.2V", "pair_uint8"),
        Tuple.Create(&H6047D, "level dac ac 12V", "pair_uint8"),
        Tuple.Create(&H6047E, "dcv trigger offset 100mV", "pair_uint8"),
        Tuple.Create(&H6047F, "dcv trigger offset 1V", "pair_uint8"),
        Tuple.Create(&H60480, "dcv trigger offset 10V", "pair_uint8"),
        Tuple.Create(&H60481, "dcv trigger offset 100V", "pair_uint8"),
        Tuple.Create(&H60482, "dcv trigger offset 1000V", "pair_uint8"),
        Tuple.Create(&H60484, "acdcv sync offset 10mV", "pair_float64"),
        Tuple.Create(&H6048C, "acdcv sync offset 100mV", "pair_float64"),
        Tuple.Create(&H60494, "acdcv sync offset 1V", "pair_float64"),
        Tuple.Create(&H6049C, "acdcv sync offset 10V", "pair_float64"),
        Tuple.Create(&H604A4, "acdcv sync offset 100V", "pair_float64"),
        Tuple.Create(&H604AC, "acdcv sync offset 1KV", "pair_float64"),
        Tuple.Create(&H604B4, "acv sync offset 10mV", "pair_float64"),
        Tuple.Create(&H604BC, "acv sync offset 100mV", "pair_float64"),
        Tuple.Create(&H604C4, "acv sync offset 1V", "pair_float64"),
        Tuple.Create(&H604CC, "acv sync offset 10V", "pair_float64"),
        Tuple.Create(&H604D4, "acv sync offset 100V", "pair_float64"),
        Tuple.Create(&H604DC, "acv sync offset 1KV", "pair_float64"),
        Tuple.Create(&H604E4, "acv sync gain 10mV", "pair_float64"),
        Tuple.Create(&H604EC, "acv sync gain 100mV", "pair_float64"),
        Tuple.Create(&H604F4, "acv sync gain 1V", "pair_float64"),
        Tuple.Create(&H604FC, "acv sync gain 10V", "pair_float64"),
        Tuple.Create(&H60504, "acv sync gain 100V", "pair_float64"),
        Tuple.Create(&H6050C, "acv sync gain 1KV", "pair_float64"),
        Tuple.Create(&H60514, "ab ratio", "pair_float64"),
        Tuple.Create(&H6051C, "gain ratio", "pair_float64"),
        Tuple.Create(&H60524, "acv ana gain 10mV", "pair_float64"),
        Tuple.Create(&H6052C, "acv ana gain 100mV", "pair_float64"),
        Tuple.Create(&H60534, "acv ana gain 1V", "pair_float64"),
        Tuple.Create(&H6053C, "acv ana gain 10V", "pair_float64"),
        Tuple.Create(&H60544, "acv ana gain 100V", "pair_float64"),
        Tuple.Create(&H6054C, "acv ana gain 1KV", "pair_float64"),
        Tuple.Create(&H60554, "acv ana offset 10mV", "pair_float64"),
        Tuple.Create(&H6055C, "acv ana offset 100mV", "pair_float64"),
        Tuple.Create(&H60564, "acv ana offset 1V", "pair_float64"),
        Tuple.Create(&H6056C, "acv ana offset 10V", "pair_float64"),
        Tuple.Create(&H60574, "acv ana offset 100V", "pair_float64"),
        Tuple.Create(&H6057C, "acv ana offset 1KV", "pair_float64"),
        Tuple.Create(&H60584, "rmsdc ratio", "pair_float64"),
        Tuple.Create(&H6058C, "sampdc ratio", "pair_float64"),
        Tuple.Create(&H60594, "aci gain", "pair_float64"),
        Tuple.Create(&H6059C, "Cal_Sum1", "pair_uint16"),
        Tuple.Create(&H6059E, "Cal_59e", "pair_float64"),
        Tuple.Create(&H605A6, "Cal_5a6", "pair_float64"),
        Tuple.Create(&H605AE, "Cal_5ae", "pair_float64"),
        Tuple.Create(&H605B6, "freq gain", "pair_float64"),
        Tuple.Create(&H605BE, "attenuator high frequency dac", "pair_uint8"),
        Tuple.Create(&H605C0, "amplifier high frequency dac 10mV", "pair_uint8"),
        Tuple.Create(&H605C1, "amplifier high frequency dac 100mV", "pair_uint8"),
        Tuple.Create(&H605C2, "amplifier high frequency dac 1V", "pair_uint8"),
        Tuple.Create(&H605C3, "amplifier high frequency dac 10V", "pair_uint8"),
        Tuple.Create(&H605C4, "amplifier high frequency dac 100V", "pair_uint8"),
        Tuple.Create(&H605C5, "amplifier high frequency dac 1KV", "pair_uint8"),
        Tuple.Create(&H605C6, "interpolator", "pair_uint8"),
        Tuple.Create(&H605C8, "Cal_Sum2", "pair_uint16"),
        Tuple.Create(&H605CA, "Calstr", "string30"),
        Tuple.Create(&H6061A, "Calnum", "uint32"),
        Tuple.Create(&H6061E, "Cal_SecureCode", "pair_uint16"),
        Tuple.Create(&H60622, "Cal_AcalSecure", "pair_uint16"),
        Tuple.Create(&H60624, "Cal_Sum3", "pair_uint16"),
        Tuple.Create(&H60626, "Destructive Overloads", "pair_uint16"),
        Tuple.Create(&H6062A, "Defeats", "pair_uint16")
    }

        Dim sb As New StringBuilder()

        ' --- NEW: emit Calstr header line if present ---
        ParseAndEmitCalStr(bytes, sb)
        ' --- NEW: CALNUM (32-bit big-endian) at base+0x615 and base+0x61A ---
        Dim ts As String = Date.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture)

        'Dim calnum As UInteger = ReadUInt32BE(bytes, &H6061A)
        'Dim off1 As Integer = &H6061A - BASE
        'Dim hexCal As String = BitConverter.ToString(bytes, off1, 4).Replace("-", " ")
        'sb.AppendLine($"{ts} 6061A [{hexCal}] - {calnum} - Calnum")

        For Each t In entries
            Dim addr As Integer = t.Item1
            Dim label As String = t.Item2
            Dim kind As String = t.Item3
            Dim off As Integer = addr - BASE
            If off < 0 OrElse off >= bytes.Length Then Continue For

            Dim curStr As String = ""
            Dim facStr As String = ""
            Dim entryByteCount As Integer = 0  ' <-- NEW

            Select Case kind
                Case "pair_float64"
                    Dim cur# = ReadFloat64BE(bytes, addr)
                    Dim fac# = ReadFloat64BE(bytes, addr + 8)
                    curStr = cur.ToString("E10", CultureInfo.InvariantCulture)
                    facStr = fac.ToString("E10", CultureInfo.InvariantCulture)
                    entryByteCount = 16   ' 8 + 8 bytes

                Case "pair_uint32"
                    Dim cur32 As UInteger = ReadUInt32BE(bytes, addr)
                    Dim fac32 As UInteger = ReadUInt32BE(bytes, addr + 4)
                    curStr = cur32.ToString("N0", CultureInfo.InvariantCulture)
                    facStr = fac32.ToString("N0", CultureInfo.InvariantCulture)
                    entryByteCount = 8    ' 4 + 4 bytes

                Case "pair_uint16"
                    Dim cur16 As UInteger = ReadUInt16BE(bytes, addr)
                    Dim fac16 As UInteger = ReadUInt16BE(bytes, addr + 2)
                    curStr = cur16.ToString("N0", CultureInfo.InvariantCulture)
                    facStr = fac16.ToString("N0", CultureInfo.InvariantCulture)
                    entryByteCount = 4    ' 2 + 2 bytes

                Case "pair_uint8"
                    Dim curB As UInteger = ReadUInt8(bytes, addr)
                    Dim facB As UInteger = ReadUInt8(bytes, addr + 1)
                    curStr = curB.ToString("N0", CultureInfo.InvariantCulture)
                    facStr = facB.ToString("N0", CultureInfo.InvariantCulture)
                    entryByteCount = 2    ' 1 + 1 byte
                Case "uint32"
                    Dim curVal As UInteger = ReadUInt32BE(bytes, addr)
                    curStr = curVal.ToString()
                    entryByteCount = 4
                Case "string30"
                    Dim txt As String = Encoding.ASCII.GetString(bytes, addr - BASE, 30).TrimEnd(ChrW(0))
                    curStr = txt
                    entryByteCount = 30
                Case "uint32"
                    Dim val32 As UInteger = ReadUInt32BE(bytes, addr)
                    curStr = val32.ToString()
                    entryByteCount = 4
            End Select

            ' Build hex for the whole pair (current + factory)
            Dim avail As Integer = Math.Max(0, Math.Min(entryByteCount, bytes.Length - (addr - BASE)))
            Dim hexData As String = If(avail > 0, BitConverter.ToString(bytes, addr - BASE, avail).Replace("-", " "), "")

            'sb.AppendLine($"{ts} {addr:X5} - {curStr} ({facStr}) - {label}")
            'sb.AppendLine($"{ts} {addr:X5} - {curStr} - {label}")
            sb.AppendLine($"{ts} {addr:X5} [{hexData}] - {curStr} - {label}")


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


            CalramStatus.Text = "SETTING UP GPIB: STB Mask, Polling, CalRam Pre-Run"
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

            CalramStatus.Text = "SETTING UP GPIB: STB Mask, Polling, Settings Pre-Run"
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


    Private Sub ButtonCalramDump3457A_Click(sender As Object, e As EventArgs) Handles ButtonCalramDump3457A.Click

        ' 3457A
        respUSERTABonly = False

        If AddressRangeA.Checked = True Then
            CalAddrStart3457A = 64
            CalAddrEnd3457A = 511
        End If
        If AddressRangeB.Checked = True Then
            CalAddrStart3457A = 20480
            CalAddrEnd3457A = 22527
        End If
        If AddressRangeF.Checked = True Then
            CalAddrStart3457A = Val(TextBox3457AFrom.Text)
            CalAddrEnd3457A = Val(TextBox3457ATo.Text)

            If (CalAddrStart3457A < 0) Then
                CalAddrStart3457A = 0
            End If

            If (CalAddrEnd3457A < 0) Then
                CalAddrEnd3457A = 0
            End If

            If (CalAddrStart3457A > 32767) Then
                CalAddrStart3457A = 32767
            End If

            If (CalAddrEnd3457A > 32767) Then
                CalAddrEnd3457A = 32767
            End If

            If (CalAddrEnd3457A < CalAddrStart3457A) Then
                CalAddrStart3457A = 0
                CalAddrEnd3457A = 32767
            End If
        End If

        LabelCounter3457A.Text = "0"
        Counter3457A = 0

        Abort3457A = False

        CalramStatus3457A.Text = "CHECKING SETUP"

        Me.Refresh()

        If ButtonDev1Run.Enabled = True Then      ' Device 1 is started

            'RAMfilename3457A = CSVfilepath.Text & "\" & "3457ACalram_" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            RAMfilename3457A = strPath & "\" & "3457ACalram_" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".bin"
            c = Chr(9)
            fs = New System.IO.FileStream(RAMfilename3457A, IO.FileMode.OpenOrCreate)
            'fs = New System.IO.FileStream(RAMfilename, IO.FileMode.Append)
            CalRamPathfile = New System.IO.BinaryWriter(fs)
            CalRamPathfile.Seek(0, System.IO.SeekOrigin.Begin)

            TextBoxCalRamFile3457A.Text = RAMfilename3457A

            CalramStatus3457A.Text = "SETTING UP GPIB"
            System.Threading.Thread.Sleep(500)     ' 500mS delay
            Me.Refresh()

            dev1.SendAsync("TRIG 4", True)      ' TRIG HOLD
            CalramStatus3457A.Text = "TRIG 4"
            System.Threading.Thread.Sleep(250)     ' 250mS delay
            Me.Refresh()

            txtr1a.Text = ""                       ' Prepare reply as empty


            ' 10 dummy reads to set the interface up (some take a read or two to start getting valid data, buffer flush maybe)
            CalramStatus3457A.Text = "DUMMY READ - BUFFER FLUSH"
            For CalAddrtemp As Integer = 1 To 10 Step 1
                Dim r As IOQuery = Nothing
                dev1.QueryBlocking("PEEK " & CalAddrStart3457A, r, False)
                Debug.WriteLine("BLOCKING DetermineQuery: ")

                Cbdev1(r)
                System.Threading.Thread.Sleep(50)     ' 50mS delay
            Next

            System.Threading.Thread.Sleep(250)     ' 250mS delay


            ' Retrieve the data
            For CalAddr3457A As Integer = CalAddrStart3457A To CalAddrEnd3457A Step 2      ' step 2 so even addresses only

                If Abort3457A = True Then
                    Exit For
                End If

                CalramStatus3457A.Text = "READING........"

                ' Send MREAD command with address and wait for reply
                Dim q As IOQuery = Nothing
                dev1.QueryBlocking("PEEK " & CalAddr3457A, q, False)
                Debug.WriteLine("BLOCKING DetermineQuery: ")

                Cbdev1(q)   ' Process reply which stores value in txtr1a.Text (see Formtest.vb)

                ' Got reply, store it in array
                CalramStore3457A(Counter3457A) = Hex(Val(txtr1a.Text))

                'Label127.Text = CalramStore3457A(Counter3457A)
                'Me.Refresh()

                If (Len(CalramStore3457A(Counter3457A)) > 4) Then     ' originally a negative number i.e. FFFFE3B9 so need to strip FFFF of beginning
                    CalramStore3457A(Counter3457A) = CalramStore3457A(Counter3457A).Remove(0, 4)  ' remove first 4 characters if 5 or more bytes long
                End If

                If (Len(CalramStore3457A(Counter3457A)) = 3) Then     ' FB9 should be 0FB9 so need to add a 0 to beginning
                    CalramStore3457A(Counter3457A) = "0" & CalramStore3457A(Counter3457A)
                End If

                If (Len(CalramStore3457A(Counter3457A)) = 1) Then     ' B should be 000B so need to add a 000 to beginning
                    CalramStore3457A(Counter3457A) = "000" & CalramStore3457A(Counter3457A)
                End If

                If (Len(CalramStore3457A(Counter3457A)) = 2) Then     ' B9 should be 00B9 so need to add a 00 to beginning
                    CalramStore3457A(Counter3457A) = "00" & CalramStore3457A(Counter3457A)
                End If

                If (Len(CalramStore3457A(Counter3457A)) = 3) Then     ' B should be 000B so need to add a 00 to beginning
                    CalramStore3457A(Counter3457A) = "0" & CalramStore3457A(Counter3457A)
                End If

                ' Now strip into two bytes
                If (Len(CalramStore3457A(Counter3457A)) = 4) Then     ' E3B9 so need to strip into two bytes
                    CalramStore3457Abyte1(Counter3457A) = CalramStore3457A(Counter3457A).Remove(2, 2)  ' remove last two characters for byte 1
                    CalramStore3457Abyte2(Counter3457A) = CalramStore3457A(Counter3457A).Remove(0, 2)  ' remove first two characters for byte 2....so xABCD becomes two bytes AB and CD
                End If

                'Label129.Text = CalramStore3457Abyte1(Counter3457A)
                'Label131.Text = CalramStore3457Abyte2(Counter3457A)

                ' Write to text box
                'Me.ListCalRam.Items.Insert(0, CalramStore3457A(Counter3457A))

                ' Write to binary file
                fs.WriteByte(Convert.ToByte(CalramStore3457Abyte1(Counter3457A), 16))
                fs.WriteByte(Convert.ToByte(CalramStore3457Abyte2(Counter3457A), 16))

                LabelCounter3457A.Text = Counter3457A
                LabelCalRamAddress3457A.Text = CalAddr3457A
                LabelCalRamAddress3457AHex.Text = String.Join(",", LabelCalRamAddress3457A.Text.Split(","c).        ' Hex conversion
                              Select(Function(x) _
                              Convert.ToInt32(x).ToString("X")))

                LabelCalRamByte3457A.Text = CalramStore3457A(Counter3457A)
                CalramStatus3457A.Text = CalAddr3457A & "=" & Int(Val(txtr1a.Text))     ' display
                Counter3457A = Counter3457A + 1   ' prepare for next loop

            Next

            ' Close file
            fs.Close()

            ' Abort display update
            If Abort3457A = True Then
                Abort3457A = False
                CalramStatus3457A.Text = "ABORTED!"
                TextBoxCalRamFile3457A.Text = ""
            Else
                ' Finished
                LabelCalRamAddress3457A.Text = CalAddrEnd3457A
                CalramStatus3457A.Text = "DONE!"
            End If

        Else

            ' GPIB Dev 1 has not been started
            CalramStatus3457A.Text = "DEVICE 1 IS NOT STARTED"

        End If
    End Sub
    Private Sub Button3458Aabort_Click(sender As Object, e As EventArgs) Handles Button3458Aabort.Click

        Abort3458A = True
        TextBoxCalRamFile.Text = ""
        TextBoxCalRamFile2.Text = ""
        respUSERTABonly = False

    End Sub
    Private Sub Button3457Aabort_Click(sender As Object, e As EventArgs) Handles Button3457Aabort.Click

        Abort3457A = True
        TextBoxCalRamFile3457A.Text = ""
        respUSERTABonly = False

    End Sub

End Class