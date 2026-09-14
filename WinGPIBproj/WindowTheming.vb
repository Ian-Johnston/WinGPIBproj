' Shared window theming helpers, used by every Form (and every runtime
' pop-up created via "New Form") across the app.

Imports System.Runtime.InteropServices

Module WindowTheming

    ' Use Win10-style square window corners on Windows 11. Harmless on
    ' earlier Windows versions - DWM just ignores an attribute it doesn't
    ' recognise.
    <DllImport("dwmapi.dll")>
    Private Function DwmSetWindowAttribute(
    hwnd As IntPtr,
    dwAttribute As Integer,
    ByRef pvAttribute As Integer,
    cbAttribute As Integer) As Integer
    End Function

    Private Const DWMWA_WINDOW_CORNER_PREFERENCE As Integer = 33
    Private Const DWMWCP_DONOTROUND As Integer = 1

    Public Sub ApplySquareCorners(f As Form)

        Dim cornerPreference As Integer = DWMWCP_DONOTROUND
        DwmSetWindowAttribute(f.Handle, DWMWA_WINDOW_CORNER_PREFERENCE, cornerPreference, Marshal.SizeOf(cornerPreference))

    End Sub

End Module
