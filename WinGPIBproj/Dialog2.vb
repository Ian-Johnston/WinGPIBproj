' pop-up dialog

Imports System.Windows.Forms
Imports System.Threading

Public Class Dialog2

    Public Warning1 As String
    Public Warning2 As String
    Public Warning3 As String

    Public Sub Dialog2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Me.Location = Windows.Forms.Cursor.Position
        LabelWarning1.Text = Warning1
        LabelWarning2.Text = Warning2
        LabelWarning3.Text = Warning3

    End Sub


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

End Class
