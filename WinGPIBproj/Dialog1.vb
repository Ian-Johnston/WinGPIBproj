' pop-up help dialog

Imports System.Windows.Forms

Public Class Dialog1

    Private Sub Dialog1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ApplySquareCorners(Me)
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

End Class
