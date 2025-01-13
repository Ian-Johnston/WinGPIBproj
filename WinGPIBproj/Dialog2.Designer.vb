<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Dialog2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.LabelWarning1 = New System.Windows.Forms.Label()
        Me.LabelWarning2 = New System.Windows.Forms.Label()
        Me.LabelWarning3 = New System.Windows.Forms.Label()
        Me.LabelWarning4 = New System.Windows.Forms.Label()
        Me.LabelWarning5 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(182, 186)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(84, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(8, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'LabelWarning1
        '
        Me.LabelWarning1.AutoSize = True
        Me.LabelWarning1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelWarning1.Location = New System.Drawing.Point(11, 19)
        Me.LabelWarning1.Name = "LabelWarning1"
        Me.LabelWarning1.Size = New System.Drawing.Size(28, 16)
        Me.LabelWarning1.TabIndex = 63
        Me.LabelWarning1.Text = "###"
        Me.LabelWarning1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelWarning2
        '
        Me.LabelWarning2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LabelWarning2.AutoSize = True
        Me.LabelWarning2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelWarning2.Location = New System.Drawing.Point(11, 48)
        Me.LabelWarning2.Name = "LabelWarning2"
        Me.LabelWarning2.Size = New System.Drawing.Size(28, 16)
        Me.LabelWarning2.TabIndex = 64
        Me.LabelWarning2.Text = "###"
        Me.LabelWarning2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelWarning3
        '
        Me.LabelWarning3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LabelWarning3.AutoSize = True
        Me.LabelWarning3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelWarning3.Location = New System.Drawing.Point(11, 76)
        Me.LabelWarning3.Name = "LabelWarning3"
        Me.LabelWarning3.Size = New System.Drawing.Size(28, 16)
        Me.LabelWarning3.TabIndex = 65
        Me.LabelWarning3.Text = "###"
        Me.LabelWarning3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelWarning4
        '
        Me.LabelWarning4.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LabelWarning4.AutoSize = True
        Me.LabelWarning4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelWarning4.Location = New System.Drawing.Point(11, 104)
        Me.LabelWarning4.Name = "LabelWarning4"
        Me.LabelWarning4.Size = New System.Drawing.Size(28, 16)
        Me.LabelWarning4.TabIndex = 66
        Me.LabelWarning4.Text = "###"
        Me.LabelWarning4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelWarning5
        '
        Me.LabelWarning5.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LabelWarning5.AutoSize = True
        Me.LabelWarning5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelWarning5.Location = New System.Drawing.Point(11, 133)
        Me.LabelWarning5.Name = "LabelWarning5"
        Me.LabelWarning5.Size = New System.Drawing.Size(28, 16)
        Me.LabelWarning5.TabIndex = 67
        Me.LabelWarning5.Text = "###"
        Me.LabelWarning5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Dialog2
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(436, 227)
        Me.ControlBox = False
        Me.Controls.Add(Me.LabelWarning5)
        Me.Controls.Add(Me.LabelWarning4)
        Me.Controls.Add(Me.LabelWarning3)
        Me.Controls.Add(Me.LabelWarning2)
        Me.Controls.Add(Me.LabelWarning1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Dialog2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "WinGPIB"
        Me.TopMost = True
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents LabelWarning1 As Label
    Friend WithEvents LabelWarning2 As Label
    Friend WithEvents LabelWarning3 As Label
    Friend WithEvents LabelWarning4 As Label
    Friend WithEvents LabelWarning5 As Label
End Class
