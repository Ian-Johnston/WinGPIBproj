<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class JsonViewer
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
        Me.TreeViewJson = New System.Windows.Forms.TreeView()
        Me.SuspendLayout()
        '
        'TreeViewJson
        '
        Me.TreeViewJson.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TreeViewJson.Location = New System.Drawing.Point(2, 1)
        Me.TreeViewJson.Name = "TreeViewJson"
        Me.TreeViewJson.Size = New System.Drawing.Size(798, 448)
        Me.TreeViewJson.TabIndex = 638
        '
        'JsonViewer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.TreeViewJson)
        Me.Name = "JsonViewer"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TreeViewJson As TreeView
End Class
