<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStatus
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
        Me.barStatus = New System.Windows.Forms.ProgressBar()
        Me.txtlabel = New System.Windows.Forms.Label()
        Me.txtlabel2 = New System.Windows.Forms.Label()
        Me.txtPercent = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'barStatus
        '
        Me.barStatus.Location = New System.Drawing.Point(29, 124)
        Me.barStatus.Name = "barStatus"
        Me.barStatus.Size = New System.Drawing.Size(493, 30)
        Me.barStatus.TabIndex = 0
        '
        'txtlabel
        '
        Me.txtlabel.AutoSize = True
        Me.txtlabel.Location = New System.Drawing.Point(26, 18)
        Me.txtlabel.Name = "txtlabel"
        Me.txtlabel.Size = New System.Drawing.Size(45, 15)
        Me.txtlabel.TabIndex = 1
        Me.txtlabel.Text = "Label1"
        '
        'txtlabel2
        '
        Me.txtlabel2.AutoSize = True
        Me.txtlabel2.Location = New System.Drawing.Point(26, 74)
        Me.txtlabel2.Name = "txtlabel2"
        Me.txtlabel2.Size = New System.Drawing.Size(45, 15)
        Me.txtlabel2.TabIndex = 2
        Me.txtlabel2.Text = "Label1"
        '
        'txtPercent
        '
        Me.txtPercent.Location = New System.Drawing.Point(538, 134)
        Me.txtPercent.Name = "txtPercent"
        Me.txtPercent.Size = New System.Drawing.Size(81, 20)
        Me.txtPercent.TabIndex = 3
        '
        'frmStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(638, 190)
        Me.Controls.Add(Me.txtPercent)
        Me.Controls.Add(Me.txtlabel2)
        Me.Controls.Add(Me.txtlabel)
        Me.Controls.Add(Me.barStatus)
        Me.Name = "frmStatus"
        Me.Text = "AXIS CADES METHODOLOGY"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents barStatus As ProgressBar
    Friend WithEvents txtlabel As Label
    Friend WithEvents txtlabel2 As Label
    Friend WithEvents txtPercent As TextBox
End Class
