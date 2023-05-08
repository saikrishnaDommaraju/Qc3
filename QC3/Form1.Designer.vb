<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Btn_help = New System.Windows.Forms.Button()
        Me.btn_run = New System.Windows.Forms.Button()
        Me.groupBox1 = New System.Windows.Forms.GroupBox()
        Me.rb_Sheetmetal = New System.Windows.Forms.RadioButton()
        Me.rb_metallic = New System.Windows.Forms.RadioButton()
        Me.rb_Composite = New System.Windows.Forms.RadioButton()
        Me.rb_single = New System.Windows.Forms.RadioButton()
        Me.rb_generalsheetCheck = New System.Windows.Forms.RadioButton()
        Me.groupBox2 = New System.Windows.Forms.GroupBox()
        Me.rb_A350_1000 = New System.Windows.Forms.RadioButton()
        Me.rb_A350_900 = New System.Windows.Forms.RadioButton()
        Me.rb_Bracketinstallation = New System.Windows.Forms.RadioButton()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        Me.groupBox4 = New System.Windows.Forms.GroupBox()
        Me.rb_Section1314 = New System.Windows.Forms.RadioButton()
        Me.groupBox3 = New System.Windows.Forms.GroupBox()
        Me.rb_Section1618 = New System.Windows.Forms.RadioButton()
        Me.groupBox5 = New System.Windows.Forms.GroupBox()
        Me.chkReport = New System.Windows.Forms.CheckBox()
        Me.groupBox1.SuspendLayout()
        Me.groupBox2.SuspendLayout()
        Me.groupBox4.SuspendLayout()
        Me.groupBox3.SuspendLayout()
        Me.groupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Btn_help
        '
        Me.Btn_help.Location = New System.Drawing.Point(468, 4)
        Me.Btn_help.Name = "Btn_help"
        Me.Btn_help.Size = New System.Drawing.Size(39, 23)
        Me.Btn_help.TabIndex = 6
        Me.Btn_help.Text = "?"
        Me.Btn_help.UseVisualStyleBackColor = True
        '
        'btn_run
        '
        Me.btn_run.Location = New System.Drawing.Point(403, 50)
        Me.btn_run.Name = "btn_run"
        Me.btn_run.Size = New System.Drawing.Size(75, 23)
        Me.btn_run.TabIndex = 4
        Me.btn_run.Text = "Run"
        Me.btn_run.UseVisualStyleBackColor = True
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.rb_Sheetmetal)
        Me.groupBox1.Controls.Add(Me.rb_metallic)
        Me.groupBox1.Controls.Add(Me.rb_Composite)
        Me.groupBox1.Location = New System.Drawing.Point(14, 23)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(154, 100)
        Me.groupBox1.TabIndex = 1
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "Type Of part / Assy /Env"
        '
        'rb_Sheetmetal
        '
        Me.rb_Sheetmetal.AutoSize = True
        Me.rb_Sheetmetal.Location = New System.Drawing.Point(20, 73)
        Me.rb_Sheetmetal.Name = "rb_Sheetmetal"
        Me.rb_Sheetmetal.Size = New System.Drawing.Size(94, 19)
        Me.rb_Sheetmetal.TabIndex = 4
        Me.rb_Sheetmetal.TabStop = True
        Me.rb_Sheetmetal.Text = "Sheet Metal"
        Me.rb_Sheetmetal.UseVisualStyleBackColor = True
        '
        'rb_metallic
        '
        Me.rb_metallic.AutoSize = True
        Me.rb_metallic.Location = New System.Drawing.Point(20, 45)
        Me.rb_metallic.Name = "rb_metallic"
        Me.rb_metallic.Size = New System.Drawing.Size(71, 19)
        Me.rb_metallic.TabIndex = 3
        Me.rb_metallic.TabStop = True
        Me.rb_metallic.Text = "Metallic"
        Me.rb_metallic.UseVisualStyleBackColor = True
        '
        'rb_Composite
        '
        Me.rb_Composite.AutoSize = True
        Me.rb_Composite.Checked = True
        Me.rb_Composite.Location = New System.Drawing.Point(20, 19)
        Me.rb_Composite.Name = "rb_Composite"
        Me.rb_Composite.Size = New System.Drawing.Size(87, 19)
        Me.rb_Composite.TabIndex = 2
        Me.rb_Composite.TabStop = True
        Me.rb_Composite.Text = "Composite"
        Me.rb_Composite.UseVisualStyleBackColor = True
        '
        'rb_single
        '
        Me.rb_single.AutoSize = True
        Me.rb_single.Location = New System.Drawing.Point(17, 49)
        Me.rb_single.Name = "rb_single"
        Me.rb_single.Size = New System.Drawing.Size(150, 19)
        Me.rb_single.TabIndex = 6
        Me.rb_single.TabStop = True
        Me.rb_single.Text = "Single / Equipped Part"
        Me.rb_single.UseVisualStyleBackColor = True
        '
        'rb_generalsheetCheck
        '
        Me.rb_generalsheetCheck.AutoSize = True
        Me.rb_generalsheetCheck.Checked = True
        Me.rb_generalsheetCheck.Location = New System.Drawing.Point(17, 23)
        Me.rb_generalsheetCheck.Name = "rb_generalsheetCheck"
        Me.rb_generalsheetCheck.Size = New System.Drawing.Size(144, 19)
        Me.rb_generalsheetCheck.TabIndex = 5
        Me.rb_generalsheetCheck.TabStop = True
        Me.rb_generalsheetCheck.Text = "General Sheet Check"
        Me.rb_generalsheetCheck.UseVisualStyleBackColor = True
        '
        'groupBox2
        '
        Me.groupBox2.Controls.Add(Me.rb_A350_1000)
        Me.groupBox2.Controls.Add(Me.rb_A350_900)
        Me.groupBox2.Location = New System.Drawing.Point(180, 23)
        Me.groupBox2.Name = "groupBox2"
        Me.groupBox2.Size = New System.Drawing.Size(200, 90)
        Me.groupBox2.TabIndex = 2
        Me.groupBox2.TabStop = False
        Me.groupBox2.Text = "Program"
        '
        'rb_A350_1000
        '
        Me.rb_A350_1000.AutoSize = True
        Me.rb_A350_1000.Location = New System.Drawing.Point(17, 59)
        Me.rb_A350_1000.Name = "rb_A350_1000"
        Me.rb_A350_1000.Size = New System.Drawing.Size(88, 19)
        Me.rb_A350_1000.TabIndex = 7
        Me.rb_A350_1000.TabStop = True
        Me.rb_A350_1000.Text = "A350-1000"
        Me.rb_A350_1000.UseVisualStyleBackColor = True
        '
        'rb_A350_900
        '
        Me.rb_A350_900.AutoSize = True
        Me.rb_A350_900.Checked = True
        Me.rb_A350_900.Location = New System.Drawing.Point(17, 27)
        Me.rb_A350_900.Name = "rb_A350_900"
        Me.rb_A350_900.Size = New System.Drawing.Size(81, 19)
        Me.rb_A350_900.TabIndex = 6
        Me.rb_A350_900.TabStop = True
        Me.rb_A350_900.Text = "A350-900"
        Me.rb_A350_900.UseVisualStyleBackColor = True
        '
        'rb_Bracketinstallation
        '
        Me.rb_Bracketinstallation.AutoSize = True
        Me.rb_Bracketinstallation.Location = New System.Drawing.Point(17, 77)
        Me.rb_Bracketinstallation.Name = "rb_Bracketinstallation"
        Me.rb_Bracketinstallation.Size = New System.Drawing.Size(131, 19)
        Me.rb_Bracketinstallation.TabIndex = 7
        Me.rb_Bracketinstallation.TabStop = True
        Me.rb_Bracketinstallation.Text = "Bracket Installation"
        Me.rb_Bracketinstallation.UseVisualStyleBackColor = True
        '
        'btn_Cancel
        '
        Me.btn_Cancel.Location = New System.Drawing.Point(403, 90)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(75, 23)
        Me.btn_Cancel.TabIndex = 5
        Me.btn_Cancel.Text = "Cancel"
        Me.btn_Cancel.UseVisualStyleBackColor = True
        '
        'groupBox4
        '
        Me.groupBox4.Controls.Add(Me.rb_Bracketinstallation)
        Me.groupBox4.Controls.Add(Me.rb_single)
        Me.groupBox4.Controls.Add(Me.rb_generalsheetCheck)
        Me.groupBox4.Location = New System.Drawing.Point(180, 119)
        Me.groupBox4.Name = "groupBox4"
        Me.groupBox4.Size = New System.Drawing.Size(200, 102)
        Me.groupBox4.TabIndex = 2
        Me.groupBox4.TabStop = False
        Me.groupBox4.Text = "Type of Drawing"
        '
        'rb_Section1314
        '
        Me.rb_Section1314.AutoSize = True
        Me.rb_Section1314.Checked = True
        Me.rb_Section1314.Location = New System.Drawing.Point(18, 29)
        Me.rb_Section1314.Name = "rb_Section1314"
        Me.rb_Section1314.Size = New System.Drawing.Size(104, 19)
        Me.rb_Section1314.TabIndex = 4
        Me.rb_Section1314.TabStop = True
        Me.rb_Section1314.Text = "Section 13-14"
        Me.rb_Section1314.UseVisualStyleBackColor = True
        '
        'groupBox3
        '
        Me.groupBox3.Controls.Add(Me.rb_Section1618)
        Me.groupBox3.Controls.Add(Me.rb_Section1314)
        Me.groupBox3.Location = New System.Drawing.Point(14, 129)
        Me.groupBox3.Name = "groupBox3"
        Me.groupBox3.Size = New System.Drawing.Size(154, 84)
        Me.groupBox3.TabIndex = 3
        Me.groupBox3.TabStop = False
        Me.groupBox3.Text = "Fuselage Section"
        '
        'rb_Section1618
        '
        Me.rb_Section1618.AutoSize = True
        Me.rb_Section1618.Location = New System.Drawing.Point(18, 59)
        Me.rb_Section1618.Name = "rb_Section1618"
        Me.rb_Section1618.Size = New System.Drawing.Size(104, 19)
        Me.rb_Section1618.TabIndex = 5
        Me.rb_Section1618.TabStop = True
        Me.rb_Section1618.Text = "Section 16-18"
        Me.rb_Section1618.UseVisualStyleBackColor = True
        '
        'groupBox5
        '
        Me.groupBox5.Controls.Add(Me.chkReport)
        Me.groupBox5.Controls.Add(Me.Btn_help)
        Me.groupBox5.Controls.Add(Me.btn_Cancel)
        Me.groupBox5.Controls.Add(Me.btn_run)
        Me.groupBox5.Controls.Add(Me.groupBox4)
        Me.groupBox5.Controls.Add(Me.groupBox1)
        Me.groupBox5.Controls.Add(Me.groupBox2)
        Me.groupBox5.Controls.Add(Me.groupBox3)
        Me.groupBox5.Location = New System.Drawing.Point(10, 8)
        Me.groupBox5.Name = "groupBox5"
        Me.groupBox5.Size = New System.Drawing.Size(519, 258)
        Me.groupBox5.TabIndex = 5
        Me.groupBox5.TabStop = False
        Me.groupBox5.Text = "Drawing QualityCheck Tool"
        '
        'chkReport
        '
        Me.chkReport.AutoSize = True
        Me.chkReport.Location = New System.Drawing.Point(34, 235)
        Me.chkReport.Name = "chkReport"
        Me.chkReport.Size = New System.Drawing.Size(163, 19)
        Me.chkReport.TabIndex = 7
        Me.chkReport.Text = "Open Report after Check"
        Me.chkReport.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(532, 268)
        Me.Controls.Add(Me.groupBox5)
        Me.Name = "Form1"
        Me.Text = "frmStart"
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        Me.groupBox2.ResumeLayout(False)
        Me.groupBox2.PerformLayout()
        Me.groupBox4.ResumeLayout(False)
        Me.groupBox4.PerformLayout()
        Me.groupBox3.ResumeLayout(False)
        Me.groupBox3.PerformLayout()
        Me.groupBox5.ResumeLayout(False)
        Me.groupBox5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Private WithEvents Btn_help As Button
    Private WithEvents btn_run As Button
    Private WithEvents groupBox1 As GroupBox
    Private WithEvents rb_Sheetmetal As RadioButton
    Private WithEvents rb_metallic As RadioButton
    Private WithEvents rb_Composite As RadioButton
    Private WithEvents rb_single As RadioButton
    Private WithEvents rb_generalsheetCheck As RadioButton
    Private WithEvents groupBox2 As GroupBox
    Private WithEvents rb_A350_1000 As RadioButton
    Private WithEvents rb_A350_900 As RadioButton
    Private WithEvents rb_Bracketinstallation As RadioButton
    Private WithEvents btn_Cancel As Button
    Private WithEvents groupBox4 As GroupBox
    Private WithEvents rb_Section1314 As RadioButton
    Private WithEvents groupBox3 As GroupBox
    Private WithEvents rb_Section1618 As RadioButton
    Private WithEvents groupBox5 As GroupBox
    Friend WithEvents chkReport As CheckBox
End Class
