<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmCreateUser
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbTitle = New System.Windows.Forms.TextBox()
        Me.tbSname = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbFname = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.rbBasicAccess = New System.Windows.Forms.RadioButton()
        Me.rbFullAccess = New System.Windows.Forms.RadioButton()
        Me.tbPassword = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbConfirmPassword = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnConfirmChanges = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Pronombre:"
        '
        'tbTitle
        '
        Me.tbTitle.Location = New System.Drawing.Point(102, 12)
        Me.tbTitle.Name = "tbTitle"
        Me.tbTitle.Size = New System.Drawing.Size(232, 26)
        Me.tbTitle.TabIndex = 0
        '
        'tbSname
        '
        Me.tbSname.Location = New System.Drawing.Point(89, 44)
        Me.tbSname.Name = "tbSname"
        Me.tbSname.Size = New System.Drawing.Size(658, 26)
        Me.tbSname.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Apellido:"
        '
        'tbFname
        '
        Me.tbFname.Location = New System.Drawing.Point(454, 12)
        Me.tbFname.Name = "tbFname"
        Me.tbFname.Size = New System.Drawing.Size(293, 26)
        Me.tbFname.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(369, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 18)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Nombre:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 103)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(114, 18)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Nivel de acceso:"
        '
        'rbBasicAccess
        '
        Me.rbBasicAccess.AutoSize = True
        Me.rbBasicAccess.Location = New System.Drawing.Point(179, 101)
        Me.rbBasicAccess.Name = "rbBasicAccess"
        Me.rbBasicAccess.Size = New System.Drawing.Size(67, 22)
        Me.rbBasicAccess.TabIndex = 3
        Me.rbBasicAccess.TabStop = True
        Me.rbBasicAccess.Text = "Basico"
        Me.rbBasicAccess.UseVisualStyleBackColor = True
        '
        'rbFullAccess
        '
        Me.rbFullAccess.AutoSize = True
        Me.rbFullAccess.Location = New System.Drawing.Point(275, 101)
        Me.rbFullAccess.Name = "rbFullAccess"
        Me.rbFullAccess.Size = New System.Drawing.Size(46, 22)
        Me.rbFullAccess.TabIndex = 4
        Me.rbFullAccess.TabStop = True
        Me.rbFullAccess.Text = "Full"
        Me.rbFullAccess.UseVisualStyleBackColor = True
        '
        'tbPassword
        '
        Me.tbPassword.Location = New System.Drawing.Point(179, 137)
        Me.tbPassword.Name = "tbPassword"
        Me.tbPassword.Size = New System.Drawing.Size(568, 26)
        Me.tbPassword.TabIndex = 5
        Me.tbPassword.UseSystemPasswordChar = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 140)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(87, 18)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Contraseña:"
        '
        'tbConfirmPassword
        '
        Me.tbConfirmPassword.Location = New System.Drawing.Point(179, 169)
        Me.tbConfirmPassword.Name = "tbConfirmPassword"
        Me.tbConfirmPassword.Size = New System.Drawing.Size(568, 26)
        Me.tbConfirmPassword.TabIndex = 6
        Me.tbConfirmPassword.UseSystemPasswordChar = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 172)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(86, 18)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Confirmarla:"
        '
        'btnConfirmChanges
        '
        Me.btnConfirmChanges.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnConfirmChanges.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnConfirmChanges.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConfirmChanges.Location = New System.Drawing.Point(630, 211)
        Me.btnConfirmChanges.Name = "btnConfirmChanges"
        Me.btnConfirmChanges.Size = New System.Drawing.Size(117, 31)
        Me.btnConfirmChanges.TabIndex = 7
        Me.btnConfirmChanges.Text = "Confirmar"
        Me.btnConfirmChanges.UseVisualStyleBackColor = False
        '
        'frmCreateUser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(759, 254)
        Me.Controls.Add(Me.btnConfirmChanges)
        Me.Controls.Add(Me.tbConfirmPassword)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.tbPassword)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.rbFullAccess)
        Me.Controls.Add(Me.rbBasicAccess)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbFname)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbSname)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbTitle)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCreateUser"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Crear nuevo usuario"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents tbTitle As TextBox
    Friend WithEvents tbSname As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents tbFname As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents rbBasicAccess As RadioButton
    Friend WithEvents rbFullAccess As RadioButton
    Friend WithEvents tbPassword As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents tbConfirmPassword As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents btnConfirmChanges As Button
End Class
