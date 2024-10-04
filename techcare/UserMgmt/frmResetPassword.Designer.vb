<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmResetPassword
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
        Me.btnConfirmChanges = New System.Windows.Forms.Button()
        Me.tbNewPassword = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbConfirmNewPassword = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnConfirmChanges
        '
        Me.btnConfirmChanges.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnConfirmChanges.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnConfirmChanges.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConfirmChanges.Location = New System.Drawing.Point(380, 148)
        Me.btnConfirmChanges.Name = "btnConfirmChanges"
        Me.btnConfirmChanges.Size = New System.Drawing.Size(117, 31)
        Me.btnConfirmChanges.TabIndex = 3
        Me.btnConfirmChanges.Text = "Confirmar"
        Me.btnConfirmChanges.UseVisualStyleBackColor = False
        '
        'tbNewPassword
        '
        Me.tbNewPassword.Location = New System.Drawing.Point(170, 12)
        Me.tbNewPassword.Name = "tbNewPassword"
        Me.tbNewPassword.Size = New System.Drawing.Size(327, 26)
        Me.tbNewPassword.TabIndex = 1
        Me.tbNewPassword.UseSystemPasswordChar = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(23, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(132, 18)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Nueva contraseña:"
        '
        'tbConfirmNewPassword
        '
        Me.tbConfirmNewPassword.Location = New System.Drawing.Point(170, 54)
        Me.tbConfirmNewPassword.Name = "tbConfirmNewPassword"
        Me.tbConfirmNewPassword.Size = New System.Drawing.Size(327, 26)
        Me.tbConfirmNewPassword.TabIndex = 2
        Me.tbConfirmNewPassword.UseSystemPasswordChar = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 57)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(153, 18)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Confirmar contraseña:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.Label2.Location = New System.Drawing.Point(13, 102)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(460, 26)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Tu nueva contraseña tiene los mismos requerimientos, 8 caracteres, 1 mayuscula, 1" & _
    " numero y " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "1 simbolo."
        '
        'frmResetPassword
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(509, 191)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbConfirmNewPassword)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbNewPassword)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnConfirmChanges)
        Me.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmResetPassword"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Resetear contraseña"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnConfirmChanges As Button
    Friend WithEvents tbNewPassword As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents tbConfirmNewPassword As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
End Class
