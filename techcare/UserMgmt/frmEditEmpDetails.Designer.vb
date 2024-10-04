<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditEmpDetails
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbTitle = New System.Windows.Forms.TextBox()
        Me.tbSurname = New System.Windows.Forms.TextBox()
        Me.tbForename = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.rbBasicAccess = New System.Windows.Forms.RadioButton()
        Me.rbFullAccess = New System.Windows.Forms.RadioButton()
        Me.lblEditSurnameWarning = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnConfirmChanges
        '
        Me.btnConfirmChanges.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnConfirmChanges.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnConfirmChanges.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConfirmChanges.Location = New System.Drawing.Point(380, 125)
        Me.btnConfirmChanges.Name = "btnConfirmChanges"
        Me.btnConfirmChanges.Size = New System.Drawing.Size(117, 31)
        Me.btnConfirmChanges.TabIndex = 6
        Me.btnConfirmChanges.Text = "Confirmar"
        Me.btnConfirmChanges.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 18)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Pronombre:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(208, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 18)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Nombre:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 49)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 18)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Apellido:"
        '
        'tbTitle
        '
        Me.tbTitle.Location = New System.Drawing.Point(103, 10)
        Me.tbTitle.Name = "tbTitle"
        Me.tbTitle.Size = New System.Drawing.Size(85, 26)
        Me.tbTitle.TabIndex = 1
        '
        'tbSurname
        '
        Me.tbSurname.Location = New System.Drawing.Point(90, 46)
        Me.tbSurname.Name = "tbSurname"
        Me.tbSurname.Size = New System.Drawing.Size(407, 26)
        Me.tbSurname.TabIndex = 3
        '
        'tbForename
        '
        Me.tbForename.Location = New System.Drawing.Point(293, 10)
        Me.tbForename.Name = "tbForename"
        Me.tbForename.Size = New System.Drawing.Size(204, 26)
        Me.tbForename.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 87)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(114, 18)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Nivel de acceso:"
        '
        'rbBasicAccess
        '
        Me.rbBasicAccess.AutoSize = True
        Me.rbBasicAccess.Location = New System.Drawing.Point(175, 85)
        Me.rbBasicAccess.Name = "rbBasicAccess"
        Me.rbBasicAccess.Size = New System.Drawing.Size(67, 22)
        Me.rbBasicAccess.TabIndex = 4
        Me.rbBasicAccess.TabStop = True
        Me.rbBasicAccess.Text = "Basico"
        Me.rbBasicAccess.UseVisualStyleBackColor = True
        '
        'rbFullAccess
        '
        Me.rbFullAccess.AutoSize = True
        Me.rbFullAccess.Location = New System.Drawing.Point(261, 85)
        Me.rbFullAccess.Name = "rbFullAccess"
        Me.rbFullAccess.Size = New System.Drawing.Size(46, 22)
        Me.rbFullAccess.TabIndex = 5
        Me.rbFullAccess.TabStop = True
        Me.rbFullAccess.Text = "Full"
        Me.rbFullAccess.UseVisualStyleBackColor = True
        '
        'lblEditSurnameWarning
        '
        Me.lblEditSurnameWarning.AutoSize = True
        Me.lblEditSurnameWarning.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEditSurnameWarning.Location = New System.Drawing.Point(13, 135)
        Me.lblEditSurnameWarning.Name = "lblEditSurnameWarning"
        Me.lblEditSurnameWarning.Size = New System.Drawing.Size(289, 13)
        Me.lblEditSurnameWarning.TabIndex = 17
        Me.lblEditSurnameWarning.Text = "Cambiando el nombre o el apellido, no cambia el username."
        Me.lblEditSurnameWarning.Visible = False
        '
        'frmEditEmpDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(509, 169)
        Me.Controls.Add(Me.lblEditSurnameWarning)
        Me.Controls.Add(Me.rbFullAccess)
        Me.Controls.Add(Me.rbBasicAccess)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbForename)
        Me.Controls.Add(Me.tbSurname)
        Me.Controls.Add(Me.tbTitle)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnConfirmChanges)
        Me.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEditEmpDetails"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Editando datos del empleado: XXXXXXXX"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnConfirmChanges As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents tbTitle As TextBox
    Friend WithEvents tbSurname As TextBox
    Friend WithEvents tbForename As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents rbBasicAccess As RadioButton
    Friend WithEvents rbFullAccess As RadioButton
    Friend WithEvents lblEditSurnameWarning As Label
End Class
