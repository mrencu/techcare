<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdateQuote
    Inherits System.Windows.Forms.Form

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

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnConfirmChanges = New System.Windows.Forms.Button()
        Me.tbNewRepairQuote = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnConfirmChanges
        '
        Me.btnConfirmChanges.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnConfirmChanges.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnConfirmChanges.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConfirmChanges.Location = New System.Drawing.Point(355, 58)
        Me.btnConfirmChanges.Name = "btnConfirmChanges"
        Me.btnConfirmChanges.Size = New System.Drawing.Size(117, 31)
        Me.btnConfirmChanges.TabIndex = 8
        Me.btnConfirmChanges.Text = "Confirmar"
        Me.btnConfirmChanges.UseVisualStyleBackColor = False
        '
        'tbNewRepairQuote
        '
        Me.tbNewRepairQuote.Location = New System.Drawing.Point(155, 12)
        Me.tbNewRepairQuote.Name = "tbNewRepairQuote"
        Me.tbNewRepairQuote.Size = New System.Drawing.Size(317, 26)
        Me.tbNewRepairQuote.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 18)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Nuevo precio:"
        '
        'frmUpdateQuote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(484, 101)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbNewRepairQuote)
        Me.Controls.Add(Me.btnConfirmChanges)
        Me.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUpdateQuote"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Actualizar precio de reparacion"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnConfirmChanges As Button
    Friend WithEvents tbNewRepairQuote As TextBox
    Friend WithEvents Label1 As Label
End Class
