<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAddRepairRemark
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
        Me.btnAddRemark = New System.Windows.Forms.Button()
        Me.tbRemark = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnAddRemark
        '
        Me.btnAddRemark.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnAddRemark.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAddRemark.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddRemark.Location = New System.Drawing.Point(355, 218)
        Me.btnAddRemark.Name = "btnAddRemark"
        Me.btnAddRemark.Size = New System.Drawing.Size(117, 31)
        Me.btnAddRemark.TabIndex = 9
        Me.btnAddRemark.Text = "Añadir"
        Me.btnAddRemark.UseVisualStyleBackColor = False
        '
        'tbRemark
        '
        Me.tbRemark.Location = New System.Drawing.Point(13, 13)
        Me.tbRemark.Multiline = True
        Me.tbRemark.Name = "tbRemark"
        Me.tbRemark.Size = New System.Drawing.Size(459, 199)
        Me.tbRemark.TabIndex = 10
        '
        'frmAddRepairRemark
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(484, 261)
        Me.Controls.Add(Me.tbRemark)
        Me.Controls.Add(Me.btnAddRemark)
        Me.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddRepairRemark"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Añadir comentario a la reparacion"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnAddRemark As Button
    Friend WithEvents tbRemark As TextBox
End Class
