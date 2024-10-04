<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdateRepairStatus
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
        Me.rbBookedIn = New System.Windows.Forms.RadioButton()
        Me.rbServiceInProgress = New System.Windows.Forms.RadioButton()
        Me.rbTransferred = New System.Windows.Forms.RadioButton()
        Me.rbServiceCompleted = New System.Windows.Forms.RadioButton()
        Me.rbAssetRemoved = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout()
        '
        'btnConfirmChanges
        '
        Me.btnConfirmChanges.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnConfirmChanges.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnConfirmChanges.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConfirmChanges.Location = New System.Drawing.Point(155, 218)
        Me.btnConfirmChanges.Name = "btnConfirmChanges"
        Me.btnConfirmChanges.Size = New System.Drawing.Size(117, 31)
        Me.btnConfirmChanges.TabIndex = 7
        Me.btnConfirmChanges.Text = "Confirmar"
        Me.btnConfirmChanges.UseVisualStyleBackColor = False
        '
        'rbBookedIn
        '
        Me.rbBookedIn.AutoSize = True
        Me.rbBookedIn.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbBookedIn.Location = New System.Drawing.Point(12, 12)
        Me.rbBookedIn.Name = "rbBookedIn"
        Me.rbBookedIn.Size = New System.Drawing.Size(92, 22)
        Me.rbBookedIn.TabIndex = 9
        Me.rbBookedIn.TabStop = True
        Me.rbBookedIn.Text = "Ingresado"
        Me.rbBookedIn.UseVisualStyleBackColor = True
        '
        'rbServiceInProgress
        '
        Me.rbServiceInProgress.AutoSize = True
        Me.rbServiceInProgress.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbServiceInProgress.Location = New System.Drawing.Point(12, 45)
        Me.rbServiceInProgress.Name = "rbServiceInProgress"
        Me.rbServiceInProgress.Size = New System.Drawing.Size(157, 22)
        Me.rbServiceInProgress.TabIndex = 10
        Me.rbServiceInProgress.TabStop = True
        Me.rbServiceInProgress.Text = "Servicio en progreso"
        Me.rbServiceInProgress.UseVisualStyleBackColor = True
        '
        'rbTransferred
        '
        Me.rbTransferred.AutoSize = True
        Me.rbTransferred.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbTransferred.Location = New System.Drawing.Point(12, 78)
        Me.rbTransferred.Name = "rbTransferred"
        Me.rbTransferred.Size = New System.Drawing.Size(171, 22)
        Me.rbTransferred.TabIndex = 11
        Me.rbTransferred.TabStop = True
        Me.rbTransferred.Text = "Transferido a terceros"
        Me.rbTransferred.UseVisualStyleBackColor = True
        '
        'rbServiceCompleted
        '
        Me.rbServiceCompleted.AutoSize = True
        Me.rbServiceCompleted.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbServiceCompleted.Location = New System.Drawing.Point(12, 111)
        Me.rbServiceCompleted.Name = "rbServiceCompleted"
        Me.rbServiceCompleted.Size = New System.Drawing.Size(154, 22)
        Me.rbServiceCompleted.TabIndex = 12
        Me.rbServiceCompleted.TabStop = True
        Me.rbServiceCompleted.Text = "Servicio completado"
        Me.rbServiceCompleted.UseVisualStyleBackColor = True
        '
        'rbAssetRemoved
        '
        Me.rbAssetRemoved.AutoSize = True
        Me.rbAssetRemoved.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbAssetRemoved.Location = New System.Drawing.Point(12, 145)
        Me.rbAssetRemoved.Name = "rbAssetRemoved"
        Me.rbAssetRemoved.Size = New System.Drawing.Size(85, 22)
        Me.rbAssetRemoved.TabIndex = 13
        Me.rbAssetRemoved.TabStop = True
        Me.rbAssetRemoved.Text = "Eliminado"
        Me.rbAssetRemoved.UseVisualStyleBackColor = True
        '
        'frmUpdateRepairStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Controls.Add(Me.rbAssetRemoved)
        Me.Controls.Add(Me.rbServiceCompleted)
        Me.Controls.Add(Me.rbTransferred)
        Me.Controls.Add(Me.rbServiceInProgress)
        Me.Controls.Add(Me.rbBookedIn)
        Me.Controls.Add(Me.btnConfirmChanges)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUpdateRepairStatus"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Actualizar estado de la reparacion"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnConfirmChanges As Button
    Friend WithEvents rbBookedIn As RadioButton
    Friend WithEvents rbServiceInProgress As RadioButton
    Friend WithEvents rbTransferred As RadioButton
    Friend WithEvents rbServiceCompleted As RadioButton
    Friend WithEvents rbAssetRemoved As RadioButton
End Class
