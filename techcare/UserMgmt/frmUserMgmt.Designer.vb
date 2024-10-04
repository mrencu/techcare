<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUserMgmt
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.sidePanel = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnCreateNewEmp = New System.Windows.Forms.Button()
        Me.lblUserPrompt = New System.Windows.Forms.Label()
        Me.btnDeleteEmp = New System.Windows.Forms.Button()
        Me.btnResetEmpPwd = New System.Windows.Forms.Button()
        Me.btnEditEmployeeDetails = New System.Windows.Forms.Button()
        Me.lblUserAccessLvl = New System.Windows.Forms.Label()
        Me.lblEmpID = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblEmpName = New System.Windows.Forms.Label()
        Me.dgvCurrentUserList = New System.Windows.Forms.DataGridView()
        Me.colEmpID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colUserID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colEmpName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colUserAccessLevel = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.sidePanel.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.dgvCurrentUserList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'sidePanel
        '
        Me.sidePanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.sidePanel.Controls.Add(Me.Panel1)
        Me.sidePanel.Controls.Add(Me.lblUserPrompt)
        Me.sidePanel.Controls.Add(Me.btnDeleteEmp)
        Me.sidePanel.Controls.Add(Me.btnResetEmpPwd)
        Me.sidePanel.Controls.Add(Me.btnEditEmployeeDetails)
        Me.sidePanel.Controls.Add(Me.lblUserAccessLvl)
        Me.sidePanel.Controls.Add(Me.lblEmpID)
        Me.sidePanel.Controls.Add(Me.Label2)
        Me.sidePanel.Controls.Add(Me.Label1)
        Me.sidePanel.Controls.Add(Me.lblEmpName)
        Me.sidePanel.Dock = System.Windows.Forms.DockStyle.Left
        Me.sidePanel.Location = New System.Drawing.Point(0, 0)
        Me.sidePanel.Name = "sidePanel"
        Me.sidePanel.Size = New System.Drawing.Size(350, 729)
        Me.sidePanel.TabIndex = 2
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnCreateNewEmp)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 678)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(348, 49)
        Me.Panel1.TabIndex = 22
        '
        'btnCreateNewEmp
        '
        Me.btnCreateNewEmp.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnCreateNewEmp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCreateNewEmp.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreateNewEmp.Location = New System.Drawing.Point(14, 3)
        Me.btnCreateNewEmp.Name = "btnCreateNewEmp"
        Me.btnCreateNewEmp.Size = New System.Drawing.Size(318, 31)
        Me.btnCreateNewEmp.TabIndex = 20
        Me.btnCreateNewEmp.Text = "Crear usuario"
        Me.btnCreateNewEmp.UseVisualStyleBackColor = False
        '
        'lblUserPrompt
        '
        Me.lblUserPrompt.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserPrompt.Location = New System.Drawing.Point(53, 308)
        Me.lblUserPrompt.Name = "lblUserPrompt"
        Me.lblUserPrompt.Size = New System.Drawing.Size(236, 116)
        Me.lblUserPrompt.TabIndex = 21
        Me.lblUserPrompt.Text = "Selecciona un empleado para ver sus detalles a la derecha."
        Me.lblUserPrompt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnDeleteEmp
        '
        Me.btnDeleteEmp.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnDeleteEmp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDeleteEmp.Location = New System.Drawing.Point(14, 246)
        Me.btnDeleteEmp.Name = "btnDeleteEmp"
        Me.btnDeleteEmp.Size = New System.Drawing.Size(318, 31)
        Me.btnDeleteEmp.TabIndex = 19
        Me.btnDeleteEmp.Text = "Borrar usuario (permanente)"
        Me.btnDeleteEmp.UseVisualStyleBackColor = False
        '
        'btnResetEmpPwd
        '
        Me.btnResetEmpPwd.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnResetEmpPwd.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnResetEmpPwd.Location = New System.Drawing.Point(14, 209)
        Me.btnResetEmpPwd.Name = "btnResetEmpPwd"
        Me.btnResetEmpPwd.Size = New System.Drawing.Size(318, 31)
        Me.btnResetEmpPwd.TabIndex = 18
        Me.btnResetEmpPwd.Text = "Resetear contraseña"
        Me.btnResetEmpPwd.UseVisualStyleBackColor = False
        '
        'btnEditEmployeeDetails
        '
        Me.btnEditEmployeeDetails.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnEditEmployeeDetails.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnEditEmployeeDetails.Location = New System.Drawing.Point(14, 172)
        Me.btnEditEmployeeDetails.Name = "btnEditEmployeeDetails"
        Me.btnEditEmployeeDetails.Size = New System.Drawing.Size(318, 31)
        Me.btnEditEmployeeDetails.TabIndex = 17
        Me.btnEditEmployeeDetails.Text = "Editar datos"
        Me.btnEditEmployeeDetails.UseVisualStyleBackColor = False
        '
        'lblUserAccessLvl
        '
        Me.lblUserAccessLvl.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserAccessLvl.Location = New System.Drawing.Point(187, 117)
        Me.lblUserAccessLvl.Name = "lblUserAccessLvl"
        Me.lblUserAccessLvl.Size = New System.Drawing.Size(145, 18)
        Me.lblUserAccessLvl.TabIndex = 4
        Me.lblUserAccessLvl.Text = "<ACCESS_LEVEL>"
        Me.lblUserAccessLvl.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblEmpID
        '
        Me.lblEmpID.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmpID.Location = New System.Drawing.Point(187, 79)
        Me.lblEmpID.Name = "lblEmpID"
        Me.lblEmpID.Size = New System.Drawing.Size(145, 18)
        Me.lblEmpID.TabIndex = 3
        Me.lblEmpID.Text = "XXXXXXXX"
        Me.lblEmpID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(11, 117)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(130, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Nivel de acceso:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(11, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 18)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "ID Empleado:"
        '
        'lblEmpName
        '
        Me.lblEmpName.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmpName.Location = New System.Drawing.Point(12, 8)
        Me.lblEmpName.Name = "lblEmpName"
        Me.lblEmpName.Size = New System.Drawing.Size(320, 58)
        Me.lblEmpName.TabIndex = 0
        Me.lblEmpName.Text = "Mr Nombre Apellido"
        Me.lblEmpName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dgvCurrentUserList
        '
        Me.dgvCurrentUserList.AllowUserToAddRows = False
        Me.dgvCurrentUserList.AllowUserToDeleteRows = False
        Me.dgvCurrentUserList.AllowUserToResizeColumns = False
        Me.dgvCurrentUserList.AllowUserToResizeRows = False
        Me.dgvCurrentUserList.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.dgvCurrentUserList.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvCurrentUserList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvCurrentUserList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCurrentUserList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colEmpID, Me.colUserID, Me.colEmpName, Me.colUserAccessLevel})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvCurrentUserList.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvCurrentUserList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCurrentUserList.GridColor = System.Drawing.Color.White
        Me.dgvCurrentUserList.Location = New System.Drawing.Point(350, 0)
        Me.dgvCurrentUserList.Name = "dgvCurrentUserList"
        Me.dgvCurrentUserList.ReadOnly = True
        Me.dgvCurrentUserList.RowHeadersVisible = False
        Me.dgvCurrentUserList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvCurrentUserList.Size = New System.Drawing.Size(658, 729)
        Me.dgvCurrentUserList.TabIndex = 5
        '
        'colEmpID
        '
        Me.colEmpID.HeaderText = "ID Empleado"
        Me.colEmpID.Name = "colEmpID"
        Me.colEmpID.ReadOnly = True
        Me.colEmpID.Width = 150
        '
        'colUserID
        '
        Me.colUserID.HeaderText = "Username"
        Me.colUserID.Name = "colUserID"
        Me.colUserID.ReadOnly = True
        Me.colUserID.Width = 105
        '
        'colEmpName
        '
        Me.colEmpName.HeaderText = "Nombre empleado"
        Me.colEmpName.Name = "colEmpName"
        Me.colEmpName.ReadOnly = True
        Me.colEmpName.Width = 200
        '
        'colUserAccessLevel
        '
        Me.colUserAccessLevel.HeaderText = "Nivel de acceso"
        Me.colUserAccessLevel.Name = "colUserAccessLevel"
        Me.colUserAccessLevel.ReadOnly = True
        Me.colUserAccessLevel.Width = 200
        '
        'frmUserMgmt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1008, 729)
        Me.Controls.Add(Me.dgvCurrentUserList)
        Me.Controls.Add(Me.sidePanel)
        Me.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "frmUserMgmt"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Gestion de usuarios"
        Me.sidePanel.ResumeLayout(False)
        Me.sidePanel.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgvCurrentUserList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents sidePanel As Panel
    Friend WithEvents dgvCurrentUserList As DataGridView
    Friend WithEvents lblEmpName As Label
    Friend WithEvents lblUserAccessLvl As Label
    Friend WithEvents lblEmpID As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents btnDeleteEmp As Button
    Friend WithEvents btnResetEmpPwd As Button
    Friend WithEvents btnEditEmployeeDetails As Button
    Friend WithEvents btnCreateNewEmp As Button
    Friend WithEvents lblUserPrompt As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents colEmpID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colUserID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEmpName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colUserAccessLevel As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
