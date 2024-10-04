<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMainWindow
    Inherits System.Windows.Forms.Form

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

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMainWindow))
        Me.sidePanel = New System.Windows.Forms.Panel()
        Me.btnApplicationSettings = New System.Windows.Forms.Button()
        Me.btnSignOut = New System.Windows.Forms.Button()
        Me.btnUserMgmt = New System.Windows.Forms.Button()
        Me.btnRepairMgmt = New System.Windows.Forms.Button()
        Me.btnNewRepair = New System.Windows.Forms.Button()
        Me.statusStrip = New System.Windows.Forms.Panel()
        Me.lblCurrentUser = New System.Windows.Forms.Label()
        Me.lblEmpID = New System.Windows.Forms.Label()
        Me.lblVersionID = New System.Windows.Forms.Label()
        Me.sidePanel.SuspendLayout()
        Me.statusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'sidePanel
        '
        Me.sidePanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.sidePanel.Controls.Add(Me.btnApplicationSettings)
        Me.sidePanel.Controls.Add(Me.btnSignOut)
        Me.sidePanel.Controls.Add(Me.btnUserMgmt)
        Me.sidePanel.Controls.Add(Me.btnRepairMgmt)
        Me.sidePanel.Controls.Add(Me.btnNewRepair)
        Me.sidePanel.Dock = System.Windows.Forms.DockStyle.Left
        Me.sidePanel.Location = New System.Drawing.Point(0, 0)
        Me.sidePanel.Name = "sidePanel"
        Me.sidePanel.Size = New System.Drawing.Size(250, 706)
        Me.sidePanel.TabIndex = 0
        Me.sidePanel.Visible = False
        '
        'btnApplicationSettings
        '
        Me.btnApplicationSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnApplicationSettings.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnApplicationSettings.Location = New System.Drawing.Point(3, 237)
        Me.btnApplicationSettings.Name = "btnApplicationSettings"
        Me.btnApplicationSettings.Size = New System.Drawing.Size(242, 72)
        Me.btnApplicationSettings.TabIndex = 5
        Me.btnApplicationSettings.Text = "Configuracion"
        Me.btnApplicationSettings.UseVisualStyleBackColor = True
        '
        'btnSignOut
        '
        Me.btnSignOut.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.btnSignOut.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSignOut.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSignOut.Location = New System.Drawing.Point(0, 632)
        Me.btnSignOut.Name = "btnSignOut"
        Me.btnSignOut.Size = New System.Drawing.Size(248, 72)
        Me.btnSignOut.TabIndex = 4
        Me.btnSignOut.Text = "Cerrar sesion"
        Me.btnSignOut.UseVisualStyleBackColor = True
        '
        'btnUserMgmt
        '
        Me.btnUserMgmt.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUserMgmt.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUserMgmt.Location = New System.Drawing.Point(3, 159)
        Me.btnUserMgmt.Name = "btnUserMgmt"
        Me.btnUserMgmt.Size = New System.Drawing.Size(242, 72)
        Me.btnUserMgmt.TabIndex = 2
        Me.btnUserMgmt.Text = "Gestionar usuarios"
        Me.btnUserMgmt.UseVisualStyleBackColor = True
        '
        'btnRepairMgmt
        '
        Me.btnRepairMgmt.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRepairMgmt.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRepairMgmt.Location = New System.Drawing.Point(3, 81)
        Me.btnRepairMgmt.Name = "btnRepairMgmt"
        Me.btnRepairMgmt.Size = New System.Drawing.Size(242, 72)
        Me.btnRepairMgmt.TabIndex = 1
        Me.btnRepairMgmt.Text = "Gestionar reparaciones"
        Me.btnRepairMgmt.UseVisualStyleBackColor = True
        '
        'btnNewRepair
        '
        Me.btnNewRepair.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNewRepair.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNewRepair.Location = New System.Drawing.Point(3, 3)
        Me.btnNewRepair.Name = "btnNewRepair"
        Me.btnNewRepair.Size = New System.Drawing.Size(242, 72)
        Me.btnNewRepair.TabIndex = 0
        Me.btnNewRepair.Text = "Crear reparacion"
        Me.btnNewRepair.UseVisualStyleBackColor = True
        '
        'statusStrip
        '
        Me.statusStrip.Controls.Add(Me.lblCurrentUser)
        Me.statusStrip.Controls.Add(Me.lblEmpID)
        Me.statusStrip.Controls.Add(Me.lblVersionID)
        Me.statusStrip.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.statusStrip.Location = New System.Drawing.Point(0, 706)
        Me.statusStrip.Name = "statusStrip"
        Me.statusStrip.Size = New System.Drawing.Size(1008, 23)
        Me.statusStrip.TabIndex = 2
        '
        'lblCurrentUser
        '
        Me.lblCurrentUser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCurrentUser.Location = New System.Drawing.Point(144, 0)
        Me.lblCurrentUser.Name = "lblCurrentUser"
        Me.lblCurrentUser.Size = New System.Drawing.Size(720, 23)
        Me.lblCurrentUser.TabIndex = 2
        Me.lblCurrentUser.Text = "<<Current User>>"
        Me.lblCurrentUser.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblCurrentUser.Visible = False
        '
        'lblEmpID
        '
        Me.lblEmpID.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblEmpID.Location = New System.Drawing.Point(0, 0)
        Me.lblEmpID.Name = "lblEmpID"
        Me.lblEmpID.Size = New System.Drawing.Size(144, 23)
        Me.lblEmpID.TabIndex = 3
        Me.lblEmpID.Text = "<<Employee ID>>"
        Me.lblEmpID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblEmpID.Visible = False
        '
        'lblVersionID
        '
        Me.lblVersionID.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblVersionID.Location = New System.Drawing.Point(864, 0)
        Me.lblVersionID.Name = "lblVersionID"
        Me.lblVersionID.Size = New System.Drawing.Size(144, 23)
        Me.lblVersionID.TabIndex = 1
        Me.lblVersionID.Text = "Version x.x.x"
        Me.lblVersionID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmMainWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.BackgroundImage = Global.techcare.My.Resources.Resources.croppedLogo
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(1008, 729)
        Me.Controls.Add(Me.sidePanel)
        Me.Controls.Add(Me.statusStrip)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(1024, 766)
        Me.Name = "frmMainWindow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "techcare"
        Me.sidePanel.ResumeLayout(False)
        Me.statusStrip.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents sidePanel As Panel
    Friend WithEvents statusStrip As Panel
    Friend WithEvents lblCurrentUser As Label
    Friend WithEvents lblEmpID As Label
    Friend WithEvents lblVersionID As Label
    Friend WithEvents btnSignOut As Button
    Friend WithEvents btnUserMgmt As Button
    Friend WithEvents btnRepairMgmt As Button
    Friend WithEvents btnNewRepair As Button
    Friend WithEvents btnApplicationSettings As Button
End Class
