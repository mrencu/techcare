<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAppConfig
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAppConfig))
        Me.sidePanel = New System.Windows.Forms.Panel()
        Me.btnResetSoftware = New System.Windows.Forms.Button()
        Me.btnAboutTechcare = New System.Windows.Forms.Button()
        Me.btnDataMgmt = New System.Windows.Forms.Button()
        Me.btnDatabaseConfig = New System.Windows.Forms.Button()
        Me.databaseConfigPanel = New System.Windows.Forms.Panel()
        Me.lblWarningDbConfig = New System.Windows.Forms.Label()
        Me.btnSaveDbConfig = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbDbUsername = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbDbPassword = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnTestDbConnection = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbDbLocation = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbDbName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dataMgmtPanel = New System.Windows.Forms.Panel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnDbMgmtRestoreData = New System.Windows.Forms.Button()
        Me.btnDbMgmtBackupData = New System.Windows.Forms.Button()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.xamppLocationDialog = New System.Windows.Forms.FolderBrowserDialog()
        Me.xamppBackupToLocationDialog = New System.Windows.Forms.SaveFileDialog()
        Me.xamppRestoreFromLocationDialog = New System.Windows.Forms.OpenFileDialog()
        Me.aboutTechcarePanel = New System.Windows.Forms.Panel()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblVersionID = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.resetSoftwarePanel = New System.Windows.Forms.Panel()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.tbResetSoftwarePassword = New System.Windows.Forms.TextBox()
        Me.tbResetSoftwareUsername = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnResetSoftwareAuthentication = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.sidePanel.SuspendLayout()
        Me.databaseConfigPanel.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.dataMgmtPanel.SuspendLayout()
        Me.aboutTechcarePanel.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.resetSoftwarePanel.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'sidePanel
        '
        Me.sidePanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.sidePanel.Controls.Add(Me.btnResetSoftware)
        Me.sidePanel.Controls.Add(Me.btnAboutTechcare)
        Me.sidePanel.Controls.Add(Me.btnDataMgmt)
        Me.sidePanel.Controls.Add(Me.btnDatabaseConfig)
        Me.sidePanel.Dock = System.Windows.Forms.DockStyle.Left
        Me.sidePanel.Location = New System.Drawing.Point(0, 0)
        Me.sidePanel.Name = "sidePanel"
        Me.sidePanel.Size = New System.Drawing.Size(250, 314)
        Me.sidePanel.TabIndex = 1
        '
        'btnResetSoftware
        '
        Me.btnResetSoftware.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnResetSoftware.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnResetSoftware.Location = New System.Drawing.Point(3, 159)
        Me.btnResetSoftware.Name = "btnResetSoftware"
        Me.btnResetSoftware.Size = New System.Drawing.Size(242, 72)
        Me.btnResetSoftware.TabIndex = 3
        Me.btnResetSoftware.Text = "Formatear techcare"
        Me.btnResetSoftware.UseVisualStyleBackColor = True
        '
        'btnAboutTechcare
        '
        Me.btnAboutTechcare.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAboutTechcare.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAboutTechcare.Location = New System.Drawing.Point(3, 237)
        Me.btnAboutTechcare.Name = "btnAboutTechcare"
        Me.btnAboutTechcare.Size = New System.Drawing.Size(242, 72)
        Me.btnAboutTechcare.TabIndex = 2
        Me.btnAboutTechcare.Text = "Sobre Techcare"
        Me.btnAboutTechcare.UseVisualStyleBackColor = True
        '
        'btnDataMgmt
        '
        Me.btnDataMgmt.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDataMgmt.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDataMgmt.Location = New System.Drawing.Point(3, 81)
        Me.btnDataMgmt.Name = "btnDataMgmt"
        Me.btnDataMgmt.Size = New System.Drawing.Size(242, 72)
        Me.btnDataMgmt.TabIndex = 1
        Me.btnDataMgmt.Text = "Gestion de la data"
        Me.btnDataMgmt.UseVisualStyleBackColor = True
        '
        'btnDatabaseConfig
        '
        Me.btnDatabaseConfig.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDatabaseConfig.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDatabaseConfig.Location = New System.Drawing.Point(3, 3)
        Me.btnDatabaseConfig.Name = "btnDatabaseConfig"
        Me.btnDatabaseConfig.Size = New System.Drawing.Size(242, 72)
        Me.btnDatabaseConfig.TabIndex = 0
        Me.btnDatabaseConfig.Text = "Configuracion BD"
        Me.btnDatabaseConfig.UseVisualStyleBackColor = True
        '
        'databaseConfigPanel
        '
        Me.databaseConfigPanel.AutoSize = True
        Me.databaseConfigPanel.Controls.Add(Me.lblWarningDbConfig)
        Me.databaseConfigPanel.Controls.Add(Me.btnSaveDbConfig)
        Me.databaseConfigPanel.Controls.Add(Me.GroupBox2)
        Me.databaseConfigPanel.Controls.Add(Me.GroupBox1)
        Me.databaseConfigPanel.Controls.Add(Me.Label1)
        Me.databaseConfigPanel.Location = New System.Drawing.Point(754, 295)
        Me.databaseConfigPanel.Name = "databaseConfigPanel"
        Me.databaseConfigPanel.Size = New System.Drawing.Size(534, 314)
        Me.databaseConfigPanel.TabIndex = 2
        Me.databaseConfigPanel.Visible = False
        '
        'lblWarningDbConfig
        '
        Me.lblWarningDbConfig.AutoSize = True
        Me.lblWarningDbConfig.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarningDbConfig.Location = New System.Drawing.Point(6, 280)
        Me.lblWarningDbConfig.Name = "lblWarningDbConfig"
        Me.lblWarningDbConfig.Size = New System.Drawing.Size(360, 26)
        Me.lblWarningDbConfig.TabIndex = 8
        Me.lblWarningDbConfig.Text = "PRECAUCION! Si cambias estos ajustes, se reiniciará techcare. Fijate que" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & " no ten" & _
    "gas ningun registro que no hayas guardado"
        Me.lblWarningDbConfig.Visible = False
        '
        'btnSaveDbConfig
        '
        Me.btnSaveDbConfig.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnSaveDbConfig.Enabled = False
        Me.btnSaveDbConfig.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSaveDbConfig.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveDbConfig.Location = New System.Drawing.Point(426, 276)
        Me.btnSaveDbConfig.Name = "btnSaveDbConfig"
        Me.btnSaveDbConfig.Size = New System.Drawing.Size(105, 34)
        Me.btnSaveDbConfig.TabIndex = 7
        Me.btnSaveDbConfig.Text = "Guardar"
        Me.btnSaveDbConfig.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.tbDbUsername)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.tbDbPassword)
        Me.GroupBox2.ForeColor = System.Drawing.Color.White
        Me.GroupBox2.Location = New System.Drawing.Point(6, 184)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(525, 89)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Informacion de autenticacion"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 28)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(106, 14)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "MySQL Username:"
        '
        'tbDbUsername
        '
        Me.tbDbUsername.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbDbUsername.Location = New System.Drawing.Point(182, 25)
        Me.tbDbUsername.Name = "tbDbUsername"
        Me.tbDbUsername.Size = New System.Drawing.Size(337, 22)
        Me.tbDbUsername.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 60)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(113, 14)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "MySQL Contraseña:"
        '
        'tbDbPassword
        '
        Me.tbDbPassword.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbDbPassword.Location = New System.Drawing.Point(182, 57)
        Me.tbDbPassword.Name = "tbDbPassword"
        Me.tbDbPassword.Size = New System.Drawing.Size(337, 22)
        Me.tbDbPassword.TabIndex = 3
        Me.tbDbPassword.UseSystemPasswordChar = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnTestDbConnection)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.tbDbLocation)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.tbDbName)
        Me.GroupBox1.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(6, 65)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(525, 117)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Informacion de ubicacion"
        '
        'btnTestDbConnection
        '
        Me.btnTestDbConnection.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnTestDbConnection.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnTestDbConnection.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTestDbConnection.Location = New System.Drawing.Point(404, 85)
        Me.btnTestDbConnection.Name = "btnTestDbConnection"
        Me.btnTestDbConnection.Size = New System.Drawing.Size(115, 26)
        Me.btnTestDbConnection.TabIndex = 8
        Me.btnTestDbConnection.Text = "Testear conexion"
        Me.btnTestDbConnection.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 28)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(127, 14)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Donde esta tu server:"
        '
        'tbDbLocation
        '
        Me.tbDbLocation.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbDbLocation.Location = New System.Drawing.Point(182, 25)
        Me.tbDbLocation.Name = "tbDbLocation"
        Me.tbDbLocation.Size = New System.Drawing.Size(337, 22)
        Me.tbDbLocation.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 60)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(165, 14)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Nombre de la base de datos:"
        '
        'tbDbName
        '
        Me.tbDbName.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbDbName.Location = New System.Drawing.Point(182, 57)
        Me.tbDbName.Name = "tbDbName"
        Me.tbDbName.Size = New System.Drawing.Size(337, 22)
        Me.tbDbName.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(5, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(526, 33)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Configuracion de la base de datos"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dataMgmtPanel
        '
        Me.dataMgmtPanel.AutoSize = True
        Me.dataMgmtPanel.Controls.Add(Me.Label6)
        Me.dataMgmtPanel.Controls.Add(Me.btnDbMgmtRestoreData)
        Me.dataMgmtPanel.Controls.Add(Me.btnDbMgmtBackupData)
        Me.dataMgmtPanel.Controls.Add(Me.Label11)
        Me.dataMgmtPanel.Location = New System.Drawing.Point(769, 248)
        Me.dataMgmtPanel.Name = "dataMgmtPanel"
        Me.dataMgmtPanel.Size = New System.Drawing.Size(534, 314)
        Me.dataMgmtPanel.TabIndex = 3
        Me.dataMgmtPanel.Visible = False
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 186)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(525, 91)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "PRECAUCION" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Esta funcion tiene fallas tdv, tengo que hacer correciones, asi q" & _
    " nc, se aleja cositas" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnDbMgmtRestoreData
        '
        Me.btnDbMgmtRestoreData.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnDbMgmtRestoreData.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDbMgmtRestoreData.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDbMgmtRestoreData.Location = New System.Drawing.Point(271, 82)
        Me.btnDbMgmtRestoreData.Name = "btnDbMgmtRestoreData"
        Me.btnDbMgmtRestoreData.Size = New System.Drawing.Size(184, 72)
        Me.btnDbMgmtRestoreData.TabIndex = 3
        Me.btnDbMgmtRestoreData.Text = "Restaurar tu data (no funca todavia)"
        Me.btnDbMgmtRestoreData.UseVisualStyleBackColor = False
        '
        'btnDbMgmtBackupData
        '
        Me.btnDbMgmtBackupData.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnDbMgmtBackupData.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDbMgmtBackupData.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDbMgmtBackupData.Location = New System.Drawing.Point(81, 82)
        Me.btnDbMgmtBackupData.Name = "btnDbMgmtBackupData"
        Me.btnDbMgmtBackupData.Size = New System.Drawing.Size(184, 72)
        Me.btnDbMgmtBackupData.TabIndex = 2
        Me.btnDbMgmtBackupData.Text = "Respaldar tu data (no funca todavia)"
        Me.btnDbMgmtBackupData.UseVisualStyleBackColor = False
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(5, 19)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(526, 33)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Getion de la data :v"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'xamppLocationDialog
        '
        Me.xamppLocationDialog.RootFolder = System.Environment.SpecialFolder.MyComputer
        Me.xamppLocationDialog.ShowNewFolderButton = False
        '
        'xamppBackupToLocationDialog
        '
        Me.xamppBackupToLocationDialog.Filter = "SQL Dump File (*sql)|*.sql|All Files(*.*)|*.*"
        '
        'xamppRestoreFromLocationDialog
        '
        Me.xamppRestoreFromLocationDialog.Filter = "SQL Dump File (*sql)|*.sql|All Files(*.*)|*.*"
        '
        'aboutTechcarePanel
        '
        Me.aboutTechcarePanel.AutoSize = True
        Me.aboutTechcarePanel.Controls.Add(Me.Label16)
        Me.aboutTechcarePanel.Controls.Add(Me.Label13)
        Me.aboutTechcarePanel.Controls.Add(Me.Label12)
        Me.aboutTechcarePanel.Controls.Add(Me.Label10)
        Me.aboutTechcarePanel.Controls.Add(Me.Label9)
        Me.aboutTechcarePanel.Controls.Add(Me.lblVersionID)
        Me.aboutTechcarePanel.Controls.Add(Me.Label7)
        Me.aboutTechcarePanel.Controls.Add(Me.PictureBox1)
        Me.aboutTechcarePanel.Location = New System.Drawing.Point(310, 303)
        Me.aboutTechcarePanel.Name = "aboutTechcarePanel"
        Me.aboutTechcarePanel.Size = New System.Drawing.Size(534, 314)
        Me.aboutTechcarePanel.TabIndex = 4
        Me.aboutTechcarePanel.Visible = False
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.DarkCyan
        Me.Label16.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(215, 261)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(117, 32)
        Me.Label16.TabIndex = 7
        Me.Label16.Text = "MI GITHUB"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(3, 180)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(527, 19)
        Me.Label13.TabIndex = 6
        Me.Label13.Text = "Sobre este software"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(3, 195)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(527, 66)
        Me.Label12.TabIndex = 5
        Me.Label12.Text = resources.GetString("Label12.Text")
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(3, 154)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(527, 19)
        Me.Label10.TabIndex = 4
        Me.Label10.Text = "Proyecto de la materia Instalacion, Mantenimiento y Reparacion de Sistemas Comput" & _
    "acionales"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(3, 128)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(527, 19)
        Me.Label9.TabIndex = 3
        Me.Label9.Text = "Desarrollado por Orellana Matias"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblVersionID
        '
        Me.lblVersionID.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersionID.Location = New System.Drawing.Point(3, 101)
        Me.lblVersionID.Name = "lblVersionID"
        Me.lblVersionID.Size = New System.Drawing.Size(527, 19)
        Me.lblVersionID.TabIndex = 2
        Me.lblVersionID.Text = "Version: X.X.X.X"
        Me.lblVersionID.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 66)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(524, 21)
        Me.Label7.TabIndex = 1
        Me.Label7.Text = "Techcare: coso q hice para guardar reparaciones"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.techcare.My.Resources.Resources.croppedLogo
        Me.PictureBox1.Location = New System.Drawing.Point(6, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(524, 39)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'resetSoftwarePanel
        '
        Me.resetSoftwarePanel.AutoSize = True
        Me.resetSoftwarePanel.Controls.Add(Me.GroupBox3)
        Me.resetSoftwarePanel.Controls.Add(Me.Label8)
        Me.resetSoftwarePanel.Controls.Add(Me.Label18)
        Me.resetSoftwarePanel.Location = New System.Drawing.Point(613, 298)
        Me.resetSoftwarePanel.Name = "resetSoftwarePanel"
        Me.resetSoftwarePanel.Size = New System.Drawing.Size(545, 313)
        Me.resetSoftwarePanel.TabIndex = 5
        Me.resetSoftwarePanel.Visible = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.tbResetSoftwarePassword)
        Me.GroupBox3.Controls.Add(Me.tbResetSoftwareUsername)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Controls.Add(Me.btnResetSoftwareAuthentication)
        Me.GroupBox3.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.White
        Me.GroupBox3.Location = New System.Drawing.Point(12, 222)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(519, 88)
        Me.GroupBox3.TabIndex = 9
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Autenticarse como usuario administrador"
        '
        'tbResetSoftwarePassword
        '
        Me.tbResetSoftwarePassword.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbResetSoftwarePassword.Location = New System.Drawing.Point(84, 53)
        Me.tbResetSoftwarePassword.Name = "tbResetSoftwarePassword"
        Me.tbResetSoftwarePassword.Size = New System.Drawing.Size(311, 22)
        Me.tbResetSoftwarePassword.TabIndex = 11
        Me.tbResetSoftwarePassword.UseSystemPasswordChar = True
        '
        'tbResetSoftwareUsername
        '
        Me.tbResetSoftwareUsername.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbResetSoftwareUsername.Location = New System.Drawing.Point(84, 24)
        Me.tbResetSoftwareUsername.Name = "tbResetSoftwareUsername"
        Me.tbResetSoftwareUsername.Size = New System.Drawing.Size(311, 22)
        Me.tbResetSoftwareUsername.TabIndex = 10
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(6, 56)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(81, 14)
        Me.Label15.TabIndex = 9
        Me.Label15.Text = "Contraseña:"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 27)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(70, 14)
        Me.Label14.TabIndex = 8
        Me.Label14.Text = "Username:"
        '
        'btnResetSoftwareAuthentication
        '
        Me.btnResetSoftwareAuthentication.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnResetSoftwareAuthentication.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnResetSoftwareAuthentication.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnResetSoftwareAuthentication.Location = New System.Drawing.Point(408, 24)
        Me.btnResetSoftwareAuthentication.Name = "btnResetSoftwareAuthentication"
        Me.btnResetSoftwareAuthentication.Size = New System.Drawing.Size(105, 51)
        Me.btnResetSoftwareAuthentication.TabIndex = 7
        Me.btnResetSoftwareAuthentication.Text = "Resetear"
        Me.btnResetSoftwareAuthentication.UseVisualStyleBackColor = False
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(9, 65)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(522, 150)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = resources.GetString("Label8.Text")
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(5, 19)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(526, 33)
        Me.Label18.TabIndex = 0
        Me.Label18.Text = "Formatear software"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmAppConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(784, 314)
        Me.Controls.Add(Me.resetSoftwarePanel)
        Me.Controls.Add(Me.aboutTechcarePanel)
        Me.Controls.Add(Me.dataMgmtPanel)
        Me.Controls.Add(Me.databaseConfigPanel)
        Me.Controls.Add(Me.sidePanel)
        Me.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAppConfig"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Configuracion de la aplicacion"
        Me.TopMost = True
        Me.sidePanel.ResumeLayout(False)
        Me.databaseConfigPanel.ResumeLayout(False)
        Me.databaseConfigPanel.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.dataMgmtPanel.ResumeLayout(False)
        Me.aboutTechcarePanel.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.resetSoftwarePanel.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents sidePanel As Panel
    Friend WithEvents btnAboutTechcare As Button
    Friend WithEvents btnDataMgmt As Button
    Friend WithEvents btnDatabaseConfig As Button
    Friend WithEvents databaseConfigPanel As Panel
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents tbDbLocation As TextBox
    Friend WithEvents tbDbName As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label4 As Label
    Friend WithEvents tbDbUsername As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents tbDbPassword As TextBox
    Friend WithEvents btnSaveDbConfig As Button
    Friend WithEvents lblWarningDbConfig As Label
    Friend WithEvents btnTestDbConnection As Button
    Friend WithEvents dataMgmtPanel As Panel
    Friend WithEvents Label11 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents btnDbMgmtRestoreData As Button
    Friend WithEvents btnDbMgmtBackupData As Button
    Friend WithEvents xamppLocationDialog As FolderBrowserDialog
    Friend WithEvents xamppBackupToLocationDialog As SaveFileDialog
    Friend WithEvents xamppRestoreFromLocationDialog As OpenFileDialog
    Friend WithEvents aboutTechcarePanel As Panel
    Friend WithEvents Label13 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents lblVersionID As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents btnResetSoftware As Button
    Friend WithEvents resetSoftwarePanel As Panel
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents tbResetSoftwarePassword As TextBox
    Friend WithEvents tbResetSoftwareUsername As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents btnResetSoftwareAuthentication As Button
    Friend WithEvents Label8 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
End Class
