<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCreateNewRepair
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbCustomerForename = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbCustomerSurname = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbCustomerAddress = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbCustomerPostcode = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbCustomerCity = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbCustomerMobilePhone = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.tbCustomerHomePhone = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tbCustomerEmail = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.tbAssetModel = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.tbAssetMake = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.tbAssetSerialNumber = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.tbFaultType = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.tbFaultDescription = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.tbEstimateCost = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.tbRepairTnC = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.btnCreateRepair = New System.Windows.Forms.Button()
        Me.cbCustomerTitle = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Prefijo:"
        '
        'tbCustomerForename
        '
        Me.tbCustomerForename.Location = New System.Drawing.Point(409, 12)
        Me.tbCustomerForename.Name = "tbCustomerForename"
        Me.tbCustomerForename.Size = New System.Drawing.Size(263, 26)
        Me.tbCustomerForename.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(324, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Nombre:"
        '
        'tbCustomerSurname
        '
        Me.tbCustomerSurname.Location = New System.Drawing.Point(121, 46)
        Me.tbCustomerSurname.Name = "tbCustomerSurname"
        Me.tbCustomerSurname.Size = New System.Drawing.Size(551, 26)
        Me.tbCustomerSurname.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 49)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 18)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Apellido:"
        '
        'tbCustomerAddress
        '
        Me.tbCustomerAddress.Location = New System.Drawing.Point(121, 78)
        Me.tbCustomerAddress.Name = "tbCustomerAddress"
        Me.tbCustomerAddress.Size = New System.Drawing.Size(551, 26)
        Me.tbCustomerAddress.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 81)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 18)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Direccion:"
        '
        'tbCustomerPostcode
        '
        Me.tbCustomerPostcode.Location = New System.Drawing.Point(409, 110)
        Me.tbCustomerPostcode.MaxLength = 8
        Me.tbCustomerPostcode.Name = "tbCustomerPostcode"
        Me.tbCustomerPostcode.Size = New System.Drawing.Size(263, 26)
        Me.tbCustomerPostcode.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(305, 113)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(99, 18)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Codigo postal:"
        '
        'tbCustomerCity
        '
        Me.tbCustomerCity.Location = New System.Drawing.Point(121, 110)
        Me.tbCustomerCity.Name = "tbCustomerCity"
        Me.tbCustomerCity.Size = New System.Drawing.Size(170, 26)
        Me.tbCustomerCity.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 113)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(70, 18)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Municipio:"
        '
        'tbCustomerMobilePhone
        '
        Me.tbCustomerMobilePhone.Location = New System.Drawing.Point(409, 196)
        Me.tbCustomerMobilePhone.MaxLength = 20
        Me.tbCustomerMobilePhone.Name = "tbCustomerMobilePhone"
        Me.tbCustomerMobilePhone.Size = New System.Drawing.Size(263, 26)
        Me.tbCustomerMobilePhone.TabIndex = 9
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(334, 199)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(55, 18)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Celular:"
        '
        'tbCustomerHomePhone
        '
        Me.tbCustomerHomePhone.Location = New System.Drawing.Point(121, 196)
        Me.tbCustomerHomePhone.MaxLength = 20
        Me.tbCustomerHomePhone.Name = "tbCustomerHomePhone"
        Me.tbCustomerHomePhone.Size = New System.Drawing.Size(170, 26)
        Me.tbCustomerHomePhone.TabIndex = 8
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(12, 199)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(70, 18)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Telefono:"
        '
        'tbCustomerEmail
        '
        Me.tbCustomerEmail.Location = New System.Drawing.Point(121, 164)
        Me.tbCustomerEmail.Name = "tbCustomerEmail"
        Me.tbCustomerEmail.Size = New System.Drawing.Size(551, 26)
        Me.tbCustomerEmail.TabIndex = 7
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(12, 167)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 18)
        Me.Label9.TabIndex = 12
        Me.Label9.Text = "Correo:"
        '
        'tbAssetModel
        '
        Me.tbAssetModel.Location = New System.Drawing.Point(409, 253)
        Me.tbAssetModel.Name = "tbAssetModel"
        Me.tbAssetModel.Size = New System.Drawing.Size(263, 26)
        Me.tbAssetModel.TabIndex = 11
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(330, 256)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(59, 18)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "Modelo:"
        '
        'tbAssetMake
        '
        Me.tbAssetMake.Location = New System.Drawing.Point(121, 253)
        Me.tbAssetMake.Name = "tbAssetMake"
        Me.tbAssetMake.Size = New System.Drawing.Size(170, 26)
        Me.tbAssetMake.TabIndex = 10
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(12, 256)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(53, 18)
        Me.Label11.TabIndex = 18
        Me.Label11.Text = "Marca:"
        '
        'tbAssetSerialNumber
        '
        Me.tbAssetSerialNumber.Location = New System.Drawing.Point(121, 285)
        Me.tbAssetSerialNumber.Name = "tbAssetSerialNumber"
        Me.tbAssetSerialNumber.Size = New System.Drawing.Size(551, 26)
        Me.tbAssetSerialNumber.TabIndex = 12
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(12, 288)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(86, 18)
        Me.Label12.TabIndex = 22
        Me.Label12.Text = "N° de serie:"
        '
        'tbFaultType
        '
        Me.tbFaultType.Location = New System.Drawing.Point(121, 341)
        Me.tbFaultType.Name = "tbFaultType"
        Me.tbFaultType.Size = New System.Drawing.Size(551, 26)
        Me.tbFaultType.TabIndex = 13
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(12, 344)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(92, 18)
        Me.Label13.TabIndex = 24
        Me.Label13.Text = "Tipo de falla:"
        '
        'tbFaultDescription
        '
        Me.tbFaultDescription.Location = New System.Drawing.Point(121, 373)
        Me.tbFaultDescription.Multiline = True
        Me.tbFaultDescription.Name = "tbFaultDescription"
        Me.tbFaultDescription.Size = New System.Drawing.Size(551, 150)
        Me.tbFaultDescription.TabIndex = 14
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(12, 376)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(91, 36)
        Me.Label14.TabIndex = 26
        Me.Label14.Text = "Informacion " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "adicional:"
        '
        'tbEstimateCost
        '
        Me.tbEstimateCost.Location = New System.Drawing.Point(121, 539)
        Me.tbEstimateCost.Name = "tbEstimateCost"
        Me.tbEstimateCost.Size = New System.Drawing.Size(551, 26)
        Me.tbEstimateCost.TabIndex = 15
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(5, 542)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(114, 18)
        Me.Label15.TabIndex = 28
        Me.Label15.Text = "Costo estimado:"
        '
        'tbRepairTnC
        '
        Me.tbRepairTnC.Location = New System.Drawing.Point(121, 580)
        Me.tbRepairTnC.Multiline = True
        Me.tbRepairTnC.Name = "tbRepairTnC"
        Me.tbRepairTnC.Size = New System.Drawing.Size(551, 89)
        Me.tbRepairTnC.TabIndex = 31
        Me.tbRepairTnC.TabStop = False
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(12, 583)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(88, 36)
        Me.Label16.TabIndex = 30
        Me.Label16.Text = "Terminos y " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Condiciones:"
        '
        'btnCreateRepair
        '
        Me.btnCreateRepair.BackColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(129, Byte), Integer), CType(CType(197, Byte), Integer))
        Me.btnCreateRepair.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCreateRepair.Location = New System.Drawing.Point(526, 686)
        Me.btnCreateRepair.Name = "btnCreateRepair"
        Me.btnCreateRepair.Size = New System.Drawing.Size(146, 31)
        Me.btnCreateRepair.TabIndex = 16
        Me.btnCreateRepair.Text = "Crear reparacion"
        Me.btnCreateRepair.UseVisualStyleBackColor = False
        '
        'cbCustomerTitle
        '
        Me.cbCustomerTitle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbCustomerTitle.FormattingEnabled = True
        Me.cbCustomerTitle.Items.AddRange(New Object() {"Mr", "Mrs", "Miss", "Mx", "Dr", "Rev", "Prof"})
        Me.cbCustomerTitle.Location = New System.Drawing.Point(121, 12)
        Me.cbCustomerTitle.Name = "cbCustomerTitle"
        Me.cbCustomerTitle.Size = New System.Drawing.Size(170, 26)
        Me.cbCustomerTitle.TabIndex = 1
        '
        'frmCreateNewRepair
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(48, Byte), Integer), CType(CType(55, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(684, 729)
        Me.Controls.Add(Me.cbCustomerTitle)
        Me.Controls.Add(Me.btnCreateRepair)
        Me.Controls.Add(Me.tbRepairTnC)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.tbEstimateCost)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.tbFaultDescription)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.tbFaultType)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.tbAssetSerialNumber)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.tbAssetModel)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.tbAssetMake)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.tbCustomerMobilePhone)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.tbCustomerHomePhone)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.tbCustomerEmail)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.tbCustomerPostcode)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.tbCustomerCity)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.tbCustomerAddress)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbCustomerSurname)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbCustomerForename)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "frmCreateNewRepair"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Crear nueva reparacion:"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents tbCustomerForename As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents tbCustomerSurname As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents tbCustomerAddress As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents tbCustomerPostcode As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents tbCustomerCity As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents tbCustomerMobilePhone As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents tbCustomerHomePhone As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents tbCustomerEmail As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents tbAssetModel As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents tbAssetMake As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents tbAssetSerialNumber As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents tbFaultType As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents tbFaultDescription As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents tbEstimateCost As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents tbRepairTnC As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents btnCreateRepair As Button
    Friend WithEvents cbCustomerTitle As ComboBox
End Class
