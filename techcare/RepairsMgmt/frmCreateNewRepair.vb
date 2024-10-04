Imports MySql.Data.MySqlClient
Public Class frmCreateNewRepair

    Private Sub frmCreateNewRepair_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
            Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT repairTnC FROM Business", dbConnection)

            dbConnection.Open()

            Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

            If dbReader.HasRows Then
                While dbReader.Read
                    If dbReader(0).ToString = "" Then
                        tbRepairTnC.Text = "You have not set any Terms and Conditions for your repair services."
                        tbRepairTnC.ReadOnly = True
                    Else
                        tbRepairTnC.Text = dbReader(0).ToString
                        tbRepairTnC.ReadOnly = True
                    End If
                End While
            Else
                tbRepairTnC.Text = "You have not set any Terms and Conditions for your repair services."
                tbRepairTnC.ReadOnly = True
            End If

            dbConnection.Close()
            dbCommand.Dispose()
            dbConnection.Dispose()
        Catch ex As Exception
            tbRepairTnC.Text = ex.Message
            tbRepairTnC.ReadOnly = True
        End Try
    End Sub

    Private Sub tbEstimateCost_Leave(sender As Object, e As EventArgs) Handles tbEstimateCost.Leave
        Try
            tbEstimateCost.Text = CDec(tbEstimateCost.Text).ToString("c")
        Catch ex As Exception
            tbEstimateCost.Clear()
            MsgBox("Cotización no válida ingresada. Por favor inténtalo de nuevo.", MsgBoxStyle.Exclamation, "techcare")
        End Try
    End Sub

    Private Sub btnCreateRepair_Click(sender As Object, e As EventArgs) Handles btnCreateRepair.Click

        If cbCustomerTitle.SelectedIndex = -1 Or tbCustomerForename.Text = "" Or tbCustomerSurname.Text = "" Or tbCustomerAddress.Text = "" _
            Or tbCustomerCity.Text = "" Or tbCustomerPostcode.Text = "" Or tbAssetMake.Text = "" Or tbAssetModel.Text = "" Or tbAssetSerialNumber.Text = "" _
            Or tbFaultType.Text = "" Or tbEstimateCost.Text = "" Then
            MsgBox("A uno o más campos les falta información requerida. Por favor verifique sus entradas e inténtelo nuevamente.", MsgBoxStyle.Exclamation, "techcare")
        Else
            Try
                Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                Dim dbCommand As MySqlCommand = New MySqlCommand("INSERT INTO Repairs (repairReference, title, forename, surname, address, city, postcode, homePhone, " &
                                                                 "mobilePhone, emailAddress, assetMake, assetModel, assetSerialNumber, currentRepairStatus, intakeDate, " &
                                                                 "faultDescription, estimateQuote) VALUES (@repairReference, @title, @forename, @surname, @address, " &
                                                                 "@city, @postcode, @homePhone, @mobilePhone, @emailAddress, @assetMake, @assetModel, @assetSerialNumber, " &
                                                                 "@currentRepairStatus, @intakeDate, @faultDescription, @estimateQuote);", dbConnection)

                Dim repairRef As String = functions.generateUid("Repairs", "repairReference", 6)

                dbCommand.Parameters.AddWithValue("@repairReference", repairRef)
                dbCommand.Parameters.AddWithValue("@title", cbCustomerTitle.SelectedItem.ToString)
                dbCommand.Parameters.AddWithValue("@forename", tbCustomerForename.Text)
                dbCommand.Parameters.AddWithValue("@surname", tbCustomerSurname.Text)
                dbCommand.Parameters.AddWithValue("@address", tbCustomerAddress.Text)
                dbCommand.Parameters.AddWithValue("@city", tbCustomerCity.Text)
                dbCommand.Parameters.AddWithValue("@postcode", tbCustomerPostcode.Text)
                dbCommand.Parameters.AddWithValue("@homePhone", tbCustomerHomePhone.Text)
                dbCommand.Parameters.AddWithValue("@mobilePhone", tbCustomerMobilePhone.Text)
                dbCommand.Parameters.AddWithValue("@emailAddress", tbCustomerEmail.Text)
                dbCommand.Parameters.AddWithValue("@assetMake", tbAssetMake.Text)
                dbCommand.Parameters.AddWithValue("@assetModel", tbAssetModel.Text)
                dbCommand.Parameters.AddWithValue("@assetSerialNumber", tbAssetSerialNumber.Text)
                dbCommand.Parameters.AddWithValue("@currentRepairStatus", "Booked In")
                dbCommand.Parameters.AddWithValue("@intakeDate", System.DateTime.Now.ToShortDateString)
                dbCommand.Parameters.AddWithValue("@faultDescription", tbFaultType.Text & ": " & tbFaultDescription.Text)
                dbCommand.Parameters.AddWithValue("@estimateQuote", tbEstimateCost.Text)

                dbConnection.Open()

                dbCommand.ExecuteNonQuery()

                dbConnection.Close()
                dbCommand.Dispose()
                dbConnection.Dispose()

                MsgBox("Reserva de reparación creada a las " & System.DateTime.Now.ToShortTimeString & " on " & System.DateTime.Now.ToShortDateString & vbNewLine & vbNewLine &
                       "Referencia de reparación: " & repairRef.ToString)

                repairRef = Nothing

                Me.Close()
            Catch ex As Exception
                MsgBox("Se produjo un error al registrar una nueva reparación en el sistema." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                Me.Close()
            End Try
        End If
    End Sub
End Class