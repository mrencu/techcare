Imports MySql.Data.MySqlClient

Public Class frmUpdateRepairStatus

    Private Sub btnConfirmChanges_Click(sender As Object, e As EventArgs) Handles btnConfirmChanges.Click
        If rbBookedIn.Checked = True Then
            updateRepairStatus("Booked In")
        ElseIf rbServiceInProgress.Checked = True Then
            updateRepairStatus("Service in Progress")
        ElseIf rbTransferred.Checked = True Then
            updateRepairStatus("Transferred to External Body")
        ElseIf rbServiceCompleted.Checked = True Then
            updateRepairStatus("Service Completed")
        ElseIf rbAssetRemoved.Checked = True Then
            updateRepairStatus("Asset Removed")
        Else
            MsgBox("Elija una opción para actualizar el estado de reparación actual.", MsgBoxStyle.Exclamation, "techcare")
        End If
    End Sub

    Public Sub updateRepairStatus(ByVal status As String)
        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
            Dim dbCommand As MySqlCommand = New MySqlCommand("UPDATE Repairs SET currentRepairStatus=@status WHERE repairReference=@repairRef;", dbConnection)

            dbCommand.Parameters.AddWithValue("@status", status)
            dbCommand.Parameters.AddWithValue("@repairRef", frmRepairMgmt.lblRepairRef.Text)

            dbConnection.Open()

            dbCommand.ExecuteNonQuery()

            dbConnection.Close()

            frmRepairMgmt.lblCurrentRepairStatus.Text = status

            MsgBox("El estado de reparación cambió a: " & status & ".", MsgBoxStyle.Information, "techcare")

            dbConnection.Close()
            dbConnection.Dispose()
            dbCommand.Dispose()

            If status = "Asset Removed" Then
                frmRepairMgmt.btnAddRepairRemark.Enabled = False
                frmRepairMgmt.btnCustomerCollection.Enabled = False
                frmRepairMgmt.btnUpdateRepairQuote.Enabled = False
                frmRepairMgmt.btnUpdateRepairStatus.Enabled = False
            End If

            Me.Close()
        Catch ex As Exception
            MsgBox("Se produjo un error al cambiar el estado de reparación actual." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
            Me.Close()
        End Try
    End Sub
End Class
