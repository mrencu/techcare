Imports MySql.Data.MySqlClient

Public Class frmUpdateQuote

    Private Sub tbNewRepairQuote_Leave(sender As Object, e As EventArgs) Handles tbNewRepairQuote.Leave
        Try
            tbNewRepairQuote.Text = CDec(tbNewRepairQuote.Text).ToString("c")
        Catch ex As Exception
            tbNewRepairQuote.Clear()
            MsgBox("Cotización no válida ingresada. Por favor inténtalo de nuevo.", MsgBoxStyle.Exclamation, "techcare")
        End Try
    End Sub

    Private Sub btnConfirmChanges_Click(sender As Object, e As EventArgs) Handles btnConfirmChanges.Click
        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
            Dim dbCommand As MySqlCommand = New MySqlCommand("UPDATE Repairs SET estimateQuote=@quote WHERE repairReference=@repairRef;", dbConnection)

            dbCommand.Parameters.AddWithValue("@quote", tbNewRepairQuote.Text)
            dbCommand.Parameters.AddWithValue("@repairRef", frmRepairMgmt.lblRepairRef.Text)

            dbConnection.Open()

            dbCommand.ExecuteNonQuery()

            dbConnection.Close()

            frmRepairMgmt.lblQuote.Text = tbNewRepairQuote.Text

            MsgBox("Cotización actualizada exitosamente.", MsgBoxStyle.Information, "techcare")

            dbConnection.Close()
            dbConnection.Dispose()
            dbCommand.Dispose()

            Me.Close()
        Catch ex As Exception
            MsgBox("Se produjo un error al actualizar la cotización de reparación." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
            Me.Close()
        End Try
    End Sub

    Private Sub frmUpdateQuote_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbNewRepairQuote.Text = frmRepairMgmt.lblQuote.Text
    End Sub
End Class