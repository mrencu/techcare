Imports MySql.Data.MySqlClient

Public Class frmEditEmpDetails

    Public empID As Integer

    Private Sub tbSurname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbSurname.KeyPress
        lblEditSurnameWarning.Visible = True
    End Sub

    Private Sub tbForename_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbForename.KeyPress
        lblEditSurnameWarning.Visible = True
    End Sub

    Private Sub btnConfirmChanges_Click(sender As Object, e As EventArgs) Handles btnConfirmChanges.Click
        ' This procedure is called upon clicking the "Confirm Changes" button. First, a validation check to ensure all fields are filled in is completed, before
        ' sending an SQL query to update the details of the selected employee, where the Employee ID is equal to the passed in employee ID through the empID variable.
        If tbTitle.TextLength = 0 Or tbForename.TextLength = 0 Or tbSurname.TextLength = 0 Or (rbBasicAccess.Checked = False And rbFullAccess.Checked = False) Then
            MsgBox("A uno o más campos les falta información requerida. Por favor verifique sus entradas e inténtelo nuevamente.", MsgBoxStyle.Critical, "techcare")
        Else
            Try
                Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                Dim dbCommand As MySqlCommand = New MySqlCommand("UPDATE Employees SET title=@empTitle, forename=@fname, surname=@sname, userAccessLevel=@ual WHERE employeeID=@empID;", dbConnection)

                Dim ual As String = Nothing

                If rbBasicAccess.Checked = True Then
                    ual = "Basic"
                ElseIf rbFullAccess.Checked = True Then
                    ual = "Full"
                End If

                dbCommand.Parameters.AddWithValue("@empTitle", tbTitle.Text)
                dbCommand.Parameters.AddWithValue("@fname", tbForename.Text)
                dbCommand.Parameters.AddWithValue("@sname", tbSurname.Text)
                dbCommand.Parameters.AddWithValue("@ual", ual)
                dbCommand.Parameters.AddWithValue("@empID", empID)

                dbConnection.Open()

                dbCommand.ExecuteNonQuery()

                dbConnection.Close()
                dbCommand.Dispose()
                dbConnection.Dispose()

                MsgBox("Los detalles se han actualizado correctamente.", MsgBoxStyle.Information, "techcare")

                frmUserMgmt.refreshEmpList()
                frmUserMgmt.lblEmpName.Text = tbTitle.Text & " " & tbForename.Text & " " & tbSurname.Text
                frmUserMgmt.lblUserAccessLvl.Text = ual
                lblEditSurnameWarning.Visible = False
                Me.Close()
            Catch ex As Exception
                MsgBox("Se ha producido un error al actualizar los datos del usuario." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                frmUserMgmt.refreshEmpList()
                Me.Close()
            End Try
        End If
    End Sub
End Class