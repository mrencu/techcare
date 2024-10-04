Imports MySql.Data.MySqlClient

Public Class frmAddRepairRemark

    Private Sub btnAddRemark_Click(sender As Object, e As EventArgs) Handles btnAddRemark.Click

        If tbRemark.Text = "" Then
            MsgBox("No se puede agregar un comentario vacío a la reparación.", MsgBoxStyle.Exclamation, "techcare")
        Else
            Try
                Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                Dim dbCommand As MySqlCommand = New MySqlCommand("INSERT INTO remarks (commentID, employeeID, jobReference, timestamp, comment) VALUES " &
                                                                 "(@commentID, @employeeID, @repairRef, @timestamp, @comment);", dbConnection)

                Dim commentID As String = functions.generateUid("remarks", "commentID", 8)

                Dim timestamp As String = System.DateTime.Now.ToShortDateString & " " & System.DateTime.Now.ToShortTimeString

                dbCommand.Parameters.AddWithValue("@commentID", commentID)
                dbCommand.Parameters.AddWithValue("@employeeID", frmMainWindow.lblEmpID.Text)
                dbCommand.Parameters.AddWithValue("@repairRef", frmRepairMgmt.lblRepairRef.Text)
                dbCommand.Parameters.AddWithValue("@timestamp", timestamp)
                dbCommand.Parameters.AddWithValue("@comment", tbRemark.Text)

                dbConnection.Open()

                dbCommand.ExecuteNonQuery()

                dbConnection.Close()

                frmRepairMgmt.dgvRepairRemarks.Rows.Add(New String() {functions.obtainEmployeeDetails(frmMainWindow.lblEmpID.Text, 2) & " " &
                                                        functions.obtainEmployeeDetails(frmMainWindow.lblEmpID.Text, 3), timestamp, tbRemark.Text})

                MsgBox("Comentario añadido:", MsgBoxStyle.Information, "techcare")

                dbConnection.Close()
                dbConnection.Dispose()
                dbCommand.Dispose()

                Me.Close()
            Catch ex As Exception
                MsgBox("Ocurrio un error." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                Me.Close()
            End Try
        End If
    End Sub
End Class