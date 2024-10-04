Imports MySql.Data.MySqlClient

Public Class frmUserMgmt

    Private Sub frmUserMgmt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' This procedure is called when the User Management window is shown. It overlays the left (details) pane with a label,
        ' which prompts the user to select an employee first to view their details / take action. The refreshEmpList procedure
        ' is also called.

        lblUserPrompt.Visible = True
        lblUserPrompt.Dock = DockStyle.Fill

        refreshEmpList()
    End Sub

    Public Sub refreshEmpList()
        ' This procedure sends a query to the database requesting a full list of employee (user) records. The only fields
        ' shown to the user (for security) are the username, employee ID, employee name, and user access level.

        dgvCurrentUserList.Rows.Clear()

        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
            Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT employeeID, title, forename, surname, userAccessLevel, username FROM Employees;", dbConnection)

            dbConnection.Open()

            Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

            If dbReader.HasRows Then
                While dbReader.Read
                    dgvCurrentUserList.Rows.Add(New String() {dbReader(0).ToString, dbReader(5).ToString, dbReader(1).ToString & " " & dbReader(2).ToString &
                                                " " & dbReader(3).ToString, dbReader(4).ToString})
                End While
            End If

            dbConnection.Close()
            dbCommand.Dispose()
            dbConnection.Dispose()
        Catch ex As Exception
            MsgBox("Se produjo un error al buscar empleados en la base de datos de techcare." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
        End Try
    End Sub

    Private Sub dgvCurrentUserList_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvCurrentUserList.CellClick
        ' This procedure is called when a row in the Current User List is clicked. It simply brings up the selected employee's details on the left pane,
        ' and hides the user prompt to allow the user to select an option regarding the selected employee.

        lblEmpName.Text = dgvCurrentUserList.SelectedRows(0).Cells(2).Value.ToString
        lblEmpID.Text = dgvCurrentUserList.SelectedRows(0).Cells(0).Value.ToString
        lblUserAccessLvl.Text = dgvCurrentUserList.SelectedRows(0).Cells(3).Value.ToString
        lblUserPrompt.Visible = False
    End Sub

    Private Sub btnEditEmployeeDetails_Click(sender As Object, e As EventArgs) Handles btnEditEmployeeDetails.Click
        ' This procedure is called when the Edit Details button is clicked. A copy of all employee details are relayed through to the
        ' form, to allow the user to edit what is required, rather than having to re-type all the employee's details again.

        frmEditEmpDetails.tbTitle.Text = functions.obtainEmployeeDetails(lblEmpID.Text, 1)
        frmEditEmpDetails.tbForename.Text = functions.obtainEmployeeDetails(lblEmpID.Text, 2)
        frmEditEmpDetails.tbSurname.Text = functions.obtainEmployeeDetails(lblEmpID.Text, 3)

        If functions.obtainEmployeeDetails(lblEmpID.Text, 4) = "Basic" Then
            frmEditEmpDetails.rbBasicAccess.Checked = True
        Else
            frmEditEmpDetails.rbFullAccess.Checked = True
        End If

        frmEditEmpDetails.Text = "Editing Details for Employee: " & lblEmpID.Text
        frmEditEmpDetails.empID = lblEmpID.Text

        frmEditEmpDetails.ShowDialog()
    End Sub

    Private Sub btnResetEmpPwd_Click(sender As Object, e As EventArgs) Handles btnResetEmpPwd.Click
        ' This procedure is called to allow a user to change the password for a given employee. Since only users with Full system
        ' access are permitted to access the User Management area, the requirement to enter the user's current password is not included.
        ' **** (This program assumes that users with full system access will not abuse this function) ****

        frmResetPassword.empID = lblEmpID.Text
        frmResetPassword.ShowDialog()
    End Sub

    Private Sub btnDeleteEmp_Click(sender As Object, e As EventArgs) Handles btnDeleteEmp.Click
        ' This procedure is called when the Delete User button is clicked. First, the program checks that the user is not attempting to delete
        ' their own account. Once this check has passed, this procedure displays a confirmation message confirming the user intends to delete
        ' a user from the system. If the user clicks YES, an SQL query is sent to the database to remove the account from the system. It should
        ' be noted that this action is irreversible.

        If lblEmpID.Text = frmMainWindow.lblEmpID.Text Then
            MsgBox("No se puede eliminar un usuario que actualmente ha iniciado sesión. Por favor, cierre sesión en esta cuenta e inténtelo de nuevo.", MsgBoxStyle.Critical, "techcare")
        Else
            Dim msg As DialogResult = MessageBox.Show("Está a punto de eliminar a un empleado del sistema." & vbNewLine & " Esta accion es irreversible." &
                                                      vbNewLine & vbNewLine & "¿Quieres continuar?", "techcare", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If msg = DialogResult.Yes Then
                Try
                    Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                    Dim dbCommand As MySqlCommand = New MySqlCommand("DELETE FROM Employees WHERE employeeID=@empID", dbConnection)

                    dbConnection.Open()

                    dbCommand.Parameters.AddWithValue("@empID", lblEmpID.Text)

                    dbCommand.ExecuteNonQuery()

                    dbConnection.Close()
                    dbCommand.Dispose()
                    dbConnection.Dispose()

                    MsgBox("Usuario eliminado exitosamente.", MsgBoxStyle.Information, "techcare")
                    refreshEmpList()
                    lblUserPrompt.Visible = True
                    lblUserPrompt.Dock = DockStyle.Fill
                Catch ex As Exception
                    MsgBox("Se ha producido una excepción al eliminar este usuario." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                End Try
            End If
        End If
    End Sub

    Private Sub btnCreateNewEmp_Click(sender As Object, e As EventArgs) Handles btnCreateNewEmp.Click
        ' This procedure is called on clicking the Create User button.
        frmCreateUser.ShowDialog()
    End Sub
End Class