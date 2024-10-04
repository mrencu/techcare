Imports System.Windows.Forms

Public Class frmLogin

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        startLogin()
    End Sub

    Private Sub tbPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles tbPassword.KeyDown
        If e.KeyCode = Keys.Enter Then startLogin()
    End Sub

    Private Sub tbUsername_KeyDown(sender As Object, e As KeyEventArgs) Handles tbUsername.KeyDown
        If e.KeyCode = Keys.Enter Then startLogin()
    End Sub

    Public Sub startLogin()
        If tbUsername.TextLength > 0 And tbPassword.TextLength > 0 Then
            Dim empID As String = Convert.ToString(functions.authenticate(tbUsername.Text, tbPassword.Text))

            If empID = "0" Then
                MsgBox("¡Error de inicio de sesión! El nombre de usuario o la contraseña son incorrectos.", MsgBoxStyle.Critical, "techcare")
                tbPassword.Clear()
                tbPassword.Focus()
            Else
                frmMainWindow.sidePanel.Visible = True
                frmMainWindow.lblEmpID.Visible = True
                frmMainWindow.lblCurrentUser.Visible = True
                frmMainWindow.lblEmpID.Text = empID
                frmMainWindow.lblCurrentUser.Text = functions.obtainEmployeeDetails(empID, 2).ToString & " " & functions.obtainEmployeeDetails(empID, 3).ToString
                frmMainWindow.Refresh()
                Me.Close()
            End If
        Else
            MsgBox("¡Error de inicio de sesión! Se requiere un nombre de usuario y contraseña válidos para iniciar sesión.", MsgBoxStyle.Critical, "techcare")
            tbPassword.Clear()
            tbPassword.Focus()
        End If
    End Sub
End Class
