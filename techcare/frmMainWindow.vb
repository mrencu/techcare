Public Class frmMainWindow

    Private Sub frmMainWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        System.Threading.Thread.Sleep(3000)
        Me.WindowState = FormWindowState.Maximized

        functions.log("Pantalla de bienvenida inicializada")

        For Each ctl As Control In Me.Controls
            If TypeOf ctl Is MdiClient Then
                ctl.BackColor = Me.BackColor
            End If
        Next ctl

        lblVersionID.Text = "Version " & My.Application.Info.Version.ToString

        If My.Settings.userAskedForConsent = False Then
            frmLogsConsent.MdiParent = Me
            frmLogsConsent.Show()
        Else
            If My.Settings.userFirstRun = True Then
                frmInitialSetup.MdiParent = Me
                frmInitialSetup.Show()
                functions.log("La configuración inicial ha comenzado.")
            Else
                While functions.databaseCheck(My.Settings.dbName) = False
                    Dim conf As Integer = MsgBox("No se puede encontrar la base de datos de techcare en el servidor." & vbNewLine & vbNewLine &
                                                 "Compruebe que el servidor MySQL se esté ejecutando y luego haga clic en Aceptar para volver a intentarlo.", MsgBoxStyle.Exclamation, "techcare")
                End While

                frmLogin.MdiParent = Me
                frmLogin.Show()
            End If
        End If

        Try
            If functions.obtainBusinessDetails(0) = "" Then
                Me.Text = "techcare"
            Else
                Me.Text = "techcare: " & functions.obtainBusinessDetails(0) & ", " & functions.obtainBusinessDetails(1) &
                                ", " & functions.obtainBusinessDetails(2)
            End If
        Catch ex As Exception
            Me.Text = "techcare"
            functions.log("Se produjo una excepción al agregar detalles de la empresa:" & ex.Message)
        End Try
    End Sub

    Private Sub frmMainWindow_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Me.Refresh()
    End Sub

    Private Sub btnSignOut_Click(sender As Object, e As EventArgs) Handles btnSignOut.Click
        functions.log("Botón Cerrar sesión presionado")

        Dim msg As Integer
        msg = MsgBox("¿Estás seguro de que deseas cerrar sesión?" & vbNewLine & vbNewLine & "Todo lo que no se guarde se perderá.", MsgBoxStyle.YesNo, "techcare")

        If msg = MsgBoxResult.Yes Then
            functions.log("Se inició la secuencia de cierre de sesión")
            frmAppConfig.Close()
            frmAddRepairRemark.Close()
            frmCreateNewRepair.Close()
            frmRepairMgmt.Close()
            frmUpdateQuote.Close()
            frmUpdateRepairStatus.Close()
            frmCreateUser.Close()
            frmEditEmpDetails.Close()
            frmResetPassword.Close()
            frmUserMgmt.Close()

            lblCurrentUser.Text = ""
            lblEmpID.Text = ""
            lblCurrentUser.Visible = False
            lblEmpID.Visible = False
            sidePanel.Visible = False
            frmLogin.MdiParent = Me
            frmLogin.Show()
            Me.Refresh()
            functions.log("Secuencia de cierre de sesión completada")
        End If
    End Sub

    Private Sub btnNewRepair_Click(sender As Object, e As EventArgs) Handles btnNewRepair.Click
        functions.log("Nuevo botón de reparación presionado")
        frmCreateNewRepair.MdiParent = Me
        frmCreateNewRepair.Show()
    End Sub

    Private Sub btnRepairMgmt_Click(sender As Object, e As EventArgs) Handles btnRepairMgmt.Click
        functions.log("Botón de gestión de reparación presionado")
        frmRepairMgmt.MdiParent = Me
        frmRepairMgmt.Show()
    End Sub

    Private Sub btnUserMgmt_Click(sender As Object, e As EventArgs) Handles btnUserMgmt.Click
        functions.log("Botón de gestión de usuarios presionado")

        If functions.obtainEmployeeDetails(lblEmpID.Text, 4) = "Full" Then
            functions.log("El usuario actual tiene permiso para acceder a la Gestión de usuarios, abriendo...")
            frmUserMgmt.MdiParent = Me
            frmUserMgmt.Show()
        Else
            functions.log("Current user does NOT have permission to access User Management")
            MsgBox("No tienes el nivel de acceso requerido para acceder a esta área.", MsgBoxStyle.Exclamation, "techcare")
        End If
    End Sub

    Private Sub btnApplicationSettings_Click(sender As Object, e As EventArgs) Handles btnApplicationSettings.Click
        functions.log("Application Settings button pressed")

        If functions.obtainEmployeeDetails(lblEmpID.Text, 4) = "Full" Then
            functions.log("Current user has permission to access Application Settings, opening...")

            frmAppConfig.MdiParent = Me
            frmAppConfig.Show()
        Else
            functions.log("Current user does NOT permission to access Application Settings")
            MsgBox("No tienes el nivel de acceso requerido para acceder a esta área.", MsgBoxStyle.Exclamation, "techcare")
        End If
    End Sub
End Class
