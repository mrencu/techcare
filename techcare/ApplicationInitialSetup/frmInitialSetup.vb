Imports MySql.Data.MySqlClient
Public Class frmInitialSetup
    Private Sub frmInitialSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        step1.Dock = DockStyle.Fill
        step1.Visible = True
    End Sub

    Private Sub btnStep1Next_Click(sender As Object, e As EventArgs) Handles btnStep1Next.Click
        step1.Dock = DockStyle.None
        step1.Visible = False

        step2.Dock = DockStyle.Fill
        step2.Visible = True
    End Sub

    Private Sub btnStep2Prev_Click(sender As Object, e As EventArgs) Handles btnStep2Prev.Click
        step2.Dock = DockStyle.None
        step2.Visible = False
        tbStep2DbLocation.Clear()
        tbStep2DbPassword.Clear()
        tbStep2DbUsername.Clear()

        step1.Dock = DockStyle.Fill
        step1.Visible = True
    End Sub

    Private Sub btnStep2Next_Click(sender As Object, e As EventArgs) Handles btnStep2Next.Click
        If tbStep2DbLocation.Text = "" Or tbStep2DbUsername.Text = "" Then
            MsgBox("Uno o mas campos no contienen informacion requerida. Porfavor revisa todos los" &
                   " campos y intentalo nuevamente.", MsgBoxStyle.Exclamation, "techcare")
        Else
            Dim connectionStatus As Boolean = False

            Try
                Dim dbConnection As MySqlConnection = New MySqlConnection(
                    "Server=" & tbStep2DbLocation.Text & ";Uid=" & tbStep2DbUsername.Text &
                    ";Pwd=" & tbStep2DbPassword.Text & ";")
                dbConnection.Open()
                connectionStatus = True
                functions.log("Conexion al servidor exitosa...")
                dbConnection.Close()
            Catch ex As Exception
                connectionStatus = False
                functions.log("Conexion al servidor fallida...abortando la mision")
            End Try

            If connectionStatus = True Then
                If functions.databaseCheck("techcare") = True Then
                    functions.log("Db techcare encontrada, necesitamos confirmacion del usuario...")
                    Dim confirmation As DialogResult
                    confirmation = MessageBox.Show("Techcare ha detectado otra base de datos llamada " &
                        "'techcare' en tu servidor MySQL." & vbNewLine & "Paraa " &
                        "completar la configuracion, debemos borrar esa bd. Tenes dos opciones: " &
                        "a", "techcare", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                    If confirmation = DialogResult.Yes Then
                        functions.log("Usuario confirma remover esta bd... borrandooo")
                        Try
                            Dim dbConnection As MySqlConnection = New MySqlConnection(
                                "Server=" & tbStep2DbLocation.Text & ";Uid=" & tbStep2DbUsername.Text _
                                & ";Pwd=" & tbStep2DbPassword.Text & ";")
                            Dim dbCommand As MySqlCommand = New MySqlCommand("DROP DATABASE techcare;",
                                                                             dbConnection)

                            dbConnection.Open()
                            dbCommand.ExecuteNonQuery()
                            dbConnection.Close()

                            My.Settings.dbLocation = tbStep2DbLocation.Text
                            My.Settings.dbUsername = tbStep2DbUsername.Text
                            My.Settings.dbPassword = tbStep2DbPassword.Text

                            step2.Dock = DockStyle.None
                            step2.Visible = False
                            tbStep2DbLocation.Clear()
                            tbStep2DbPassword.Clear()
                            tbStep2DbUsername.Clear()

                            step3.Dock = DockStyle.Fill
                            step3.Visible = True
                            functions.log("BD borrada... siguiente paso")
                        Catch ex As Exception
                            MsgBox("La config no se pudo concretar por un error de conexion a la BD" &
                                   " A." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical,
                                                                                    "techcare")
                            functions.log("ERROR CATASTROFICO... CHAU" & ex.InnerException.Message)
                            Application.Exit()
                        End Try
                    Else
                        functions.log("Usuario rechazo borrar la bd, abortando")
                        MsgBox("La configuracion no puede continuar", MsgBoxStyle.Exclamation, "techcare")
                        Application.Exit()
                    End If
                Else
                    My.Settings.dbLocation = tbStep2DbLocation.Text
                    My.Settings.dbUsername = tbStep2DbUsername.Text
                    My.Settings.dbPassword = tbStep2DbPassword.Text

                    step2.Dock = DockStyle.None
                    step2.Visible = False
                    tbStep2DbLocation.Clear()
                    tbStep2DbPassword.Clear()
                    tbStep2DbUsername.Clear()

                    step3.Dock = DockStyle.Fill
                    step3.Visible = True

                    functions.log("sisi, todo okey, todo fino")
                End If
            Else
                MsgBox("No se ha podido conectar con las credenciales que diste" & vbNewLine & vbNewLine &
                       "Revisa estas cositas:" & vbNewLine & "A) El servidor esta corriendo y las credenciales son correctas" &
                       vbNewLine & "B) El user y la pass esten registrados en el servidorr" &
                       vbNewLine & "C) El usuario tiene permisos de lectura y escritura.",
                       MsgBoxStyle.Exclamation, "techcare")
                functions.log("Conexion fallida, nc xq")
            End If
        End If
    End Sub



    Private Sub btnStep3Previous_Click(sender As Object, e As EventArgs) Handles btnStep3Previous.Click
        step3.Dock = DockStyle.None
        step3.Visible = False

        step2.Dock = DockStyle.Fill
        step2.Visible = True
    End Sub

    Private Sub btnStep3RestoreFromBackup_Click(sender As Object, e As EventArgs) Handles btnStep3RestoreFromBackup.Click

        If functions.databaseCheck("techcare") = False Then
            functions.log("Reconstruyhendo base de datos...")
            functions.rebuildDatabase()
        End If

        functions.log("Restaurando de un backup...")

        My.Settings.dbName = "techcare"
        My.Settings.Save()
        My.Settings.Reload()

        Dim xamppLocation As String = ""
        Dim dbDumpLocation As String = ""

        xamppLocationDialog.Description = "Seleccionar carpeta XAMPP servidor"                                                                                    
        If xamppLocationDialog.ShowDialog = DialogResult.OK Then
            xamppLocation = xamppLocationDialog.SelectedPath & "\mysql\bin\"

            functions.log("Attempting to find mysqldump.exe")

            If System.IO.File.Exists(xamppLocation & "mysqldump.exe") = True Then
                If xamppRestoreFromLocationDialog.ShowDialog = DialogResult.OK Then
                    Try
                        Dim dbConnection As MySqlConnection
                        functions.log("Archivo encontrado, restaurando")
                        dbConnection = New MySqlConnection("Server=" & My.Settings.dbLocation &
                                                           ";Database=" & My.Settings.dbName &
                                                           ";Uid=" & My.Settings.dbUsername &
                                                           ";Pwd=" & My.Settings.dbPassword & ";")

                        dbConnection.Open()

                        dbDumpLocation = xamppRestoreFromLocationDialog.FileName

                        Dim backup As New Process
                        backup.StartInfo.FileName = "cmd.exe"
                        backup.StartInfo.UseShellExecute = False
                        backup.StartInfo.WorkingDirectory = xamppLocation
                        backup.StartInfo.RedirectStandardInput = True
                        backup.StartInfo.RedirectStandardOutput = True

                        backup.Start()

                        Dim backupStream As System.IO.StreamWriter = backup.StandardInput
                        Dim myStreamReader As System.IO.StreamReader = backup.StandardOutput

                        backupStream.WriteLine("mysql -u " & My.Settings.dbUsername &
                                               " -p " & My.Settings.dbName & " < """ &
                                               dbDumpLocation & """")

                        backupStream.Close()
                        backup.WaitForExit()
                        backup.Close()

                        dbConnection.Close()

                        step3.Dock = DockStyle.None
                        step3.Visible = False

                        step4a.Dock = DockStyle.Fill
                        step4a.Visible = True

                        functions.log("Restauracion completada, finalizando...")
                    Catch ex As Exception
                        MsgBox("Ocurrio un error mientras restauraba la BD, ayuda" &
                               vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                        functions.log("Restauracion fallida " & ex.InnerException.Message)
                    End Try
                End If
            End If
        Else
            MsgBox("MySQL.exe was Not found in the specified XAMPP directory. Please check" &
                   " that the XAMPP directory has been selected " &
                   "and try again.", MsgBoxStyle.Critical, "techcare")
            functions.log("MySQL not found in directory specified by user.")
        End If
    End Sub

    Private Sub btnStep3StartFromScratch_Click(sender As Object, e As EventArgs) Handles btnStep3StartFromScratch.Click
        If functions.databaseCheck("techcare") = False Then
            functions.log("Reconstruyendo BD...")
            functions.rebuildDatabase()
        End If

        My.Settings.dbName = "techcare"
        My.Settings.Save()
        My.Settings.Reload()

        step3.Dock = DockStyle.None
        step3.Visible = False

        step4b.Dock = DockStyle.Fill
        step4b.Visible = True

        functions.log("NOS VAMOS AL PASO 4.")
    End Sub

    Private Sub btnStep4aFinish_Click(sender As Object, e As EventArgs) Handles btnStep4aFinish.Click

        step4a.Dock = DockStyle.None
        step4a.Visible = False
        My.Settings.userFirstRun = False
        My.Settings.Save()
        My.Settings.Reload()

        functions.log("** Reiniciando programa **")

        Application.Restart()
    End Sub


    Private Sub btnStep4BPrevious_Click(sender As Object, e As EventArgs) Handles btnStep4BPrevious.Click
        step4b.Dock = DockStyle.None
        step4b.Visible = False
        tbStep4BTitle.Clear()
        tbStep4BFname.Clear()
        tbStep4BSurname.Clear()
        tbStep4BPassword.Clear()
        tbStep4BConfirmPassword.Clear()

        step3.Visible = True
        step3.Dock = DockStyle.Fill
    End Sub

    Private Sub btnStep4BNext_Click(sender As Object, e As EventArgs) Handles btnStep4BNext.Click
        If tbStep4BTitle.Text = "" Or tbStep4BFname.Text = "" Or tbStep4BSurname.Text = "" Or tbStep4BPassword.Text = "" Or tbStep4BConfirmPassword.Text = "" Then
            MsgBox("Uno o mas campos le falta informacion, revisar eso" &
                   " , que todos los campos tengan bien la info", MsgBoxStyle.Exclamation, "techcare")
            tbStep4BPassword.Clear()
            tbStep4BConfirmPassword.Clear()
        Else
            If tbStep4BPassword.Text = tbStep4BConfirmPassword.Text Then
                If functions.validateNewPassword(tbStep4BPassword.Text) = True Then
                    Try
                        Dim username As String = functions.generateUsername(tbStep4BFname.Text,
                                                                            tbStep4BSurname.Text)

                        If username Is Nothing Then
                            Throw New System.Exception("No es posible generar una cuenta de usuario" &
                                                       ". Proceso terminado")
                        Else
                            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation &
                                                                                      ";Database=" & My.Settings.dbName &
                                                                                      ";Uid=" & My.Settings.dbUsername &
                                                                                      ";Pwd=" & My.Settings.dbPassword & ";")
                            Dim dbCommand As MySqlCommand = New MySqlCommand("INSERT INTO Employees (employeeID, title, " &
                                                                             "forename, surname, userAccessLevel, username," &
                                                                             " password) VALUES (@empID, @empTitle, @fname," &
                                                                             "@sname, @ual, @username, @password);", dbConnection)

                            dbConnection.Open()

                            dbCommand.Parameters.AddWithValue("@empID", functions.generateUid("Employees", "employeeID", 7))
                            dbCommand.Parameters.AddWithValue("@empTitle", tbStep4BTitle.Text)
                            dbCommand.Parameters.AddWithValue("@fname", tbStep4BFname.Text)
                            dbCommand.Parameters.AddWithValue("@sname", tbStep4BSurname.Text)
                            dbCommand.Parameters.AddWithValue("@ual", "Full")
                            dbCommand.Parameters.AddWithValue("@username", username)
                            dbCommand.Parameters.AddWithValue("@password", tbStep4BPassword.Text)

                            dbCommand.ExecuteNonQuery()

                            dbConnection.Close()
                            dbCommand.Dispose()
                            dbConnection.Dispose()

                            MsgBox("Cuenta de usuario creada!" & vbNewLine & vbNewLine & "El usuario que le hice es: " & username &
                                   ". Recordarlo xfa", MsgBoxStyle.Information, "techcare")

                            step4b.Dock = DockStyle.None
                            step4b.Visible = False
                            tbStep4BTitle.Clear()
                            tbStep4BFname.Clear()
                            tbStep4BSurname.Clear()
                            tbStep4BPassword.Clear()
                            tbStep4BConfirmPassword.Clear()

                            step5.Visible = True
                            step5.Dock = DockStyle.Fill
                        End If
                    Catch ex As Exception
                        MsgBox("Ocurrio un error mientras se creaba la cuenta de usuario" & vbNewLine & vbNewLine &
                               ex.Message, MsgBoxStyle.Critical, "techcare")
                        tbStep4BPassword.Clear()                                                                                    
                    End Try
                Else
                    MsgBox("La contraseña ingresada no cumple con los requerimientos.",
                           MsgBoxStyle.Exclamation, "techcare")
                    tbStep4BConfirmPassword.Clear()
                    tbStep4BPassword.Clear()
                End If
            Else
                MsgBox("Las contraseñas no coinciden, intenta otra vezz", MsgBoxStyle.Exclamation, "techcare")
                tbStep4BPassword.Clear()
                tbStep4BConfirmPassword.Clear()
            End If
        End If
    End Sub

    Private Sub btnStep5Next_Click(sender As Object, e As EventArgs) Handles btnStep5Next.Click
        If tbStep5BusName.Text = "" Or tbStep5BusAddress.Text = "" Or tbStep5BusContact.Text = "" Then
            MsgBox("Revisa que te faltan campos",
                   MsgBoxStyle.Exclamation, "techcare")
        Else
            Try
                Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation &
                                                                          ";Database=" & My.Settings.dbName &
                                                                          ";Uid=" & My.Settings.dbUsername &
                                                                          ";Pwd=" & My.Settings.dbPassword & ";")
                Dim dbCommand As MySqlCommand = New MySqlCommand("INSERT INTO Business (name, address, phoneNumber, repairTnC)" &
                                                                 "VALUES (@busName, @busAddr, @busPhn, @busRepairTNC);",
                                                                 dbConnection)

                dbConnection.Open()

                dbCommand.Parameters.AddWithValue("@busName", tbStep5BusName.Text)
                dbCommand.Parameters.AddWithValue("@busAddr", tbStep5BusAddress.Text)
                dbCommand.Parameters.AddWithValue("@busPhn", tbStep5BusContact.Text)
                dbCommand.Parameters.AddWithValue("@busRepairTNC", tbStep5BusRepairTnC.Text)

                dbCommand.ExecuteNonQuery()

                dbConnection.Close()
                dbCommand.Dispose()
                dbConnection.Dispose()

                step5.Visible = False
                step5.Dock = DockStyle.None
                tbStep5BusName.Clear()
                tbStep5BusAddress.Clear()
                tbStep5BusContact.Clear()
                tbStep5BusRepairTnC.Clear()

                step6.Visible = True
                step6.Dock = DockStyle.Fill
            Catch ex As Exception
                MsgBox("No se han podido agregar los detalles de la empresa" & vbNewLine & vbNewLine & ex.Message,
                       MsgBoxStyle.Critical, "techcare")
            End Try
        End If
    End Sub

    Private Sub tbStep5BusContact_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbStep5BusContact.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnStep6Finish_Click(sender As Object, e As EventArgs) Handles btnStep6Finish.Click
        step6.Dock = DockStyle.None
        step6.Visible = False
        My.Settings.userFirstRun = False
        My.Settings.Save()
        My.Settings.Reload()

        Application.Restart()
    End Sub


End Class