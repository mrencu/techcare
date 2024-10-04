Imports MySql.Data.MySqlClient
Public Class frmAppConfig

    Public Sub HideAllPanels()
        databaseConfigPanel.Visible = False
        databaseConfigPanel.Dock = DockStyle.None
        btnDatabaseConfig.BackColor = Color.FromArgb(44, 48, 55)

        dataMgmtPanel.Visible = False
        dataMgmtPanel.Dock = DockStyle.None
        btnDataMgmt.BackColor = Color.FromArgb(44, 48, 55)

        aboutTechcarePanel.Visible = False
        aboutTechcarePanel.Dock = DockStyle.None
        btnAboutTechcare.BackColor = Color.FromArgb(44, 48, 55)

        resetSoftwarePanel.Visible = False
        resetSoftwarePanel.Dock = DockStyle.None
        btnResetSoftware.BackColor = Color.FromArgb(44, 48, 55)
    End Sub
    Private Sub btnDatabaseConfig_Click(sender As Object, e As EventArgs) Handles btnDatabaseConfig.Click
        HideAllPanels()

        databaseConfigPanel.Dock = DockStyle.Fill
        databaseConfigPanel.Visible = True

        tbDbName.Text = My.Settings.dbName
        tbDbName.ReadOnly = True
        tbDbLocation.Text = My.Settings.dbLocation
        tbDbUsername.Text = My.Settings.dbUsername
        tbDbPassword.Text = My.Settings.dbPassword

        btnDatabaseConfig.BackColor = Color.FromArgb(32, 129, 197)
    End Sub

    Private Sub btnSaveDbConfig_Click(sender As Object, e As EventArgs) Handles btnSaveDbConfig.Click
        Dim confirmation As DialogResult

        confirmation = MessageBox.Show("Precaucion!!" & vbNewLine & vbNewLine &
                                       "Si las configuraciones enumeradas aquí no coinciden con su servidor, este software dejará de funcionar." &
                                       vbNewLine & "Techcare debe reiniciarse para aplicar sus cambios. ¿Quieres continuar?",
                                       "techcare", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If confirmation = DialogResult.Yes Then
            If tbDbName.TextLength = 0 Or tbDbLocation.TextLength = 0 Then
                MsgBox("Debe tener una ubicación válida del servidor MySQL y un nombre de base de datos para continuar.", MsgBoxStyle.Critical, "techcare")
            Else
                My.Settings.dbLocation = tbDbLocation.Text
                My.Settings.dbName = tbDbName.Text
                My.Settings.dbUsername = tbDbUsername.Text
                My.Settings.dbPassword = tbDbPassword.Text
                My.Settings.Save()
                My.Settings.Reload()
                MsgBox("Configuración de la base de datos actualizada correctamente. Techcare ahora se reiniciará.", MsgBoxStyle.Information, "techcare")

                Application.Restart()
            End If
        End If
    End Sub

    Private Sub btnTestDbConnection_Click(sender As Object, e As EventArgs) Handles btnTestDbConnection.Click
        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & tbDbLocation.Text & ";Database=" & tbDbName.Text & ";Uid=" &
                                                                      tbDbUsername.Text & ";Pwd=" & tbDbPassword.Text & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT schema_name FROM information_schema.schemata WHERE schema_name = @dbName;",
                                                             dbConnection)

            dbCommand.Parameters.AddWithValue("@dbName", tbDbName.Text)

            dbConnection.Open()

            If dbCommand.ExecuteScalar = tbDbName.Text Then
                MsgBox("¡La conexión a la base de datos se ha realizado correctamente!", MsgBoxStyle.Information, "techcare")
                btnSaveDbConfig.Enabled = True
                lblWarningDbConfig.Visible = True
                tbDbName.ReadOnly = True
                tbDbLocation.ReadOnly = True
                tbDbUsername.ReadOnly = True
                tbDbPassword.ReadOnly = True
            Else
                MsgBox("La conexión a la base de datos ha fallado. Por favor verifique que toda la información sea correcta e inténtelo nuevamente.", MsgBoxStyle.Critical, "techcare")
                btnSaveDbConfig.Enabled = False
                lblWarningDbConfig.Visible = False
                tbDbName.ReadOnly = True
                tbDbLocation.ReadOnly = False
                tbDbUsername.ReadOnly = False
                tbDbPassword.ReadOnly = False
            End If

            dbConnection.Close()
        Catch ex As Exception
            MsgBox("La conexión a la base de datos ha fallado. Por favor verifique que toda la información sea correcta e inténtelo nuevamente." & vbNewLine & vbNewLine &
                   "Más información relacionada con el problema se encuentra a continuación:" & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
        End Try
    End Sub
    Private Sub btnDataMgmt_Click(sender As Object, e As EventArgs) Handles btnDataMgmt.Click
        HideAllPanels()

        dataMgmtPanel.Dock = DockStyle.Fill
        dataMgmtPanel.Visible = True

        btnDataMgmt.BackColor = Color.FromArgb(32, 129, 197)
    End Sub

    Private Sub btnDbMgmtBackupData_Click(sender As Object, e As EventArgs) Handles btnDbMgmtBackupData.Click

        Dim xamppLocation As String = ""
        Dim dbDumpLocation As String = ""

        xamppBackupToLocationDialog.FileName = "Techcare DB Backup " & DateTime.Now.ToString("dd-MM-yyyy HH-mm") & ".sql"
        xamppLocationDialog.Description = "Select XAMPP Server folder."

        If xamppLocationDialog.ShowDialog = DialogResult.OK Then
            xamppLocation = xamppLocationDialog.SelectedPath & "\mysql\bin\"

            If System.IO.File.Exists(xamppLocation & "mysqldump.exe") = True Then
                If xamppBackupToLocationDialog.ShowDialog = DialogResult.OK Then
                    Try
                        Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName & ";Uid=" &
                                                                                 My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
                        dbConnection.Open()

                        dbDumpLocation = xamppBackupToLocationDialog.FileName

                        Dim backup As New Process
                        backup.StartInfo.FileName = "cmd.exe"
                        backup.StartInfo.UseShellExecute = False
                        backup.StartInfo.WorkingDirectory = xamppLocation
                        backup.StartInfo.RedirectStandardInput = True
                        backup.StartInfo.RedirectStandardOutput = True

                        backup.Start()

                        Dim backupStream As System.IO.StreamWriter = backup.StandardInput
                        Dim myStreamReader As System.IO.StreamReader = backup.StandardOutput

                        backupStream.WriteLine("mysqldump --user=" & My.Settings.dbUsername & " --password=" & My.Settings.dbPassword & " --host=" &
                                               My.Settings.dbLocation & " --databases " & My.Settings.dbName & " > """ & dbDumpLocation & """")

                        backupStream.Close()
                        backup.WaitForExit()
                        backup.Close()

                        dbConnection.Close()

                        MsgBox("Backup completado!", MsgBoxStyle.Information, "techcare")
                    Catch ex As Exception
                        MsgBox("Error mientras hacia mi backup" & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                    End Try
                End If
            Else
                MsgBox("MySQLDump.exe no se encontró en el directorio XAMPP especificado. Verifique que se haya seleccionado el directorio XAMPP" &
                       "y intentalo nuevamente", MsgBoxStyle.Critical, "techcare")
            End If
        End If
    End Sub

    Private Sub btnDbMgmtRestoreData_Click(sender As Object, e As EventArgs) Handles btnDbMgmtRestoreData.Click
        Dim xamppLocation As String = ""
        Dim dbDumpLocation As String = ""

        xamppLocationDialog.Description = "Select XAMPP Server folder."

        If xamppLocationDialog.ShowDialog = DialogResult.OK Then
            xamppLocation = xamppLocationDialog.SelectedPath & "\mysql\bin\"

            If System.IO.File.Exists(xamppLocation & "mysqldump.exe") = True Then
                Dim confirmation As DialogResult

                confirmation = MessageBox.Show("La restauración de datos reiniciará techcare. ¿Quieres continuar?",
                                               "techcare", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If confirmation = DialogResult.Yes Then
                    If xamppRestoreFromLocationDialog.ShowDialog = DialogResult.OK Then
                        Try
                            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName & ";Uid=" &
                                                                                     My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")

                            dbConnection.Open()

                            Dim dbCommand As MySqlCommand = New MySqlCommand("DROP DATABASE techcare;", dbConnection)
                            dbCommand.ExecuteNonQuery()

                            dbConnection.Close()
                            dbConnection.Dispose()
                            dbCommand.Dispose()
                        Catch ex As Exception
                            MsgBox("No se puede eliminar la base de datos anterior." & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                        End Try

                        functions.rebuildDatabase()

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

                            MsgBox("Restauracion completada!" & vbNewLine & vbNewLine & " Techcare se reiniciará", MsgBoxStyle.Information, "techcare")

                            Application.Restart()
                        Catch ex As Exception
                            MsgBox("Se ha producido un error al restaurar datos de la base de datos." &
                               vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                            functions.log("Restore failed. " & ex.InnerException.Message)
                        End Try
                    End If
                End If
            Else
                MsgBox("MySQL.exe no se encontró en el directorio XAMPP especificado. Por favor verifique que el directorio XAMPP haya sido seleccionado " &
                       "y proba otra vez", MsgBoxStyle.Critical, "techcare")
            End If
        End If
    End Sub

    Private Sub btnResetSoftware_Click(sender As Object, e As EventArgs) Handles btnResetSoftware.Click
        HideAllPanels()

        resetSoftwarePanel.Visible = True
        resetSoftwarePanel.Dock = DockStyle.Fill
        btnResetSoftware.BackColor = Color.FromArgb(32, 129, 197)
    End Sub

    Private Sub btnResetSoftwareAuthentication_Click(sender As Object, e As EventArgs) Handles btnResetSoftwareAuthentication.Click

        Dim empID As Integer = functions.authenticate(tbResetSoftwareUsername.Text, tbResetSoftwarePassword.Text)

        If empID = 0 Then
            MsgBox("¡Se proporcionó un nombre de usuario o contraseña incorrectos!", MsgBoxStyle.Critical, "techcare")
        Else
            If functions.obtainEmployeeDetails(empID, 4) = "Full" Then
                Dim confirmation As DialogResult

                confirmation = MessageBox.Show("Al restablecer techcare se eliminarán TODOS los usuarios, reparaciones, detalles comerciales y configuraciones personalizadas." & vbNewLine & vbNewLine &
                                               "Esta acción NO se puede deshacer. ¿Quieres continuar?", "techcare", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If confirmation = DialogResult.Yes Then
                    Dim accruedErrors As List(Of String) = New List(Of String)
                    Try
                        Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
                        Dim dbCommand As MySqlCommand = New MySqlCommand("DROP DATABASE techcare;", dbConnection)

                        dbConnection.Open()
                        dbCommand.ExecuteNonQuery()
                        dbConnection.Close()
                    Catch ex As Exception
                        accruedErrors.Add("MySQL FATAL ERROR: " & ex.Message)
                    End Try

                    Try
                        My.Settings.dbLocation = ""
                        My.Settings.dbName = ""
                        My.Settings.dbUsername = ""
                        My.Settings.dbPassword = ""
                        My.Settings.userFirstRun = True
                        My.Settings.userAskedForConsent = False
                        My.Settings.userHasConsented = False
                        My.Settings.Save()
                        My.Settings.Reload()
                    Catch ex As Exception
                        accruedErrors.Add("FATAL ERROR: " & ex.Message)
                    End Try

                    If accruedErrors.Count > 0 Then
                        MsgBox("TECHCARE RESET FAILED!" & vbNewLine & accruedErrors.ToString, MsgBoxStyle.Critical, "techcare")
                        Application.Restart()
                    Else
                        MsgBox("¡Reinicio completado! Techcare ahora se reiniciará.", MsgBoxStyle.Information, "techcare")
                        Application.Restart()
                    End If
                End If
            Else
                MsgBox("No se puede restablecer Techcare. " & vbNewLine & vbNewLine & "Sólo los usuarios con acceso de administrador están autorizados a restablecer techcare.",
                       MsgBoxStyle.Critical, "techcare")
            End If
        End If
    End Sub
    Private Sub btnAboutTechcare_Click(sender As Object, e As EventArgs) Handles btnAboutTechcare.Click
        HideAllPanels()

        aboutTechcarePanel.Visible = True
        aboutTechcarePanel.Dock = DockStyle.Fill
        btnAboutTechcare.BackColor = Color.FromArgb(32, 129, 197)

        lblVersionID.Text = "Version " & My.Application.Info.Version.ToString
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        Dim url As String = "https://github.com/OrellanaMatias/"

        Process.Start(url)
    End Sub
End Class