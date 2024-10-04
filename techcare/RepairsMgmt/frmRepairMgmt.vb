Imports MySql.Data.MySqlClient
Public Class frmRepairMgmt

    Private Sub frmRepairMgmt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dgvRepairRemarks.Rows.Clear()
        dgvRepairSearchResults.Rows.Clear()
        lblUserPrompt.Show()
        lblUserPrompt.Dock = DockStyle.Fill
    End Sub

    Private Sub btnSearchRepairs_Click(sender As Object, e As EventArgs) Handles btnSearchRepairs.Click
        dgvRepairSearchResults.Rows.Clear()

        If tbRepairRef.TextLength > 0 And tbSurname.TextLength > 0 Then
            MsgBox("Se produjo un error al buscar reparaciones." & vbNewLine & vbNewLine &
                   "Sólo puedes buscar por apellido o por referencia de reparación.", MsgBoxStyle.Exclamation, "techcare")
        Else
            If tbSurname.TextLength > 0 Then
                Try
                    Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                    Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT repairReference, surname, assetMake, assetModel, currentRepairStatus FROM Repairs" &
                                                                     " WHERE surname Like @surname;", dbConnection)

                    dbCommand.Parameters.AddWithValue("@surname", tbSurname.Text & "%")

                    dbConnection.Open()

                    Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

                    If dbReader.HasRows Then
                        While dbReader.Read
                            dgvRepairSearchResults.Rows.Add(New String() {dbReader(0).ToString, dbReader(1).ToString, dbReader(2).ToString & " " & dbReader(3).ToString,
                                dbReader(4).ToString})
                        End While
                    Else
                        MsgBox("No se encontraron resultados. Intente refinar los resultados de su búsqueda.", MsgBoxStyle.Exclamation, "techcare")
                    End If

                    dbConnection.Close()
                    dbCommand.Dispose()
                    dbConnection.Dispose()
                Catch ex As Exception
                    MsgBox("Se ha producido un error al buscar reparaciones en la base de datos de techcare." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                End Try

            ElseIf tbRepairRef.TextLength > 0 Then
                Try
                    Dim repairRefs As List(Of String) = New List(Of String)

                    Try
                        Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                        Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT repairReference FROM Repairs ORDER BY repairReference ASC;", dbConnection)

                        dbConnection.Open()

                        Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

                        If dbReader.HasRows Then
                            While dbReader.Read
                                repairRefs.Add(dbReader(0).ToString)
                            End While
                        End If

                        dbConnection.Close()
                        dbConnection.Dispose()
                        dbCommand.Dispose()
                    Catch ex As Exception
                        MsgBox("Se ha producido un error al buscar reparaciones en la base de datos de techcare." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                    End Try

                    If repairRefs.Count > 0 Then
                        Dim minimum As Integer = 0
                        Dim maximum As Integer = repairRefs.Count
                        Dim guess As Integer
                        Dim position As Integer = -1

                        While repairRefs.Count >= minimum
                            guess = Math.Floor((maximum + minimum) / 2)
                            If repairRefs(guess).ToString = tbRepairRef.Text Then
                                position = guess
                                Exit While
                            ElseIf Convert.ToInt32(repairRefs(guess).ToString) > Convert.ToInt32(tbRepairRef.Text) Then
                                maximum = guess - 1
                            Else
                                minimum = guess + 1
                            End If
                        End While

                        If position = -1 Then
                            MsgBox("No se encontraron resultados. Intente refinar los resultados de su búsqueda.", MsgBoxStyle.Exclamation, "techcare")
                        Else
                            Try
                                Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                                Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT repairReference, surname, assetMake, assetModel, currentRepairStatus FROM Repairs " &
                                                                                 "WHERE repairReference = @repairRef;", dbConnection)

                                dbCommand.Parameters.AddWithValue("@repairRef", repairRefs(position).ToString)

                                dbConnection.Open()

                                Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

                                If dbReader.HasRows Then
                                    While dbReader.Read
                                        dgvRepairSearchResults.Rows.Add(New String() {dbReader(0).ToString, dbReader(1).ToString, dbReader(2).ToString & " " &
                                                                        dbReader(3).ToString, dbReader(4).ToString})
                                    End While
                                End If

                                dbConnection.Close()
                                dbConnection.Dispose()
                                dbCommand.Dispose()
                            Catch ex As Exception
                                MsgBox("Se ha producido un error al buscar reparaciones en la base de datos de techcare." & vbNewLine & vbNewLine & ex.Message,
                                       MsgBoxStyle.Critical, "techcare")
                            End Try
                        End If
                    Else
                        MsgBox("No se encontraron resultados. Intente refinar los resultados de su búsqueda.", MsgBoxStyle.Exclamation, "techcare")
                    End If
                Catch ex As Exception
                    MsgBox("No se encontraron resultados. Intente refinar los resultados de su búsqueda.", MsgBoxStyle.Exclamation, "techcare")
                End Try
            Else
                Try
                    Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                    Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT repairReference, surname, assetMake, assetModel, currentRepairStatus FROM Repairs;", dbConnection)

                    dbConnection.Open()

                    Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

                    If dbReader.HasRows Then
                        While dbReader.Read
                            dgvRepairSearchResults.Rows.Add(New String() {dbReader(0).ToString, dbReader(1).ToString, dbReader(2).ToString & " " & dbReader(3).ToString,
                                                            dbReader(4).ToString})
                        End While
                    Else
                        MsgBox("No existen reparaciones en el sistema.", MsgBoxStyle.Exclamation, "techcare")
                    End If

                    dbConnection.Close()
                    dbCommand.Dispose()
                    dbConnection.Dispose()
                Catch ex As Exception
                    MsgBox("Se ha producido un error al buscar reparaciones en la base de datos de techcare." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                End Try
            End If
        End If
    End Sub

    Private Sub dgvRepairSearchResults_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRepairSearchResults.CellClick

        If dgvRepairSearchResults.SelectedRows.Count = 1 Then
            Try
                Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT * FROM Repairs WHERE repairReference = @repairRef;", dbConnection)

                dbCommand.Parameters.AddWithValue("@repairRef", dgvRepairSearchResults.SelectedRows(0).Cells(0).Value.ToString)

                dbConnection.Open()
                Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

                If dbReader.HasRows Then
                    While dbReader.Read
                        lblRepairRef.Text = dbReader(0).ToString
                        lblCustomerNameAddress.Text = dbReader(1).ToString & " " & dbReader(2).ToString & " " & dbReader(3).ToString & vbNewLine &
                            dbReader(4).ToString & vbNewLine & dbReader(5).ToString & vbNewLine & dbReader(6).ToString
                        lblHomePhn.Text = dbReader(7).ToString
                        lblMobilePhn.Text = dbReader(8).ToString
                        lblEmailAddress.Text = dbReader(9).ToString
                        lblAssetDetails.Text = dbReader(10).ToString & vbNewLine & dbReader(11).ToString & vbNewLine & dbReader(12).ToString
                        lblCurrentRepairStatus.Text = dbReader(13).ToString
                        lblIntakeDate.Text = dbReader(14).ToString
                        lblFaultDesc.Text = dbReader(15).ToString
                        lblQuote.Text = dbReader(16).ToString
                    End While
                End If

                If lblCurrentRepairStatus.Text = "Guardado" Or lblCurrentRepairStatus.Text = "Removido" Then
                    btnAddRepairRemark.Enabled = False
                    btnUpdateRepairQuote.Enabled = False
                    btnUpdateRepairStatus.Enabled = False
                    btnCustomerCollection.Enabled = False
                Else
                    btnAddRepairRemark.Enabled = True
                    btnUpdateRepairQuote.Enabled = True
                    btnUpdateRepairStatus.Enabled = True
                    btnCustomerCollection.Enabled = True
                End If

                dbConnection.Close()
                dbCommand.Dispose()
                dbConnection.Dispose()
            Catch ex As Exception
                MsgBox("Se ha producido un error al buscar reparaciones en la base de datos de techcare." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
            End Try

            Try
                dgvRepairRemarks.Rows.Clear()

                Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT * FROM Remarks WHERE jobReference = @repairRef ORDER BY timestamp ASC;", dbConnection)

                dbCommand.Parameters.AddWithValue("@repairRef", dgvRepairSearchResults.SelectedRows(0).Cells(0).Value.ToString)

                dbConnection.Open()

                Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

                If dbReader.HasRows Then
                    While dbReader.Read
                        dgvRepairRemarks.Rows.Add(New String() {functions.obtainEmployeeDetails(dbReader(1).ToString, 2) & " " &
                                                  functions.obtainEmployeeDetails(dbReader(1).ToString, 3), dbReader(3).ToString, dbReader(4).ToString})
                    End While
                End If

                dbConnection.Close()
                dbCommand.Dispose()
                dbConnection.Dispose()
            Catch ex As Exception
                MsgBox("Se ha producido un error al buscar los comentarios de la reparación seleccionada en la base de datos de techcare." & vbNewLine & vbNewLine & ex.Message,
                       MsgBoxStyle.Critical, "techcare")
            End Try

            lblUserPrompt.Hide()
            lblUserPrompt.Dock = DockStyle.None
        End If
    End Sub

    Private Sub btnUpdateRepairStatus_Click(sender As Object, e As EventArgs) Handles btnUpdateRepairStatus.Click
        If lblCurrentRepairStatus.Text = "Registrado" Then
            frmUpdateRepairStatus.rbBookedIn.Checked = True
        ElseIf lblCurrentRepairStatus.Text = "Servicio en progeso" Then
            frmUpdateRepairStatus.rbServiceInProgress.Checked = True
        ElseIf lblCurrentRepairStatus.Text = "Transferido a un tercero" Then
            frmUpdateRepairStatus.rbTransferred.Checked = True
        ElseIf lblCurrentRepairStatus.Text = "Servicio completado" Then
            frmUpdateRepairStatus.rbServiceCompleted.Checked = True
        End If
        frmUpdateRepairStatus.ShowDialog()
    End Sub

    Private Sub btnUpdateRepairQuote_Click(sender As Object, e As EventArgs) Handles btnUpdateRepairQuote.Click
        frmUpdateQuote.tbNewRepairQuote.Text = lblQuote.Text
        frmUpdateQuote.ShowDialog()
    End Sub

    Private Sub btnAddRepairRemark_Click(sender As Object, e As EventArgs) Handles btnAddRepairRemark.Click
        frmAddRepairRemark.ShowDialog()
    End Sub

    Private Sub btnCustomerCollection_Click(sender As Object, e As EventArgs) Handles btnCustomerCollection.Click

        Dim confirmCollection As DialogResult
        confirmCollection = MessageBox.Show("Entrega al Cliente:" & vbNewLine & vbNewLine & "¿Ha/está el cliente:" & vbNewLine & "1) Pagado por alguna(s) pieza(s) o mano de obra?" &
                                          vbNewLine & "2) Recibido su producto y está satisfecho con la condición?" & vbNewLine & "3) Satisfecho de que la reparación está completa?" &
                                          vbNewLine & vbNewLine & "Haga clic en SÍ para confirmar la entrega, o haga clic en NO para detener el proceso.",
                                          "techcare", MessageBoxButtons.YesNo, MessageBoxIcon.Question)


        If confirmCollection = DialogResult.Yes Then
            Try
                Dim dbConnection As MySqlConnection = New MySqlConnection("Server=localhost;Database=techcare;Uid=techcare;Pwd=techcare;")
                Dim dbCommand As MySqlCommand = New MySqlCommand("UPDATE Repairs SET currentRepairStatus='Collected' WHERE repairReference=@repairRef;", dbConnection)

                dbCommand.Parameters.AddWithValue("@repairRef", lblRepairRef.Text)

                dbConnection.Open()

                dbCommand.ExecuteNonQuery()

                dbConnection.Close()
                dbCommand.Dispose()
                dbConnection.Dispose()

                MsgBox("Reparacion ha sido marcada como registrada", MsgBoxStyle.Information, "techcare")

                lblUserPrompt.Dock = DockStyle.Fill
                lblUserPrompt.Visible = True
            Catch ex As Exception
                MsgBox("Se ha producido un error al completar la recogida al cliente para esta reparación." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
            End Try
        End If
    End Sub

    Private Sub btnNewRepair_Click(sender As Object, e As EventArgs) Handles btnNewRepair.Click
        frmCreateNewRepair.MdiParent = frmMainWindow
        frmCreateNewRepair.Show()
    End Sub

End Class