Imports MySql.Data.MySqlClient
Imports System.IO
Public Class functions
    Public Shared Function databaseCheck(ByVal dbName As String)
        Dim hasDB As Boolean = False

        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT schema_name FROM information_schema.schemata WHERE schema_name = @dbName;", dbConnection)

            dbCommand.Parameters.AddWithValue("@dbName", dbName)

            dbConnection.Open()

            Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

            If dbReader.HasRows Then
                While dbReader.Read
                    hasDB = True
                End While
            Else
                hasDB = False
            End If

            dbConnection.Close()
        Catch ex As Exception
            log("Se produjo un error durante la verificación de la base de datos: " & ex.Message)
            hasDB = False
        End Try

        Return hasDB
    End Function

    Public Shared Sub rebuildDatabase()
        Dim accruedErrors As List(Of String) = New List(Of String)

        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("CREATE DATABASE techcare;", dbConnection)

            My.Settings.dbName = "techcare"
            My.Settings.Save()
            My.Settings.Reload()

            dbConnection.Open()
            dbCommand.ExecuteNonQuery()
            dbConnection.Close()
        Catch ex As Exception
            accruedErrors.Add("No se pudo crear la base de datos techcare en el servidor:" & ex.Message)
            log("Error occurred during db rebuild: " & ex.Message)
        End Try

        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("CREATE TABLE Employees (employeeID INT NOT NULL, title VARCHAR(256) NOT NULL, forename VARCHAR(256) NOT NULL, " &
                                                             "surname VARCHAR(256) NOT NULL, userAccessLevel VARCHAR(256) NOT NULL, username VARCHAR(256) NOT NULL," &
                                                             "password VARCHAR(256) NOT NULL);", dbConnection)
            dbConnection.Open()
            dbCommand.ExecuteNonQuery()
            dbConnection.Close()
        Catch ex As Exception
            accruedErrors.Add("Failed to create Employees table in techcare database: " & ex.Message)
            log("Error occurred during db rebuild: " & ex.Message)
        End Try

        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("CREATE TABLE Repairs (repairReference INT NOT NULL, title VARCHAR(256) NOT NULL, forename VARCHAR(256) NOT NULL, " &
                                                             "surname VARCHAR(256) NOT NULL, address VARCHAR(256) NOT NULL, city VARCHAR(256) NOT NULL, postcode VARCHAR(256) NOT NULL, " &
                                                             "homePhone VARCHAR(256) NULL, mobilePhone VARCHAR(256) NULL, emailAddress VARCHAR(256) NULL, " &
                                                             "assetMake VARCHAR(256) NOT NULL, assetModel VARCHAR(256) NOT NULL, assetSerialNumber VARCHAR(256) NOT NULL, " &
                                                             "currentRepairStatus VARCHAR(256) NOT NULL, intakeDate VARCHAR(256) NOT NULL, faultDescription VARCHAR(256) NOT NULL, " &
                                                             "estimateQuote VARCHAR(256) NOT NULL);", dbConnection)
            dbConnection.Open()
            dbCommand.ExecuteNonQuery()
            dbConnection.Close()
        Catch ex As Exception
            accruedErrors.Add("Failed to create Repairs table in techcare database: " & ex.Message)
            log("Error occurred during db rebuild: " & ex.Message)
        End Try

        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("CREATE TABLE Remarks (commentID VARCHAR(256) NOT NULL, employeeID VARCHAR(256) NOT NULL, jobReference VARCHAR(256) NOT NULL, " &
                                                             "timestamp VARCHAR(256) NOT NULL, comment VARCHAR(256) NOT NULL);", dbConnection)
            dbConnection.Open()
            dbCommand.ExecuteNonQuery()
            dbConnection.Close()
        Catch ex As Exception
            accruedErrors.Add("Failed to create Remarks table in techcare database: " & ex.Message)
            log("Error occurred during db rebuild: " & ex.Message)
        End Try

        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("CREATE TABLE Business (name VARCHAR(256) NOT NULL, address VARCHAR(256) NOT NULL, phoneNumber VARCHAR(256) NOT NULL, " &
                                                             "repairTnC VARCHAR(256) NOT NULL);", dbConnection)

            dbConnection.Open()
            dbCommand.ExecuteNonQuery()
            dbConnection.Close()
        Catch ex As Exception
            accruedErrors.Add("Failed to create Business table in techcare database: " & ex.Message)
            log("Error occurred during db rebuild: " & ex.Message)
        End Try

        If accruedErrors.Count > 0 Then
            Dim errors As String = Nothing
            For i As Integer = 0 To accruedErrors.Count - 1
                errors = errors & vbNewLine & vbNewLine & accruedErrors(i).ToString
            Next
            MsgBox("Se han producido errores al crear la base de datos de techcare. Es posible que Techcare no funcione según lo previsto. Revise estos errores a continuación:" & errors, MsgBoxStyle.Critical, "techcare")
        Else
            If databaseCheck(My.Settings.dbName) = True Then
                log("Reconstrucción exitosa, se ha logrado la conexión.")
            End If
        End If
    End Sub

    Public Shared Function generateUid(ByVal table As String, ByVal tablePk As String, ByVal length As Integer) As String
        Dim hasUniqueIdentifier As Boolean = False
        Dim uniqueIdentifier As String = Nothing

        While hasUniqueIdentifier = False
            Dim possibleChars As String = "1234567890"
            Dim charArray() As Char = possibleChars.ToCharArray
            Dim random As New Random
            Dim sb As New System.Text.StringBuilder

            For index As Integer = 1 To length
                sb.Append(charArray(random.Next(0, charArray.Length)))
            Next

            Dim generatedID As String = sb.ToString

            Try
                Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
                Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT * FROM " & table & " WHERE " & tablePk & " = '" & generatedID & "';", dbConnection)


                dbConnection.Open()
                Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

                If dbReader.HasRows Then
                Else
                    uniqueIdentifier = generatedID
                    hasUniqueIdentifier = True
                End If

                dbConnection.Close()
                dbCommand.Dispose()
                dbConnection.Dispose()
            Catch ex As Exception
                MsgBox("Se ha producido una excepción al generar un identificador único." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
                log("Error occurred during UID creation: " & ex.Message)
            End Try
        End While

        Return uniqueIdentifier
    End Function

    Public Shared Function authenticate(ByVal username As String, ByVal password As String)
        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT employeeID FROM Employees WHERE username=@uid AND password=@pwd", dbConnection)
            Dim validDetails As Integer = 0

            dbConnection.Open()

            dbCommand.Parameters.AddWithValue("@uid", username)
            dbCommand.Parameters.AddWithValue("@pwd", password)

            Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

            If dbReader.HasRows Then
                While dbReader.Read
                    validDetails = Convert.ToInt32(dbReader(0).ToString)
                End While
            End If

            dbConnection.Close()
            dbCommand.Dispose()
            dbConnection.Dispose()

            Return validDetails
        Catch ex As Exception
            MsgBox("Se ha producido una excepción al autenticar al usuario." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
            log("Se produjo un error durante la autenticación del usuario: " & ex.Message)
            Return 0
        End Try
    End Function

    Public Shared Function obtainEmployeeDetails(ByVal employeeID As String, ByVal col As Integer)
        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT * FROM Employees WHERE employeeID=@empID", dbConnection)
            Dim returningValue As String = Nothing

            dbConnection.Open()

            dbCommand.Parameters.AddWithValue("@empID", employeeID)

            Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

            If dbReader.HasRows Then
                While dbReader.Read
                    returningValue = dbReader(col).ToString
                End While
            End If

            dbConnection.Close()
            dbCommand.Dispose()
            dbConnection.Dispose()

            Return returningValue
        Catch ex As Exception
            MsgBox("Se ha producido una excepción al obtener los datos del empleado." & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "techcare")
            log("Error occurred while obtaining employee details: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Shared Function validateNewPassword(ByVal password As String) As Boolean
        Dim upper As New System.Text.RegularExpressions.Regex("[A-Z]")
        Dim lower As New System.Text.RegularExpressions.Regex("[a-z]")
        Dim number As New System.Text.RegularExpressions.Regex("[0-9]")
        Dim special As New System.Text.RegularExpressions.Regex("[^a-zA-Z0-9]")

        If password.Length < 8 Then Return False
        If upper.Matches(password).Count < 1 Then Return False
        If lower.Matches(password).Count < 1 Then Return False
        If number.Matches(password).Count < 1 Then Return False
        If special.Matches(password).Count < 1 Then Return False

        Return True
    End Function

    ' //////////////////////////////////////////////////////////////////////////////////////////////////////////
    ' //////////////////////////////////////////////////////////////////////////////////////////////////////////
    ' //////////////////////////////////////////////////////////////////////////////////////////////////////////
    ' //////////////////////////////////////////////////////////////////////////////////////////////////////////
    ' //////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Shared Function generateUsername(ByVal forename As String, ByVal surname As String)
        Try
            Dim newUsername As String

            surname = surname.Replace(" ", "")
            surname = surname.Replace("-", "")
            surname = surname.Replace("'", "")

            If surname.ToString.Length > 5 Then
                newUsername = surname.ToString.Substring(0, 5)
            Else
                newUsername = surname.ToString
            End If

            newUsername = newUsername + forename.ToString.Substring(0, 1)

            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName & ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT * FROM Employees WHERE username LIKE @uid", dbConnection)
            Dim counter As Integer = 1

            dbCommand.Parameters.AddWithValue("@uid", "%" & newUsername.ToString.Substring(0, Len(newUsername)) & "%")
            dbConnection.Open()

            Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

            If dbReader.HasRows Then
                While dbReader.Read
                    counter += 1
                End While
            End If

            newUsername = newUsername + counter.ToString("00")

            dbConnection.Close()
            dbConnection.Dispose()
            dbCommand.Dispose()

            Return newUsername.ToLower
        Catch ex As Exception
            log("Error occurred during username generation: " & ex.StackTrace)
            Return Nothing
        End Try
    End Function

    Public Shared Function obtainBusinessDetails(ByVal col As Integer)
        Try
            Dim dbConnection As MySqlConnection = New MySqlConnection("Server=" & My.Settings.dbLocation & ";Database=" & My.Settings.dbName &
                                                                      ";Uid=" & My.Settings.dbUsername & ";Pwd=" & My.Settings.dbPassword & ";")
            Dim dbCommand As MySqlCommand = New MySqlCommand("SELECT * FROM Business", dbConnection)
            Dim returningValue As String = Nothing

            dbConnection.Open()

            Dim dbReader As MySqlDataReader = dbCommand.ExecuteReader

            If dbReader.HasRows Then
                While dbReader.Read
                    returningValue = dbReader(col).ToString
                End While
            End If

            dbConnection.Close()
            dbCommand.Dispose()
            dbConnection.Dispose()

            Return returningValue
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Shared Sub log(ByVal infoToLog As String)
        If My.Settings.userHasConsented = True Then
            Dim strFile As String = "log.txt"
            Dim fileExists As Boolean = File.Exists(strFile)
            Using sw As New StreamWriter(File.Open(strFile, FileMode.Append))
                sw.WriteLine(If(fileExists, "[" & DateTime.Now.ToShortDateString & " - " & DateTime.Now.ToShortTimeString & "] " & infoToLog,
                             "====NEW ERROR LOG====" & vbNewLine & My.Computer.Info.OSFullName & vbNewLine & My.Computer.Info.TotalPhysicalMemory &
                             vbNewLine & My.Computer.Name & "====END OF COMPUTER INFO COLLECTION===="))
            End Using
        End If
    End Sub
End Class
