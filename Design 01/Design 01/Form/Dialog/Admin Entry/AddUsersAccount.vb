Imports System.Data.SqlClient

Public Class AddUsersAccount

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox3.Text.Trim <> "" Or TextBox4.Text.Trim <> "" Or TextBox5.Text.Trim <> "" Or ComboBox1.Text <> "" Then

            Dim existing As Boolean

            Call ConnectionNetwork()

            Try
                sql = "SELECT * FROM Users_Tbl WHERE AND Username= '" & TextBox4.Text.Trim & "' AND UserType = '" & ComboBox1.Text & "'"
                cmd = New SqlCommand(sql, cnn)
                cnn.Open()

                reader = cmd.ExecuteReader

                If reader.Read Then

                    If reader.HasRows Then

                        existing = True

                    Else

                        existing = False

                    End If

                End If

            Catch ex As Exception
                MsgBox("error found ;" & ex.Message & ex.StackTrace)
            Finally

                cnn.Close()

                If existing = False Then

                    Try

                        sql = ("INSERT INTO Users_Tbl (Username,Userpassword,Lastname, Firstname, MiddleInitial,UserType,Status) VALUES ('" & TextBox4.Text.Trim & "', '" & TextBox5.Text.Trim & "', '" & TextBox1.Text.Trim & "', '" & TextBox2.Text.Trim & "', '" & TextBox3.Text.Trim + "." & "', '" & "0" & "', '" & "Active" & "')")
                        cmd = New SqlCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()

                        MsgBox("Successfully Added")
                        cnn.Close()
                        cmd.Dispose()

                        sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & Admin.Label1.Text & "', 'Add User Information (' + '" & TextBox4.Text + ")" & "', '" & Date.Now & "')"
                        cmd = New SqlCommand(sql, cnn)
                        cnn.Open()


                        cmd.ExecuteNonQuery()
                        cnn.Close()

                        Me.Dispose()

                    Catch ex As SystemException

                        MsgBox("error " & ex.Message & " " & ex.StackTrace)

                        cnn.Close()

                    Finally

                        cnn.Close()
                        cmd.Dispose()

                    End Try

                Else

                    MsgBox("Account is already registered")

                End If

            End Try

        Else

            MsgBox("Fill all the blanks")

        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox3.Text.Trim <> "" Or TextBox4.Text.Trim <> "" Or TextBox5.Text.Trim <> "" Or ComboBox1.Text <> "" Then

            Dim existing As Boolean

            Call ConnectionNetwork()

            Try
                sql = "SELECT * FROM Users_Tbl WHERE Lastname= '" & TextBox1.Text.Trim & "' AND Firstname= '" & TextBox2.Text.Trim & "' AND MiddleInitial= '" & TextBox3.Text.Trim & "' AND Username= '" & TextBox4.Text.Trim & "'"
                cmd = New SqlCommand(sql, cnn)
                cnn.Open()

                reader = cmd.ExecuteReader

                If reader.Read Then

                    If reader.HasRows Then

                        existing = True

                    Else

                        existing = False

                    End If

                End If

            Catch ex As Exception
                MsgBox("error found ;" & ex.Message & ex.StackTrace)
            Finally

                cnn.Close()

                If existing = False Then

                    Try

                        sql = ("UPDATE Users_Tbl SET Username = '" & TextBox4.Text & "', Userpassword = '" & TextBox5.Text & "', Lastname = '" & TextBox1.Text & "', Firstname = '" & TextBox2.Text & "', MiddleInitial = '" & TextBox3.Text & "', Status = '" & ComboBox2.Text & "' WHERE ID = " & Label7.Text & "")
                        cmd = New SqlCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()

                        MsgBox("Successfully Added")
                        cnn.Close()
                        cmd.Dispose()

                        sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & Admin.Label1.Text & "', 'Update User Information (' + '" & TextBox4.Text + ")" & "', '" & Date.Now & "')"
                        cmd = New SqlCommand(sql, cnn)
                        cnn.Open()


                        cmd.ExecuteNonQuery()
                        cnn.Close()

                        Me.Dispose()

                    Catch ex As SystemException

                        MsgBox("error " & ex.Message & " " & ex.StackTrace)

                        cnn.Close()

                    Finally

                        cnn.Close()
                        cmd.Dispose()

                    End Try

                Else

                    MsgBox("Account is already registered")

                End If

            End Try

        Else

            MsgBox("fill all the blanks")

        End If
    End Sub

    Private Sub AddUsersAccount_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        Me.Opacity = 1
        Dim tmr As New Timer
        tmr.Interval = 1
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity -= 0.1
                                 If Me.Opacity = 0 Then
                                     e.Cancel = False
                                     tmr.Stop()
                                     Me.Dispose()
                                 End If
                             End Sub
    End Sub

    Private Sub AddUsersAccount_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Opacity = 0
        Dim tmr As New Timer
        tmr.Interval = 2
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity += 0.1
                                 If Me.Opacity = 1 Then tmr.Stop()
                             End Sub
    End Sub
End Class