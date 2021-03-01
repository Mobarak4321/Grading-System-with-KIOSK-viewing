Imports System.Data.SqlClient

Public Class AddBooks

    Private Sub AddBooks_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Opacity = 0
        Dim tmr As New Timer
        tmr.Interval = 2
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity += 0.1
                                 If Me.Opacity = 1 Then tmr.Stop()
                             End Sub
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox3.Text.Trim <> "" Or TextBox4.Text.Trim <> "" Or ComboBox1.Text.Trim <> "" Then

            Dim existing As Boolean

            Call ConnectionNetwork()

            Try
                sql = "SELECT * FROM Books_Tbl WHERE BooksID= '" & TextBox1.Text.Trim & "'"
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

                        sql = ("INSERT INTO Books_Tbl (BooksID, Title, Publisher, Writter, Date, Category, Status) VALUES ('" & TextBox1.Text.Trim & "', '" & TextBox2.Text.Trim & "', '" & TextBox3.Text.Trim & "', '" & TextBox4.Text.Trim + "." & "','" & DateTimePicker1.Text.Trim & "','" & ComboBox1.Text.Trim & "', 'Good' )")
                        cmd = New SqlCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()

                        MsgBox("Successfully Added")
                        cnn.Close()
                        cmd.Dispose()

                        'sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & Registrar.Label1.Text & "', 'Add Student Information (' + '" & TextBox1.Text + ")" & "', '" & Date.Now & "')"
                        'cmd = New SqlCommand(sql, cnn)
                        'cnn.Open()


                        'cmd.ExecuteNonQuery()
                        'cnn.Close()

                        Me.Dispose()

                    Catch ex As SystemException

                        MsgBox("error " & ex.Message & " " & ex.StackTrace)

                        cnn.Close()

                    Finally

                        cnn.Close()
                        'cmd.Dispose()
                        'Call Clear()

                    End Try

                Else

                    MsgBox("Books are already registered.")

                End If

            End Try

        Else

            MsgBox("fill all the blanks")

        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox3.Text.Trim <> "" Or TextBox4.Text.Trim <> "" Or ComboBox1.Text.Trim <> "" Then

            Dim existing As Boolean

            Call ConnectionNetwork()

            Try
                sql = "SELECT * FROM Books_Tbl WHERE BooksID= '" & TextBox1.Text.Trim & "'"
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

                        sql = ("UPDATE Books_Tbl SET BooksID = '" & TextBox1.Text & "', Title = '" & TextBox2.Text & "', Publisher = '" & TextBox3.Text & "', Writter = '" & TextBox4.Text & "', Date = '" & DateTimePicker1.Text & "', Category = '" & ComboBox1.Text & "', Status = '" & ComboBox2.Text & "' WHERE BooksID = '" & Label9.Text & "' ")
                        cmd = New SqlCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()

                        DataSaved.ShowDialog()
                        cnn.Close()
                        cmd.Dispose()

                        'sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & Registrar.Label1.Text & "', 'Add Student Information (' + '" & TextBox1.Text + ")" & "', '" & Date.Now & "')"
                        'cmd = New SqlCommand(sql, cnn)
                        'cnn.Open()


                        'cmd.ExecuteNonQuery()
                        'cnn.Close()

                        Me.Dispose()

                    Catch ex As SystemException

                        MsgBox("error " & ex.Message & " " & ex.StackTrace)

                        cnn.Close()

                    Finally

                        cnn.Close()
                        'cmd.Dispose()
                        'Call Clear()

                    End Try

                Else

                    MsgBox("Books are already registered.")

                End If

            End Try

        Else

            MsgBox("fill all the blanks")

        End If
    End Sub
End Class