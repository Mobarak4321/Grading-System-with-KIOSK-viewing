Imports System.Data.SqlClient

Public Class AddStudents

    Protected Overrides ReadOnly Property createparams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            Return cp
        End Get
    End Property

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox3.Text.Trim <> "" Or TextBox4.Text.Trim <> "" Or ComboBox1.Text.Trim <> "" Then

            Dim existing As Boolean

            Call ConnectionNetwork()

            Try
                sql = "SELECT * FROM Students_Tbl WHERE StudentID= '" & TextBox1.Text.Trim & "'"
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

                        sql = ("INSERT INTO Students_Tbl (StudentID, Lastname, Firstname, MiddleInitial, Course, YearLevel) VALUES ('" & TextBox1.Text.Trim & "', '" & TextBox2.Text.Trim & "', '" & TextBox3.Text.Trim & "', '" & TextBox4.Text.Trim + "." & "','" & ComboBox1.Text.Trim & "','" & ComboBox2.Text.Trim & "')")
                        cmd = New SqlCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()

                        MsgBox("Successfully Added")
                        cnn.Close()
                        cmd.Dispose()

                        sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & Registrar.Label1.Text & "', 'Add Student Information (' + '" & TextBox1.Text + ")" & "', '" & Date.Now & "')"
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
                        'cmd.Dispose()
                        'Call Clear()

                    End Try

                Else

                    MsgBox("already registered")

                End If

            End Try

        Else

            MsgBox("fill all the blanks")

        End If
    End Sub

    Private Sub AddStudents_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub AddStudents_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Opacity = 0
        Dim tmr As New Timer
        tmr.Interval = 2
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity += 0.1
                                 If Me.Opacity = 1 Then tmr.Stop()
                             End Sub
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox3.Text.Trim <> "" Or TextBox4.Text.Trim <> "" Or ComboBox1.Text.Trim <> "" Then

            Dim existing As Boolean

            Call ConnectionNetwork()

            Try
                sql = "SELECT * FROM Students_Tbl WHERE StudentID= '" & TextBox1.Text.Trim & "'"
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

                        sql = ("UPDATE Students_Tbl SET StudentID = '" & TextBox1.Text & "', Lastname = '" & TextBox2.Text & "', MiddleInitial = '" & TextBox3.Text & "', Course = '" & ComboBox1.Text & "', YearLevel = '" & ComboBox2.Text & "' WHERE StudentID = '" & Label7.Text & "' ")
                        cmd = New SqlCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()

                        MsgBox("Successfully Saved")
                        cnn.Close()
                        cmd.Dispose()

                        sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & Registrar.Label1.Text & "', 'Update Student Information (' + '" & TextBox1.Text + ")" & "', '" & Date.Now & "')"
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
                        'cmd.Dispose()
                        'Call Clear()

                    End Try

                Else

                    MsgBox("already registered")

                End If

            End Try

        Else

            MsgBox("fill all the blanks")

        End If
    End Sub
End Class