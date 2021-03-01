Imports System.Data.SqlClient

Public Class Login

    Protected Overrides ReadOnly Property createparams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            Return cp
        End Get
    End Property

    Sub Login()
        Call ConnectionNetwork()

        Try

            Dim userFound As Boolean = False
            Dim Lastname As String = ""
            Dim Firstname As String = ""
            Dim Middlename As String = ""

            sql = "SELECT * FROM Users_Tbl WHERE Username=@Username AND Userpassword=@Userpassword AND Status = 'Active' "
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()

            cmd.Parameters.AddWithValue("@Username", TextBox1.Text)
            cmd.Parameters.AddWithValue("@Userpassword", TextBox2.Text)

            reader = cmd.ExecuteReader
            If TextBox1.Text <> "Username" Or TextBox2.Text <> "Password" Then

                If reader.Read Then

                    If reader.HasRows Then

                        UserName = reader.Item(1).ToString

                        If reader.Item(6).ToString = "Admin" Then

                            UserName = reader.Item(0).ToString

                            Admin.Label1.Text = reader.Item(1).ToString
                            Admin.Label2.Text = reader.Item(3).ToString + ", " + reader.Item(4).ToString + " " + reader.Item(5).ToString


                            cnn.Close()

                            sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & textBox1.Text & "', 'Sign In', '" & Date.Now & "')"
                            cmd = New SqlCommand(sql, cnn)
                            cnn.Open()


                            cmd.ExecuteNonQuery()
                            cnn.Close()

                            Admin.Show()
                            Me.Hide()

                        ElseIf reader.Item(6).ToString = "Registrar" Then

                            UserName = reader.Item(0).ToString

                            Registrar.Label1.Text = reader.Item(1).ToString
                            Registrar.Label2.Text = reader.Item(3).ToString + ", " + reader.Item(4).ToString + " " + reader.Item(5).ToString

                            cnn.Close()

                            sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & textBox1.Text & "', 'Sign In', '" & Date.Now & "')"
                            cmd = New SqlCommand(sql, cnn)
                            cnn.Open()


                            cmd.ExecuteNonQuery()
                            cnn.Close()

                            Registrar.Show()
                            Me.Hide()

                        ElseIf reader.Item(6).ToString = "Teacher" Then
                            UserName = reader.Item(0).ToString

                            Teacher.Label3.Text = reader.Item(3).ToString
                            Teacher.Label4.Text = reader.Item(4).ToString
                            Teacher.Label5.Text = reader.Item(5).ToString


                            Teacher.Label2.Text = reader.Item(3).ToString + ", " + reader.Item(4).ToString + " " + reader.Item(5).ToString
                            Teacher.Label1.Text = reader.Item(1).ToString

                            TeacherProfile.TextBox2.Text = reader.Item(3).ToString + ", " + reader.Item(4).ToString + " " + reader.Item(5).ToString
                            TeacherProfile.TextBox1.Text = reader.Item(1).ToString
                            TeacherProfile.TextBox3.Text = reader.Item(1).ToString
                            TeacherProfile.TextBox4.Text = reader.Item(2).ToString

                            cnn.Close()

                            sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & textBox1.Text & "', 'Sign In', '" & Date.Now & "')"
                            cmd = New SqlCommand(sql, cnn)
                            cnn.Open()


                            cmd.ExecuteNonQuery()
                            cnn.Close()

                            TeacherGrade.SchoolYear()
                            Teacher.Show()
                            Me.Hide()

                        ElseIf reader.Item(6).ToString = "Librarian" Then
                            UserName = reader.Item(0).ToString

                            cnn.Close()

                            sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & textBox1.Text & "', 'Sign In', '" & Date.Now & "')"
                            cmd = New SqlCommand(sql, cnn)
                            cnn.Open()


                            cmd.ExecuteNonQuery()
                            cnn.Close()

                            'TeacherGrade.SchoolYear()
                            Librarian.Show()
                            Me.Hide()

                        Else
                            UserName = reader.Item(0).ToString

                            Student.Label2.Text = reader.Item(3).ToString + ", " + reader.Item(4).ToString + " " + reader.Item(5).ToString
                            Student.Label1.Text = reader.Item(1).ToString

                            cnn.Close()

                            sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & textBox1.Text & "', 'Sign In', '" & Date.Now & "')"
                            cmd = New SqlCommand(sql, cnn)
                            cnn.Open()


                            cmd.ExecuteNonQuery()
                            cnn.Close()

                            Student.Show()
                            Me.Hide()

                        End If
                    Else
                        cnn.Close()
                    End If
                Else

                    cnn.Close()
                    MsgBox("Username or Password is incorrect")

                    TextBox1.Text = "Username"
                    TextBox2.Text = "Password"

                End If

            Else

                MsgBox("Please enter your username and password")
                textBox1.Text = "Username"
                textBox2.Text = "Password"
                cnn.Close()

            End If

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles textBox1.Click
        If textBox1.Text = "Username" Then
            textBox1.Clear()
        End If
    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles textBox2.Click
        If textBox2.Text = "Password" Then
            textBox2.Clear()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        Login()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Login_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Opacity = 0
        Dim tmr As New Timer
        tmr.Interval = 2
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity += 0.1
                                 If Me.Opacity = 1 Then tmr.Stop()
                             End Sub
    End Sub

    Private Sub linkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles linkLabel1.LinkClicked
        Process.Start("C:\xampp\xampp_start.exe")

    End Sub
End Class