Imports System.Data.SqlClient

Public Class AddSubjectEnrolled

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

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        Call ConnectionNetwork()

        Try
            sql = "SELECT * FROM Students_Tbl WHERE StudentID LIKE '%" & TextBox4.Text & "%' "
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()


            reader = cmd.ExecuteReader

            Do While reader.Read = True

                TextBox1.Text = (reader("StudentID").ToString)
                TextBox2.Text = (reader("Lastname").ToString) + ", " + (reader("Firstname").ToString) + " " + (reader("MiddleInitial").ToString)

            Loop

        Catch ex As Exception
            MsgBox("Error Found ; " & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cmd.Dispose()
            cnn.Close()
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call ConnectionNetwork()


        Try
            sql = "SELECT  *  FROM TeacherAdvisories_Tbl WHERE Subject = '" & ComboBox1.Text & "' AND AcademicYear = '" & Dashboard.TextBox1.Text & "' AND Semester = '" & Dashboard.TextBox2.Text & "' AND Laps ='1'"
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Do While reader.Read = True

                TextBox3.Text = (reader("Lastname").ToString) + ", " + (reader("Firstname").ToString) + " " + (reader("MiddleInitial").ToString)

            Loop
        Catch ex As Exception
            MsgBox("Error Found ; " & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox3.Text.Trim <> "" Or ComboBox1.Text.Trim <> "" Then

            Dim existing As Boolean

            Call ConnectionNetwork()

            Try
                sql = "SELECT * FROM StudentSubjects_Tbl WHERE StudentID = '" & TextBox1.Text.Trim & "' AND Subject = '" & ComboBox1.Text & "' AND AcademicYear = '" & Dashboard.TextBox1.Text & "' AND Semester = '" & Dashboard.TextBox2.Text & "' "
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

                        sql = ("INSERT INTO StudentSubjects_Tbl (StudentID, Name, Subject, Professor, AcademicYear, Semester) VALUES ('" & TextBox1.Text.Trim & "', '" & TextBox2.Text.Trim & "', '" & ComboBox1.Text.Trim & "', '" & TextBox3.Text.Trim & "','" & Dashboard.TextBox1.Text.Trim & "','" & Dashboard.TextBox2.Text.Trim & "')")
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

    Private Sub AddSubjectEnrolled_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub AddSubjectEnrolled_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Opacity = 0
        Dim tmr As New Timer
        tmr.Interval = 2
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity += 0.1
                                 If Me.Opacity = 1 Then tmr.Stop()
                             End Sub
        Try

            SubjectsEnrolled.LoadRecords()
            Call ConnectionNetwork()


            Dim command2 As New SqlCommand("SELECT * FROM Subjects_Tbl ", cnn)
            Dim adapter2 As New SqlDataAdapter(command2)

            Dim Table2 As New DataTable

            adapter2.Fill(Table2)

            ComboBox1.DataSource = Table2

            ComboBox1.DisplayMember = "SubjectDescription"
            ComboBox1.ValueMember = "SubjectID"


            cnn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)


        End Try
    End Sub
End Class