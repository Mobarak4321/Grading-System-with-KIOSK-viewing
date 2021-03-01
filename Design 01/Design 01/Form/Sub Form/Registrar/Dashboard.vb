Imports System.Data.SqlClient

Public Class Dashboard
    Sub AcademicYear()
        Call ConnectionNetwork()

        Try
            sql = "SELECT  *  FROM AcademicYear_Tbl "
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim Year As Boolean = False

            Dim AcademicYear As String = ""
            Dim Semester As String = ""

            While reader.Read

                Year = True

                AcademicYear = reader("AcademicYear").ToString
                Label12.Text = reader("AcademicYear").ToString
                Semester = reader("Semester").ToString

            End While

            If Year = True Then

                TextBox1.Text = AcademicYear
                TextBox2.Text = Semester
                cnn.Close()

            End If
        Catch ex As Exception
            MsgBox("Error Found ; " & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try

    End Sub

    Private Sub Dashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AcademicYear()
        Count_Students()
        Count_Teachers()
        Count_Subjects()
        Count_Courses()
    End Sub

    Sub Count_Students()
        Try
            Call ConnectionNetwork()

            Dim command As New SqlCommand("SELECT COUNT('StudentID') FROM Students_Tbl ", cnn)
            cnn.Open()

            Label2.Text = command.ExecuteScalar().ToString()

            cnn.Close()
        Catch ex As Exception
            MsgBox("Error Found: " & ex.Message)
        End Try

    End Sub

    Sub Count_Teachers()
        Try
            Call ConnectionNetwork()

            Dim command As New SqlCommand("SELECT COUNT('EmployeeID') FROM Teachers_Tbl ", cnn)
            cnn.Open()

            Label3.Text = command.ExecuteScalar().ToString()

            cnn.Close()
        Catch ex As Exception
            MsgBox("Error Found: " & ex.Message)
        End Try

    End Sub

    Sub Count_Subjects()
        Try
            Call ConnectionNetwork()

            Dim command As New SqlCommand("SELECT COUNT('SubjectID') FROM Subjects_Tbl ", cnn)
            cnn.Open()

            Label5.Text = command.ExecuteScalar().ToString()

            cnn.Close()
        Catch ex As Exception
            MsgBox("Error Found: " & ex.Message)
        End Try

    End Sub

    Sub Count_Courses()
        Try
            Call ConnectionNetwork()

            Dim command As New SqlCommand("SELECT COUNT('CourseID') FROM Courses_Tbl ", cnn)
            cnn.Open()

            Label7.Text = command.ExecuteScalar().ToString()

            cnn.Close()
        Catch ex As Exception
            MsgBox("Error Found: " & ex.Message)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text <> "" Then

            Call ConnectionNetwork()

            Try
                sql = ("UPDATE AcademicYear_Tbl SET AcademicYear= '" & TextBox1.Text.Trim & "', Semester= '" & TextBox2.Text.Trim & "' WHERE AcademicYear= '" & Label12.Text & "'")
                cmd = New SqlCommand(sql, cnn)
                cnn.Open()

                cmd.ExecuteNonQuery()

                MsgBox("Successfully Updated")
                cnn.Close()
                cmd.Dispose()

            Catch ex As Exception
                MsgBox("error found ;" & ex.Message & ex.StackTrace)
            Finally
                cnn.Close()
            End Try

        End If
    End Sub
End Class