Imports System.Data.SqlClient

Public Class TeachersAdvisory

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With AddTeacherAdvisories
            .ShowDialog()
        End With
    End Sub

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM TeacherAdvisories_Tbl WHERE AcademicYear = '" & Dashboard.TextBox1.Text & "' AND Semester = '" & Dashboard.TextBox2.Text & "' AND Laps = '1'", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("AdvisoryID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, reader.Item("Subject").ToString, reader.Item("AcademicYear").ToString, reader.Item("Semester").ToString)
        End While
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        With Teachers
            .TopLevel = False
            Registrar.Panel2.Controls.Add(Teachers)
            .BringToFront()
            .RadioButton1.Checked = True
            .Show()
            .LoadRecords()
            .RadioButton1.Checked = True
        End With
        RadioButton1.Checked = False
    End Sub
End Class