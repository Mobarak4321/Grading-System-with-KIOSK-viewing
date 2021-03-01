Imports System.Data.SqlClient

Public Class SubjectsEnrolled

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With AddSubjectEnrolled
            .ShowDialog()
        End With
    End Sub

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM StudentSubjects_Tbl", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("ID").ToString, reader.Item("StudentID").ToString, reader.Item("Name").ToString, reader.Item("Subject").ToString, reader.Item("Professor").ToString, reader.Item("AcademicYear").ToString, reader.Item("Semester").ToString)
        End While
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        With Students
            .TopLevel = False
            Registrar.Panel2.Controls.Add(Students)
            .BringToFront()
            .RadioButton1.Checked = True
            .Show()
            .LoadRecords()
        End With
        RadioButton1.Checked = False
    End Sub
End Class