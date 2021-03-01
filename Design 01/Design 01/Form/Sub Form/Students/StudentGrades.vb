Imports System.Data.SqlClient

Public Class StudentGrades

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Grades_Tbl WHERE StudentID = '" & Student.Label1.Text & "' ORDER BY AcademicYear,Semester", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("Subject").ToString, reader.Item("Prelim").ToString, reader.Item("Midterm").ToString, reader.Item("PreFinal").ToString, reader.Item("Final").ToString, reader.Item("Grade").ToString, reader.Item("AcademicYear").ToString, reader.Item("Semester").ToString)
        End While
        reader.Close()
        cnn.Close()
    End Sub

End Class