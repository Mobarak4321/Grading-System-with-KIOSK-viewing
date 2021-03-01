Imports System.Data.SqlClient

Public Class Grades

    Dim _ID, _StudentID, _AcademicYear, _Semester, _Subject, _Prelim, _Midterm, _PreFinal, _Final, _Grade, _Remarks As String

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT  *  FROM Grades_Tbl WHERE AcademicYear = '" & Dashboard.TextBox1.Text & "' AND Semester = '" & Dashboard.TextBox2.Text & "' ORDER BY StudentID", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("ID").ToString, reader.Item("StudentID").ToString, reader.Item("Name").ToString, reader.Item("AcademicYear").ToString, reader.Item("Semester").ToString, reader.Item("Professor").ToString, reader.Item("Subject").ToString, reader.Item("Prelim").ToString, reader.Item("Midterm").ToString, reader.Item("PreFinal").ToString, reader.Item("Final").ToString, reader.Item("Grade").ToString, reader.Item("Remarks").ToString, "EDIT")
        End While
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With PrintGrade
            .ShowDialog()
        End With
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim colName As String = DataGridView1.Columns(e.ColumnIndex).Name

        If colName = "Column14" Then
            With UpdateGrades
                .Label12.Text = _ID
                .TextBox1.Text = _StudentID
                .Label9.Text = _AcademicYear
                .Label10.Text = _Semester
                .TextBox2.Text = _Subject
                .TextBox3.Text = _Prelim
                .TextBox4.Text = _Midterm
                .TextBox5.Text = _PreFinal
                .TextBox6.Text = _Final
                .TextBox7.Text = _Grade
                .TextBox8.Text = _Remarks

                .ShowDialog()
            End With
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Dim i As Integer = DataGridView1.CurrentRow.Index

        _ID = DataGridView1.Item(0, i).Value
        _StudentID = DataGridView1.Item(1, i).Value
        _AcademicYear = DataGridView1.Item(3, i).Value
        _Semester = DataGridView1.Item(4, i).Value
        _Subject = DataGridView1.Item(6, i).Value
        _Prelim = DataGridView1.Item(7, i).Value
        _Midterm = DataGridView1.Item(8, i).Value
        _PreFinal = DataGridView1.Item(9, i).Value
        _Final = DataGridView1.Item(10, i).Value
        _Grade = DataGridView1.Item(11, i).Value
        _Remarks = DataGridView1.Item(12, i).Value


    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Try
            Call ConnectionNetwork()

            Dim i As Integer = 0
            Dim cmd2 As SqlCommand
            DataGridView1.Rows.Clear()
            cnn.Open()
            cmd = New SqlCommand("SELECT  *  FROM Grades_Tbl WHERE StudentID LIKE '%" & TextBox1.Text & "%' AND AcademicYear = '" & Dashboard.TextBox1.Text & "' AND Semester = '" & Dashboard.TextBox2.Text & "' ORDER BY AcademicYear, Semester", cnn)
            reader = cmd.ExecuteReader

            While reader.Read
                i += 1
                DataGridView1.Rows.Add(reader.Item("ID").ToString, reader.Item("StudentID").ToString, reader.Item("Name").ToString, reader.Item("AcademicYear").ToString, reader.Item("Semester").ToString, reader.Item("Professor").ToString, reader.Item("Subject").ToString, reader.Item("Prelim").ToString, reader.Item("Midterm").ToString, reader.Item("PreFinal").ToString, reader.Item("Final").ToString, reader.Item("Grade").ToString, reader.Item("Remarks").ToString, "EDIT")
            End While
            reader.Close()
            cnn.Close()

            cnn.Open()
            cmd2 = New SqlCommand("SELECT  Course,YearLevel  FROM Students_Tbl WHERE StudentID LIKE '%" & TextBox1.Text & "%' ", cnn)
            reader = cmd2.ExecuteReader

            While reader.Read
                i += 1
                Label4.Text = reader("Course").ToString
                Label5.Text = reader("YearLevel").ToString
            End While
            reader.Close()
            cnn.Close()
        Catch ex As Exception
            MsgBox("Error Found in :" & ex.Message)
        End Try

    End Sub
End Class