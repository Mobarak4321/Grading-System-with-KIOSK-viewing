Imports System.Data.SqlClient

Public Class Students

    Dim _StudentID, _Lastname, _Firstname, _MiddleInitial, _Course, _YearLevel As String

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Students_Tbl", cnn)
        reader = cmd.ExecuteReader

            While reader.Read
                i += 1
            DataGridView1.Rows.Add(reader.Item("StudentID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, reader.Item("Course").ToString, reader.Item("YearLevel").ToString, "EDIT ")
            End While
            reader.Close()
            cnn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With AddStudents
            .ShowDialog()
        End With
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Students_Tbl WHERE StudentID LIKE '%" & TextBox1.Text & "%' OR Lastname LIKE '%" & TextBox1.Text & "%' OR Firstname LIKE '%" & TextBox1.Text & "%' OR Course LIKE '%" & TextBox1.Text & "%' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("StudentID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, reader.Item("Course").ToString, reader.Item("YearLevel").ToString, "EDIT ")
        End While

        reader.Close()
        cnn.Close()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        With SubjectsEnrolled
            .TopLevel = False
            Registrar.Panel2.Controls.Add(SubjectsEnrolled)
            .BringToFront()
            .RadioButton2.Checked = True
            .Show()
            .LoadRecords()
        End With
        RadioButton2.Checked = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With PrintStudent
            .ShowDialog()
        End With
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim colName As String = DataGridView1.Columns(e.ColumnIndex).Name

        If colName = "Column7" Then
            With AddStudents
                .Label7.Text = _StudentID
                .TextBox1.Text = _StudentID
                .TextBox2.Text = _Lastname
                .TextBox3.Text = _Firstname
                .TextBox4.Text = _MiddleInitial
                .ComboBox1.Text = _Course
                .ComboBox2.Text = _YearLevel
                .btnAdd.Hide()
                .btnSave.Show()
                .ShowDialog()
            End With
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Dim i As Integer = DataGridView1.CurrentRow.Index

        _StudentID = DataGridView1.Item(0, i).Value
        _Lastname = DataGridView1.Item(1, i).Value
        _Firstname = DataGridView1.Item(2, i).Value
        _MiddleInitial = DataGridView1.Item(3, i).Value
        _Course = DataGridView1.Item(4, i).Value
        _YearLevel = DataGridView1.Item(5, i).Value

    End Sub
End Class