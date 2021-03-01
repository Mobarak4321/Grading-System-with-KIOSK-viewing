Imports System.Data.SqlClient

Public Class Subjects

    Dim _SubjectID, _SubjectCode, _SubjectDescription, Units As String

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Subjects_Tbl", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("SubjectID").ToString, reader.Item("SubjectCode").ToString, reader.Item("SubjectDescription").ToString, reader.Item("Unit").ToString, "EDIT")
        End While
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With AddSubjects
            .ShowDialog()
        End With
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Dim i As Integer = DataGridView1.CurrentRow.Index

        _SubjectID = DataGridView1.Item(0, i).Value
        _SubjectCode = DataGridView1.Item(1, i).Value
        _SubjectDescription = DataGridView1.Item(2, i).Value
        Units = DataGridView1.Item(3, i).Value
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim colName As String = DataGridView1.Columns(e.ColumnIndex).Name

        If colName = "Column5" Then
            With AddSubjects
                .TextBox1.Text = _SubjectID
                .TextBox2.Text = _SubjectCode
                .TextBox3.Text = _SubjectDescription
                .TextBox4.Text = Units
                .btnAdd.Hide()
                .btnSave.Show()
                .ShowDialog()
            End With
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With PrintSubject
            .ShowDialog()
        End With
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Subjects_Tbl WHERE SubjectCode LIKE '%" & TextBox1.Text & "%' OR SubjectDescription LIKE '%" & TextBox1.Text & "%' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("SubjectID").ToString, reader.Item("SubjectCode").ToString, reader.Item("SubjectDescription").ToString, reader.Item("Unit").ToString, "EDIT")
        End While
        reader.Close()
        cnn.Close()
    End Sub
End Class