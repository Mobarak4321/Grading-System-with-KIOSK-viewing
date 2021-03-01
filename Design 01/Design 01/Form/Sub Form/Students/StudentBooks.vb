Imports System.Data.SqlClient

Public Class StudentBooks

    Dim _BooksID, _Title, _Publisher, _Writter, _Date, _Category As String

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Books_Tbl WHERE Status = 'Good' OR Status = 'Borrow' OR Status = 'Reserved' ORDER BY Title, Writter", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString, "BORROW")
        End While

        Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Books_Tbl WHERE Status = 'Good' AND Title LIKE '%" & TextBox1.Text & "%' OR Publisher LIKE '%" & TextBox1.Text & "%' OR Category LIKE '%" & TextBox1.Text & "%' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString, "BORROW")
        End While

        Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Dim i As Integer = DataGridView1.CurrentRow.Index

        _BooksID = DataGridView1.Item(0, i).Value
        _Title = DataGridView1.Item(1, i).Value
        _Publisher = DataGridView1.Item(2, i).Value
        _Writter = DataGridView1.Item(3, i).Value
        _Date = DataGridView1.Item(4, i).Value
        _Category = DataGridView1.Item(5, i).Value
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim colName As String = DataGridView1.Columns(e.ColumnIndex).Name

        If colName = "Column8" Then
            With AddStudentIssued
                .TextBox1.Text = _BooksID
                .TextBox2.Text = _Title
                .TextBox3.Text = _Publisher
                .TextBox4.Text = _Writter
                .DateTimePicker1.Text = _Date
                .ComboBox1.Text = _Category
                .btnSave.Show()
                .ShowDialog()
            End With
        End If
    End Sub
End Class