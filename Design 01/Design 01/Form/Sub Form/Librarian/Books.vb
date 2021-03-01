Imports System.Data.SqlClient

Public Class Books

    Dim _BooksID, _Title, _Publisher, _Writter, _Date, _Category, _Status As String

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Books_Tbl", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With AddBooks
            .ShowDialog()
        End With
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim colName As String = DataGridView1.Columns(e.ColumnIndex).Name

        If colName = "Column8" Then
            With AddBooks
                .Label9.Text = _BooksID
                .TextBox1.Text = _BooksID
                .TextBox2.Text = _Title
                .TextBox3.Text = _Publisher
                .TextBox4.Text = _Writter
                .DateTimePicker1.Text = _Date
                .ComboBox1.Text = _Category
                .ComboBox2.Text = _Status
                .ComboBox2.Show()
                .Label8.Show()
                .Panel8.Show()
                .btnAdd.Hide()
                .btnSave.Show()
                .ShowDialog()
            End With
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Dim i As Integer = DataGridView1.CurrentRow.Index

        _BooksID = DataGridView1.Item(0, i).Value
        _Title = DataGridView1.Item(1, i).Value
        _Publisher = DataGridView1.Item(2, i).Value
        _Writter = DataGridView1.Item(3, i).Value
        _Date = DataGridView1.Item(4, i).Value
        _Category = DataGridView1.Item(5, i).Value
        _Status = DataGridView1.Item(6, i).Value
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Books_Tbl WHERE Title LIKE '%" & TextBox1.Text & "%' OR Writter LIKE '%" & TextBox1.Text & "%' OR Category LIKE '%" & TextBox1.Text & "%' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub

    Sub Archived_Status()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Books_Tbl WHERE Status = 'Archived' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        reader.Close()
        cnn.Close()
    End Sub

    Sub Disposed_Status()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Books_Tbl WHERE Status = 'Disposed' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        reader.Close()
        cnn.Close()
    End Sub

    Sub Good_Status()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Books_Tbl WHERE Status = 'Good' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        reader.Close()
        cnn.Close()
    End Sub

    Sub Lost_Status()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Books_Tbl WHERE Status = 'Lost'", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        reader.Close()
        cnn.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "Archived" Then
            Archived_Status()

        ElseIf ComboBox1.Text = "Disposed" Then
            Disposed_Status()

        ElseIf ComboBox1.Text = "Good" Then
            Good_Status()

        ElseIf ComboBox1.Text = "Lost" Then
            Lost_Status()
        End If
    End Sub
End Class