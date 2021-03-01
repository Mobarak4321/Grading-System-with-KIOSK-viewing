Imports System.Data.SqlClient

Public Class ReportBooks

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Books_Tbl WHERE Status = 'Good' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString)
        End While

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
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString)
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
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString)
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
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString)
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
            DataGridView1.Rows.Add(reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("Publisher").ToString, reader.Item("Writter").ToString, reader.Item("Date").ToString, reader.Item("Category").ToString, reader.Item("Status").ToString)
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