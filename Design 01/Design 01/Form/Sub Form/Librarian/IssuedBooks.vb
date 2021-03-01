Imports System.Data.SqlClient

Public Class IssuedBooks

    Dim _No, _BooksID, _Title, _MemberID, _Name, _IssuedDate, _ReturnedDate, _Status As String

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM BooksIssued_Tbl WHERE Status = 'Reserved' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("IssueNo").ToString, reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("MemberID").ToString, reader.Item("Name").ToString, reader.Item("IssuedDate").ToString, reader.Item("ReturnedDate").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        'Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub

    Sub Cancel()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM BooksIssued_Tbl WHERE Status = 'Cancel' ORDER BY IssuedDate ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("IssueNo").ToString, reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("MemberID").ToString, reader.Item("Name").ToString, reader.Item("IssuedDate").ToString, reader.Item("ReturnedDate").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        'Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub

    Sub Damaged()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM BooksIssued_Tbl WHERE Status = 'Damaged' ORDER BY IssuedDate ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("IssueNo").ToString, reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("MemberID").ToString, reader.Item("Name").ToString, reader.Item("IssuedDate").ToString, reader.Item("ReturnedDate").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        'Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub

    Sub Lost()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM BooksIssued_Tbl WHERE Status = 'Lost' ORDER BY IssuedDate ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("IssueNo").ToString, reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("MemberID").ToString, reader.Item("Name").ToString, reader.Item("IssuedDate").ToString, reader.Item("ReturnedDate").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        'Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub

    Sub Pending()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM BooksIssued_Tbl WHERE Status = 'Pending' ORDER BY IssuedDate ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("IssueNo").ToString, reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("MemberID").ToString, reader.Item("Name").ToString, reader.Item("IssuedDate").ToString, reader.Item("ReturnedDate").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        'Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub

    Sub Returned()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM BooksIssued_Tbl WHERE Status = 'Returned' ORDER BY IssuedDate ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("IssueNo").ToString, reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("MemberID").ToString, reader.Item("Name").ToString, reader.Item("IssuedDate").ToString, reader.Item("ReturnedDate").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        'Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim colName As String = DataGridView1.Columns(e.ColumnIndex).Name
            If colName = "Column9" Then
                With AddIssued
                    .TextBox1.Text = _No
                    .TextBox2.Text = _BooksID
                    .TextBox3.Text = _Title
                    .TextBox4.Text = _MemberID
                    .TextBox5.Text = _Name
                    .DateTimePicker1.Text = _IssuedDate
                    .DateTimePicker2.Text = _ReturnedDate
                    .ComboBox1.Text = _Status
                    .btnSave.Show()
                    .ShowDialog()
                End With
            End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Dim i As Integer = DataGridView1.CurrentRow.Index

        _No = DataGridView1.Item(0, i).Value
        _BooksID = DataGridView1.Item(1, i).Value
        _Title = DataGridView1.Item(2, i).Value
        _MemberID = DataGridView1.Item(3, i).Value
        _Name = DataGridView1.Item(4, i).Value
        _IssuedDate = DataGridView1.Item(5, i).Value
        _ReturnedDate = DataGridView1.Item(6, i).Value
        _Status = DataGridView1.Item(7, i).Value
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "Cancel" Then
            Cancel()
        ElseIf ComboBox1.Text = "Damaged" Then
            Damaged()
        ElseIf ComboBox1.Text = "Lost" Then
            Lost()
        ElseIf ComboBox1.Text = "Pending" Then
            Pending()
        ElseIf ComboBox1.Text = "Returned" Then
            Returned()
        End If
    End Sub
End Class