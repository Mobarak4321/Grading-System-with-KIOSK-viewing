Imports System.Data.SqlClient

Public Class StudentIssued

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM BooksIssued_Tbl WHERE MemberID = '" & Student.Label1.Text & "' ORDER By IssuedDate", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("IssueNo").ToString, reader.Item("BooksID").ToString, reader.Item("Title").ToString, reader.Item("IssuedDate").ToString, reader.Item("ReturnedDate").ToString, reader.Item("Status").ToString)
        End While

        'Label2.Text = "(" & DataGridView1.RowCount & ") record(s) found."
        reader.Close()
        cnn.Close()
    End Sub
End Class