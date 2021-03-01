Imports System.Data.SqlClient

Public Class AddIssued

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If ComboBox1.Text = "Cancel" Then
            Cancel_Status()
        ElseIf ComboBox1.Text = "Damaged" Then
            Damaged_Status()
        ElseIf ComboBox1.Text = "Lost" Then
            Lost_Status()
        ElseIf ComboBox1.Text = "Returned" Then
            Returned_Status()
        ElseIf ComboBox1.Text = "Pending" Then
            Pending_Status()
        End If
    End Sub

    Sub Cancel_Status()
        Call ConnectionNetwork()

        Try

            sql = ("UPDATE Books_Tbl SET Status = 'Good' WHERE BooksID = '" & TextBox2.Text & "' ")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
            cmd.Dispose()

            sql = ("UPDATE BooksIssued_Tbl SET Status = 'Cancel' WHERE IssueNo = '" & TextBox1.Text & "'")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            DataSaved.ShowDialog()
            cnn.Close()
            cmd.Dispose()

            Me.Dispose()

        Catch ex As SystemException

            MsgBox("error " & ex.Message & " " & ex.StackTrace)
            cnn.Close()

        Finally

            cnn.Close()

        End Try
    End Sub

    Sub Damaged_Status()
        Call ConnectionNetwork()

        Try

            sql = ("UPDATE Books_Tbl SET Status = 'Good' WHERE BooksID = '" & TextBox2.Text & "' ")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
            cmd.Dispose()

            sql = ("UPDATE BooksIssued_Tbl SET ReturnedDate = '" & Date.Now & "', Status = 'Damaged' WHERE IssueNo = '" & TextBox1.Text & "'")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            DataSaved.ShowDialog()
            cnn.Close()
            cmd.Dispose()

            Me.Dispose()

        Catch ex As SystemException

            MsgBox("error " & ex.Message & " " & ex.StackTrace)
            cnn.Close()

        Finally

            cnn.Close()

        End Try
    End Sub

    Sub Lost_Status()
        Call ConnectionNetwork()

        Try

            sql = ("UPDATE Books_Tbl SET Status = 'Lost' WHERE BooksID = '" & TextBox2.Text & "' ")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
            cmd.Dispose()

            sql = ("UPDATE BooksIssued_Tbl SET , Status = 'Lost' WHERE IssueNo = '" & TextBox1.Text & "'")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            DataSaved.ShowDialog()
            cnn.Close()
            cmd.Dispose()

            Me.Dispose()

        Catch ex As SystemException

            MsgBox("error " & ex.Message & " " & ex.StackTrace)
            cnn.Close()

        Finally

            cnn.Close()

        End Try
    End Sub

    Sub Returned_Status()
        Call ConnectionNetwork()

        Try

            sql = ("UPDATE Books_Tbl SET Status = 'Good' WHERE BooksID = '" & TextBox2.Text & "' ")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
            cmd.Dispose()

            sql = ("UPDATE BooksIssued_Tbl SET ReturnedDate = '" & Date.Now & "', Status = 'Returned' WHERE IssueNo = '" & TextBox1.Text & "'")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            DataSaved.ShowDialog()
            cnn.Close()
            cmd.Dispose()

            Me.Dispose()

        Catch ex As SystemException

            MsgBox("error " & ex.Message & " " & ex.StackTrace)
            cnn.Close()

        Finally

            cnn.Close()

        End Try
    End Sub

    Sub Pending_Status()
        Call ConnectionNetwork()

        Try

            sql = ("UPDATE Books_Tbl SET Status = 'Borrow' WHERE BooksID = '" & TextBox2.Text & "' ")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
            cmd.Dispose()

            sql = ("UPDATE BooksIssued_Tbl SET Status = 'Pending' WHERE IssueNo = '" & TextBox1.Text & "'")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            DataSaved.ShowDialog()
            cnn.Close()
            cmd.Dispose()

            Me.Dispose()

        Catch ex As SystemException

            MsgBox("error " & ex.Message & " " & ex.StackTrace)
            cnn.Close()

        Finally

            cnn.Close()

        End Try
    End Sub
End Class