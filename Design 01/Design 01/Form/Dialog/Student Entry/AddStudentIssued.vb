Imports System.Data.SqlClient

Public Class AddStudentIssued

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Call ConnectionNetwork()

        Try

            sql = ("UPDATE Books_Tbl SET Status = 'Reserved' WHERE BooksID = '" & TextBox1.Text & "' ")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
            cmd.Dispose()

            sql = ("INSERT INTO BooksIssued_Tbl (BooksID, Title, MemberID, Name, IssuedDate, Status) VALUES ('" & TextBox1.Text.Trim & "', '" & TextBox2.Text.Trim & "', '" & Student.Label1.Text.Trim & "', '" & Student.Label2.Text.Trim + "." & "','" & Date.Now & "','Reserved' )")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            MsgBox("Successfully")
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub
End Class