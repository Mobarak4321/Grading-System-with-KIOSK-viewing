Imports System.Data.SqlClient

Public Class UserActivity

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM UsersLog_Tbl ORDER BY Time", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("UserID").ToString, reader.Item("Activity").ToString, reader.Item("Time").ToString)
        End While
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM UsersLog_Tbl WHERE UserID LIKE '%" & TextBox1.Text & "%' ORDER BY Time", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("UserID").ToString, reader.Item("Activity").ToString, reader.Item("Time").ToString)
        End While
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM UsersLog_Tbl WHERE Time = '" & TextBox1.Text & "' ORDER BY Time", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("UserID").ToString, reader.Item("Activity").ToString, reader.Item("Time").ToString)
        End While
        reader.Close()
        cnn.Close()
    End Sub
End Class