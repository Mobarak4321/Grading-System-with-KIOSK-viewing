Imports System.Data.SqlClient

Public Class UsersAccounts

    Dim _ID, _Lastname, _Firstname, _MiddleInitial, _Username, _Password, _Status As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With AddUsersAccount
            .ShowDialog()
        End With
    End Sub

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Users_Tbl WHERE UserType = 'Admin' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("ID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, reader.Item("Username").ToString, reader.Item("UserPassword").ToString, reader.Item("Status").ToString, "EDIT")
        End While
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Users_Tbl WHERE Username LIKE '%" & TextBox1.Text & "%' OR Lastname LIKE '%" & TextBox1.Text & "%' OR Firstname LIKE '%" & TextBox1.Text & "%' ", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("ID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, reader.Item("Username").ToString, reader.Item("UserPassword").ToString, reader.Item("Status").ToString, "EDIT")
        End While

        reader.Close()
        cnn.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If ComboBox1.Text = "Admin" Then
            Call ConnectionNetwork()

            Dim i As Integer = 0
            DataGridView1.Rows.Clear()
            cnn.Open()
            cmd = New SqlCommand("SELECT * FROM Users_Tbl WHERE UserType = 'Admin' ", cnn)
            reader = cmd.ExecuteReader

            While reader.Read
                i += 1
                DataGridView1.Rows.Add(reader.Item("ID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, reader.Item("Username").ToString, reader.Item("UserPassword").ToString, reader.Item("Status").ToString, "EDIT")
            End While

            reader.Close()
            cnn.Close()

        ElseIf ComboBox1.Text = "Librarian" Then
            Call ConnectionNetwork()

            Dim i As Integer = 0
            DataGridView1.Rows.Clear()
            cnn.Open()
            cmd = New SqlCommand("SELECT * FROM Users_Tbl WHERE UserType = 'Librarian' ", cnn)
            reader = cmd.ExecuteReader

            While reader.Read
                i += 1
                DataGridView1.Rows.Add(reader.Item("ID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, reader.Item("Username").ToString, reader.Item("UserPassword").ToString, reader.Item("Status").ToString, "EDIT")
            End While

            reader.Close()
            cnn.Close()

        ElseIf ComboBox1.Text = "Registrar" Then
            Call ConnectionNetwork()

            Dim i As Integer = 0
            DataGridView1.Rows.Clear()
            cnn.Open()
            cmd = New SqlCommand("SELECT * FROM Users_Tbl WHERE UserType = 'Registrar' ", cnn)
            reader = cmd.ExecuteReader

            While reader.Read
                i += 1
                DataGridView1.Rows.Add(reader.Item("ID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, reader.Item("Username").ToString, reader.Item("UserPassword").ToString, reader.Item("Status").ToString, "EDIT")
            End While

            reader.Close()
            cnn.Close()

        ElseIf ComboBox1.Text = "Teacher" Then

            Call ConnectionNetwork()

            Dim i As Integer = 0
            DataGridView1.Rows.Clear()
            cnn.Open()
            cmd = New SqlCommand("SELECT * FROM Users_Tbl WHERE UserType = 'Teacher' ", cnn)
            reader = cmd.ExecuteReader

            While reader.Read
                i += 1
                DataGridView1.Rows.Add(reader.Item("ID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, reader.Item("Username").ToString, reader.Item("UserPassword").ToString, reader.Item("Status").ToString, "EDIT")
            End While

            reader.Close()
            cnn.Close()

        ElseIf ComboBox1.Text = "Student" Then

            Call ConnectionNetwork()

            Dim i As Integer = 0
            DataGridView1.Rows.Clear()
            cnn.Open()
            cmd = New SqlCommand("SELECT * FROM Users_Tbl WHERE UserType = 'Student' ", cnn)
            reader = cmd.ExecuteReader

            While reader.Read
                i += 1
                DataGridView1.Rows.Add(reader.Item("ID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, reader.Item("Username").ToString, reader.Item("UserPassword").ToString, reader.Item("Status").ToString, "EDIT")
            End While

            reader.Close()
            cnn.Close()

        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim colName As String = DataGridView1.Columns(e.ColumnIndex).Name

        If colName = "Column8" Then
            With AddUsersAccount
                .Label7.Text = _ID
                .TextBox1.Text = _Lastname
                .TextBox2.Text = _Firstname
                .TextBox3.Text = _MiddleInitial
                .TextBox4.Text = _Username
                .TextBox5.Text = _Password
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

        _ID = DataGridView1.Item(0, i).Value
        _Lastname = DataGridView1.Item(1, i).Value
        _Firstname = DataGridView1.Item(2, i).Value
        _MiddleInitial = DataGridView1.Item(3, i).Value
        _Username = DataGridView1.Item(4, i).Value
        _Password = DataGridView1.Item(5, i).Value
        _Status = DataGridView1.Item(6, i).Value
    End Sub

End Class