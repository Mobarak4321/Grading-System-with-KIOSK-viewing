Imports System.Data.SqlClient

Public Class Teachers

    Dim _EmpID, _Lastname, _Firstname, _MiddleInitial As String
    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Teachers_Tbl", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("EmployeeID").ToString, reader.Item("Lastname").ToString, reader.Item("Firstname").ToString, reader.Item("MiddleInitial").ToString, "EDIT")
        End While
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With AddTeachers
            .ShowDialog()
        End With
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        With TeachersAdvisory
            .TopLevel = False
            Registrar.Panel2.Controls.Add(TeachersAdvisory)
            .BringToFront()
            .RadioButton2.Checked = True
            .Show()
            .LoadRecords()
        End With
        RadioButton2.Checked = False
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Dim i As Integer = DataGridView1.CurrentRow.Index

        _EmpID = DataGridView1.Item(0, i).Value
        _Lastname = DataGridView1.Item(1, i).Value
        _Firstname = DataGridView1.Item(2, i).Value
        _MiddleInitial = DataGridView1.Item(3, i).Value
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim colName As String = DataGridView1.Columns(e.ColumnIndex).Name

        If colName = "Column5" Then
            With AddTeachers
                .Label6.Text = _EmpID
                .TextBox1.Text = _EmpID
                .TextBox2.Text = _Lastname
                .TextBox3.Text = _Firstname
                .TextBox4.Text = _MiddleInitial
                .btnAdd.Hide()
                .btnSave.Show()
                .ShowDialog()
            End With
        End If
    End Sub
End Class