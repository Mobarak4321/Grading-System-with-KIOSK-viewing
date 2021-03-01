Imports System.Data.SqlClient

Public Class Courses

    Dim _CourseID, _CourseCode, _CourseDescription, _Status As String

    Sub LoadRecords()
        Call ConnectionNetwork()

        Dim i As Integer = 0
        DataGridView1.Rows.Clear()
        cnn.Open()
        cmd = New SqlCommand("SELECT * FROM Courses_Tbl", cnn)
        reader = cmd.ExecuteReader

        While reader.Read
            i += 1
            DataGridView1.Rows.Add(reader.Item("CourseID").ToString, reader.Item("CourseCode").ToString, reader.Item("CourseDescription").ToString, reader.Item("Status").ToString, "EDIT")
        End While
        reader.Close()
        cnn.Close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With AddCourses
            .ShowDialog()
        End With
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim colName As String = DataGridView1.Columns(e.ColumnIndex).Name

        If colName = "Column5" Then
            With AddCourses
                .TextBox1.Text = _CourseID
                .TextBox2.Text = _CourseCode
                .TextBox3.Text = _CourseDescription
                .ComboBox2.Text = _Status
                .ComboBox2.Show()
                .Label5.Show()
                .Panel5.Show()
                .btnAdd.Hide()
                .btnSave.Show()
                .ShowDialog()
            End With
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Dim i As Integer = DataGridView1.CurrentRow.Index

        _CourseID = DataGridView1.Item(0, i).Value
        _CourseCode = DataGridView1.Item(1, i).Value
        _CourseDescription = DataGridView1.Item(2, i).Value
        _Status = DataGridView1.Item(3, i).Value

    End Sub
End Class