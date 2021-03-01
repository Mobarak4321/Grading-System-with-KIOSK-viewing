Imports System.Data.SqlClient

Public Class BackupDatabase

    Sub Backup_Database()
        Dim SQLCon As New SqlConnection With {.ConnectionString = "server =" & TextBox1.Text & ";database=" & TextBox2.Text & ";integrated security=SSPI"}

        SaveFileDialog1.FileName = DateAndTime.DateString + " - SchoolDb"
        SaveFileDialog1.Filter = "SQL Server database backup files |*.bak"
        SaveFileDialog1.ShowDialog()

        Dim cmds As New SqlCommand("BACKUP DATABASE SchoolDb TO disk='" & SaveFileDialog1.FileName & "'", SQLCon)
        SQLCon.Open()

        cmds.ExecuteNonQuery()
        SQLCon.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Backup_Database()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Dispose()
    End Sub

    Private Sub BackupDatabase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Opacity = 0
        Dim tmr As New Timer
        tmr.Interval = 2
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity += 0.1
                                 If Me.Opacity = 1 Then tmr.Stop()
                             End Sub
    End Sub
End Class