Imports System.Data.SqlClient

Public Class RestoreDatabase

    Sub Restore_Database()


        Dim servername As String = TextBox1.Text
        Dim databaseaname As String = TextBox2.Text

        Dim con As SqlConnection = New SqlConnection("Data Source =" + servername + ";Initial Catalog=" + databaseaname + ";Integrated security=True")

        con.Open()

        Dim str As String = "USE master;"
        Dim str1 As String = "ALTER DATABASE " + databaseaname + " SET SINGLE_USER WITH ROLLBACK IMMEDIATE;"
        Dim str3 As String = "RESTORE DATABASE " + databaseaname + " FROM DISK = '" + TextBox3.Text + "' WITH REPLACE "

        Dim cmd As SqlCommand = New SqlCommand(str, con)
        Dim cmd1 As SqlCommand = New SqlCommand(str1, con)
        Dim cmd3 As SqlCommand = New SqlCommand(str3, con)

        cmd.ExecuteNonQuery()
        cmd1.ExecuteNonQuery()
        cmd3.ExecuteNonQuery()

        MessageBox.Show("DATABASE RECOVERED Successfull. Restart the system to recovered the data.")
        con.Close()
        Application.Exit()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        OpenFileDialog1 = New OpenFileDialog()

        OpenFileDialog1.InitialDirectory = "D:\Shares\SQL SERVER BACKUP\"
        OpenFileDialog1.Title = "Browse Text File"

        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.CheckPathExists = True

        OpenFileDialog1.DefaultExt = "BAK"
        OpenFileDialog1.Filter = "SQL Server Database (*.bak)|*.bak"
        OpenFileDialog1.FilterIndex = 2
        OpenFileDialog1.RestoreDirectory = True

        OpenFileDialog1.ReadOnlyChecked = True
        OpenFileDialog1.ShowReadOnly = True

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox3.Text = OpenFileDialog1.FileName

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Dispose()
    End Sub

    Private Sub RestoreDatabase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Opacity = 0
        Dim tmr As New Timer
        tmr.Interval = 2
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity += 0.1
                                 If Me.Opacity = 1 Then tmr.Stop()
                             End Sub
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Restore_Database()
    End Sub
End Class