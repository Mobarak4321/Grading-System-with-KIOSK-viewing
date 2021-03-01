Public Class Database

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With BackupDatabase
            .ShowDialog()
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With RestoreDatabase
            .ShowDialog()
        End With
    End Sub
End Class