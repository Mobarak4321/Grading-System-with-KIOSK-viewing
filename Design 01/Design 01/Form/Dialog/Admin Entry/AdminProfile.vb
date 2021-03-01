Public Class AdminProfile

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox5.Visible = True
        TextBox6.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        Panel3.Visible = True
        Panel4.Visible = True
        Button2.Visible = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox5.Visible = False
        TextBox6.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
        Button2.Visible = False
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

            If TextBox4.UseSystemPasswordChar = True Then
                TextBox4.UseSystemPasswordChar = False
                TextBox5.UseSystemPasswordChar = False
                TextBox6.UseSystemPasswordChar = False
            Else
                TextBox4.UseSystemPasswordChar = True
                TextBox5.UseSystemPasswordChar = True
                TextBox6.UseSystemPasswordChar = True
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Dispose()
    End Sub

    Private Sub AdminProfile_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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