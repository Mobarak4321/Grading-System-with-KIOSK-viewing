Public Class Librarian

    Private Sub Librarian_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        Me.Opacity = 1
        Dim tmr As New Timer
        tmr.Interval = 1
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity -= 0.1
                                 If Me.Opacity = 0 Then
                                     e.Cancel = False
                                     tmr.Stop()
                                     Me.Dispose()
                                 End If
                             End Sub
    End Sub

    Private Sub Librarian_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With Books
            .TopLevel = False
            Panel2.Controls.Add(Books)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With

        Me.Opacity = 0
        Dim tmr As New Timer
        tmr.Interval = 2
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity += 0.1
                                 If Me.Opacity = 1 Then tmr.Stop()
                             End Sub
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With IssuedBooks
            .TopLevel = False
            Panel2.Controls.Add(IssuedBooks)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With Books
            .TopLevel = False
            Panel2.Controls.Add(Books)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        With ReportBooks
            .TopLevel = False
            Panel2.Controls.Add(ReportBooks)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Application.Restart()
    End Sub
End Class