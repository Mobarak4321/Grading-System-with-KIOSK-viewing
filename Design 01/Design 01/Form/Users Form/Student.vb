Imports System.Data.SqlClient

Public Class Student

    Private Sub Student_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With StudentGrades
            .TopLevel = False
            Panel2.Controls.Add(StudentGrades)
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With StudentGrades
            .TopLevel = False
            Panel2.Controls.Add(StudentGrades)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With StudentBooks
            .TopLevel = False
            Panel2.Controls.Add(StudentBooks)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        With StudentIssued
            .TopLevel = False
            Panel2.Controls.Add(StudentIssued)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Logout.ShowDialog()
    End Sub
End Class