Public Class Registrar

    Protected Overrides ReadOnly Property createparams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            Return cp
        End Get
    End Property

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With Dashboard
            .TopLevel = False
            Panel2.Controls.Add(Dashboard)
            .BringToFront()
            .Show()
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'switchPanel(Students)
        With Students
            .TopLevel = False
            Panel2.Controls.Add(Students)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        With Teachers
            .TopLevel = False
            Panel2.Controls.Add(Teachers)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        With Subjects
            .TopLevel = False
            Panel2.Controls.Add(Subjects)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        With Courses
            .TopLevel = False
            Panel2.Controls.Add(Courses)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        With Grades
            .TopLevel = False
            Panel2.Controls.Add(Grades)
            .BringToFront()
            .Show()
            .LoadRecords()
        End With
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Logout.ShowDialog()
    End Sub

    Private Sub Registrar_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub Registrar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With Dashboard
            .TopLevel = False
            Panel2.Controls.Add(Dashboard)
            .BringToFront()
            .Show()
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
End Class
