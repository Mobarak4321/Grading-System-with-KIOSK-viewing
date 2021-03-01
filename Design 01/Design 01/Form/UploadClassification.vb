Public Class UploadClassification

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        Me.Dispose()
        TeacherGrade.SaveExcelData()
        TeacherGrade.Button1.Enabled = True
        TeacherGrade.Button3.Enabled = True
        TeacherGrade.Button4.Enabled = True
        TeacherGrade.Button5.Enabled = False
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        Me.Dispose()
        TeacherGrade.Button1.Enabled = True
        TeacherGrade.Button3.Enabled = True
        TeacherGrade.Button4.Enabled = True
        TeacherGrade.Button5.Enabled = False
    End Sub

    Private Sub UploadClassification_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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