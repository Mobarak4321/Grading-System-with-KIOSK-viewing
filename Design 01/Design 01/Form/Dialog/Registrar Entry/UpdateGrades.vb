Imports System.Data.SqlClient

Public Class UpdateGrades

    Sub Compute_Grades()
        Prelim = Val(TextBox3.Text)
        TextBox3.Text = Prelim

        Midterm = Val(TextBox4.Text)
        TextBox4.Text = Midterm

        PreFinal = Val(TextBox5.Text)
        TextBox5.Text = PreFinal

        Final = Val(TextBox6.Text)
        TextBox6.Text = Final

        If TextBox3.Text = 0 Or TextBox4.Text = 0 Or TextBox5.Text = 0 Or TextBox6.Text = 0 Then
            TextBox7.Text = "INC"
        Else
            FinalGrade = Val(TextBox3.Text * 0.2 + TextBox4.Text * 0.2 + TextBox5.Text * 0.2 + TextBox6.Text * 0.4)
            TextBox7.Text = FinalGrade
        End If

        If FinalGrade >= 97 Then
            TextBox7.Text = "1"
            TextBox8.Text = "PASS"
        ElseIf FinalGrade >= 94 Then
            TextBox7.Text = "1.25"
            TextBox8.Text = "PASS"
        ElseIf FinalGrade >= 91 Then
            TextBox7.Text = "1.5"
            TextBox8.Text = "PASS"
        ElseIf FinalGrade >= 88 Then
            TextBox7.Text = "1.75"
            TextBox8.Text = "PASS"
        ElseIf FinalGrade >= 85 Then
            TextBox7.Text = "2"
            TextBox8.Text = "PASS"
        ElseIf FinalGrade >= 82 Then
            TextBox7.Text = "2.25"
            TextBox8.Text = "PASS"
        ElseIf FinalGrade >= 80 Then
            TextBox7.Text = "2.5"
            TextBox8.Text = "PASS"
        ElseIf FinalGrade >= 76 Then
            TextBox7.Text = "2.75"
            TextBox8.Text = "PASS"
        ElseIf FinalGrade = 75 Then
            TextBox7.Text = "3"
            TextBox8.Text = "PASS"
        ElseIf FinalGrade = 9 Then
            TextBox7.Text = "UW"
            TextBox8.Text = "UW"
        ElseIf FinalGrade = 8 Then
            TextBox7.Text = "CF"
            TextBox8.Text = "CF"
        ElseIf FinalGrade = 7 Then
            TextBox8.Text = "AW"
            TextBox8.Text = "AW"
        ElseIf FinalGrade = 6 Then
            TextBox7.Text = "UW"
            TextBox8.Text = "UW"
        ElseIf FinalGrade = 5 Then
            TextBox7.Text = "5"
            TextBox8.Text = "Failed"
        ElseIf FinalGrade = 4 Then
            TextBox8.Text = "4"
            TextBox8.Text = "INC"
        ElseIf FinalGrade <= 74 Then
            TextBox8.Text = "5"
            TextBox8.Text = "FAILED"
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox4.Text.Trim <> "" Then

            Call ConnectionNetwork()

            Try
                sql = ("UPDATE Grades_Tbl SET Prelim ='" & TextBox3.Text & "', Midterm ='" & TextBox4.Text & "', PreFinal= '" & TextBox5.Text & "', Final= '" & TextBox6.Text & "',  Grade= '" & TextBox7.Text.Trim & "', Remarks= '" & TextBox8.Text.Trim & "' WHERE ID= " & Label12.Text.Trim & "")
                cmd = New SqlCommand(sql, cnn)
                cnn.Open()

                cmd.ExecuteNonQuery()

                MsgBox("Successfully Updated")
                cnn.Close()
                cmd.Dispose()
                Me.Dispose()
                'Registrar.switchPanel(Grades)

            Catch ex As Exception
                MsgBox("error found ;" & ex.Message & ex.StackTrace)
            Finally
                cnn.Close()
            End Try

        End If
    End Sub

    Private Sub UpdateGrades_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Opacity = 0
        Dim tmr As New Timer
        tmr.Interval = 2
        tmr.Start()

        AddHandler tmr.Tick, Sub()
                                 Me.Opacity += 0.1
                                 If Me.Opacity = 1 Then tmr.Stop()
                             End Sub

        If TextBox8.Text = "INC" Then
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            Button1.Show()

            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Compute_Grades()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Compute_Grades()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Compute_Grades()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Compute_Grades()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub
End Class