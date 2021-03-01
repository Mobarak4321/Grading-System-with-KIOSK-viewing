Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms

Public Class PrintGrade

    Private Sub PrintGrade_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call ConnectionNetwork()
        LoadGradesReport()
    End Sub

    Sub LoadGradesReport()
        Dim rptDt As ReportDataSource
        Me.ReportViewer1.RefreshReport()
        Try
            With ReportViewer1.LocalReport
                .ReportPath = Application.StartupPath & "\Reports\GradesReport.rdlc"
                .DataSources.Clear()
            End With

            Dim ds As New DataSetGrade
            Dim da As New SqlDataAdapter

            cnn.Open()
            da.SelectCommand = New SqlCommand("SELECT StudentID, Name, Subject, Grade, Remarks FROM Grades_Tbl WHERE StudentID = '" & Grades.TextBox1.Text & "' AND AcademicYear = '" & Grades.TextBox2.Text & "' AND Semester = '" & Grades.ComboBox1.Text & "'", cnn)
            da.Fill(ds.Tables("dtGrades"))
            cnn.Close()

            rptDt = New ReportDataSource("DataSet1", ds.Tables("dtGrades"))
            ReportViewer1.LocalReport.DataSources.Add(rptDt)
            ReportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)
            ReportViewer1.ZoomMode = ZoomMode.Percent
            ReportViewer1.ZoomPercent = 30
        Catch ex As Exception
            cnn.Close()
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Dispose()
    End Sub
End Class