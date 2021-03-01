Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms

Public Class PrintStudent

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Dispose()
    End Sub

    Sub LocalStudentReport()
        Dim rptDt As ReportDataSource
        Me.ReportViewer1.RefreshReport()
        Try
            With ReportViewer1.LocalReport
                .ReportPath = Application.StartupPath & "\Reports\StudentsReport.rdlc"
                .DataSources.Clear()
            End With

            Dim ds As New DataSetGrade
            Dim da As New SqlDataAdapter

            cnn.Open()
            da.SelectCommand = New SqlCommand("SELECT StudentID, Lastname, Firstname, MiddleInitial, Course, YearLevel FROM Students_Tbl ", cnn)
            da.Fill(ds.Tables("dtStudents"))
            cnn.Close()

            rptDt = New ReportDataSource("DataSet1", ds.Tables("dtStudents"))
            ReportViewer1.LocalReport.DataSources.Add(rptDt)
            ReportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)
            ReportViewer1.ZoomMode = ZoomMode.Percent
            ReportViewer1.ZoomPercent = 30
        Catch ex As Exception
            cnn.Close()
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub PrintStudent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call ConnectionNetwork()
        LocalStudentReport()
    End Sub
End Class