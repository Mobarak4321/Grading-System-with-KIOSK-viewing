Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms

Public Class PrintSubject

    Sub LocalSubjectsReport()

        Dim rptDt As ReportDataSource
        Me.ReportViewer1.RefreshReport()
        Try
            With ReportViewer1.LocalReport
                .ReportPath = Application.StartupPath & "\Reports\SubjectsReport.rdlc"
                .DataSources.Clear()
            End With

            Dim ds As New DataSetGrade
            Dim da As New SqlDataAdapter

            cnn.Open()
            da.SelectCommand = New SqlCommand("SELECT SubjectCode, SubjectDescription, Unit FROM Subjects_Tbl ", cnn)
            da.Fill(ds.Tables("dtSubjects"))
            cnn.Close()

            rptDt = New ReportDataSource("DataSet1", ds.Tables("dtSubjects"))
            ReportViewer1.LocalReport.DataSources.Add(rptDt)
            ReportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)
            ReportViewer1.ZoomMode = ZoomMode.Percent
            ReportViewer1.ZoomPercent = 30
        Catch ex As Exception
            cnn.Close()
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub PrintSubject_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call ConnectionNetwork()
        LocalSubjectsReport()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Dispose()
    End Sub
End Class