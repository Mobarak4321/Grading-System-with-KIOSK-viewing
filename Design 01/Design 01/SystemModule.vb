Imports System.Data.SqlClient

Module SystemModule

    Public cnn As New SqlConnection
    Public cmd As New SqlCommand
    Public da As SqlDataAdapter
    Public reader As SqlDataReader
    Public conn_string As String
    Public sql As String

    Public server As String

    Public UserName As String
    Public Prelim, Midterm, PreFinal, Final, FinalGrade As Integer

    Public Sub ConnectionNetwork()
        Try
            Dim iofile As New System.IO.StreamReader(Application.StartupPath + "/Database Connection.txt")
            Dim ioline As String
            ioline = iofile.ReadToEnd

            cnn.ConnectionString = ioline
            cmd.Connection = cnn


        Catch ex As Exception
            MsgBox("Critical Error the system will shutdown the error found:" & ex.Message & "Please Contact System Administrator", MsgBoxStyle.Critical)
            conn_string = "Offline"
            cnn.Close()

            End

        Finally

            cnn.Close()

        End Try
    End Sub

End Module
