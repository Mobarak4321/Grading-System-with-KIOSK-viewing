Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient

Public Class TeacherGrade

    Public Sub ExportToExcel()

        Dim Data As Integer = 0
        Label4.Text = Data
        Dim default_location As String = "D:\Database Folder\Excel\GRADETEMPLATE.xlsx"
        Try
            Dim dset As New DataSet
            dset.Tables.Add()

            For i As Integer = 0 To DataGridView1.ColumnCount - 1
                dset.Tables(0).Columns.Add(DataGridView1.Columns(i).HeaderText)
            Next
            '  add rows to the table
            Dim dr1 As DataRow
            For i As Integer = 0 To DataGridView1.RowCount - 1
                dr1 = dset.Tables(0).NewRow
                For j As Integer = 1 To DataGridView1.Columns.Count - 1
                    dr1(j) = DataGridView1.Rows(i).Cells(j).Value
                Next
                dset.Tables(0).Rows.Add(dr1)
            Next

            Dim excel As Microsoft.Office.Interop.Excel.Application
            excel = New Microsoft.Office.Interop.Excel.Application
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

            excel.Visible = True
            excel.UserControl = True

            wBook = excel.Workbooks.Open(default_location)
            wSheet = wBook.Sheets("Grading_Sheets")
            wSheet.Unprotect("12345")
            excel.Range("B50:C50").ColumnWidth = 35.0
            With wBook
                .Sheets("Grading_Sheets").Select()
                .Sheets(1).Name = "Grading_Sheets"

            End With

            Dim dt As System.Data.DataTable = dset.Tables(0)
            For Each col As DataGridViewColumn In DataGridView1.Columns
                wSheet.Cells(1, col.Index + 1) = col.HeaderText("0").ToString
            Next


            For i = 0 To DataGridView1.RowCount - 1
                For j = 0 To DataGridView1.ColumnCount - 1
                    wSheet.Columns.NumberFormat = "@"
                    wSheet.Cells(i + 11, j + 2).value = DataGridView1.Rows(i).Cells(j).Value.ToString
                    wSheet.Cells(5, 2).value = Label3.Text
                    wSheet.Cells(6, 2).value = Label4.Text
                    wSheet.Cells(7, 2).value = Teacher.Label2.Text
                    wSheet.Cells(8, 2).value = Label15.Text
                Next j
            Next i

            Data = 1

            excel.ActiveWorkbook.ActiveSheet.Protect("12345")

            Dim blnFileOpen As Boolean = False
            Try
                Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(default_location)
                fileTemp.Close()
            Catch ex As Exception
                blnFileOpen = False
            End Try
            wBook.SaveAs("D:\Database Folder\Excel\Export Excel\" & Label15.Text & ".xls")
            wBook.Close()
            excel.Quit()

            releaseObject(excel)
            releaseObject(wBook)
            releaseObject(wSheet)


            DataGridView1.DataSource = Nothing

        Catch ex As Exception
            MsgBox("ERROR IN :" & ex.Message & " Please contact the Administrator")

        End Try

        MsgBox("Exporting  List into Excel is Complete")


        'End If
    End Sub

    Sub KillExcel()
        Dim Application() As Process = Process.GetProcessesByName("EXCEL")
        For Each Process As Process In Application
            If Process.GetProcessesByName("EXCEL").Count <> 0 Then
                Process.Kill()
            End If
        Next

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

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
            'TextBox9.Text = "0"
        Else
            FinalGrade = Val(TextBox3.Text * 0.2 + TextBox4.Text * 0.2 + TextBox5.Text * 0.2 + TextBox6.Text * 0.4)
            TextBox7.Text = FinalGrade
        End If

        If FinalGrade >= 97 Then
            TextBox8.Text = "1"
            TextBox9.Text = "PASS"
        ElseIf FinalGrade >= 94 Then
            TextBox8.Text = "1.25"
            TextBox9.Text = "PASS"
        ElseIf FinalGrade >= 91 Then
            TextBox8.Text = "1.5"
            TextBox9.Text = "PASS"
        ElseIf FinalGrade >= 88 Then
            TextBox8.Text = "1.75"
            TextBox9.Text = "PASS"
        ElseIf FinalGrade >= 85 Then
            TextBox8.Text = "2"
            TextBox9.Text = "PASS"
        ElseIf FinalGrade >= 82 Then
            TextBox8.Text = "2.25"
            TextBox9.Text = "PASS"
        ElseIf FinalGrade >= 80 Then
            TextBox8.Text = "2.5"
            TextBox9.Text = "PASS"
        ElseIf FinalGrade >= 76 Then
            TextBox8.Text = "2.75"
            TextBox9.Text = "PASS"
        ElseIf FinalGrade = 75 Then
            TextBox8.Text = "3"
            TextBox9.Text = "PASS"
        ElseIf FinalGrade = 9 Then
            TextBox8.Text = "UW"
            TextBox9.Text = "UW"
        ElseIf FinalGrade = 8 Then
            TextBox8.Text = "CF"
            TextBox9.Text = "CF"
        ElseIf FinalGrade = 7 Then
            TextBox8.Text = "AW"
            TextBox9.Text = "AW"
        ElseIf FinalGrade = 6 Then
            TextBox8.Text = "UW"
            TextBox9.Text = "UW"
        ElseIf FinalGrade = 5 Then
            TextBox8.Text = "5"
            TextBox9.Text = "Failed"
        ElseIf FinalGrade = 4 Then
            TextBox8.Text = "4"
            TextBox9.Text = "INC"
        ElseIf FinalGrade <= 74 Then
            TextBox8.Text = "5"
            TextBox9.Text = "FAILED"
        End If

    End Sub

    Sub SchoolYear()
        Call ConnectionNetwork()

        Try
            sql = "SELECT  *  FROM AcademicYear_Tbl"
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim Year As Boolean = False

            Dim AcademicYear As String = ""
            Dim Semester As String = ""

            While reader.Read

                Year = True

                AcademicYear = reader("AcademicYear").ToString
                Semester = reader("Semester").ToString

            End While

            If Year = True Then

                Label4.Text = Semester
                Label3.Text = AcademicYear
                cnn.Close()

            End If
        Catch ex As Exception
            MsgBox("Error Found ; " & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Sub Subject_Listed()
        Call ConnectionNetwork()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM TeacherAdvisories_Tbl WHERE Lastname = '" & Teacher.Label3.Text & "' AND Firstname = '" & Teacher.Label4.Text & "' AND MiddleInitial = '" & Teacher.Label5.Text & "' AND Semester = '" & Label4.Text & "' AND AcademicYear = '" & Label3.Text & "'  "
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader
            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("Subject").ToString)
                ListView1.Items.Add(x)

                ListView1.Focus()
                ListView1.Items(0).Selected = True
            Loop
        Catch ex As Exception
            MsgBox("Error Found ; " & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cmd.Dispose()
            cnn.Close()
        End Try
    End Sub

    Sub Student_List()

        Call ConnectionNetwork()

        DataGridView1.Columns.Clear()

        Try

            sql = ("SELECT StudentID, Name FROM StudentSubjects_Tbl WHERE StudentID != '' AND Subject = '" & Label15.Text & "' AND Professor= '" & Teacher.Label2.Text & "' AND AcademicYear= '" & Label3.Text & "' AND Semester= '" & Label4.Text & "' ")
            cmd = New SqlCommand(sql, cnn)
            cnn.Open()

            Dim adapater As New SqlDataAdapter(cmd)
            Dim table As New DataTable

            adapater.Fill(table)

            DataGridView1.DataSource = table

            cnn.Close()
            DataGridView1.ColumnHeadersVisible = False

        Catch ex As Exception
            MsgBox("Error Found ; " & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Sub Load_Excel32()
        Try

            DataGridView1.DataSource = ""

            Dim ExcelProvider As New System.IO.StreamReader(Application.StartupPath + "/Excel Provider.txt")
            Dim ExcelProviderLine As String
            ExcelProviderLine = ExcelProvider.ReadToEnd

            Dim ExcelProperties As New System.IO.StreamReader(Application.StartupPath + "/Excel Properties.txt")
            Dim ExcelPropertiesLine As String
            ExcelPropertiesLine = ExcelProperties.ReadToEnd

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataset As System.Data.DataSet
            Dim Mycommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = Label14.Text

            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.JET.OLEDB.4.0; Data source =" + path + "; Extended Properties = Excel 8.0; ")
            Mycommand = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [GRADE_TEMPLATE$B4:J50]", MyConnection)

            dataset = New System.Data.DataSet
            Mycommand.Fill(dataset)
            DataGridView1.DataSource = dataset.Tables(0)

            MyConnection.Close()

        Catch ex As Exception
            MsgBox("ERROR IN :" & ex.Message & " Please contact the Administrator")
            Load_Excel64()
        End Try
    End Sub

    Sub Load_Excel64()
        Try

            DataGridView1.DataSource = ""

            Dim ExcelProvider As New System.IO.StreamReader(Application.StartupPath + "/Excel Provider.txt")
            Dim ExcelProviderLine As String
            ExcelProviderLine = ExcelProvider.ReadToEnd

            Dim ExcelProperties As New System.IO.StreamReader(Application.StartupPath + "/Excel Properties.txt")
            Dim ExcelPropertiesLine As String
            ExcelPropertiesLine = ExcelProperties.ReadToEnd

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataset As System.Data.DataSet
            Dim Mycommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = Label14.Text

            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0; Data source =" + path + "; Extended Properties = Excel 8.0; ")
            Mycommand = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [GRADE_TEMPLATE$B6:J50]", MyConnection)

            dataset = New System.Data.DataSet
            Mycommand.Fill(dataset)
            DataGridView1.DataSource = dataset.Tables(0)

            MyConnection.Close()

        Catch ex As Exception
            MsgBox("ERROR IN :" & ex.Message & " Please contact the Administrator")
            Button3.Enabled = False
            Button2.Enabled = True
            Button4.Enabled = True
            Button6.Enabled = True
        End Try
    End Sub

    Sub SaveExcelData()

        Try

            If DataGridView1.Rows(2).Cells(0).Value = Teacher.Label2.Text And DataGridView1.Rows(0).Cells(0).Value = Label3.Text And DataGridView1.Rows(1).Cells(0).Value = Label4.Text Then
                Call ConnectionNetwork()
                Dim result As Object

                For x As Integer = 6 To 43
                    Dim StudentID As String = DataGridView1.Rows(x).Cells(0).Value
                    Dim Name As String = DataGridView1.Rows(x).Cells(1).Value
                    Dim AcademicYear As String = DataGridView1.Rows(0).Cells(0).Value
                    Dim Semester As String = DataGridView1.Rows(1).Cells(0).Value
                    Dim Professor As String = DataGridView1.Rows(2).Cells(0).Value
                    Dim Subject As String = DataGridView1.Rows(3).Cells(0).Value
                    Dim Prelim As String = DataGridView1.Rows(x).Cells(2).Value
                    Dim Midterm As String = DataGridView1.Rows(x).Cells(3).Value
                    Dim PreFinal As String = DataGridView1.Rows(x).Cells(4).Value
                    Dim Final As String = DataGridView1.Rows(x).Cells(5).Value
                    Dim Grade As String = DataGridView1.Rows(x).Cells(7).Value
                    Dim Remarks As String = DataGridView1.Rows(x).Cells(8).Value

                    If (String.IsNullOrEmpty(StudentID)) Or (String.IsNullOrWhiteSpace(StudentID)) Then

                    Else
                        sql = "SELECT * FROM Grades_Tbl WHERE StudentID = '" & StudentID & "' AND Subject = '" & Subject & "' AND AcademicYear = '" & AcademicYear & "' AND Semester = '" & Semester & "' "
                        Dim cmd As New SqlCommand(sql, cnn)
                        cnn.Open()

                        result = cmd.ExecuteScalar

                        cnn.Close()

                        If result IsNot Nothing = False Then
                            Try
                                Dim cmd2 As New SqlCommand
                                Dim cmd3 As New SqlCommand
                                Dim sqlinsert As String
                                sqlinsert = "INSERT INTO Grades_Tbl (StudentID, Name, AcademicYear, Semester, Professor, Subject, Prelim, Midterm, PreFinal, Final, Grade, Remarks) VALUES ('" & StudentID & "', '" & Name & "', '" & AcademicYear & "' , '" & Semester & "' , '" & Professor & "' , '" & Subject & "', '" & Prelim & "', '" & Midterm & "', '" & PreFinal & "', '" & Final & "', '" & Grade & "', '" & Remarks & "')"
                                cmd2 = New SqlCommand(sqlinsert, cnn)
                                cnn.Open()

                                cmd2.ExecuteNonQuery()
                            Catch ex As Exception
                                MsgBox("error " & ex.Message & " " & ex.StackTrace)
                                cnn.Close()

                            Finally
                                cnn.Close()
                            End Try

                        Else
                            MsgBox("Student  " & StudentID & "  (" & Name & ") Grade on " & Subject & " is already submitted.")
                        End If
                    End If

                Next


                MsgBox("Uploading grades is complete")
                cnn.Close()

            Else
                MsgBox("The File is outdated or not belongs to you")

            End If

        Catch ex As Exception
            'MsgBox("Importing Grades is finish")
            cnn.Close()
        End Try

    End Sub

    Private Sub TeacherGrade_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Subject_Listed()
    End Sub

    Private Sub ListView1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick
        Dim subject As String = ListView1.SelectedItems(0).SubItems(0).Text

        Label15.Text = subject
    End Sub

    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick
        Student_List()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim index As Integer
        index = e.RowIndex

        Dim selectedRow As DataGridViewRow

        selectedRow = DataGridView1.Rows(index)


        TextBox1.Text = selectedRow.Cells(0).Value.ToString
        TextBox2.Text = selectedRow.Cells(1).Value.ToString
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox5.Text.Trim <> "" Or TextBox6.Text.Trim <> "" Or TextBox7.Text.Trim <> "" Or TextBox8.Text.Trim <> "" Then

            Dim existing As Boolean

            Call ConnectionNetwork()

            Try
                sql = "SELECT * FROM Grades_Tbl WHERE StudentID= '" & TextBox1.Text.Trim & "' AND AcademicYear = '" & Label3.Text.Trim & "' AND Semester = '" & Label4.Text.Trim & "' AND Subject = '" & Label15.Text & "' "
                cmd = New SqlCommand(sql, cnn)
                cnn.Open()

                reader = cmd.ExecuteReader

                If reader.Read Then

                    If reader.HasRows Then

                        existing = True

                    Else

                        existing = False

                    End If

                End If

            Catch ex As Exception
                MsgBox("error found ;" & ex.Message & ex.StackTrace)
            Finally

                cnn.Close()

                If existing = False Then

                    Try

                        sql = ("INSERT INTO Grades_Tbl (StudentID, Name, AcademicYear, Semester, Professor, Subject, Prelim, Midterm, PreFinal, Final, Grade, Remarks) VALUES ('" & TextBox1.Text.Trim & "', '" & TextBox2.Text.Trim & "','" & Label3.Text.Trim & "', '" & Label4.Text.Trim & "', '" & Teacher.Label2.Text & "','" & Label15.Text.Trim & "','" & TextBox3.Text.Trim & "', '" & TextBox4.Text.Trim & "','" & TextBox5.Text.Trim & "', '" & TextBox6.Text.Trim & "', '" & TextBox7.Text.Trim & "', '" & TextBox9.Text.Trim & "')")
                        cmd = New SqlCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()
                        cnn.Close()
                        cmd.Dispose()

                        MsgBox("Successfully Added")

                        sql = "INSERT INTO UsersLog_Tbl (UserID, Activity, Time) VALUES ('" & Teacher.Label1.Text & "', 'Upload Grades for (' + '" & TextBox1.Text + ")" & "', '" & Date.Now & "')"
                        cmd = New SqlCommand(sql, cnn)
                        cnn.Open()


                        cmd.ExecuteNonQuery()

                        cnn.Close()

                    Catch ex As SystemException

                        MsgBox("error " & ex.Message & " " & ex.StackTrace)

                        cnn.Close()

                    Finally

                        cnn.Close()
                        cmd.Dispose()

                    End Try

                Else

                    MsgBox("already registered")

                End If

            End Try

        Else

            MsgBox("fill all the blanks")

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        DataGridView1.DataSource = ""
        Dim OpenFile As New OpenFileDialog

        OpenFileDialog1.Filter = "Excel File(.xls) | *.xls"

        OpenFile.Filter = "Exel File (.xlsx) | *.xlsx | Excel File (.xls)| *.xls "

        If OpenFileDialog1.ShowDialog Then
            Label14.Text = OpenFileDialog1.FileName

            Button5.Enabled = True
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Load_Excel32()

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ExportToExcel()
        DataGridView1.DataSource = ""
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        UploadClassification.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        DataGridView1.DataSource = ""
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
    End Sub
End Class