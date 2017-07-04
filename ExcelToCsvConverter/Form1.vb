Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.IO
Imports Excel

Public Class frmXlsToCsv
    Dim result As DataSet = New DataSet
    Private Sub btnBrows_Click(sender As Object, e As EventArgs) Handles btnBrows.Click
        Dim Chosen_File As String = ""
        If (OpenFileDialog1.ShowDialog = DialogResult.OK) Then
            Chosen_File = OpenFileDialog1.FileName
        End If

        If (Chosen_File = String.Empty) Then
            Return
        End If

        txtXlx.Text = Chosen_File
        getExcelData(txtXlx.Text)
    End Sub
    Private Sub DefaultValues()
        'txtXlx.Text = "C: \Users\rathma\Desktop\New folder\Sample Audit Trail.xlsx"
        'cmbSheet.Text = "Sheet1"
        'txtFolder.Text = "C: \Users\rathma\Desktop\New folder"
        'txtCsv.Text = "test"
    End Sub

    Private Sub getExcelData(ByVal tfile As String)
        If tfile.EndsWith(".xlsx") Then
            ' Reading from a binary Excel file (format; *.xlsx)
            Dim stream As FileStream = File.Open(tfile, FileMode.Open, FileAccess.Read)
            Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream)
            result = excelReader.AsDataSet
            excelReader.Close()
        End If

        If tfile.EndsWith(".xls") Then
            ' Reading from a binary Excel file ('97-2003 format; *.xls)
            Dim stream As FileStream = File.Open(tfile, FileMode.Open, FileAccess.Read)
            Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateBinaryReader(stream)
            result = excelReader.AsDataSet
            excelReader.Close()
        End If

        Dim items As List(Of String) = New List(Of String)
        Dim i As Integer = 0
        Do While (i < result.Tables.Count)
            items.Add(result.Tables(i).TableName.ToString)
            i = (i + 1)
        Loop

        cmbSheet.DataSource = items
    End Sub

    Private Sub converToCSV(ByVal ind As Integer)
        ' sheets in excel file becomes tables in dataset
        'result.Tables[0].TableName.ToString(); // to get sheet name (table name)
        Dim a As String = ""
        Dim row_no As Integer = 0
        'Dim dttbl As DataTable = result.Tables(0)
        'result.Tables.Remove(dttbl)
        'result.Tables.Add(GetTable)

        Dim dttbl As DataTable = GetTable()


        'for column1 :Employee Profile
        Dim i As Integer = 0
            Dim newTableRowCnt As Integer = 1
        Do While (i < result.Tables(ind).Columns.Count)
            If i = 0 Then

                a = result.Tables(ind).Rows(row_no)(i).ToString.LastIndexOf(">")
                Dim words As String() = result.Tables(ind).Rows(row_no)(i).ToString.Split(">")
                If words.Length > 1 Then
                    dttbl.Rows.Add()
                    dttbl.Rows(newTableRowCnt)(0) = words(words.Length - 1).ToString
                    newTableRowCnt = newTableRowCnt + 1
                End If
                row_no = (row_no + 1)

            End If
            If row_no = result.Tables(ind).Rows.Count Then
                row_no = 0
                i = i + 1
                newTableRowCnt = 1
                Exit Do
            End If

        Loop


        'for column1 :Employee 
        newTableRowCnt = 1
        While newTableRowCnt < dttbl.Rows.Count - 1
            i = 1
            a = result.Tables(ind).Rows(row_no)(i).ToString
                Dim words As String() = a.Split(":")
                If words(0).ToLower = "employee" Then

                dttbl.Rows(newTableRowCnt)(1) = words(words.Length - 1).ToString.Replace(",", " ")
                a = result.Tables(ind).Rows(row_no + 1)(i).ToString
                Dim words1 As String() = a.Split(":")
                If words1(0).ToLower = "associate id" Then
                    If (words1.Length = 3) Then
                        dttbl.Rows(newTableRowCnt)(3) = words1(words1.Length - 1).ToString.Substring(0, 10)
                        words1(1) = words1(1).ToString.Substring(0, 10)
                        dttbl.Rows(newTableRowCnt)(2) = words1(1).ToString
                    Else
                        words1(1) = words1(1).ToString.Substring(0, 10)
                        dttbl.Rows(newTableRowCnt)(2) = words1(1).ToString
                    End If


                End If
                    newTableRowCnt = newTableRowCnt + 1
                End If
                row_no = (row_no + 1)


                If row_no = result.Tables(ind).Rows.Count Then
                    row_no = 0
                    i = i + 1
                    newTableRowCnt = 1
                End If

            End While

        'for column1 :Action,Effective Date
        newTableRowCnt = 1
        row_no = 0

        While newTableRowCnt < dttbl.Rows.Count - 1

            i = 2

            a = result.Tables(ind).Rows(row_no)(i).ToString
            Dim words As String() = a.Split(":")
            If words(0).ToLower = "action" Then

                dttbl.Rows(newTableRowCnt)(4) = words(words.Length - 1).ToString
                a = result.Tables(ind).Rows(row_no + 1)(i).ToString
                Dim words1 As String() = a.Split(":")
                If words1(0).ToLower = "effective date" Then

                    dttbl.Rows(newTableRowCnt)(5) = words1(words1.Length - 1).ToString

                End If

                newTableRowCnt = newTableRowCnt + 1
            End If
            row_no = (row_no + 1)


            If row_no = result.Tables(ind).Rows.Count Then
                row_no = 0
                i = i + 1
                newTableRowCnt = 1
            End If

        End While

        'for column1 :Change To
        newTableRowCnt = 1
        row_no = 0

        While newTableRowCnt < dttbl.Rows.Count - 1

            i = 3

            a = result.Tables(ind).Rows(row_no)(i).ToString

            If a <> "" And a <> "Changed To" Then

                dttbl.Rows(newTableRowCnt)(6) = a.ToString
                newTableRowCnt = newTableRowCnt + 1
            End If
            row_no = (row_no + 1)


            If row_no = result.Tables(ind).Rows.Count Then
                row_no = 0
                i = i + 1
                newTableRowCnt = 1
                Exit While
            End If

        End While

        'for column1 :User,Role,Date
        newTableRowCnt = 1
        row_no = 0

        While newTableRowCnt < dttbl.Rows.Count - 1

            i = 4

            a = result.Tables(ind).Rows(row_no)(i).ToString

            If a <> "" And a <> "Changed By" Then

                dttbl.Rows(newTableRowCnt)(7) = a.ToString
                Try
                    a = result.Tables(ind).Rows(row_no + 1)(i).ToString
                Catch ex As Exception
                    Exit While
                End Try

                Dim words As String() = a.Split(":")
                If words(0).ToLower = "role" Then

                    dttbl.Rows(newTableRowCnt)(8) = words(words.Length - 1).ToString
                    a = result.Tables(ind).Rows(row_no + 2)(i).ToString
                    Dim words1 As String() = a.Split(":")
                    If words1(0).ToLower = "date" Then
                        words1(1) = words1(1).ToString.Substring(0, words1(1).ToString.Length - 3)
                        dttbl.Rows(newTableRowCnt)(9) = words1(1).ToString

                    End If

                    newTableRowCnt = newTableRowCnt + 1
                End If
            End If
            row_no = (row_no + 1)


            If row_no = result.Tables(ind).Rows.Count Then
                row_no = 0
                i = i + 1
                newTableRowCnt = 1
                Exit While
            End If

        End While

        'write data set to Excel

        'to be removed
        Dim dttbl1 As DataTable = result.Tables(0)
        result.Tables.Remove(dttbl1)
        result.Tables.Add(dttbl)
        row_no = 0


        a = ""
        While (row_no < result.Tables(ind).Rows.Count)
            i = 0
            Do While (i < result.Tables(ind).Columns.Count)
                a = (a _
                            + (result.Tables(ind).Rows(row_no)(i).ToString + ","))
                i = (i + 1)
            Loop

            row_no = (row_no + 1)
            a = (a + "" & vbLf)

        End While

        Dim output As String = (txtFolder.Text + ("\" _
                    + (txtCsv.Text + ".csv")))
        Dim csv As StreamWriter = New StreamWriter(output, False)
        csv.Write(a)
        csv.Close()
        MessageBox.Show("File converted succussfully")
        txtXlx.Text = ""
        txtFolder.Text = ""
        txtCsv.Text = ""
        cmbSheet.DataSource = Nothing
        Return
    End Sub

    Private Sub btnBrowsFolder_Click(sender As Object, e As EventArgs) Handles btnBrowsFolder.Click
        Dim result As DialogResult = Me.FolderBrowserDialog1.ShowDialog
        Dim foldername As String = ""
        If (result = DialogResult.OK) Then
            foldername = Me.FolderBrowserDialog1.SelectedPath
        End If

        txtFolder.Text = foldername
    End Sub

    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
        Dim fileName As String = ""
        fileName = txtCsv.Text
        If (fileName = "") Then
            MessageBox.Show("Enter Valid file name")
            Return
        End If

        converToCSV(cmbSheet.SelectedIndex)
    End Sub

    Function GetTable() As DataTable
        ' Create new DataTable instance.
        Dim table As New DataTable

        ' Create four typed columns in the DataTable.
        table.Columns.Add("Area Changed", GetType(String))
        table.Columns.Add("Name", GetType(String))
        table.Columns.Add("Associate ID", GetType(String))
        table.Columns.Add("Position ID", GetType(String))
        table.Columns.Add("Action", GetType(String))
        table.Columns.Add("Effective Date", GetType(String))
        table.Columns.Add("Changed To", GetType(String))
        table.Columns.Add("Changed By", GetType(String))
        table.Columns.Add("Change By Role", GetType(String))
        table.Columns.Add("Processed Date", GetType(String))
        ' Add five rows with those columns filled in the DataTable.
        table.Rows.Add("Area Changed", "Name", "Associate ID", "Position ID", "Action", "Effective Date", "Changed To", "Changed By", "Change By Role", "Processed Date")

        Return table
    End Function

    Private Sub frmXlsToCsv_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DefaultValues()
    End Sub
    Private Sub DatasetToXls()

    End Sub
End Class
