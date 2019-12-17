Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class Import_Data
    Dim connexcel As OleDbConnection
    Dim daexcel As OleDbDataAdapter
    Dim dsexcel As DataSet
    Dim cmdexcel As OleDbCommand
    Dim drexcel As OleDbDataReader



    Dim connsql As SqlConnection
    Dim dasql As SqlDataAdapter
    Dim dssql As DataSet
    Dim cmdsql As SqlCommand
    Dim drsql As SqlDataReader

    Private Sub Import_Data_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
    End Sub
    
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error Resume Next
        'OpenFileDialog1.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()
        TextBox1.Text = OpenFileDialog1.FileName
        connexcel = New OleDbConnection("provider=Microsoft.ace.OLEDB.12.0;data source=" & TextBox1.Text & ";Extended Properties=Excel 8.0;")
        connexcel.Open()

        Dim dtSheets As DataTable = connexcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        Dim listSheet As New List(Of String)
        Dim drSheet As DataRow

        For Each drSheet In dtSheets.Rows
            listSheet.Add(drSheet("TABLE_NAME").ToString())
        Next

        For Each sheet As String In listSheet
            ListBox1.Items.Add(sheet)
        Next

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        ListBox1.Items.Clear()
        TextBox1.Text = String.Empty
    End Sub
    Sub Koneksisql()
        connsql = New SqlConnection("Data Source=localhost;Initial Catalog=userad;Integrated Security=True")
        'connsql = New SqlConnection("Data Source=localhost;Initial Catalog=OLAH_AXIST;Integrated Security=True")
        connsql.Open()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For baris As Integer = 0 To DataGridView1.RowCount - 2
            Call Koneksisql()
            Dim simpan As String = "insert into tbpegawai values('" & DataGridView1.Rows(baris).Cells(0).Value & "','" & DataGridView1.Rows(baris).Cells(1).Value & "','" & DataGridView1.Rows(baris).Cells(2).Value & "','" & DataGridView1.Rows(baris).Cells(3).Value & "')"
            'Dim simpan As String = "insert into axto values('" & DataGridView1.Rows(baris).Cells(0).Value & "','" & DataGridView1.Rows(baris).Cells(1).Value & "','" & DataGridView1.Rows(baris).Cells(2).Value & "','" & DataGridView1.Rows(baris).Cells(3).Value & "','" & DataGridView1.Rows(baris).Cells(4).Value & "','" & DataGridView1.Rows(baris).Cells(5).Value & "','" & DataGridView1.Rows(baris).Cells(6).Value & "','" & DataGridView1.Rows(baris).Cells(7).Value & "','" & DataGridView1.Rows(baris).Cells(8).Value & "','" & DataGridView1.Rows(baris).Cells(9).Value & "','" & DataGridView1.Rows(baris).Cells(10).Value & "','" & DataGridView1.Rows(baris).Cells(11).Value & "','" & DataGridView1.Rows(baris).Cells(12).Value & "','" & DataGridView1.Rows(baris).Cells(13).Value & "','" & DataGridView1.Rows(baris).Cells(14).Value & "','" & DataGridView1.Rows(baris).Cells(15).Value & "','" & DataGridView1.Rows(baris).Cells(16).Value & "','" & DataGridView1.Rows(baris).Cells(17).Value & "','" & DataGridView1.Rows(baris).Cells(18).Value & "','" & DataGridView1.Rows(baris).Cells(19).Value & "','" & DataGridView1.Rows(baris).Cells(20).Value & "','" & DataGridView1.Rows(baris).Cells(21).Value & "','" & DataGridView1.Rows(baris).Cells(22).Value & "','" & DataGridView1.Rows(baris).Cells(23).Value & "')"
            Timer1.Start()
            cmdsql = New SqlCommand(simpan, connsql)
            cmdsql.ExecuteNonQuery()
        Next

        DataGridView1.Columns.Clear()
        connsql.Close()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        daexcel = New OleDbDataAdapter("select * from [" & ListBox1.Text & "]", connexcel)
        dsexcel = New DataSet
        daexcel.Fill(dsexcel)
        DataGridView1.DataSource = dsexcel.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value = 100 Then
            Timer1.Stop()
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Value += 1
        End If
    End Sub
End Class
