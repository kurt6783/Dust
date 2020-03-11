Imports System.IO.Ports
Imports System.Threading
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Public Class Form1
    Dim Creat As String
    Dim MyPort As Array
    Dim ReceiveData As String
    Dim File As System.IO.StreamWriter
    Dim Result As String
    Dim Sn As Integer
    Dim StartDate As DateTime
    Dim StartTime As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1.Enabled = True
        Button3.Enabled = False
        Sn = 1
        StartDate = DateTime.Now
        StartTime = StartDate.ToString("yyyyMMdd_HHmmss")
        MyPort = IO.Ports.SerialPort.GetPortNames()
        ComboBox1.Items.AddRange(MyPort)
        Creat = "Sn" & "," & "Date" & "," & "Time" & "," & "Sigma0.3" & "," & "Sigma0.5" & "," & "Sigma1.0" & "," & "Sigma3.0" & "," & "Sigma5.0" & "," & "Sigma10.0" & "," & "Delta0.3" & "," & "Delta0.5" & "," & "Delta1.0" & "," & "Delta3.0" & "," & "Delta5.0" & "," & "Delta10.0"
        Log(Creat)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SerialPort1.PortName = ComboBox1.Text
        SerialPort1.BaudRate = ComboBox2.Text
        SerialPort1.Parity = Parity.None
        SerialPort1.DataBits = 8
        SerialPort1.StopBits = StopBits.One
        SerialPort1.Open()
        If SerialPort1.IsOpen = True Then
            Button1.Enabled = False
            Button3.Enabled = True
        ElseIf SerialPort1.IsOpen = False Then
            Button1.Enabled = True
            Button3.Enabled = False
            MsgBox("連線失敗，請重試。")
        End If
    End Sub

    Sub Log(ByVal Result As String)
        ListBox1.Items.Add(Result)
    End Sub

    Private Sub SerialPort1_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Dim DataStr As String
        Dim DataArray(100) As String
        Dim Total_03 As Integer
        Dim Total_05 As Integer
        Dim Total_10 As Integer
        Dim Total_30 As Integer
        Dim Total_50 As Integer
        Dim Total_100 As Integer
        Thread.Sleep(500)
        DataStr = SerialPort1.ReadExisting
        DataArray = DataStr.Split(",")
        Total_100 = Val(DataArray(27))
        Total_50 = Total_100 + Val(DataArray(25))
        Total_30 = Total_50 + Val(DataArray(23))
        Total_10 = Total_30 + Val(DataArray(21))
        Total_05 = Total_10 + Val(DataArray(19))
        Total_03 = Total_05 + Val(DataArray(17))
        Result = Sn & "," & DataArray(0).TrimStart & "," & DataArray(1) & "," & Total_03.ToString & "," & Total_05.ToString & "," & Total_10.ToString & "," & Total_30.ToString & "," & Total_50.ToString & "," & Total_100.ToString & "," & DataArray(17) & "," & DataArray(19) & "," & DataArray(21) & "," & DataArray(23) & "," & DataArray(25)
        Log(Result)
        Sn = Sn + 1
    End Sub

   

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        SerialPort1.Close()
        If SerialPort1.IsOpen = True Then
            Button1.Enabled = False
            Button3.Enabled = True
        ElseIf SerialPort1.IsOpen = False Then
            Button1.Enabled = True
            Button3.Enabled = False
            MsgBox("連線已中斷。")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        File = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\DustData\" & StartTime & ".csv", True)
        For i = 0 To ListBox1.Items.Count - 1
            File.WriteLine(ListBox1.Items.Item(i))
        Next
        File.Close()


        '1.讀取落塵值
        Button2.Text = "資料轉換中"
        Dim app As New Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        Dim range As Excel.Range
        app.DisplayAlerts = False
        app.Visible = False
        
        Dim DustData(,) As Integer = Nothing
        Dim Row As Integer
        Select Case ComboBox6.Text
            Case "CAM房"
                ReDim DustData(7, 1)
                Row = 8
            Case "FQC"
                ReDim DustData(44, 1)
                Row = 45
            Case "內層"
                ReDim DustData(62, 1)
                Row = 63
            Case "外層"
                ReDim DustData(25, 1)
                Row = 26
            Case "印刷(UCoater)"
                ReDim DustData(38, 1)
                Row = 39
            Case "印刷"
                ReDim DustData(34, 1)
                Row = 35
            Case "綠漆"
                ReDim DustData(76, 1)
                Row = 77
            Case "線路"
                ReDim DustData(52, 1)
                Row = 53
            Case "壓合一廠疊板房"
                ReDim DustData(5, 1)
                Row = 6
            Case "壓合二廠疊板房"
                ReDim DustData(11, 1)
                Row = 12
        End Select

        book = app.Workbooks.Open(Application.StartupPath & "\DustData\" & StartTime & ".csv")
        sheet = book.Sheets(1)

        If ComboBox6.Text = "FQC" Or ComboBox6.Text = "線路" Or ComboBox6.Text = "綠漆" Then
            For i = 0 To Row - 1
                range = sheet.Cells(i + 2, 5) 'Sigma 0.5
                DustData(i, 0) = range.Value
                range = sheet.Cells(i + 2, 8) 'Sigma 5
                DustData(i, 1) = range.Value
            Next
        ElseIf ComboBox6.Text = "CAM房" Or ComboBox6.Text = "內層" Or ComboBox6.Text = "外層" Or ComboBox6.Text = "印刷(UCoater)" Or ComboBox6.Text = "印刷" Or ComboBox6.Text = "壓合一廠疊板房" Or ComboBox6.Text = "壓合二廠疊板房" Then
            For i = 0 To Row - 1
                range = sheet.Cells(i + 2, 12) ' Delta 1
                DustData(i, 0) = range.Value
                range = sheet.Cells(i + 2, 5) ' Sigma 0.5
                DustData(i, 1) = range.Value
            Next
        End If
        book.Close()
        '2.讀取表單設定
        Dim TXT As New System.IO.StreamReader(Application.StartupPath & "\DustSetting\" & ComboBox6.Text & ".txt", System.Text.Encoding.GetEncoding("BIG5"))
        Dim DustSetting(Row - 1, 4)
        Dim Dust(4) As String
        Dim x As Integer = 0
        Dim Str As String
        Str = TXT.ReadLine
        Do Until Str = Nothing
            Dust = Str.Split(",")
            DustSetting(x, 0) = Dust(0) 'Sn
            DustSetting(x, 1) = Dust(1) '0.5x
            DustSetting(x, 2) = Dust(2) '0.5y
            DustSetting(x, 3) = Dust(3) '5x
            DustSetting(x, 4) = Dust(4) '5y
            x = x + 1
            Str = TXT.ReadLine
        Loop
        '3.讀取基底表單、填入表單、另存表單
        book = app.Workbooks.Open(Application.StartupPath & "\DustList\" & ComboBox6.Text & ".xls")
        sheet = book.Sheets(1)
        For i = 0 To Row - 1
            range = sheet.Cells(CInt(DustSetting(i, 2)), CInt(DustSetting(i, 1)))
            range.Value = DustData(i, 0)
            range = sheet.Cells(CInt(DustSetting(i, 4)), CInt(DustSetting(i, 3)))
            range.Value = DustData(i, 1)
            Button2.Text = "資料轉換中" & "第" & i & "筆"
        Next
        book.SaveAs((Application.StartupPath & "\ExcelData\" & ComboBox6.Text & "_" & StartTime & ".xls"))
        book.Close()
        app.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(book)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(app)
        GC.Collect()

        MsgBox("資料轉換完成。")
        Button2.Text = "資料轉換"
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
        Sn = Sn - 1
    End Sub
End Class
