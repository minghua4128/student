Imports System
Imports System.Data
Imports System.Data.OleDb
Module Module1
    ''' <summary>
    ''' 學生成績資料結構
    ''' </summary>
    Public Structure StudentData
        ''' <summary>
        ''' 學生姓名
        ''' </summary>
        Dim Name As String
        Dim Test As Integer
        Dim Project As Integer
        Dim Quizzes As Integer
        Dim Exam As Integer
        Dim CAMarks As Double
        Dim ModlueGrade As String
        Dim ModuleMarks As Double
        Dim Remarks As String
    End Structure
    Public Sub CleanAll()
        '清除所有文字方塊的資料
        Form1.TextBox1.Text = ""
        Form1.InputBox1.Text = ""
        Form1.InputBox2.Text = ""
        Form1.InputBox3.Text = ""
        Form1.InputBox4.Text = ""
        Form1.OutputBox1.Text = "不能為空白"
        Form1.OutputBox2.Text = "不能為空白"
        Form1.OutputBox3.Text = "不能為空白"
        Form1.OutputBox4.Text = "不能為空白"
        Form1.ResultBox1.Text = ""
        Form1.ResultBox2.Text = ""
        Form1.ResultBox3.Text = ""
        Form1.ResultBox4.Text = ""
        Form1.Statistics1.Text = ""
        Form1.Statistics2.Text = ""
        Form1.Statistics3.Text = ""
        Form1.Statistics4.Text = ""
        Form1.ChkButton.Enabled = False
    End Sub
    ''' <summary>
    ''' 檢查資料是否合乎規則，回傳值為布林值
    ''' </summary>
    ''' <param name="InputBox">成績輸入資料，類型為TextBox</param>
    ''' <param name="OutputBox">錯誤訊息輸出資料，類型為Label</param>
    ''' <returns>回傳值為布林值</returns>
    Public Function ChkInput(ByVal InputBox As TextBox, ByVal OutputBox As Label) As Boolean
        Dim ErrorStr As String = ""
        If InputBox.Text.Length = 0 Then
            ErrorStr += "不能為空白"
        ElseIf Not IsNumeric(InputBox.Text) Then
            ErrorStr += "必須為數字"
        ElseIf CInt(InputBox.Text) < 1 Or CInt(InputBox.Text) > 100 Then
            ErrorStr += "必須介於1-100之間"
        End If
        OutputBox.Text = ErrorStr
        If ErrorStr.Length = 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 計算成績資料，回傳值為計算完畢的成績資料，資料結構為StudentData
    ''' </summary>
    ''' <param name="InputData">參數是資料結構為StudentData的成績資料</param>
    ''' <returns>回傳值為計算完畢的成績資料，資料結構為StudentData</returns>
    Public Function ChkMark(ByRef InputData As StudentData) As StudentData
        InputData.CAMarks = Math.Round((InputData.Test * 0.5) + (InputData.Project * 0.3) + (InputData.Quizzes * 0.2), 2)
        InputData.ModuleMarks = Math.Round((InputData.CAMarks * 0.4) + (InputData.Exam * 0.6), 2)
        InputData.ModlueGrade = ""
        If InputData.CAMarks < 40 Or InputData.Exam < 40 Then
            InputData.ModlueGrade = "F"
        ElseIf InputData.CAMarks >= 40 And InputData.Exam >= 40 And InputData.ModuleMarks >= 40 And InputData.ModuleMarks < 65 Then
            InputData.ModlueGrade = "C"
        ElseIf InputData.CAMarks >= 40 And InputData.Exam >= 40 And InputData.ModuleMarks >= 65 And InputData.ModuleMarks < 75 Then
            InputData.ModlueGrade = "B"
        ElseIf InputData.CAMarks >= 40 And InputData.Exam >= 40 And InputData.ModuleMarks >= 75 And InputData.ModuleMarks <= 100 Then
            InputData.ModlueGrade = "A"
        End If
        InputData.Remarks = ""
        If InputData.ModlueGrade = "A" Or InputData.ModlueGrade = "B" Or InputData.ModlueGrade = "C" Then
            InputData.Remarks = "Pass"
        ElseIf InputData.ModlueGrade = "F" And InputData.ModuleMarks >= 30 Then
            InputData.Remarks = "Re - sit Exam"
        ElseIf InputData.ModlueGrade = "F" And InputData.ModuleMarks < 30 Then
            InputData.Remarks = "Restudy"
        End If
        Return InputData
    End Function
    ''' <summary>
    ''' 送出SQL要求，回傳值為帶索引的陣列
    ''' </summary>
    ''' <param name="SQLStr">SQL要求，為一個SQL命令字串</param>
    ''' <returns>回傳值為帶索引的陣列</returns>
    Public Function SqlQuery(ByVal SQLStr As String) As Hashtable()
        On Error Resume Next
        Dim cn As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\db.mdb")
        Dim adp As OleDbDataAdapter = New OleDbDataAdapter(SQLStr, cn)
        Dim Dset As New DataSet
        adp.Fill(Dset, "res")
        Dim ResCounts As Integer = Dset.Tables("res").Rows.Count
        Dim MyIndex As Integer = Dset.Tables("res").Columns.Count
        Dim Res(ResCounts - 1) As Hashtable
        For i = 0 To ResCounts - 1
            Res(i) = New Hashtable
            For j = 0 To MyIndex - 1
                Dim n_index = Dset.Tables("res").Columns(j).ToString
                Res(i).Add(n_index, Dset.Tables("res").Rows(i).Item(j))
            Next j
        Next i
        adp.Dispose()
        cn.Close()
        Return Res
    End Function
End Module
