Public Class Form1
    ''' <summary>
    ''' 存放新輸入的資料，資料型態為StudentData
    ''' </summary>
    Private ThisStudent As StudentData
    ''' <summary>
    ''' 存放資料庫裡的所有資料，資料型態為StudentData
    ''' </summary>
    Private SqlStudent() As StudentData
    ''' <summary>
    ''' 資料庫所有分數的總和，用來計算平均分數
    ''' </summary>
    Private AllModuleMarks As Double = 0
    ''' <summary>
    ''' 評分為A的計數器
    ''' </summary>
    Private CountA As Integer = 0
    ''' <summary>
    ''' 評分為F的計數器
    ''' </summary>
    Private CountF As Integer = 0

    Private Sub ChkErrorStr()
        '檢查輸入的文字方塊是否符合規則
        Dim ErrorStr As String = OutputBox1.Text & OutputBox2.Text & OutputBox3.Text & OutputBox4.Text
        If TextBox1.Text.Length = 0 Then
            ChkButton.Enabled = False
            Exit Sub
        End If
        If ErrorStr.Length = 0 Then
            ChkButton.Enabled = True
        Else
            ChkButton.Enabled = False
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '開啟程式時，清除所有文字方塊的資料，並且讀取資料庫取得資料庫裡的數據
        CleanAll()
        UpSqlData()
    End Sub
    Private Sub InputBox1_TextChanged(sender As Object, e As EventArgs) Handles InputBox1.TextChanged
        '成績輸入方塊有變化時，檢查最新的輸入是否合乎規則
        ChkInput(InputBox1, OutputBox1)
        ChkErrorStr()
    End Sub
    Private Sub InputBox2_TextChanged(sender As Object, e As EventArgs) Handles InputBox2.TextChanged
        '成績輸入方塊有變化時，檢查最新的輸入是否合乎規則
        ChkInput(InputBox2, OutputBox2)
        ChkErrorStr()
    End Sub
    Private Sub InputBox3_TextChanged(sender As Object, e As EventArgs) Handles InputBox3.TextChanged
        '成績輸入方塊有變化時，檢查最新的輸入是否合乎規則
        ChkInput(InputBox3, OutputBox3)
        ChkErrorStr()
    End Sub
    Private Sub InputBox4_TextChanged(sender As Object, e As EventArgs) Handles InputBox4.TextChanged
        '成績輸入方塊有變化時，檢查最新的輸入是否合乎規則
        ChkInput(InputBox4, OutputBox4)
        ChkErrorStr()
    End Sub
    Private Sub InputBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles InputBox1.KeyPress, InputBox2.KeyPress, InputBox3.KeyPress, InputBox4.KeyPress
        '限制四個成績輸入方塊只允許輸入數字跟BackSpace鍵
        If Char.IsDigit(e.KeyChar) Or e.KeyChar = vbBack Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub ChkButton_Click(sender As Object, e As EventArgs) Handles ChkButton.Click
        '確認輸入成績資料，並計算成績，寫入資料庫
        If TextBox1.Text.Length = 0 Then
            MsgBox("必須輸入姓名")
        End If
        ThisStudent.Name = TextBox1.Text
        ThisStudent.Test = CInt(InputBox1.Text)
        ThisStudent.Project = CInt(InputBox2.Text)
        ThisStudent.Quizzes = CInt(InputBox3.Text)
        ThisStudent.Exam = CInt(InputBox4.Text)
        ChkMark(ThisStudent)
        ResultBox1.Text = ThisStudent.CAMarks
        ResultBox2.Text = ThisStudent.ModlueGrade
        ResultBox3.Text = ThisStudent.ModuleMarks
        ResultBox4.Text = ThisStudent.Remarks
        SqlQuery("INSERT INTO student (name,test,project,quizzes,exam) VALUES ('" & ThisStudent.Name & "'," & ThisStudent.Test & "," & ThisStudent.Project & "," & ThisStudent.Quizzes & "," & ThisStudent.Exam & ")")
        UpSqlData()
    End Sub
    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        '點選Listbox時，即時顯示出該同學的各科成績
        Dim SelectIndex As Integer = ListBox1.SelectedIndex
        TextBox1.Text = SqlStudent(SelectIndex).Name
        InputBox1.Text = SqlStudent(SelectIndex).Test
        InputBox2.Text = SqlStudent(SelectIndex).Project
        InputBox3.Text = SqlStudent(SelectIndex).Quizzes
        InputBox4.Text = SqlStudent(SelectIndex).Exam
        ResultBox1.Text = SqlStudent(SelectIndex).CAMarks
        ResultBox2.Text = SqlStudent(SelectIndex).ModlueGrade
        ResultBox3.Text = SqlStudent(SelectIndex).ModuleMarks
        ResultBox4.Text = SqlStudent(SelectIndex).Remarks
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '尋找同學姓名，當有找到符合資料，反白該同學，並列出成績資料，若無符合資料，利用Msgbox來提示
        If FindBox1.Text.Length = 0 Then Exit Sub
        For i As Integer = 0 To SqlStudent.Count - 1
            If FindBox1.Text = SqlStudent(i).Name Then
                ListBox1.SelectedIndex = i
                ListBox1_Click(sender, e)
                Exit Sub
            End If
        Next i
        MsgBox("NotFound!!")
    End Sub
    ''' <summary>
    ''' 讀取資料庫資料，並將資料放在變數SqlStudent()裡面
    ''' </summary>
    Private Sub UpSqlData()
        Dim Res = SqlQuery("SELECT * FROM student ORDER BY name ASC")
        If Res.Count = 0 Then
            ReDim SqlStudent(0)
            Exit Sub
        End If
        ReDim SqlStudent(Res.Count - 1)
        ListBox1.Items.Clear()
        For i As Integer = 0 To Res.Count - 1
            ListBox1.Items.Add(Res(i)("name"))
            SqlStudent(i).Name = Res(i)("name")
            SqlStudent(i).Test = Res(i)("test")
            SqlStudent(i).Project = Res(i)("project")
            SqlStudent(i).Quizzes = Res(i)("quizzes")
            SqlStudent(i).Exam = Res(i)("exam")
            ChkMark(SqlStudent(i))
        Next i
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '刪除所有資料，先做vbYesNo重複確認，避免不小心刪除資料
        Dim Ans = MsgBox("是否刪除資料庫所有資料", vbYesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Exclamation, "刪除資料")
        If Ans = vbYes Then
            '刪除資料
            SqlQuery("DELETE * FROM student")
            CleanAll()
            UpSqlData()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '計算統計資料
        AllModuleMarks = 0
        CountA = 0
        CountF = 0
        For i As Integer = 0 To SqlStudent.Count - 1
            AllModuleMarks += SqlStudent(i).ModuleMarks
            If SqlStudent(i).ModlueGrade = "A" Then
                CountA += 1
            ElseIf SqlStudent(i).ModlueGrade = "F" Then
                CountF += 1
            End If
        Next i
        Statistics1.Text = SqlStudent.Count
        If SqlStudent.Count = 0 Then
            Statistics2.Text = "0"
        Else
            Statistics2.Text = Math.Round(AllModuleMarks / SqlStudent.Count, 2)
        End If
        Statistics3.Text = CountA
        Statistics4.Text = CountF
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress, FindBox1.KeyPress
        '限制輸入姓名格式，只允許輸入字母、BackSpace、空白鍵
        If Char.IsLetter(e.KeyChar) Or e.KeyChar = vbBack Or e.KeyChar = " " Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        '姓名輸入方塊有變化時，檢查最新的輸入是否合乎規則
        ChkErrorStr()
    End Sub

End Class
