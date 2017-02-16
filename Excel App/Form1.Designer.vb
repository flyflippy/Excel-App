Imports System
Imports System.IO
Imports Microsoft.Office.Interop.Excel
'Imports System.Diagnostics


<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnGetFiles = New System.Windows.Forms.Button()
        Me.btnBench = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(18, 12)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(245, 20)
        Me.TextBox1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(18, 38)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(245, 32)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "write txt to b5"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(18, 76)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(245, 32)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "прочитать b5 в txt"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(18, 114)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(245, 32)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Заполнить 1 лист"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(18, 152)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(245, 32)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Текст в Массив"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(291, 27)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(361, 20)
        Me.TextBox2.TabIndex = 5
        Me.TextBox2.Text = "1604 2RS(ZEN)"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(288, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(176, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Текстовая строка для обработки"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(289, 98)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Выходная строка 1"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(291, 114)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(361, 20)
        Me.TextBox3.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(289, 137)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Выходная строка 2"
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(292, 153)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(360, 20)
        Me.TextBox4.TabIndex = 10
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(292, 218)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(362, 21)
        Me.Button5.TabIndex = 11
        Me.Button5.Text = "Строки"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(288, 57)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(101, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Условная строка1"
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(389, 53)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(263, 20)
        Me.TextBox5.TabIndex = 13
        Me.TextBox5.Text = "("
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(389, 76)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(263, 20)
        Me.TextBox6.TabIndex = 15
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(288, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(101, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Условная строка2"
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(292, 192)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(360, 20)
        Me.TextBox7.TabIndex = 17
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(289, 176)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(104, 13)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Выходная строка 3"
        '
        'btnGetFiles
        '
        Me.btnGetFiles.Location = New System.Drawing.Point(18, 187)
        Me.btnGetFiles.Name = "btnGetFiles"
        Me.btnGetFiles.Size = New System.Drawing.Size(245, 29)
        Me.btnGetFiles.TabIndex = 18
        Me.btnGetFiles.Text = "Обработать в папке"
        Me.btnGetFiles.UseVisualStyleBackColor = True
        '
        'btnBench
        '
        Me.btnBench.Location = New System.Drawing.Point(291, 245)
        Me.btnBench.Name = "btnBench"
        Me.btnBench.Size = New System.Drawing.Size(363, 28)
        Me.btnBench.TabIndex = 19
        Me.btnBench.Text = "Бенчмарк"
        Me.btnBench.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(696, 313)
        Me.Controls.Add(Me.btnBench)
        Me.Controls.Add(Me.btnGetFiles)
        Me.Controls.Add(Me.TextBox7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button

    Private Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        Dim _Excel As New Application 'Приложение Excel
        Dim Книга As Workbook = _Excel.Workbooks.Open("c:\Intel\test.xlsx") 'Открываем книгу
        Dim Лист As Worksheet = CType(Книга.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet) 'Первый лист книги

        Try
            Лист.Range("B5").Value = TextBox1.Text 'запись в ячейку
            _Excel.Workbooks(1).Save()
            _Excel.Workbooks(1).Close() ' сохранение и закрытие


            Лист = Nothing
            Книга = Nothing
            _Excel.Quit()
            _Excel = Nothing
        Catch ex As Exception ' обработка ошибок
            MsgBox(ex.ToString, Microsoft.VisualBasic.MsgBoxStyle.Critical, "В программе произошла ошибка") ' описание ошибки
            _Excel.Workbooks(1).Save()
            _Excel.Workbooks(1).Close() ' в случаи ошибки экземпляр Excel повиснит в памяти, а так мы закроем процесс не сохраняя

            Лист = Nothing
            Книга = Nothing
            _Excel.Quit()
            _Excel = Nothing
        End Try
        GC.Collect()
    End Sub



    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click

        Dim _Excel As New Application 'Приложение Excel
        Dim Книга As Workbook = _Excel.Workbooks.Open("c:\Intel\test.xlsx") 'Открываем книгу
        Dim Лист As Worksheet = CType(Книга.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet) 'Первый лист книги


        Try
            TextBox1.Text = Лист.Range("B5").Value ' читаем
            Книга.Close(SaveChanges:=False)


            Лист = Nothing
            Книга = Nothing
            _Excel.Quit()
            _Excel = Nothing

        Catch ex As Exception
            MsgBox(ex.ToString, Microsoft.VisualBasic.MsgBoxStyle.Critical, "В программе произошла ошибка") ' описание ошибки
            _Excel.Workbooks(1).Close(SaveChanges:=False) ' в случаи ошибки экземпляр Excel повиснит в памяти, а так мы закроем процесс не сохраняя

            Лист = Nothing
            Книга = Nothing
            _Excel.Quit()
            _Excel = Nothing
        End Try
        GC.Collect()

    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click



        'Открыть новую книгу Excel
        Dim _Excel As New Application 'Приложение Excel
        Dim Книга As Workbook = _Excel.Workbooks.Open("c:\Intel\test.xlsx") 'Открываем книгу
        Dim Лист As Worksheet = CType(Книга.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet) 'Первый лист книги

        Dim NumRows As Integer = 3

        MsgBox(Лист.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row)

        'Создать массив с 3 столбцами и 100 строками
        Dim DataArray(0 To NumRows, 0 To 2) As String
        Dim r As Integer
        For r = 0 To NumRows
            DataArray(r, 0) = "ID" & Format(r, "0000")
            DataArray(r, 1) = CInt(Int((6 * Rnd()) + 1))
            DataArray(r, 2) = DataArray(r, 1) * 0.7
        Next

        'Добавить заголовки в строку 1
        Лист.Range("A1:C1").Value = {"Первый столбец", "Второй стобец", "Третий столбец"}

        'Передать массив на лист, начиная с ячейки A2
        Лист.Range("A2").Resize(NumRows + 1, 3).Value = DataArray

        MsgBox("prompt2")

        'Сохранить книгу и закрыть Excel
        Книга.Close(True)


        Лист = Nothing
        Книга = Nothing
        _Excel.Quit()
        _Excel = Nothing
        GC.Collect()
    End Sub

    Friend WithEvents Button4 As System.Windows.Forms.Button

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ProcessManyXLS("c:\intel\test.xlsx")
    End Sub

    Private Sub ProcessManyXLS(filePath As String)




        'Dim replacements(,) As String = {{"КАТАЛОГ FAG", "Заменено FAG"}, {"КАТАЛОГ SKF", "Заменено SKF"}, {"курукузя", "каракузя"}}

        'MsgBox(" dim 0: " + replacements.GetUpperBound(0).ToString + " dim1: " + replacements.GetUpperBound(1).ToString)


        'For i = 0 To replacements.GetUpperBound(0)


        '    MsgBox(replacements(i, 0) + " ///  " + replacements(i, 1))
        'Next

        'End


        'Открыть новую книгу Excel
        Dim _Excel As New Application 'Приложение Excel
        _Excel.DisplayAlerts = False
        Dim Книга As Workbook
        Dim Лист As Worksheet

        Dim NumRows As Integer

        ' MsgBox(Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row)



        Книга = _Excel.Workbooks.Open(filePath) 'Открываем книгу
        Лист = CType(Книга.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet) 'Первый лист книги
        NumRows = Лист.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row
        'Создать массив с 3 столбцами и 100 строками
        Dim DataArray(NumRows - 1, 1) As Object
        Dim lastArray(NumRows, 1) As String

        Dim upper0 As Int32 = DataArray.GetUpperBound(0)
        Dim lower0 As Int32 = DataArray.GetLowerBound(0)
        Dim upper1 As Int32 = DataArray.GetUpperBound(1)
        Dim lower1 As Int32 = DataArray.GetLowerBound(1)



        'MsgBox("DataArray Rank: " + DataArray.Rank.ToString + vbCrLf + "lower0: " + lower0.ToString + vbCrLf + "lower1: " + lower1.ToString + vbCrLf + "upper0: " + upper0.ToString + vbCrLf + "upper1: " + upper1.ToString + vbCrLf + "NumRows :" + NumRows.ToString + vbCrLf + "Ubound DataArray(1): " + UBound(DataArray, 1).ToString + vbCrLf + "Ubound DataArray(2): " + UBound(DataArray, 2).ToString + vbCrLf + "Lbound DataArray(1): " + LBound(DataArray, 1).ToString + vbCrLf + "Lbound DataArray(2): " + LBound(DataArray, 2).ToString)


        ' Dim r As Integer
        'For r = 0 To NumRows
        'DataArray(r, 0) = "ID" & Format(r, "0000")
        'DataArray(r, 1) = CInt(Int((6 * Rnd()) + 1))
        'DataArray(r, 2) = DataArray(r, 1) * 0.7
        'Next

        'MsgBox("prompt1")



        'Добавить заголовки в строку 1
        'Лист.Range("A1:C1").Value = {"Первый столбец", "Второй стобец", "Третий столбец"}

        'Передать массив на лист, начиная с ячейки A2
        DataArray = Лист.Range("A2").Resize(NumRows - 1, 1).Value


        upper0 = DataArray.GetUpperBound(0)
        lower0 = DataArray.GetLowerBound(0)
        upper1 = DataArray.GetUpperBound(1)
        lower1 = DataArray.GetLowerBound(1)

        MsgBox("DataArray Rank: " + DataArray.Rank.ToString + vbCrLf + "lower0: " + lower0.ToString + vbCrLf + "lower1: " + lower1.ToString + vbCrLf + "upper0: " + upper0.ToString + vbCrLf + "upper1: " + upper1.ToString + vbCrLf + "NumRows :" + NumRows.ToString + vbCrLf + "Ubound DataArray(1): " + UBound(DataArray, 1).ToString + vbCrLf + "Ubound DataArray(2): " + UBound(DataArray, 2).ToString + vbCrLf + "Lbound DataArray(1): " + LBound(DataArray, 1).ToString + vbCrLf + "Lbound DataArray(2): " + LBound(DataArray, 2).ToString)


        '================================

        Dim tempstr As String
        Dim xstr As Char
        Dim finstr As String = ""
        Dim finstr2 As String = ""
        Dim x, lenstr As Integer


        xstr = "("
        MsgBox("Пошла обработка массива")
        For i = lower0 To upper0
            tempstr = DataArray(i, 1)

            'x = InStrRev(tempstr, xstr, , CompareMethod.Text)
            'Оптимизировано по скорости
            x = tempstr.LastIndexOf(xstr) + 1


            lenstr = Len(tempstr)

            finstr = Mid(tempstr, x + 1, lenstr - x - 1)

            If x <> 0 Then
                finstr = Mid(tempstr, x + 1, lenstr - x - 1)
                finstr2 = RTrim(Strings.Left(tempstr, x - 1))
            Else

                'finstr.IndexOf()

                finstr2 = Trim(tempstr)
                finstr = ""


            End If



            lastArray(i - 1, 1) = finstr
            lastArray(i - 1, 0) = finstr2
            'MsgBox("finstr:" + finstr + "finstr2:" + finstr2)

        Next


        ' Запихивание в лист

        'Передать массив на лист, начиная с ячейки A2
        Лист.Range("E2").Resize(NumRows + 1, 2).Value = lastArray


        TextBox7.Text = x.ToString
        TextBox3.Text = finstr
        TextBox4.Text = finstr2





        '===============================



        'For i = 1 To upper0
        'For j = 1 To 3
        'MsgBox("DataArray(" + Str(i) + "," + Str(j) + ") :" + DataArray(i, j))
        'Next
        'Next

        'Сохранить книгу и закрыть Excel
        Книга.Close(True)


        ' Catch ex As Exception
        'MsgBox("Exeptions" & vbCrLf & ex.Message)



        Лист = Nothing
        Книга = Nothing
        _Excel.Quit()
        _Excel = Nothing
        GC.Collect()

        ' MsgBox("Finaly")


    End Sub

    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Button5 As System.Windows.Forms.Button

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim tempstr As String
        Dim xstr As Char
        Dim finstr As String
        Dim finstr2 As String
        Dim x, lenstr As Integer

        tempstr = TextBox2.Text
        xstr = TextBox5.Text

        x = InStrRev(tempstr, xstr, , CompareMethod.Text)
        lenstr = Len(tempstr)

        'finstr = Mid(tempstr, x + 1, lenstr - x - 1)

        finstr = tempstr.Substring(x, lenstr - x - 1) 'оптимизировано вместо mid


        finstr2 = RTrim(Strings.Left(tempstr, x - 1))

        TextBox7.Text = x.ToString
        TextBox3.Text = finstr
        TextBox4.Text = finstr2


    End Sub

    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnGetFiles As System.Windows.Forms.Button

    Private Sub btnGetFiles_Click(sender As Object, e As EventArgs) Handles btnGetFiles.Click

        'Try
        ' Only get files that begin with the letter "c."
        Dim dirs As String() = Directory.GetFiles("c:\intel\xlsx", "*.xlsx")
        'Debug.WriteLine("The number of files starting with c is {0}.", dirs.Length)
        Dim dir As String
        For Each dir In dirs
            'Debug.WriteLine(dir)
            ProcessManyXLS(dir)
        Next
        'Catch ex As Exception
        'Debug.WriteLine("The process failed: {0}", ex.ToString())
        'End Try

    End Sub

    Friend WithEvents btnBench As System.Windows.Forms.Button

    Private Sub btnBench_Click(sender As Object, e As EventArgs) Handles btnBench.Click

        Dim sb As New Text.StringBuilder()
        Dim benchString As String = "1604 2RS(ZEN)"
        Dim tempStr As String
        Dim benchChar As Char = "("
        Dim x As Integer
        Dim startTime, endTime As DateTime


        startTime = DateTime.Now


        For i = 0 To 50000000


            'x = InStrRev(benchString, benchChar, , CompareMethod.Text)
            'x = InStr(benchString, benchChar, CompareMethod.Text)
            'x = benchString.IndexOf(benchChar)
            'x = benchString.LastIndexOf(benchChar)

            'x = Len(benchString)
            'x = benchString.Length

            'tempStr = Trim(benchString)
            'tempStr = benchString.Trim

            'tempStr = RTrim(benchString)
            'tempStr = benchString.TrimEnd

            'Mid benchmark
            'tempStr = Mid(benchString, 4, 5) 'debug 641  release 706
            'tempStr = benchString.Substring(4 - 1, 5) 'debug 515 release 564



        Next
        endTime = DateTime.Now

        MsgBox("Miliseconds: " & (endTime - startTime).TotalMilliseconds.ToString & vbCrLf & "x=" & (x).ToString & vbCrLf & "tempStr:" & tempStr)


    End Sub
End Class
