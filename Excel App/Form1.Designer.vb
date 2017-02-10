Imports System
Imports Microsoft.Office.Interop.Excel


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
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 261)
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

        'Открыть новую книгу Excel
        Dim _Excel As New Application 'Приложение Excel
        Dim Книга As Workbook = _Excel.Workbooks.Open("c:\Intel\test.xlsx") 'Открываем книгу
        Dim Лист As Worksheet = CType(Книга.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet) 'Первый лист книги

        Dim NumRows As Integer = Лист.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row

        ' MsgBox(Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row)


        'Создать массив с 3 столбцами и 100 строками
        Dim DataArray(NumRows - 1, 2) As Object
        Dim upper0 As Int32 = DataArray.GetUpperBound(0)
        Dim lower0 As Int32 = DataArray.GetLowerBound(0)
        Dim upper1 As Int32 = DataArray.GetUpperBound(1)
        Dim lower1 As Int32 = DataArray.GetLowerBound(1)




        MsgBox("DataArray Rank: " + DataArray.Rank.ToString + vbCrLf + "lower0: " + lower0.ToString + vbCrLf + "lower1: " + lower1.ToString + vbCrLf + "upper0: " + upper0.ToString + vbCrLf + "upper1: " + upper1.ToString + vbCrLf + "NumRows :" + NumRows.ToString + vbCrLf + "Ubound DataArray(1): " + UBound(DataArray, 1).ToString + vbCrLf + "Ubound DataArray(2): " + UBound(DataArray, 2).ToString + vbCrLf + "Lbound DataArray(1): " + LBound(DataArray, 1).ToString + vbCrLf + "Lbound DataArray(2): " + LBound(DataArray, 2).ToString)


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
        DataArray = Лист.Range("A2").Resize(NumRows - 1, 3).Value

        upper0 = DataArray.GetUpperBound(0)
        lower0 = DataArray.GetLowerBound(0)
        upper1 = DataArray.GetUpperBound(1)
        lower1 = DataArray.GetLowerBound(1)

        MsgBox("DataArray Rank: " + DataArray.Rank.ToString + vbCrLf + "lower0: " + lower0.ToString + vbCrLf + "lower1: " + lower1.ToString + vbCrLf + "upper0: " + upper0.ToString + vbCrLf + "upper1: " + upper1.ToString + vbCrLf + "NumRows :" + NumRows.ToString + vbCrLf + "Ubound DataArray(1): " + UBound(DataArray, 1).ToString + vbCrLf + "Ubound DataArray(2): " + UBound(DataArray, 2).ToString + vbCrLf + "Lbound DataArray(1): " + LBound(DataArray, 1).ToString + vbCrLf + "Lbound DataArray(2): " + LBound(DataArray, 2).ToString)

        'For i = 1 To upper0
        'For j = 1 To 3
        'MsgBox("DataArray(" + Str(i) + "," + Str(j) + ") :" + DataArray(i, j))
        'Next
        'Next

        'Сохранить книгу и закрыть Excel
        'Книга.Save(True)
        Книга.Close(True)

        Лист = Nothing
        Книга = Nothing
        _Excel.Quit()
        _Excel = Nothing
        GC.Collect()


    End Sub
End Class
