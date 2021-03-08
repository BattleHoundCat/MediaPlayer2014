Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class CheckForm
    Dim xlApp As Excel.Application = New Excel.Application
    Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open("f:\Хаос\MusicList.xls")
    Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Worksheets("Лист1")
    Dim musicname As String
    Dim sourcemusic As String
    Dim Category As String
    Private Sub Yepbtn_Click(sender As Object, e As EventArgs) Handles Yepbtn.Click
            If CheckDir(TextBox1.Text) Then
                Category = "Классно"
                sourcemusic = Form1.wmp.currentMedia.sourceURL
                musicname = IO.Path.GetFileName(Form1.wmp.currentMedia.sourceURL)
                'если есть директории в папке Классно
                Dim pathes() As String
                pathes = Directory.GetDirectories(TextBox1.Text)
            For i = 0 To pathes.Length - 1
                Dim DI As New IO.DirectoryInfo(pathes(i))
                ComboBox1.Items.Add(DI.Name)
            Next
            Label1.Text = pathes.Length
        Else
            'если их нет
            Category = "Классно"
            sourcemusic = Form1.wmp.currentMedia.sourceURL
            musicname = IO.Path.GetFileName(Form1.wmp.currentMedia.sourceURL)
        End If
        Fonbtn.Enabled = False
        Yepbtn.Enabled = False
        ConfirmBtn.Enabled = True
        ComboBox1.Enabled = True
    End Sub

    Private Sub Fonbtn_Click(sender As Object, e As EventArgs) Handles Fonbtn.Click
            If CheckDir(TextBox2.Text) Then
                Category = "Фон"
                sourcemusic = Form1.wmp.currentMedia.sourceURL
                musicname = IO.Path.GetFileName(Form1.wmp.currentMedia.sourceURL)
                'если есть директории в папке Классно
                Dim pathes() As String
                pathes = Directory.GetDirectories(TextBox2.Text)
                For i = 0 To pathes.Length - 1
                    Dim DI As New IO.DirectoryInfo(pathes(i))
                    ComboBox1.Items.Add(DI.Name)
            Next
            Label1.Text = pathes.Length
            Else
                'если их нет
                Category = "Фон"
                sourcemusic = Form1.wmp.currentMedia.sourceURL
                musicname = IO.Path.GetFileName(Form1.wmp.currentMedia.sourceURL)
        End If
        Fonbtn.Enabled = False
        Yepbtn.Enabled = False
        ConfirmBtn.Enabled = True
        ComboBox1.Enabled = True
    End Sub

    Private Sub Delbtn_Click(sender As Object, e As EventArgs) Handles Delbtn.Click
        File.Delete(Form1.wmp.currentMedia.sourceURL)
        MsgBox("Эта хреновая музыка удалена !" & vbCrLf & Form1.wmp.currentMedia.sourceURL)
        Me.Close()
    End Sub

    Private Sub CheckForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        TextBox1.Text = "D:\Новая жизнь\Новое\Хобби\Развлечения\Музыка\Автоматом\Классно\"
        TextBox2.Text = "D:\Новая жизнь\Новое\Хобби\Развлечения\Музыка\Автоматом\Фон\"
        ConfirmBtn.Enabled = False
    End Sub
    Private Sub MoveFile(ByVal sourceFile As String, ByVal FileName As String, ByVal path As String, ByVal Category As String, ByVal localpath As String)
        If File.Exists(sourceFile) Then
            'если исходный файл существует,то перемещаем его в нужную папку
            If Directory.Exists(path & localpath) Then
                'если есть директория,то проверяем файл
                If File.Exists(path & "\" & localpath & "\" & FileName) Then
                    'здесь типа проверка,существует ли подобный файл с подобным названием в уже целевой папке
                    MsgBox("Файл с таким названием уже существует,он будет переименован")
                    Dim i As Integer = Form1.Random(1, 100000)

                    WriteExcel(i & FileName, localpath, Category, path & localpath)
                    File.Move(sourceFile, path & "\" & localpath & "\" & i & FileName)
                Else
                    WriteExcel(FileName, localpath, Category, path & localpath)
                    File.Move(sourceFile, path & "\" & localpath & "\" & FileName)
                End If
            Else
                'если ее нет,то создаем
                Directory.CreateDirectory(path & localpath)
                WriteExcel(FileName, localpath, Category, path & localpath)
                File.Move(sourceFile, path & localpath & "\" & FileName)
            End If

        End If
    End Sub
    Public Sub WriteExcel(ByVal filename As String, ByVal artist As String, ByVal Category As String, ByVal finishPath As String)
        Dim cell_text As Integer
        Dim xlCells As Excel.Range = Nothing
        Dim range As Excel.Range = Nothing
        xlCells = xlWorkSheet.Cells(1, 1)
        cell_text = Val(xlCells.Value)
        Dim i As Integer = 0
        If cell_text > 0 Then
            'записываем в след.колонку
            xlWorkSheet.Cells(cell_text + 3, 1).Value = cell_text + 1
            xlWorkSheet.Cells(cell_text + 3, 2).Value = filename
            xlWorkSheet.Cells(cell_text + 3, 3).Value = artist
            xlWorkSheet.Cells(cell_text + 3, 4).Value = Category
            xlWorkSheet.Cells(cell_text + 3, 5).Value = finishPath
            xlWorkBook.Save()
        Else
            'иначе записываем в первую колонку
            xlWorkSheet.Cells(i + 3, 1).Value = 1
            xlWorkSheet.Cells(i + 3, 2).Value = filename
            xlWorkSheet.Cells(i + 3, 3).Value = artist
            xlWorkSheet.Cells(i + 3, 4).Value = Category
            xlWorkSheet.Cells(i + 3, 5).Value = finishPath
            xlWorkBook.Save()
        End If
    End Sub
    Private Sub Form1_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        xlWorkBook.Close()
        xlApp.Quit()
        'xlApp.Dispose()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
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
    Private Function CheckDir(ByVal targetDir As String) As Boolean
        If Directory.Exists(targetDir) Then
            Dim str() As String
            str = Directory.GetDirectories(targetDir)
            If str.Length > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Directory.CreateDirectory(targetDir)
        End If
    End Function

    Private Sub ConfirmBtn_Click(sender As Object, e As EventArgs) Handles ConfirmBtn.Click
        If Category <> "" Then
            If Category = "Классно" Then
                If Not IsNothing(ComboBox1.SelectedItem) Then
                    MoveFile(sourcemusic, musicname, TextBox1.Text, Category, ComboBox1.SelectedItem.ToString)
                    Me.Close()
                Else
                    If ComboBox1.Text <> "" Then
                        MoveFile(sourcemusic, musicname, TextBox1.Text, Category, ComboBox1.Text)
                        Me.Close()
                    Else
                        MsgBox("Выберите папку или напишите ее в названии")
                    End If
                End If
            Else
                If Not IsNothing(ComboBox1.SelectedItem) Then
                    MoveFile(sourcemusic, musicname, TextBox2.Text, Category, ComboBox1.SelectedItem.ToString)
                    Me.Close()
                Else
                    If ComboBox1.Text <> "" Then
                        MoveFile(sourcemusic, musicname, TextBox2.Text, Category, ComboBox1.Text)
                        Me.Close()
                    Else
                        MsgBox("Выберите папку или напишите ее в названии")
                    End If
                End If
            End If
            
        Else
            MsgBox("Выберите категорию музыки:Классно,Фон,или Удалить")
        End If
    End Sub

    Private Sub ClearBtn_Click(sender As Object, e As EventArgs) Handles ClearBtn.Click
        ComboBox1.Items.Clear()
        ComboBox1.Text = ""
        ConfirmBtn.Enabled = False
        ComboBox1.Enabled = False
        Fonbtn.Enabled = True
        Yepbtn.Enabled = True
        Label1.Text = 0
    End Sub
End Class