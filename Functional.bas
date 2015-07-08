Attribute VB_Name = "Functional"
Sub Import()
'
' Макрос2 Макрос
'

'
    Dim FilePath As String
    FilePath = GetFilePath()
    Workbooks.OpenText FileName:=FilePath _
        , Origin:=1251, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, Tab:=True, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(11, 1)), TrailingMinusNumbers:=True
        
    FreezeHead
End Sub

Private Function GetFilePath() As String
    Dim FileName As String
    
    'FileName = Dir(ThisWorkbook.Path & "\*.txt")
    If FileName <> "" Then
        ' Если в текущей дирректории нет других txt файлов
        If Dir() = "" Then GetFilePath = ThisWorkbook.Path & "\" & FileName
    Else
        Application.DefaultFilePath = ThisWorkbook.Path
        ChDir (ThisWorkbook.Path)
        GetFilePath = Application.GetOpenFilename("Firework data in TXT-format (*.txt), *.txt", , "Choose a Firework file (TXT)")
    End If

End Function

Sub LoadFireworks()
Attribute LoadFireworks.VB_ProcData.VB_Invoke_Func = "r\n14"
    Misc.ClearCells
    Dim oSalut As New clsSalut
  
    oSalut.Title = Workbooks(2).ActiveSheet.Name
    
    Dim i As Integer
    
    i = 2
    Do
        Set oFirework = New clsFirework
        
        
        oFirework.Load (Workbooks(2).ActiveSheet.Cells(i, 1))
        oSalut.AddFirework (oFirework)
        
        i = i + 1
        
    Loop Until IsEmpty(Workbooks(2).ActiveSheet.Cells(i, 1))
 
    oSalut.Display (Workbooks(1).Worksheets(1).Cells(1, 1))
    
    Debug.Print " "
    
    
End Sub


Sub FreezeHead()
'
' Закрепление верхней строки с заголовками
'

'
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub

Sub CountMax()

' Определяет максимальное количество зарядов

    Dim i As Integer
    i = 0
    Do
        i = i + 1
        
    Loop Until IsEmpty(Workbooks(2).ActiveSheet.Cells(i, 1))
    
    Debug.Print i
    
    
End Sub


Private Sub Import2()
'
' Макрос1 Макрос
'

'
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\timur.tatarshaov\Documents\fireworks\80000 Янченко сбербанк 2010 3_54.txt" _
        , Destination:=Range("$A$1"))
        .Name = "80000 Янченко сбербанк 2010 3_54"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1251
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub


