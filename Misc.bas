Attribute VB_Name = "Misc"
Public Sub Prepare()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
End Sub

Public Sub Ended()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
End Sub

Public Function InCollection(ByVal col As Collection, key As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long

  InCollection = False
  Set var = Nothing

  Err.Clear
  On Error Resume Next
    var = col.Item(key)
    errNumber = CLng(Err.Number)
  On Error GoTo 0

  '5 is not in, 0 and 438 represent incollection
  If errNumber = 5 Then ' it is 5 if not in collection
    InCollection = False
  Else
    InCollection = True
  End If

End Function

Sub ClearCells()
Attribute ClearCells.VB_ProcData.VB_Invoke_Func = "t\n14"

    Cells.Select
    'Range("A51").Activate
    Selection.Delete Shift:=xlUp
End Sub

Public Sub MergeCells(ByRef rngCurrentCell As Range)

    With rngCurrentCell
             .Merge
             .NumberFormat = "0.00"
             .Font.Bold = True
             .HorizontalAlignment = xlCenter
             .Select
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
        End With

End Sub

Static Function SheafCounter(Optional ByVal btSheafCounter As Integer = -1)

    Dim m_btSheafCounter As Integer
    
    If btSheafCounter <> -1 Then _
        m_btSheafCounter = btSheafCounter
        
    SheafCounter = m_btSheafCounter
    
End Function
