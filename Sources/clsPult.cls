VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_btNumber As Byte
Private m_clsLines(0 To 8) As clsLine
Private Const m_rowOffset As Byte = 5
Private Const m_columnOffset As Byte = 0


Public Property Get Number() As Byte

    Number = m_btNumber

End Property

Public Property Let Number(ByVal btNumber As Byte)
    
    m_btNumber = btNumber

End Property

Property Get Lines(Index As Integer) As clsLine
    
    Set Lines = m_clsLines(Index)
    
End Property

Property Let Lines(Index As Integer, ByVal Line As clsLine)
    
    If Index < LBound(m_clsLines) Or Index > UBound(m_clsLines) Then Err.Raise 9
    Set m_clsLines(Index) = Line
    
End Property

Public Property Get RowOffset() As Byte

    Dim Index As Integer
    For Index = 1 To UBound(m_clsLines)
        If Me.Lines(Index).Filled <> False Then
          RowOffset = RowOffset + Me.Lines(Index).RowOffset
        End If
    Next Index

End Property

Public Property Get ColumnOffset() As Byte
        
    ColumnOffset = m_columnOffset

End Property

Public Sub AddFirework(ByRef oFirework As clsFirework)

    Me.Lines(oFirework.Line).AddFirework oFirework

End Sub

Public Function Display(ByRef rngCurrentCell As Range) As Range
    
    Dim Index As Integer
    Dim rngTemp As Range
    
    Set rngTemp = rngCurrentCell.OffSet(ColumnOffset:=1)
    rngTemp.Value = "Ïóëüò ¹" & Me.Number
    rngTemp.Font.Bold = True
     
    Set rngTemp = rngCurrentCell.OffSet(RowOffset:=1, ColumnOffset:=0)
    
    
    
    For Index = 1 To UBound(m_clsLines)
            If Me.Lines(Index).Filled <> False Then
                Set rngTemp = Me.Lines(Index).Display(rngCurrentCell:=rngTemp)
            End If
    Next Index
    
    Set Display = rngCurrentCell.OffSet(RowOffset:=RowOffset + m_rowOffset, ColumnOffset:=ColumnOffset + m_columnOffset)
    
End Function

Private Sub Class_Initialize()
    Dim Index As Integer
    For Index = 0 To UBound(m_clsLines)
        Set m_clsLines(Index) = New clsLine
        m_clsLines(Index).Number = Index
    Next Index
End Sub
