VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_btNumber As Byte
Private m_blFilled As Boolean
Private Const m_btLength As Byte = 16
Private m_clsFireworks(1 To 16) As clsFirework
Private Const m_rowOffset As Byte = 3
Private Const m_columnOffset As Byte = 0

Private m_sSheaf As String
Private m_btSheafCount As Byte


Public Property Get Sheaf() As String

    Sheaf = m_sSheaf

End Property

Public Property Let Sheaf(ByVal sSheaf As String)
    
    m_sSheaf = sSheaf

End Property

Public Property Get SheafCount() As Byte

    If Me.Sheaf <> "" Then _
        SheafCount = Right(Me.Sheaf, 1)

End Property


Public Property Get Number() As Byte

    Number = m_btNumber

End Property

Public Property Let Number(ByVal btNumber As Byte)
    
    m_btNumber = btNumber
    
    If Me.Number <> 0 Then
        Dim Index As Integer
        For Index = 1 To UBound(m_clsFireworks)
            m_clsFireworks(Index).Channel = (Me.Number - 1) * m_btLength + Index
        Next Index
    End If

End Property

Public Property Get Filled() As Boolean

    Filled = m_blFilled

End Property

Public Property Let Filled(ByVal blFilled As Boolean)
    
    m_blFilled = blFilled

End Property


Property Get Fireworks(Index As Integer) As clsFirework
    
    Set Fireworks = m_clsFireworks(Index)
    
End Property

Property Let Fireworks(Index As Integer, ByVal Firework As clsFirework)
    
    If Index < LBound(m_clsFireworks) Or Index > UBound(m_clsFireworks) Then Err.Raise 9
    Set m_clsFireworks(Index) = Firework
    Me.Filled = True
    
End Property

Public Property Get RowOffset() As Byte
        
    RowOffset = m_rowOffset

End Property

Public Property Get ColumnOffset() As Byte
        
    ColumnOffset = m_columnOffset

End Property

Public Sub AddFirework(ByRef oFirework As clsFirework)
        
    Set oFirework.ContainingLine = Me
    If Me.Fireworks(oFirework.Position).Title <> "" Then
        oFirework.Duplicate = True
    End If
        
    Me.Fireworks(oFirework.Position) = oFirework
    
    If oFirework.Position = 1 Then
        If oFirework.Title Like "*B# *" Then
            Dim btSheafPos As Byte
            btSheafPos = InStr(oFirework.Title, "B")
            Me.Sheaf = Mid(oFirework.Title, btSheafPos, 2)
        End If
    End If
    
End Sub

Public Function Display(ByRef rngCurrentCell As Range) As Range
    
    Dim Index As Integer
    Dim rngTemp As Range
    Dim rngOrig As Range
    Dim Diameter As Single
    Dim Position As Byte
    Dim oPrevFirework As New clsFirework
    Dim oCurentFirework As New clsFirework
    Position = 1
    Diameter = Me.Fireworks(1).Diameter
    
    Set rngTemp = rngCurrentCell.OffSet(RowOffset:=0, ColumnOffset:=0)
    rngTemp.Value = "¹" & Me.Number
    rngTemp.Font.Bold = True
    
    Set rngTemp = rngCurrentCell.OffSet(RowOffset:=0, ColumnOffset:=1)
    Set rngOrig = rngTemp
    For Index = 1 To UBound(m_clsFireworks)
            
            'If Me.Fireworks(Index).Diameter <> Diameter Then
             '       Set rngOrig = rngTemp.Offset(RowOffset:=1, ColumnOffset:=-(Me.Fireworks(Index).Position - Position))
             ' Application.DisplayAlerts = False
              '      With rngOrig.Resize(columnSize:=(Me.Fireworks(Index).Position - Position))
               '             .Borders.LineStyle = xlThick
                '            .HorizontalAlignment = xlCenter
                 '           .ReadingOrder = xlContext
                  '          .MergeCells = True
                            '.Merge
                   ' End With
               ' Position = Me.Fireworks(Index).Position
            'End If
            
            Set oCurentFirework = Me.Fireworks(Index)
            'oCurentFirework.Diameter = "'" & oCurentFirework.Diameter
            
            If Me.Sheaf = "" And oCurentFirework.Title <> "" And oPrevFirework.Title <> "" And oPrevFirework.Diameter = oCurentFirework.Diameter Then
                Misc.MergeCells rngCurrentCell:=rngTemp.OffSet(RowOffset:=1, ColumnOffset:=-1).Resize(columnSize:=2)
            End If
            If Me.Sheaf <> "" Then
                If (m_btLength - (m_btLength Mod Me.SheafCount)) >= oCurentFirework.Position Then
                    If oCurentFirework.Title = "" Then
                        oCurentFirework.Title = "XXX"
                    Else
                        oCurentFirework.Title = CStr(oCurentFirework.Diameter) & " " & oCurentFirework.Title
                    End If
                    If ((oCurentFirework.Position - 1) Mod (Me.SheafCount)) <> 0 Then _
                        Misc.MergeCells rngCurrentCell:=rngTemp.OffSet(RowOffset:=1, ColumnOffset:=-1).Resize(columnSize:=2)
                End If
            End If
            If oCurentFirework.Title <> "" Then
                Misc.MergeCells rngCurrentCell:=rngTemp.OffSet(RowOffset:=1)
        
            End If
            'If oCurentFirework.Title <> "" And oPrevFirework.Title <> "" And oPrevFirework.Diameter = oCurentFirework.Diameter And oPrevFirework.Title = oCurentFirework.Title Then
            Set rngTemp = Me.Fireworks(Index).Display(rngCurrentCell:=rngTemp, oPrevFirework:=oPrevFirework)
            
           
            
            Set oPrevFirework = Me.Fireworks(Index)
    Next Index
    
    Set Display = rngCurrentCell.OffSet(RowOffset:=m_rowOffset, ColumnOffset:=m_columnOffset)
    
End Function

Private Sub Class_Initialize()
    
    Dim Index As Integer
    For Index = 1 To UBound(m_clsFireworks)
        Set m_clsFireworks(Index) = New clsFirework
        m_clsFireworks(Index).Position = Index
        Set m_clsFireworks(Index).ContainingLine = Me
    Next Index
End Sub
