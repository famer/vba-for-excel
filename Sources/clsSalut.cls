VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sTitle As String
Private m_collPults As New Collection
Private m_clsPults() As New clsPult


Public Property Get Title() As String

    Title = m_sTitle

End Property

Public Property Let Title(ByVal sTitle As String)
    
    m_sTitle = sTitle
    
End Property

Property Get Pults(Index As String) As clsPult
    Set Pults = m_collPults.Item(Index)
End Property

Property Let Pults(Index As String, ByVal Pult As clsPult)
    ReDim m_clsPults(CByte(Index))
    'm_clsPults(CByte(Index)) = Pult
    If Misc.InCollection(m_collPults, Index) Then
        m_collPults.Remove Index
    End If
    m_collPults.Add Item:=Pult, key:=(Index)
    
End Property

Public Sub AddFirework(ByVal oFirework As clsFirework)
            
    If oFirework.Filled = True Then
        If Not Misc.InCollection(m_collPults, oFirework.Pult) Then
            Dim Pult As New clsPult
            Pult.Number = oFirework.Pult
            'Me.Pults(oFirework.Pult) = Pult
            m_collPults.Add Item:=Pult, key:=CStr(oFirework.Pult)
        End If
        
        Me.Pults(oFirework.Pult).AddFirework oFirework
    End If
End Sub

Public Function Display(ByRef rngCurrentCell As Range)
    
    Misc.Prepare
    
    Dim oPult As clsPult
    Dim rngTemp As Range
    
    Set rngTemp = rngCurrentCell.OffSet(ColumnOffset:=1)
    rngTemp.Value = Me.Title
    rngTemp.Font.Bold = True
     
    Set rngTemp = rngCurrentCell.OffSet(RowOffset:=2)
    
    For Each oPult In m_collPults
            Set rngTemp = oPult.Display(rngCurrentCell:=rngTemp)
    Next oPult
    
    Misc.Ended
    
End Function



Private Sub Class_Initialize()
    Dim Index As Integer
    
End Sub
