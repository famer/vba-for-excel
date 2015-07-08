VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFirework"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_blFilled As Boolean
Private m_sLaunchTime As String
Private m_btPult As Byte
Private m_btLine As Byte
Private m_btPosition As Byte
Private m_btChannel As Byte
Private m_btStation As Byte
Private m_btBarrel As Byte
Private m_sTitle As String
Private m_iX As Integer
Private m_iY As Integer
Private m_iZ As Integer
Private m_iPitch As Integer
Private m_iTurn As Integer

Private m_snDiameter As Single
Private Const LineLength As Byte = 16
Private Const m_btTitleLength As Byte = 8


Private Const m_rowOffset As Byte = 0
Private Const m_columnOffset As Byte = 1

Private m_blBoldTitle As Boolean
Private m_blDuplicate As Boolean
Private m_btSheafCounter As Byte
Private m_clsLine As clsLine

Public Property Get ContainingLine() As clsLine
    
    Set ContainingLine = m_clsLine

End Property


Public Property Set ContainingLine(ByRef oLine As clsLine)
    
    Set m_clsLine = oLine

End Property


Public Property Get BoldTitle() As Boolean

    BoldTitle = m_blBoldTitle

End Property

Public Property Let BoldTitle(ByVal blBoldTitle As Boolean)
    
    m_blBoldTitle = blBoldTitle

End Property

Public Property Get Duplicate() As Boolean

    Duplicate = m_blDuplicate

End Property

Public Property Let Duplicate(ByVal blDuplicate As Boolean)
    
    m_blDuplicate = blDuplicate

End Property

Public Property Get Filled() As Boolean

    Filled = m_blFilled

End Property

Public Property Let Filled(ByVal blFilled As Boolean)
    
    m_blFilled = blFilled

End Property


Public Property Get LaunchTime() As String

    LaunchTime = m_sLaunchTime

End Property

Public Property Let LaunchTime(ByVal sLaunchTime As String)
    
    If sLaunchTime = "" Then Err.Raise 1001, , "Can't be empty!"
    m_sLaunchTime = sLaunchTime

End Property

Public Property Get Pult() As Byte

    Pult = m_btPult

End Property

Public Property Let Pult(ByVal btPult As Byte)
    
    If btPult = 0 Then Me.Filled = False 'Err.Raise 1001, , "Can't be empty!"
    m_btPult = btPult

End Property

Public Property Get Channel() As Byte

    Channel = m_btChannel

End Property

Public Property Let Channel(ByVal btChannel As Byte)
    
    m_btChannel = btChannel
    Me.Line = LineByChannel(btChannel)
    Me.Position = PositionByChannel(btChannel)
    
End Property


Public Property Get Station() As Byte

    Station = m_btStation

End Property

Public Property Let Station(ByVal btStation As Byte)
    
    m_btStation = btStation

End Property

Public Property Get Barrel() As Byte

    Barrel = m_btBarrel

End Property

Public Property Let Barrel(ByVal btBarrel As Byte)
    
    m_btBarrel = btBarrel

End Property

Public Property Get Title() As String

    Title = Left(m_sTitle, m_btTitleLength)

End Property

Public Property Let Title(ByVal sTitle As String)
    
    
    Me.Diameter = DiameterByTitle(sTitle)
    
    If Me.Diameter Then
        m_sTitle = Replace(sTitle, Me.Diameter, "", 1, 1)
        
        'Dont like that
        If Left(m_sTitle, 2) = "' " Then _
            m_sTitle = Mid(m_sTitle, 3)
    Else
        m_sTitle = sTitle
    End If
    
    If Me.Title = "Дубль+" Then Me.BoldTitle = True
    
End Property


Public Property Get X() As Integer

    X = m_iX

End Property

Public Property Let X(ByVal iX As Integer)
    
    m_iX = iX

End Property


Public Property Get Y() As Integer

    Y = m_iY

End Property

Public Property Let Y(ByVal iY As Integer)
    
    m_iY = iY

End Property


Public Property Get Z() As Integer

    Z = m_iZ

End Property

Public Property Let Z(ByVal iZ As Integer)
    
    m_iZ = iZ

End Property


Public Property Get Pitch() As Integer

    Pitch = m_iPitch

End Property

Public Property Let Pitch(ByVal iPitch As Integer)
    
    m_iPitch = iPitch

End Property

Public Property Get Turn() As Integer

    Turn = m_iTurn

End Property

Public Property Let Turn(ByVal iTurn As Integer)
    
    m_iTurn = iTurn

End Property
Public Property Get Position() As Byte

    Position = m_btPosition

End Property

Public Property Let Position(ByVal btPosition As Byte)
    
    m_btPosition = btPosition

End Property

Public Property Get Line() As Byte

    Line = m_btLine

End Property

Public Property Let Line(ByVal btLine As Byte)
    
    m_btLine = btLine

End Property

Public Property Get Diameter() As Single

    Diameter = m_snDiameter

End Property

Public Property Let Diameter(ByVal snDiameter As Single)
    
    m_snDiameter = snDiameter

End Property

Private Function PositionByChannel(ByVal btChannel As Byte) As Byte
    If btChannel <> 0 Then
        PositionByChannel = (btChannel Mod LineLength)
        If PositionByChannel = 0 Then PositionByChannel = 16
    Else
        PositionByChannel = 1
    End If

    
End Function

Private Function LineByChannel(ByVal btChannel As Byte) As Byte

    Dim Ceil, Remainder As Byte
    Ceil = btChannel \ LineLength
    Remainder = btChannel Mod LineLength
    If Remainder > 0 Then
        LineByChannel = Ceil + 1
    Else
        LineByChannel = Ceil
    End If
    
End Function

Private Function DiameterByTitle(ByVal sTitle As String) As Single
    
    '/^[0-9,]+'/
    Dim StartPos, EndPos As Byte
    Dim stTmpStr As String
    Dim blHasDiameter, blHasApostrophe, blHasComma As Boolean
    
    stTmpStr = Left(sTitle, 1)
    blHasDiameter = stTmpStr Like "#"
    
    If blHasDiameter Then
        EndPos = InStr(1, sTitle, " ") - 1
        
        blHasApostrophe = InStr(1, Left(sTitle, EndPos), "'")
        If blHasApostrophe Then _
            EndPos = EndPos - 1
            
        blHasComma = InStr(1, Left(sTitle, EndPos), ",")
        
        If EndPos > 3 Then
            If blHasComma = True Then
                EndPos = 4
            Else
                EndPos = 3
            End If
        End If
        
        DiameterByTitle = Left(sTitle, EndPos)
        
    End If
    
    
End Function

Public Sub Load(ByRef rngCurrentCell As Range)
      
    With Me
            .Filled = True
            .LaunchTime = rngCurrentCell.OffSet(ColumnOffset:=0).Value
            .Pult = rngCurrentCell.OffSet(ColumnOffset:=1).Value
            .Channel = rngCurrentCell.OffSet(ColumnOffset:=2).Value
            .Station = rngCurrentCell.OffSet(ColumnOffset:=3).Value
            .Barrel = rngCurrentCell.OffSet(ColumnOffset:=4).Value
            .Title = rngCurrentCell.OffSet(ColumnOffset:=5).Value
            .X = rngCurrentCell.OffSet(ColumnOffset:=6).Value
            .Y = rngCurrentCell.OffSet(ColumnOffset:=7).Value
            .Z = rngCurrentCell.OffSet(ColumnOffset:=8).Value
            .Pitch = rngCurrentCell.OffSet(ColumnOffset:=9).Value
            .Turn = rngCurrentCell.OffSet(ColumnOffset:=10).Value
        End With
End Sub

Public Function Display(ByRef rngCurrentCell As Range, ByRef oPrevFirework As clsFirework) As Range
    
    rngCurrentCell.Value = Me.Channel
    rngCurrentCell.Font.Bold = True
    rngCurrentCell.HorizontalAlignment = xlCenter

    'If Me.Title <> "" And oPrevFirework.Title <> "" And oPrevFirework.Diameter = Me.Diameter And oPrevFirework.Title = Me.Title Then
    '    Misc.MergeCells rngCurrentCell:=rngCurrentCell.OffSet(RowOffset:=1, ColumnOffset:=-1).Resize(columnSize:=2)

    'Else
        If Me.Diameter <> 0 Then
            If Me.ContainingLine.Sheaf = "" Then
                rngCurrentCell.OffSet(RowOffset:=1).Value = "'" & Me.Diameter
            End If
            rngCurrentCell.OffSet(RowOffset:=1).Font.Bold = True
        End If
        If Me.ContainingLine.Sheaf <> "" And Me.Title <> "" Then
            rngCurrentCell.OffSet(RowOffset:=1).Value = Me.ContainingLine.Sheaf
        End If
        'If Me.Title <> "" Then
        '    Misc.MergeCells rngCurrentCell:=rngCurrentCell.OffSet(RowOffset:=1)
        
        'End If
    'End If

    'Debug.Print Misc.SheafCounter(Misc.SheafCounter + 1)
    If Me.ContainingLine.Sheaf <> "" And Me.Title <> "" And Me.Title <> "XXX" Then
        rngCurrentCell.OffSet(RowOffset:=2).Value = Me.Diameter & "" & Replace(Me.Title, Me.ContainingLine.Sheaf & " ", "", 1, 1)
    Else
        rngCurrentCell.OffSet(RowOffset:=2).Value = Me.Title
    End If
    
    If Me.BoldTitle = True Then _
        rngCurrentCell.OffSet(RowOffset:=2).Font.Bold = True
        
     If Me.Duplicate = True Then _
        rngCurrentCell.OffSet(RowOffset:=0).Interior.Color = RGB(255, 0, 0)
    
    
    
    '.Interior.Color = RGB(255, 0, 0)
    
    
    Set Display = rngCurrentCell.OffSet(RowOffset:=m_rowOffset, ColumnOffset:=m_columnOffset)

End Function

