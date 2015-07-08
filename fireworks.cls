VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ёта нига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Workbook_Open()

        
    Functional.Import
    Workbooks(1).Worksheets(1).Activate
    Functional.LoadFireworks
    'Workbooks(2).Close False

End Sub





