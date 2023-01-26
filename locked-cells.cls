VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ' Locking certain cells in spreadsheet when opening
    ' Make sure that all cells in Sheet2 are unlocked first
    ' Alignment > bottom right corner > Protection tab > Uncheck "Locked"
    
    ' Method where columns H of Sheet2 are locked when opening the spreadsheet
    Worksheets("Sheet2").Activate
    
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False
    LockColumn ("H")
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("A1").Select
    
    ' Method where columns A and H of Sheet1 are locked when opening the spreadsheet
    Worksheets("Sheet1").Activate
    
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False
    ' If you want to lock more columns, copy the LockColumn method and insert cell letter between the quotation marks
    LockColumn ("H")
    LockColumn ("A")
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("A1").Select
    
End Sub

Private Sub LockColumn(Column As String)
    Columns(Column + ":" + Column).EntireColumn.Locked = True
End Sub

