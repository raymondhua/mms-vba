VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub 'detects when cell is deleted
    Dim InputValue As String
    Dim EmptyOrNot As Boolean: EmptyOrNot = True
    Dim NumberInput As Boolean
    Dim YearEnd As Boolean
    
    On Error GoTo ErrorHandler
    If CorrectYear(Target.Value2) Then
    InputValue = Target.Value2
    Else
    InputValue = Target.Value
    End If
    If Not IsEmpty(Target.Value) Then EmptyOrNot = False
    ChangeValue Target, InputValue, EmptyOrNot
    Exit Sub
ErrorHandler:
    InputValue = Target.Value2
    If Not IsEmpty(Target.Value2) Then EmptyOrNot = False
    Err.Clear
    ChangeValue Target, InputValue, EmptyOrNot
End Sub
Function NumbersOnly(DateValue As String) As String
    NumbersOnly = DateValue
    If IsNumeric(DateValue) Then
        If Len(NumbersOnly) = 6 Then
            NumbersOnly = Left(DateValue, 2) & "/" & Mid(DateValue, 3, 2) & "/20" & Right(DateValue, 2)
        ElseIf Len(NumbersOnly) = 5 Then
            MsgBox "Date format invalid for formatting, please enter the correct format (DD/MM/YYYY)"
            'NumbersOnly = Left(DateValue, 1) & "/" & Mid(DateValue, 2, 2) & "/20" & Right(DateValue, 2)
            Exit Function
        ElseIf Len(NumbersOnly) = 7 Then
            'NumbersOnly = Left(DateValue, 1) & "/" & Mid(DateValue, 3, 1) & "/" & Right(DateValue, 4)
            MsgBox "Date format invalid for formatting, please enter the correct format (DD/MM/YYYY)"
            Exit Function
        ElseIf Len(NumbersOnly) = 8 Then
            NumbersOnly = Left(DateValue, 2) & "/" & Mid(DateValue, 3, 2) & "/" & Right(DateValue, 4)
        ElseIf Len(NumbersOnly) = 4 Then
            NumbersOnly = Left(DateValue, 1) & "/" & Mid(DateValue, 2, 1) & "/20" & Right(DateValue, 2)
        End If
    End If
End Function
Sub ChangeValue(ByVal Target As Range, InputValue As String, EmptyOrNot As Boolean)
    If Not EmptyOrNot And (Target.NumberFormat = "dd/mm/yyyy" Or Target.NumberFormat = "dd/mm/yyyy;@") Then
        Application.EnableEvents = False
        If CorrectYear(InputValue) Then
        Target.Value = CheckDate(InputValue)
        Else
        MsgBox "Please check if year is right with the format (dd/mm/YYYY)"
        Target.ClearContents
        End If
        Application.EnableEvents = True
    End If
End Sub
Function CheckDate(DateValue As String) As String

    If Not IsNumeric(DateValue) Then
    CheckDate = DateValue
    Else
    CheckDate = NumbersOnly(DateValue)
    End If
    
    If Not ValidDate(CheckDate) Then CheckDate = FormatDate(CheckDate)
    End If
    
    On Error GoTo ErrorHandler
    CheckDate = Format(CheckDate, "MM/DD/YYYY")
    
    Exit Function
ErrorHandler:
        MsgBox "Date format invalid for formatting, please enter the correct format (DD/MM/YYYY)"
End Function
' Checks if date is a valid format (DD/MM/YYYY)
Function ValidDate(DateValue As String) As Boolean
    Dim ArrInput As Variant: ArrInput = Split(DateValue, "/")
    
    If UBound(ArrInput) = 2 Then 'make sure there are exactly two slashes in the date
        Dim MyValidDate As Date
        MyValidDate = DateSerial(ArrInput(2), ArrInput(1), ArrInput(0))
    
        If Day(MyValidDate) = CLng(ArrInput(0)) And _
           Month(MyValidDate) = CLng(ArrInput(1)) And _
           Year(MyValidDate) = CLng(ArrInput(2)) Then
            ValidDate = True
            Exit Function
        End If
    End If
End Function
' Formats the date by removing special characters and replace them with slashes
Function FormatDate(DateValue As String) As String
    Dim xChars As String: xChars = "*[\!%^:~#|@.;'`\*'$,""]*"
    If DateValue Like xChars Then
        Dim I As Long
        For I = 1 To Len(xChars)
            DateValue = Replace$(DateValue, Mid$(xChars, I, 1), "/")
        Next
    End If
    FormatDate = DateValue
End Function
' This function is to check if the year is correct (2022-2024)
' NOTE: Change this function at the end of 2024 if the spreadsheet is still in use by then
Function CorrectYear(DateValue As String) As Boolean
    If Right(DateValue, 2) >= 22 And Right(DateValue, 2) <= 24 Then 'modify this after 2024 to an approriate column (e.g. 25)
    CorrectYear = True
    ElseIf Right(DateValue, 4) >= 2022 And Right(DateValue, 4) <= 2024 Then 'modify this after 2024 to an approriate column (e.g. 2025)
    CorrectYear = True
    End If
End Function



