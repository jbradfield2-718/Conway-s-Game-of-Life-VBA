Attribute VB_Name = "Module1"
Public pause_flag As Integer
Public num_generations As Integer


Sub Main()
    pause_flag = 0
    num_generations = 0
    Life
End Sub

Sub Life()
    Seed
End Sub

Sub Seed()
Freeze_screen
Workbooks("Game of Life.xlsm").Sheets("Previous Array").Range("a1:az40").Name = "last_array"
Workbooks("Game of Life.xlsm").Sheets("Board").Range("a1:az40").Name = "board"
Refresh
Sheets("Previous Array").Select
Range("last_array").Select
Range("last_array").Formula = "=rand()"

' For loop seeds the array with either ones or zeroes based on the
' output of the rand function.
For Each c In Worksheets("Previous Array").Range("last_array").Cells
    
    If (c.Value) > 0.5 Then
        c.Value = 1
    Else
        c.Value = 0
    End If
    
    Next
Range("last_array").Select
Range("last_array").Copy
Sheets("Board").Select
Range("board").PasteSpecial (xlPasteValues)
Color
Unfreeze_screen
Application.ScreenUpdating = True
PauseEvent (0.25)
Tick
End Sub
Sub Refresh()
    Range("board").Interior.ColorIndex = 0
End Sub
Sub Freeze_screen()
    Application.ScreenUpdating = False
End Sub
Sub Unfreeze_screen()
    Sheets("Board").Select
    Application.ScreenUpdating = True
End Sub
Sub Color()
    Refresh
    For Each c In Worksheets("Board").Range("board").Cells
        If (c.Value) > 0.5 Then c.Interior.ColorIndex = 5
    Next
End Sub

Sub Tick()
    Dim sum As Integer
    Dim row As Integer
    Dim column As Integer
    
    num_generations = num_generations + 1
    Freeze_screen
    For Each c In Worksheets("Previous Array").Range("last_array").Cells
        sum = Calc_Neighbors(c)
        row = c.row
        column = c.column
        
        If c.Value = 1 Then
            If sum < 2 Or sum > 3 Then
                Sheets("Board").Select
                Cells(row, column) = 0
            End If
            
        ElseIf c.Value = 0 And sum = 3 Then
                Sheets("Board").Select
                Cells(row, column) = 1
        End If
    Next
    
    Sheets("Board").Select
    Range("bd13").Value = num_generations
    Range("board").Select
    Range("board").Copy
    Sheets("Previous Array").Select
    Range("last_array").PasteSpecial (xlPasteValues)

    Color
    
    
    Unfreeze_screen
    PauseEvent (0.25)
    
    Tick
    
End Sub
Sub Pause_Button()
    
    If pause_flag = 0 Then
        pause_flag = 1
        Do While pause_flag = 1
            DoEvents
        Loop
        
    Else
        pause_flag = 0
        Exit Sub
    End If
    
End Sub

Function Calc_Neighbors(c)
 Dim row As Integer
 Dim column As Integer
 Dim max_rows As Integer
 Dim max_columns As Integer
 Dim sum As Integer
  
 Sheets("Previous Array").Select
 row = c.row
 column = c.column
 max_rows = Range("last_array").Rows.Count
 max_columns = Range("last_array").Columns.Count

'-------------------------------------------------------------------------------------------------------
' UPPER LEFT HAND CORNER
'-------------------------------------------------------------------------------------------------------
 If row = 1 And column = 1 Then
    sum = Cells(row, column + 1) + Cells(row + 1, column + 1) + Cells(row + 1, column) + _
                    Cells(row, max_columns) + Cells(row + 1, max_columns) + Cells(max_rows, column) + _
                    Cells(max_rows, column + 1) + Cells(max_rows, max_columns)
    Calc_Neighbors = sum
    Exit Function
'-------------------------------------------------------------------------------------------------------
' UPPER RIGHT HAND CORNER
'-------------------------------------------------------------------------------------------------------
 ElseIf row = 1 And column = max_columns Then
    sum = Cells(row, column - 1) + Cells(row + 1, column - 1) + Cells(row + 1, column) + _
                    Cells(1, 1) + Cells(2, 1) + Cells(max_rows, 1) + Cells(max_rows, max_columns) + _
                    Cells(max_rows, max_columns - 1)
    Calc_Neighbors = sum
    Exit Function
'-------------------------------------------------------------------------------------------------------
' LOWER LEFT HAND CORNER
'-------------------------------------------------------------------------------------------------------
 ElseIf row = max_rows And column = 1 Then
    sum = Cells(max_rows - 1, column) + Cells(max_rows - 1, column + 1) + Cells(row, column + 1) + _
                    Cells(1, max_columns) + Cells(1, 2) + Cells(1, 1) + Cells(max_rows, max_columns) + _
                    Cells(max_rows - 1, max_columns)
    Calc_Neighbors = sum
    Exit Function
'-------------------------------------------------------------------------------------------------------
' LOWER RIGHT HAND CORNER
'-------------------------------------------------------------------------------------------------------
 ElseIf row = max_rows And column = max_columns Then
    sum = Cells(max_rows, column - 1) + Cells(max_rows - 1, column - 1) + Cells(max_rows - 1, column) + _
                    Cells(max_rows - 1, 1) + Cells(max_rows, 1) + Cells(1, 1) + Cells(1, max_columns) + _
                    Cells(1, max_columns - 1)
    Calc_Neighbors = sum
    Exit Function
'-------------------------------------------------------------------------------------------------------
' TOP ROW
'-------------------------------------------------------------------------------------------------------
 ElseIf row = 1 Then
    sum = Cells(max_rows, column - 1) + Cells(max_rows, column) + Cells(max_rows, column + 1) + _
                    Cells(1, column + 1) + Cells(2, column + 1) + Cells(2, column) + Cells(2, column - 1) + _
                    Cells(1, column - 1)
    Calc_Neighbors = sum
    Exit Function
'-------------------------------------------------------------------------------------------------------
' BOTTOM ROW
'-------------------------------------------------------------------------------------------------------
 ElseIf row = max_rows Then
    sum = Cells(max_rows - 1, column - 1) + Cells(max_rows - 1, column) + Cells(max_rows - 1, column + 1) + _
                    Cells(max_rows, column + 1) + Cells(1, column + 1) + Cells(1, column) + Cells(1, column - 1) + _
                    Cells(max_rows, column - 1)
    Calc_Neighbors = sum
    Exit Function
'-------------------------------------------------------------------------------------------------------
' RIGHT EDGE
'-------------------------------------------------------------------------------------------------------
 ElseIf column = max_columns Then
    sum = Cells(row - 1, column - 1) + Cells(row - 1, column) + Cells(row - 1, 1) + _
                    Cells(row, 1) + Cells(row + 1, 1) + Cells(row + 1, column) + Cells(row + 1, column - 1) + _
                    Cells(row, column - 1)
    Calc_Neighbors = sum
    Exit Function
'-------------------------------------------------------------------------------------------------------
' LEFT EDGE
'-------------------------------------------------------------------------------------------------------
 ElseIf column = 1 Then
    sum = Cells(row - 1, column) + Cells(row - 1, column + 1) + Cells(row, column + 1) + _
                    Cells(row + 1, column + 1) + Cells(row + 1, column) + Cells(row + 1, max_columns) + Cells(row, max_columns) + _
                    Cells(row - 1, max_columns)
    Calc_Neighbors = sum
    Exit Function
'-------------------------------------------------------------------------------------------------------
' BASE CASE, CENTER OF BOARD
'-------------------------------------------------------------------------------------------------------
 Else
    sum = Cells(row - 1, column) + Cells(row - 1, column + 1) + Cells(row, column + 1) + _
                    Cells(row + 1, column + 1) + Cells(row + 1, column) + Cells(row + 1, column - 1) + Cells(row, column - 1) + _
                    Cells(row - 1, column - 1)
    Calc_Neighbors = sum
    Exit Function
 End If
 
End Function
Function PauseEvent(ByVal Delay As Double)
  Dim dblEndTime As Double
  dblEndTime = Timer + Delay
  Do While Timer < dblEndTime
    DoEvents
  Loop
End Function
