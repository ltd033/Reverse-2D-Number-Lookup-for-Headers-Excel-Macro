Attribute VB_Name = "Module1"		
'If copy and pasting then copy below HERE"

Sub FindNearestValue()

    ' Get user input for the value to search for and the range of the matrix
    Dim searchValue As Variant
    searchValue = InputBox("Enter the value to search for:")
    
    Dim matrixRange As Range
    Set matrixRange = Application.InputBox("Select the range of the matrix:", Type:=8)
    
    ' Find the nearest value in the matrix and its position
    Dim nearestValue As Variant
    Dim nearestRowIndex As Long
    Dim nearestColIndex As Long
    
    nearestValue = matrixRange.Cells(1, 1).Value
    nearestRowIndex = 1
    nearestColIndex = 1
    
    For i = 1 To matrixRange.Rows.Count
        For j = 1 To matrixRange.Columns.Count
            If Abs(matrixRange.Cells(i, j).Value - searchValue) < Abs(nearestValue - searchValue) Then
                nearestValue = matrixRange.Cells(i, j).Value
                nearestRowIndex = i
                nearestColIndex = j
            End If
        Next j
    Next i
    
    ' Get the row and column headers based on the position of the matrix
    Dim headerRow As Range
    Set headerRow = matrixRange.Offset(-1, 0).Resize(1, matrixRange.Columns.Count)
    
    Dim headerCol As Range
    Set headerCol = matrixRange.Offset(0, -1).Resize(matrixRange.Rows.Count, 1)
    
    ' Output the nearest value and its position
    Range(matrixRange(1, matrixRange.Columns.Count + 2).Address).Offset(-1, 0) = "Input Value"
    Range(matrixRange(1, matrixRange.Columns.Count + 3).Address).Offset(-1, 0) = searchValue
    Range(matrixRange(2, matrixRange.Columns.Count + 2).Address).Offset(-1, 0) = "Nearest"
    Range(matrixRange(2, matrixRange.Columns.Count + 3).Address).Offset(-1, 0) = nearestValue
    Range(matrixRange(3, matrixRange.Columns.Count + 2).Address).Offset(-1, 0) = "Row Index"
    Range(matrixRange(3, matrixRange.Columns.Count + 3).Address).Offset(-1, 0) = nearestRowIndex
    Range(matrixRange(4, matrixRange.Columns.Count + 2).Address).Offset(-1, 0) = "Column Index"
    Range(matrixRange(4, matrixRange.Columns.Count + 3).Address).Offset(-1, 0) = nearestColIndex
    
    ' Output the row and column headers
    Range(matrixRange(6, matrixRange.Columns.Count + 2).Address).Offset(-1, 0) = "Column Header"
    Range(matrixRange(6, matrixRange.Columns.Count + 3).Address).Offset(-1, 0) = headerRow.Cells(1, nearestColIndex).Value
    Range(matrixRange(5, matrixRange.Columns.Count + 2).Address).Offset(-1, 0) = "Row Header"
    Range(matrixRange(5, matrixRange.Columns.Count + 3).Address).Offset(-1, 0) = headerCol.Cells(nearestRowIndex, 1).Value

End Sub

