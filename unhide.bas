Attribute VB_Name = "unhide"
Sub unhide_above()

Dim actv_row As Integer
Dim actv_col As Integer

actv_row = ActiveCell.Row
actv_col = ActiveCell.Column


If actv_row > 1 Then
    ActiveSheet.rows(actv_row - 1).EntireRow.Hidden = False
    Cells(actv_row - 1, actv_col).Select
End If


End Sub


Sub unhide_below()

Dim actv_row As Integer
Dim actv_col As Integer

actv_row = ActiveCell.Row
actv_col = ActiveCell.Column

If actv_row < 1048576 Then
    ActiveSheet.rows(actv_row + 1).EntireRow.Hidden = False
    Cells(actv_row + 1, actv_col).Select
End If


End Sub


Sub unhide_left()

Dim actv_row As Integer
Dim actv_col As Integer

actv_row = ActiveCell.Row
actv_col = ActiveCell.Column

If actv_col > 1 Then
    ActiveSheet.Columns(actv_col - 1).EntireColumn.Hidden = False
    Cells(actv_row, actv_col - 1).Select
End If



End Sub


Sub unhide_right()

Dim actv_row As Integer
Dim actv_col As Integer

actv_row = ActiveCell.Row
actv_col = ActiveCell.Column

If actv_col < 16384 Then
    ActiveSheet.Columns(actv_col + 1).EntireColumn.Hidden = False
    Cells(actv_row, actv_col + 1).Select
End If


End Sub
