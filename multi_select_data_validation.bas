Attribute VB_Name = "multi_select_data_validation"
Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents = False
    
    Dim KeyCells As Range
    Set KeyCells = Range("H3:H1000")
    
'Check if it happened in range of interest
If Not Application.Intersect(KeyCells, Range(Target.Address)) _
    Is Nothing Then
    
    'Call sub to update field

    Call update_component_category(Target.Address)

    End If
    
Application.EnableEvents = True
End Sub


Sub update_component_category(cell_address As String)

Dim old_value
Dim new_value
Dim output_val


new_value = Range(cell_address).Value
Application.Undo
old_value = Range(cell_address).Value
'Application.Repeat

'case clearing cell
If new_value = Empty Then
    Range(cell_address).ClearContents
    Exit Sub
'was blank
ElseIf old_value = Empty Then
    output_val = new_value
ElseIf InStr(old_value, new_value) = 0 Then
    'add to end
    output_val = old_value & ", " & new_value
Else
    'case 1: only value
    If Len(new_value) = Len(old_value) Then
        Range(cell_address).ClearContents
    'case 2: remove from start
    ElseIf Left(old_value, Len(new_value)) = new_value Then
        output_val = Replace(old_value, new_value & ", ", "")
    'case 3: remove from middle or end
    Else
        output_val = Replace(old_value, ", " & new_value, "")
    End If
    
End If

Range(cell_address).Value = output_val



End Sub

