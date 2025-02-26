Attribute VB_Name = "scroll"
Sub assign_scroll_keys()

Application.OnKey "^%{LEFT}", "keys_scroll_left"
Application.OnKey "^%{RIGHT}", "keys_scroll_right"
Application.OnKey "^%{UP}", "keys_scroll_up"
Application.OnKey "^%{DOWN}", "keys_scroll_down"

End Sub

Sub keys_scroll_left()

Dim actv_col As Long
actv_col = ActiveWindow.ScrollColumn

If actv_col > 1 Then
    ActiveWindow.ScrollColumn = actv_col - 1
    End If

End Sub



Sub keys_scroll_right()

Dim actv_col As Long
actv_col = ActiveWindow.ScrollColumn

If actv_col < 16384 Then
    ActiveWindow.ScrollColumn = actv_col + 1
    End If

End Sub


Sub keys_scroll_up()

Dim actv_row As Long
actv_row = ActiveWindow.ScrollRow

If actv_row > 1 Then
    ActiveWindow.ScrollRow = actv_row - 1
    End If

End Sub


Sub keys_scroll_down()

Dim actv_row As Long
actv_row = ActiveWindow.ScrollRow

If actv_row < 1048576 Then
    ActiveWindow.ScrollRow = actv_row + 1
    End If

End Sub
