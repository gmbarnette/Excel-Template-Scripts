
Private Sub Worksheet_Change(ByVal Target As Range)
If Target.Column = 1 And Target.Row = 2 Then
    On Error Resume Next
    If Target.Value = "" Then
        Target.Value = "Enter Plate Name Here!"
    End If
    
End If

If Target.Column = 5 And Target.Row = 2 Then
    On Error Resume Next
    If Target.Value = "" Then
        Target.Value = "Enter Plate Name Here!"
    End If
    
End If

If Target.Column = 9 And Target.Row = 2 Then
    On Error Resume Next
    If Target.Value = "" Then
        Target.Value = "Enter Plate Name Here!"
    End If
    
End If

If Target.Column = 13 And Target.Row = 2 Then
    On Error Resume Next
    If Target.Value = "" Then
        Target.Value = "Enter Plate Name Here!"
    End If
    
End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
