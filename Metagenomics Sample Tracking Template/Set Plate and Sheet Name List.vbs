
Private Sub Worksheet_Change(ByVal Target As Range)
If Target.Column = 1 And Target.Row = 2 Then
    On Error Resume Next
    If Target.Value = "" Then
        Target.Value = "Enter Plate Name Here!"
    End If
    If Target.Value = "Enter Plate Name Here!" Then
        Sheet2.Name = "Plate Map 1"
    Else
        Sheet2.Name = Target.Value
    End If
End If

If Target.Column = 7 And Target.Row = 2 Then
    On Error Resume Next
    If Target.Value = "" Then
        Target.Value = "Enter Plate Name Here!"
    End If
    If Target.Value = "Enter Plate Name Here!" Then
        Sheet9.Name = "Plate Map 2"
    Else
        Sheet9.Name = Target.Value
    End If
End If

If Target.Column = 13 And Target.Row = 2 Then
    On Error Resume Next
    If Target.Value = "" Then
        Target.Value = "Enter Plate Name Here!"
    End If
    If Target.Value = "Enter Plate Name Here!" Then
        Sheet8.Name = "Plate Map 3"
    Else
        Sheet8.Name = Target.Value
    End If
End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
