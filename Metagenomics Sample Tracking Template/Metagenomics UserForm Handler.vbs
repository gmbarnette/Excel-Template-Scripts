
If OptionButton1 = True Then

    Sheets("Enter Sample Names Here").Visible = xlSheetVeryHidden
    Sheets("Enter Sample Names Here Plate").Visible = xlSheetVisible
    Sheets("Enter Primer Sets Here").Visible = xlSheetVisible
    
    Sheets("Plate Map 1").Visible = xlSheetVisible
    Sheets("Plate Map 2").Visible = xlSheetVisible
    Sheets("Plate Map 3").Visible = xlSheetVisible
    Sheets("Sample Sheet Converter").Visible = xlSheetVisible
    
    Sheets("Enter Sample Names Here Plate").Select
    
    Unload Me
    
ElseIf OptionButton2 = True Then
    
    Sheets("Enter Sample Names Here").Visible = xlSheetVisible
    Sheets("Enter Sample Names Here Plate").Visible = xlSheetVeryHidden
    Sheets("Enter Primer Sets Here").Visible = xlSheetVeryHidden
    
    Sheets("Plate Map 1").Visible = xlSheetVisible
    Sheets("Plate Map 2").Visible = xlSheetVisible
    Sheets("Plate Map 3").Visible = xlSheetVisible
    Sheets("Sample Sheet Converter").Visible = xlSheetVisible
    
    Sheets("Enter Sample Names Here").Select

    Unload Me
    
End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub
