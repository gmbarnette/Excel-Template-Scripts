Private Sub Workbook_Open()
Dim numOfSamples
Dim numOfSamples2
Dim totSamples
Dim todayDate
Dim workSheetToCount As Worksheet
Dim workSheetToCount2 As Worksheet
Set workSheetToCount = Sheets("Enter Sample Names Here")
Set workSheetToCount2 = Sheets("Enter Primer Sets Here")

todayDate = Format(Date, "dd-mmm-yy")
numOfSamples = Application.CountA(workSheetToCount.Range("B4:D99")) + Application.CountA(workSheetToCount.Range("H4:J99")) + Application.CountA(workSheetToCount.Range("N4:P99"))
numOfSamples2 = Application.CountIf(workSheetToCount2.Range("B4:D99"), "*?") + Application.CountIf(workSheetToCount2.Range("H4:J99"), "*?") + Application.CountIf(workSheetToCount2.Range("N4:P99"), "*?")
totSamples = numOfSamples + numOfSamples2


If numOfSamples > 0 And numOfSamples2 > 0 Then
    MsgBox (numOfSamples)
    MsgBox (numOfSamples2)
    MsgBox ("Error:  Data Entered in both List and Format Form")
    Exit Sub
End If
If totSamples > 0 Then
    If numOfSamples > 0 Then
        Sheet1.Visible = xlSheetVisible
        Sheet11.Visible = xlSheetHidden
        Sheet10.Visible = xlSheetHidden
        Exit Sub
    Else
        Sheet1.Visible = xlSheetHidden
        Sheet11.Visible = xlSheetVisible
        Sheet10.Visible = xlSheetVisible
        Exit Sub
    End If

Else
    Sheet1.Visible = xlSheetVisible 'Enter Sample Names Here
    Sheet11.Visible = xlSheetHidden 'Enter Sample Names Here Plate
    Sheet10.Visible = xlSheetHidden 'Enter Primer Sets
    Sheet2.Visible = xlSheetVisible 'Plate A
    Sheet9.Visible = xlSheetVisible 'Plate B
    Sheet8.Visible = xlSheetVisible 'Plate C
    Sheet4.Visible = xlSheetVisible 'Sample Sheet Converter
    Sheet6.Visible = xlSheetVeryHidden  'Template Math Page
    Sheet3.Visible = xlSheetHidden  'Barcode Sequences
    Sheet5.Visible = xlSheetHidden  'Barcode Plates
    
    Sheet2.Range("M1") = todayDate  'Plate A Date
    Sheet9.Range("M1") = todayDate  'Plate B Date
    Sheet8.Range("M1") = todayDate  'Plate C Date
    
End If
    Sheet1.Select

UserForm1.Show


End Sub
