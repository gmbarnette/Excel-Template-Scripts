Private Sub Worksheet_Activate()
Set plateAPrimerRange = Sheet10.Range("P4:P99")
Set listAPrimerRange = Sheet1.Range("P4:P99")
Set plateASampleRange = Sheet10.Range("N4:N99")
Set listASampleRange = Sheet1.Range("N4:N99")

Dim primerArray(3) As String
Dim numOfPrimerPlate
Dim numOfPrimerList
Dim numOfSamplePlate
Dim numOfSampleList
Dim primerCounter
Dim positionCounter
Dim arrayCounter
Dim arrayPosition
Dim arrayBoolean

numOfSamplePlate = Application.CountIf(plateASampleRange, "*?") + Application.Count(plateASampleRange)
numOfSampleList = Application.CountIf(listASampleRange, "*?") + Application.Count(listASampleRange)
numOfPrimerList = Application.CountIf(listAPrimerRange, "*?") + Application.Count(listAPrimerRange)
numOfPrimerPlate = Application.CountIf(plateAPrimerRange, "*?") + Application.Count(plateAPrimerRange)
primerCounter = 1
positionCounter = 4
arrayCounter = 0
arrayPosition = 0
arrayBoolean = 0


If numOfSamplePlate > numOfPrimerPlate Or numOfSampleList > numOfPrimerList Then
    MsgBox ("Please Enter a Primer Set for Every Sample")
    If numOfSamplePlate + numOfPrimerPlate > 0 Then
        Sheet10.Select
    ElseIf numOfSampleList + numOfPrimerList > 0 Then
        Sheet1.Select
    End If
    Exit Sub
ElseIf numOfSamplePlate < numOfPrimerPlate Or numOfSampleList < numOfPrimerList Then
    MsgBox ("More Primer Sets Selected Than Samples Entered.  Please delete excess Primer Sets")
    If numOfSamplePlate + numOfPrimerPlate > 0 Then
        Sheet10.Select
    ElseIf numOfSampleList + numOfPrimerList > 0 Then
        Sheet1.Select
    End If
    Exit Sub
End If

If numOfPrimerPlate > 0 Then
    Do Until primerCounter > numOfPrimerPlate
        
            Do Until arrayCounter > arrayPosition
                If primerArray(arrayCounter) = Sheet10.Range("P" & positionCounter) Then
                    arrayBoolean = 1
                End If
                arrayCounter = arrayCounter + 1
             Loop
        
        If arrayBoolean = 0 Then
            primerArray(arrayPosition) = Sheet10.Range("P" & positionCounter)
            arrayPosition = arrayPosition + 1
        End If
    If Sheet10.Range("P" & positionCounter) <> "" Then
        primerCounter = primerCounter + 1
    End If
    
    arrayCounter = 0
    positionCounter = positionCounter + 1
    arrayBoolean = 0
    
    Loop
    


End If

If numOfPrimerList > 0 Then

Do Until primerCounter > numOfPrimerList
        
            Do Until arrayCounter > arrayPosition
                If primerArray(arrayCounter) = Sheet1.Range("P" & positionCounter) Then
                    arrayBoolean = 1
                End If
                arrayCounter = arrayCounter + 1
             Loop
        
        If arrayBoolean = 0 Then
            primerArray(arrayPosition) = Sheet1.Range("P" & positionCounter)
            arrayPosition = arrayPosition + 1
        End If
    If Sheet1.Range("P" & positionCounter) <> "" Then
        primerCounter = primerCounter + 1
    End If
    
    arrayCounter = 0
    positionCounter = positionCounter + 1
    arrayBoolean = 0
    
    Loop
End If


Sheet8.Range("B28") = arrayPosition
Sheet8.Range("B30") = primerArray(0)
Sheet8.Range("E30") = primerArray(1)
Sheet8.Range("H30") = primerArray(2)
Sheet8.Range("K30") = primerArray(3)
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
