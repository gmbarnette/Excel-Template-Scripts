
Sub CreateNewPool_Click()

'displays a prompt asking for a name for the New Pool to be created and ends the Macro of Cancel is pushed or a null string is returned
Dim poolName As Variant
Dim poolCounter As Integer
createPool:

poolCounter = 8


poolName = InputBox("What is the name of the new pool?")

If poolName = vbNullString Then
    Exit Sub
End If
    
Do Until poolCounter > Worksheets.Count
    If (poolName = Worksheets(poolCounter).Name) Then
        MsgBox ("Please pick a different pool name")
        GoTo createPool
    Else
        poolCounter = poolCounter + 1
    End If
Loop
        
'creates new Pool Worksheet and Names it from User input above-------------------------------------------------------------------------
Sheets(Worksheets.Count).Select
Sheets(Worksheets.Count).Copy After:=Sheets(Worksheets.Count)

Sheets(Worksheets.Count).Name = poolName

Dim numSamples As Integer
Dim myRange As Range
Dim keeper As Integer
Dim placeKeeper As Integer
    
placeKeeper = 1
keeper = 3
Set myRange = Sheets(Worksheets.Count).Range("B3:B400")
numSamples = Application.WorksheetFunction.CountIf(myRange, "*?") + Application.Count(myRange)
    
    'Deletes Rows that are already included in a pool---------------------------------
    Do Until placeKeeper > numSamples
        If Sheets(Worksheets.Count).Range("B" & keeper) = "" Then
            keeper = keeper + 1
            placeKeeper = placeKeeper
            
        ElseIf Sheets(Worksheets.Count).Range("I" & keeper) = "Y" Then
            Sheets(Worksheets.Count).Rows(keeper).Delete
            keeper = keeper
            placeKeeper = placeKeeper
            numSamples = numSamples - 1
            
        Else
            keeper = keeper + 1
            placeKeeper = placeKeeper + 1
        End If
    Loop
End Sub
Sub DeleteWhiteSamples_Click()

'This button deletes all samples from all pools that are not to be inlcuded in each individual pool (the white samples)

If MsgBox("Are you sure you want to delete samples that are not to be included in each pool? This is permanent and cannot be undone", vbYesNo) = vbNo Then Exit Sub

    Dim sheetNumber As Integer
    Dim numSamples As Integer
    Dim totSamples As Integer
    Dim myRange As Range
    Dim numSamples1 As Integer
    Dim myRange1 As Range
    Dim keeper1 As Integer
    Dim placeKeeper1 As Integer
    Dim sheetCount As Integer
    
    sheetNumber = 8
    
    If Worksheets(Worksheets.Count).Name = "Pool_Data" Then
        sheetCount = Worksheets.Count - 1
    
        Do Until sheetNumber > Worksheets.Count - 1
        Set myRange = Worksheets(sheetNumber).Range("B3:B400")
        numSamples = Worksheets(sheetNumber).Range("F1")
        totSamples = Worksheets(sheetNumber).Application.WorksheetFunction.CountIf(myRange, "*?") + Application.Count(myRange)
        
             'fix this because on last pool doesn't delete gaps in samples
           
                placeKeeper1 = 1
                keeper1 = 3
                Set myRange1 = Worksheets(sheetNumber).Range("B3:B400")
                numSamples1 = Worksheets(sheetNumber).Application.WorksheetFunction.CountIf(myRange1, "*?") + Application.Count(myRange1)
    
                'Deletes rows that are not to be included in the pool----------------------------------
                Do Until placeKeeper1 > numSamples1
                        
                    If Worksheets(sheetNumber).Range("I" & keeper1) = "" Then
                        If Worksheets(sheetNumber).Range("B" & keeper1) = "" Then
                            numSamples1 = numSamples1
                        Else
                            numSamples1 = numSamples1 - 1
                        End If
                        
                        Worksheets(sheetNumber).Rows(keeper1).Delete
                        keeper1 = keeper1
                        placeKeeper1 = placeKeeper1
                       
                    Else
                        keeper1 = keeper1 + 1
                        placeKeeper1 = placeKeeper1 + 1
                    End If
                Loop
        
            
            
        sheetNumber = sheetNumber + 1
        
    Loop
    
    Else
    
    Do Until sheetNumber > Worksheets.Count
        Set myRange = Worksheets(sheetNumber).Range("B3:B400")
        numSamples = Worksheets(sheetNumber).Range("F1")
        totSamples = Worksheets(sheetNumber).Application.WorksheetFunction.CountIf(myRange, "*?") + Application.Count(myRange)
        
    
           
            placeKeeper1 = 1
            keeper1 = 3
            Set myRange1 = Worksheets(sheetNumber).Range("B3:B400")
            numSamples1 = Worksheets(sheetNumber).Application.WorksheetFunction.CountIf(myRange1, "*?") + Application.Count(myRange1)
    
            'Deletes rows that are not to be included in the pool----------------------------------
            Do Until placeKeeper1 > numSamples1
                
                If Worksheets(sheetNumber).Range("I" & keeper1) = "" Then
                    If Worksheets(sheetNumber).Range("B" & keeper1) = "" Then
                            numSamples1 = numSamples1
                    Else
                            numSamples1 = numSamples1 - 1
                    End If
                    
                    Worksheets(sheetNumber).Rows(keeper1).Delete
                    keeper1 = keeper1
                    placeKeeper1 = placeKeeper1
                   
            
                Else
                    keeper1 = keeper1 + 1
                    placeKeeper1 = placeKeeper1 + 1
                End If
            Loop
        
        
            
        sheetNumber = sheetNumber + 1
        
    Loop
    End If
End Sub
Sub MakeBiomekFile_Click()

If MsgBox("Are you sure you want to create the BIOMEK Robot File?", vbYesNo) = vbNo Then Exit Sub

'Deletes the Pool Data Worksheet if it already exists----------------------
    With ThisWorkbook
        
        Application.DisplayAlerts = False
        
        If .Worksheets(.Worksheets.Count).Name = "Pool_Data" Then
            .Worksheets(.Worksheets.Count).Delete
            .Worksheets(.Worksheets.Count).Delete
         End If
        Application.DisplayAlerts = True
        
    End With
    
    Dim sheetNumber As Integer
    Dim numSamples As Integer
    Dim placeKeeper As Integer
    Dim keeper As Integer
    Dim rowKeeper As Integer
    Dim totSamples As Integer
    Dim myRange As Range
    Dim poolPos() As Variant
    Dim locKeeper As Integer
    
    sheetNumber = 8
    
    If Worksheets(8).Range("F1") = 0 Then
       MsgBox ("Please make pools before trying to create robot file")
       Exit Sub
    End If
    
    Do Until sheetNumber > Worksheets.Count
      Set myRange = Worksheets(sheetNumber).Range("B3:B400")
      numSamples = Worksheets(sheetNumber).Range("F1")
      totSamples = Worksheets(sheetNumber).Application.WorksheetFunction.CountIf(myRange, "*?") + Application.Count(myRange)
      
      If numSamples <> totSamples Then
        MsgBox ("Please delete all samples not to be included in pools before creating the robot file")
        Exit Sub
      End If
      
      sheetNumber = sheetNumber + 1
    
    Loop
    
    Dim robotPool As Worksheet
    Dim robotLoc As Worksheet
    
   'Creates the new Worksheet Pool_Data in which all pool data will be added----------
    With ThisWorkbook
        Set robotLoc = .Sheets.Add(After:=.Sheets(Sheets.Count))
        robotLoc.Name = "Robot Pool Location"
        robotLoc.Range("A1") = "Pool Name"
        robotLoc.Range("B1") = "Robot Location"
        
        Set robotPool = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        robotPool.Name = "Pool_Data"
        robotPool.Range("A1") = "Source Plate"
        robotPool.Range("B1") = "S-well"
        robotPool.Range("C1") = "Destination Plate"
        robotPool.Range("D1") = "D-well"
        robotPool.Range("E1") = "Volume"
        
        
    End With
    
    poolPos() = Array("A01", "B01", "C01", "D01", "A02", "B02", "C02", "D02", "A03", "B03", "C03", "D03", "A04", "B04", "C04", "D04", "A05", "B05", "C05", "D05", "A06", "B06", "C06", "D06")
    poolNum = 0
    rowKeeper = 2
    sheetNumber = 8
    locKeeper = 2
    
    Do Until sheetNumber > Worksheets.Count - 2
        Set myRange = Worksheets(sheetNumber).Range("B3:B400")
        numSamples = Worksheets(sheetNumber).Range("F1")
        placeKeeper = 1
        keeper = 3
        totSamples = Worksheets(sheetNumber).Application.WorksheetFunction.CountIf(myRange, "*?") + Application.Count(myRange)
        
        'THIS FUCNTION OF THE CREATE ROBOT FILE BUTTON WAS REMOVED AND A CHECK WAS INSTEAD ADDED TO THE
        'BEGGINING OF THIS FUNCTION TO WARN USERS TO DELETE ALL NON POOL SAMPLES BEFORE CONTINUING.
        'If numSamples <> totSamples Then
           
           
            'Dim numSamples1 As Integer
            'Dim myRange1 As Range
            'Dim keeper1 As Integer
            'Dim placeKeeper1 As Integer
    
            'placeKeeper1 = 1
            'keeper1 = 3
            'Set myRange1 = Worksheets(sheetNumber).Range("B3:B400")
            'numSamples1 = Worksheets(sheetNumber).Application.WorksheetFunction.CountA(myRange1)
    
            'Deletes rows that are not to be included in the pool----------------------------------
            'Do Until placeKeeper1 > numSamples1
                'If Worksheets(sheetNumber).Range("I" & keeper1) = "" Then
                'Worksheets(sheetNumber).Rows(keeper1).Delete
                'keeper1 = keeper1
                'placeKeeper1 = placeKeeper1
                'numSamples1 = numSamples1 - 1
            
            'Else
                'keeper1 = keeper1 + 1
                'placeKeeper1 = placeKeeper1 + 1
            'End If
            'Loop
        
        
        'End If
            
        Do Until placeKeeper > numSamples
            robotPool.Range("A" & rowKeeper) = Worksheets(sheetNumber).Range("C" & keeper)
            robotPool.Range("B" & rowKeeper) = Worksheets(sheetNumber).Range("D" & keeper)
            robotPool.Range("C" & rowKeeper) = "Pools"
            robotPool.Range("E" & rowKeeper) = Format(Worksheets(sheetNumber).Range("H" & keeper), "##.00")
            robotPool.Range("D" & rowKeeper) = poolPos(poolNum)
            placeKeeper = placeKeeper + 1
            rowKeeper = rowKeeper + 1
            keeper = keeper + 1
            
        Loop
        
        Worksheets("Robot Pool Location").Range("A" & locKeeper) = Worksheets(sheetNumber).Name
        Worksheets("Robot Pool Location").Range("B" & locKeeper) = poolPos(poolNum)
        
        locKeeper = locKeeper + 1
        poolNum = poolNum + 1
        sheetNumber = sheetNumber + 1
        
    Loop
    
    robotPool.Copy
    MsgBox ("Robot File Complete. Save the file as a .CSV file in order to use it with the robot!")
End Sub
Sub MakeTecanFile_Click()

If MsgBox("Are you sure you want to create the TECAN Robot File?", vbYesNo) = vbNo Then Exit Sub

'Deletes the Pool Data Worksheet if it already exists----------------------
    With ThisWorkbook
        
        Application.DisplayAlerts = False
        
        If .Worksheets(.Worksheets.Count).Name = "Pool_Data" Then
            .Worksheets(.Worksheets.Count).Delete
            .Worksheets(.Worksheets.Count).Delete
         End If
        Application.DisplayAlerts = True
        
    End With
    
    Dim sheetNumber As Integer
    Dim numSamples As Integer
    Dim placeKeeper As Integer
    Dim keeper As Integer
    Dim rowKeeper As Integer
    Dim totSamples As Integer
    Dim myRange As Range
    Dim poolPos() As Variant
    Dim locKeeper As Integer
    Dim sourcePosRange As Range
    
    sheetNumber = 8
    
    If Worksheets(8).Range("F1") = 0 Then
        MsgBox ("Please make pools before trying to create robot file")
        Exit Sub
    End If
        
    Do Until sheetNumber > Worksheets.Count
      Set myRange = Worksheets(sheetNumber).Range("B3:B400")
      numSamples = Worksheets(sheetNumber).Range("F1")
      totSamples = Worksheets(sheetNumber).Application.WorksheetFunction.CountIf(myRange, "*?") + Application.Count(myRange)
      
      If numSamples <> totSamples Then
        MsgBox ("Please delete all samples not to be included in pools before creating the robot file")
        Exit Sub
      End If
      
      sheetNumber = sheetNumber + 1
    
    Loop
    
    Dim robotPool As Worksheet
    Dim robotLoc As Worksheet
    
   'Creates the new Worksheet Pool_Data in which all pool data will be added----------
    With ThisWorkbook
        Set robotLoc = .Sheets.Add(After:=.Sheets(Sheets.Count))
        robotLoc.Name = "Robot Pool Location"
        robotLoc.Range("A1") = "Pool Name"
        robotLoc.Range("B1") = "Robot Location"
        
        Set robotPool = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        robotPool.Name = "Pool_Data"
        robotPool.Range("A1") = "SourceLabel"
        robotPool.Range("B1") = "SourceWell"
        robotPool.Range("C1") = "DestLabel"
        robotPool.Range("D1") = "DestWell"
        robotPool.Range("E1") = "Volume"
        
        
    End With
    
    
    
    
    poolNum = 1
    rowKeeper = 2
    sheetNumber = 8
    locKeeper = 2
    Set sourcePosRange = Worksheets("Sheet3").Range("H20:I115")
    
    Do Until sheetNumber > Worksheets.Count - 2
        Set myRange = Worksheets(sheetNumber).Range("B3:B400")
        numSamples = Worksheets(sheetNumber).Range("F1")
        placeKeeper = 1
        keeper = 3
        totSamples = Worksheets(sheetNumber).Application.WorksheetFunction.CountIf(myRange, "*?") + Application.Count(myRange)
        'THIS FUCNTION OF THE CREATE ROBOT FILE BUTTON WAS REMOVED AND A CHECK WAS INSTEAD ADDED TO THE
        'BEGGINING OF THIS FUNCTION TO WARN USERS TO DELETE ALL NON POOL SAMPLES BEFORE CONTINUING.
        'If numSamples <> totSamples Then
           
            'Dim numSamples1 As Integer
            'Dim myRange1 As Range
            'Dim keeper1 As Integer
            'Dim placeKeeper1 As Integer
    
            'placeKeeper1 = 1
            'keeper1 = 3
            'Set myRange1 = Worksheets(sheetNumber).Range("B3:B400")
            'numSamples1 = Worksheets(sheetNumber).Application.WorksheetFunction.CountA(myRange1)
    
            'Deletes rows that are not to be included in the pool----------------------------------
            'Do Until placeKeeper1 > numSamples1
                'If Worksheets(sheetNumber).Range("I" & keeper1) = "" Then
                'Worksheets(sheetNumber).Rows(keeper1).Delete
                'keeper1 = keeper1
                'placeKeeper1 = placeKeeper1
                'numSamples1 = numSamples1 - 1
            
            'Else
                'keeper1 = keeper1 + 1
                'placeKeeper1 = placeKeeper1 + 1
            'End If
            'Loop
        
        
        'End If
            
        Do Until placeKeeper > numSamples
        
            robotPool.Range("A" & rowKeeper) = Replace(Worksheets(sheetNumber).Range("C" & keeper), "_", " ")
            robotPool.Range("B" & rowKeeper) = Application.WorksheetFunction.VLookup(Worksheets(sheetNumber).Range("D" & keeper), sourcePosRange, 2, 0)
            robotPool.Range("C" & rowKeeper) = "Pools"
            robotPool.Range("E" & rowKeeper) = Format(Worksheets(sheetNumber).Range("H" & keeper), "##.00")
            robotPool.Range("D" & rowKeeper) = poolNum
            placeKeeper = placeKeeper + 1
            rowKeeper = rowKeeper + 1
            keeper = keeper + 1
            
        Loop
        
        Worksheets("Robot Pool Location").Range("A" & locKeeper) = Worksheets(sheetNumber).Name
        Worksheets("Robot Pool Location").Range("B" & locKeeper) = poolNum
        
        locKeeper = locKeeper + 1
        poolNum = poolNum + 1
        sheetNumber = sheetNumber + 1
        
    Loop
    
    robotPool.Copy
    MsgBox ("Robot File Complete. Save the file as a .CSV file in order to use it with the robot!")
End Sub
