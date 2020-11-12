Attribute VB_Name = "Module2"
Sub QuickTooling()

Dim wB As Workbook

If Range("M5").Value > 0 Then
    On Error Resume Next
    Set wB = Workbooks("Test_wkbk.xlsx")
    If wB Is Nothing Then
        Set wB = Workbooks.Open(Filename:="C:\Users\englandt\Desktop\scratch\Test_wkbk.xlsx")
    End If
    On Error GoTo 0
End If

Call ToolingPricing(wB)

If Range("M5").Value > 0 Then
    wB.Activate
End If

End Sub

Function ToolingPricing(wB As Workbook)

Dim nuM As Integer, i As Integer, x As Boolean
Dim augWB As Workbook, augWbPath As String

augWbPath = "K:\EnglandT\MATEER"
'path for Auger Output Data workbook

Application.ScreenUpdating = False

nuM = 0
i = 5
x = False

On Error GoTo errhandler

ThisWorkbook.Activate

If Range("M5").Value > 0 Then
    Sheet1.Activate
    Sheet1.Unprotect
    If Range("M5").Value > 0 Then
        If Range("M6").Value > 0 Then
            Range("M5").Select
            Selection.End(xlDown).Select
            nuM = ActiveCell.Row
            
            For i = 14 To 20
                If Cells(nuM + 1, i).Value > 0 Then
                    x = True
                End If
            Next
            
            If x = True Then
                MsgBox "Tooling will not be found for rows without a 'Product' description."
            End If
           
        Else
            nuM = 5
        End If
    End If
    
    x = True
    
    On Error Resume Next
    Set augWB = Workbooks("Auger Output Data.xlsx")
    If augWB Is Nothing Then
        Set augWB = Workbooks.Open(Filename:=augWbPath & "\Auger Output Data.xlsx")
    End If
    On Error GoTo errhandler
    
    If nuM > 0 Then
        For i = 5 To nuM
                x = FindTooling(i, wB, augWB)
        Next

    Else
        MsgBox "Each line item must contain a 'Product' description."
    End If
    
    augWB.Close savechanges:=False
    wB.Activate
    Columns(2).AutoFit
    Range("C5:C" & nuM).NumberFormat = "0.000"
    wB.Worksheets(1).Activate
    ThisWorkbook.Activate
    
End If

Sheet1.Activate
Range("A1").Select

Exit Function

errhandler:
    On Error GoTo errhandler2
    wB.Activate
    wB.Worksheets(1).Activate
    MsgBox "Error in ToolingPricing function"
    Exit Function

errhandler2:
    MsgBox "Error in ToolingPricing function"

End Function

Function FindTooling(i As Integer, wB As Workbook, augWB As Workbook) As Boolean

Dim x As Integer, y As Integer, j As Integer, augLen As Single, funDia As Single, diaDiff As Single
Dim wkBk As Workbook, cB As Boolean, rrY As Boolean, vffS As Boolean, indTime As Single, ffNffCol As Integer
Dim flowType As String, deN As Variant, openDia As Variant, cpmReq As Variant, ratE As Integer, numAugs As Integer
Dim searchVal As Variant, timE As Variant

Application.ScreenUpdating = False

x = 0
cB = False
rrY = False
vffS = False
Set wkBk = ThisWorkbook

wkBk.Activate
Sheet1.Activate
ActiveSheet.Unprotect
Range("M1").Formula = "=iferror(value(right(left(D6,2),1)),18)"
numAugs = 1 'more for dual head rotaries

If Range("M1").Value = 9 Then
    cB = True
ElseIf Range("M1").Value > 5 And Range("M1").Value < 8 Then
    rrY = True
    If Range("M1").Value = 7 Then
        numAugs = 2 * Range("D15").Value
        If numAugs = 0 Then numAugs = 1
    End If
End If
Range("M1").Clear

If UCase(Range("D13").Value) = "YES" Or UCase(Range("D13").Value) = "Y" Or Range("D13").Value = 1 Then
    vffS = True
End If

On Error GoTo errhandler

If Application.WorksheetFunction.IsNumber(Range("N" & i).Value) Then
    x = 1
    'tooling is given. find other stuff?
Else
    If Application.WorksheetFunction.IsNumber(Range("R" & i).Value) Then
        If Application.WorksheetFunction.IsNumber(Range("P" & i).Value) Then
            'fill weight & density are known
            x = 5
        Else
            x = 3
        End If
    Else
        x = 4
    End If
End If

wB.Activate
wB.Worksheets("Tooling").Activate

'new wB row "y"
If Range("A2").Value > 0 Then
    If Range("A3").Value > 0 Then
        Range("A2").Select
        Selection.End(xlDown).Select
        y = ActiveCell.Row + 1
    Else
        y = 3
    End If
Else
    y = 2
End If

On Error Resume Next
    ThisWorkbook.Activate
    Sheet3.Activate
    indTime = Range("C20").Value
On Error GoTo errhandler
If indTime = 0 Then
    indTime = 4
End If

If rrY Then
    indTime = 0
End If

Sheet1.Activate

If Range("Q" & i).Value = 0 Then
    MsgBox "Tooling row #" & i & " (" & Range("M" & i).Value & ") doesn't specify flow type. Free flow will be assumed."
    Range("Q" & i).Value = "FF"
End If

If x = 1 Then
    If Range("P" & i).Value > 0 And IsNumeric(Range("P" & i).Value) And Range("R" & i).Value > 0 And IsNumeric(Range("R" & i).Value) Then
        prodName = Range("M" & i).Value
        tooL = Range("N" & i).Value
        fillWeight = Range("P" & i).Value
        flowType = Range("Q" & i).Value
        deN = Range("R" & i).Value
        
        augWB.Activate
        Range("L2").Value = fillWeight / numAugs
        Range("L3").Value = deN
        Range("B2").Value = flowType
        If cB Then
            Range("B3").Value = 19
        End If
        
        timE = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:M50"), 13, False)
        numTurns = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:M50"), 12, False)
        numTurns = Round(numTurns, 2)
        ratE = WorksheetFunction.Floor_Math(60 / (timE + indTime))
        
        wB.Activate
        Range("A" & y).Value = fillWeight
        Range("B" & y).Value = prodName
        Range("C" & y).Value = deN
        Range("D" & y).Value = timE
        If cpmReq = 0 Then
            Range("E" & y).Formula = "=Floor.Math(60 / (D" & y & "+K" & y & "))"
        Else
            Range("E" & y).Formula = "=MIN(" & cpmReq & ",Floor.Math(60 / (D" & y & "+K" & y & ")))"
        End If
        Range("F" & y).Value = numTurns
        Range("G" & y).Value = tooL
        Range("J" & y).Value = flowType
        Range("K" & y).Value = indTime
        
    Else
        fillWeight = Range("P" & i).Value
        If fillWeight = 0 Then
            fillWeight = 0.001
        End If
        prodName = Range("M" & i).Value
        tooL = Range("N" & i).Value
        flowType = Range("Q" & i).Value
        
        wB.Activate
        If fillWeight < 1 Then
            Range("A" & y).Value = "TBD"
        Else
            Range("A" & y).Value = fillWeight
        End If
        Range("B" & y).Value = prodName
        Range("G" & y).Value = tooL
        Range("J" & y).Value = flowType
        Range("K" & y).Value = indTime
        
        For j = 3 To 6
            ActiveSheet.Cells(y, j).Value = "TBD"
        Next
        
    End If
    
ElseIf x = 5 Then

    prodName = Range("M" & i).Value
    fillWeight = Range("P" & i).Value
    flowType = Range("Q" & i).Value
    deN = Range("R" & i).Value
    cpmReq = 0
    
    If IsNumeric(Range("T" & i).Value) And Range("T" & i).Value > 0 Then
        cpmReq = Range("T" & i).Value
    End If
    
    augWB.Activate
    Range("L2").Value = fillWeight / numAugs
    Range("L3").Value = deN
    Range("B2").Value = flowType
    
    If flowType = "FF" Then
        ffNffCol = 15
    Else
        ffNffCol = 14
    End If
    
    If cB Then
        Range("B3").Value = 19
    End If
    
    On Error Resume Next
    tooL = WorksheetFunction.indeX(ActiveSheet.Range("A9:A29"), Application.Match(1, ActiveSheet.Range("B9:B29"), 0))
    On Error GoTo errhandler
    
    If tooL = 0 Then
        '52 is not big enough
        tooL = 52
    End If
    
    If rrY And cpmReq > 0 Then 'get tool size by index time
        Cells(Application.WorksheetFunction.Match(tooL, Range("A:A"), 0), 13).Select
        Do While ActiveCell.Value > 60 / cpmReq
            Selection.Offset(1, 0).Select
        Loop
        
        If ActiveCell.Value = 0 Then Selection.Offset(-1, 0).Select
        tooL = Range("A" & ActiveCell.Row).Value
        
    End If
    
    timE = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:M50"), 13, False)
    numTurns = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:M50"), 12, False)
    numTurns = Round(numTurns, 2)
    ratE = WorksheetFunction.Floor_Math(60 / (timE + indTime))
    funDia = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:O50"), ffNffCol, False)

    wkBk.Activate
    If cpmReq > 0 And Not rrY Then
        If ratE < cpmReq Then
            augWB.Activate
            Do While ratE < cpmReq
            
                tooL = tooL + 2
                
                If tooL = 34 Or tooL = 42 Or tooL = 46 Then
                    tooL = tooL + 2
                End If
                
                If tooL < 6 Or tooL > 52 Then
                    MsgBox "Tooling must be calculated manually (item #" & i - 4 & ")"
                    FindTooling = False
                    Exit Function
                End If
                'msgbox "tooL: " & tooL
                timE = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:M50"), 13, False)
                ratE = WorksheetFunction.Floor_Math(60 / (timE + indTime))
                
            Loop
            
            numTurns = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:M50"), 12, False)
            numTurns = Round(numTurns, 2)
            funDia = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:O50"), ffNffCol, False)
            
        End If
    End If
    
    If Not rrY Then 'rotaries don't care about opening diam
        If Application.WorksheetFunction.IsNumber(Range("S" & i).Value) And Range("S" & i).Value > 0 Then
            openDia = Range("S" & i).Value
            If vffS Then
                diaDiff = 0.12
            Else
                diaDiff = 0.3
            End If
            
            If openDia < funDia + diaDiff And rrY = False Then
                augWB.Activate
                
                Do While openDia < funDia + diaDiff
                
                    tooL = tooL - 2
                    
                    If tooL = 34 Or tooL = 42 Or tooL = 46 Then
                        tooL = tooL - 2
                    End If
                    
                    If tooL < 6 Or tooL > 52 Then
                        MsgBox "Tooling must be calculated manually (item #" & i - 4 & ")"
                        FindTooling = False
                        Exit Function
                    End If
                    
                    funDia = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:O50"), ffNffCol, False)
                    
                Loop
                
                timE = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:M50"), 13, False)
                numTurns = WorksheetFunction.VLookup(tooL, ActiveSheet.Range("A9:M50"), 12, False)
                numTurns = Round(numTurns, 2)
                ratE = WorksheetFunction.Floor_Math(60 / (timE + indTime))
                
                If ratE < cpmReq Then
                    MsgBox "The container for " & prodName & " (item #" & i - 4 & ") requires the rate not to be met."
                End If
                        
            End If
        End If
    End If
    
    If cpmReq > 0 And ratE > cpmReq Then
        ratE = cpmReq 'Dont want to promise more than the customer is seeking
    End If
    
    wB.Activate
    Range("A" & y).Value = fillWeight
    Range("B" & y).Value = prodName
    Range("C" & y).Value = deN
    Range("D" & y).Value = timE
    If cpmReq = 0 Then
        Range("E" & y).Formula = "=Floor.Math(60 / (D" & y & "+K" & y & "))"
    Else
        Range("E" & y).Formula = "=MIN(" & cpmReq & ",Floor.Math(60 / (D" & y & "+K" & y & ")))"
    End If
    Range("F" & y).Value = numTurns
    Range("G" & y).Value = tooL
    Range("J" & y).Value = flowType
    Range("K" & y).Value = indTime
    
Else
    prodName = Range("M" & i).Value
    tooL = Range("N" & i).Value
    fillWeight = Range("P" & i).Value
    flowType = Range("Q" & i).Value
    deN = Range("R" & i).Value
    
    wB.Activate
    Range("A" & y).Value = fillWeight
    Range("B" & y).Value = prodName
    Range("C" & y).Value = deN
    Range("D" & y).Value = timE
    Range("E" & y).Formula = "=Floor.Math(60 / (D" & y & "+K" & y & "))"
    Range("F" & y).Value = numTurns
    Range("G" & y).Value = tooL
    Range("J" & y).Value = flowType
    Range("K" & y).Value = indTime
    
    If Range("A" & y).Value = 0 Then
        Range("A" & y).Value = "TBD"
    End If
    
    For j = 3 To 7
        If ActiveSheet.Cells(y, j).Value = 0 Then
            ActiveSheet.Cells(y, j).Value = "TBD"
        End If
    Next
    
End If

wkBk.Activate
Sheet1.Activate
augLen = Range("O" & i).Value
augLen = Round(augLen, 1)
wB.Activate
FindTooling = False

If x > 0 Then
    If vffS Then
        Range("I" & y).Value = "OEM"
    ElseIf augLen > 0 Then
        If augLen = 13.1 Then
            Range("I" & y).Value = "Short"
        ElseIf augLen = 20.1 Then
            Range("I" & y).Value = "Long"
        Else
            Range("I" & y).Value = "OEM"
        End If
    Else
        If rrY Then
            Range("I" & y).Value = "Long"
        Else
            Range("I" & y).Value = "Short"
        End If
    End If
    
    FindTooling = True
End If

Exit Function

errhandler: MsgBox "Error in FindTooling function"

End Function

Function CopyToolingPricing(wB As Workbook, templFile As String, steelType As String)

Dim costBook As Workbook, priceBook As Workbook, lastRow As Integer, i As Integer
Dim costBookPath As String, priceBookPath As String, optionsRow As Integer

On Error GoTo errhandler

'path designations
''''''''''''''''''
costBookPath = ThisWorkbook.Path
priceBookPath = costBookPath
''''''''''''''''''
'path designations

Application.ScreenUpdating = False

On Error Resume Next
    Set costBook = Workbooks("CostBook_Mateer.xlsm")
    If costBook Is Nothing Then
        Set costBook = Workbooks.Open(Filename:=costBookPath & "\CostBook_Mateer.xlsm")
    End If
    Set priceBook = Workbooks("PriceBook_Mateer.xlsm")
    If priceBook Is Nothing Then
        Set priceBook = Workbooks.Open(Filename:=priceBookPath & "\PriceBook_Mateer.xlsm")
    End If
On Error GoTo errhandler

wB.Activate
wB.Worksheets("Tooling").Activate

If templFile = "Pricing_Template" Then
    If Range("A2").Value > 0 Then
        Call ToolingLineItem(2, steelType, wB, costBook, priceBook, templFile)
        lastRow = 2
        
        wB.Worksheets(1).Activate
        wB.Worksheets("Tooling").Activate
        
        If Range("A3").Value > 0 Then
            Range("A2").Select
            Selection.End(xlDown).Select
            lastRow = ActiveCell.Row
            
            wB.Worksheets(1).Activate
            Rows("12:" & lastRow + 9).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            wB.Worksheets("Tooling").Activate

            For i = 3 To lastRow
                Call ToolingLineItem(i, steelType, wB, costBook, priceBook, templFile)
            Next
            
        End If
        
        wB.Worksheets(1).Activate
    
        Range("B1").Formula = "=match(" & """" & "Zepf*" & """" & ",A:A,0)"
        On Error Resume Next
            optionsRow = Range("B1").Value + 3
            If optionsRow = 0 Then
                Range("B1").Formula = "=match(" & """" & "Options" & """" & ",A:A,0)"
                optionsRow = Range("B1").Value
                If optionsRow = 0 Then
                    Range("B1").Formula = "=match(" & """" & "Spare Parts Estimate" & """" & ",A:A,0)"
                    optionsRow = Range("B1").Value
                End If
            End If
            
        On Error GoTo errhandler
        Range("B1").Clear
        Range("F" & optionsRow).Select
        Selection.Copy
        Range("F" & optionsRow - 1).Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        
        
         If lastRow > 2 And Range("A12").Value > 0 Then
            Range("B11:D11").Select
            Selection.AutoFill Destination:=Range("B11:D" & optionsRow - 2)
            
            Range("G11:H11").Select
            Selection.AutoFill Destination:=Range("G11:H" & optionsRow - 2)
            
            Range("J11").Select
            Selection.AutoFill Destination:=Range("J11:J" & optionsRow - 2)
            
            Range("Q11").Select
        End If
        
        Range("A1").Select
        
    End If
    
ElseIf templFile = "AM_Pricing" Then
    
    If Range("A2").Value > 0 Then
    
        Call ToolingLineItem(2, steelType, wB, costBook, priceBook, templFile)
        lastRow = 2
        
        wB.Worksheets("Tooling").Activate

        If Range("A3").Value > 0 Then
            Range("A2").Select
            Selection.End(xlDown).Select
            lastRow = ActiveCell.Row

            For i = 3 To lastRow
                Call ToolingLineItem(i, steelType, wB, costBook, priceBook, templFile)
            Next

        End If
        
        wB.Worksheets(1).Activate
        
        Range("A4").Select
        Selection.End(xlDown).Select
        lastRow = ActiveCell.Row
        
        If lastRow < 1000 Then
            Range("F4:I4").Select
            Selection.AutoFill Destination:=Range("F4:I" & lastRow)
            
            Range("K4").Select
            Selection.AutoFill Destination:=Range("K4:K" & lastRow)
            
            Range("M4").Select
            Selection.AutoFill Destination:=Range("M4:M" & lastRow)

            For i = 5 To lastRow
                If Range("J" & i).Value = 0 Then
                    Range("J" & i).Value = 1
                End If
            Next
            
        End If
        
    End If
    
    Range("A1").Select

End If

costBook.Close savechanges:=False
priceBook.Close savechanges:=False

Exit Function

errhandler:
On Error Resume Next
priceBook.Close savechanges:=False
costBook.Close savechanges:=False
MsgBox "Error in CopyToolingPricing function"

End Function

Function CopyBudgTooling(wB As Workbook, templFile As String)

Dim costBook As Workbook, priceBook As Workbook, steelType As String, vffS As Boolean, engRate As Integer
Dim ffCost As Single, ffPrice As Single, nffCost As Single, nffPrice As Single, mfgRate As Integer
Dim costBookPath As String, priceBookPath As String, toolingSpec As Boolean, numTool As Integer, oemCol As Integer
Dim augWB As Workbook, newWB As Workbook, augWbPath As String, x As Boolean, i As Integer, oeM As Boolean

On Error GoTo errhandler

'path designations
''''''''''''''''''
augWbPath = "K:\EnglandT\MATEER"
costBookPath = ThisWorkbook.Path
priceBookPath = costBookPath
''''''''''''''''''
'path designations

Application.ScreenUpdating = False

steelType = "304"
vffS = False
ffCost = 100000
ffPrice = 10000
nffCost = 100000
nffPrice = 10000

ThisWorkbook.Activate
Sheet1.Activate

If Application.WorksheetFunction.VLookup("316 Product Contact Parts", Range("C:E"), 3, False) Then
    steelType = "316"
End If

If UCase(Range("D13").Value) = "YES" Or UCase(Range("D13").Value) = "Y" Or Range("D13").Value = 1 Then
    vffS = True
End If

If Range("M5").Value > 0 Then
    toolingSpec = True
    numTool = 1
    'one tooling spec

    If Range("M6").Value > 0 Then
        Range("M5").Select
        Selection.End(xlDown).Select
        numTool = ActiveCell.Row - 4
        'mult tooling spec
        
    End If
    
Else
    toolingSpec = False
    numTool = 2
    'previous function put one FF line and
    'one NFF line on pricing sheet
End If
    
If Not toolingSpec Then
    
    On Error Resume Next
    Set costBook = Workbooks("CostBook_Mateer.xlsm")
    If costBook Is Nothing Then
        Set costBook = Workbooks.Open(Filename:=costBookPath & "\CostBook_Mateer.xlsm")
    End If
    Set priceBook = Workbooks("PriceBook_Mateer.xlsm")
    If priceBook Is Nothing Then
        Set priceBook = Workbooks.Open(Filename:=priceBookPath & "\PriceBook_Mateer.xlsm")
    End If
    On Error GoTo errhandler
    
    costBook.Activate
    x = True 'find the labor rates (ENG)
    i = 1
    Do While x = True And i < 10
        If Right(costBook.Sheets(1).Range("E2").Value, i) > "" Then
            engRate = Right(costBook.Sheets(1).Range("E2").Value, i)
            x = IsNumeric(Right(costBook.Sheets(1).Range("E2").Value, i + 1))
        End If
        i = i + 1
    Loop
    MsgBox "Alpha: ENG rate = " & engRate
    
    x = True 'find the labor rates (MFG)
    i = 1
    Do While x = True And i < 10
        If Right(costBook.Sheets(1).Range("F2").Value, i) > "" Then
            mfgRate = Right(costBook.Sheets(1).Range("F2").Value, i)
            x = IsNumeric(Right(costBook.Sheets(1).Range("F2").Value, i + 1))
        End If
        i = i + 1
    Loop
    
    costBook.Worksheets("Tooling").Activate

    Range("E55").Value = 52

    If vffS = False Then
        ffCost = Range("E56").Value
        nffCost = Range("E58").Value
    Else
        ffCost = Range("E57").Value
        Range("D61").Value = 50
        nffCost = Range("E60").Value
        Range("D61").ClearContents
    End If
    
    costBook.Close savechanges:=False

    priceBook.Activate
    priceBook.Worksheets("Tooling").Activate

    Range("J3").Value = 52
    
    If vffS = False Then
        Range("H3").Value = "FF"
        Range("N3").Value = "Set"
        ffPrice = Range("P3").Value
        
        Range("H3").Value = "NFF"
        Range("L3").Value = 5.5625
        nffPrice = Range("P3").Value
    Else
        Range("H3").Value = "FF"
        Range("N3").Value = "Set & Tube"
        ffPrice = Range("P3").Value
        
        Range("H3").Value = "NFF"
        Range("L3").Value = 42.5625
        Range("N3").Value = "Set"
        nffPrice = Range("P3").Value
        
        If oeM Then
            oemCol = Application.WorksheetFunction.Match("OEM", Range("6:6"), False) - 1
            nffPrice = Application.WorksheetFunction.VLookup(tooL, Range("B:U"), oemCol, False) 'aug & funn
            nffPrice = nffPrice + Range("AH7").Value 'SSA
        End If
    End If
    
    priceBook.Close savechanges:=False
    
    If steelType = "316" Then
        ffCost = ffCost * 1.5
        nffCost = nffCost * 1.5
        ffPrice = ffPrice * 1.5
        nffPrice = nffPrice * 1.5
    End If
    
    wB.Activate
    
    If templFile = "AM_Pricing" Then
        Range("A5").Value = "#52 FF tooling: 304 SST, short with SSA"
        Range("A6").Value = "#52 NFF tooling: 304 SST"
        If vffS Then
            Range("A5").Value = Range("A5").Value & " & drop tube)"
            Range("A6").Value = Range("A6").Value & ", OEM length with SSA"
        Else
            Range("A6").Value = Range("A6").Value & ", short with SSA"
        End If
        Range("B5").Value = ffPrice
        Range("B6").Value = nffPrice
        Range("J5").Value = 1
        Range("J6").Value = 1
        Range("L5").Value = ffCost
        Range("L6").Value = nffCost
        Range("A1").Select
    Else
        Range("A11").Value = "#52 free flow tooling: 304 stainless steel, standard short, with slow speed agitator"
        If vffS Then
            Range("A11").Value = Range("A11").Value & " and drop tube"
        End If
        Range("L11").Value = ffCost
        Range("M11").Value = "0"
        Range("N11").Value = "0"
        Range("F11").Value = ffPrice
        Range("M1").Value = engRate
        Range("N1").Value = mfgRate
        Range("P11").Formula = "=L11+$M$1*M11+$N$1*N11"
    End If

Else
    
    If templFile = "Pricing_Template" Then
        Exit Function
    End If
    '
    'Following code redundant? Isn't ToolingPricing already doing this??
    '
    If numTool > 0 Then
    
    wB.Activate
    Range("A5:L6").ClearContents
        
        x = True
        On Error Resume Next
        Set augWB = Workbooks("Auger Output Data.xlsx")
        If augWB Is Nothing Then
            Set augWB = Workbooks.Open(Filename:=augWbPath & "\Auger Output Data.xlsx")
        End If
        On Error GoTo errhandler
        
        ThisWorkbook.Activate
        
        For i = 5 To numTool + 4
            If x = True Then
                x = FindTooling(i, wB, augWB)
            End If
        Next
        
        
        augWB.Close savechanges:=False
        
        Call CopyToolingPricing(wB, templFile, steelType)
        
        wB.Activate
        wB.Worksheets(1).Activate
        ThisWorkbook.Activate
        
    Else
        MsgBox "Each line item must contain a 'Product' description."
    End If
    
End If

Exit Function

errhandler: MsgBox "Error in CopyBudgTooling function"

End Function

Function ToolingLineItem(i As Integer, steelType As String, wB As Workbook, costBook As Workbook, priceBook As Workbook, templFile As String)

Dim tooL As Integer, toolType As String, flowType As String, toolLen As String, toolRow As Integer
Dim outpuT As String, pricE As Single, cosT As Single, funLen As Single, uniquE As Boolean, j As Integer
Dim engRate As Integer, mfgRate As Integer, x As Boolean, oeM As Boolean, oemCol As Integer

On Error GoTo errhandler

Application.ScreenUpdating = False

ThisWorkbook.Activate
Sheet1.Activate
If Range("E8").Value = "OEM" Then
    oeM = True
Else
    oeM = False
End If

wB.Activate
wB.Worksheets("Tooling").Activate

If Range("J" & i).Value = "FF" Or Range("I" & i).Value <> "OEM" Then
    oeM = False 'oem pricing is for NFF tooling
End If

If IsNumeric(Range("G" & i).Value) Then
    tooL = Range("G" & i).Value
Else
    tooL = 52
End If

'tooling length
If Range("I" & i).Value = "Short" Then
    toolLen = "standard short"
    funLen = 5.5625
ElseIf Range("I" & i).Value = "Long" Then
    toolLen = "standard long"
    funLen = 12.5625
Else
    toolLen = "OEM length"
    ThisWorkbook.Activate
    If Range("O" & i + 3).Value > 0 Then
        funLen = Range("O" & i + 3).Value - 7.5625
    Else
        funLen = 42.5625
    End If
    wB.Activate
End If
    
'flow type
If Range("J" & i).Value = "SF/NFF" Then
    toolType = "non-free flow"
    flowType = "NFF"
Else
    toolType = "free flow"
    flowType = "FF"
    toolLen = "standard short"
    oeM = False
End If

outpuT = "#" & tooL & " " & toolType & " product tooling: " & steelType & " stainless steel " & toolLen & " tooling with slow speed agitator"

'get tooling cost
costBook.Activate

x = True 'find the labor rates (ENG)
j = 1
Do While x = True And j < 10
    If Right(costBook.Sheets(1).Range("E2").Value, j) > "" Then
        engRate = Right(costBook.Sheets(1).Range("E2").Value, j)
        x = IsNumeric(Right(costBook.Sheets(1).Range("E2").Value, j + 1))
    End If
    j = j + 1
Loop
MsgBox "Beta: ENG rate = " & engRate

x = True 'find the labor rates (MFG)
j = 1
Do While x = True And j < 10
    If Right(costBook.Sheets(1).Range("F2").Value, j) > "" Then
        mfgRate = Right(costBook.Sheets(1).Range("F2").Value, j)
        x = IsNumeric(Right(costBook.Sheets(1).Range("F2").Value, j + 1))
    End If
    j = j + 1
Loop

costBook.Worksheets("Tooling").Activate
Range("E55").Value = tooL

If flowType = "FF" Then
    If wB.Worksheets("Tooling").Range("I" & i).Value = "Short" Then
        cosT = Range("E56").Value
    Else
        cosT = Range("E57").Value 'adder for drop tube
        outpuT = outpuT & " and drop tube"
    End If
Else
    If toolLen = "standard short" Then
        cosT = Range("E58").Value
    ElseIf toolLen = "standard long" Then
        cosT = Range("E59").Value
    Else
        Range("D61").Value = funLen + 7.5625
        cosT = Range("E60").Value
        Range("D61").ClearContents
    End If
End If

'get tooling price
priceBook.Activate
priceBook.Worksheets("Tooling").Activate
Range("J3").Value = tooL
Range("N3").Value = "Set"

If flowType = "FF" Then
    Range("H3").Value = "FF"
    Range("L3").Value = 5.5625
    If wB.Worksheets("Tooling").Range("I" & i).Value <> "Short" Then
        Range("N3").Value = "Set & Tube"
    End If
Else
    Range("H3").Value = "NFF"
    Range("L3").Value = funLen
End If

pricE = Range("P3").Value

If oeM Then 'oem tooling
    oemCol = Application.WorksheetFunction.Match("OEM", Range("6:6"), False) - 1
    pricE = Application.WorksheetFunction.VLookup(tooL, Range("B:U"), oemCol, False) 'aug & funn
    pricE = pricE + Range("AH7").Value 'SSA
End If

If steelType = "316" Then
    cosT = cosT * 1.5
    pricE = pricE * 1.5
End If

'put on front sheet
uniquE = True
wB.Activate
wB.Worksheets(1).Activate

If templFile = "Pricing_Template" Then

    For j = 11 To 11 + i
        If Range("A" & j).Value = outpuT Then
            uniquE = False
        End If
    Next
    
    toolRow = 11
    Do While Range("A" & toolRow).Value Like "*tooling*"
        toolRow = toolRow + 1
    Loop
    
    If i = 2 Then
        toolRow = 11
    End If
    
    If uniquE = True Or i = 2 Then
        Range("A" & toolRow).Value = outpuT
        Range("F" & toolRow).Value = pricE
        Range("L" & toolRow).Value = cosT
        Range("M1").Value = engRate
        Range("M" & toolRow).Value = "0"
        Range("N1").Value = mfgRate
        Range("N" & toolRow).Value = "0"
        Range("P" & toolRow).Formula = "=L" & toolRow & "+$M$1*M" & toolRow & "+$N$1*N" & toolRow
    Else
        Rows(toolRow & ":" & toolRow).Delete
    End If
    
    wB.Worksheets("Tooling").Activate


ElseIf templFile = "AM_Pricing" Then
    outpuT = tooL & " tooling (" & flowType & ", " & steelType & " SST, " & toolLen & ")"
    
    For j = 1 To 3 + i
        If Range("A" & j).Value = outpuT Then
            uniquE = False
        End If
    Next
    
    If uniquE = True Then
        If Range("A" & 2 + i).Value = 0 Then
            Range("A" & 2 + i).Value = outpuT
            Range("B" & 2 + i).Value = pricE
            Range("L" & 2 + i).Value = cosT
        Else
            Rows(3 + i & ":" & 3 + i).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("A" & 3 + i).Value = outpuT
            Range("B" & 3 + i).Value = pricE
            Range("L" & 3 + i).Value = cosT
        End If
        
    End If
End If

Exit Function

errhandler: MsgBox "Error in ToolingLineItem function"

End Function


Function Option316(wB As Workbook)

Dim nomPrice As Integer, addeR As Integer, increaseD As Integer, finPrice As Integer

On Error GoTo errhandler

Application.ScreenUpdating = False

finPrice = 0

ThisWorkbook.Activate
Sheet1.Activate

If Application.WorksheetFunction.VLookup("316 Product Contact Parts", Range("C:E"), 3, False) Then
    
    Sheet4.Activate
    nomPrice = Application.WorksheetFunction.VLookup("316 Product Contact Parts", Range("B:G"), 6, False)
    addeR = 0
    Sheet1.Activate
    
    If Application.WorksheetFunction.VLookup("High Speed Agitator", Range("C:E"), 3, False) Then
        Sheet4.Activate
        addeR = 0.5 * Application.WorksheetFunction.VLookup("High Speed Agitator", Range("B:G"), 6, False)
        Sheet1.Activate
    End If
    
    increaseD = nomPrice + addeR
    
    wB.Activate
    Sheet1.Activate
    addeR = 0
    
    'was going to add price increase for 316 tooling but
    'eventually decided to keep tooling markup separate
'
'    If Range("F11").Value > 0 Then
'        If Range("F12").Value > 0 Then
'            Range("F11").Select
'            Selection.End(xlDown).Select
'            For i = 11 To ActiveCell.row
'                addeR = addeR + (0.5 * Range("F" & i).Value)
'            Next
'        Else
'            addeR = 0.5 * Range("F11").Value
'        End If
'    End If
    
    finPrice = increaseD + addeR

End If



If finPrice > nomPrice Then
    wB.Activate
    wB.Worksheets(1).Activate
    Range("B1").Formula = "=IFERROR(MATCH(" & """" & "316 Product Contact Parts" & """" & ",A:A,0),6)"
    Range("F" & Range("B1").Value).Select
    ActiveCell.Value = finPrice
    Range("J" & Range("B1").Value).Select
    ActiveCell.Value = "Price book (standard + HSA)"
    Range("A1").Select
    Range("B1").Clear
    ThisWorkbook.Activate
    
End If

Exit Function

errhandler: MsgBox "Error in Option316 function"

End Function
Function LinkAppSummary(quotE As String, oQuote As String, yeaR As Integer, wDoc As Object)
'
' Pastes pricing sheet into word doc
'

Dim testFol As String, testFil As String, wApp As Word.Application
Dim wBook As Workbook, rangeEnd As Integer, quoteFolderPath As String

On Error GoTo errhandler

'path designations
''''''''''''''''''
quoteFolderPath = "T:\Quotes\Mateer\" & yeaR & " Quotes"
''''''''''''''''''
'path designations

Application.DisplayAlerts = False

testFol = Dir(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
'see if quote folder exists

If testFol <> "" Then

    testFil = Dir(quoteFolderPath & "\" & testFol & "\" & quotE & "*" & ".xls*")
    'see if excel doc exists

    If testFil <> "" Then
    
        On Error Resume Next
        Set wBook = Workbooks(testFil)
        If wBook Is Nothing Then
            Set wBook = Workbooks.Open(Filename:=quoteFolderPath & "\" & testFol & "\" & testFil, UpdateLinks:=True)
        End If
        On Error GoTo errhandler
        
        wBook.Activate
        wBook.Worksheets("Tooling").Activate
   
        If Range("A2").Value > 0 Then
            Range("A1").Select
            Selection.End(xlDown).Select
            rangeEnd = Selection.Row
            Range("A1:G" & rangeEnd).Copy
        Else
            Range("A1:G2").Select
            Selection.Copy
        End If
    
            wDoc.Bookmarks("AppSummary").Range.PasteExcelTable LinkedToExcel:=False, _
                WordFormatting:=False, RTF:=False
                
            wBook.Close (savechanges = False)
        
        
'        msgbox "w"
'        wApp.Visible = True
'        msgbox "x"
'        wApp.Activate
'        msgbox "y"
'        wDoc.Bookmarks("AppSummary").Range.PasteExcelTable LinkedToExcel:=False, _
'                WordFormatting:=False, RTF:=False
'        msgbox "z"
    End If

End If

Application.DisplayAlerts = True
wDoc.Activate

Exit Function

errhandler: MsgBox "Error in LinkAppSummary function"

End Function

