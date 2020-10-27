Attribute VB_Name = "Module3"
Public Function CopyEmailWkbks(oQuote As String, yeaR As Integer)

Dim quoteFolderPath As String, i As Integer, salesRep As String, nameArray() As String, customeR As String
Dim outlookApp As Object, eMail As MailItem, testInspect As Inspector, numEmails As Integer, senderSurname As String
Dim emailContent As String, emailFound As Boolean

''''''''''
quoteFolderPath = "T:\Quotes\Mateer\" & yeaR & " Quotes"
''''''''''

On Error Resume Next
Set outlookApp = GetObject(, "Outlook.Application")
On Error GoTo errhandler

If outlookApp Is Nothing Then
    Exit Function
End If
    
quoteFolder = Dir$(quoteFolderPath & "\" & oQuote & "*", vbDirectory)

If quoteFolder = "" Then
    Exit Function
End If

Set testInspect = outlookApp.ActiveInspector

If testInspect Is Nothing Then
    Exit Function
End If

testInspect = Dir$(quoteFolderPath & "\" & quoteFolder & "\" & oQuote & "*.msg")

If testInspect <> "" Then
    Exit Function
End If

numEmails = outlookApp.Inspectors.Count

If numEmails = 0 Then
    Exit Function
End If

ThisWorkbook.Activate
Sheet3.Activate
salesRep = Range("C7").Value
customeR = Range("C8").Value
nameArray() = Split(salesRep)
salesRep = nameArray(1)
emailFound = False

For i = 1 To numEmails
    
    Set eMail = outlookApp.Inspectors.Item(i).CurrentItem
    
    senderSurname = eMail.SenderName
    
    nameArray() = Split(senderSurname, ",")
    senderSurname = nameArray(0)
    
    If eMail.Sent Then
        If senderSurname = salesRep Then
            emailContent = eMail.Body
            If InStr(1, emailContent, customeR) > 0 Then
                emailFound = True
                Exit For
            End If
        End If
    End If

Next

If emailFound = False Then
    Exit Function
End If

If Not eMail.Sent Then
    Exit Function
End If

eMail.SaveAs quoteFolderPath & "\" & quoteFolder & "\" & oQuote & ".msg"

'go through email's attachments
'look for *customer* or *cew* and *.xls*
'save file into directory with quote number
'open cew & risk assessment?

Exit Function

errhandler: MsgBox "Error in CopyEmailWkbks function"

End Function

Function PricingSheetRev(preEx As String, revLet As String, oQuote As String, quoteFolderPath As String, _
    quoteFolder As String, templFile As String, modeL As String, desiG As String) As Boolean

Dim wbOrig As Workbook, wbRev As Workbook, filenameOrig As String, filenameEnd As String, filenameRev As String, templOrig As String
Dim modelOrig As String, sameModel As Boolean, pricE As Integer, cosT As Integer, evoRev As Single, evoOrig As Single
Dim partRev As Single, partOrig As Single, qtY As Integer, indexRow As Integer, lastRow As Integer, i As Integer
Dim optionDesc As String, optionQty As Integer, firstRow As Integer

On Error GoTo errhandler

filenameOrig = Dir$(quoteFolderPath & "\" & quoteFolder & "\" & oQuote & preEx & "*.xls*")
'check for preexisting pricing sheet

If filenameOrig = "" Then
    Exit Function
End If

PricingSheetRev = False

On Error Resume Next
Set wbOrig = Workbooks(oQuote & preEx & "*.xls*")
If wbOrig Is Nothing Then
    Set wbOrig = Workbooks.Open(Filename:=quoteFolderPath & "\" & quoteFolder & "\" & oQuote & preEx & "*.xls*", UpdateLinks:=True)
End If
On Error GoTo errhandler

If Not preEx = "-" Then
    filenameEnd = Right(filenameOrig, Len(filenameOrig) - (Len(oQuote) + Len(preEx)))
    'this means there's a rev letter in the filename
Else
    filenameEnd = Right(filenameOrig, Len(filenameOrig) - Len(oQuote))
End If

filenameRev = quoteFolderPath & "\" & quoteFolder & "\" & oQuote & revLet & filenameEnd

wbOrig.SaveAs Filename:=filenameRev
wbOrig.Close savechanges:=False

On Error Resume Next
Set wbRev = Workbooks(filenameRev)
If wbRev Is Nothing Then
    Set wbRev = Workbooks.Open(filenameRev)
End If
On Error GoTo errhandler

wbRev.Activate

If Range("A1").Value > 0 Then
    templOrig = "AM_Pricing"
Else
    templOrig = "Pricing_Template"
End If

evoRev = ThisWorkbook.Sheets(1).Range("G4").Value
partRev = ThisWorkbook.Sheets(1).Range("G5").Value

If Not (templFile = templOrig) Then
    'new wb and old wb are diff templates
    
    msg = "Last revision of this quote was a different format. Redo the quote from scratch?"
    ans = MsgBox(msg, vbYesNo, "Budgetary/Aftermarket vs. New Machine")

    If ans = vbYes Then
        On Error Resume Next
        wbRev.Close savechanges:=False
        Exit Function
    End If
    
End If

templFile = templOrig

If templFile = "Pricing_Template" Then
    'check machine & add to wkbk if nec.
    modelOrig = Range("A7").Value
    sameModel = False
    If InStr(1, modelOrig, modeL) > 0 Then
        sameModel = True
    End If
    
    If sameModel = False Then
        'replace machine
        ThisWorkbook.Activate
        Sheet4.Activate
        pricE = Range("L3").Value
        cosT = Range("M3").Value
        wbRev.Activate
        Range("A7").Value = "Mateer® Model " & modeL & " " & desiG & " Filler"
        Range("F7").Value = pricE
        Range("J7").Value = "Price book"
        Range("L7").Value = cosT
    End If
    
    'check options & add to wkbk if nec.
    ThisWorkbook.Activate
    Sheet4.Activate
    Range("A1").Formula = "=SUM(C:C)"
    
    If Range("A1").Value > 0 Then
    
        'some options
        qtY = Range("A1").Value
        Sheet4.AutoFilterMode = False
        Range("A1").Formula = "=MATCH(1,C:C,0)"
        wbRev.Activate
        indexRow = 0
        Range("B1").Formula = "=MATCH(" & """" & "Options" & """" & ",A:A,FALSE)"
        
        On Error Resume Next
        firstRow = Range("B1").Value
        On Error GoTo errhandler
        
        If firstRow = 0 Then
            
          'add options section and one option (lev cont)
            ThisWorkbook.Activate
            Sheet4.Activate
            
            indexRow = Range("C3").Value 'using this to avoid an extra variable
            Range("C3").Value = 1
            optionDesc = Range("F3").Value
            pricE = Range("G3").Value
            cosT = Range("H3").Value
            optionQty = Range("I3").Value
            Range("C3").Value = indexRow 'putting C3 back to the way it was
            
            wbRev.Activate
            
            If Range("A12").Value > 0 Then
                Range("A11").Select
                Selection.End(xlDown).Select
                indexRow = ActiveCell.Row + 1 'line under tooling
            Else
                indexRow = 12
            End If
            
            Rows(indexRow & ":" & indexRow).Select
            For i = 1 To 4
                Selection.Insert Shift:=xlDown
            Next
            Rows("8:11").Select
            Selection.Copy
            Rows(indexRow & ":" & indexRow + 3).Select
            Selection.PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
                        
            Range("A" & indexRow + 1).Value = "Options"
            Range("A" & indexRow + 2).Value = "Description"
            Range("B" & indexRow + 2).Value = "Price Each"
            Range("C" & indexRow + 2).Value = "Qty"
            Range("D" & indexRow + 2).Value = "Price"
            
            Range("A" & indexRow + 3).Value = optionDesc
            Range("B" & indexRow + 3).Formula = "=IFERROR(SUM(F" & indexRow + 3 & ":H" & indexRow + 3 & ")," & """" & "TBD" & """" & ")"
            Range("C" & indexRow + 3).Value = optionQty
            Range("D" & indexRow + 3).Formula = "=IFERROR(C" & indexRow + 3 & "*B" & indexRow + 3 & "," & """" & "TBD" & """" & ")"
            
            Range("F" & indexRow + 3).Value = pricE
            Range("J" & indexRow + 3).Value = "Price book"
            Range("L" & indexRow + 3).Value = cosT
            Range("M" & indexRow + 3).Formula = "=1-(L" & indexRow + 3 & "/F" & indexRow + 3 & ")"
            firstRow = Range("B1").Value
        End If
        
        firstRow = firstRow + 2 'first option
        indexRow = firstRow 'last option
        If Range("A" & firstRow + 1).Value > 0 Then
            Range("A" & firstRow).Select
            Selection.End(xlDown).Select
            indexRow = ActiveCell.Row
        End If
        indexRow = indexRow + 1
        Range("B1").Clear
    
        ThisWorkbook.Activate
        Sheet4.Activate
    
        For i = 3 To Application.WorksheetFunction.Sum(Range("I:I")) + 50
            Range("J" & i).Formula = "=IFERROR(IF(F" & i & ">" & """" & """" & ",VLOOKUP(F" & i & ",'" & quoteFolderPath & "\" & quoteFolder & "\[" & oQuote & revLet & filenameEnd & _
                "]Mateer Filler'!$A:$A,1,FALSE),0),1)"
            On Error Resume Next
            If Range("C" & i).Value = 1 And Range("J" & i).Value = 1 Then
                On Error GoTo errhandler
                optionDesc = Range("F" & i).Value
                pricE = Range("G" & i).Value
                cosT = Range("H" & i).Value
                optionQty = Range("I" & i).Value
                
                wbRev.Activate
                Rows(indexRow & ":" & indexRow).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                Range("A" & indexRow).Value = optionDesc
                Range("C" & indexRow).Value = optionQty
                Range("F" & indexRow).Value = pricE
                Range("J" & indexRow).Value = "Price book"
                Range("L" & indexRow).Value = cosT
        
                indexRow = indexRow + 1
                ThisWorkbook.Activate
            End If
        Next
        
        On Error GoTo errhandler
        
        Range("J:J").Clear
        indexRow = indexRow - 1

        wbRev.Activate
        Rows(firstRow - 2 & ":" & indexRow).EntireRow.AutoFit

        Range("B" & firstRow).Select
        Selection.AutoFill Destination:=Range("B" & firstRow & ":B" & indexRow)
        
        Range("D" & firstRow).Select
        Selection.AutoFill Destination:=Range("D" & firstRow & ":D" & indexRow)
        
        Range("G" & firstRow & ":H" & firstRow).Select
        Selection.AutoFill Destination:=Range("G" & firstRow & ":H" & indexRow)
        
        Range("M" & firstRow).Select
        Selection.AutoFill Destination:=Range("M" & firstRow & ":M" & indexRow)
    End If
    
    ThisWorkbook.Activate
    Sheet4.Activate
    Range("A1").ClearContents
    Sheet1.Activate
    
    'check EVO & replace if nec.
    wbRev.Activate
    If evoRev < 1 Then
        evoOrig = Range("H6").Value
        If Not evoOrig = evoRev Then
            Range("H6").Value = evoRev
        End If
    Else
        evoOrig = Range("H7").Value
        If Not evoOrig = evoRev Then
            Range("H7").Value = evoRev
        End If
    End If
    
    If partRev < 1 Then
        partOrig = Range("G6").Value
        If Not partOrig = partRev Then
            Range("G6").Value = partRev
        End If
    Else
        partOrig = Range("G7").Value
        If Not partOrig = partRev Then
            Range("G7").Value = partRev
        End If
    End If
    
ElseIf templFile = "AM_Pricing" Then

    'check machine & add to wkbk if nec.
    sameModel = False
    For i = 1 To 100
        If InStr(Range("A" & i).Value, modeL) > 0 Then
            sameModel = True
        End If
    Next
    
    If sameModel = False Then
        'add row for base machine (may result in two machines being listed if one was already present)
        Rows("5:5").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ThisWorkbook.Activate
        Sheet4.Activate
        pricE = Range("L3").Value
        cosT = Range("M3").Value
        wbRev.Activate
        Range("A5").Value = modeL
        Range("B5").Value = pricE
        Range("K5").Value = cosT
        Range("F4:J4").Select
        Selection.AutoFill Destination:=Range("F4:J5")
        
        Range("L4:L4").Select
        Selection.AutoFill Destination:=Range("L4:L5")
        
        Range("I5").Value = 1
    End If
    
    
    'check options & add to wkbk if nec.
    
    wbRev.Activate
    Range("A4").Select
    Selection.End(xlDown).Select
    If ActiveCell.Row < 200 Then
        indexRow = ActiveCell.Row
    Else
        indexRow = 5
    End If
    
    ThisWorkbook.Activate
    Sheet4.Activate
    
    If Application.WorksheetFunction.Sum(Range("C:C")) > 0 Then
        'some options
        
        For i = 3 To Application.WorksheetFunction.Sum(Range("I:I")) + 50
            Range("J" & i).Value = "=IFERROR(VLOOKUP(B" & i & ",'" & quoteFolderPath & "\" & quoteFolder & "\[" & oQuote & revLet & filenameEnd & _
                "]Sheet1'!$A:$A,1,FALSE),1)"
            On Error Resume Next
            If Range("C" & i).Value = 1 And Range("J" & i).Value = 1 Then
                On Error GoTo errhandler
                optionDesc = Range("B" & i).Value
                pricE = Range("G" & i).Value
                cosT = Range("H" & i).Value
                optionQty = Range("I" & i).Value
                
                wbRev.Activate
                Range("A" & indexRow).Value = optionDesc
                Range("B" & indexRow).Value = pricE
                Range("I" & indexRow).Value = optionQty
                Range("K" & indexRow).Value = cosT
        
                indexRow = indexRow + 1
                ThisWorkbook.Activate
            End If
        Next
        
    End If

    Sheet4.Activate
    Range("J:J").Clear
    Range("A1").ClearContents
    Sheet1.Activate
    
    indexRow = indexRow - 1
    wbRev.Activate
    
    Range("F4:H4").Select
    Selection.AutoFill Destination:=Range("F4:H" & indexRow)
    
    Range("J4").Select
    Selection.AutoFill Destination:=Range("J4:J" & indexRow)
    
    Range("L4").Select
    Selection.AutoFill Destination:=Range("L4:L" & indexRow)
    
    'check EVO & replace if nec.
    wbRev.Activate
    If evoRev > 1 Then
        MsgBox "EVO needs to be added manually to make sure it's counted properly."
    Else
        evoOrig = Range("G3").Value
        If evoOrig > 0 Then
            If Not evoOrig = evoRev Then
                Range("G3").Value = evoRev
            End If
        End If
    End If
    
End If

wbRev.Activate
Range("A1").Select
PricingSheetRev = True

Exit Function

errhandler: MsgBox "Error in PricingSheetRev function."
PricingSheetRev = False

End Function

Function QuoteRev(preEx As String, revLet As String, oQuote As String, quoteFolderPath As String, quoteFolder As String, templFile As String) As Boolean

Dim wordApp As Object, docOrig As Object, docRev As Object, filenameOrig As String, filenameEnd As String, filenameRev As String
Dim templOrig As String, pagesOrig As Integer, pricingOrig As String, docRange As Word.Range, dateLine As String

On Error GoTo errhandler

filenameRev = Dir$(quoteFolderPath & "\" & quoteFolder & "\" & oQuote & preEx & "*.doc*")
'check for last rev doc
'used filenameRev to avoid creating another variable

If filenameRev = "" Then
    Exit Function
End If

QuoteRev = False

Set wordApp = GetObject(, "Word.Application")
filenameOrig = Dir$(quoteFolderPath & "\" & quoteFolder & "\" & oQuote & preEx & "*.doc*")
Set docOrig = wordApp.Documents.Open(quoteFolderPath & "\" & quoteFolder & "\" & filenameOrig)
docOrig.Activate

'check templfiles (infer using number of pages)
pagesOrig = docOrig.Range.Information(wdNumberOfPagesInDocument)

If pagesOrig > 12 Then
    templOrig = "Pricing_Template"
Else
    templOrig = "AM_Pricing"
End If

'if different, ask user
If Not (templOrig = templFile) Then
    
    msg = "Last revision of this quote was a different format. Redo the quote from scratch?"
    ans = MsgBox(msg, vbYesNo, "Budgetary/Aftermarket vs. New Machine")

    If ans = vbYes Then
        On Error Resume Next
        docOrig.Close savechanges:=False
        Exit Function
    End If
End If

templFile = templOrig

If preEx = "-" Then
    filenameEnd = Right(filenameOrig, Len(filenameOrig) - (Len(oQuote) + Len(preEx)))
Else
    filenameEnd = Right(filenameOrig, Len(filenameOrig) - (Len(oQuote)))
End If
filenameRev = quoteFolderPath & "\" & quoteFolder & "\" & oQuote & revLet & filenameEnd

docOrig.SaveAs Filename:=filenameRev
docOrig.Close savechanges:=False
Set docRev = wordApp.Documents.Open(filenameRev)

'replace quote number
If templOrig = "Pricing_Template" Then
    For Each docRange In docRev.StoryRanges
        With docRange.Find
            .Text = oQuote & preEx
            .Replacement.Text = oQuote & revLet
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next docRange
    
Else
        'for some reason the code for the
        'replacement used on the other template
        'won't work here (???)
        With docRev.Content.Find
            .Text = oQuote & preEx
            .Replacement.Text = oQuote & revLet
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
End If

'replace date
ThisWorkbook.Activate
Sheet3.Activate
Range("C18").Select
dateLine = ActiveCell.Value

If templOrig = "Pricing_Template" Then
    docRev.Paragraphs(32).Range.Text = dateLine
Else
    docRev.Paragraphs(1).Range.Text = dateLine & vbCrLf
End If

If templOrig = "Pricing_Template" Then
    pricingOrig = Dir$(quotePath & "\" & quoteFolder & "\" & oQuote & "\" & revLet & "*.xls*")
    
    If Not pricingOrig = "" Then
        MsgBox "replace pricing stuff"
        'delete existing app summary
        'delete existing pricing summary
        'copy both from wbrev
        'update table of contents
        
    End If
End If

wordApp.Visible = True
docRev.Activate
QuoteRev = True

Exit Function

errhandler: MsgBox "Error in QuoteRev function."
On Error Resume Next
docRev.Close savechanges:=False
QuoteRev = False

End Function

Function RotaryOptionDisclaimer(modeL As String, newWB As Workbook, templFile As String)

On Error GoTo errhandler

If modeL <> "6600" And modeL <> "6700" Then
    Exit Function
End If

ThisWorkbook.Activate
If Range("D" & Application.WorksheetFunction.Match("*Pitch*Diam*", Range("B:B"), 0)).Value > 36 Then
    Exit Function
End If

If templFile = "AM_Pricing" Then
    Exit Function
End If

'see if both options are offered
Dim matchTest As Integer, columnJack As Boolean, turretJack As Boolean
newWB.Activate
matchTest = 0
columnJack = False
turretJack = False
On Error Resume Next
matchTest = Application.WorksheetFunction.Match("*turret*jack*", Range("A:A"), 0)
If matchTest > 0 Then
    turretJack = True
End If
matchTest = Application.WorksheetFunction.Match("*column*jack*", Range("A:A"), 0)
If matchTest > 0 Then
    columnJack = True
End If

If columnJack = False Or turretJack = False Then
    Exit Function
End If

On Error GoTo errhandler
'disclaimer in newwb
Dim spRow As Integer
spRow = Application.WorksheetFunction.Match("*Spare*Parts*", Range("A:A"), 0)

matchTest = Application.WorksheetFunction.Match("*turret*jack*", Range("A:A"), 0)
Range("A" & matchTest).Value = Range("A" & matchTest).Value & " ***"

matchTest = Application.WorksheetFunction.Match("*column*jack*", Range("A:A"), 0)
Range("A" & matchTest).Value = Range("A" & matchTest).Value & " ***"

Rows(spRow - 1 & ":" & spRow - 1).Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

Range("A" & spRow - 1 & ":D" & spRow - 1).Merge
Range("A" & spRow - 1).Value = "*** If the turret jack and column jack options are" & _
                                "both purchased, a larger size machine is required."

Range("A1").Select
ThisWorkbook.Activate
Sheet1.Activate

Exit Function
errhandler:
MsgBox "Error in RotaryOptionDisclaimer function"
End Function

Function CreateCostForm(quoteNum As String, custName As String, salesRep As String, sN As String, machTyp As String, destPath As String)

Dim wBook As Workbook, wbPath As String, templFile As String, newWbName As String
Dim outlookApp As Object, eMail As MailItem, testInspect As Inspector
Dim numEmails As Integer, nameArray() As String, senderSurname As String, emailContent As String
Dim emailFound As Boolean, emailLines As Variant, i As Integer, startRow As Integer

wbPath = "\\PSACLW02\PROJDATA\ENGLANDT\MATEER\QUOTES\"
templFile = "Cost_Estimate_Form.xlsm"

'see if form is required
msg = "Do you want to create a cost estimate form for this quote?"
ans = MsgBox(msg, vbYesNo, "Cost Estimate Form")

If ans = vbYes Then 'create form
    If Right(destPath, 1) <> "\" Then
        destPath = destPath & "\"
    End If
    newWbName = destPath & "Cost_Estimate-" & quoteNum & ".xlsm"
    
    Set wBook = Workbooks.Open(Filename:=wbPath & templFile, UpdateLinks:=True)
    
    'populate as possible
    wBook.Activate
    wBook.Worksheets(1).Activate
    Range("B1").Value = quoteNum
    Range("B2").Value = custName
    Range("B3").Value = Format(Date, "mm/dd/yy")
    If sN <> "" Then
        Range("B7").Value = sN
    End If
    Range("B8").Value = machTyp
    
    Range("B:B").HorizontalAlignment = xlCenter
    If Len(custName) > 10 Then
        Range("B2").HorizontalAlignment = xlLeft
    End If
    
    'see if email is open
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    On Error GoTo errhandler
    
    If Not outlookApp Is Nothing Then
        
        Set testInspect = outlookApp.ActiveInspector
        
        If Not testInspect Is Nothing Then
            
            numEmails = outlookApp.Inspectors.Count
            
            nameArray() = Split(salesRep)
            salesRep = nameArray(1) 'last name
            emailFound = False
            
            If numEmails > 0 Then
                
                For i = 1 To numEmails
                    
                    Set eMail = outlookApp.Inspectors.Item(i).CurrentItem
                    
                    senderSurname = eMail.SenderName
                    
                    nameArray() = Split(senderSurname, ",")
                    senderSurname = nameArray(0)
                    
                    If eMail.Sent Then
                        If senderSurname = salesRep Then
                            emailContent = eMail.Body
                            If InStr(1, emailContent, custName) > 0 Then
                                emailFound = True
                                Exit For
                            End If
                        End If
                    End If
                
                Next
                
                If emailFound = True And eMail.Sent Then
                    wBook.Activate
                    wBook.Worksheets(2).Activate
                    emailLines = Split(eMail.Body, vbCrLf)
                    startRow = 3
                    For i = 0 To UBound(emailLines)
                        Cells(startRow + 1, 1).Value = emailLines(i)
                        startRow = startRow + 1
                    Next
                    wBook.Worksheets(1).Activate
                End If
                
            End If
            
        End If
        
    End If
    
    'close wbook & save changes
    wBook.SaveAs Filename:=(newWbName)
    Application.Run ("'Cost_Estimate-" & quoteNum & ".xlsm'!CheckForLinks")
    Exit Function
    
Else
    Exit Function
End If

errhandler:
MsgBox "Error in CreateCostForm function"

End Function
