Attribute VB_Name = "Module1"
Option Explicit
Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hWnd As LongLong) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As LongLong
Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As LongLong
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As LongLong) As LongLong
Private Declare PtrSafe Function GetClipboardData Lib "user32.dll" (ByVal wFormat As LongLong) As LongLong
Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As LongLong, ByVal hMem As LongLong) As LongLong
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As LongLong, ByVal dwBytes As LongLong) As LongLong
Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongLong) As LongLong
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As LongLong) As LongLong
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongLong) As LongLong
Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As LongLong, ByVal lpString2 As LongLong) As LongLong

Public Sub SetClipboard(sUniText As String)
    Dim iStrPtr As LongLong
    Dim iLen As LongLong
    Dim iLock As LongLong
    Const GMEM_MOVEABLE As LongLong = &H2
    Const GMEM_ZEROINIT As LongLong = &H40
    Const CF_UNICODETEXT As LongLong = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

Sub Populate_Quote()
'
' Make a new quote (New or Aftermarket)
'
Dim pW As String

pW = InputBox("Enter password to use this function", "Password Required")

If pW = "" Then
    Exit Sub
ElseIf pW <> "T" Then
    MsgBox "Password is incorrect."
    Exit Sub
End If

    Dim typ As String
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Sheet3.Activate
    typ = Range("C5").Value
    
    If Range("G4").Value = True Then
        Call Populate_AM_Quote
        Exit Sub
    End If
        
    If typ = "NM" Then
        Call Populate_NM_Quote
    ElseIf typ = "AM" Then
        Call Populate_AM_Quote
    ElseIf typ = "CP" Then
        Call Populate_AM_Quote
    Else:
        Sheet1.Activate
        MsgBox "Fix the input errors before generating this"
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub Populate_Pricing()
'
' Make a new pricing workbook for a particular quote
'

    Dim wbNew As Workbook, wbOrig As Workbook, wkSht As Worksheet, newWbName As String
    Dim preEx As String, preExInd As String, oQuote As String, partnerProp As Single
    Dim typ As String, companyName As String, company1 As String, qLength As String, partnerDir As Integer
    Dim yeaR As Integer, templFile As String, rev As String, desiG As String
    Dim modeL As String, quotE As String, budG As Boolean, revLet As String
    Dim origText As String, newText As String, steelType As String, namePos As Integer
    Dim templPath As String, quoteFolderPath As String, newRev As Boolean
    Dim pW As String, sN As String, salesRep As String, i As Integer

    pW = InputBox("Enter password to use this function", "Password Required")
    
    If pW = "" Then
        Exit Sub
    ElseIf pW <> "T" Then
        MsgBox "Password is incorrect."
        Exit Sub
    End If
    
    On Error GoTo errhandler

    Sheet3.Activate
    Sheet3.Unprotect
    Set wkSht = Sheet3

    quotE = wkSht.Range("C4").Value
    typ = wkSht.Range("C5").Value
    modeL = wkSht.Range("C6").Value
    salesRep = wkSht.Range("C7").Value
    If modeL = "" Then
        modeL = "NO MODEL"
    End If

    company1 = wkSht.Range("C17").Value
    
    namePos = InStr("abc" & company1, "Bw")
    
    If namePos > 0 Then
    
        company1 = Application.WorksheetFunction.Substitute(Range("C8").Value, " ", "")
    End If

    desiG = wkSht.Range("C19").Value

    wkSht.Range("A1").Formula = "=LEFT(C4,12)"

    oQuote = wkSht.Range("A1").Value
    wkSht.Range("A1").Formula = "=LEN(C4)"
    qLength = wkSht.Range("A1").Value
    Range("A1").Formula = "=LEFT(C4,2)"
    yeaR = "20" & Range("A1").Value
    wkSht.Range("A1").Clear
    partnerProp = Range("G16").Value
    partnerDir = Range("G17").Value
    newRev = False

    'path designations
   ''''''''''''''''''
    templPath = "K:\EnglandT\MATEER\QUOTES\Quote Templates\"
    quoteFolderPath = "T:\Quotes\Mateer\" & yeaR & " Quotes"
   ''''''''''''''''''
   'path designations
    
    If Sheet1.Range("I1").Value <> "READY" Then
        Sheet1.Activate
    
        MsgBox "Fix the input errors before generating this"
        Exit Sub
        
    End If
    
    budG = False
    templFile = "Pricing_Template"
    
    If typ = "AM" Then
        templFile = "AM_Pricing"
    ElseIf typ = "CP" Then
        templFile = "AM_Pricing"
    ElseIf Range("G4").Value = True Then
        budG = True
        templFile = "AM_Pricing"
    End If

    preEx = Dir$(templPath & templFile & ".xlsx")
    
    If preEx = "" Then
        
        MsgBox "Must create template..." & vbCrLf & templFile & ".xlsx" & vbCrLf & vbCrLf & "This file must be in folder..." & vbCrLf & templPath
        Sheet1.Activate
        Range("A1").Select
        Exit Sub
    
    End If

    On Error Resume Next
    Set wbNew = Workbooks(templFile & ".xlsx")
    If wbNew Is Nothing Then
        Set wbNew = Workbooks.Open(Filename:=templPath & templFile & ".xlsx", UpdateLinks:=True)
    End If
    On Error GoTo errhandler

    preEx = Dir$(quoteFolderPath & "*", vbDirectory)
    'check for umbrella quote folder
    
    If preEx = "" Then
        'create umbrella quote folder
        MkDir Path:=quoteFolderPath
    End If
    
    preExInd = Dir$(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
    'check for indiv. quote folder
    
    If preExInd = "" Then
        'create indiv. quote folder
        If modeL = "NO MODEL" Then
        MkDir Path:=quoteFolderPath & "\" & oQuote & "-" & typ & "-" & company1
        Else: MkDir Path:=quoteFolderPath & "\" & oQuote & "-" & typ & "-" & company1 & "-" & modeL
        End If
        preExInd = Dir(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
    End If
    
    Call CopyEmailWkbks(oQuote, yeaR)
    
    preEx = Dir$(quoteFolderPath & "\" & preExInd & "\" & oQuote & "*" & ".xlsx*")
    'check for wb

    If preEx = "" Then
    
        If qLength > 12 Then
        
            MsgBox "There is no original version of this quote. Ensure the quote number is correct."
        
        End If
    
    Else:
        
        Do While preEx <> ""
        
            preExInd = preEx
            'using preExInd to avoid another variable
            preEx = Dir$()
        
        Loop
        
        wkSht.Range("A1").Value = preExInd
        wkSht.Range("A2").Formula = "=LEFT(RIGHT(A1,LEN(A1)-12),1)"
        preEx = wkSht.Range("A2").Value
        
        If preEx = "-" Then
            preEx = "@"
            '@ is the character before "A" in the ASCII code library
            'must be changed for program to run on Macintosh
        End If
        
        wkSht.Range("A1").Formula = "=RIGHT(C4,1)"
        revLet = wkSht.Range("A1").Value
        wkSht.Range("A2").Value = preEx
        wkSht.Range("A3").Formula = "=CODE(A1)-CODE(A2)"
        
        preExInd = Dir(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
        'setting preExInd so file is named correctly after If block

        If wkSht.Range("A3").Value <> 1 Then
        
            If preEx = "@" Then
                preEx = "original"
            End If
            
            preExInd = Dir$(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
            
            If revLet = "0" Then
                revLet = "the initial one"
            End If
            
            MsgBox "Different revision letter (not " & revLet & ") of this workbook is needed for this quote number." & vbCrLf & vbCrLf & _
                "Most up-to-date existing revision: " & preEx & vbCrLf & vbCrLf & "Stored in folder: " & quoteFolderPath & "\" & preExInd
            
            On Error Resume Next
            Sheet3.Activate
            Range("A1:A3").Clear
            Range("A1").Select
            ActiveSheet.Protect
            Sheet1.Activate
            Range("A1").Select
            Application.ScreenUpdating = True
            
            Exit Sub
        Else
            If preEx = "@" Then
                preEx = "-"
            End If
            newRev = True
        End If
        
        wkSht.Range("A1:A3").Clear
    
    End If

    newWbName = quoteFolderPath & "\" & preExInd & "\" & quotE & "-" & company1 & "-Pricing.xlsx"
    quoteFolderPath = quoteFolderPath & "\" & preExInd & "\" 'used for cost estimate form
    
    If newRev = True Then
        newRev = PricingSheetRev(preEx, revLet, oQuote, quoteFolderPath, preExInd, templFile, modeL, desiG)
        If newRev = True Then
            On Error Resume Next
            wbNew.Close savechanges:=False
            Exit Sub
        End If
    End If
    
    wbNew.Activate
    wbNew.Worksheets(1).Activate
    
    On Error Resume Next
    For i = 4 To 50
    
        origText = "<<template_" & wkSht.Range("B" & i).Value & ">>"
        newText = wkSht.Range("C" & i).Value
        
        Cells.Replace What:=origText, Replacement:=newText, LookAt _
            :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
            
    Next i
    On Error GoTo errhandler
    
    If partnerDir > 0 Then
        If templFile = "Pricing_Template" Then
            Range("G7").Formula = "=(F7/(1-$G$6)-F7)+" & partnerDir
        Else
            Range("H4").Formula = "=IFERROR(VALUE(" & partnerDir & ")+$F$4/(1-H3)-$F$4,$F$4)"
        End If
    End If
    ' ^v this and this are to do the EVO for a partner
    If partnerProp > 0 Then
        If templFile = "Pricing_Template" Then
            Range("G6").Value = partnerProp
        Else
            Range("H3").Value = partnerProp
        End If
    End If

    Application.ScreenUpdating = False
    
    ThisWorkbook.Activate
    Sheet1.Activate
    companyName = Range("D8").Value 'for cost estimate function
    
    steelType = "304"
    If Application.WorksheetFunction.VLookup("316 Product Contact Parts", Range("C:E"), 3, False) Then
        steelType = "316"
    End If
    
    If budG = True Then
        wbNew.Activate
        Range("A4").Value = modeL
        Range("A5").Value = "Standard FF tooling (" & steelType & ")"
        Range("A6").Value = "Standard NFF tooling (" & steelType & ")"
    End If
    Application.ScreenUpdating = False
    
    Call CopyBudgTooling(wbNew, templFile)
    'puts pricing for FF tooling
    
    ThisWorkbook.Activate
    Sheet1.Activate
    
    If Range("M5").Value > 0 Then
        Call ToolingPricing(wbNew)
        'put tooling on sheet 2 of newwb
    End If
    Application.ScreenUpdating = False

    Call PricingOptions(ThisWorkbook, wbNew, templFile)
    'populate options and their prices/costs
    Application.ScreenUpdating = False

    ThisWorkbook.Activate
    Sheet1.Activate
    
    Call RotaryOptionDisclaimer(modeL, wbNew, templFile)
    
    If Range("M5").Value > 0 Then
        Call CopyToolingPricing(wbNew, templFile, steelType)
        'put tooling & its pricing on sheet 1 of newwb
        'comes after options because the options need to
        'always start on the same row
    End If
    Application.ScreenUpdating = False
    
    If templFile = "AM_Pricing" Then
        wbNew.Activate
        Columns("A:L").EntireColumn.AutoFit
        Range("A1").Select
    End If
    
    ThisWorkbook.Activate
    Sheet1.Activate
    
    If Application.WorksheetFunction.VLookup("316 Product Contact Parts", Range("C:E"), 3, False) Then
        If templFile = "Pricing_Template" Then
            Call Option316(wbNew)
        End If
    End If
    Application.ScreenUpdating = False

    wbNew.SaveAs Filename:=(newWbName), _
        FileFormat:=xlOpenXMLWorkbook
        
    ThisWorkbook.Activate
    Sheet1.Activate
    
    On Error Resume Next
        sN = Application.WorksheetFunction.VLookup("Serial*", Range("B:D"), 3, 0)
    On Error GoTo errhandler
    
    Call CreateCostForm(oQuote, companyName, salesRep, sN, modeL, quoteFolderPath)
    
    ThisWorkbook.Activate
    Sheet3.Protect
    Sheet1.Activate
    Sheet1.Range("A1").Select
    Sheet1.Protect
    Application.ScreenUpdating = True
    wbNew.Activate
    wbNew.Worksheets(1).Activate
        
    Exit Sub
    
errhandler:     MsgBox "unspecified error"

End Sub


Sub Populate_AM_Quote()
'
' Populate_Afetermarket_Quote Macro
'

    Dim wApp As Object, check As Boolean
    Dim wDoc As Object
    Dim newDocName As String, preEx As String, preExInd As String, seriaL As String
    Dim wRange As Word.Range, para As Paragraph, newRev As Boolean, namePos As Integer
    Dim numVars As Integer, yeaR As Integer, qLength As Integer, quoteFolder As String
    Dim quotE As String, company As String, modeL As String, oQuote As String
    Dim txt As String, tempStr As String, typ As String, errFree As String, i As Integer
    Dim templPath As String, quoteFolderPath As String, revLet As String, company1 As String
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Sheet3.Activate
    ActiveSheet.Unprotect

    errFree = Sheet1.Range("K1").Value
    If errFree <> "READY" Then
        ActiveSheet.Protect
        Sheet1.Activate
        Range("A1").Select
        Application.ScreenUpdating = True
        MsgBox "Fix the input errors before generating this"
        Exit Sub
    End If

    Set wApp = GetObject(, "Word.Application") 'open word
    
    If Err.Number = 429 Then
        'no instance of word is open
        Err.Clear
        Set wApp = CreateObject("Word.Application")
    End If
    
    On Error GoTo errhandler

    quotE = Range("C4").Value
    typ = Range("C5").Value
    
    If Range("B13").Value = "VFFS" Then
        seriaL = ""
    Else: seriaL = Range("C13").Value
    End If
    
    company = Range("C8").Value
    company1 = Range("C17").Value
    modeL = Application.WorksheetFunction.Substitute(Range("C6"), " ", "")

    numVars = 50
    
    Range("A1").Formula = "=LEFT(C4,12)"
    oQuote = Range("A1").Value
    Range("A1").Formula = "=LEN(C4)"
    qLength = Range("A1").Value
    Range("A1").Formula = "=LEFT(C4,2)"
    yeaR = "20" & Range("A1").Value
    Range("A1").Select
    Range("A1").Clear
    newRev = False
    
    namePos = InStr("abc" & company1, "Bw")
    
    If namePos > 0 Then
    
        company1 = Application.WorksheetFunction.Substitute(Range("C8").Value, " ", "")
    End If
    
   'path designations
   ''''''''''''''''''
   templPath = "K:\EnglandT\MATEER\QUOTES\Quote Templates"
   quoteFolderPath = "T:\Quotes\Mateer\" & yeaR & " Quotes"
   ''''''''''''''''''
   'path designations
 
    wApp.Visible = True

    preEx = Dir$(quoteFolderPath & "*", vbDirectory)
    'check for umbrella quote folder
    
    If preEx = "" Then
        'create umbrella quote folder
        MkDir Path:=quoteFolderPath
    End If
    
    preExInd = Dir$(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
    'check for indiv. quote folder
    
    If preExInd = "" Then
        'create indiv. quote folder
        If seriaL > "" Then
            If seriaL <> "ERROR" Then
                MkDir Path:=quoteFolderPath & "\" & oQuote & "-" & typ & "-" & company1 & "-" & modeL & "-" & seriaL
            Else: MkDir Path:=quoteFolderPath & "\" & oQuote & "-" & typ & "-" & company1 & "-" & modeL
            End If
        Else: MkDir Path:=quoteFolderPath & "\" & oQuote & "-" & typ & "-" & company1 & "-" & modeL
        End If
        preExInd = Dir(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
    End If
    
    preEx = Dir$(quoteFolderPath & "\" & preExInd & "\" & oQuote & "*" & ".doc*")
    'check for doc
    
    If preEx = "" Then
        'no doc yet
    
        If qLength > 12 Then
    
            MsgBox "There is no original version of this quote. Ensure the quote number is correct."
    
        End If
    
    Else:
        'doc exists
        
        Do While preEx <> ""
        
            preExInd = preEx
            'using preExInd to avoid another variable
            
            preEx = Dir$()
        
        Loop
        
        Range("A1").Value = preExInd
        Range("A2").Formula = "=LEFT(RIGHT(A1,LEN(A1)-12),1)"
        preEx = Range("A2").Value 'latest rev letter
        Range("A2").Clear
        
        'check if intended revision is next in line
        Range("A1").Formula = "=RIGHT(C4,1)"
        revLet = Range("A1").Value 'intended rev
        
        If preEx = "-" Then
            preEx = "@"
        End If
        
        Range("A1").Formula = "=CODE(" & """" & revLet & """" & ")-CODE(" & """" & preEx & """" & ")"
        
        If preEx = "@" Then
            preEx = "original"
        End If
        
        If Range("A1").Value < 1 Then
            
            preExInd = Dir$(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
            
            MsgBox "A different revision letter (not " & revLet & ") of this quote number is needed." & vbCrLf & vbCrLf & _
                "Most up-to-date existing revision: " & preEx & vbCrLf & vbCrLf & "Stored in folder: " & preExInd
            
            Range("A1:A2").Clear
            ActiveSheet.Protect
            Sheet1.Activate
            Sheet1.Range("A1").Select
            Application.ScreenUpdating = True
            
            wDoc.Close savechanges:=False
            Exit Sub
            
        ElseIf Range("A1").Value > 1 Then
            MsgBox "Latest rev is """ & preEx & """. You are generating rev """ & revLet & """ right now."
        
        Else
            newRev = True
        
        End If
    
    End If
    
    Range("A1").Clear
    
    If newRev = True Then
    'new rev of existing quote
        quoteFolder = Dir$(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
        newRev = QuoteRev(preEx, revLet, oQuote, quoteFolderPath, quoteFolder, "AM_Pricing")
        'If newRev = True Then
            'use existing quote instead of remaking
            On Error Resume Next
            wDoc.Close savechanges:=False
            Exit Sub
        'End If
    End If
    
    preExInd = Dir$(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
    newDocName = quoteFolderPath & "\" & preExInd & "\" & quotE & "-" & typ & "-" & company1 & "-" & modeL & ".docx"
    
    If Dir(templPath & "\Aftermarket_Quote_Template.dotx") = "" Then
        MsgBox "Must create template..." & vbCrLf & "Aftermarket_Quote_Template.dotx" & vbCrLf & vbCrLf & "This file must be in folder..." & vbCrLf & templPath
        ActiveSheet.Protect
        Sheet1.Activate
        Sheet1.Range("A1").Select
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    If Range("G4").Value = True Then
    
        'check for budgetary template
        If Dir(templPath & "\Budgetary_Quote_Template.dotx") = "" Then
            MsgBox "Must create template..." & vbCrLf & "Budgetary_Quote_Template.dotx" & vbCrLf & vbCrLf & "This file must be in folder..." & vbCrLf & templPath
            ActiveSheet.Protect
            Sheet1.Activate
            Sheet1.Range("A1").Select
            Application.ScreenUpdating = True
            Exit Sub
        Else
            Set wDoc = wApp.Documents.Add(template:=templPath & "\Budgetary_Quote_Template.dotx", NewTemplate:=False, DocumentType:=0)
        End If
        
    Else
        Set wDoc = wApp.Documents.Add(template:=templPath & "\Aftermarket_Quote_Template.dotx", NewTemplate:=False, DocumentType:=0)
    
    End If

    For i = 4 To numVars
    
        With wDoc.Range.Find
        
            .Text = "<<template_" & Range("B" & i).Value & ">>"
            .Replacement.Text = Range("C" & i).Value
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll

        End With
    Next i
    MsgBox newDocName
    wDoc.SaveAs Filename:=(newDocName), _
        FileFormat:=wdFormatXMLDocument, AddtoRecentFiles:=True

    Range("A1:A2").Clear
    ActiveSheet.Protect
    Sheet1.Activate
    Sheet1.Range("A1").Select
    Application.ScreenUpdating = True
    wApp.Activate
    
    Exit Sub
    
errhandler:     MsgBox "unspecified error"

End Sub


Sub Populate_NM_Quote()
Attribute Populate_NM_Quote.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Populate_New_Quote Macro
'

'
    Dim wApp As Object, wDoc As Object, check As Boolean, newRev As Boolean
    Dim newDocName As String, preEx As String, preExInd As String, namePos As Integer
    Dim wRange As Word.Range, para As Paragraph, length As Integer, quoteFolder As String
    Dim numVars As Integer, yeaR As Integer, qLength As Integer, model1 As String
    Dim quotE As String, company1 As String, modeL As String, oQuote As String
    Dim txt As String, tempStr As String, typ As String, errFree As String, txtRange As Range
    Dim cc As ContentControl, shpPic As Word.InlineShape, imageFile As String
    Dim templPath As String, quoteFolderPath As String, revLet As String, i As Integer

    On Error Resume Next
    
    Application.ScreenUpdating = False
    Sheet3.Activate
    ActiveSheet.Unprotect
    
    errFree = Sheet1.Range("K1").Value
    If errFree <> "READY" Then
        ActiveSheet.Protect
        Sheet1.Activate
        Range("A1").Select
        Application.ScreenUpdating = True
        MsgBox "Fix the input errors before generating this"
        Exit Sub
    End If
    
    Set wApp = GetObject(, "Word.Application") 'open word
    
    If Err.Number = 429 Then
        'no instance of word is open
        Err.Clear
        Set wApp = CreateObject("Word.Application")
    End If
    
    On Error GoTo errhandler
    
    quotE = Range("C4").Value
    typ = Range("C5").Value
    modeL = Range("C6").Value
    model1 = Application.WorksheetFunction.Substitute(modeL, " ", "")
    company1 = Range("C17").Value
    
    Range("B4").Select
    Selection.End(xlDown).Select
    numVars = ActiveCell.Row

    Range("A1").Formula = "=LEFT(C4,12)"
    oQuote = Range("A1").Value
    Range("A1").Formula = "=LEN(C4)"
    qLength = Range("A1").Value
    Range("A1").Formula = "=LEFT(C4,2)"
    yeaR = "20" & Range("A1").Value
    Range("A1").Select
    Range("A1").Clear
    
    namePos = InStr("abc" & company1, "Bw")
    
    If namePos > 0 Then
    
        company1 = Application.WorksheetFunction.Substitute(Range("C8").Value, " ", "")
    End If
    
   'path designations
   ''''''''''''''''''
   templPath = "K:\EnglandT\MATEER\QUOTES\Quote Templates"
   quoteFolderPath = "T:\Quotes\Mateer\" & yeaR & " Quotes"
   ''''''''''''''''''
   'path designations
   
    preEx = Dir$(quoteFolderPath & "*", vbDirectory)
    'check for umbrella quote folder (current year)
    
    If preEx = "" Then
        'create umbrella quote folder
        MkDir Path:=quoteFolderPath
    End If
    
    preExInd = Dir$(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
    'check for indiv. quote folder
    
    If preExInd = "" Then
        'create indiv. quote folder
        MkDir Path:=quoteFolderPath & "\" & oQuote & "-" & typ & "-" & company1 & "-" & model1
        preExInd = Dir(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
    End If
    
    preEx = Dir$(quoteFolderPath & "\" & preExInd & "\" & oQuote & "*" & ".doc*")
    'check for doc
    
    If preEx = "" Then
    
        If qLength > 12 Then
    
            MsgBox "There is no original version of this quote, but you are creating a version with a revision letter."
    
        End If
    
    Else:
        
        Do While preEx <> ""
        
            preExInd = preEx
            'using preExInd to avoid another variable
            
            preEx = Dir$()
        
        Loop
        
        Range("A1").Value = preExInd
        Range("A2").Formula = "=LEFT(RIGHT(A1,LEN(A1)-12),1)"
        preEx = Range("A2").Value 'latest rev letter
        Range("A2").Clear
        
        'check if intended revision is next in line
        Range("A1").Formula = "=RIGHT(C4,1)"
        revLet = Range("A1").Value 'intended rev
        
        If preEx = "-" Then
            preEx = "@"
            '@ is the character before "A" in the ASCII code library
            'must be changed for program to run on Macintosh
        End If
        
        Range("A1").Formula = "=CODE(" & """" & revLet & """" & ")-CODE(" & """" & preEx & """" & ")"
        
        If preEx = "@" Then
            preEx = "original"
        End If
        
        preExInd = Dir$(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
        
        If Range("A1").Value < 1 Then
        'intended rev not next rev
            
            MsgBox "A different revision letter (not " & revLet & ") of this quote number is needed." & vbCrLf & vbCrLf & _
                "Most up-to-date existing revision: " & preEx & vbCrLf & vbCrLf & "Stored in folder: " & preExInd
            
            Range("A1:A2").Clear
            Range("A1").Select
            Application.ScreenUpdating = True
            ActiveSheet.Protect
            
            wDoc.Close savechanges:=False
            Exit Sub
            
        ElseIf Range("A1").Value > 1 Then
            MsgBox "Latest rev is """ & preEx & """. You are generating rev """ & revLet & """ right now."
            
        Else
            newRev = True
        End If
    
    End If
    
    Range("A1").Clear
    
    preExInd = Dir$(quoteFolderPath & "\" & oQuote & "*", vbDirectory)
    newDocName = quoteFolderPath & "\" & preExInd & "\" & quotE & "-" & typ & "-" & company1 & "-" & model1 & ".docx"

    If Dir(templPath & "\New_Quote_Template.docm") = "" Then
        MsgBox "Must create template..." & vbCrLf & "New_Quote_Template.docm" & vbCrLf & vbCrLf & "This file must be in folder..." & vbCrLf & templPath
        Range("A1").Select
        Application.ScreenUpdating = True
        ActiveSheet.Protect
        Exit Sub
    End If
    
    On Error Resume Next
    'set correct photo based on model
    Range("C1").Formula = "=LEFT(C6,2)"
    imageFile = Dir(templPath & "\Images\" & Range("C1").Value & "*.jpg")
    
    If imageFile > "" Then
        imageFile = templPath & "\Images\" & imageFile
        
        'belt driven / direct driven
        Range("C2").Formula = "=RIGHT(LEFT(C6,5),1)"
        
        If Range("C2").Value = "D" Then
            'check for Direct servo photo
            tempStr = Dir(templPath & "\Images\" & Range("C1").Value & "*D*.jpg")
            If tempStr > "" Then
                imageFile = templPath & "\Images\" & tempStr
            End If
        ElseIf Range("C2").Value = "B" Then
            'check for belt servo photo
            tempStr = Dir(templPath & "\Images\" & Range("C1").Value & "*B*.jpg")
            If tempStr > "" Then
                imageFile = templPath & "\Images\" & tempStr
            End If
        End If
        
        Range("C1").Clear
        Range("C2").Clear
    Else
        MsgBox "Error inserting image of the machine"
    End If
    
    Set wDoc = wApp.Documents.Add(template:=templPath & "\New_Quote_Template.docm", NewTemplate:=False, DocumentType:=0)
    wApp.Visible = True

    wDoc.Bookmarks("ModelPic").Range _
        .InlineShapes.AddPicture Filename:=imageFile

    For Each wRange In wDoc.StoryRanges
        
        For i = 4 To numVars
        
            With wRange.Find
            
                .Text = "<<template_" & Range("B" & i).Value & ">>"
                If Len(Range("C" & i).Value) < 250 Then
                    .Replacement.Text = Range("C" & i).Value
                Else
                    .Replacement.Text = "^c"                'workaround to overcome 255 char limit for Find/Replace
                    Call SetClipboard(Range("C" & i).Value) '^c as the Replace term gives the contents of the clipboard
'                    Dim dataObj As New DataObject          'which can be any length
'                    With dataObj
'                        .SetText Sheet3.Range("C" & i).Text
'                        .PutInClipboard
'                    End With
                End If
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
                
                If i > 200 Then
                    i = numVars
                End If

            End With
        Next i
        
    Next wRange
    
    
    Call LinkAppSummary(quotE, oQuote, yeaR, wDoc)
    
    Call LinkPricing(quotE, oQuote, yeaR, wDoc)
    
    For Each para In wDoc.Paragraphs
        txt = para.Range.Text
        tempStr = LCase(txt)
        check = InStr(tempStr, "template")

        If check = True Then
            para.Range.Delete
        End If

        length = Len(tempStr)

        If length = 0 Then
            para.Range.Delete
        End If
        
    Next
    
    For Each wRange In wDoc.StoryRanges

        With wRange.Find

                .Text = "^p^p^p"
                .Replacement.Text = "^p^p"
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll

        End With

    Next wRange
    
    For Each para In wDoc.Paragraphs
        txt = para.Range.Text
        tempStr = LCase(txt)
        length = Len(tempStr)

        If length = 0 Then
            para.Range.Delete
        End If
        
        Set txtRange = para.Range
        If txtRange.End(xlToRight) = txtRange.Start + 1 Then
            txtRange.Collapse wdCollapseStart
            txtRange.Delete (wdCharacter - 1)
        End If

    Next
    
    For Each wRange In wDoc.StoryRanges

        With wRange.Find

                .Text = "^p^p^p"
                .Replacement.Text = "^p^p"
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll

        End With

    Next wRange
    
    On Error Resume Next
        'run bullets macro
        wApp.Run "Bullets"
        'if error (e.g. macro was removed from template) shouldn't be an issue
    On Error GoTo errhandler
    
    If wDoc.TablesOfContents.Count = 1 Then
        wDoc.TablesOfContents(1).Update
    Else: MsgBox "WARNING: Table of Contents issue"
    End If
    
    wDoc.SaveAs Filename:=(newDocName), _
        FileFormat:=wdFormatXMLDocument, AddtoRecentFiles:=True

    Range("A1").Select
    Range("A1").Clear
    ActiveSheet.Protect
    Sheet1.Activate
    Range("A1").Select
    Application.ScreenUpdating = True
    
    wApp.Activate
    
    Exit Sub
    
errhandler:     MsgBox "unspecified error"

End Sub

Private Function LinkPricing(quotE As String, oQuote As String, yeaR As Integer, wDoc As Object)
'
' Pastes pricing sheet into word doc
'

Dim testFol As String, testFil As String
Dim wBook As Workbook, pricinG As Range
Dim rangeEnd As Integer
Dim quoteFolderPath As String

Application.DisplayAlerts = False

quoteFolderPath = "T:\Quotes\Mateer\" & yeaR & " Quotes"
'path designation for quote folder

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
        wBook.Worksheets(1).Activate
        
        If Range("A1").Value > 0 Then
            wBook.Close savechanges:=False
            Exit Function
        Else
            Range("A200").Select
            Selection.End(xlUp).Select
            rangeEnd = Selection.Row
            Range("A5:D" & rangeEnd).Copy
    
            wDoc.Bookmarks("Pricing").Range.PasteExcelTable LinkedToExcel:=False, _
                WordFormatting:=False, RTF:=False
                
            wBook.Close savechanges:=False
        End If
    End If
    
End If

Application.DisplayAlerts = True

Exit Function

errhandler: MsgBox "Error in LinkPricing function."

End Function

Public Function Options(machine1 As String, machine2 As String, heads As Integer)

Dim Row As Integer
Dim phrase As String

Application.ScreenUpdating = False

ThisWorkbook.Activate
Sheet1.Activate
Sheet1.Unprotect
Range("G9").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Clear
Range("C20:D100").Select
Selection.UnMerge
Selection.ClearContents

Sheet4.Activate
Sheet4.Unprotect
Sheet4.AutoFilterMode = False
Sheet4.Range("$D:$D").AutoFilter Field:=1, Criteria1:="<>0"
Range("B3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheet1.Activate
Range("C20").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

Sheet4.Activate
Range("C1").Value = machine1
'machine1 is either rotary or non-rotary
Range("B1").Formula = "=match(C1,A:A,0)"
Row = Range("B1").Value

Range("B" & Row).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheet1.Activate
Range("G9").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

If machine2 > "" Then
    
    'means non-rotary... machine2 is semiautomatic or automatic
    Sheet4.Activate
    Range("C1").Value = machine2
    Range("B1").Formula = "=match(C1,A:A,0)"
    Row = Range("B1").Value

    Range("B" & Row).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    Sheet1.Activate
    Range("G9").Select
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=True, Transpose:=False

    'options related to the number of heads
    Sheet4.Activate
    
    phrase = "Single Head"
    
    If machine2 = "Semiautomatic" Then
        If heads = 2 Then
            phrase = "Twin Head"
        End If
    End If
    
    Range("B1").Formula = "=match(" & """" & phrase & """" & ",A:A,0)"
    Row = Range("B1").Value
    
    Range("B" & Row).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Sheet1.Activate
    Range("G9").Select
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Else
    Sheet4.Activate
    Range("B1").Formula = "=match(" & """" & "Twin Head" & """" & ",A:A,0)"
    Row = Range("B1").Value
    
    Range("B" & Row).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Sheet1.Activate
    Range("G9").Select
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
End If

Sheet4.Range("B1:C1").Clear
Sheet4.ShowAllData
Sheet4.Range("A1").Select
Application.CutCopyMode = False

End Function

Public Function PricingOptions(wbOld, wbNew, templFile)

Dim qtY As Integer, indeX As Integer, amIndex As Integer, clasS As Integer, shiP As Integer, pricE As Long, cosT As Long
Dim matL As Single, eHrs As Single, aHrs As Single, pasteRow As Integer, eRate As Single, aRate As Single

Application.ScreenUpdating = False

ThisWorkbook.Activate
Sheet4.Activate

On Error GoTo errhandler

ActiveSheet.Unprotect
Range("A1").Formula = "=SUM(C:C)"
qtY = Range("A1").Value 'O3 keeps getting cleared?
pricE = Range("P3").Value
cosT = Range("Q3").Value
matL = Range("R3").Value
eHrs = Range("S3").Value
eRate = Range("S1").Value
aHrs = Range("T3").Value
aRate = Range("T1").Value
clasS = Sheet3.Range("C21").Value
shiP = 1100

If qtY > 0 Then
    Sheet4.Activate
    Sheet4.AutoFilterMode = False
    ActiveSheet.Range("$F:$F").AutoFilter Field:=1, Criteria1:="<>"
    
    Range("A1").Formula = "=MATCH(1,C:C,0)"
    indeX = Range("A1").Value
    Range("A1").ClearContents
Else
    indeX = 1
End If

If templFile = "Pricing_Template" Then

    If clasS = 2 Then
        shiP = 3500
        'shipping price for rotary
    ElseIf clasS = 1 Then
        shiP = 2500
        'shipping price automatics
    End If

    If qtY = 1 Then
        
        wbNew.Activate
        
        Range("B" & Application.WorksheetFunction.Match("*Skidding*", Range("A:A"), 0) + 2).Value = shiP
        
        If clasS <> 2 Then 'not a rotary -> delete rotary options
            Rows(Application.WorksheetFunction.Match("Zepf*", Range("A:A"), 0) & ":" & _
            Application.WorksheetFunction.Match("Zepf*", Range("A:A"), 0) + 1).Delete
        End If
        
        qtY = qtY + Application.WorksheetFunction.Match("Options", Range("A:A"), 0) + 1
        
        wbOld.Activate
        Range("F" & indeX).Select
        Selection.Copy
       
        
        wbNew.Activate
        pasteRow = Application.WorksheetFunction.Match("Options", Range("A:A"), 0) + 2
        Range("A" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

        wbOld.Activate
        Range("G" & indeX).Select
        Selection.Copy
        
        wbNew.Activate
        Range("F" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Range("P" & pasteRow).Formula = "=L" & pasteRow & "+$M$1*M" & pasteRow & "+$N$1*N" & pasteRow
        
        wbOld.Activate
        Range("I" & indeX).Select
        Selection.Copy
        
        wbNew.Activate
        Range("C" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        wbOld.Activate
        Range("K" & indeX).Select
        Selection.Copy
        
        wbNew.Activate
        Range("L" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        wbOld.Activate
        Range("L" & indeX).Select
        Selection.Copy
        
        wbNew.Activate
        Range("M" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        wbOld.Activate
        Range("M" & indeX).Select
        Selection.Copy
        
        wbNew.Activate
        Range("N" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Range("J" & pasteRow).Value = "Price book"
        Range("F7").Value = pricE
        Range("L7").Value = matL
        Range("M7").Value = eHrs
        Range("N7").Value = aHrs
        Range("M1").Value = eRate
        Range("N1").Value = aRate
        Range("P7").Formula = "=L7+$M$1*M7+$N$1*N7"

    ElseIf qtY > 1 Then
        
        wbNew.Activate
        
        If clasS <> 2 Then 'not a rotary -> delete rotary options
            Rows(Application.WorksheetFunction.Match("Zepf*", Range("A:A"), 0) & ":" & _
            Application.WorksheetFunction.Match("Zepf*", Range("A:A"), 0) + 1).Delete
        End If
        
        qtY = qtY + Application.WorksheetFunction.Match("Options", Range("A:A"), 0) + 1
        
        Range("B" & Application.WorksheetFunction.Match("*Skidding*", Range("A:A"), 0) + 2).Value = shiP
        
        pasteRow = Application.WorksheetFunction.Match("Options", Range("A:A"), 0) + 2
    
        Rows(pasteRow + 1 & ":" & qtY).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        wbOld.Activate
        Sheet4.Range("F" & indeX).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy

        wbNew.Activate
        Range("A" & Application.WorksheetFunction.Match("Options", Range("A:A"), 0) + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

        wbOld.Activate
        Sheet4.Range("G" & indeX).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        
        wbNew.Activate
        Range("F" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Range("P" & pasteRow).Formula = "=L" & pasteRow & "+$M$1*M" & pasteRow & "+$N$1*N" & pasteRow
        
        wbOld.Activate
        Sheet4.Range("I" & indeX).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        
        wbNew.Activate
        Range("C" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        wbOld.Activate
        Sheet4.Range("K" & indeX).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        
        wbNew.Activate
        Range("L" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        wbOld.Activate
        Sheet4.Range("L" & indeX).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        
        wbNew.Activate
        Range("M" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        wbOld.Activate
        Sheet4.Range("M" & indeX).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        
        wbNew.Activate
        Range("N" & pasteRow).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Range("J" & pasteRow).Value = "Price book"
        Range("F7").Value = pricE
        Range("L7").Value = matL
        Range("M7").Value = eHrs
        Range("N7").Value = aHrs
        Range("M1").Value = eRate
        Range("N1").Value = aRate
        Range("P7").Formula = "=L7+$M$1*M7+$N$1*N7"
    
    Else:
        wbNew.Activate
        
        If clasS <> 2 Then 'not a rotary -> delete rotary options
            Rows(Application.WorksheetFunction.Match("Zepf*", Range("A:A"), 0) & ":" & _
            Application.WorksheetFunction.Match("Zepf*", Range("A:A"), 0) + 1).Delete
        End If
        
        Range("B" & Application.WorksheetFunction.Match("*Skidding*", Range("A:A"), 0) + 2).Value = shiP
        Rows(Application.WorksheetFunction.Match("Options", Range("A:A"), 0) & ":" & _
            Application.WorksheetFunction.Match("Options", Range("A:A"), 0) + 3).Delete
        Range("F7").Value = pricE
        Range("L7").Value = matL
        Range("M7").Value = eHrs
        Range("N7").Value = aHrs
        Range("P7").Value = cosT
        'add labor rates
        Range("M1").Value = eRate
        Range("N1").Value = aRate
        wbOld.Activate
        Exit Function
    
    End If

    wbOld.Activate
    Sheet4.Activate
    Range("A1").Select
    ActiveSheet.ShowAllData
    wbNew.Activate
    
    amIndex = Application.WorksheetFunction.Match("Options", Range("A:A"), 0) + 2 'just chose unused variable
    
    If qtY > amIndex Then
        
        Range("B" & amIndex).Select
        Selection.AutoFill Destination:=Range("B" & amIndex & ":B" & qtY)
        
        Range("D" & amIndex).Select
        Selection.AutoFill Destination:=Range("D" & amIndex & ":D" & qtY)
        
        Range("G" & amIndex & ":H" & amIndex).Select
        Selection.AutoFill Destination:=Range("G" & amIndex & ":H" & qtY)
        
        Range("J" & amIndex).Select
        Selection.AutoFill Destination:=Range("J" & amIndex & ":J" & qtY)
        
        Range("P" & amIndex & ":Q" & amIndex).Select
        Selection.AutoFill Destination:=Range("P" & amIndex & ":Q" & qtY)
        
        Rows("5:" & qtY).EntireRow.AutoFit
        
        Rows("5:" & qtY).EntireRow.AutoFit
        
    Else: Rows("5:" & amIndex).EntireRow.AutoFit
    End If
    
    Range("A1").Select

ElseIf templFile = "AM_Pricing" Then
    wbNew.Activate
    
    If qtY = 1 Then
        If Range("A4").Value > "" Then
            If Range("A5").Value > "" Then
            
                Range("A4").Select
                Selection.End(xlDown).Select
                Selection.Offset(1, 0).Select
                amIndex = ActiveCell.Row
                
                wbOld.Activate
                Range("E" & indeX).Select
                Selection.Copy
                
                wbNew.Activate
                Range("A" & amIndex).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("G" & indeX).Select
                Selection.Copy
                
                wbNew.Activate
                Range("B" & amIndex).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("H" & indeX).Select
                Selection.Copy
                
                wbNew.Activate
                Range("L" & amIndex).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("I" & indeX).Select
                Selection.Copy
                
                wbNew.Activate
                Range("J" & amIndex).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Else
                wbOld.Activate
                Range("E" & indeX).Select
                Selection.Copy
                
                wbNew.Activate
                Range("A5").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("G" & indeX).Select
                Selection.Copy
                
                wbNew.Activate
                Range("B5").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("H" & indeX).Select
                Selection.Copy
                
                wbNew.Activate
                Range("L5").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("I" & indeX).Select
                Selection.Copy
                
                wbNew.Activate
                Range("J5").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            End If
        Else
            wbOld.Activate
            Range("E" & indeX).Select
            Selection.Copy
            
            wbNew.Activate
            Range("A4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            wbOld.Activate
            Range("G" & indeX).Select
            Selection.Copy
            
            wbNew.Activate
            Range("B4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            wbOld.Activate
            Range("H" & indeX).Select
            Selection.Copy
            
            wbNew.Activate
            Range("L4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            wbOld.Activate
            Range("I" & indeX).Select
            Selection.Copy
            
            wbNew.Activate
            Range("J4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        End If
        
    ElseIf qtY > 1 Then
        If Range("A4").Value > "" Then
            If Range("A5").Value > "" Then
                Range("A4").Select
                Selection.End(xlDown).Select
                Selection.Offset(1, 0).Select
                amIndex = ActiveCell.Row
                
                wbOld.Activate
                Range("E" & indeX).Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.SpecialCells(xlCellTypeVisible).Select
                Selection.Copy
                
                wbNew.Activate
                Range("A" & amIndex).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("G" & indeX).Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.SpecialCells(xlCellTypeVisible).Select
                Selection.Copy
                
                wbNew.Activate
                Range("B" & amIndex).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("H" & indeX).Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.SpecialCells(xlCellTypeVisible).Select
                Selection.Copy
                
                wbNew.Activate
                Range("L" & amIndex).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("I" & indeX).Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.SpecialCells(xlCellTypeVisible).Select
                Selection.Copy
                
                wbNew.Activate
                Range("J" & amIndex).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Else
                wbOld.Activate
                Range("E" & indeX).Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.SpecialCells(xlCellTypeVisible).Select
                Selection.Copy
                
                wbNew.Activate
                Range("A5").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("G" & indeX).Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.SpecialCells(xlCellTypeVisible).Select
                Selection.Copy
                
                wbNew.Activate
                Range("B5").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("H" & indeX).Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.SpecialCells(xlCellTypeVisible).Select
                Selection.Copy
                
                wbNew.Activate
                Range("L5").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                wbOld.Activate
                Range("I" & indeX).Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.SpecialCells(xlCellTypeVisible).Select
                Selection.Copy
                
                wbNew.Activate
                Range("J5").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            End If
        Else
            wbOld.Activate
            Range("E" & indeX).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.SpecialCells(xlCellTypeVisible).Select
            Selection.Copy
            
            wbNew.Activate
            Range("A4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            wbOld.Activate
            Range("G" & indeX).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.SpecialCells(xlCellTypeVisible).Select
            Selection.Copy
            
            wbNew.Activate
            Range("B4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            wbOld.Activate
            Range("H" & indeX).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.SpecialCells(xlCellTypeVisible).Select
            Selection.Copy
            
            wbNew.Activate
            Range("L4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            wbOld.Activate
            Range("I" & indeX).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.SpecialCells(xlCellTypeVisible).Select
            Selection.Copy
            
            wbNew.Activate
            Range("J4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        End If

    End If
    
    wbOld.Activate
    Sheet4.Activate
    Range("A1").Select
    If qtY > 0 Then
        ActiveSheet.ShowAllData
    End If
    wbNew.Activate
    Range("A4").Select
    
    If Range("A5").Value > 0 Then
        Selection.End(xlDown).Select
        amIndex = ActiveCell.Row
        Range("F4:I4").Select
        Selection.AutoFill Destination:=Range("F4:I" & amIndex)
        
        Range("K4").Select
        Selection.AutoFill Destination:=Range("K4:K" & amIndex)
        
        Range("M4").Select
        Selection.AutoFill Destination:=Range("M4:M" & amIndex)
    End If

    If Range("B4").Value = 0 Then
        Range("B4").Value = pricE
        Range("L4").Value = cosT
    End If
    
    'autofit columns has been moved out of the function
    'because it must come after tooling has been entered
    
Else: MsgBox "Pricing population aborted. Template file is named differently than expected."
End If

Exit Function

errhandler:
MsgBox "Error in PricingOptions function"

End Function
