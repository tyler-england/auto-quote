VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim heads As Integer, machine1 As String, machine2 As String, row1 As Integer
    
    On Error Resume Next

    Set typ = Range("D5")
    'quote type changes
    
    Set modeL = Range("D6")
    'Model changes
    
    Set salesRep = Range("D7")
    
    If Not Application.Intersect(modeL, Range(Target.Address)) _
           Is Nothing Then

        ActiveSheet.Unprotect
        Application.EnableEvents = False
           
        Application.ScreenUpdating = False
        
        On Error Resume Next
        
        heads = Sheet3.Range("C22").Value
        
        If Sheet3.Range("C21").Value > 0 Then
        'automatic/rotary
'            Range("C14").Select
'            Selection.Copy
'            Range("D14").Select
'            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'                SkipBlanks:=False, Transpose:=False
'            Application.CutCopyMode = False
            'does not require VFFS answer
            Range("D13").Interior.ColorIndex = 0
            Range("D13").ClearFormats
            Range("D13").HorizontalAlignment = xlCenter
            Range("D13").Borders.LineStyle = xlContinuous
            Range("D13").Formula = "No"
            
        ElseIf heads = 2 Then
        '2 heads
            'does not require VFFS answer
            Range("D13").Interior.ColorIndex = 0
            Range("D13").ClearFormats
            Range("D13").HorizontalAlignment = xlCenter
            Range("D13").Borders.LineStyle = xlContinuous
            Range("D13").Formula = "No"
            
        Else:
            'requires VFFS answer
            Range("D12").Select
            Selection.Copy
            Range("D13").Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            
            Range("D13").Formula = ""
        
        End If
        
        If Sheet3.Range("C21").Value = 2 Then 'rotary
            
            Range("D12").Select
            Selection.Copy
            Range("D14:D16").Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            Range("D14").Formula = ""
            Range("D15").Formula = ""
            Range("D16").Formula = ""
            
            If Range("D6").Value = 6600 Then
                Range("D15").Interior.ColorIndex = 0
                Range("D15").ClearFormats
                Range("D13").HorizontalAlignment = xlCenter
                Range("D15").Borders.LineStyle = xlContinuous
                Range("D15").Formula = "N/A"
                Range("D15").HorizontalAlignment = xlCenter
            End If
            
        ElseIf Sheet3.Range("C21").Value <> 2 Then
            Range("D14:D16").Select
            Selection.Interior.ColorIndex = 0
            Selection.ClearFormats
            Selection.HorizontalAlignment = xlCenter
            Range("D13").HorizontalAlignment = xlCenter
            Selection.Borders.LineStyle = xlContinuous
            Range("D14").Formula = "N/A"
            Range("D15").Formula = "N/A"
            Range("D16").Formula = "N/A"
            
        End If
        
        If UCase(Range("D5").Value) = "AM" Then
        'AM quote
            
            Range("D13").Clear
            Range("D12").Select
            Selection.Copy
            Range("D13").Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            
            Range("D14:D16").Select
            Selection.ClearFormats
            Selection.Borders.LineStyle = xlContinuous
            Selection.Interior.ColorIndex = 0

            'does not require VFFS answer
            Range("D14").Formula = "N/A"
            Range("D15").Formula = "N/A"
            Range("D16").Formula = "N/A"
        
        End If
        
        If Sheet3.Range("C21").Value = 0 Then
            machine1 = "Non-Rotaries"
            machine2 = "Semiautomatic"
        ElseIf Sheet3.Range("C21").Value = 1 Then
            machine1 = "Non-Rotaries"
            machine2 = "Automatic"
        Else
            machine1 = "Rotaries"
            machine2 = ""
        End If
        
        If Range("D6").Value = "" Then
        
            Range("F9:K100").ClearContents
            Range("F10:K100").Borders.LineStyle = xlNone
            Range("G9:K9").Select
            Range("G9:K9").Merge
            Range("F9:K9").Borders.LineStyle = xlContinuous
            
        ElseIf Range("D6").Value <> "" Then
    
            Call Options(machine1, machine2, heads)
            'populates machine specific options
    
            Range("G9").Select
            Selection.End(xlDown).Select
            If ActiveCell.Row < 100 Then
                row1 = ActiveCell.Row
            Else: row1 = 100
            End If
    
            Range("F" & row1 + 1 & ":K101").Select
            Selection.Borders.LineStyle = xlNone
            
            Range("F9:F" & row1).Select
            Selection.Borders.LineStyle = xlContinuous
            
            Range("G9:K9").Select
            Range("G9:K9").Merge
            Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
            Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
            Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
            Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
            Range("G9:K9").Copy
    
            Range("G9:K" & row1).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            Selection.HorizontalAlignment = xlLeft
    
            Range("F9:K" & row1).Select
            'Selection.Borders.LineStyle = xlContinuous
        End If
        
        Range("D7").Select
        ActiveSheet.Protect

        Application.EnableEvents = True
        Application.ScreenUpdating = True
    
    ElseIf Not Application.Intersect(typ, Range(Target.Address)) _
           Is Nothing Then
           
        Dim temp As String
        
        ActiveSheet.Unprotect
        
        temp = Range("D6").Value
        Range("D13").Clear
        Range("D12").Select
            Selection.Copy
            Range("D13").Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            
        ActiveSheet.Protect
        
        Range("D6").Value = temp
        Range("D6").Select
        
    ElseIf Not Application.Intersect(salesRep, Range(Target.Address)) _
           Is Nothing Then
    
        If UCase(salesRep.Value) Like "MASSIMO*" Then
            Range("G5").Value = 0.025
        ElseIf UCase(salesRep.Value) Like "IVAN*TAN*" Then
            Range("G5").Value = 0.15
        ElseIf salesRep.Value = 0 Then
            Range("G5").ClearContents
        End If
    
    End If

End Sub
