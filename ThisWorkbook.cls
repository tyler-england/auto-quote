VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If InStr(UCase(Application.UserName), "ENGLAND") > 0 Then
        x = ExportModules
    End If
End Sub

Private Sub Workbook_Open()
On Error Resume Next
    Application.EnableAutoComplete = False
    ThisWorkbook.AutoSaveOn = False
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
On Error Resume Next
    Application.EnableAutoComplete = False
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
On Error Resume Next
    Application.EnableAutoComplete = True
End Sub


