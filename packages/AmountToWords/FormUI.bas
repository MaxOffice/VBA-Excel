Attribute VB_Name = "FormUI"
Option Explicit

Private Const MACROTITLE = "Numbers To Words"

Public Sub ShowInsertFunctionForm()
    On Error GoTo ShowInsertFunctionFormErr
    
    If Selection.Cells.Count = 1 Then
        InsertFunctionForm.Show vbModal
    Else
        MsgBox "Please select exactly one cell.", vbExclamation, MACROTITLE
    End If

    Exit Sub
ShowInsertFunctionFormErr:
    MsgBox "Please select a blank cell in a worksheet and invoke this macro.", _
            vbExclamation, MACROTITLE
End Sub

Public Sub ShowInsertFunctionFormUIAction(button As IRibbonControl)
    ShowInsertFunctionForm
End Sub

Public Sub OpenDocumentation(button As IRibbonControl)
    On Error GoTo OpenDocumentationErr   
        
    ActiveWorkbook.FollowHyperlink Address:="https://efficiency365.com/2017/01/02/amount-to-words-macro/", NewWindow:=True
    
    Exit Sub
OpenDocumentationErr:
    MsgBox "Could not open documentation. At least one workbook needs to be open.", _
            vbExclamation, MACROTITLE
End Sub
