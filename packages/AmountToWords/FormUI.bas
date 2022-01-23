Attribute VB_Name = "FormUI"
Option Explicit


Public Sub ShowInsertFunctionForm()
    If Not Selection Is Nothing Then
        If Selection.Cells.Count = 1 Then
            InsertFunctionForm.Show vbModal
        Else
            MsgBox "Please select exactly one cell."
        End If
    End If
End Sub

Public Sub ShowInsertFunctionFormUIAction(button As IRibbonControl)
    ShowInsertFunctionForm
End Sub

Public Sub OpenDocumentation(button As IRibbonControl)
    On Error Resume Next
    
    If Not ActiveWorkbook Is Nothing Then
        ActiveWorkbook.FollowHyperlink Address:="https://efficiency365.com/2017/01/02/amount-to-words-macro/", NewWindow:=True
    End If
End Sub
