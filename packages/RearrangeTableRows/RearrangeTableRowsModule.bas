Attribute VB_Name = "RearrangeTableRowsModule"
Option Explicit

Private Const MACROTITLE = "Rearrange Table Rows"

Public Sub MoveTableRowUpUIAction(ctl As IRibbonControl)
    MoveTableRowUp
End Sub

Public Sub MoveTableRowDownUIAction(ctl As IRibbonControl)
    MoveTableRowDown
End Sub

Public Sub MoveTableRowUp()
    MoveTableRow 1
End Sub

Public Sub MoveTableRowDown()
    MoveTableRow -1
End Sub

Private Sub MoveTableRow(direction As Integer)
    On Error GoTo MoveTableRowErr
    
    Dim lo As ListObject
    Dim tbl As ListObject
    Dim rowIdx As Long
    
    If ActiveCell Is Nothing Then
        MsgBox "Please select a cell inside a valid Excel table and try again.", _
                vbExclamation + vbOKOnly, MACROTITLE
        Exit Sub
    End If

    ' Check if selection is inside a table
    On Error Resume Next
    Set lo = ActiveCell.ListObject
    On Error GoTo MoveTableRowErr

    If lo Is Nothing Then
        MsgBox "Please select a cell inside a valid Excel table and try again.", _
                vbExclamation + vbOKOnly, MACROTITLE
        Exit Sub
    End If

    Set tbl = lo

    ' Find the table row index (relative to DataBodyRange)
    rowIdx = ActiveCell.Row - tbl.DataBodyRange.Rows(1).Row + 1

    If rowIdx < 1 Or rowIdx > tbl.ListRows.Count Then
        MsgBox "Please select a data row inside the table.", _
                vbExclamation + vbOKOnly, MACROTITLE
        Exit Sub
    End If

    ' Check move direction and bounds
    If direction = 1 Then ' Move Up
        If rowIdx = 1 Then
            MsgBox "Already at the first row. Cannot move up.", _
                    vbExclamation + vbOKOnly, MACROTITLE
            Exit Sub
        End If
        ' Swap with row above
        SwapTableRows tbl, rowIdx, rowIdx - 1
        tbl.DataBodyRange.Rows(rowIdx - 1).Cells(1).Select
    ElseIf direction = -1 Then ' Move Down
        If rowIdx = tbl.ListRows.Count Then
            MsgBox "Already at the last row. Cannot move down.", _
                    vbExclamation + vbOKOnly, MACROTITLE
            Exit Sub
        End If
        ' Swap with row below
        SwapTableRows tbl, rowIdx, rowIdx + 1
        tbl.DataBodyRange.Rows(rowIdx + 1).Cells(1).Select
    Else
        MsgBox "Invalid direction parameter. Use 1 for up, -1 for down.", _
            vbCritical + vbOKOnly, MACROTITLE
    End If
    Exit Sub
MoveTableRowErr:
    MsgBox "Sorry, something went wrong: " & Err.Description & Err.Source, _
            vbExclamation + vbOKOnly, MACROTITLE
End Sub

' Helper function to swap two rows in a table
Private Sub SwapTableRows(tbl As ListObject, idx1 As Long, idx2 As Long)
    Dim arr1, arr2
    arr1 = tbl.DataBodyRange.Rows(idx1).Value
    arr2 = tbl.DataBodyRange.Rows(idx2).Value
    tbl.DataBodyRange.Rows(idx1).Value = arr2
    tbl.DataBodyRange.Rows(idx2).Value = arr1
End Sub

