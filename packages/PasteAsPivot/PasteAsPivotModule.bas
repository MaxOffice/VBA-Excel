Attribute VB_Name = "PasteAsPivotModule"
Option Explicit

Private Const MACROTITLE = "Paste as Pivot by Dr Nitin Paranjape"
Private Const SELECTIONERR = "Please select adjacent cells from the data area of any Pivot Table."

Public Sub PasteAsPivot()
    ' Created by Dr Nitin Paranjape
    ' Purpose: Provide the ability to copy pivot data cells and bulk paste them as GetPivotData calls
    ' Benefit: Useful to keep live link with multiple cells form a pivot table for creating custom reports
    ' In absence of this tool, you would need to generate GetPivotData cell by cell making it impractical for use with multiple cells
    
    ' Psueudocode:
    ' get the range
    '
    ' work on each cell
    ' if at least one cell is inside datarange of pivot then proceed
    ' else give error on status bar and exit
    '
    ' ask user to specify where to paste
    
    '   if inside the datarange return getpivotdata function
    '   if not return nothing but keep the cell empty
    '   use the original offset while pasting
    '   at a later date insert a data validation to insert a tooltip as well
    '   right now insert the whole syntax as the tooltip
    '       Activecell.Validation.Add xlValidateInputOnly
    '       ActiveCell.Validation.InputMessage = "it works"
    
    
    'get the range

    On Error GoTo PasteAsPivotSelectionErr
    
    If Selection Is Nothing Then
        generror SELECTIONERR
        Exit Sub
    End If
    
    Dim rg As Range, curcell As Range
    
    Set rg = Selection
    ' If multiple ranges are selected, give error and exit
    If rg.Areas.Count > 1 Then
        generror "Please select only one area."
        Exit Sub
    End If
    
    Dim validcells As Integer
    validcells = 0
    
    Dim rcnt As Integer, ccnt As Integer, f As Integer, g As Integer
    Dim addrcheck As String
        
    rcnt = rg.Rows.Count
    ccnt = rg.Columns.Count
    
    For f = 1 To rcnt
        For g = 1 To ccnt
            Set curcell = rg.Cells(f, g)
            
            ' Check if current cell is in a PivotTable
            ' If there is no error accessing the PivotCell
            ' property, it is.
            On Error Resume Next
            addrcheck = curcell.PivotCell.Range.Address
            If Err.Number = 0 Then
                validcells = validcells + 1
            End If
            On Error GoTo 0
            
        Next
    Next
    
    If validcells = 0 Then
        generror SELECTIONERR
        Exit Sub
    End If

    On Error GoTo PasteAsPivotErr
    
    Dim dest As Worksheet
    Set dest = ActiveWorkbook.Sheets.Add
    
    
    Dim deststart As Range
    Set deststart = dest.Cells(1, 1)
    Dim destCell As Range
    Dim errno As Long
    
    ' process each cell in selected range one by one
    For f = 1 To rcnt
    
        For g = 1 To ccnt
        
            Set curcell = rg.Cells(f, g)
            
            'use the original offset while pasting
            Set destCell = deststart.Offset(f - 1, g - 1)
            
            ' Paste if valid
            SetGetPivotDataFormula curcell, destCell
        Next
    Next
    Exit Sub
PasteAsPivotSelectionErr:
    generror SELECTIONERR
    Exit Sub
PasteAsPivotErr:
    generror "We are sorry. An unexpected error occured."
End Sub

Private Sub SetGetPivotDataFormula(ByVal singleCell As Range, ByVal destCell As Range)
    On Error GoTo SetGetPivotDataFormulaErr
    
    Dim pt As PivotTable
    Set pt = singleCell.PivotTable
    
    Dim dataname As String
    dataname = singleCell.PivotCell.DataField
    
    Dim tooltip As String
    tooltip = dataname & ":" & vbCrLf
    
    ' Define base string for getpivotdata
    Dim base As String
    base = "=GetPivotdata(""" & dataname & """, '" & singleCell.Worksheet.Name & "'!" & pt.TableRange1.Cells(1, 1).Address
    
    Dim item As PivotItem
    Dim itemvalue As String, fieldname As String
    
    ' Add rowitems  to formula and tooltip
    For Each item In singleCell.PivotCell.RowItems
        itemvalue = item.Value
        fieldname = item.Parent
    
        base = base & ", """ & fieldname
        base = base & """"
        base = base & ", """ & itemvalue
        base = base & """"
        tooltip = tooltip & fieldname & " = " & itemvalue & vbCrLf
    Next
    
    ' Add column items to formula and tooltip
    For Each item In singleCell.PivotCell.ColumnItems
        itemvalue = item.Value
        fieldname = item.Parent
        base = base & ", """ & fieldname
        base = base & """"
        base = base & ", """ & itemvalue
        base = base & """"
        tooltip = tooltip & fieldname & " = " & itemvalue & vbCrLf
    Next
    
    base = base & ")"
    
    ' If no errors so far, paste the formula
    destCell.Formula = base
    
    ' and tooltip
    destCell.Validation.Add xlValidateInputOnly
    destCell.Validation.InputMessage = tooltip
    
    Exit Sub
SetGetPivotDataFormulaErr:
    ' On any error, just do nothing.
End Sub

Private Sub generror(str As String)
    MsgBox Trim(str), vbOKOnly, MACROTITLE
End Sub
