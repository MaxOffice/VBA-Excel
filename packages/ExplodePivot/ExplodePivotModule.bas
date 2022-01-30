Attribute VB_Name = "ExplodePivotModule"
Option Explicit


Private Const MACROTITLE = "Pivot Filter Explode"

Public Sub PivotFilterExplode()
    On Error GoTo PivotFilterExplodeErr
    
    Dim baseWorkbook As Workbook
    Dim baseSheet As Worksheet
    Dim basepivot As PivotTable
    Dim basefield As PivotField
    Dim rg As Range
       
    If ActiveWorkbook Is Nothing Then
        MsgBox "Please select a filter field on a Pivot Table.", vbExclamation, MACROTITLE
        Exit Sub
    End If
    
    If ActiveSheet Is Nothing Then
        MsgBox "Please select a filter field on a Pivot Table.", vbExclamation, MACROTITLE
        Exit Sub
    End If
    
    Set baseWorkbook = ActiveWorkbook
    Set baseSheet = ActiveSheet
    Set rg = ActiveCell

    ' Check if cursor is in pivot table field
    On Error Resume Next
    Set basefield = rg.PivotField
    If Err.Number <> 0 Then
        MsgBox "Pivot field not selected. You must select the Pivot Table filter field which you want to explode" & _
                vbCrLf & "and then run this macro.", vbExclamation, MACROTITLE
        Exit Sub
    End If
    On Error GoTo PivotFilterExplodeErr
    
    ' Check if the field selected is a page filter field
    Set basepivot = basefield.Parent
    If Application.Intersect(rg, basepivot.PageRange) Is Nothing Then
        MsgBox "Wrong field selected. You must select the Pivot Table filter field which you want to explode" & _
                vbCrLf & "and then run this macro.", vbExclamation, MACROTITLE
            Exit Sub
    End If
        
    ' check existing sheets - currsheets
    ' check items in filter area - filtercnt
    ' check if currsheets+filtercnt <255 (it is just a sample number)
    ' in reality there is no limit on number of sheets (as long as memory is not exhausted)
    
    ' If you want to implement a limit just uncomment the code below and put the number you like
    
    
    '    If basefield.PivotItems.Count + baseWorkbook.Sheets.Count > 255 Then
    '        MsgBox "The number of sheets to be created exceeds the limit of 255." & _
    '        vbCrLf & "Please reduce number of filter items and try again", _
    '        vbCritical
    '        Exit Sub
    '    End If
            
    
    ' If filter field has only one item selected, then there is no point in going ahead
    Dim selectedFilterValues As Integer
    selectedFilterValues = SelectedItems(basefield)
    
    If selectedFilterValues < 2 Then
        MsgBox "There is only one item in the selected field." & _
                vbCrLf & "There is no need to explode this Pivot Table.", vbExclamation, MACROTITLE
        Exit Sub
    End If
    
    ' Inform user and get consent to proceed
    If MsgBox( _
        "The Pivot Table will be exploded by items in the " & basefield.Name & " field. " & _
        vbCrLf & selectedFilterValues & " new workbooks will be created." & _
        vbCrLf & "If the data is missing for a particular filter item, the pivot table will be empty." & _
        vbCrLf & "Proceed?", _
        vbYesNo + vbInformation, _
        MACROTITLE _
    ) = vbNo Then
        Exit Sub
    End If

        
    '---------------------------------------------------------------------
    ' Actual explode
    '---------------------------------------------------------------------
    
    
    ' Keep list of existing sheet names
    Dim sht As Worksheet
    
    Dim oldSheets As String
    oldSheets = ","
    For Each sht In baseWorkbook.Sheets
        oldSheets = oldSheets & LCase$(sht.Name) & ","
    Next
    
    ' Explode the pivot table using the .ShowPages method
    basepivot.ShowPages (basefield)
    
    For Each sht In baseWorkbook.Sheets
        ' If it is a new sheet
        If InStr(1, oldSheets, "," & LCase$(sht.Name) & ",") < 1 Then
        
            ' .Move will create a new workbook with this sheet
            sht.Move
        
            ' New workbook is automatically activated so get reference to the
            ' first (and only) pivot table in it
            Dim newWorkbook As Workbook
            Dim newPivotSheet As Worksheet
            Set newWorkbook = ActiveWorkbook
            Set newPivotSheet = ActiveSheet
            
            ' Get reference to the filtered pivot table
            Dim newPivot As PivotTable
            Set newPivot = newPivotSheet.PivotTables(1)
            
            ' Get datarange and go to the bottom right cell
            '   range.cells(rows.count, columns.count)
            Dim currDR As Range
            Set currDR = newPivot.DataBodyRange
            
            ' Find the last cell in the data range - which is basically grand total
            Dim grandRG As Range
            Set grandRG = currDR.Cells(currDR.Rows.Count, currDR.Columns.Count)
            
            ' Use .ShowDetail to create the raw data sheet
            grandRG.ShowDetail = True

            ' This sheet has table containing the raw data. Table is always called Table1
            ' Right now, the pivot table on newSheet is still connected to the original data
            ' Now we need to connect it to the filtered data which we just created using ShowDetail
            ' Create new pivotcache by connecting it to this table
            Dim newPivotCache As PivotCache
            Set newPivotCache = newWorkbook.PivotCaches.Create( _
                                        SourceType:=xlDatabase, _
                                        SourceData:="Table1", _
                                        Version:=xlPivotTableVersion14 _
            )
            
            ' Connect pivot table to the new pivotcache
            newPivot.ChangePivotCache newPivotCache
            
            ' Change the name of the Pivot Table = the filter field item name
            ' A nice touch!
            ' Pivot Table names can start with numbers, special characters and can contain spaces :)
            newPivot.Name = newPivotSheet.Name
            
            newPivotSheet.Activate
            
            ' DO NOT save the newly created file (not really required)
            
        End If ' If it is a new sheet
    
    Next ' Move to next sheet
    
    ' Inform about what just happened
    
    MsgBox selectedFilterValues & " new workbooks have been created." & _
            vbCrLf & "Each one has data filtered by items in the " & basefield.Name & " field." & _
            vbCrLf & "These files have not been saved." & _
            vbCrLf & "You need to save these files or discard them, as required." _
            , vbInformation, MACROTITLE
    
    
    ' That's it. Job done!!
    Exit Sub
    
PivotFilterExplodeErr:
    MsgBox "Sorry.Something went wrong.", vbExclamation, MACROTITLE
End Sub


'---------------------------------------------------------------------
'Pseudocode
'(actual code may not be exactly in sync with this, but this will give you an idea of the logic used)

'---------------------------------------------------------------------
'set up environment
'---------------------------------------------------------------------
'one lengthy processing flag - longprocess bool
        'how to detect potentially long processing
        'file size
        'if more than x no of items in filter area

'---------------------------------------------------------------------
'pre-error handling
'---------------------------------------------------------------------
        
'check if cursor is in pivot table
    'if not error and exit
    'if yes continue
    
'check if pivot has anything in filter area
    'if not msg that filter is required and exit
    'if yes
        'check if multiple fields in filter area
        'if yes ask which field to explode by
        'check existing sheets - currsheets
        'check items in filter area - filtercnt
        'check if currsheets+filtercnt <255
            'if yes, error and exit
            'if no, continue
            
        
'explode selected field
        'message that explode will be done on field x and n no of new files will be created
    
'question: should the file be saved before the operation?
'most macros dont do this

'---------------------------------------------------------------------
'actual explode
'---------------------------------------------------------------------

'keep list of existing sheet names

'explode the pivot table using the .showpages method

'get list of sheet names again

'compare the two and iterate on new sheets

'keep reference to the original file

'---------------------------------------------------------------------
'process each newly created sheet
'---------------------------------------------------------------------
'for each sheet
    '.move to create a new workbook with this sheet
    'it is automatically activated so get reference to the first pivot in it
    'get datarange and go to the bottom right cell (range.cells(rows.count, columns.count)
    'use .showdetail to create the raw data sheet
    'this sheet has table containing the raw data. table is always called Table1
    '(in future research on this. if this is dicey, get the reference to this table)
    
    'now create new pivotcache by connecting it to this table
    'change the name of the pivot table = the filter field item name
    
    '(decide whether to delete that sheet - better to keep it)
    'activate the pivot table sheet
    'DO NOT save the newly created file (not really required)
'move to the next sheet in base file

'show end message with no of files created
'end

'--------------------------------------------

' This function is required because the .VisibleItems
' property of a PivotField does not work as expected.
Private Function SelectedItems(f As PivotField) As Integer
    Dim result As Integer
    If f Is Nothing Then
        result = 0
    Else
        Dim item As PivotItem
        result = 0
        For Each item In f.PivotItems
            If item.Visible Then
                result = result + 1
            End If
        Next
        ' If (all) is selected, result is 0
        If result = 0 Then
            result = f.ParentItems.Count
        End If
    End If
    SelectedItems = result
End Function

Public Sub ExplodePivotUIAction(button As IRibbonControl)
    PivotFilterExplode
End Sub
