Attribute VB_Name = "ExplodePivotModule"
Option Explicit

Private Const MACROTITLE = "Pivot Filter Explode"

Public Const EMPTYBACKCOLOR As Long = &HFFFFFF       ' White
Public Const INVALIDBACKCOLOR As Long = &HC8C8FF     ' Red
Public Const VALIDBACKCOLOR As Long = &HFFDCD        ' Green

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
   
    ' If filter field has only one item selected, then there is no point in going ahead
    Dim selectedFilterValuesCount As Integer
    selectedFilterValuesCount = SelectedItems(basefield)
    
    If selectedFilterValuesCount < 2 Then
        MsgBox "There is only one item in the selected field." & _
                vbCrLf & "There is no need to explode this Pivot Table." & _
                vbCrLf & "Please enable 'Select Multiple Items', and choose two or more.", _
                vbExclamation, MACROTITLE
        Exit Sub
    End If
    
    ' Check existing sheets - baseWorkbook.Sheets.Count
    ' Check items in filter area - selectedFilterValuesCount
    ' Check if baseWorkbook.Sheets.Count + selectedFilterValuesCount <255 (it is just a sample number)
    ' In reality there is no limit on number of sheets (as long as memory is not exhausted)
    ' If you want to implement a limit just uncomment the code below and put the number you like
    '
    ' If baseWorkbook.Sheets.Count + selectedFilterValuesCount > 255 Then
    '     MsgBox "The number of sheets to be created exceeds the limit of 255." & _
    '     vbCrLf & "Please reduce number of filter items and try again", _
    '     vbInformation + vbOkOnly
    '     Exit Sub
    ' End If

    
    ' Inform user and get consent and other parameters to proceed
    Dim f As ExplodePivotForm
    Set f = New ExplodePivotForm
    f.Caption = MACROTITLE
    Set f.SelectionField = basefield
    f.Show vbModal
    
    ' If user chose to cancel, get out now
    If f.result = vbNo Then
        Exit Sub
    End If
        
    '---------------------------------------------------------------------
    ' Actual explode
    '---------------------------------------------------------------------
        
    ' Keep list of existing sheet names
    Dim sht As Worksheet
    
    ' Compile a list of worksheets already present in the workbook
    Dim oldSheets As String
    oldSheets = ","
    For Each sht In baseWorkbook.Sheets
        oldSheets = oldSheets & LCase$(sht.Name) & ","
    Next
    
    ' Explode the pivot table using the .ShowPages method
    basepivot.ShowPages (basefield)
    
    Dim doNotEmail As Boolean
    
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
            
            ' There are only two sheets in the new workbook
            Dim rawDataSheet As Worksheet
            If newPivotSheet.Index = 1 Then
                Set rawDataSheet = newWorkbook.Sheets(2)
            Else
                Set rawDataSheet = newWorkbook.Sheets(1)
            End If
            
            ' Delete or rename the raw data sheet
            If f.ReduceSize Then
                Application.DisplayAlerts = False
                rawDataSheet.Delete
                Application.DisplayAlerts = True
            Else
                rawDataSheet.Name = "Data"
            End If
            
            Dim fileName As String
            
            ' Save the newly created workbook, if required
            If f.SaveSheets Then
                fileName = Left$(f.BaseFilename, Len(f.BaseFilename) - 5) & " - " & newPivot.Name
                If Len(fileName) > 250 Then
                    fileName = Left$(fileName, 250)
                End If
                fileName = fileName & ".xlsx"
                newWorkbook.SaveAs fileName:=fileName, AddToMru:=False
            End If
            
            ' Attempt to Email the newly created workbook, if required
            If f.EmailSheets And Not doNotEmail Then
                Dim recipient As String
                Dim pmItem As PivotEmailItem
                
                
                Set pmItem = f.EmailItems(newPivot.Name)
                If Not pmItem Is Nothing Then
                    If Not pmItem.EmailEmpty Then
                        recipient = pmItem.Email
                        fileName = newPivot.Name & ".xlsx"
                        
                        On Error Resume Next
                        Err.Clear
                        
                        newWorkbook.SendMail Recipients:=recipient, Subject:="Attached: " & fileName
                        
                        If Err.Number <> 0 Then
                            Dim proceed As VbMsgBoxResult
                            proceed = MsgBox( _
                                "An error occured while trying to email " & fileName & ":" & _
                                vbCrLf & Err.Description & _
                                vbCrLf & "Should I try to email the rest of the sheets?" & _
                                vbCrLf & "Choose Cancel to stop the explode operation completely", _
                                vbInformation + vbYesNoCancel, _
                                MACROTITLE _
                            )
                            
                            If proceed = vbCancel Then
                                Exit Sub
                            ElseIf proceed = vbNo Then
                                doNotEmail = True
                            End If
                        End If
                        
                        On Error GoTo PivotFilterExplodeErr
                    End If
                End If
            End If
            
        End If ' If it is a new sheet
    
    Next ' Move to next sheet
    
    ' Inform about what just happened
    Dim finalMessage As String
    finalMessage = selectedFilterValuesCount & " new workbooks have been created." & _
                vbCrLf & "Each one has data filtered by items in the " & basefield.Name & " field."
    If f.SaveSheets Then
        finalMessage = finalMessage & vbCrLf & "These files have been saved as " & _
                        Left$(f.BaseFilename, Len(f.BaseFilename) - 5) & "*.xslx."
    Else
        finalMessage = finalMessage & vbCrLf & "These files have not been saved." & _
            vbCrLf & "You need to save these files or discard them, as required."
    End If
            
    MsgBox finalMessage, vbInformation, MACROTITLE
    
    ' That's it. Job done!!
    Exit Sub
    
PivotFilterExplodeErr:
    MsgBox "Sorry. Something went wrong." & Err.Description & Err.Source, vbExclamation, MACROTITLE
End Sub

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
