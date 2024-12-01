VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExplodePivotForm 
   ClientHeight    =   4610
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6690
   OleObjectBlob   =   "ExplodePivotForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExplodePivotForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_selectionField As PivotField
Private m_selectionCount As Long
Private m_result As VbMsgBoxResult
Private m_emailAddresses As Collection

Public Property Get SelectionField() As PivotField
    Set SelectionField = m_selectionField
End Property

Public Property Set SelectionField(value As PivotField)
    Set m_selectionField = value
    
    Dim itemcount As Long
    
    If value.EnableMultiplePageItems Then
        ' Multiple Item Selection is enabled.
        ' We need to count the visible items.
        
        Dim item As PivotItem
        itemcount = 0
        For Each item In value.PivotItems
            If item.Visible Then
                itemcount = itemcount + 1
            End If
        Next
        
        ' If (all) is selected, itemcount is 0
        If itemcount = 0 Then
            itemcount = value.ParentItems.Count
        End If
    Else
        ' Multiple item selection is not enabled.
        ' If (All) is selected, the count is all
        ' the items in the table, otherwise one.
        If value.AllItemsVisible Then
            itemcount = value.PivotItems.Count
        Else
            itemcount = 1
        End If
    End If
    
    lblItemHeader = value.Name
    
    lblMessage = "The Pivot Table will be exploded by items in the " & value.Name & " field. " & _
        vbCrLf & itemcount & " new workbook(s) will be created." & _
        vbCrLf & "If the data is missing for a particular filter item, the pivot table will be empty."
    
    m_selectionCount = itemcount
    
    PopulateEmailsUI
End Property

Public Property Get SaveSheets() As Boolean
    SaveSheets = chkSaveSheets.value
End Property

Public Property Get BaseFilename() As String
    BaseFilename = txtBaseFilename.Text
End Property

Public Property Let BaseFilename(newValue As String)
    txtBaseFilename.Text = newValue
End Property

Public Property Get EmailSheets() As Boolean
    EmailSheets = chkEmailSheets.value
End Property

Public Property Get EmailItems() As Collection
    Set EmailItems = m_emailAddresses
End Property

Public Property Get ReduceSize() As Boolean
    ReduceSize = chkReduceSize.value
End Property

Public Property Get result() As VbMsgBoxResult
    result = m_result
End Property

Private Sub chkEmailSheets_Change()
    Dim item As PivotEmailItem
    Dim desired As Boolean
    
    desired = chkEmailSheets.value
    For Each item In m_emailAddresses
        item.ItemLabel.Visible = desired
        item.ItemTextBox.Visible = desired
    Next

    If desired Then
        Set item = m_emailAddresses.item(1)
        item.ItemTextBox.SetFocus
    End If
End Sub

Private Sub chkSaveSheets_Change()
    Dim desired As Boolean
    desired = chkSaveSheets.value
    
    If desired Then
        cmdBrowse.Enabled = True
    Else
        txtBaseFilename.Text = ""
        cmdBrowse.Enabled = False
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim result As Variant
    
    result = Application.GetSaveAsFilename( _
                            InitialFileName:=Caption, _
                            FileFilter:="Excel Workbook Files (*.xlsx),*.xlsx", _
                            Title:=Caption _
    )
    
    If VarType(result) = vbString Then
        BaseFilename = CStr(result)
    End If
End Sub

Private Sub cmdCancel_Click()
    m_result = vbNo
    Hide
End Sub

Private Sub cmdOk_Click()
    If SaveSheets Then
        If Trim$(txtBaseFilename.Text) = "" Then
            MsgBox "Please choose a base file name to save exploded sheets.", _
                vbInformation + vbOKOnly, _
                Caption
            Exit Sub
        End If
    End If
    If EmailSheets Then
        If Not ValidateEmails() Then
            MsgBox "At least one email is invalid. Please provide valid email ids, or leave them empty", _
                        vbInformation + vbOKOnly, _
                        Caption
            Exit Sub
        End If
    End If
    
    m_result = vbYes
    Hide
End Sub

Private Sub UserForm_Initialize()
    m_result = vbNo
    m_selectionCount = 0
    Set m_emailAddresses = New Collection
End Sub

Private Sub PopulateEmailsUI()
    Dim allSelected As Boolean
    
    If m_selectionField.EnableMultiplePageItems Then
        allSelected = (m_selectionCount = m_selectionField.ParentItems.Count)
    Else
        allSelected = m_selectionField.AllItemsVisible
    End If
    
    Dim item As PivotItem
    Dim newItem As PivotEmailItem
    Dim newLabel As MSForms.Label
    Dim newTextBox As MSForms.TextBox
    
    Dim newLabelTop As Long
    
    newLabelTop = lblEmailHeader.Top + lblEmailHeader.Height + 5
    
    For Each item In m_selectionField.PivotItems
        If item.Visible Or allSelected Then
            Set newItem = New PivotEmailItem
            
            Set newLabel = fmEmailDetails.Controls.Add("Forms.Label.1")
            With newLabel
                .Caption = item.value
                .Top = newLabelTop
                .Left = 5
                .Visible = False
            End With
            
            Set newTextBox = fmEmailDetails.Controls.Add("Forms.TextBox.1")
            With newTextBox
                .Top = newLabelTop - 5
                .Left = 150
                .Width = fmEmailDetails.Width - .Left - 30
                .BackColor = EMPTYBACKCOLOR
                .Visible = False
            End With
            
            newLabelTop = newLabelTop + newLabel.Height + 5
            
            Set newItem.ItemLabel = newLabel
            Set newItem.ItemTextBox = newTextBox
            
            m_emailAddresses.Add newItem, item.value
        End If
    Next
    
    If newLabelTop > fmEmailDetails.Height Then
        fmEmailDetails.ScrollHeight = newLabelTop
        fmEmailDetails.ScrollBars = fmScrollBarsVertical
    Else
        fmEmailDetails.ScrollHeight = fmEmailDetails.Height
        fmEmailDetails.ScrollBars = fmScrollBarsNone
    End If
End Sub

Private Function ValidateEmails() As Boolean
    Dim item As PivotEmailItem
    
    For Each item In m_emailAddresses
        If Not item.EmailEmpty Then
            If Not item.EmailValid Then
                ValidateEmails = False
                Exit Function
            End If
        End If
    Next
    
    ValidateEmails = True
End Function
