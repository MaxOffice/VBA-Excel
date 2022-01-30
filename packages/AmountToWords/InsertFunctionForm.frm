VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertFunctionForm 
   Caption         =   "Use the AmountToWords function"
   ClientHeight    =   5595
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   10210
   OleObjectBlob   =   "InsertFunctionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InsertFunctionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const MACROTITLE = "Numbers To Words"

Private initialized As Boolean
Private WithEvents app As Application
Attribute app.VB_VarHelpID = -1
Private selectedRange As Range

Private Sub ClearSample()
    lblSample = ""
    lblFunctionSyntax = ""
End Sub

Private Sub ShowSample()
    If Not initialized Then
        Exit Sub
    End If
      
    On Error GoTo ShowSampleErr
    Dim amount As Currency
    
    If IsNumeric(txtAmount.Text) Then
        amount = CCur(txtAmount.Text)
        
        lblSample = AmountToWords( _
                amount, _
                chkIncludeCommas, _
                chkIncludeOnly, _
                chkShowRupees, _
                chkShowPaise, _
                chkLastAnd, _
                chkRupeesAfter, _
                chkPaiseAfter, _
                cmbResultCase.List(cmbResultCase.ListIndex, 1), _
                txtRupees, _
                txtPaise, _
                txtLakhs, _
                txtZeroPaise _
        )
        
        Dim syntax As String
        syntax = "=AmountToWords(" & _
                         txtAmount.Text & ", " & _
                         chkIncludeCommas & ", " & _
                         chkIncludeOnly & ", " & _
                         chkShowRupees & ", " & _
                         chkShowPaise & ", " & _
                         chkLastAnd & ", " & _
                         chkRupeesAfter & ", " & _
                         chkPaiseAfter & ", " & _
                         """" & cmbResultCase.List(cmbResultCase.ListIndex, 1) & """, " & _
                         """" & txtRupees & """, " & _
                         """" & txtPaise & """, " & _
                         """" & txtLakhs & """, " & _
                         """" & txtZeroPaise & """) "
        lblFunctionSyntax = syntax
        
        cmdInsert.Enabled = True
    Else
        ClearSample
        cmdInsert.Enabled = False
    End If
        
    Exit Sub
ShowSampleErr:
    If Err.number = 6 Then
        MsgBox "The amount provided is too large for the AmountToWords function.", _
                vbExclamation, MACROTITLE
        ClearSample
    Else
        MsgBox "An unforseen error happened.", _
                vbExclamation, MACROTITLE
    End If
End Sub

Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    cmdInsert.Enabled = False
    Set selectedRange = Nothing
    If Not Sh Is Nothing Then
        If Target.Cells.Count = 1 Then
            Set selectedRange = Target
            If lblFunctionSyntax.Caption <> "" Then
                cmdInsert.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub chkShowPaise_Change()
    ShowSample
    txtPaise.Enabled = chkShowPaise
    chkPaiseAfter.Enabled = chkShowPaise
    lblZeroPaise.Enabled = chkShowPaise
    txtZeroPaise.Enabled = chkShowPaise
End Sub

Private Sub chkShowRupees_Change()
    ShowSample
    txtRupees.Enabled = chkShowRupees
    chkRupeesAfter.Enabled = chkShowRupees
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    If lblFunctionSyntax <> "" Then
        Dim funcText As String
        funcText = lblFunctionSyntax.Caption
        Dim copyData As DataObject
        Set copyData = New DataObject
        With copyData
            .Clear
            .SetText funcText, 1
            .PutInClipboard
        End With
        Set copyData = Nothing
    End If
End Sub

Private Sub cmdInsert_Click()
    If lblFunctionSyntax <> "" Then
        If Not selectedRange Is Nothing Then
            selectedRange.Formula = lblFunctionSyntax.Caption
        End If
    End If
End Sub

Private Sub cmdResetOptions_Click()
    ClearSample
    txtAmount = ""
    txtLakhs = "lakhs"
    txtPaise = "paise"
    txtRupees = "rupees"
    txtZeroPaise = "zero"
    chkIncludeCommas = False
    chkIncludeOnly = False
    chkLastAnd = False
    chkPaiseAfter = False
    chkRupeesAfter = False
    chkShowPaise = True
    chkShowRupees = True
    cmbResultCase.ListIndex = 0
End Sub

Private Sub txtAmount_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With txtAmount
        Select Case KeyAscii
            Case Asc("0") To Asc("9")
            Case Asc("-")
                If InStr(1, .Text, "-") > 0 Or .SelStart > 0 Then
                    KeyAscii = 0
                End If
            Case Asc(".")
                If InStr(1, .Text, ".") > 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With
End Sub

Private Sub txtAmount_Change()
    ShowSample
End Sub

Private Sub txtLakhs_Change()
    ShowSample
End Sub

Private Sub txtPaise_Change()
    ShowSample
End Sub

Private Sub txtRupees_Change()
    ShowSample
End Sub

Private Sub txtZeroPaise_Change()
    ShowSample
End Sub

Private Sub chkIncludeCommas_Change()
    ShowSample
End Sub

Private Sub chkIncludeOnly_Change()
    ShowSample
End Sub

Private Sub chkLastAnd_Change()
    ShowSample
End Sub

Private Sub chkPaiseAfter_Change()
    ShowSample
End Sub

Private Sub chkRupeesAfter_Change()
    ShowSample
End Sub

Private Sub cmbResultCase_Change()
    ShowSample
End Sub

Private Sub UserForm_Initialize()
    With cmbResultCase
        .AddItem
        .AddItem
        .AddItem
        .AddItem
        .List(0, 0) = "UPPER CASE"
        .List(0, 1) = "u"
        .List(1, 0) = "lower case"
        .List(1, 1) = "l"
        .List(2, 0) = "Title Case"
        .List(2, 1) = "t"
        .List(3, 0) = "Sentence case"
        .List(3, 1) = "s"
        .ListIndex = 0
    End With
    Set app = Application
    With app
        Dim selectedRange As Range
        Set selectedRange = .Selection
        If selectedRange.Cells.Count = 1 Then
            If IsNumeric(selectedRange.Value) Then
                txtAmount = selectedRange.Value
            End If
        End If
        app_SheetSelectionChange .ActiveSheet, selectedRange
        
    End With
    initialized = True
End Sub

