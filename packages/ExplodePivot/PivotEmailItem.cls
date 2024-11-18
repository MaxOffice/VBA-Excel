VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotEmailItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ItemLabel As MSForms.Label
Public WithEvents ItemTextBox As MSForms.TextBox
Attribute ItemTextBox.VB_VarHelpID = -1

Public Property Get Email() As String
    Email = ItemTextBox.Text
End Property

Public Property Get EmailEmpty() As Boolean
    EmailEmpty = (Trim$(ItemTextBox.Text) = "")
End Property

Public Property Get EmailValid() As Boolean
    EmailValid = IsValidEmail(ItemTextBox.Text)
End Property

Private Sub ItemTextBox_Change()
    If EmailEmpty Then
        ItemTextBox.BackColor = vbYellow
    ElseIf EmailValid Then
        ItemTextBox.BackColor = vbWhite
    Else
        ItemTextBox.BackColor = vbRed
    End If
End Sub

Private Function IsValidEmail(strEmail As String) As Boolean
    Dim intAtSignPos As Integer, intPeriodPos As Integer
    Dim strUsername As String, strDomain As String, strTLD As String

    ' Check for empty string
    If Len(strEmail) = 0 Then
        IsValidEmail = False
        Exit Function
    End If

    ' Find the position of the @ symbol
    intAtSignPos = InStr(1, strEmail, "@")
    If intAtSignPos = 0 Then
        IsValidEmail = False
        Exit Function
    End If

    ' Find the position of the period after the @ symbol
    intPeriodPos = InStr(intAtSignPos + 1, strEmail, ".")
    If intPeriodPos = 0 Then
        IsValidEmail = False
        Exit Function
    End If

    ' Extract username, domain, and TLD
    strUsername = Left$(strEmail, intAtSignPos - 1)
    strDomain = Mid$(strEmail, intAtSignPos + 1, intPeriodPos - intAtSignPos - 1)
    strTLD = Right$(strEmail, Len(strEmail) - intPeriodPos)

    ' Check for invalid characters in username, domain, and TLD
    If InStr(strUsername & strDomain & strTLD, """<>[]\|") > 0 Then
        IsValidEmail = False
        Exit Function
    End If

    ' Check for minimum length of domain and TLD
    If Len(strDomain) < 2 Or Len(strTLD) < 2 Then
        IsValidEmail = False
        Exit Function
    End If

    ' Basic validation passed
    IsValidEmail = True
End Function
