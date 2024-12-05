Attribute VB_Name = "ExcelEmailModule"
Option Explicit

Private m_OutlookApplication As Object
    
Public Const ErrEmailNotAvailable As Long = 429
    
Public Function Init() As Boolean
    On Error GoTo InitErr
    Set m_OutlookApplication = GetObject(, "Outlook.Application")
    Init = True
    Exit Function
InitErr:
    Init = False
End Function
    
Public Sub SendEmail(recipient As String, subject As String, bodyText As String, attachmentFileName As String)
    Dim OutMail As Object
    
    Set OutMail = m_OutlookApplication.CreateItem(0) ' Email Item
    
    With OutMail
        .To = recipient
        .subject = subject
        .Body = bodyText
        .Attachments.Add attachmentFileName
        .Send
    End With
    
    Set OutMail = Nothing
End Sub
    
Public Function Dispose()
    Set m_OutlookApplication = Nothing
End Function

