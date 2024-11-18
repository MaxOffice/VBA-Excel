VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DrawDataViewForm 
   Caption         =   "Draw Data"
   ClientHeight    =   2390
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5390
   OleObjectBlob   =   "DrawDataViewForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DrawDataViewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_numberOfDataPoints As Long
Private m_maxDataPoints As Long

Public Event ApplyChanges(numberOfDataPoints As Long, ByRef outputRange As Variant, movingAveragePeriod As Long)

Public Property Get MaxDataPoints() As Long
    MaxDataPoints = m_maxDataPoints
End Property

Public Property Let MaxDataPoints(value As Long)
    m_maxDataPoints = value
    txtNumberOfDataPoints = value
    lblMax = "(maximum " & m_maxDataPoints & ".)"
    spnMovingAverage.Max = m_maxDataPoints - 1
End Property


Private Sub chkSmoothen_Change()
    txtPeriod.Enabled = chkSmoothen.value
    spnMovingAverage.Enabled = chkSmoothen.value
End Sub

Private Sub cmdApply_Click()
    If CLng(txtNumberOfDataPoints) > m_maxDataPoints Then
        txtNumberOfDataPoints = m_maxDataPoints
    End If
    
    Dim movingAveragePeriod As Long
    If chkSmoothen.value Then
        movingAveragePeriod = CLng(txtPeriod)
        If movingAveragePeriod < 2 Then
            txtPeriod = "2"
        Else
            If movingAveragePeriod >= m_maxDataPoints Then
                txtPeriod = CStr(m_maxDataPoints - 1)
            End If
        End If
        movingAveragePeriod = CLng(txtPeriod)
    Else
        movingAveragePeriod = 1
    End If
    
    Dim outputRange As Variant
    outputRange = reStartCell.value
    RaiseEvent ApplyChanges(CLng(txtNumberOfDataPoints.Text), outputRange, movingAveragePeriod)
    reStartCell.value = outputRange
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub spnMovingAverage_Change()
    txtPeriod = spnMovingAverage.value
End Sub
