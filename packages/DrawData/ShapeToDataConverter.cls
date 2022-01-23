VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeToDataConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_shape As Shape
Private m_inputArr() As Single
Private m_outputArr() As Variant
Private m_windowHeight As Double
Private m_maxDataPoints As Long

Private m_dataPoints As Long
Private m_movingAveragePeriod As Long

Private Sub GenerateOutput()
    ReDim m_outputArr(m_dataPoints, 1)
    
    Dim stepValue As Long
    Dim i As Long
    
    stepValue = m_maxDataPoints \ m_dataPoints
    
    For i = 1 To m_maxDataPoints Step stepValue
        Dim result As Single
        ' subtract Y position from Window height, because
        ' normal graphs assume 0 at bottom, and window co-ordinates
        ' assume 0 at top
        result = m_windowHeight - m_inputArr(i, 2)
        
        ' remove extras, because we do not want the entire window to be taken
        
        ' scale
        Dim outputLocation As Long
        If stepValue > 1 Then
            outputLocation = (i - 1) \ stepValue
            
        Else
            outputLocation = i - 1
        End If
        If outputLocation <= m_dataPoints Then
            m_outputArr(outputLocation, 0) = result
        End If
    Next
End Sub

Public Property Get numberOfDataPoints() As Long
    numberOfDataPoints = m_dataPoints
End Property

Public Property Let numberOfDataPoints(value As Long)
    If m_dataPoints <> value Then
        m_dataPoints = value
        GenerateOutput
    End If
End Property

Public Property Get MaxDataPoints() As Long
    MaxDataPoints = m_maxDataPoints
End Property

Public Property Get Output() As Variant()
    Output = m_outputArr
End Property

Public Function GetSmoothenedOutput(movingAveragePeriod As Long) As Variant()
    If movingAveragePeriod < 2 Or movingAveragePeriod > (m_dataPoints - 1) Then
        GetSmoothenedOutput = m_outputArr
    Else
        Dim resultArr() As Variant
        ReDim resultArr(m_dataPoints - movingAveragePeriod, 1)
        Dim i As Long
        Dim runningTotal As Single
        runningTotal = 0
        For i = 0 To (movingAveragePeriod - 1)
            runningTotal = runningTotal + m_outputArr(i, 0)
        Next
        For i = movingAveragePeriod To m_dataPoints
            resultArr(i - movingAveragePeriod, 0) = runningTotal / movingAveragePeriod
            runningTotal = runningTotal - m_outputArr(i - movingAveragePeriod, 0)
            runningTotal = runningTotal + m_outputArr(i, 0)
        Next
        GetSmoothenedOutput = resultArr
    End If
End Function

Public Sub Init(shp As Shape, height As Long)
    Set m_shape = shp
    m_windowHeight = shp.height + shp.Top
    m_inputArr = m_shape.Vertices
    m_maxDataPoints = UBound(m_inputArr, 1)
    m_dataPoints = m_maxDataPoints
    GenerateOutput
End Sub
