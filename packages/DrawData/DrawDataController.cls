VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrawDataController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MACROTITLE = "Draw Data"

Private WithEvents m_workbook As Workbook
Attribute m_workbook.VB_VarHelpID = -1
Private WithEvents m_view As DrawDataViewForm
Attribute m_view.VB_VarHelpID = -1
Private m_converters As Collection
Private m_maximumDataPoints As Long

Private Function GetResultSheet() As Worksheet
    On Error GoTo GetResultSheetErr
    Dim result As Worksheet
    Set result = m_workbook.Sheets("Result")
    Set GetResultSheet = result
    Exit Function
GetResultSheetErr:
    If Err.Number = 9 Then
        Set result = m_workbook.Sheets.Add()
        result.Name = "Result"
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Public Sub Init(Workbook As Workbook)
    Set m_workbook = Workbook
    Set m_converters = New Collection
End Sub

Public Sub AddShape(shp As Shape)
    Dim m_conv As ShapeToDataConverter
    Set m_conv = New ShapeToDataConverter
    m_conv.Init shp, shp.height
    
    m_converters.Add m_conv
    
    If m_converters.Count = 1 Then
        m_maximumDataPoints = m_conv.MaxDataPoints
    Else
        If m_conv.MaxDataPoints < m_maximumDataPoints Then
            m_maximumDataPoints = m_conv.MaxDataPoints
        End If
    End If
End Sub

Public Sub Show()
    Set m_view = New DrawDataViewForm
    m_view.MaxDataPoints = m_maximumDataPoints
    m_view.Show vbModal
End Sub

Private Sub OutputResults(numberOfDataPoints As Long, outputRange As Range, movingAveragePeriod As Long)
    Dim resultRange As Range
    Dim resultRangeOrigin As Range
    Dim rowsInResult As Long
    
    If outputRange Is Nothing Then
        Dim resultSheet As Worksheet
        
        Set resultSheet = GetResultSheet
        resultSheet.Select
        DoEvents
        
        
        Set resultRangeOrigin = resultSheet.Range("B3")
    Else
        Set resultRangeOrigin = outputRange.Cells(1, 1)
    End If
    
    Dim resultTable As ListObject
    Set resultTable = resultRangeOrigin.ListObject
    If Not resultTable Is Nothing Then
        resultTable.Delete
    End If
    
    If movingAveragePeriod > 1 Then
        rowsInResult = numberOfDataPoints - (movingAveragePeriod - 1)
    Else
        rowsInResult = numberOfDataPoints
    End If
    
    Dim m_conv As ShapeToDataConverter
    Dim cnt As Long
    cnt = 0
    
    For Each m_conv In m_converters
        m_conv.numberOfDataPoints = numberOfDataPoints
        Set resultRange = resultRangeOrigin.Offset(0, cnt)
        Set resultRange = resultRange.Resize(rowsInResult, 1)
        resultRange.Clear
     
        If movingAveragePeriod > 1 Then
            resultRange.value = m_conv.GetSmoothenedOutput(movingAveragePeriod)
        Else
            resultRange.value = m_conv.Output
        End If
        cnt = cnt + 1
    Next
    
    Set resultRange = resultRangeOrigin
    Set resultRange = resultRange.Resize(rowsInResult, m_converters.Count)
    ConvertResultRangeToTable resultRange
    resultRange.Select
End Sub

Private Sub ConvertResultRangeToTable(resultRange As Range)
    On Error GoTo ConvertResultRangeToTableErr
    resultRange.Worksheet.ListObjects.Add SourceType:=xlSrcRange, Source:=resultRange, XlListObjectHasHeaders:=xlNo
    Exit Sub
ConvertResultRangeToTableErr:
    If Err.Number = 1004 Then
        MsgBox "The output range will overlap an existing table. Please choose a different location.", _
                vbExclamation, MACROTITLE
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Private Sub m_view_ApplyChanges(numberOfDataPoints As Long, outputRange As Variant, movingAveragePeriod As Long)
    Dim outRange As Range
    Set outRange = Nothing
    
    If Not IsEmpty(outputRange) Then
        If Len(outputRange) > 0 Then
            
            On Error GoTo m_view_ApplyChangesErr
            Set outRange = m_workbook.Application.Range(outputRange)
            Set outRange = outRange.Resize(1, 1)
            outputRange = outRange.Address
        End If
    End If
m_view_output:
    OutputResults numberOfDataPoints, outRange, movingAveragePeriod
    Exit Sub
m_view_ApplyChangesErr:
    Dim errorCount As Long
    If Err.Number = 1004 Then
        errorCount = errorCount + 1
        If errorCount < 10 Then
            Resume m_view_output
        Else
            MsgBox "An unforeseen error has happened. We are sorry about this. Please inform the developer that error no. 1004 has happened.", _
                    vbExclamation, MACROTITLE
            errorCount = 0
        End If
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

