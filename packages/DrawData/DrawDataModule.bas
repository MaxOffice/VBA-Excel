Attribute VB_Name = "DrawDataModule"
Option Explicit

Private Const SELECTERRMSG = "Please select one or more ""squiggle"" shapes."
Private Const MACROTITLE = "Draw Data"

Public Sub DrawData()
    On Error GoTo DrawDataErr:
    If Not Selection Is Nothing Then
        Dim selectedShapes As ShapeRange
        Set selectedShapes = Selection.ShapeRange
        Dim selectedShape As Shape
        Dim cont As DrawDataController
        
        For Each selectedShape In selectedShapes
            ' Best guess for squiggle
            If selectedShape.AutoShapeType = msoShapeNotPrimitive Then
                If cont Is Nothing Then
                    Set cont = New DrawDataController
                    cont.Init selectedShape.Parent.Parent
                End If
                cont.AddShape selectedShape
            End If
        Next
        If Not cont Is Nothing Then
            cont.Show
        Else
            MsgBox SELECTERRMSG, vbExclamation, MACROTITLE
        End If
    Else
        MsgBox SELECTERRMSG, vbExclamation, MACROTITLE
    End If
    Exit Sub
DrawDataErr:
    If Err.Number = 438 Then
        MsgBox SELECTERRMSG, vbExclamation, MACROTITLE
    Else
        MsgBox "Sorry. An unforeseen error happened. " & _
                vbCrlf & "Please inform the developers that error number " & Err.Number & " happened at startup.", _ 
                vbExclamation, MACROTITLE
    End If
End Sub

Public Sub DrawDataUIAction(button As IRibbonControl)
    DrawData
End Sub
