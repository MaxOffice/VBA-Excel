Attribute VB_Name = "DrawDataModule"
Option Explicit

Private Const SELECTERRMSG = "Please select one or more ""squiggle"" shapes."

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
            MsgBox SELECTERRMSG
        End If
    Else
        MsgBox SELECTERRMSG
    End If
    Exit Sub
DrawDataErr:
    If Err.Number = 438 Then
        MsgBox SELECTERRMSG
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Public Sub DrawDataUIAction(button As IRibbonControl)
    DrawData
End Sub
