Attribute VB_Name = "AmountToWordFunctions"
Option Explicit

Private Const RupeeConst As String = "÷"
Private Const PaisaConst As String = "Ë"
Private Const RspConnectorConst As String = "°"
Private Const LacsConst As String = "¬"
Private Const ZeroPaisaConst As String = "¿"
Private Const AndConst As String = "½"

Private Function Decorate( _
        numberaswords As String, _
        paisaPosition As Long, _
        SpecifyRupees As Boolean, _
        RupeeName As String, _
        SpecifyPaisa As Boolean, _
        PaisaName As String, _
        CharacterCase As String, _
        RupeesAfter As Boolean, _
        PaisaAfter As Boolean, _
        AddOnly As Boolean, _
        LacsName As String, _
        ZeroPaisaName As String _
        ) As String
    
    Dim result As String
    result = numberaswords
    
    ' Do rupees
    If SpecifyRupees Then
        If RupeesAfter Then
            If paisaPosition <> 0 Then
                result = Left$(result, paisaPosition - 1) & RupeeConst & " " & Mid$(result, paisaPosition)
            Else
                result = result & RupeeConst & " "
            End If
        Else
            ' Add the rupee name before
            result = RupeeConst & " " & result
        End If
        paisaPosition = paisaPosition + Len(RupeeConst) + 1
    End If

    ' Do paise
    If SpecifyPaisa Then
        If PaisaAfter Then
            result = result & " " & PaisaConst
        Else
            Dim actualPaisaPosition As Long
            actualPaisaPosition = paisaPosition + Len(RspConnectorConst)
            result = Left$(result, actualPaisaPosition) & PaisaConst & " " & Mid$(result, actualPaisaPosition + 1)
        End If
    End If
    
    ' Do Only
    If AddOnly Then
        result = result & " only"
    End If
    
    ConvertCase result, CharacterCase, RupeeName, PaisaName, LacsName, ZeroPaisaName
    
    Decorate = result
End Function

Private Function PCase$(ByVal str As String)
    If Len(str) > 0 Then
        str = LCase$(str)
        Mid$(str, 1, 1) = UCase$(Mid$(str, 1, 1))
    End If
    PCase$ = str
End Function

Private Sub ConvertCase( _
                str As String, _
                ByRef CharacterCase As String, _
                ByVal RupeeName As String, _
                ByVal PaisaName As String, _
                ByVal LacsName As String, _
                ByVal ZeroPaisaName As String _
                )
    Dim rupeesAndPaisaConnector As String
    Dim LastAndName As String
    rupeesAndPaisaConnector = "and"
    LastAndName = "and"
    
    RupeeName = LCase$(RupeeName)
    PaisaName = LCase$(PaisaName)
    LacsName = LCase$(LacsName)
    ZeroPaisaName = LCase$(ZeroPaisaName)
    
    Select Case LCase$(CharacterCase)
        Case "u"
            str = UCase$(str)
            RupeeName = UCase$(RupeeName)
            PaisaName = UCase$(PaisaName)
            LacsName = UCase$(LacsName)
            ZeroPaisaName = UCase$(ZeroPaisaName)
            rupeesAndPaisaConnector = "AND"
            LastAndName = "AND"
        Case "l"
            str = LCase$(str)
        Case "t"
            Mid$(str, 1, 1) = UCase$(Mid$(str, 1, 1))
            Dim spacePos As Long
            spacePos = InStr(1, str, " ")
            Do While spacePos <> 0
                Mid$(str, spacePos + 1, 1) = UCase$(Mid$(str, spacePos + 1, 1))
                spacePos = InStr(spacePos + 1, str, " ")
            Loop
            RupeeName = PCase$(RupeeName)
            PaisaName = PCase$(PaisaName)
            LacsName = PCase$(LacsName)
            ZeroPaisaName = PCase$(ZeroPaisaName)
    End Select
    
    str = Replace(str, RupeeConst, RupeeName)
    str = Replace(str, PaisaConst, PaisaName)
    str = Replace(str, RspConnectorConst, rupeesAndPaisaConnector)
    str = Replace(str, LacsConst, LacsName)
    str = Replace(str, ZeroPaisaConst, ZeroPaisaName)
    str = Replace(str, AndConst, LastAndName)
    
    If LCase$(CharacterCase) = "s" Then
        str = PCase$(str)
    End If
End Sub

Private Function ConvertToWords(ByVal Value As Currency, Optional ShowPaisa As Boolean = False, Optional AddCommas As Boolean = True, Optional LastAnd As Boolean = False) As String
    Dim result As String
    
    
    Dim Units(1 To 19) As String
    Dim Tens(2 To 9) As String
    Dim isNegative As Boolean
    Dim suppressLastAnd As Boolean
    
    If Value < 0 Then
        Value = Abs(Value)
        isNegative = True
    End If
    
    If Value < 100 Then
        suppressLastAnd = True
    End If
    
    ' Initialize
    Units(1) = "one"
    Units(2) = "two"
    Units(3) = "three"
    Units(4) = "four"
    Units(5) = "five"
    Units(6) = "six"
    Units(7) = "seven"
    Units(8) = "eight"
    Units(9) = "nine"
    Units(10) = "ten"
    Units(11) = "eleven"
    Units(12) = "twelve"
    Units(13) = "thirteen"
    Units(14) = "fourteen"
    Units(15) = "fifteen"
    Units(16) = "sixteen"
    Units(17) = "seventeen"
    Units(18) = "eighteen"
    Units(19) = "nineteen"
    
    Tens(2) = "twenty"
    Tens(3) = "thirty"
    Tens(4) = "forty"
    Tens(5) = "fifty"
    Tens(6) = "sixty"
    Tens(7) = "seventy"
    Tens(8) = "eighty"
    Tens(9) = "ninety"
    
    Dim PaisePart As Long
    
    PaisePart = CLng((Value@ - Fix(Value@)) * 100@)
    
    ' Paise part cannot be greater than 99
    If PaisePart > 99 Then
        PaisePart = 99
    End If
    
    Value = Fix(Value)
    
    If Value >= 10000000 Then
        Dim CrorePlaces As Currency
        CrorePlaces = Fix(CCur(Value@ / 10000000@))
        
        result = result & ConvertToWords(CrorePlaces) & " crore"
        If AddCommas Then
            result = result & ","
        End If
        result = result & " "
        Value = Value - (10000000@ * CrorePlaces)
    End If
    
    If Value \ 100000 > 0 Then
        result = result & ConvertToWords(Value \ 100000) & " " & LacsConst
        If AddCommas Then
            result = result & ","
        End If
        result = result & " "
        Value = Value Mod 100000@
    End If
    
    If Value \ 1000 > 0 Then
        result = result & ConvertToWords(Value \ 1000) & " thousand"
        If AddCommas Then
            result = result & ","
        End If
        result = result & " "
        Value = Value Mod 1000
    End If
    If Value \ 100 > 0 Then
        result = result & ConvertToWords(Value \ 100) & " hundred "
        Value = Value Mod 100
    End If

    Dim lastAndAdded As Boolean
    lastAndAdded = False
    
    If Value > 19 Then
        If LastAnd Then
            If result <> "" Then
                result = result & AndConst & " "
                lastAndAdded = True
            End If
        End If
        result = result & Tens(Value \ 10) & " "
        Value = Value Mod 10
    End If
    
    If Value > 0 Then
        If LastAnd Then
            If result <> "" Then
                If (Not lastAndAdded) And (Not suppressLastAnd) Then
                    result = result & AndConst & " "
                    lastAndAdded = True
                End If
            End If
        End If
        result = result & Units(Value) & " "
    End If
    
    ' Add the and word
    If ShowPaisa Then

        result = result & RspConnectorConst & " "
        
        If PaisePart > 0 Then
            result = result & ConvertToWords(PaisePart)
        Else
            result = result & ZeroPaisaConst
        End If
    End If

    If isNegative Then
        result = "minus " & result
    End If
    
    result = RTrim$(result)
    
    
    ConvertToWords = result

End Function

Public Function AmountToWords( _
        ByVal Value As Currency, _
        Optional AddCommas As Boolean = True, _
        Optional AddOnly As Boolean = True, _
        Optional ShowRupeeCaption As Boolean = True, _
        Optional IncludePaisa As Boolean = True, _
        Optional LastAnd As Boolean = True, _
        Optional RupeesAfter As Boolean = False, _
        Optional PaisaAfter As Boolean = True, _
        Optional CharacterCase As String = "s", _
        Optional RupeeName As String = "rupees", _
        Optional PaisaName As String = "paisa", _
        Optional LacsName As String = "lakhs", _
        Optional ZeroPaisaName As String = "zero" _
        ) As Variant
'ConvertNumberToWords

    Dim result As String
    
    Dim paisaPosition As Long
    paisaPosition = 0
    
    ' Check for invalid value
    Select Case LCase$(CharacterCase)
        Case "s", "u", "l", "t"
            ' Okay
        Case Else
            AmountToWords = CVErr(2015) ' #VALUE
            Exit Function
    End Select
    
    If Value = 0 Then
        result = "zero "
        If IncludePaisa Then
            paisaPosition = 6
            result = result & RspConnectorConst & " " & ZeroPaisaConst
        Else
            paisaPosition = 0
        End If
        AmountToWords = Decorate( _
                                result, _
                                    paisaPosition, _
                                    ShowRupeeCaption, _
                                    RupeeName, _
                                    IncludePaisa, _
                                    PaisaName, _
                                    CharacterCase, _
                                    RupeesAfter, _
                                    PaisaAfter, _
                                    AddOnly, _
                                    LacsName, _
                                    ZeroPaisaName _
                                )
        Exit Function
    End If
    
    result = ConvertToWords(Value, IncludePaisa, AddCommas, LastAnd)
    paisaPosition = InStr(1, result, RspConnectorConst)
    If paisaPosition > 0 Then
        paisaPosition = paisaPosition
    End If
    
    result = Decorate( _
                    result, _
                    paisaPosition, _
                    ShowRupeeCaption, _
                    RupeeName, _
                    IncludePaisa, _
                    PaisaName, _
                    CharacterCase, _
                    RupeesAfter, _
                    PaisaAfter, _
                    AddOnly, _
                    LacsName, _
                    ZeroPaisaName _
            )
    
    AmountToWords = result
End Function

Public Function RUPEESTEXT(ByVal number As Currency) As Variant
    RUPEESTEXT = AmountToWords(Value:=number, CharacterCase:="T")
End Function




