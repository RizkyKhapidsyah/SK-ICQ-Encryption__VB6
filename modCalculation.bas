Attribute VB_Name = "modCalculation"
Function hMul(Hex1$, Hex2$) As String
    Hex1Len = Len(Hex1$)
    Hex2Len = Len(Hex2$)

    For i = Hex2Len To 1 Step -1
        Temp1 = Val("&H" + Mid$(Hex2$, i, 1))
        Temp4$ = "0"
        Temp5$ = ""
        For j = Hex1Len To 1 Step -1
            Temp2$ = Hex$(Temp1 * ("&H" + Mid$(Hex1$, j, 1)) + Val("&H" + Temp4$))
            Temp3$ = Right$(Temp2$, 1)
            Temp4$ = Left$(Temp2$, Len(Temp2$) - 1)
            Temp5$ = Temp3$ + Temp5$
        Next j
        Temp5$ = Temp4$ + Temp5$ + String$(Hex2Len - i, "0")
        TempOut$ = hAdd(Temp5$, TempOut$)
    Next i
    hMul = TempOut$
End Function

Function hAdd(Hex1$, Hex2$) As String
    Hex1Len = Len(Hex1$)
    Hex2Len = Len(Hex2$)
    If Hex2Len > Hex1Len Then Hex1$ = String$(Hex2Len - Hex1Len, "0") + Hex1$
    If Hex1Len > Hex2Len Then Hex2$ = String$(Hex1Len - Hex2Len, "0") + Hex2$
    HexLen = Len(Hex1$)
    
    Temp4$ = "0"
    For i = HexLen To 1 Step -1
        Temp1$ = Mid$(Hex1$, i, 1)
        Temp2$ = Mid$(Hex2$, i, 1)
        Temp3$ = Hex(("&H" + Temp1$) + Val("&H" + Temp2$) + Val("&H" + Temp5$))
        Temp4$ = Right$(Temp3$, 1)
        Temp5$ = Left$(Temp3$, Len(Temp3$) - 1)
        TempOut$ = Temp4$ + TempOut$
    Next i
    TempOut$ = Temp5$ + TempOut$
    hAdd = TempOut$
End Function

Public Function BitShift(InpHex$, NoofBits, Direction As BitShiftDirection) As String
    Dim TempValue, Temp1$, Temp2, TempBinary$, TempLen, TempHex$, HexLen
    If Val("&H" + InpHex$) = 0 Then BitShift = "0": Exit Function
    
    HexLen = Len(InpHex$)
    
    TempBinary$ = Hex2Bin(InpHex$)
    '-- Start Shifting --
    If Direction = LeftShift Then
        TempBinary$ = TempBinary$ + String$(NoofBits, "0")
        TempLen = Len(TempBinary$)
        TempBinary$ = String$(4 - (TempLen Mod 4), "0") + TempBinary$
    Else
        FromLeft = NoofBits
        FromRight = Len(TempBinary$) - NoofBits
        TempBinary$ = String$(FromLeft, "0") + Left$(TempBinary$, FromRight)
    End If

    BitShift = Right$(Bin2Hex(TempBinary$), HexLen)
End Function

Function hAnd(Hex1$, Hex2$) As String
    Dim Bin1$, Bin2$, Val1, Val2, TempOut$
    Bin1$ = Hex2Bin(Hex1$)
    Bin2$ = Hex2Bin(Hex2$)
    
    If Len(Bin1$) < Len(Bin2$) Then
        Bin1$ = String$(Len(Bin2$) - Len(Bin1$), "0") + Bin1$
    Else
        Bin2$ = String$(Len(Bin1$) - Len(Bin2$), "0") + Bin2$
    End If

    TempOut$ = ""
    For i = 1 To Len(Bin1$)
        Val1 = Val(Mid$(Bin1$, i, 1))
        Val2 = Val(Mid$(Bin2$, i, 1))
        
        If Val1 = 1 And Val2 = 1 Then
            TempOut$ = TempOut$ + "1"
        Else
            TempOut$ = TempOut$ + "0"
        End If
    Next i
    
    hAnd = Bin2Hex(TempOut$)
End Function

Function hXor(Hex1$, Hex2$) As String
    Dim Bin1$, Bin2$, Val1, Val2, TempOut$
    Bin1$ = Hex2Bin(Hex1$)
    Bin2$ = Hex2Bin(Hex2$)
    If Len(Bin1$) < Len(Bin2$) Then
        Bin1$ = String$(Len(Bin2$) - Len(Bin1$), "0") + Bin1$
    Else
        Bin2$ = String$(Len(Bin1$) - Len(Bin2$), "0") + Bin2$
    End If

    TempOut$ = ""
    For i = 1 To Len(Bin1$)
        Val1 = Val(Mid$(Bin1$, i, 1))
        Val2 = Val(Mid$(Bin2$, i, 1))
        
        If Val1 <> Val2 Then
            TempOut$ = TempOut$ + "1"
        Else
            TempOut$ = TempOut$ + "0"
        End If
    Next i
    hXor = Bin2Hex(TempOut$)
End Function

Function hOr(Hex1$, Hex2$) As String
    Dim Bin1$, Bin2$, Val1, Val2, TempOut$
    Bin1$ = Hex2Bin(Hex1$)
    Bin2$ = Hex2Bin(Hex2$)
    If Len(Bin1$) < Len(Bin2$) Then
        Bin1$ = String$(Len(Bin2$) - Len(Bin1$), "0") + Bin1$
    Else
        Bin2$ = String$(Len(Bin1$) - Len(Bin2$), "0") + Bin2$
    End If

    TempOut$ = ""
    For i = 1 To Len(Bin1$)
        Val1 = Val(Mid$(Bin1$, i, 1))
        Val2 = Val(Mid$(Bin2$, i, 1))
        
        If Val1 = 1 Or Val2 = 1 Then
            TempOut$ = TempOut$ + "1"
        Else
            TempOut$ = TempOut$ + "0"
        End If
    Next i
    
    hOr = Bin2Hex(TempOut$)
End Function

Function hNot(Hex1$) As String
    Dim Bin1$, Val1, TempOut$
    Bin1$ = Hex2Bin(Hex1$)

    TempOut$ = ""
    For i = 1 To Len(Bin1$)
        Val1 = Val(Mid$(Bin1$, i, 1))
        
        If Val1 = 0 Then
            TempOut$ = TempOut$ + "1"
        Else
            TempOut$ = TempOut$ + "0"
        End If
    Next i
    
    hNot = Bin2Hex(TempOut$)
End Function
