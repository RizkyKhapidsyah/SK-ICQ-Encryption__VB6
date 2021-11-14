Attribute VB_Name = "modTypeConversion"
Public Const HexBinary = _
    "0000" + "0001" + "0010" + "0011" + _
    "0100" + "0101" + "0110" + "0111" + _
    "1000" + "1001" + "1010" + "1011" + _
    "1100" + "1101" + "1110" + "1111"
    
Function Hex2Bin(InpHex$) As String
    '-- Convert to Binary First --
    For i = 1 To Len(InpHex$)
        TempValue = Val("&H" + Mid$(InpHex$, i, 1))
        Temp1$ = Peek(HexBinary, TempValue * 2, 2)
        TempBinary$ = TempBinary$ + Temp1$
    Next i
    If Len(TempBinary$) = 0 Then TempBinary$ = "0000"
    Hex2Bin = TempBinary$
End Function

Function Bin2Hex(InpBin$) As String
    '-- Convert back to Hex --
    For i = 1 To Len(InpBin$) Step 4
        Temp1$ = Mid$(InpBin$, i, 4)
        TempValue = 0
        
        If Val(Temp1$) > 1 Then
            For j = 1 To 4
                Temp2 = Val(Mid$(Temp1$, j, 1))
                TempValue = TempValue + ((2 ^ (4 - j)) * Temp2)
            Next j
        Else
            TempValue = Val(Temp1$)
        End If
            
        TempHex$ = TempHex$ + Hex$(TempValue)
    Next i
    
    If Len(TempHex$) = 0 Then TempHex$ = "0"
    
    Bin2Hex = TempHex$
End Function

Function HexB$(InpHex)
    Temp1$ = Hex$(InpHex)
    If Len(Temp1$) Mod 2 = 1 Then Temp1$ = "0" + Temp1$
    HexB$ = Temp1$
End Function

Function Dec2Hex(InpNumber, OutType As VarType) As String
    'This module is to output a Hexadecimal string.
    On Error GoTo ErrorProc
    Select Case OutType
        Case tByte: Temp0 = CByte(InpNumber)
        Case tInt:  Temp0 = CInt(InpNumber)
        Case tLong: Temp0 = CLng(InpNumber)
    End Select
    
    Temp1$ = Hex$(Temp0)
    Temp1$ = String$((OutType * 2) - Len(Temp1$), "0") + Temp1$
    If OutType = tInt Or OutType = tLong Then Temp1$ = hDump(Temp1$)
    
    Dec2Hex = Temp1$
    
ErrorProc:
End Function

Function Hex2Str(InpHex$) As String
    'Convert Hex strings (eg. "00123ADBBE") to Character strings
    For i = 1 To Len(InpHex$) Step 2
        Temp1$ = Chr$(Val("&H" + Mid$(InpHex$, i, 2)))
        Temp2$ = Temp2$ + Temp1$
    Next i
    Hex2Str = Temp2$
End Function

Function hDump(InpHex$) As String
    If Len(InpHex$) Mod 2 = 1 Then InpHex$ = "0" + InpHex$
    Temp$ = ""
    For i = 1 To Len(InpHex$) Step 2
        Temp$ = Mid$(InpHex$, i, 2) + Temp$
    Next i
    
    hDump = Temp$
End Function

Function Str2Hex(InpStr$) As String
    'Convert Character Strings to Hex strings
    For i = 1 To Len(InpStr$)
        Temp1$ = HexB$(Asc(Mid$(InpStr$, i, 1)))
        Temp2$ = Temp2$ + Temp1$
    Next i
    Str2Hex = Temp2$
End Function

Function IP2Hex$(InpIP$)
    If Len(InpIP$) = 0 Then IP2Hex$ = "00000000": Exit Function
    TempDot1 = InStr(1, InpIP$, ".", vbBinaryCompare)
    TempDot2 = InStr(TempDot1 + 1, InpIP$, ".", vbBinaryCompare)
    TempDot3 = InStr(TempDot2 + 1, InpIP$, ".", vbBinaryCompare)
    TempDot4 = Len(InpIP$)
       
    TempIP1$ = Left$(InpIP$, TempDot1 - 1)
    TempIP2$ = Mid$(InpIP$, TempDot1 + 1, TempDot2 - TempDot1 - 1)
    TempIP3$ = Mid$(InpIP$, TempDot2 + 1, TempDot3 - TempDot2 - 1)
    TempIP4$ = Right$(InpIP$, Len(InpIP$) - TempDot3)
    
    'Debug.Print Str$(TempDot1); " "; Str$(TempDot2); " "; Str$(TempDot3); " "; Str$(TempDot4); " ";
    'Debug.Print TempIP1$; " "; TempIP2$; " "; TempIP3$; " "; TempIP4$;
    
    IP2Hex$ = HexB$(TempIP1$) + HexB$(TempIP2$) + HexB$(TempIP3$) + HexB$(TempIP4$)
    
End Function

Function Hex2Ip$(InpHex$)
    Dim IP$(4)
    For i = 1 To 4
        Temp1$ = Mid$(InpHex$, (i * 2) - 1, 2)
        Temp2$ = Trim(Str$(Val("&H" + Temp1$)))
        Temp3$ = Temp3$ + Temp2$ + "."
    Next i
    Hex2Ip$ = Left$(Temp3$, Len(Temp3$) - 1)
End Function

Function PeekByte(Data$, Location) As String
    PeekByte = Mid$(Data$, Location * 2 + 1, 2)
End Function

Function Peek(Data$, Location, NumberofBytes) As String
    Peek = ""
    For i = Location To Location + NumberofBytes - 1
        Peek = Peek + PeekByte(Data$, i)
    Next i
End Function

Function Poke(Data$, PokeInp$, Location) As String
    If Len(PokeInp$) = 0 Then
        Poke = Data$
        Exit Function
    End If
    targetlen = Location * 2 + Len(PokeInp$)
    CurrentLen = Len(Data$)
    'If CurrentLen < targetlen Then Data$ = Data$ + String$(targetlen - CurrentLen, "0")
    Mid$(Data$, Location * 2 + 1) = UCase$(PokeInp$)
    Poke = Data$
End Function

Function hFill(InpHex$, OutByte) As String
    targetlen = OutByte * 2
    CurLen = Len(InpHex$)
    
    If CurLen = targetlen Then hFill = InpHex$: Exit Function
    If CurLen < targetlen Then
        hFill = String$(targetlen - CurLen, "0") + InpHex$
    Else
        hFill = Right$(InpHex$, targetlen)
    End If
End Function
