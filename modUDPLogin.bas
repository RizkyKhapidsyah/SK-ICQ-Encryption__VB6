Attribute VB_Name = "modUDP"
Option Explicit
Public Const HEADER_ZERO = "00000000"

Public Const ICQTable$ = _
    "5960376B6562464853614C5960575B3D5E346D36503F6F6753614C5940476339" + _
    "505F5F3F6F47436948333164355A4A425640675341076C49583B4D4668436948" + _
    "333144656246485341076C69483351545D4E6C49384B554A6246483351346D36" + _
    "505F5F5F3F6F4763594067333164355A6A526E3C51346D36505F5F3F4F374B35" + _
    "5A4A6266583B4D66585B5D4E6C49583B4D66583B4D464853614C594067333164" + _
    "556A323E4445526E3C3164556A524E6C694853614C39306F47635960575B3D3E" + _
    "64353A3A5A6A524E6C694853616C49583B4D46686339505F5F3F6F6753412541" + _
    "3C51543D5E545D4E4C39505F5F5F3F6F474369483351545D6E3C3164355A0000"
    
'Public Const ICQTableV4$ = _
    "0A5B315D20596F752063616E206D6F646966792074686520736F756E647320494351206D616B65732E204A7573742073" + _
    "656C6563742022536F756E6473222066726F6D207468652022707265666572656E6365732F6D6973632220696E20494351206F722066726F6D20746865202253" + _
    "6F756E64732220696E2074686520636F6E74726F6C2070616E656C2E204372656469743A204572616E0A5B325D2043616E27742072656D656D62657220776861" + _
    "742077617320736169643F2020446F75626C652D636C69636B206F6E2061207573657220746F206765742061206469616C6F67206F6620616C6C206D65737361" + _
    "6765732073656E7420696E636F6D696E"

Function CreatePacket(Header As ClientHeader) As String
    Dim Packet$, PacketSeqNum1$, PacketSeqNum2$, TempCheckCode$
        
    PacketSeqNum1$ = Dec2Hex(Header.SeqNum1, tInt)
    PacketSeqNum2$ = Dec2Hex(Header.SeqNum2, tInt)
    TempCheckCode$ = "00000000"
    
    Header.SeqNum1 = Header.SeqNum1 + 1
    Header.SeqNum2 = Header.SeqNum2 + 1
    
    'Assemble the packet
    Packet$ = _
        INF_UDP_VERSION + _
        HEADER_ZERO + _
        Dec2Hex(Header.UIN, tLong) + _
        Dec2Hex(Header.SessionID, tLong) + _
        Header.Command + _
        PacketSeqNum1$ + _
        PacketSeqNum2$ + _
        TempCheckCode$ + _
        Header.Parameter
    
    CreatePacket = EncryptPacket$(Packet$)
End Function



Function CryptPacket$(Packet$, CheckCode$)
    Dim PacketLength, Code1$, Code2$, Code3$, Pos, N, T, Temp$, Temp1$
    
    PacketLength = Int(Len(Packet$) / 2)
    
    Code1$ = hFill(hMul(Hex$(PacketLength), "68656C6C"), tLong)
    Code2$ = hFill(hAdd(Code1$, CheckCode$), tLong)
    
    Pos = &HA
    N = PacketLength
    Do While Pos < N
        T = Pos Mod &H100
        Code3$ = hFill(hAdd(Code2$, PeekByte(ICQTable$, T)), tLong)
        Temp$ = hDump(Peek(Packet$, Pos, tLong))
        Temp1$ = hXor(Temp$, Code3$)
        Temp1$ = hDump(Right$(Temp1$, 8))
        Packet$ = hFill(Poke(Packet$, Temp1$, Pos), PacketLength)
        Pos = Pos + 4
    Loop
    
    CryptPacket$ = Packet$
End Function

Function DecryptPacket$(Packet$)
    Dim TempCode$, Temp1$, Temp2$

    TempCode$ = Descramble$(Packet$)

    Packet$ = CryptPacket$(Packet$, TempCode$)
        
    DecryptPacket$ = Packet$
End Function

Function EncryptPacket$(Packet$)
    Dim TempCode$
    TempCode$ = CalcCheckCode$(Packet$)
    Packet$ = CryptPacket(Packet$, hDump(TempCode$))
    
    TempCode$ = Scramble$(hDump(TempCode$))
    Packet$ = Poke(Packet$, TempCode$, 20)
    
    EncryptPacket$ = Packet$
End Function

Function CalcCheckCode$(Packet$)
    Dim TempB2$, TempB4$, TempB6$, TempB8$, Number1$, _
        TempCode, PacketLength, R1, R2, _
        TempX4$, TempX3$, TempX2$, TempX1$, Number2$
    
    TempB2$ = PeekByte(Packet$, 2)
    TempB4$ = PeekByte(Packet$, 4)
    TempB6$ = PeekByte(Packet$, 6)
    TempB8$ = PeekByte(Packet$, 8)
    Number1$ = TempB8$ + TempB4$ + TempB2$ + TempB6$
 
    PacketLength = Len(Packet$) / 2
    Randomize Timer
    R1 = Abs(&H18 + Int(Rnd(Timer) * ((PacketLength - &H19) Mod &H100)))
    R2 = Abs(Int(Rnd(Timer) * &HFF))
     
    TempX4$ = HexB$(R1)
    TempX3$ = HexB$(Not (CByte(Val("&H" + PeekByte(Packet$, R1)))))
    TempX2$ = HexB$(R2)
    TempX1$ = HexB$(Not (CByte(Val("&H" + PeekByte(ICQTable$, R2)))))
    Number2$ = TempX4$ + TempX3$ + TempX2$ + TempX1$
     
    TempCode = Val("&H" + Number1$) Xor Val("&H" + Number2$)
    CalcCheckCode$ = hDump(Dec2Hex(TempCode, tLong))
End Function

Function Scramble$(CheckCode$)
    Dim a0$, a1$, a2$, a3$, a4$, TempOut$
    
    a0$ = hAnd(CheckCode$, "1F")
    a1$ = hAnd(CheckCode$, "3E003E0")
    a2$ = hAnd(CheckCode$, "F8000400")
    a3$ = hAnd(CheckCode$, "F800")
    a4$ = hAnd(CheckCode$, "41F0000")
    
    a0$ = BitShift(a0$, &HC, LeftShift)
    a1$ = BitShift(a1$, &H1, LeftShift)
    a2$ = BitShift(a2$, &HA, RightShift)
    a3$ = BitShift(a3$, &H10, LeftShift)
    a4$ = BitShift(a4$, &HF, RightShift)
    
    TempOut$ = "00"
    TempOut$ = hAdd(TempOut$, a0$)
    TempOut$ = hAdd(TempOut$, a1$)
    TempOut$ = hAdd(TempOut$, a2$)
    TempOut$ = hAdd(TempOut$, a3$)
    TempOut$ = hAdd(TempOut$, a4$)
    
    Scramble$ = hDump(TempOut$)
End Function

Function Descramble$(Packet$)
    Dim CheckCode$, a0$, a1$, a2$, a3$, a4$, TempOut$
    CheckCode$ = hDump(Peek(Packet$, &H14, tLong))
    
    a0$ = hAnd(CheckCode$, "1F000")
    a1$ = hAnd(CheckCode$, "7C007C0")
    a2$ = hAnd(CheckCode$, "3E0001")
    a3$ = hAnd(CheckCode$, "F8000000")
    a4$ = hAnd(CheckCode$, "83E")
    
    a0$ = BitShift(a0$, &HC, RightShift)
    a1$ = BitShift(a1$, &H1, RightShift)
    a2$ = BitShift(a2$, &HA, LeftShift)
    a3$ = BitShift(a3$, &H10, RightShift)
    a4$ = BitShift(a4$, &HF, LeftShift)

    TempOut$ = "00"
    TempOut$ = hAdd(TempOut$, a0$)
    TempOut$ = hAdd(TempOut$, a1$)
    TempOut$ = hAdd(TempOut$, a2$)
    TempOut$ = hAdd(TempOut$, a3$)
    TempOut$ = hAdd(TempOut$, a4$)
    
    Descramble$ = hFill(TempOut$, tLong)
End Function
