Attribute VB_Name = "modEncryptionTest"
Function EncText(Text$) As String
    Packet$ = String$(48, "0") + Str2Hex$(Text$)
    
    TempTxt$ = EncryptPacket$(Packet$)
    TempTxt$ = Right$(TempTxt$, Len(TempTxt$) - 40)
    TempTxt$ = Hex2Str$(TempTxt$)
    
    EncText = TempTxt$
End Function

Function DecText$(Packet$)
    Packet$ = Str2Hex$(Packet$)
    TempTxt = DecryptPacket$(String$(40, "0") + Packet$)
    DecText = Hex2Str$(Right$(TempTxt, Len(TempTxt) - 48))
End Function
