Attribute VB_Name = "modSend"
Option Explicit

Public Const LOGIN_PACKET_SIZE = 47
Public Const LOGIN_X1_DEF = "D5000000"
Public Const LOGIN_X2_DEF = "0000"
Public Const LOGIN_X3_DEF = "00000000"
Public Const LOGIN_X4_DEF = "0800D500"
Public Const LOGIN_X5_DEF = "50000000"
Public Const LOGIN_X6_DEF = "03000000"
Public Const LOGIN_X7_DEF = "00000000"

Public Enum LOGIN_PROXY_STATUS
    LOGIN_FIRWALL = 1
    LOGIN_SNDONLY_TCP = 2
    LOGIN_SNDRCV_TCP = 4
End Enum

Function SendLogin(Header As ClientHeader, TCP_port, ProxyStatus As LOGIN_PROXY_STATUS, OnlineStatus$) As String
    Dim TempTime1 As Long, ParamTime$, ParamTCPPort$, _
    ParamPassLen$, ParamPassword$, ParamIP$, ParamProxyStatus$
    
    Header.SessionID = CLng(Rnd(Timer) * &H3FFFFFFF)
    Header.SeqNum1 = CInt(Rnd(Timer) * &H7FFF)
    Header.SeqNum2 = 1
    Header.Command = CMD_LOGIN
    
    TempTime1 = DateDiff("d", "1-1-1971", Now()) * 24 * 60 * 60
    TempTime1 = TempTime1 + Timer
    ParamTime$ = Dec2Hex(TempTime1, tLong)
    
    ParamTCPPort$ = Dec2Hex(TCP_port, tLong)
    ParamPassLen$ = Dec2Hex(Len(Header.Password) + 1, tInt)
    ParamPassword$ = Str2Hex(Header.Password + vbNullChar)
    ParamProxyStatus$ = Dec2Hex(ProxyStatus, tByte)
    ParamIP$ = IP2Hex$(IPAddr)
        
    Owner.Parameter = _
        ParamTime$ + _
        ParamTCPPort$ + _
        ParamPassLen$ + _
        ParamPassword$ + _
        LOGIN_X1_DEF + _
        ParamIP$ + _
        ParamProxyStatus$ + _
        OnlineStatus$ + _
        INF_TCP_VERSION + _
        LOGIN_X2_DEF + _
        LOGIN_X3_DEF + _
        LOGIN_X4_DEF + _
        LOGIN_X5_DEF + _
        LOGIN_X6_DEF + _
        LOGIN_X7_DEF
                        
    SendLogin = CreatePacket(Owner)
End Function

Function SendACK(SvrHeader As ServerHeader) As String
    Dim TempHeader As ClientHeader
    TempHeader.Command = CMD_ACK
    TempHeader.SeqNum1 = SvrHeader.SeqNum1
    TempHeader.SeqNum2 = SvrHeader.SeqNum2
    TempHeader.UIN = Owner.UIN
    TempHeader.SessionID = Owner.SessionID
    TempHeader.Parameter = Dec2Hex(CLng(Rnd(Timer) * &H7FFFFFFF), tLong)
    
    SendACK = CreatePacket(TempHeader)
End Function


