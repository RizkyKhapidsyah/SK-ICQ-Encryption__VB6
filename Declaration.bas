Attribute VB_Name = "Declaration"
Enum VarType
    tByte = 1: tInt = 2: tLong = 4
End Enum

Enum BitShiftDirection
    RightShift = 1: LeftShift = -1
End Enum

Public Type ClientHeader
    UIN As Long
    Password As String
    SessionID As Long
    Command As String * 4
    SeqNum1 As Integer
    SeqNum2 As Integer
    Parameter As String
End Type

Public Type ServerHeader
    SessionID As Long
    Command As String * 4
    SeqNum1 As Integer
    SeqNum2 As Integer
    UIN As Long
    Parameter As String
End Type

'   --- Constant Info
Public Const INF_UDP_VERSION = "0500"
Public Const INF_TCP_VERSION = "0600"
Public Const V5_HEADER_SIZE = "24"


Public Const CMD_ACK = "0A00"
Public Const CMD_SEND_MESSAGE = "0E01"
Public Const CMD_LOGIN = "E803"
Public Const CMD_REG_NEW_USER = "FC03"
Public Const CMD_CONTACT_LIST = "0604"
Public Const CMD_SEARCH_UIN = "1A04"
Public Const CMD_SEARCH_USER = "2404"
Public Const CMD_KEEP_ALIVE = "2E04"
Public Const CMD_SEND_TEXT_CODE = "3804"
Public Const CMD_ACK_MESSAGES = "4204"
Public Const CMD_LOGIN_1 = "4C04"
Public Const CMD_MSG_TO_NEW_USER = "5604"
Public Const CMD_INFO_REQ = "6004"
Public Const CMD_EXT_INFO_REQ = "6A04"
Public Const CMD_CHANGE_PW = "9C04"
Public Const CMD_NEW_USER_INFO = "A604"
Public Const CMD_UPDATE_EXT_INFO = "B004"
Public Const CMD_QUERY_SERVERS = "BA04"
Public Const CMD_QUERY_ADDONS = "C404"
Public Const CMD_STATUS_CHANGE = "D804"
Public Const CMD_NEW_USER_1 = "EC04"
Public Const CMD_UPDATE_INFO = "0A05"
Public Const CMD_AUTH_UPDATE = "1405"
Public Const CMD_KEEP_ALIVE2 = "1E05"
Public Const CMD_LOGIN_2 = "2805"
Public Const CMD_ADD_TO_LIST = "3C05"
Public Const CMD_RAND_SET = "6405"
Public Const CMD_RAND_SEARCH = "6E05"
Public Const CMD_META_USER = "4A06"
Public Const CMD_INVIS_LIST = "A406"
Public Const CMD_VIS_LIST = "AE06"
Public Const CMD_UPDATE_LIST = "B806"

Public Const SRV_ACK = "0A00"
Public Const SRV_LOGIN_REPLY = "5A00"
Public Const SRV_USER_ONLINE = "6E00"
Public Const SRV_USER_OFFLINE = "7800"
Public Const SRV_USER_FOUND = "8C00"
Public Const SRV_RECV_MESSAGE = "DC00"
Public Const SRV_END_OF_SEARCH = "A000"
Public Const SRV_INFO_REPLY = "1801"
Public Const SRV_EXT_INFO_REPLY = "2201"
Public Const SRV_STATUS_UPDATE = "A401"
Public Const SRV_X1 = "1C02"
Public Const SRV_X2 = "E600"
Public Const SRV_UPDATE_EXT = "C800"
Public Const SRV_NEW_UIN = "4600"
Public Const SRV_NEW_USER = "B400"
Public Const SRV_QUERY = "8200"
Public Const SRV_SYSTEM_MESSAGE = "C201"
Public Const SRV_SYS_DELIVERED_MESS = "0401"
Public Const SRV_GO_AWAY = "2800"
Public Const SRV_NOT_CONNECTED = "F000"
Public Const SRV_BAD_PASS = "6400"
Public Const SRV_TRY_AGAIN = "FA00"
Public Const SRV_UPDATE_FAIL = "EA01"
Public Const SRV_UPDATE_SUCCESS = "E001"
Public Const SRV_MULTI_PACKET = "1202"
Public Const SRV_META_USER = "DE03"
Public Const SRV_RAND_USER = "4E02"
Public Const SRV_AUTH_UPDATE = "F401"
 
Public Const META_INFO_SET = "E803"
Public Const META_INFO_REQ = "B004"
Public Const META_INFO_SECURE = "2404"
Public Const META_INFO_PASS = "2E04"
Public Const META_SRV_GEN = "C800"
Public Const META_SRV_MORE = "DC00"
Public Const META_SRV_WORK = "D200"
Public Const META_SRV_ABOUT = "E600"
Public Const META_SRV_PASS = "AA00"
Public Const META_SRV_GEN_UPDATE = "6400"
Public Const META_SRV_ABOUT_UPDATE = "8200"
Public Const META_SRV_OTHER_UPDATE = "7800"
Public Const META_INFO_ABOUT = "0604"
Public Const META_INFO_MORE = "FC03"

Public Const STATUS_ONLINE = "00000000"
Public Const STATUS_INVISIBLE = "00010000"
Public Const STATUS_NA = "04000000"
Public Const STATUS_OCCUPIED = "10000000"
Public Const STATUS_AWAY = "01000000"
Public Const STATUS_DND = "13000000"
Public Const STATUS_CHAT = "20000000"

Public Const AUTH_MESSAGE = "0800"

Public Const MSG_USER_ADDED = "0C00"
Public Const MSG_AUTH_REQ = "0600"
Public Const MSG_AUTH = "0800"
Public Const MSG_URL = "0400"
Public Const MSG_TEXT = "0100"
Public Const MSG_CONTACTS = "1300"

Public IPAddr As String
Public Owner As ClientHeader
