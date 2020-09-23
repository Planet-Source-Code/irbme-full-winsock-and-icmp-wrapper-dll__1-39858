Attribute VB_Name = "modWinsock"
Option Explicit

'Winsock Initialization and termination
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

'String functions
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'Socket Functions
Private Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function WSACloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

'Data transfer functions
Private Declare Function WSARecv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Private Declare Function WSASend Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

'Network byte ordering functions
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

'End point information
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByRef namelen As Long) As Long
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByRef namelen As Long) As Long

'Hostname resolving functions
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

'---------------------------------------------------------------------------

'Winsock messages that will go to the window handler
Public Enum WSAMessage
    FD_READ = &H1&      'Data is ready to be read from the buffer
    FD_CONNECT = &H10&  'Connection esatblished
    FD_CLOSE = &H20&    'Connection closed
    FD_ACCEPT = &H8&    'Connection request pending
End Enum

'---------------------------------------------------------------------------

Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129

'Winsock Data structure
Public Type WSAData
    wVersion       As Integer                       'Version
    wHighVersion   As Integer                       'High Version
    szDescription  As String * WSADESCRIPTION_LEN   'Description
    szSystemStatus As String * WSASYS_STATUS_LEN    'Status of system
    iMaxSockets    As Integer                       'Maximum number of sockets allowed
    iMaxUdpDg      As Integer                       'Maximum UDP datagrams
    lpVendorInfo   As Long                          'Vendor Info
End Type

'HostEnt Structure
Public Type HOSTENT
    hName     As Long       'Host Name
    hAliases  As Long       'Alias
    hAddrType As Integer    'Address Type
    hLength   As Integer    'Length
    hAddrList As Long       'Address List
End Type

'Socket Address structure
Public Type SOCKADDR_IN
    sin_family       As Integer 'Address familly
    sin_port         As Integer 'Port
    sin_addr         As Long    'Long address
    sin_zero(1 To 8) As Byte
End Type

'---------------------------------------------------------------------------

'Windows Socket types
Private Const SOCK_STREAM = 1     'Stream socket

'Address family
Public Const AF_INET = 2          'Internetwork: UDP, TCP, etc.

'Socket Protocol
Private Const IPPROTO_TCP = 6     'tcp

'Data type conversion constants
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

'---------------------------------------------------------------------------

Public WSAStarted As Boolean

'---------------------------------------------------------------------------

Public Function UnsignedToLong(Value As Double) As Long
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Converts an unsigned double value into a long value.
'Arguments :Value  - The unsigned double to convert
'Returns   :The converted long value
'********************************************************************************

    If Value < 0 Or Value >= OFFSET_4 Then Error 6  'Overflow

    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If
    
End Function


Public Function LongToUnsigned(Value As Long) As Double
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Converts a long value into an unsigned double value
'Arguments :Value  - The long to convert
'Returns   :The converted unsigned double value
'********************************************************************************

    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If
    
End Function


Public Function UnsignedToInteger(Value As Long) As Integer
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Converts an unsigned long value into an integer
'Arguments :Value  - The unsigned long to convert
'Returns   :The converted integer value
'********************************************************************************

    If Value < 0 Or Value >= OFFSET_2 Then Error 6  'Overflow
    
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If

End Function


Public Function IntegerToUnsigned(Value As Integer) As Long
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Converts an integer value into an unsigned long
'Arguments :Value  - The integer value to convert
'Returns   :The converted unsigned long value
'********************************************************************************

    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
    
End Function


Public Function StringFromPointer(ByVal lngPointer As Long) As String
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Retrieves the string value given a pointer to it
'Arguments :lPointer  - A pointer to the string
'Returns   :The string value stored at the pointer - lngPointer
'********************************************************************************

  Dim strTemp As String
  Dim lRetVal As Long
    
    strTemp = String$(lstrlen(ByVal lngPointer), 0)    'prepare the strTemp buffer
    lRetVal = lstrcpy(ByVal strTemp, ByVal lngPointer) 'copy the string into the strTemp buffer
    If lRetVal Then StringFromPointer = strTemp        'return the string

End Function


Public Function HiWord(lngValue As Long) As Long
'********************************************************************************
'Date      :15 October 2002
'Purpose   :Retrieves the HIWord from a long value
'Arguments :lngValue  - The long value
'Returns   :The HIWord of the long value
'********************************************************************************

    If (lngValue And &H80000000) = &H80000000 Then
        HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        HiWord = (lngValue And &HFFFF0000) \ &H10000
    End If
    
End Function


Public Function mCreateSocket() As Long
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Creates a new socket to be used with the other function calls
'Arguments :VOID
'Returns   :If no error then the socket handle, else INVALID_SOCKET is returned
'********************************************************************************

    'Call the socket Winsock API function in order create a new socket
    mCreateSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    'Force the Winsock service to send the network event notifications
    'to the window handler
    Call WSAAsyncSelect(mCreateSocket, WindowHandle, WinsockMessage, FD_CONNECT Or FD_READ Or FD_CLOSE Or FD_ACCEPT)

End Function


Public Function mRecv(ByVal lngSocket As Long, strBuffer As String) As Long
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Recieves data from the Winsock buffer of a specific port
'Arguments :lngSocket - The socket handle to read from
'           strBuffer - The string buffer to place the data into
'Returns   :The number of bytes read
'********************************************************************************

  Const MAX_BUFFER_LENGTH As Long = 8192

  Dim arrBuffer(1 To MAX_BUFFER_LENGTH)   As Byte
  Dim lngBytesReceived                    As Long
  Dim strTempBuffer                       As String
    
    'Call the recv Winsock API function in order to read data from the buffer
    lngBytesReceived = WSARecv(lngSocket, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)

    If lngBytesReceived > 0 Then
        'If we have received some data, convert it to the Unicode
        'string that is suitable for the Visual Basic String data type
        strTempBuffer = StrConv(arrBuffer, vbUnicode)

        'Remove unused bytes
        strBuffer = Left$(strTempBuffer, lngBytesReceived)
    End If
        
    mRecv = lngBytesReceived

End Function


Public Function mSend(ByVal lngSocket As Long, strData As String) As Long
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Sends data to the Winsock buffer of a specific port
'Arguments :lngSocket - The socket handle to send data to
'           strData - The string buffer to send
'Returns   :The number of bytes written to the buffer
'********************************************************************************

  Dim arrBuffer()     As Byte

    'Convert the data string to a byte array
    arrBuffer() = StrConv(strData, vbFromUnicode)
    'Call the send Winsock API function in order to send data
    mSend = WSASend(lngSocket, arrBuffer(0), Len(strData), 0&)

End Function

