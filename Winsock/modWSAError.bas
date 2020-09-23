Attribute VB_Name = "modWSAError"
Option Explicit


'Basic Winsock error results.
Public Enum WSABaseErrors
    INADDR_NONE = &HFFFF
    SOCKET_ERROR = -1
    INVALID_SOCKET = -1
End Enum


'Winsock error offset
Private Const WSABASEERR = 10000

'Winsock error constants
Public Enum WSAErrorConstants

'Windows Sockets definitions of regular Microsoft C error constants
    WSAEINTR = (WSABASEERR + 4)
    WSAEBADF = (WSABASEERR + 9)
    WSAEACCES = (WSABASEERR + 13)
    WSAEFAULT = (WSABASEERR + 14)
    WSAEINVAL = (WSABASEERR + 22)
    WSAEMFILE = (WSABASEERR + 24)

'Windows Sockets definitions of regular Berkeley error constants
    WSAEWOULDBLOCK = (WSABASEERR + 35)
    WSAEINPROGRESS = (WSABASEERR + 36)
    WSAEALREADY = (WSABASEERR + 37)
    WSAENOTSOCK = (WSABASEERR + 38)
    WSAEDESTADDRREQ = (WSABASEERR + 39)
    WSAEMSGSIZE = (WSABASEERR + 40)
    WSAEPROTOTYPE = (WSABASEERR + 41)
    WSAENOPROTOOPT = (WSABASEERR + 42)
    WSAEPROTONOSUPPORT = (WSABASEERR + 43)
    WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
    WSAEOPNOTSUPP = (WSABASEERR + 45)
    WSAEPFNOSUPPORT = (WSABASEERR + 46)
    WSAEAFNOSUPPORT = (WSABASEERR + 47)
    WSAEADDRINUSE = (WSABASEERR + 48)
    WSAEADDRNOTAVAIL = (WSABASEERR + 49)
    WSAENETDOWN = (WSABASEERR + 50)
    WSAENETUNREACH = (WSABASEERR + 51)
    WSAENETRESET = (WSABASEERR + 52)
    WSAECONNABORTED = (WSABASEERR + 53)
    WSAECONNRESET = (WSABASEERR + 54)
    WSAENOBUFS = (WSABASEERR + 55)
    WSAEISCONN = (WSABASEERR + 56)
    WSAENOTCONN = (WSABASEERR + 57)
    WSAESHUTDOWN = (WSABASEERR + 58)
    WSAETOOMANYREFS = (WSABASEERR + 59)
    WSAETIMEDOUT = (WSABASEERR + 60)
    WSAECONNREFUSED = (WSABASEERR + 61)
    WSAELOOP = (WSABASEERR + 62)
    WSAENAMETOOLONG = (WSABASEERR + 63)
    WSAEHOSTDOWN = (WSABASEERR + 64)
    WSAEHOSTUNREACH = (WSABASEERR + 65)
    WSAENOTEMPTY = (WSABASEERR + 66)
    WSAEPROCLIM = (WSABASEERR + 67)
    WSAEUSERS = (WSABASEERR + 68)
    WSAEDQUOT = (WSABASEERR + 69)
    WSAESTALE = (WSABASEERR + 70)
    WSAEREMOTE = (WSABASEERR + 71)

'Extended Windows Sockets error constant definitions
    WSASYSNOTREADY = (WSABASEERR + 91)
    WSAVERNOTSUPPORTED = (WSABASEERR + 92)
    WSANOTINITIALISED = (WSABASEERR + 93)
    WSAEDISCON = (WSABASEERR + 101)
    WSAENOMORE = (WSABASEERR + 102)
    WSAECANCELLED = (WSABASEERR + 103)
    WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
    WSAEINVALIDPROVIDER = (WSABASEERR + 105)
    WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
    WSASYSCALLFAILURE = (WSABASEERR + 107)
    WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
    WSATYPE_NOT_FOUND = (WSABASEERR + 109)
    WSA_E_NO_MORE = (WSABASEERR + 110)
    WSA_E_CANCELLED = (WSABASEERR + 111)
    WSAEREFUSED = (WSABASEERR + 112)

    WSAHOST_NOT_FOUND = 11001
    WSATRY_AGAIN = 11002
    WSANO_RECOVERY = 11003
    WSANO_DATA = 11004

    FD_SETSIZE = 64
End Enum

'---------------------------------------------------------------------------

Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Gives a meaningful error description from a code
'Arguments :lngErrorCode - The error code
'Returns   :Returns a description of the error.
'********************************************************************************

  Dim strDesc As String
    
    Select Case lngErrorCode
        Case WSAEACCES
            strDesc = "Permission denied."
        Case WSAEADDRINUSE
            strDesc = "Address already in use."
        Case WSAEADDRNOTAVAIL
            strDesc = "Cannot assign requested address."
        Case WSAEAFNOSUPPORT
            strDesc = "Address family not supported by protocol family."
        Case WSAEALREADY
            strDesc = "Operation already in progress."
        Case WSAECONNABORTED
            strDesc = "Software caused connection abort."
        Case WSAECONNREFUSED
            strDesc = "Connection refused."
        Case WSAECONNRESET
            strDesc = "Connection reset by peer."
        Case WSAEDESTADDRREQ
            strDesc = "Destination address required."
        Case WSAEFAULT
            strDesc = "Bad address."
        Case WSAEHOSTDOWN
            strDesc = "Host is down."
        Case WSAEHOSTUNREACH
            strDesc = "No route to host."
        Case WSAEINPROGRESS
            strDesc = "Operation now in progress."
        Case WSAEINTR
            strDesc = "Interrupted function call."
        Case WSAEINVAL
            strDesc = "Invalid argument."
        Case WSAEISCONN
            strDesc = "Socket is already connected."
        Case WSAEMFILE
            strDesc = "Too many open files."
        Case WSAEMSGSIZE
            strDesc = "Message too long."
        Case WSAENETDOWN
            strDesc = "Network is down."
        Case WSAENETRESET
            strDesc = "Network dropped connection on reset."
        Case WSAENETUNREACH
            strDesc = "Network is unreachable."
        Case WSAENOBUFS
            strDesc = "No buffer space available."
        Case WSAENOPROTOOPT
            strDesc = "Bad protocol option."
        Case WSAENOTCONN
            strDesc = "Socket is not connected."
        Case WSAENOTSOCK
            strDesc = "Socket operation on nonsocket."
        Case WSAEOPNOTSUPP
            strDesc = "Operation not supported."
        Case WSAEPFNOSUPPORT
            strDesc = "Protocol family not supported."
        Case WSAEPROCLIM
            strDesc = "Too many processes."
        Case WSAEPROTONOSUPPORT
            strDesc = "Protocol not supported."
        Case WSAEPROTOTYPE
            strDesc = "Protocol wrong type for socket."
        Case WSAESHUTDOWN
            strDesc = "Cannot send after socket shutdown."
        Case WSAESOCKTNOSUPPORT
            strDesc = "Socket type not supported."
        Case WSAETIMEDOUT
            strDesc = "Connection timed out."
        Case WSATYPE_NOT_FOUND
            strDesc = "Class type not found."
        Case WSAEWOULDBLOCK
            strDesc = "Resource temporarily unavailable."
        Case WSAHOST_NOT_FOUND
            strDesc = "Host not found."
        Case WSANOTINITIALISED
            strDesc = "Successful WSAStartup not yet performed."
        Case WSANO_DATA
            strDesc = "Valid name, no data record of requested type."
        Case WSANO_RECOVERY
            strDesc = "This is a nonrecoverable error."
        Case WSASYSCALLFAILURE
            strDesc = "System call failure."
        Case WSASYSNOTREADY
            strDesc = "Network subsystem is unavailable."
        Case WSATRY_AGAIN
            strDesc = "Nonauthoritative host not found."
        Case WSAVERNOTSUPPORTED
            strDesc = "Winsock.dll version out of range."
        Case WSAEDISCON
            strDesc = "Graceful shutdown in progress."
        Case Else
            strDesc = "Unknown error."
    End Select
    
    GetErrorDescription = strDesc
    
End Function


