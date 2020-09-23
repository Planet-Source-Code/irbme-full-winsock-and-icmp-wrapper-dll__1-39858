VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function WSAConnect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long

'---------------------------------------------------------------------------

'Public Socket Events
Public Event OnClose(lngSocket As Long)
Public Event OnConnect(lngSocket As Long)
Public Event OnDataArrive(lngSocket As Long)
Public Event OnError(lngRetCode As Long, strDescription As String)

'---------------------------------------------------------------------------

Private intConnectionPort   As Integer
Private lngConnectionSocket As Long

'---------------------------------------------------------------------------

Friend Sub ResolveHostMessage(lngAddress As Long)
'********************************************************************************
'Date      :14 October 2002
'Purpose   :This function is called by the WindowProc when the host is
'           resolved, or as the case may be, unresolved.
'Arguments :lngAddress - The long address format of the host
'Returns   :VOID
'********************************************************************************

  Dim udtSocketAddress As SOCKADDR_IN

    If lngAddress <> INADDR_NONE Then
        With udtSocketAddress
            .sin_addr = lngAddress
            .sin_port = htons(UnsignedToInteger(CLng(intConnectionPort)))
            .sin_family = AF_INET
        End With
        
        'Call the connect API function
        Call WSAConnect(lngConnectionSocket, udtSocketAddress, LenB(udtSocketAddress))
    Else
        RaiseEvent OnError(WSAHOST_NOT_FOUND, "Unable to resolve host")
    End If
    
End Sub


Friend Sub WinsockMessage(lngMessage As Long, lngSocket As Long)
'********************************************************************************
'Date      :13 October 2002
'Purpose   :This function is called by the WindowProc when a Winsock message
'           has been sent.
'Arguments :lngMessage - The winsock message sent by Windows.
'Returns   :VOID
'********************************************************************************

    Select Case lngMessage
        Case FD_CONNECT     'The connection is established
            RaiseEvent OnConnect(lngSocket)
        Case FD_READ        'We are ready to read from the winsock buffer
            RaiseEvent OnDataArrive(lngSocket)
        Case FD_CLOSE       'The other end point wishes to close the connection
            RaiseEvent OnClose(lngSocket)
    End Select
    
End Sub


Private Sub Class_Initialize()
'********************************************************************************
'Date      :14 October 2002
'Purpose   :This procedure initiializes the DLL by starting up winsock and
'           also the message handling module.
'Arguments :VOID
'Returns   :VOID
'********************************************************************************

  Dim WSAInfo As WSAData
    
    If Not WSAStarted Then
        'Create the message handler window and corresponding messages.
        CreateMessageHandler
        
        'Start version 1.1 of the Microsoft Winsock API
        WSAStartup &H101, WSAInfo
        
        WSAStarted = True
    End If
    
    'Store the instance of this class for the use of the modules so that they
    'can have access to the friend message handler subroutines.
    Set clsClnt = Me
    
End Sub


Private Sub Class_Terminate()
'********************************************************************************
'Date      :14 October 2002
'Purpose   :This procedure cleans up the DLL by terminating up winsock and
'           also the message handling module.
'Arguments :VOID
'Returns   :VOID
'********************************************************************************
    
    If WSAStarted Then
        'Destroy the message handler window
        DestroyMessageHandler
        
        'Cleanup Winsock
        WSACleanup
        
        WSAStarted = False
    End If
    
End Sub


Public Sub Connect(ByVal lngSocket As Long, ByVal strRemoteHost As String, ByVal intRemotePort As Integer)
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Connects the socket to a remote host on a specific port
'Arguments :lngSocket     - The socket handle to connect
'           strRemoteHost - The remote host name or IP
'           intRemotePort - The port to connect through
'Returns   :VOID
'********************************************************************************

    'Check the socket handle
    If Not lngSocket > 0 Then RaiseEvent OnError(0, "Invalid socket handle"): Exit Sub
    'Check the remote host address argument
    If Len(strRemoteHost) = 0 Then RaiseEvent OnError(0, "Invalid hostname"): Exit Sub
    'Check the port number
    If Not intRemotePort > 0 Then RaiseEvent OnError(0, "Invalid port number"): Exit Sub
    
    ResolveHost strRemoteHost
    
    intConnectionPort = intRemotePort
    lngConnectionSocket = lngSocket
    
End Sub


Public Function Send(ByVal lngSocket As Long, strData As String) As Long
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Sends data to the Winsock buffer of a specific port
'Arguments :lngSocket - The socket handle to send data to
'           strData - The string buffer to send
'Returns   :The number of bytes written to the buffer
'********************************************************************************
    
    Send = mSend(lngSocket, strData)
End Function


Public Function Recv(ByVal lngSocket As Long, strBuffer As String) As Long
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Recieves data from the Winsock buffer of a specific port
'Arguments :lngSocket - The socket handle to read from
'           strBuffer - The string buffer to place the data into
'Returns   :The number of bytes read
'********************************************************************************
    
    Recv = mRecv(lngSocket, strBuffer)
End Function


Public Sub CloseSocket(lngSocket As Long)
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Closes a socket given it's handle
'Arguments :lngSocket - The socket handle to close
'Returns   :VOID
'********************************************************************************
    
    WSACloseSocket lngSocket
    lngSocket = 0
    
End Sub


Public Function CreateSocket() As Long
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Creates a new socket to be used with the other function calls
'Arguments :VOID
'Returns   :If no error then the socket handle, else INVALID_SOCKET is returned
'********************************************************************************
    
    CreateSocket = mCreateSocket
End Function
