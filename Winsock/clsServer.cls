VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Server side Winsock API functions
Private Declare Function WSABind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByRef namelen As Long) As Long
Private Declare Function WSAListen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function WSAAccept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByRef addrlen As Long) As Long

'---------------------------------------------------------------------------

'Maximum queue length specifiable by listen.
Private Const SOMAXCONN = &H7FFFFFFF

'---------------------------------------------------------------------------

Public Event OnClose(lngSocket As Long)
Public Event OnConnectRequest(lngSocket As Long)
Public Event OnDataArrive(lngSocket As Long)
Public Event OnError(lngRetCode As Long, strDescription As String)

'---------------------------------------------------------------------------

Friend Sub WinsockMessage(lngMessage As Long, lngSocket As Long)
'********************************************************************************
'Date      :13 October 2002
'Purpose   :This function is called by the WindowProc when a Winsock message
'           has been sent.
'Arguments :lngMessage - The winsock message sent by Windows.
'Returns   :VOID
'********************************************************************************

    Select Case lngMessage
        Case FD_ACCEPT      'There is a connection request pending
            RaiseEvent OnConnectRequest(lngSocket)
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
    Set clsSvr = Me
    
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


Public Function Accept(ByVal lngSocketHandle As Long) As Long
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Accepts a remote connection on a socket
'Arguments :lngSocketHandle - The socket handle to accept the connection
'Returns   :The new socket handle that has the connection established
'********************************************************************************

  Dim udtSocketAddress As SOCKADDR_IN

    'Call the accept Winsock API function in order to create a new socket
    Accept = WSAAccept(lngSocketHandle, udtSocketAddress, LenB(udtSocketAddress))

End Function


Public Sub Listen(ByVal lngSocket As Long, lngLocalPort As Long)
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Binds a socket to the local port and sets the state of a socket,
'           given it's handle, to a listening state
'Arguments :lngSocketHandle - The socket handle to be set to listen
'Returns   :VOID
'********************************************************************************

  Dim udtSocketAddress As SOCKADDR_IN
  Dim lngReturnValue   As Long
  Dim lngAddress       As Long

    'Notice that the first part functino is very similar to the vbConnect one.
    'This is because Binding is just like connecting to the local machine

    If Not lngSocket > 0 Then Exit Sub      'Check the socket handle
    If Not lngLocalPort > 0 Then Exit Sub   'Check the port number
    
    With udtSocketAddress
        .sin_addr = inet_addr("0.0.0.0")                    'We will accept any IP
        .sin_port = htons(UnsignedToInteger(lngLocalPort))  'The port to listen on
        .sin_family = AF_INET                 'The internet family (TCP/IP, UDP etc)
    End With
        
    'Call the bind method
    Call WSABind(lngSocket, udtSocketAddress, LenB(udtSocketAddress))

    'Call the listen function
    Call WSAListen(lngSocket, SOMAXCONN)
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
