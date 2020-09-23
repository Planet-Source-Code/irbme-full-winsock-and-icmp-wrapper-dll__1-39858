Attribute VB_Name = "modHostResolver"
Option Explicit

'Winsock API functions for resolving hostnames and IP's
Private Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal buflen As Long) As Long
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Private Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long

'Memory copy and move functions
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

'---------------------------------------------------------------------------

'End Point of connection information
Public Enum IPEndPointFields
    LOCAL_HOST          'Local hostname
    LOCAL_HOST_IP       'Local IP
    LOCAL_PORT          'Local port
    REMOTE_HOST         'Remote hostname
    REMOTE_HOST_IP      'Remote IP
    REMOTE_PORT         'Remote port
End Enum

'---------------------------------------------------------------------------

Private Const GMEM_FIXED = &H0      'Fixed memory flag for GlobalAlloc

'---------------------------------------------------------------------------

Public lngMemoryHandle    As Long 'handle of the allocated memory block object
Public lngMemoryPointer   As Long 'address of the memory block

'---------------------------------------------------------------------------

Public Sub ResolveHost(strHostName As String)
'********************************************************************************
'Date      :14 October 2002
'Purpose   :Tries to resolve an IP address, or hostname into a long address
'Arguments :strHostName - The IP address or hostname to resolve
'Returns   :VOID
'********************************************************************************

  Dim lngAddress As Long

    'Try and resolve the address. This will work if it was an IP we were given
    lngAddress = inet_addr(strHostName)
    
    'We were unable to resolve it so we will have to go for the long way
    If lngAddress = INADDR_NONE Then
        'Allocate 1Kb of fixed memory
        lngMemoryHandle = GlobalAlloc(GMEM_FIXED, 1024)
        
        If lngMemoryHandle > 0 Then
            'Lock the memory block just to get the address
            lngMemoryPointer = GlobalLock(lngMemoryHandle)

            If lngMemoryPointer = 0 Then
                'Memory allocation error
                Call GlobalFree(lngMemoryHandle)
                Exit Sub
            Else
                'Unlock the memory block
                GlobalUnlock (lngMemoryHandle)
            End If
        Else
            'Memory allocation error
            Exit Sub
        End If
        
        'Get the host by the name. This is an Asynchroneous call. This means
        'that the call will not freeze the app. It will post a message
        'to the WindowProc when it has finished.
        WSAAsyncGetHostByName WindowHandle, ResolveHostMessage, strHostName, ByVal lngMemoryPointer, 1024
    Else
        If ObjPtr(clsClnt) Then clsClnt.ResolveHostMessage (lngAddress)
        If ObjPtr(clsPng) Then clsPng.ResolveHostMessage (lngAddress)
    End If

End Sub


Public Function GetIPEndPointField(ByVal lngSocket As Long, ByVal EndpointField As IPEndPointFields) As Variant
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Function that Retrieves, address, host name or port number of an
'           end-point of the connection established on the socket.
'Arguments :lngSocket     - The socket handle on which the connection is
'                           established.
'           EndPointField - The request for information.
'Returns   :The information requested or -1 if there was an error
'********************************************************************************


  Dim udtSocketAddress    As SOCKADDR_IN
  Dim lngReturnValue      As Long
  Dim lngPtrToAddress     As Long
  Dim strIPAddress        As String
  Dim lngAddress          As Long

    Select Case EndpointField
        Case LOCAL_HOST, LOCAL_HOST_IP, LOCAL_PORT

            'If the info of a local end-point of the connection is
            'requested, call the getsockname Winsock API function
            lngReturnValue = getsockname(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
        Case REMOTE_HOST, REMOTE_HOST_IP, REMOTE_PORT
            
            'If the info of a remote end-point of the connection is
            'requested, call the getpeername Winsock API function
            lngReturnValue = getpeername(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
    End Select
    
    
    If lngReturnValue = 0 Then
        'If no errors occurred, the getsockname or getpeername function returns 0.

        Select Case EndpointField
            Case LOCAL_PORT, REMOTE_PORT
                'Get the port number from the sin_port field and convert the byte ordering
                GetIPEndPointField = IntegerToUnsigned(ntohs(udtSocketAddress.sin_port))
            
            Case LOCAL_HOST_IP, REMOTE_HOST_IP
  
                'Get pointer to the string that contains the IP address
                lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
                
                'Retrieve that string by the pointer
                GetIPEndPointField = StringFromPointer(lngPtrToAddress)
            Case LOCAL_HOST, REMOTE_HOST

                'The same procedure as for an IP address only using GetHostNameByAddress
                lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
                strIPAddress = StringFromPointer(lngPtrToAddress)
                lngAddress = inet_addr(strIPAddress)
                GetIPEndPointField = GetHostNameByAddress(lngAddress)

        End Select
    'An error occured
    Else
        GetIPEndPointField = SOCKET_ERROR
    End If
    
End Function


Private Function GetHostNameByAddress(lngInetAdr As Long) As String
'********************************************************************************
'Date      :13 October 2002
'Purpose   :Returns the hostname given an address in long format
'Arguments :lngInetAdr - The address to resolve into a hostname
'Returns   :Returns the hostname as a string
'********************************************************************************

  Dim lngPtrHostEnt As Long
  Dim udtHostEnt    As HOSTENT
  Dim strHostName   As String
  
    'Get the pointer to the HOSTENT structure
    lngPtrHostEnt = gethostbyaddr(lngInetAdr, 4, AF_INET)
    
    'Copy data into the HOSTENT structure
    RtlMoveMemory udtHostEnt, ByVal lngPtrHostEnt, LenB(udtHostEnt)
    
    'Prepare the buffer to receive a string
    strHostName = String(256, 0)
    
    'Copy the host name into the strHostName variable
    RtlMoveMemory ByVal strHostName, ByVal udtHostEnt.hName, 256
    
    'Cut received string by first chr(0) character
    GetHostNameByAddress = Left(strHostName, InStr(1, strHostName, Chr(0)) - 1)

End Function

