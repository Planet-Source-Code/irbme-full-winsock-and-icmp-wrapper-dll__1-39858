Attribute VB_Name = "modMessageHandler"
Option Explicit

'Window creation and destruction functions
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

'Subclassing functions
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Messaging functions
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

'Memory allocation functions
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

'Memory copy and move functions
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

'---------------------------------------------------------------------------

Private Const GWL_WNDPROC = (-4)    'Window Procedure flag for SetWindowLong

'---------------------------------------------------------------------------

Public clsClnt             As clsClient       'Instance of the client class
Public clsSvr              As clsServer       'Instance of the server class
Public clsPng              As clsPing         'Instance of the Ping class

Public WinsockMessage      As Long            'Winsock resolve host message
Public ResolveHostMessage  As Long            'General Winsock message

Private PrevProc           As Long            'Previous Window Procedure
Public WindowHandle        As Long            'Window handle

'---------------------------------------------------------------------------

Public Sub CreateMessageHandler()
'********************************************************************************
'Date      :14 October 2002
'Purpose   :Creates a blank window to be subclassed. It also creates the 2
'           messages we will look for.
'Arguments :VOID
'Returns   :VOID
'********************************************************************************

    'Create a blank, invisible window
    WindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

    'Create 2 unique message numbers for our messages
    WinsockMessage = RegisterWindowMessage(App.EXEName & ".WinsockMessage")
    ResolveHostMessage = RegisterWindowMessage(App.EXEName & ".ResolveHostMessage")
    
    'Subclass the window
    PrevProc = SetWindowLong(WindowHandle, GWL_WNDPROC, AddressOf WindowProc)
    
End Sub


Public Sub DestroyMessageHandler()
'********************************************************************************
'Date      :14 October 2002
'Purpose   :Destroys the message handler window created with CreateMessageHandler
'Arguments :VOID
'Returns   :VOID
'********************************************************************************
    
    If PrevProc <> 0 Then   'If we have subclassed the window
        'Return control to the previous window handler so it can close.
        SetWindowLong WindowHandle, GWL_WNDPROC, PrevProc
        
        'Destroy the window
        DestroyWindow WindowHandle
        
        PrevProc = 0
    End If
    
End Sub


Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'********************************************************************************
'Date      :14 October 2002
'Purpose   :Windows Callback function
'Arguments :hwnd   - The handle of the owner window
'           uMsg   - The message code
'           wParam - Dependant on message
'           lParam - Dependant on message
'Returns   :The return value from CallWindowProc if not a winsock message
'********************************************************************************
    
  Dim lngErrorCode As Long
    
    lngErrorCode = HiWord(lParam)
    
    'A general Winsock message
    If uMsg = WinsockMessage Then
        If ObjPtr(clsClnt) Then clsClnt.WinsockMessage lParam, wParam
        If ObjPtr(clsSvr) Then clsSvr.WinsockMessage lParam, wParam
        
    'A host resolving message from Winsock
    ElseIf uMsg = ResolveHostMessage Then
    
      Dim udtHost           As HOSTENT
      Dim lngIpAddrPtr      As Long
      Dim lngHostAddress    As Long
            
        If Not lngErrorCode > 0 Then
            'Extract the host name from the memory block
            RtlMoveMemory udtHost, ByVal lngMemoryPointer, Len(udtHost)
            RtlMoveMemory lngIpAddrPtr, ByVal udtHost.hAddrList, 4
            RtlMoveMemory lngHostAddress, ByVal lngIpAddrPtr, 4
            
            'Free the allocated memory block
            Call GlobalUnlock(lngMemoryHandle)
            Call GlobalFree(lngMemoryHandle)
        Else
            lngHostAddress = INADDR_NONE
        End If
        
        If ObjPtr(clsClnt) Then clsClnt.ResolveHostMessage lngHostAddress
        If ObjPtr(clsPng) Then clsPng.ResolveHostMessage lngHostAddress

    Else
        'A non-Winsock related message. Call the default message handler.
        WindowProc = CallWindowProc(PrevProc, hWnd, uMsg, wParam, lParam)
    End If
    
End Function
