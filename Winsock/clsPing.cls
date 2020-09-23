VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ICMP ping API functions
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean

'---------------------------------------------------------------------------

'Echo request options
Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type

'Ping reply information
Private Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type

'---------------------------------------------------------------------------

Public Event OnPingTimedOut(strIPAddress As String)
Public Event OnPingResponce(lngRoundTripTime As Long, strIPAddress As String)
Public Event OnError(lngRetCode As Long, strDescription As String)

'---------------------------------------------------------------------------

Private ReplyPing As Boolean

'---------------------------------------------------------------------------

Friend Sub ResolveHostMessage(lngAddress As Long)
'********************************************************************************
'Date      :15 October 2002
'Purpose   :This function is called by the WindowProc when the host is
'           resolved, or as the case may be, unresolved.
'Arguments :lngAddress - The long address format of the host
'Returns   :VOID
'********************************************************************************

  Dim OptInfo   As IP_OPTION_INFORMATION
  Dim EchoReply As IP_ECHO_REPLY
  Dim hFile     As Long
  
    'Only trigger an event if this was for us
    If Not ReplyPing Then Exit Sub
      
    'We only want to reply once
    ReplyPing = False
    
    If lngAddress <> INADDR_NONE Then

        'Create a ICMP file
        hFile = IcmpCreateFile()
        
        OptInfo.TTL = 255
        
        'Send the echo
        If IcmpSendEcho(hFile, lngAddress, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
            'We got a response back
            RaiseEvent OnPingResponce(EchoReply.RoundTripTime, EchoReply.Address(0) & "." & EchoReply.Address(1) & "." & EchoReply.Address(2) & "." & EchoReply.Address(3))
        Else
            'The host timed out
            RaiseEvent OnPingTimedOut(EchoReply.Address(0) & "." & EchoReply.Address(1) & "." & EchoReply.Address(2) & "." & EchoReply.Address(3))
        End If
        
        'Close the ICMP file
        IcmpCloseHandle hFile
    Else
        RaiseEvent OnError(WSAHOST_NOT_FOUND, "Unable to resolve host")
    End If
    
End Sub


Private Sub Class_Initialize()
'********************************************************************************
'Date      :15 October 2002
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
    Set clsPng = Me
    
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


Public Sub ICMPPing(strHost As String)
'********************************************************************************
'Date      :15 October 2002
'Purpose   :This procedure pings a hostname
'Arguments :strHost - The hostname or IP address to ping
'Returns   :VOID
'********************************************************************************
    ReplyPing = True
    ResolveHost strHost
End Sub
