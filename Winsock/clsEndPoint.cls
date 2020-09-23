VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEndPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function GetLocalHost(lngSocket As Long) As String
'********************************************************************************
'Date      :15 October 2002
'Purpose   :Function that Retrieves the local hostname from a connected socket
'Arguments :lngSocket     - The socket handle on which the connection is
'                           established.
'Returns   :The information requested or -1 if there was an error
'********************************************************************************
    
    GetLocalHost = GetIPEndPointField(lngSocket, LOCAL_HOST)
End Function


Public Function GetLocalIP(lngSocket As Long) As String
'********************************************************************************
'Date      :15 October 2002
'Purpose   :Function that Retrieves the local IP from a connected socket
'Arguments :lngSocket     - The socket handle on which the connection is
'                           established.
'Returns   :The information requested or -1 if there was an error
'********************************************************************************
    
    GetLocalIP = GetIPEndPointField(lngSocket, LOCAL_HOST_IP)
End Function


Public Function GetLocalPort(lngSocket As Long) As Long
'********************************************************************************
'Date      :15 October 2002
'Purpose   :Function that Retrieves the local port from a connected socket
'Arguments :lngSocket     - The socket handle on which the connection is
'                           established.
'Returns   :The information requested or -1 if there was an error
'********************************************************************************
    
    GetLocalPort = GetIPEndPointField(lngSocket, LOCAL_PORT)
End Function


Public Function GetRemoteHost(lngSocket As Long) As String
'********************************************************************************
'Date      :15 October 2002
'Purpose   :Function that Retrieves the remote hostname from a connected socket
'Arguments :lngSocket     - The socket handle on which the connection is
'                           established.
'Returns   :The information requested or -1 if there was an error
'********************************************************************************
    
    GetRemoteHost = GetIPEndPointField(lngSocket, REMOTE_HOST)
End Function


Public Function GetRemoteIP(lngSocket As Long) As String
'********************************************************************************
'Date      :15 October 2002
'Purpose   :Function that Retrieves the remote IP from a connected socket
'Arguments :lngSocket     - The socket handle on which the connection is
'                           established.
'Returns   :The information requested or -1 if there was an error
'********************************************************************************
    
    GetRemoteIP = GetIPEndPointField(lngSocket, REMOTE_HOST_IP)
End Function


Public Function GetRemotePort(lngSocket As Long) As Long
'********************************************************************************
'Date      :15 October 2002
'Purpose   :Function that Retrieves the remote port from a connected socket
'Arguments :lngSocket     - The socket handle on which the connection is
'                           established.
'Returns   :The information requested or -1 if there was an error
'********************************************************************************
    
    GetRemotePort = GetIPEndPointField(lngSocket, REMOTE_PORT)
End Function

