VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGSwsHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' DGSwsHTTP Class Module by RegX
' Copyright 2002 DGS
' You may freely use this code as long as
' All Copyright information remains intact
'=============================================
'Requires:
'cAppendString.cls
'Microsoft Scripting Runtime
'MSWinsockLib.Winsock (Mswinsck.ocx)
'=============================================
'Description
'Makes downloading HTML files with progress a snap
'Exposes Response Header and FileData seperately

'Usage ---------------------------------------------------------------
'Dim WithEvents wsc As DGSwsHTTP
'Set wsc = New DGSwsHTTP
'wsc.geturl "http://www.microsoft.com/"
'---------------------------------------------------------------------
Option Explicit
Dim cstring As cAppendString
Dim WithEvents ws As MSWinsockLib.Winsock
Attribute ws.VB_VarHelpID = -1
Public Event ProgressChanged(ByVal bytesreceived As Long)
Public URL As String
Private mvarRemotePort As String
Private mvarRemoteHost As String
Public totalbytesreceived As Long
Public Event DownloadComplete()
Public FileSize As Variant
Public ResponseHeaderString As String
Private dicResponseHeader As Dictionary
Public filedata As Variant
Public ResponseCode As String
Public Event httpError(errmsg As String, Scode As String)

Public Property Let RemoteHost(ByVal vData As String)
    mvarRemoteHost = vData
End Property

Public Property Get RemoteHost() As String
    RemoteHost = mvarRemoteHost
End Property

Public Property Let RemotePort(ByVal vData As String)
    mvarRemotePort = vData
End Property

Public Property Get RemotePort() As String
    RemotePort = mvarRemotePort
End Property


Public Sub geturl(URL As String)
            If URL & "" = "" Then RaiseEvent httpError("No URL specified", 0)
            cstring.Clear
            Dim RHstart As Long
            Dim RHend As Long
            Dim RH As String
            RHstart = InStr(1, URL, "://", vbTextCompare)
            If RHstart = 0 Then RHstart = 1 Else RHstart = RHstart + 3
            If Len(URL) <= RHstart + 3 Then
                RaiseEvent httpError("Invalid URL", 0)
                Exit Sub
            End If
            RHend = InStr(RHstart + 3, URL, "/", vbTextCompare)
            If RHend = 0 Then RHend = Len(URL) + 1
            RH = Mid(URL, RHstart, RHend - (RHstart))
            
            Me.URL = URL
            ws.Close
            ws.RemotePort = Me.RemotePort
            ws.RemoteHost = RH
            ws.Connect
End Sub

Private Sub Class_Initialize()
Set ws = New MSWinsockLib.Winsock
Set cstring = New cAppendString
Set dicResponseHeader = New Dictionary
Me.RemotePort = 80
End Sub

Private Sub Class_Terminate()
Set dicResponseHeader = Nothing
Set cstring = Nothing
Set ws = Nothing
End Sub

Private Sub ws_Close()
filedata = cstring.Value
cstring.Clear
RaiseEvent DownloadComplete
End Sub

Private Sub ws_Connect()
    Me.totalbytesreceived = 0
    Me.FileSize = 0
    Me.ResponseHeaderString = ""
    Me.filedata = ""
        Dim Request As String
        Request = "GET " & Me.URL & " HTTP/1.0" & vbCrLf & vbCrLf
        ws.SendData Request
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
dicResponseHeader.RemoveAll
On Error Resume Next
Dim data As String
Dim arrheader() As String
Dim headkeys() As String
Dim headvals() As String
Dim Item As Variant
Dim X As Long
ws.GetData data
        If totalbytesreceived = 0 Then
            ResponseHeaderString = Left(data, InStr(1, data, vbCrLf & vbCrLf, vbTextCompare) + 2)
            totalbytesreceived = bytesTotal - (Len(ResponseHeaderString) + 2)
            cstring.Append Mid(data, Len(ResponseHeaderString) + 2)
            arrheader = Split(ResponseHeaderString, vbCrLf, -1, vbTextCompare)
            For Each Item In arrheader
                X = InStr(1, Item, " ", vbTextCompare)
                If X > 1 Then
                    dicResponseHeader.Add Left(Item, X - 1), Mid(Item, X + 1)
                    Debug.Print Left(Item, X - 1) & Mid(Item, X + 1)
                End If
            Next
                    ResponseCode = Left(dicResponseHeader.Items(0), 3)
                    Me.FileSize = dicResponseHeader.Item("Content-Length:")
                    If ResponseCode <> "200" Then
                        RaiseEvent httpError(CStr(dicResponseHeader.Items(0)), ResponseCode)
                    End If
        Else
        cstring.Append data
            totalbytesreceived = totalbytesreceived + bytesTotal
            RaiseEvent ProgressChanged(totalbytesreceived)
        End If
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent httpError(Description, CStr(Scode))
End Sub
