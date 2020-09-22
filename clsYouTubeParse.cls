VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsYouTubeParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsYouTubeParse
'©2008 9InchWorM Software
'gollum_nz@hotmail.com
Option Explicit

'YouTube Video Title
Public Function GetYouTubeVidTitle$(Data$)
    Dim lngA As Long, lngB As Long, vidTitle As Long, YTitle As String
    lngA = InStr(1, Data, "<title>") + 7
    lngB = InStr(lngA, Data, "</title>")
    GetYouTubeVidTitle = Mid(Data, lngA, lngB - lngA)
End Function

'YouTube Video ID String
Public Function GetYouTubeID$(Data$)
    Dim lngA As Long, lngB As Long
    lngA = InStr(1, Data, "&video_id=") + 10
    lngB = InStr(lngA, Data, "&title=")
    GetYouTubeID = Mid(Data, lngA, lngB - lngA)
End Function

'YouTube Video Views
Public Function GetYouTubeViews$(Data$)
    Dim lngA As Long, lngB As Long
    lngA = InStr(1, Data, "watch-view-count" & Chr(34) & ">") + 18
    lngB = InStr(lngA, Data, "</span>")
    GetYouTubeViews = Mid(Data, lngA, lngB - lngA)
End Function

'YouTube Video Poster
Public Function GetYouTubeUserID$(Data$)
    Dim lngA As Long, lngB As Long
    lngA = InStr(1, Data, "urchinTracker('/VideoWatch/ChannelNameLink');" & Chr(34) & ">") + 47
    lngB = InStr(lngA, Data, "</a>")
    GetYouTubeUserID = Mid(Data, lngA, lngB - lngA)
End Function

'YouTube Date Added
Public Function GetYouTubeDate$(Data$)
    Dim lngA As Long, lngB As Long
    lngA = InStr(1, Data, "<span class=" & Chr(34) & "watch-video-added" & Chr(34) & ">") + 32
    lngB = InStr(lngA, Data, "</span>")
    GetYouTubeDate = Mid(Data, lngA, lngB - lngA)
End Function

'YouTube Category
Public Function GetYouTubeCategory$(Data$)
    Dim lngA As Long, lngB As Long
    lngA = InStr(1, Data, "urchinTracker('/VideoWatch/VideoCategoryLink');" & Chr(34) & ">") + 49
    lngB = InStr(lngA, Data, "</a>")
    GetYouTubeCategory = Mid(Data, lngA, lngB - lngA)
End Function

'YouTube Description
Public Function GetYouTubeDes$(Data$)
    Dim lngA As Long, lngB As Long
    lngA = InStr(1, Data, "<div  class=" & Chr(34) & "watch-video-desc" & Chr(34) & ">") + 32
    lngB = InStr(lngA, Data, "</span>")
    GetYouTubeDes = Mid(Data, lngA, lngB - lngA)
End Function
