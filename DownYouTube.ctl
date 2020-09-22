VERSION 5.00
Begin VB.UserControl DownYouTube 
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   Picture         =   "DownYouTube.ctx":0000
   ScaleHeight     =   2130
   ScaleWidth      =   2760
   ToolboxBitmap   =   "DownYouTube.ctx":0312
End
Attribute VB_Name = "DownYouTube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
Event DownloadError(SaveFile As String)
Event DownloadComplete(MaxBytes As Long, SaveFile As String)
Event DownloadAllComplete(FileNotDownload() As String)
Event DownloadStage(sStage As String)

Private AsyncPropertyName() As String
Private AsyncStatusCode() As Byte

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    On Error Resume Next

        If AsyncProp.BytesMax <> 0 Then
            RaiseEvent DownloadProgress(CLng(AsyncProp.BytesRead), CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
        End If

        Select Case AsyncProp.StatusCode
          Case vbAsyncStatusCodeSendingRequest
          RaiseEvent DownloadStage("Attempting Connection")
            Debug.Print "Attempting to connect", AsyncProp.Target
          Case vbAsyncStatusCodeConnecting
          RaiseEvent DownloadStage("Connecting...")
            Debug.Print "Connecting", AsyncProp.Status
          Case vbAsyncStatusCodeBeginDownloadData
          RaiseEvent DownloadStage("Beginning to Download")
          RaiseEvent DownloadProgress(CLng(AsyncProp.BytesRead), CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
            Debug.Print "Begin downloading", AsyncProp.Status
            'Case vbAsyncStatusCodeDownloadingData
            '  Debug.Print "Downloading", AsyncProp.Status 'show target URL
          Case vbAsyncStatusCodeRedirecting
            Debug.Print "Redirecting", AsyncProp.Status
          Case vbAsyncStatusCodeEndDownloadData
          RaiseEvent DownloadStage("Complete!")
            Debug.Print "Download complete", AsyncProp.Status
          Case vbAsyncStatusCodeError
          RaiseEvent DownloadStage("Error!!")
            Debug.Print "Error...aborting transfer", AsyncProp.Status
            CancelAsyncRead AsyncProp.PropertyName
        End Select

End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
  Dim f() As Byte, fn As Long
  Dim i As Integer

    On Error Resume Next

        Select Case AsyncProp.StatusCode
          Case vbAsyncStatusCodeEndDownloadData
            fn = FreeFile
            f = AsyncProp.Value
            RaiseEvent DownloadStage("Finalizing")
            Debug.Print "Writting to file " & AsyncProp.PropertyName
            Open AsyncProp.PropertyName For Binary Access Write As #fn
            Put #fn, , f
            Close #fn

            RaiseEvent DownloadComplete(CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)

          Case vbAsyncStatusCodeError
            CancelAsyncRead AsyncProp.PropertyName
            RaiseEvent DownloadError(AsyncProp.PropertyName)
        End Select

        For i = 1 To UBound(AsyncPropertyName)
            If AsyncPropertyName(i) = AsyncProp.PropertyName Then
                AsyncStatusCode(i) = AsyncProp.StatusCode
                Exit For
            End If
        Next i

        CheckAllDownloadComplete
End Sub

Private Sub UserControl_Initialize()
    SizeIt
    ReDim AsyncPropertyName(0)
    ReDim AsyncStatusCode(0)
End Sub

Private Sub UserControl_Resize()
    SizeIt
End Sub

Private Sub UserControl_Terminate()
    If UBound(AsyncPropertyName) > 0 Then CancelAllDownload
End Sub

Private Sub SizeIt()
    On Error GoTo ErrorSizeIt
    With UserControl
        .Width = ScaleX(32, vbPixels, vbTwips)
        .Height = ScaleY(32, vbPixels, vbTwips)
    End With
Exit Sub

ErrorSizeIt:
    MsgBox err & ":Error in call to SizeIt()." _
           & vbCrLf & vbCrLf & "Error Description: " & err.Description, vbCritical, "Warning"

Exit Sub
End Sub

Public Sub BeginDownload(URL As String, SaveFile As String, Optional AsyncReadOptions = vbAsyncReadForceUpdate)
    On Error GoTo ErrorBeginDownload
    UserControl.AsyncRead URL, vbAsyncTypeByteArray, SaveFile, AsyncReadOptions

    ReDim Preserve AsyncPropertyName(UBound(AsyncPropertyName) + 1)
    AsyncPropertyName(UBound(AsyncPropertyName)) = SaveFile
    ReDim Preserve AsyncStatusCode(UBound(AsyncStatusCode) + 1)
    AsyncStatusCode(UBound(AsyncStatusCode)) = 255

Exit Sub

ErrorBeginDownload:
    MsgBox err & ":Error in call to BeginDownload()." _
           & vbCrLf & vbCrLf & "Error Description: " & err.Description, vbCritical, "Warning"

Exit Sub
End Sub

Public Function CancelAllDownload() As Boolean
  Dim i As Integer

    On Error Resume Next

        For i = 1 To UBound(AsyncPropertyName)
            CancelAsyncRead AsyncPropertyName(i)
            RaiseEvent DownloadStage("Cancelling")
            Debug.Print "Killing download " & AsyncPropertyName(i)
        Next i

        ReDim AsyncPropertyName(0)
        ReDim AsyncStatusCode(0)

        CancelAllDownload = True
End Function

Private Function CheckAllDownloadComplete()
  Dim i As Integer
  Dim FileNotDownload() As String
  Dim AllDownloadComplete As Boolean
    ReDim FileNotDownload(0)
    AllDownloadComplete = True
    For i = 1 To UBound(AsyncStatusCode)
        If AsyncStatusCode(i) = vbAsyncStatusCodeError Then
            ReDim Preserve FileNotDownload(UBound(FileNotDownload) + 1)
            FileNotDownload(UBound(FileNotDownload)) = AsyncPropertyName(i)
          ElseIf AsyncStatusCode(i) <> vbAsyncStatusCodeEndDownloadData Then
            AllDownloadComplete = False
            Exit For
        End If
    Next i
    If AllDownloadComplete Then
        CancelAllDownload
        RaiseEvent DownloadAllComplete(FileNotDownload)
    End If
End Function
