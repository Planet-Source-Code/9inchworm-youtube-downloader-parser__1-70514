VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download YouTube Video"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5805
      TabIndex        =   8
      Top             =   3645
      Width           =   1140
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4380
      TabIndex        =   7
      Top             =   4185
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7020
      TabIndex        =   6
      Top             =   3615
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstDownload 
      Height          =   2640
      Left            =   60
      TabIndex        =   4
      Top             =   585
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   4657
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "SmallImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Video:"
         Object.Width           =   9596
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "vID:"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status:"
         Object.Width           =   2540
      EndProperty
   End
   Begin Project1.DownYouTube DownYouTube 
      Height          =   480
      Left            =   1155
      TabIndex        =   1
      Top             =   2580
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.ProgressBar prgBAR 
      Height          =   270
      Left            =   1080
      TabIndex        =   2
      Top             =   4965
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   870
      Top             =   4005
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDown.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDTo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1860
      TabIndex        =   11
      Top             =   3330
      Width           =   6390
   End
   Begin VB.Label Label3 
      Caption         =   "All Files Downloaded To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   10
      Top             =   3315
      Width           =   1860
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "* Do Not Close This Window While Your File Is Downloading"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   75
      TabIndex        =   9
      Top             =   3675
      Width           =   5910
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "YouTube Video Download Que:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   90
      TabIndex        =   5
      Top             =   135
      Width           =   3240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF862D&
      BackStyle       =   1  'Opaque
      Height          =   540
      Left            =   0
      Top             =   0
      Width           =   8370
   End
   Begin VB.Label lblLABEL 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1095
      TabIndex        =   3
      Top             =   5310
      Width           =   6465
   End
   Begin VB.Label lblTitle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1125
      TabIndex        =   0
      Top             =   4605
      Width           =   6510
   End
End
Attribute VB_Name = "frmDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
DownYouTube.CancelAllDownload
cmdClear_Click
End Sub

Private Sub cmdClear_Click()
On Error Resume Next
lstDownload.ListItems.Remove (lstDownload.SelectedItem.Index)
Call Form1.SaveList
End Sub

Private Sub Form_Load()
lblDTo.Caption = App.Path & "\Downloaded Files\"
    Call LoadList
End Sub

Public Function DownloadYT(URL$, FileName$)
Dim Item As ListItem
For Each Item In lstDownload.ListItems
    DownYouTube.BeginDownload URL, App.Path & "\Downloaded Files\" & FileName & ".flv"
Next
End Function

Private Sub DownYouTube_DownloadAllComplete(FileNotDownload() As String)
  Dim i As Integer
    Debug.Print "Finished all download"
    cmdDownload.Enabled = True
    cmdCancel.Enabled = False
    If UBound(FileNotDownload) > 0 Then
        For i = 1 To UBound(FileNotDownload)
            Debug.Print "File not downloaded: " & FileNotDownload(i)
        Next i
    End If
End Sub

Private Sub DownYouTube_DownloadStage(sString As String)
'
End Sub

Private Sub DownYouTube_DownloadComplete(MaxBytes As Long, SaveFile As String)
  Dim i As Integer
    Debug.Print "Completed " & SaveFile & ", Size = " & MaxBytes
    MsgBox "Completed!", vbInformation, "Success"
    With lstDownload
        For i = 1 To .ListItems.Count
            If .ListItems(i).Key = SaveFile Then
                .ListItems(i).SubItems(2) = "Completed"
            End If
        Next i
    End With
cmdClear_Click
Unload frmDown
Call Form1.ClearConts
Form1.Show
End Sub

Private Sub DownYouTube_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
  Dim i As Integer
  Dim RemBytes As Long
    With lstDownload
        For i = 1 To .ListItems.Count
                RemBytes = MaxBytes - CurBytes
                If RemBytes < 2 ^ 20 Then
                    .ListItems(i).SubItems(2) = Format((MaxBytes - CurBytes) / 2 ^ 10, "#0.0 KB") & _
                               " (" & Format(CurBytes / MaxBytes, "#0.0%") & ")"
                  Else
                    .ListItems(i).SubItems(2) = Format((MaxBytes - CurBytes) / 2 ^ 20, "#0.00 MB") & _
                               " (" & Format(CurBytes / MaxBytes, "#0.0%") & ")"
            End If
        Next i
    End With
End Sub

Private Sub DownYouTube_DownloadError(SaveFile As String)
  Dim i As Integer
    Debug.Print "Error downloading " & SaveFile
    With lstDownload
        For i = 1 To .ListItems.Count
            If .ListItems(i).Key = SaveFile Then
                .ListItems(i).SubItems(2) = "Error"
            End If
        Next i
    End With
End Sub

Private Function GetFilename$(URL$)
  Dim i As Integer
    For i = Len(URL) To 1 Step -1
        If Mid(URL, i, 1) = "/" Then
            GetFilename = Right(URL, Len(URL) - i)
            Exit For
        End If
    Next i
End Function

Public Sub LoadList()
    Dim fso As New FileSystemObject
    Dim strm As TextStream
    Dim Item As ListItem
    Dim Tmp1 As String
    Dim Tmp2 As String
    Dim Tmp3 As String
    
On Error GoTo err
    Set strm = fso.OpenTextFile(App.Path & "\List", ForReading)
On Error Resume Next
    lstDownload.ListItems.Clear
    Do Until strm.AtEndOfStream
        Tmp1 = strm.ReadLine
        Tmp2 = strm.ReadLine
        Tmp3 = strm.ReadLine
        Set Item = lstDownload.ListItems.Add(, Tmp1, Tmp2)
        Item.SubItems(1) = Tmp1
        Item.SubItems(2) = Tmp3
    Loop
    strm.Close
    Exit Sub
err:
    Set strm = fso.CreateTextFile(App.Path & "\List")
    strm.Close
    Dim File As File
    Set File = fso.GetFile(App.Path & "\List")
    File.Attributes = Hidden + System
    LoadList
End Sub

