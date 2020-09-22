VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "9InchWorM's YouTube Downloader"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDown 
      Caption         =   "Download Video"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   22
      Top             =   6435
      Width           =   1860
   End
   Begin VB.CommandButton cmndViewQ 
      Caption         =   "View Download Que"
      Height          =   270
      Left            =   4140
      TabIndex        =   21
      Top             =   6060
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data: "
      Height          =   2985
      Left            =   135
      TabIndex        =   7
      Top             =   465
      Width           =   8430
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H80000008&
         Height          =   2025
         Left            =   5430
         ScaleHeight     =   1995
         ScaleWidth      =   2820
         TabIndex        =   23
         Top             =   495
         Width           =   2850
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Preview Currently Unavailable"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   885
            Width           =   2595
         End
      End
      Begin VB.TextBox txtDes 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   900
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1605
         Width           =   4305
      End
      Begin VB.CommandButton cmdAddQ 
         Caption         =   "Add To Download Que"
         Height          =   270
         Left            =   150
         TabIndex        =   12
         Top             =   2610
         Width           =   1860
      End
      Begin VB.Label Label5 
         Caption         =   "Description:"
         Height          =   225
         Left            =   90
         TabIndex        =   20
         Top             =   1365
         Width           =   1755
      End
      Begin VB.Label lblCat 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   885
         TabIndex        =   18
         Top             =   1155
         Width           =   4830
      End
      Begin VB.Label Label4 
         Caption         =   "Category:"
         Height          =   240
         Left            =   90
         TabIndex        =   17
         Top             =   1140
         Width           =   1380
      End
      Begin VB.Label lblDate 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1080
         TabIndex        =   16
         Top             =   915
         Width           =   2970
      End
      Begin VB.Label Label3 
         Caption         =   "Date Added:"
         Height          =   225
         Left            =   90
         TabIndex        =   15
         Top             =   915
         Width           =   1125
      End
      Begin VB.Label lblytUser 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1125
         TabIndex        =   14
         Top             =   675
         Width           =   6405
      End
      Begin VB.Label Label1 
         Caption         =   "Uploaded By:"
         Height          =   285
         Left            =   90
         TabIndex        =   13
         Top             =   675
         Width           =   1260
      End
      Begin VB.Label lblTC 
         Caption         =   "Title of Video To Download:"
         Height          =   225
         Left            =   75
         TabIndex        =   11
         Top             =   240
         Width           =   2010
      End
      Begin VB.Label lblTitle 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2100
         TabIndex        =   10
         Top             =   240
         Width           =   6045
      End
      Begin VB.Label Label2 
         Caption         =   "Current Views:"
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   450
         Width           =   1125
      End
      Begin VB.Label lblViews 
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1230
         TabIndex        =   8
         Top             =   450
         Width           =   4725
      End
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   8160
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7410
      TabIndex        =   6
      Top             =   5760
      Width           =   660
   End
   Begin RichTextLib.RichTextBox rtf1 
      Height          =   915
      Left            =   1170
      TabIndex        =   5
      Top             =   6420
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   1614
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   -60
      TabIndex        =   3
      Top             =   5715
      Visible         =   0   'False
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txturl 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   1425
      TabIndex        =   2
      Text            =   "http://"
      Top             =   120
      Width           =   6240
   End
   Begin VB.PictureBox picstat1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   8610
      TabIndex        =   1
      Top             =   3495
      Width           =   8640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get It"
      Height          =   285
      Left            =   7800
      TabIndex        =   0
      Top             =   90
      Width           =   735
   End
   Begin VB.Label lblurl 
      Caption         =   "YouTube URL:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   4
      Top             =   165
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents wsc As DGSwsHTTP
Attribute wsc.VB_VarHelpID = -1
Dim yt As clsYouTubeParse
Dim ytID$, ytUID$, ytCat$, ytDes$, sF$, Added$

Private Sub cmdAddQ_Click()
If frmDown.lstDownload.ListItems.Count = 1 Then
MsgBox "Sorry, multiple downloading is not implemented yet", vbInformation, "Sorry"
Exit Sub
Else
sF = lblTitle & " - " & ytUID
Call Add2DownQue(ytID, sF, lblTitle)
Call ClearConts
Added = lblTitle
cmndViewQ_Click
End If
End Sub

Private Sub cmndViewQ_Click()
Load frmDown
frmDown.Show
End Sub

Private Sub Form_Load()
Set wsc = New DGSwsHTTP
Set yt = New clsYouTubeParse
End Sub
Private Sub Command1_Click()
If InStr(1, txturl, "http://www.youtube.com/watch?v=", vbTextCompare) Then
stat1 "Downloading Data ..."
wsc.geturl txturl
Command1.Enabled = False
Else
MsgBox "Not a valid YouTube Address", vbCritical, "Error"
End If
End Sub

Private Sub wsc_DownloadComplete()
Dim YTitle$
stat1 "Data Recieved / YouTube VidID Recieved"
rtf1.Text = wsc.filedata
YTitle = yt.GetYouTubeVidTitle(wsc.filedata)
lblTitle = YTitle 'Right(YTitle, Len(YTitle) - InStrRev(YTitle, "- "))
ytID = "http://www.youtube.com/get_video.php?video_id=" & yt.GetYouTubeID(wsc.filedata)
lblViews = yt.GetYouTubeViews(wsc.filedata)
lblytUser = yt.GetYouTubeUserID(wsc.filedata)
ytUID = yt.GetYouTubeUserID(wsc.filedata)
ytCat = yt.GetYouTubeCategory(wsc.filedata)
ytDes = yt.GetYouTubeDes(wsc.filedata)
lblCat = Replace(ytCat, "&", "&&")
lblDate = yt.GetYouTubeDate(wsc.filedata)
txtDes = Replace(ytDes, "<span >", "")
cmdDown.Enabled = True
End Sub

Private Sub stat1(statusmsg As String)
picstat1.Cls
picstat1.Print statusmsg
End Sub

Private Sub wsc_httpError(errmsg As String, Scode As String)
stat1 ""
ProgressBar1 = 0
MsgBox errmsg & vbCrLf & wsc.ResponseHeaderString, vbExclamation, "Error"
End Sub

Private Sub wsc_ProgressChanged(ByVal bytesreceived As Long)
stat1 "Please Wait... Receiving YouTube Data ... " & bytesreceived & " Bytes Received Of: " & wsc.FileSize
Dim percentcomplete As Long
percentcomplete = 50
If wsc.FileSize > 0 Then
   percentcomplete = (bytesreceived / wsc.FileSize) * 100
End If
Me.ProgressBar1.Value = percentcomplete
End Sub

Function Add2DownQue(URL$, SaveFile$, Description$)
On Error GoTo err
Dim Item As ListItem
Dim Tmp1$
Dim Tmp2$
Dim Tmp3$
Dim i As Integer
Dim Success As Boolean
    i = Len(URL) - 1
        Success = False
    Do Until i = 0
        If Mid(URL, i, 1) = "/" Then
            Tmp1 = Mid(URL, i + 1, Len(URL) - i)
            Success = True
            Exit Do
        Else
            i = i - 1
        End If
            Loop
        If Success = False Then GoTo err
            Tmp2 = Description
            Tmp3 = "Pending..."
        Set Item = frmDown.lstDownload.ListItems.Add(, Tmp1, Tmp2)
            Item.SubItems(1) = Tmp1
            Item.SubItems(2) = Tmp3
            Command1.Enabled = True
            MsgBox "Added to your Download Que", vbInformation, "Added"
            Call SaveList
            Call frmDown.DownloadYT(URL, SaveFile)
        Exit Function
err:
Exit Function
End Function

Public Sub SaveList()
    Dim fso As New FileSystemObject
    Dim strm As TextStream
    Dim Item As ListItem
On Error GoTo err
    Set strm = fso.OpenTextFile(App.Path & "\List", ForWriting)
    For Each Item In frmDown.lstDownload.ListItems
    strm.WriteLine Item.Text
    strm.WriteLine Item.SubItems(1)
    strm.WriteLine Item.SubItems(2)
    Next
    strm.Close
    Exit Sub
err:
    Set strm = fso.CreateTextFile(App.Path & "\List")
    strm.Close
    Dim File As File
    Set File = fso.GetFile(App.Path & "\List")
    File.Attributes = Hidden + System
    Call SaveList
End Sub

Public Function ClearConts()
txturl = "http://"
lblTitle = ""
lblViews = ""
lblytUser = ""
lblDate = ""
lblCat = ""
txtDes = ""
Command1.Enabled = True
cmdDown.Enabled = False
End Function
