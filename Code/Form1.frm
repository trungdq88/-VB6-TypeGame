VERSION 5.00
Object = "{41F9B345-9609-4DFD-8D8C-32BAAF40C7AF}#1.0#0"; "UNIRICHEDIT.OCX"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmNoiDung 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Type - Game"
   ClientHeight    =   7920
   ClientLeft      =   4185
   ClientTop       =   2775
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10095
   StartUpPosition =   1  'CenterOwner
   Begin UniControls.UniLabel UniLabel5 
      Height          =   255
      Left            =   5040
      Top             =   7560
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Nha61n Phi1m TAB D9e63 Chuye63n Sang Khung Go4..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483631
   End
   Begin VB.Timer tmrChat2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6960
      Top             =   1680
   End
   Begin VB.Timer tmrChat1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7440
      Top             =   1200
   End
   Begin UniControls.UniLabel lblChat1 
      Height          =   255
      Left            =   3720
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      BackColor       =   12648447
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniButton cmdSend 
      Height          =   375
      Left            =   9120
      TabIndex        =   14
      Top             =   7440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Icon            =   "Form1.frx":617A
      Style           =   1
      Caption         =   "Gu73i"
      IconAlign       =   3
      Enabled         =   0   'False
      BackColor       =   15398133
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
   End
   Begin UniControls.UniLabel UniLabel4 
      Height          =   375
      Left            =   120
      Top             =   7440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Alignment       =   2
      Caption         =   ">>>>>>> Cha1t O73 D9a6y:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniTextBox txtChat 
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   7440
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   ""
   End
   Begin VB.Timer tmrEXIT 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   480
   End
   Begin UniControls.UniButton Button1 
      Height          =   375
      Left            =   9240
      TabIndex        =   12
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Icon            =   "Form1.frx":6196
      Style           =   1
      Caption         =   "Thoa1t"
      IconAlign       =   3
      MaskColor       =   16711935
      FontColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
   End
   Begin UniControls.UniLabel OtoName 
      Height          =   255
      Index           =   0
      Left            =   240
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrUU 
      Interval        =   500
      Left            =   3000
      Top             =   480
   End
   Begin UniControls.UniLabel UniLabel3 
      Height          =   255
      Left            =   5400
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      Caption         =   "So61 Ngu7o72i Co1 Trong Pho2ng Na2y:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Caption         =   "D9i5a Chi3 Pho2ng:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock WinSock1 
      Index           =   0
      Left            =   9600
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer DeleTrim 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   480
   End
   Begin UniControls.UniTabStrip Tab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7858
      TabCount        =   2
      TabCaption(0)   =   "Tab 0"
      TabContCtrlCnt(0)=   2
      Tab(0)ContCtrlCap(1)=   "List1"
      Tab(0)ContCtrlCap(2)=   "Box1"
      TabCaption(1)   =   "Tab 1"
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "UniLabel1"
      TabStyle        =   1
      TabTheme        =   1
      ActiveTabBackStartColor=   16514555
      ActiveTabBackEndColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin UniControls.UniLabel UniLabel1 
         Height          =   255
         Left            =   -72720
         Top             =   2040
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         BackColor       =   -2147483634
         Caption         =   "Chu71c Na8ng Chu7a D9u7o75c Ca65p Nha65t Trong Phie6n Ba3n Na2y"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483631
         Link            =   ""
      End
      Begin UniControls.UniListBox List1 
         Height          =   3735
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   6588
         IconMaskColor   =   16711935
         Picture         =   "Form1.frx":61B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   19
      End
      Begin UnicodeRichEdit.UniRichTextbox Box1 
         Height          =   3735
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6588
         Version         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   16744576
         ViewMode        =   1
         Border          =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         AutoURLDetect   =   0   'False
         Transparent     =   -1  'True
      End
   End
   Begin UniControls.UniButton Command1 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Icon            =   "Form1.frx":61CE
      Style           =   1
      IconAlign       =   3
      Enabled         =   0   'False
      BackColor       =   15398133
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
   End
   Begin TypeGame.UniTextBox YourTBSpeed 
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      Enabled         =   0   'False
      BorderStyle     =   0
   End
   Begin TypeGame.UniTextBox Text2 
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   6720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Text            =   ""
      Enabled         =   0   'False
   End
   Begin VB.Timer Speed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   480
   End
   Begin VB.Timer Start 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   480
   End
   Begin TypeGame.UniTextBox Uni1 
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      Enabled         =   0   'False
      BorderStyle     =   0
   End
   Begin UniControls.UniLabel OtoName 
      Height          =   255
      Index           =   1
      Left            =   240
      Top             =   1200
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniLabel lblChat2 
      Height          =   255
      Left            =   3960
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BackColor       =   12648447
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000016&
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Oto 
      Height          =   720
      Index           =   1
      Left            =   240
      Picture         =   "Form1.frx":61EA
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image Oto 
      Height          =   720
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":B9CC
      Top             =   600
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   180
      Left            =   240
      Picture         =   "Form1.frx":111AE
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   7890
   End
   Begin VB.Label lblOnl 
      BackColor       =   &H80000016&
      Caption         =   "0"
      Height          =   255
      Left            =   8280
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblIP 
      BackColor       =   &H80000016&
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label VTTB 
      BackColor       =   &H80000016&
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
      Left            =   7680
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image DuongDua 
      Height          =   180
      Left            =   240
      Picture         =   "Form1.frx":148E0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   7890
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   8760
      Picture         =   "Form1.frx":18012
      Stretch         =   -1  'True
      Top             =   600
      Width           =   405
   End
End
Attribute VB_Name = "frmNoiDung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x
Dim ThoiGianBatDau
Dim DVDua
Dim TimeSP
Dim SpeedGo
Dim DoDaiKT
Dim SpeedTB(1 To 100)
Dim STTGO
Dim VTB
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long
Dim iSock As Integer ' Luot Truy Cap
Dim Online As Integer ' NGuoi dang Online
Dim sUser(0 To 4) As String
Dim sSSS As Boolean
Dim strMe As String
Dim strYou As String
Dim StrNoConnect As String
Dim sYouOnline As Boolean
Dim sDangChoi As Boolean
Dim sExit As Boolean
 Dim strNhanVe
Dim sChat
Dim sYouWin As Boolean
Dim sYouDvD
Private Sub Button1_Click()
If sDangChoi = True Then
    If UniMsgBox("B" & ChrW$(&H1EA1) & "n " & ChrW$(&H110) & "ang Trong Cu" & ChrW$(&H1ED9) & "c Ch" & ChrW$(&H1A1) & "i, N" & ChrW$(&H1EBF) & "u Thoát Ra, B" & ChrW$(&H1EA1) & "n S" & ChrW$(&H1EBD) & " B" & ChrW$(&H1ECB) & " Thua" & vbCrLf & "B" & ChrW$(&H1EA1) & "n Ch" & ChrW$(&H1EAF) & "c Ch" & ChrW$(&H1EAF) & "n ?", vbCritical + vbYesNo, "Thông Báo", Me.hWnd) = vbYes Then
       sExitGame
        tmrEXIT.Enabled = True
    End If
Else
    If UniMsgBox(ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EAF) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EAF) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H75) & ChrW$(&H1ED1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&HE1) & ChrW$(&H74) & ChrW$(&H3F), vbQuestion + vbYesNo, "Thông Báo", Me.hWnd) = vbYes Then
        On Error Resume Next
        If WinSock1(1).State = sckConnected Then sExitGame 'Gui Thong Bao Thoat
        tmrEXIT.Enabled = True
        frmMain.Show
        
    End If

End If

End Sub

Private Sub cmdSend_Click()
On Error Resume Next
If txtChat.Text <> "" Then
Dim strSend
strSend = Unicode2UTF8(txtChat.Text)
If WinSock1(1).State = sckConnected Then WinSock1(1).SendData strSend & "CHA"


    lblChat1.Visible = True
    lblChat1.Left = Oto(0).Left
    lblChat1.Top = Oto(0).Top - lblChat1.Height
    lblChat1.Caption = txtChat.Text
    lblChat1.AutoSize = True
    tmrChat1.Enabled = False
    tmrChat1.Enabled = True
    txtChat.Text = ""
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = ChrW$(&H42) & ChrW$(&H1ECF) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H75) & ChrW$(&H1ED9) & ChrW$(&H63) Then
    If UniMsgBox(ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H68) & ChrW$(&H1EAF) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H68) & ChrW$(&H1EAF) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H42) & ChrW$(&H1ECF) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H75) & ChrW$(&H1ED9) & ChrW$(&H63) & ChrW$(&H3F), vbQuestion + vbYesNo, "Thông Báo", Me.hWnd) = vbYes Then
    KetThucCuocDua
    Command1.Caption = ChrW$(&H42) & ChrW$(&H1EAF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EA7) & ChrW$(&H75)
    If WinSock1(1).State = sckConnected Then WinSock1(1).SendData "123456EXT" Else UniMsgBox StrNoConnect
    
    sDangChoi = False
    End If
Else
If sSSS = False Then
    UniMsgBox ChrW$(&H4E) & ChrW$(&H67) & ChrW$(&H1B0) & ChrW$(&H1EDD) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1A1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&H1EB5) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&HE0) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbInformation + vbOKOnly, "Thông Báo", Me.hWnd
Else
'If WinSock1(0).State = sckConnected Then
If WinSock1(1).State = sckConnected Then WinSock1(1).SendData "123456STR" Else UniMsgBox StrNoConnect
OtoName(1).Caption = strYou

sDangChoi = True
Command1.Enabled = False
DVDua = DuongDua.Width / SoKyTu(Box1.Text)
sYouDvD = DVDua
lblStart.Visible = True
lblStart.Left = 0
lblStart.Top = 0
lblStart.Width = Me.Width
lblStart.Top = 120
lblStart.Height = 1695
lblStart.Caption = ThoiGianBatDau
Start.Enabled = True
Text2.Text = ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&H1EB5) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&HE0) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&HF5) & ChrW$(&H20) & ChrW$(&H21)
Command1.Caption = ChrW$(&H42) & ChrW$(&H1ECF) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H75) & ChrW$(&H1ED9) & ChrW$(&H63)
List1.Visible = False
Box1.Left = 120
Box1.Width = 9615
Oto(0).Left = 240
Oto(1).Left = 240

Tab1.Enabled = False
ToKyTu Box1.Text, 1
STTGO = 0
VTTB.Caption = ""
SpeedGo = 0
TimeSP = 1
End If
End If
End Sub





Private Sub DeleTrim_Timer()
Text2.Text = Trim(Text2.Text)
DeleTrim.Enabled = False

End Sub

Private Sub Form_Load()
sDangChoi = False
'************ Winsock Setting *************
frmNoiDung.WinSock1(0).LocalPort = 8818 'khai bao cong ket noi la cong 1306, co the khai bao cong khac tuy thik
frmNoiDung.WinSock1(0).Listen 'bat buoc phai kha bao dong nay
frmNoiDung.lblIP.Caption = frmNoiDung.WinSock1(0).LocalIP
'******************
StrNoConnect = "Không Có " & ChrW$(&H4B) & ChrW$(&H1EBF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H1ED1) & ChrW$(&H69) & ChrW$(&H21)
sSSS = False
Oto(1).Visible = False
sYouWin = False
Command1.Enabled = False
TimeSP = 1
x = 1
ThoiGianBatDau = 3
SpeedGo = 0
Box1.readOnly = True
Box1.Enabled = False
'DANH SACH
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\DOANVAN.MDB")
    Set db = DBEngine.OpenDatabase(DbFile, False, False)
    Set rs = db.OpenRecordset("DoanVan", dbOpenTable)
    Max = rs.RecordCount
    If rs.RecordCount = 0 Then
Exit Sub
Else
rs.MoveFirst
List1.Clear
For i = 1 To Max
    List1.AddItem rs!Ten
    rs.MoveNext
Next i
List1.ListIndex = 0
End If
'((((((((((((


'Set Uni Caption
YourTBSpeed.Text = ChrW$(&H56) & ChrW$(&H1EAD) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H1ED1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&HEC) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H3A)
Uni1.Text = ChrW$(&H54) & ChrW$(&H1EEB) & ChrW$(&H20) & ChrW$(&H2F) & ChrW$(&H20) & ChrW$(&H50) & ChrW$(&H68) & ChrW$(&HFA) & ChrW$(&H74)
Command1.Caption = ChrW$(&H42) & ChrW$(&H1EAF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EA7) & ChrW$(&H75)
Tab1.TabCaption(0) = ChrW$(&H47) & ChrW$(&HF5) & " Nhanh"
Tab1.TabCaption(1) = "Tính Nhanh"



End Sub





Private Sub Form_Unload(Cancel As Integer)
If sExit = False Then
Cancel = 1
UniMsgBox ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EA5) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HFA) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H22) & ChrW$(&H54) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&HE1) & ChrW$(&H74) & ChrW$(&H22) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H72) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1ECF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&HF2) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H2E), vbOKOnly, "Thông Báo", Me.hWnd
End If
End Sub


Private Sub List1_Click()
Command1.Enabled = True
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\DOANVAN.mdb")
    Set db = DBEngine.OpenDatabase(DbFile, False, False)
    Set rs = db.OpenRecordset("DoanVan", dbOpenTable)
    
    Set rs = db.OpenRecordset("Select * from DoanVan where TEN = '" & Trim(List1.List(List1.ListIndex)) & "'")
Box1.Text = rs.Fields("NOIDUNG")

'box1.SelFontColour = &HFF8080
Box1.ForeColor = &HFF8080
On Error Resume Next

WinSock1(1).SendData List1.ListIndex & "NDD"
End Sub





Private Sub Speed_Timer()
TimeSP = TimeSP + 1

End Sub

Private Sub Start_Timer()
On Error Resume Next
ThoiGianBatDau = ThoiGianBatDau - 1
lblStart.Caption = ThoiGianBatDau
If ThoiGianBatDau = 0 Then
Text2.Text = ""
Text2.Enabled = True
Text2.SetFocus
Start.Enabled = False
lblStart.Visible = False
Speed.Enabled = True
Command1.Enabled = True
End If
End Sub



Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then Unload frmGoSai
'On Error Resume Next
If KeyCode = 32 Then
    
    If Text2.Text = GetKyTu(Box1.Text, x) Then
    x = x + 1
                If x > SoKyTu(Box1.Text) Then
                
Oto(0).Left = Oto(0).Left + DVDua
                KetThucCuocDua

                End If
     'Khi Go Dung ********************
    Oto(0).Left = Oto(0).Left + DVDua
    WinSock1(1).SendData "123456GOO" ' Gui du lieu chung to Server dang di toi'
    Text2.BackColor = vbWhite

                     STTGO = STTGO + 1
                     SpeedGo = STTGO / TimeSP * 60
                     VTTB.Caption = Round(SpeedGo, 3)
                     
                     ToKyTu Box1.Text, STTGO + 1
    Text2.Text = ""
    Else
    frmGoSai.Show , frmNoiDung
    frmGoSai.Left = Text2.Left + Text2.Width
    frmGoSai.Top = Text2.Top - Text2.Height * 2
    frmGoSai.ChuSai.Caption = GetKyTu(Box1.Text, x)
    frmGoSai.lblC.Caption = "1"
    frmGoSai.Timer1.Enabled = True
    Text2.BackColor = vbRed
    Exit Sub
    End If
    'Text2.Text = Trim(Text2.Text)
    DeleTrim.Enabled = True
End If

End Sub
Private Sub KetThucCuocDua()
'*******************************
                '******** KET THUC CUOC DUA ***********
                
                x = 0
                DVDua = 0
                ThoiGianBatDau = 3
                Box1.Text = ""
                Text2.Text = ""
                Text2.Enabled = False
                'Oto(0).Left = 240
                List1.Visible = True
                Box1.Left = 2400
                Box1.Width = 7335
                Tab1.Enabled = True
                Speed.Enabled = False
                Command1.Caption = ChrW$(&H42) & ChrW$(&H1EAF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EA7) & ChrW$(&H75)
                MsgBox "You Are Winner !", vbOKOnly, "Win!"
            
End Sub


Private Sub tmrChat1_Timer()
lblChat1.Visible = False
tmrChat1.Enabled = False
End Sub

Private Sub tmrChat2_Timer()
lblChat2.Visible = False
tmrChat2.Enabled = False
End Sub

Private Sub tmrEXIT_Timer()
sExit = True
Unload Me
End Sub

Private Sub tmrUU_Timer()
Dim TT
For TT = 0 To 1
    OtoName(TT).Left = Oto(TT).Left
Next TT

End Sub

Private Sub UniButton1_Click()

End Sub

Private Sub txtChat_Changed()
If txtChat.Text = "" Then
    cmdSend.Enabled = False
Else
    cmdSend.Enabled = True
End If

End Sub

Private Sub txtChat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSend_Click
End Sub

Private Sub Winsock1_ConnectionRequest(index As Integer, ByVal requestID As Long)
On Error Resume Next
WinSock1(1).Close
Load WinSock1(1) 'load ket noi 1 ve server
WinSock1(1).Accept requestID 'xac nhan ket noi tu client (ket noi 1)
End Sub
Private Sub Winsock1_DataArrival(index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim sData As String
Dim Data3 As String
Dim Data2 As String
Dim sDataOut As String
WinSock1(index).GetData sData, vbString 'nhan du lieu tu client
'***********************'
'Xu Ly Du Lieu O Day ...'
Data3 = Right(sData, 3)
Data2 = Left(sData, Len(sData) - 3)
If Data3 = "DKY" Then
        If sYouOnline = True Then
            WinSock1(index).SendData "123456QAT"
        Else
            If UniMsgBox(Data2 & " Xin Phép Tham Gia Phòng " & ChrW$(&H43) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H2E) & vbCrLf & ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&H1ED3) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & "Ý Cho Tham Gia Không?", vbInformation + vbYesNo, "Thông Báo", Me.hWnd) = vbYes Then
                If WinSock1(index).State = sckConnected Then WinSock1(index).SendData OtoName(0).Caption & "DOY" Else UniMsgBox StrNoConnect
                Oto(1).Visible = True
                OtoName(1).Caption = Data2
                strYou = Data2
                sYouOnline = True
            Else
                If WinSock1(index).State = sckConnected Then WinSock1(index).SendData "12345KDY" Else UniMsgBox StrNoConnect
            End If
        End If
ElseIf Data3 = "EXT" Then
    OtoName(1).Caption = strYou & " (Ðã B" & ChrW$(&H1ECF) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H75) & ChrW$(&H1ED9) & ChrW$(&H63) & ChrW$(&H29)
    UniMsgBox strYou & " Ðã B" & ChrW$(&H1ECF) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H75) & ChrW$(&H1ED9) & ChrW$(&H63), , , Me.hWnd
    KetThucCuocDua
    sSSS = False
    sDangChoi = False
ElseIf Data3 = "SSS" Then
    OtoName(1).Caption = strYou & " (Ðã S" & ChrW$(&H1EB5) & "n Sàng)"
    sSSS = True
ElseIf Data3 = "GOO" Then
    Oto(1).Left = Oto(1).Left + sYouDvD
ElseIf Data3 = "QUI" Then
    sYouOnline = False
    OtoName(1).Caption = "(Ðã Thoát Ra)"
    sSSS = False
    Oto(1).Visible = False
    If sDangChoi = True Then
    KetThucCuocDua
    Command1.Caption = ChrW$(&H42) & ChrW$(&H1EAF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EA7) & ChrW$(&H75)
    sDangChoi = False
    End If
ElseIf Data3 = "CHA" Then
   
    strNhanVe = UTF82Unicode(Data2)
    lblChat2.Visible = True
    lblChat2.Left = Oto(1).Left
    lblChat2.Top = Image3.Top + Image3.Height
    lblChat2.Caption = strNhanVe
    lblChat2.AutoSize = True
    tmrChat2.Enabled = False
    tmrChat2.Enabled = True

End If
End Sub

Private Sub WinWin()
MsgBox "You Are Winner"
End Sub
' Chu Thich Cac Mau Du Lieu O Data3
'QAT - Qua Tai
'DOY - Dong Y
'KDY - Ko Dong Y
'DKY - Dang Ky
'US1 - Nguoi Choi 1
Private Sub sExitGame()
If WinSock1(1).State = sckConnected Then WinSock1(1).SendData "123456QUI"  'Gui Thong Bao Thoat
End Sub
