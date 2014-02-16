VERSION 5.00
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmGoSai 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   Picture         =   "frmGoSai.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   1560
   End
   Begin UniControls.UniLabel ChuSai 
      Height          =   495
      Left            =   840
      Top             =   1200
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      Alignment       =   1
      BackStyle       =   0
      Caption         =   "UniLabel3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Link            =   ""
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   375
      Left            =   1320
      Top             =   1800
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BackStyle       =   0
      Caption         =   "Ha4y Xoa1 D9i Và Vie61t La5i Cho D9u1ng!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      Link            =   ""
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   255
      Left            =   1920
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Ba5n D9a4 Go4 Sai Chu74:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin VB.Label lblC 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmGoSai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    SetWindow Me.hWnd, &HFFFFFF, 0, LWA_COLORKEY
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Timer1_Timer()

If lblC.Caption = 1 Then frmNoiDung.Text2.SetFocus
If lblC.Caption = 0 Then frmThamGia.Text2.SetFocus

Timer1.Enabled = False
End Sub
