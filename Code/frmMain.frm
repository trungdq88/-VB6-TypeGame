VERSION 5.00
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin UniControls.UniLabel UniLabel3 
      Height          =   255
      Left            =   120
      Top             =   2160
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "Vui Lo2ng Nha65p Tho6ng Tin Va2o Ca1c O6 D9u7o75c To6 D9o3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      Link            =   ""
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   120
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Te6n Hie63n Thi5 Cu3a Ba5n:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   255
      Left            =   120
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "D9i5a Chi3 Pho2ng Tham Gia:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniButton Button3 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Icon            =   "frmMain.frx":0000
      Style           =   1
      Caption         =   "Thu75c Hie65n"
      IconAlign       =   3
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.TextBox txtName 
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
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtDiaChi 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin UniControls.UniButton Button2 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      Icon            =   "frmMain.frx":001C
      Style           =   1
      Caption         =   "Tham Gia Pho2ng Kha1c"
      IconAlign       =   3
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
   End
   Begin UniControls.UniButton Button1 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Icon            =   "frmMain.frx":0038
      Style           =   1
      Caption         =   "Ta5o Pho2ng"
      IconAlign       =   3
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
   End
   Begin UniControls.UniButton Button4 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Icon            =   "frmMain.frx":0054
      Style           =   1
      Caption         =   "Thoa1t"
      IconAlign       =   3
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AppPath As String

Private Sub Button1_Click()
Me.Caption = "Create Room"
txtDiaChi.Enabled = False
txtDiaChi.BackColor = vbWhite
frmMain.txtName.Enabled = True
frmMain.txtName.BackColor = vbRed
frmMain.Button3.Enabled = True
frmMain.txtName.SetFocus

End Sub
Function FileExists(sFile As String) As Boolean
 On Error Resume Next
 FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function
Private Sub Button2_Click()
Me.Caption = "Join Room"
txtDiaChi.Enabled = True
frmMain.txtDiaChi.BackColor = vbRed
txtName.Enabled = True
frmMain.txtName.BackColor = vbRed
Button3.Enabled = True
txtDiaChi.SetFocus
End Sub

Private Sub Button3_Click()
'"Join Room"
'"Create Room"
If Me.Caption = "Join Room" Then
    If txtDiaChi.Text <> "" And txtName.Text <> "" Then
        frmThamGia.Show
        frmThamGia.WinSock1.Close
        frmThamGia.WinSock1.Connect frmMain.txtDiaChi.Text, 8818
        Me.Hide
    End If
ElseIf Me.Caption = "Create Room" Then
    If txtName.Text <> "" Then
        frmNoiDung.Show
        frmNoiDung.OtoName(0).Caption = txtName.Text
        frmNoiDung.lblName.Caption = txtName.Text
        Unload frmMain
    End If
End If
End Sub

Private Sub Button4_Click()
End
End Sub


Private Sub Form_Load()

AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
If FileExists(AppPath & "\DOANVAN.MDB") = False Then
UniMsgBox "Không tìm th" & ChrW$(&H1EA5) & "y c" & ChrW$(&H1A1) & " s" & ChrW$(&H1EDF) & " d" & ChrW$(&H1EEF) & " li" & ChrW$(&H1EC7) & "u ! Xin th" & ChrW$(&H1EED) & " l" & ChrW$(&H1EA1) & "i sau vài giây !"
End
Else
txtDiaChi.Enabled = False
txtName.Enabled = False
Button3.Enabled = False

End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Button3_Click
End Sub

