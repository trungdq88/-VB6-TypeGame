VERSION 5.00
Begin VB.Form frmPlay 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Play Game !"
   ClientHeight    =   3090
   ClientLeft      =   4200
   ClientTop       =   5595
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   1560
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   1920
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
