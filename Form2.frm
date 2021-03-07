VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3690
   ClientLeft      =   12015
   ClientTop       =   6570
   ClientWidth     =   3030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   3030
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1920
      Top             =   1560
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form2.frx":0000
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Telp.     021 9389 4481"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "YM     : Wahyu_nhidayat@yahoo.com"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail : Wahyunhidayat@gmail.com"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wahyu Nur Hidayat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   1200
      Picture         =   "Form2.frx":00D4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   825
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
