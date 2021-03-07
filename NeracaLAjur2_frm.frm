VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form NeracaLAjur2_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9270
   ClientLeft      =   15
   ClientTop       =   1755
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   15330
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   44
      Text            =   "Text3"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   43
      Text            =   "Text2"
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "BACK"
      DownPicture     =   "NeracaLAjur2_frm.frx":0000
      Height          =   735
      Index           =   2
      Left            =   10200
      MouseIcon       =   "NeracaLAjur2_frm.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "NeracaLAjur2_frm.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   11
      Left            =   13560
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   10
      Left            =   11880
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   9
      Left            =   13560
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   8
      Left            =   11880
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   10200
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   8520
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   13560
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   11880
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   10200
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   8520
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   6360
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   4800
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "KELUAR"
      DownPicture     =   "NeracaLAjur2_frm.frx":2956
      Height          =   735
      Index           =   1
      Left            =   13560
      MouseIcon       =   "NeracaLAjur2_frm.frx":3620
      MousePointer    =   99  'Custom
      Picture         =   "NeracaLAjur2_frm.frx":392A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CETAK"
      DownPicture     =   "NeracaLAjur2_frm.frx":45F4
      Height          =   735
      Index           =   0
      Left            =   11880
      MouseIcon       =   "NeracaLAjur2_frm.frx":52BE
      MousePointer    =   99  'Custom
      Picture         =   "NeracaLAjur2_frm.frx":55C8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape6 
      Height          =   6735
      Left            =   4800
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KESEIMBANGAN"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   16
      Left            =   9960
      TabIndex        =   29
      Top             =   8760
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABA-RUGI"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   15
      Left            =   7200
      TabIndex        =   28
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KESEIMBANGAN"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   14
      Left            =   3000
      TabIndex        =   27
      Top             =   8040
      Width           =   1680
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   6495
      Index           =   9
      Left            =   13440
      TabIndex        =   26
      Top             =   1440
      Width           =   1695
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;11456"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   6495
      Index           =   8
      Left            =   11760
      TabIndex        =   25
      Top             =   1440
      Width           =   1695
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;11456"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   6495
      Index           =   7
      Left            =   10080
      TabIndex        =   24
      Top             =   1440
      Width           =   1695
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;11456"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   6495
      Index           =   6
      Left            =   8400
      TabIndex        =   23
      Top             =   1440
      Width           =   1695
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;11456"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   6495
      Index           =   5
      Left            =   7920
      TabIndex        =   22
      Top             =   1440
      Width           =   1695
      VariousPropertyBits=   746588185
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;11456"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   6495
      Index           =   4
      Left            =   6360
      TabIndex        =   21
      Top             =   1440
      Width           =   1575
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;11456"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   6495
      Index           =   3
      Left            =   4800
      TabIndex        =   20
      Top             =   1440
      Width           =   1575
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;11095"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   6495
      Index           =   2
      Left            =   4080
      TabIndex        =   19
      Top             =   1440
      Width           =   1695
      VariousPropertyBits=   746588185
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;11456"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   6495
      Index           =   1
      Left            =   1200
      TabIndex        =   18
      Top             =   1440
      Width           =   2895
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "5106;11456"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   6495
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Width           =   1575
      VariousPropertyBits=   746588185
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;11456"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape13 
      Height          =   6735
      Left            =   13440
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Shape Shape12 
      Height          =   6735
      Left            =   11760
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Shape Shape11 
      Height          =   6735
      Left            =   10080
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Shape Shape10 
      Height          =   6735
      Left            =   8400
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Shape Shape9 
      Height          =   6975
      Left            =   8400
      Top             =   960
      Width           =   3375
   End
   Begin VB.Shape Shape8 
      Height          =   6975
      Left            =   7920
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape7 
      Height          =   6735
      Left            =   6360
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape Shape5 
      Height          =   6975
      Left            =   4800
      Top             =   960
      Width           =   3135
   End
   Begin VB.Shape Shape4 
      Height          =   6975
      Left            =   4080
      Top             =   960
      Width           =   735
   End
   Begin VB.Shape Shape3 
      Height          =   6975
      Left            =   1200
      Top             =   960
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      Height          =   6975
      Left            =   120
      Top             =   960
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   120
      Top             =   960
      Width           =   15015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KREDIT"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   13
      Left            =   13920
      TabIndex        =   16
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEBIT"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   12
      Left            =   12360
      TabIndex        =   15
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NERACA"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   11
      Left            =   12240
      TabIndex        =   14
      Top             =   960
      Width           =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KREDIT"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   10
      Left            =   10680
      TabIndex        =   13
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEBIT"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   9
      Left            =   9000
      TabIndex        =   12
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABA-RUGI"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   8
      Left            =   8880
      TabIndex        =   11
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NR LR"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   465
      Index           =   7
      Left            =   7920
      TabIndex        =   10
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KREDIT"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   6
      Left            =   6720
      TabIndex        =   9
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEBIT"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   5
      Left            =   5280
      TabIndex        =   8
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NERACA SALDO"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   4
      Left            =   5040
      TabIndex        =   7
      Top             =   960
      Width           =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AKUN D/K"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   465
      Index           =   3
      Left            =   4080
      TabIndex        =   6
      Top             =   960
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA AKUN"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO.AKUN"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE : "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NERACA LAJUR"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "NeracaLAjur2_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Command1(0).Visible = False
    Command1(1).Visible = False
    Command1(2).Visible = False
    CommonDialog1.ShowPrinter
    NeracaLAjur2_frm.PrintForm
    Command1(0).Visible = True
    Command1(1).Visible = True
    Command1(2).Visible = True
    
Case 1
    Akuntansi_frm.Enabled = True
    Unload Me
    Akuntansi_frm.Show
Case 2
    Me.Hide
    NeracaLAjur_frm.Show
End Select
End Sub



Sub kosong()
Dim a As Byte
a = 0
Do Until a = 10
    ListBox1(a).Clear
    a = a + 1
Loop
a = 0
Do Until a = 12
    Text1(a) = ""
    a = a + 1
Loop
End Sub

Private Sub ListBox1_Click(Index As Integer)
Select Case Index
Case 0
    ListBox1(1).ListIndex = ListBox1(0).ListIndex
    ListBox1(2).ListIndex = ListBox1(0).ListIndex
    ListBox1(3).ListIndex = ListBox1(0).ListIndex
    ListBox1(4).ListIndex = ListBox1(0).ListIndex
    ListBox1(5).ListIndex = ListBox1(0).ListIndex
    ListBox1(6).ListIndex = ListBox1(0).ListIndex
    ListBox1(7).ListIndex = ListBox1(0).ListIndex
    ListBox1(8).ListIndex = ListBox1(0).ListIndex
    ListBox1(9).ListIndex = ListBox1(0).ListIndex
Case 1
    ListBox1(0).ListIndex = ListBox1(1).ListIndex
    ListBox1(2).ListIndex = ListBox1(1).ListIndex
    ListBox1(3).ListIndex = ListBox1(1).ListIndex
    ListBox1(4).ListIndex = ListBox1(1).ListIndex
    ListBox1(5).ListIndex = ListBox1(1).ListIndex
    ListBox1(6).ListIndex = ListBox1(1).ListIndex
    ListBox1(7).ListIndex = ListBox1(1).ListIndex
    ListBox1(8).ListIndex = ListBox1(1).ListIndex
    ListBox1(9).ListIndex = ListBox1(1).ListIndex
Case 2
    ListBox1(1).ListIndex = ListBox1(2).ListIndex
    ListBox1(0).ListIndex = ListBox1(2).ListIndex
    ListBox1(3).ListIndex = ListBox1(2).ListIndex
    ListBox1(4).ListIndex = ListBox1(2).ListIndex
    ListBox1(5).ListIndex = ListBox1(2).ListIndex
    ListBox1(6).ListIndex = ListBox1(2).ListIndex
    ListBox1(7).ListIndex = ListBox1(2).ListIndex
    ListBox1(8).ListIndex = ListBox1(2).ListIndex
    ListBox1(9).ListIndex = ListBox1(2).ListIndex
Case 3
    ListBox1(1).ListIndex = ListBox1(3).ListIndex
    ListBox1(2).ListIndex = ListBox1(3).ListIndex
    ListBox1(0).ListIndex = ListBox1(3).ListIndex
    ListBox1(4).ListIndex = ListBox1(3).ListIndex
    ListBox1(5).ListIndex = ListBox1(3).ListIndex
    ListBox1(6).ListIndex = ListBox1(3).ListIndex
    ListBox1(7).ListIndex = ListBox1(3).ListIndex
    ListBox1(8).ListIndex = ListBox1(3).ListIndex
    ListBox1(9).ListIndex = ListBox1(3).ListIndex
Case 4
    ListBox1(1).ListIndex = ListBox1(4).ListIndex
    ListBox1(2).ListIndex = ListBox1(4).ListIndex
    ListBox1(3).ListIndex = ListBox1(4).ListIndex
    ListBox1(0).ListIndex = ListBox1(4).ListIndex
    ListBox1(5).ListIndex = ListBox1(4).ListIndex
    ListBox1(6).ListIndex = ListBox1(4).ListIndex
    ListBox1(7).ListIndex = ListBox1(4).ListIndex
    ListBox1(8).ListIndex = ListBox1(4).ListIndex
    ListBox1(9).ListIndex = ListBox1(4).ListIndex
Case 5
    ListBox1(1).ListIndex = ListBox1(5).ListIndex
    ListBox1(2).ListIndex = ListBox1(5).ListIndex
    ListBox1(3).ListIndex = ListBox1(5).ListIndex
    ListBox1(4).ListIndex = ListBox1(5).ListIndex
    ListBox1(0).ListIndex = ListBox1(5).ListIndex
    ListBox1(6).ListIndex = ListBox1(5).ListIndex
    ListBox1(7).ListIndex = ListBox1(5).ListIndex
    ListBox1(8).ListIndex = ListBox1(5).ListIndex
    ListBox1(9).ListIndex = ListBox1(5).ListIndex
Case 6
    ListBox1(1).ListIndex = ListBox1(6).ListIndex
    ListBox1(2).ListIndex = ListBox1(6).ListIndex
    ListBox1(3).ListIndex = ListBox1(6).ListIndex
    ListBox1(4).ListIndex = ListBox1(6).ListIndex
    ListBox1(5).ListIndex = ListBox1(6).ListIndex
    ListBox1(0).ListIndex = ListBox1(6).ListIndex
    ListBox1(7).ListIndex = ListBox1(6).ListIndex
    ListBox1(8).ListIndex = ListBox1(6).ListIndex
    ListBox1(9).ListIndex = ListBox1(6).ListIndex
Case 7
    ListBox1(1).ListIndex = ListBox1(7).ListIndex
    ListBox1(2).ListIndex = ListBox1(7).ListIndex
    ListBox1(3).ListIndex = ListBox1(7).ListIndex
    ListBox1(4).ListIndex = ListBox1(7).ListIndex
    ListBox1(5).ListIndex = ListBox1(7).ListIndex
    ListBox1(6).ListIndex = ListBox1(7).ListIndex
    ListBox1(0).ListIndex = ListBox1(7).ListIndex
    ListBox1(8).ListIndex = ListBox1(7).ListIndex
    ListBox1(9).ListIndex = ListBox1(7).ListIndex
Case 8
    ListBox1(1).ListIndex = ListBox1(8).ListIndex
    ListBox1(2).ListIndex = ListBox1(8).ListIndex
    ListBox1(3).ListIndex = ListBox1(8).ListIndex
    ListBox1(4).ListIndex = ListBox1(8).ListIndex
    ListBox1(5).ListIndex = ListBox1(8).ListIndex
    ListBox1(6).ListIndex = ListBox1(8).ListIndex
    ListBox1(7).ListIndex = ListBox1(8).ListIndex
    ListBox1(0).ListIndex = ListBox1(8).ListIndex
    ListBox1(9).ListIndex = ListBox1(8).ListIndex
Case 9
    ListBox1(1).ListIndex = ListBox1(9).ListIndex
    ListBox1(2).ListIndex = ListBox1(9).ListIndex
    ListBox1(3).ListIndex = ListBox1(9).ListIndex
    ListBox1(4).ListIndex = ListBox1(9).ListIndex
    ListBox1(5).ListIndex = ListBox1(9).ListIndex
    ListBox1(6).ListIndex = ListBox1(9).ListIndex
    ListBox1(7).ListIndex = ListBox1(9).ListIndex
    ListBox1(8).ListIndex = ListBox1(9).ListIndex
    ListBox1(0).ListIndex = ListBox1(9).ListIndex
End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
