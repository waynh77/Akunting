VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form cetakAkun_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9405
   ClientLeft      =   15
   ClientTop       =   1755
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   15330
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "KELUAR"
      DownPicture     =   "cetakAkun_frm.frx":0000
      Height          =   855
      Index           =   1
      Left            =   13440
      MouseIcon       =   "cetakAkun_frm.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "cetakAkun_frm.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CETAK"
      DownPicture     =   "cetakAkun_frm.frx":1C9E
      Height          =   855
      Index           =   0
      Left            =   11640
      MouseIcon       =   "cetakAkun_frm.frx":2968
      MousePointer    =   99  'Custom
      Picture         =   "cetakAkun_frm.frx":2C72
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Index           =   1
      Left            =   13440
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Index           =   1
      Left            =   12480
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Index           =   1
      Left            =   8400
      Top             =   480
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Index           =   1
      Left            =   7440
      Top             =   480
      Width           =   975
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7095
      Index           =   7
      Left            =   8400
      TabIndex        =   18
      Top             =   1080
      Width           =   4095
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7223;12515"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7095
      Index           =   6
      Left            =   11760
      TabIndex        =   17
      Top             =   1080
      Width           =   1695
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;12515"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7095
      Index           =   5
      Left            =   12840
      TabIndex        =   16
      Top             =   1080
      Width           =   1815
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3201;12515"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7095
      Index           =   4
      Left            =   7440
      TabIndex        =   15
      Top             =   1080
      Width           =   1935
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3413;12515"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NERACA/ LABA-RUGI"
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
      Left            =   13320
      TabIndex        =   14
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEBIT/ KREDIT"
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
      Height          =   585
      Index           =   6
      Left            =   12360
      TabIndex        =   13
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   315
      Index           =   5
      Left            =   8400
      TabIndex        =   12
      Top             =   600
      Width           =   4080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO. AKUN"
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
      Height          =   435
      Index           =   4
      Left            =   7440
      TabIndex        =   11
      Top             =   600
      Width           =   840
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Index           =   0
      Left            =   6240
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Index           =   0
      Left            =   5280
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Index           =   0
      Left            =   1200
      Top             =   480
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Index           =   0
      Left            =   240
      Top             =   480
      Width           =   975
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7095
      Index           =   3
      Left            =   6240
      TabIndex        =   10
      Top             =   1080
      Width           =   1815
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3201;12515"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7095
      Index           =   2
      Left            =   5280
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;12515"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7095
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   1080
      Width           =   4095
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7223;12515"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7095
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3836;12515"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NERACA/ LABA-RUGI"
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
      Left            =   6120
      TabIndex        =   6
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEBIT/ KREDIT"
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
      Height          =   585
      Index           =   2
      Left            =   5160
      TabIndex        =   5
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   4080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO. AKUN"
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
      Height          =   435
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TABEL AKUN"
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1590
   End
End
Attribute VB_Name = "cetakAkun_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Command1(0).Visible = False
    Command1(1).Visible = False
    CommonDialog1.ShowPrinter
    cetakAkun_frm.PrintForm
    Command1(0).Visible = True
    Command1(1).Visible = True
Case 1
    AkunTbl_frm.Enabled = True
    Unload Me
    AkunTbl_frm.Show
End Select
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data1.RecordSource = "akuntbl"
Data1.Refresh
isi_data
End Sub

Sub isi_data()
Dim a As Single
With Data1.Recordset
ListBox1(0).Clear
ListBox1(1).Clear
ListBox1(2).Clear
ListBox1(3).Clear
If Not .BOF Then
    .MoveFirst
    a = 0
    Do While Not .EOF
        If a < 36 Then
            ListBox1(0).AddItem !no_akun
            ListBox1(1).AddItem !nm_akun
            ListBox1(2).AddItem !dk
            ListBox1(3).AddItem !nrlr
            a = a + 1
        End If
        If a >= 36 Then
            ListBox1(4).AddItem !no_akun
            ListBox1(7).AddItem !nm_akun
            ListBox1(6).AddItem !dk
            ListBox1(5).AddItem !nrlr
            a = a + 1
        End If
        .MoveNext
    Loop
End If
End With
End Sub
