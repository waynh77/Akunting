VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Neraca_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9420
   ClientLeft      =   15
   ClientTop       =   1590
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   15330
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CETAK"
      DownPicture     =   "Neraca_frm.frx":0000
      Height          =   855
      Index           =   0
      Left            =   11760
      MouseIcon       =   "Neraca_frm.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "Neraca_frm.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "KELUAR"
      DownPicture     =   "Neraca_frm.frx":1C9E
      Height          =   855
      Index           =   1
      Left            =   13560
      MouseIcon       =   "Neraca_frm.frx":2968
      MousePointer    =   99  'Custom
      Picture         =   "Neraca_frm.frx":2C72
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   480
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   735
      Index           =   17
      Left            =   12720
      TabIndex        =   47
      Top             =   4440
      Width           =   2415
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "4260;1296"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   735
      Index           =   16
      Left            =   8640
      TabIndex        =   46
      Top             =   4440
      Width           =   4095
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7223;1296"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   735
      Index           =   15
      Left            =   7560
      TabIndex        =   45
      Top             =   4440
      Width           =   2295
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "4048;1296"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   24
      Left            =   12960
      TabIndex        =   44
      Top             =   5760
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   23
      Left            =   12960
      TabIndex        =   43
      Top             =   5400
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   22
      Left            =   12960
      TabIndex        =   42
      Top             =   5160
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   21
      Left            =   12960
      TabIndex        =   41
      Top             =   3120
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   20
      Left            =   5400
      TabIndex        =   40
      Top             =   9000
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   19
      Left            =   5400
      TabIndex        =   39
      Top             =   8640
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   18
      Left            =   5400
      TabIndex        =   38
      Top             =   8400
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   17
      Left            =   5400
      TabIndex        =   37
      Top             =   6840
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   16
      Left            =   5400
      TabIndex        =   36
      Top             =   5160
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   15
      Left            =   11880
      TabIndex        =   35
      Top             =   5760
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH EKUITAS"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   14
      Left            =   8880
      TabIndex        =   34
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABA DITAHAN"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   13
      Left            =   8880
      TabIndex        =   33
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH KEWAJIBAN LANCAR"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   12
      Left            =   8880
      TabIndex        =   32
      Top             =   3120
      Width           =   2925
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   11
      Left            =   4560
      TabIndex        =   31
      Top             =   9000
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NILAI BUKU AKTIVA TETAP"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   10
      Left            =   1560
      TabIndex        =   30
      Top             =   8640
      Width           =   2685
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH AKUM. PENYUSUTAN"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   9
      Left            =   1560
      TabIndex        =   29
      Top             =   8400
      Width           =   2970
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH HARGA PEROLEHAN"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   8
      Left            =   1560
      TabIndex        =   28
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH AKTIVA LANCAR"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Index           =   7
      Left            =   1560
      TabIndex        =   27
      Top             =   5160
      Width           =   2505
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   855
      Index           =   14
      Left            =   12720
      TabIndex        =   26
      Top             =   3480
      Width           =   2415
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "4260;1508"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   855
      Index           =   13
      Left            =   8640
      TabIndex        =   25
      Top             =   3480
      Width           =   4095
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7223;1508"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   855
      Index           =   12
      Left            =   7560
      TabIndex        =   24
      Top             =   3480
      Width           =   1815
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3201;1508"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   1695
      Index           =   11
      Left            =   12720
      TabIndex        =   23
      Top             =   1440
      Width           =   2415
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "4260;2990"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   1695
      Index           =   10
      Left            =   8640
      TabIndex        =   22
      Top             =   1440
      Width           =   4095
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7223;2990"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   1695
      Index           =   9
      Left            =   7560
      TabIndex        =   21
      Top             =   1440
      Width           =   1935
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3413;2990"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   1215
      Index           =   8
      Left            =   5400
      TabIndex        =   20
      Top             =   7200
      Width           =   2175
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3836;2143"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   1215
      Index           =   7
      Left            =   1320
      TabIndex        =   19
      Top             =   7200
      Width           =   4095
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7223;2143"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   1215
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   7200
      Width           =   1695
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;2143"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   1335
      Index           =   5
      Left            =   5400
      TabIndex        =   17
      Top             =   5520
      Width           =   2175
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3836;2355"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   1335
      Index           =   4
      Left            =   1320
      TabIndex        =   16
      Top             =   5520
      Width           =   4095
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7223;2355"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   1335
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   5520
      Width           =   1695
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;2355"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   3735
      Index           =   2
      Left            =   5400
      TabIndex        =   14
      Top             =   1440
      Width           =   2175
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3836;6588"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   3735
      Index           =   1
      Left            =   1320
      TabIndex        =   13
      Top             =   1440
      Width           =   4095
      VariousPropertyBits=   746586139
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7223;6588"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   3735
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
      VariousPropertyBits=   746586137
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3201;6588"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Height          =   315
      Index           =   6
      Left            =   7560
      TabIndex        =   11
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "POS-POS"
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
      Left            =   10080
      TabIndex        =   10
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOMINAL"
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
      Index           =   4
      Left            =   12720
      TabIndex        =   9
      Top             =   960
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   7560
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   1
      Left            =   8640
      Top             =   840
      Width           =   4095
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Index           =   1
      Left            =   12720
      Top             =   840
      Width           =   2415
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
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "POS-POS"
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
      Index           =   2
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOMINAL"
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
      Index           =   3
      Left            =   5400
      TabIndex        =   6
      Top             =   960
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   0
      Left            =   1320
      Top             =   840
      Width           =   4095
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Index           =   0
      Left            =   5400
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NERACA"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1065
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
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Neraca_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Data1.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "'"
Data1.Refresh
isi_data
End Sub

Private Sub Combo2_Click()
Data1.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "'"
Data1.Refresh
isi_data
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Command1(0).Visible = False
    Command1(1).Visible = False
    CommonDialog1.ShowPrinter
    Neraca_frm.PrintForm
    Command1(0).Visible = True
    Command1(1).Visible = True
Case 1
    Akuntansi_frm.Enabled = True
    Unload Me
    Akuntansi_frm.Show
End Select
End Sub

Sub isi_cmb1()
Dim a As Date
Dim b As Byte
Combo1.Clear
b = 1
a = 15
Do Until b = 13
    Combo1.AddItem (Format(a, "mmmm"))
    b = b + 1
    a = a + 31
Loop
Combo1.ListIndex = Month(Date) - 1
End Sub

Sub isi_cmb2()
Dim a, b As Single
a = 1
b = 2008
Combo2.Clear
Do Until a = 50
    Combo2.AddItem (b)
    a = a + 1
    b = b + 1
Loop
Combo2 = Format(Date, "yyyy")
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Activate()
Data1.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "'"
Data1.Refresh
Data2.RecordSource = "select * from akuntbl order by no_akun asc"
Data2.Refresh
isi_cmb1
isi_cmb2
isi_data
End Sub

Sub kosong()
Dim a As Byte
a = 0
Do Until a = 18
    ListBox1(a).Clear
    a = a + 1
Loop
End Sub

Sub isi_data()
Dim a As Boolean
Dim b, c As Double
Dim d, e As Double
Dim jmlAktivaLancar, jmlHrgPerolehan, jmlPeny As Double
Dim modal, jmlAktiva, jmlhtgLancar As Double
Dim jmlEkuitas, htgJkPanjang, jmlKewajibanEkuitas As Double
b = 0
c = 0
d = 0
e = 0
kosong
With Data2.Recordset
    If Not .BOF And Not Data1.Recordset.BOF Then
        .MoveFirst
        Do While Not .EOF
            If Trim(Mid(.Fields("no_akun"), 1, 3)) = "1-1" Or Trim(Mid(.Fields("no_akun"), 1, 3)) = "1-0" Then
                a = False
                ListBox1(0).AddItem !no_akun
                ListBox1(1).AddItem !nm_akun
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        ListBox1(2).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.00")
                        jmlAktivaLancar = jmlAktivaLancar + Val(Data1.Recordset!saldo_akhir)
                        Data1.Recordset.MoveLast
                        a = True
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    ListBox1(2).AddItem ""
                End If
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 4)) = "1-20" Or Trim(Mid(.Fields("no_akun"), 1, 4)) = "1-21" Or Trim(Mid(.Fields("no_akun"), 1, 4)) = "1-22" Or Trim(Mid(.Fields("no_akun"), 1, 4)) = "1-23" Or Trim(Mid(.Fields("no_akun"), 1, 4)) = "1-24" Or Trim(Mid(.Fields("no_akun"), 1, 4)) = "1-25" Or Trim(Mid(.Fields("no_akun"), 1, 4)) = "1-26" Then
                a = False
                ListBox1(3).AddItem !no_akun
                ListBox1(4).AddItem !nm_akun
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        ListBox1(5).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.00")
                        jmlHrgPerolehan = jmlHrgPerolehan + Val(Data1.Recordset!saldo_akhir)
                        Data1.Recordset.MoveLast
                        a = True
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    ListBox1(5).AddItem ""
                End If
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 4)) = "1-27" Or Trim(Mid(.Fields("no_akun"), 1, 4)) = "1-28" Or Trim(Mid(.Fields("no_akun"), 1, 4)) = "1-29" Then
                a = False
                ListBox1(6).AddItem !no_akun
                ListBox1(7).AddItem !nm_akun
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        ListBox1(8).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.00")
                        jmlPeny = jmlPeny + Val(Data1.Recordset!saldo_akhir)
                        Data1.Recordset.MoveLast
                        a = True
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    ListBox1(8).AddItem ""
                End If
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 3)) = "2-1" Or Trim(Mid(.Fields("no_akun"), 1, 3)) = "2-0" Then
                a = False
                ListBox1(9).AddItem !no_akun
                ListBox1(10).AddItem !nm_akun
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        ListBox1(11).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.00")
                        jmlhtgLancar = jmlhtgLancar + Val(Data1.Recordset!saldo_akhir)
                        Data1.Recordset.MoveLast
                        a = True
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    ListBox1(11).AddItem ""
                End If
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 3)) = "2-2" Then
                a = False
                ListBox1(12).AddItem !no_akun
                ListBox1(13).AddItem !nm_akun
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        ListBox1(14).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.00")
                        htgJkPanjang = htgJkPanjang + Val(Data1.Recordset!saldo_akhir)
                        Data1.Recordset.MoveLast
                        a = True
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    ListBox1(14).AddItem ""
                End If
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 2)) = "3-" Then
                a = False
                ListBox1(15).AddItem !no_akun
                ListBox1(16).AddItem !nm_akun
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        ListBox1(17).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.00")
                        modal = modal + Val(Data1.Recordset!saldo_akhir)
                        Data1.Recordset.MoveLast
                        a = True
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    ListBox1(17).AddItem ""
                End If
            End If
            
            If Trim(Mid(.Fields("no_akun"), 1, 1)) = "4" Then
                a = False
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        b = b + Val(Data1.Recordset!saldo_akhir)
                        Data1.Recordset.MoveLast
                    End If
                    Data1.Recordset.MoveNext
                Loop
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 1)) = "5" Then
                a = False
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        c = c + Data1.Recordset!saldo_akhir
                        Data1.Recordset.MoveLast
                    End If
                    Data1.Recordset.MoveNext
                Loop
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 1)) = "6" Then
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        d = d + Data1.Recordset!saldo_akhir
                        Data1.Recordset.MoveLast
                    End If
                    Data1.Recordset.MoveNext
                Loop
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 1)) = "7" Then
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        e = e + Data1.Recordset!saldo_akhir
                        Data1.Recordset.MoveLast
                    End If
                    Data1.Recordset.MoveNext
                Loop
            End If
            .MoveNext
        Loop
        Label2(22).Caption = Format(b - c - d + e, "###,###.00")
        Label2(16).Caption = Format(jmlAktivaLancar, "###,###.00")
        Label2(17).Caption = Format(jmlHrgPerolehan, "###,###.00")
        Label2(18).Caption = Format(jmlPeny, "###,###.00")
        Label2(19).Caption = Format(jmlHrgPerolehan + jmlPeny, "###,###.00")
        Label2(20).Caption = Format(jmlAktivaLancar + jmlHrgPerolehan + jmlPeny, "###,###.00")
        Label2(21).Caption = Format(jmlhtgLancar, "###,###.00")
        Label2(23).Caption = Format(modal + b - c - d + e, "###,###.00")
        Label2(24).Caption = Format(jmlhtgLancar + htgJkPanjang + modal + b - c - d + e, "###,###.00")
    Else
        Label2(22).Caption = 0
        Label2(16).Caption = 0
        Label2(17).Caption = 0
        Label2(18).Caption = 0
        Label2(19).Caption = 0
        Label2(20).Caption = 0
        Label2(21).Caption = 0
        Label2(23).Caption = 0
        Label2(24).Caption = 0
    End If
End With
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data2.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data1.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "'"
Data1.Refresh
Data2.RecordSource = "select * from akuntbl order by no_akun asc"
Data2.Refresh
isi_cmb1
isi_cmb2
isi_data
End Sub

