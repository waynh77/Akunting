VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form NeracaLAjur_frm 
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
      Left            =   7680
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "NEXT PAGE"
      DownPicture     =   "NeracaLAjur_frm.frx":0000
      Height          =   735
      Index           =   2
      Left            =   10200
      MouseIcon       =   "NeracaLAjur_frm.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "NeracaLAjur_frm.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "KELUAR"
      DownPicture     =   "NeracaLAjur_frm.frx":2956
      Height          =   735
      Index           =   1
      Left            =   13560
      MouseIcon       =   "NeracaLAjur_frm.frx":3620
      MousePointer    =   99  'Custom
      Picture         =   "NeracaLAjur_frm.frx":392A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CETAK"
      DownPicture     =   "NeracaLAjur_frm.frx":45F4
      Height          =   735
      Index           =   0
      Left            =   11880
      MouseIcon       =   "NeracaLAjur_frm.frx":52BE
      MousePointer    =   99  'Custom
      Picture         =   "NeracaLAjur_frm.frx":55C8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   480
      Width           =   855
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
   Begin MSForms.ListBox ListBox1 
      Height          =   7575
      Index           =   6
      Left            =   8880
      TabIndex        =   25
      Top             =   1440
      Width           =   1575
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;13361"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7575
      Index           =   9
      Left            =   13560
      TabIndex        =   28
      Top             =   1440
      Width           =   1575
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;13361"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7575
      Index           =   8
      Left            =   12000
      TabIndex        =   27
      Top             =   1440
      Width           =   1575
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;13361"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7575
      Index           =   7
      Left            =   10440
      TabIndex        =   26
      Top             =   1440
      Width           =   1575
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;13361"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7575
      Index           =   4
      Left            =   6720
      TabIndex        =   23
      Top             =   1440
      Width           =   1575
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;13361"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7575
      Index           =   3
      Left            =   5160
      TabIndex        =   22
      Top             =   1440
      Width           =   1575
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;13361"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7575
      Index           =   2
      Left            =   4080
      TabIndex        =   21
      Top             =   1440
      Width           =   2175
      VariousPropertyBits=   746588189
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3836;13361"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7575
      Index           =   1
      Left            =   1200
      TabIndex        =   20
      Top             =   1440
      Width           =   2895
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "5106;13361"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7575
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   1695
      VariousPropertyBits=   746588189
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;13361"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape13 
      Height          =   6735
      Left            =   13560
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape Shape12 
      Height          =   6735
      Left            =   12000
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape Shape11 
      Height          =   6735
      Left            =   10440
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape Shape10 
      Height          =   6735
      Left            =   8880
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape Shape9 
      Height          =   6975
      Left            =   8880
      Top             =   960
      Width           =   3135
   End
   Begin VB.Shape Shape8 
      Height          =   8055
      Left            =   8280
      Top             =   960
      Width           =   615
   End
   Begin VB.Shape Shape7 
      Height          =   6735
      Left            =   6720
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape Shape6 
      Height          =   6735
      Left            =   5160
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape Shape5 
      Height          =   6975
      Left            =   5160
      Top             =   960
      Width           =   3135
   End
   Begin VB.Shape Shape4 
      Height          =   6975
      Left            =   4080
      Top             =   960
      Width           =   1095
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
      Left            =   14040
      TabIndex        =   18
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
      Left            =   12480
      TabIndex        =   17
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
      Left            =   12360
      TabIndex        =   16
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
      Left            =   10920
      TabIndex        =   15
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
      Left            =   9240
      TabIndex        =   14
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
      Left            =   9120
      TabIndex        =   13
      Top             =   960
      Width           =   2535
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
      Left            =   8400
      TabIndex        =   12
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
      Left            =   7200
      TabIndex        =   11
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
      Left            =   5640
      TabIndex        =   10
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
      Left            =   5280
      TabIndex        =   9
      Top             =   960
      Width           =   3000
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
      Left            =   4320
      TabIndex        =   8
      Top             =   960
      Width           =   660
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   3
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
   Begin MSForms.ListBox ListBox1 
      Height          =   7575
      Index           =   5
      Left            =   8280
      TabIndex        =   24
      Top             =   1440
      Width           =   1575
      VariousPropertyBits=   746588189
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;13361"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "NeracaLAjur_frm"
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
    Command1(2).Visible = False
    CommonDialog1.ShowPrinter
    NeracaLAjur_frm.PrintForm
    Command1(0).Visible = True
    Command1(1).Visible = True
    Command1(2).Visible = True
Case 1
    Akuntansi_frm.Enabled = True
    Unload Me
    Akuntansi_frm.Show
Case 2
    Me.Hide
    NeracaLAjur2_frm.Show
    NeracaLAjur2_frm.Text2 = Combo1
    NeracaLAjur2_frm.Text3 = Combo2
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
isi_data
End Sub

Sub isi_data()
Dim a As Boolean
Dim b, c, d, e, f, g, h, I As Double
Dim brs As Single
Data1.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "'"
Data1.Refresh

With Data1.Recordset
kosong
NeracaLAjur2_frm.kosong
If Not .BOF Then
    brs = 1
    b = 0
    c = 0
    d = 0
    e = 0
    f = 0
    g = 0
    h = 0
    I = 0
    Data2.Recordset.MoveFirst
    Do While Not Data2.Recordset.EOF
        If brs < 35 Then
            ListBox1(0).AddItem Data2.Recordset!no_akun
            ListBox1(1).AddItem Data2.Recordset!nm_akun
            ListBox1(2).AddItem Data2.Recordset!dk
            ListBox1(5).AddItem Data2.Recordset!nrlr
            a = False
            .MoveFirst
            Do While Not .EOF
                If !no_akun = Data2.Recordset!no_akun And !bulan = Combo1 And !tahun = Combo2 Then
                    If Data2.Recordset!dk = "D" Then
                        ListBox1(3).AddItem Format(!saldo_akhir, "###,###.##")
                        ListBox1(4).AddItem 0
                        b = b + !saldo_akhir
                    Else
                        ListBox1(3).AddItem 0
                        ListBox1(4).AddItem Format(!saldo_akhir, "###,###.##")
                        c = c + !saldo_akhir
                    End If
                    If Data2.Recordset!nrlr = "NR" Then
                        ListBox1(6).AddItem 0
                        ListBox1(7).AddItem 0
                        If Data2.Recordset!dk = "D" Then
                            ListBox1(8).AddItem Format(!saldo_akhir, "###,###.##")
                            ListBox1(9).AddItem 0
                            f = f + !saldo_akhir
                        Else
                            ListBox1(8).AddItem 0
                            ListBox1(9).AddItem Format(!saldo_akhir, "###,###.##")
                            g = g + !saldo_akhir
                        End If
                    Else
                        ListBox1(8).AddItem 0
                        ListBox1(9).AddItem 0
                        If Data2.Recordset!dk = "D" Then
                            ListBox1(6).AddItem Format(!saldo_akhir, "###,###.##")
                            ListBox1(7).AddItem 0
                            d = d + !saldo_akhir
                        Else
                            ListBox1(6).AddItem 0
                            ListBox1(7).AddItem Format(!saldo_akhir, "###,###.##")
                            e = e + !saldo_akhir
                        End If
                    End If
                    a = True
                    .MoveLast
                End If
                .MoveNext
            Loop
            If a = False Then
                ListBox1(3).AddItem 0
                ListBox1(4).AddItem 0
                ListBox1(6).AddItem 0
                ListBox1(7).AddItem 0
                ListBox1(8).AddItem 0
                ListBox1(9).AddItem 0
            End If
        End If
        If brs >= 35 Then
            Command1(2).Visible = True
            With NeracaLAjur2_frm
                .Text2 = Combo1
                .Text3 = Combo2
                .ListBox1(0).AddItem Data2.Recordset!no_akun
                .ListBox1(1).AddItem Data2.Recordset!nm_akun
                .ListBox1(2).AddItem Data2.Recordset!dk
                .ListBox1(5).AddItem Data2.Recordset!nrlr
                a = False
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If Data1.Recordset!no_akun = Data2.Recordset!no_akun And Data1.Recordset!bulan = Combo1 And Data1.Recordset!tahun = Combo2 Then
                        If Data2.Recordset!dk = "D" Then
                            .ListBox1(3).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.##")
                            .ListBox1(4).AddItem 0
                            b = b + Data1.Recordset!saldo_akhir
                        Else
                            .ListBox1(3).AddItem 0
                            .ListBox1(4).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.##")
                            c = c + Data1.Recordset!saldo_akhir
                        End If
                        If Data2.Recordset!nrlr = "NR" Then
                            .ListBox1(6).AddItem 0
                            .ListBox1(7).AddItem 0
                            If Data2.Recordset!dk = "D" Then
                                .ListBox1(8).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.##")
                                .ListBox1(9).AddItem 0
                                f = f + Data1.Recordset!saldo_akhir
                            Else
                                .ListBox1(8).AddItem 0
                                .ListBox1(9).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.##")
                                g = g + Data1.Recordset!saldo_akhir
                            End If
                        Else
                            .ListBox1(8).AddItem 0
                            .ListBox1(9).AddItem 0
                            If Data2.Recordset!dk = "D" Then
                                .ListBox1(6).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.##")
                                .ListBox1(7).AddItem 0
                                d = d + Data1.Recordset!saldo_akhir
                            Else
                                .ListBox1(6).AddItem 0
                                .ListBox1(7).AddItem Format(Data1.Recordset!saldo_akhir, "###,###.##")
                                f = f + Data1.Recordset!saldo_akhir
                            End If
                        End If
                        a = True
                        Data1.Recordset.MoveLast
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    .ListBox1(3).AddItem 0
                    .ListBox1(4).AddItem 0
                    .ListBox1(6).AddItem 0
                    .ListBox1(7).AddItem 0
                    .ListBox1(8).AddItem 0
                    .ListBox1(9).AddItem 0
                End If
            End With
        End If
        Data2.Recordset.MoveNext
        brs = brs + 1
    Loop
    Data1.Refresh
    Data2.Refresh
    NeracaLAjur2_frm.Text1(0) = Format(b, "###,###.##")
    NeracaLAjur2_frm.Text1(1) = Format(c, "###,###.##")
    NeracaLAjur2_frm.Text1(2) = Format(d, "###,###.##")
    NeracaLAjur2_frm.Text1(3) = Format(e, "###,###.##")
    NeracaLAjur2_frm.Text1(4) = Format(f, "###,###.##")
    NeracaLAjur2_frm.Text1(5) = Format(g, "###,###.##")
    If d > e Then
        NeracaLAjur2_frm.Text1(6) = Format(d - e, "###,###.##")
        NeracaLAjur2_frm.Text1(7) = 0
    Else
        NeracaLAjur2_frm.Text1(7) = Format(e - d, "###,###.##")
        NeracaLAjur2_frm.Text1(6) = 0
    End If
    If f < g Then
        NeracaLAjur2_frm.Text1(8) = Format(g - h, "###,###.##")
        NeracaLAjur2_frm.Text1(9) = 0
        h = f - g
        I = 0
    Else
        NeracaLAjur2_frm.Text1(9) = Format(f - g, "###,###.##")
        NeracaLAjur2_frm.Text1(8) = 0
        I = f - g
        h = 0
    End If
    NeracaLAjur2_frm.Text1(10) = Format(f + h, "###,###.##")
    NeracaLAjur2_frm.Text1(11) = Format(g + I, "###,###.##")
Else
    kosong
End If
End With
End Sub

Sub kosong()
Dim a As Byte
a = 0
Do Until a = 10
    ListBox1(a).Clear
    a = a + 1
Loop
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data2.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data1.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "'"
Data2.RecordSource = "akuntbl"
Data1.Refresh
Data2.Refresh
isi_cmb1
isi_cmb2
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
