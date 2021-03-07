VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form LabaRugi_frm 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9180
   ClientLeft      =   4020
   ClientTop       =   1755
   ClientWidth     =   7890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   7890
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CETAK"
      DownPicture     =   "LabaRugi_frm.frx":0000
      Height          =   855
      Index           =   0
      Left            =   4200
      MouseIcon       =   "LabaRugi_frm.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "LabaRugi_frm.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "KELUAR"
      DownPicture     =   "LabaRugi_frm.frx":1C9E
      Height          =   855
      Index           =   1
      Left            =   6120
      MouseIcon       =   "LabaRugi_frm.frx":2968
      MousePointer    =   99  'Custom
      Picture         =   "LabaRugi_frm.frx":2C72
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PENJUALAN BERSIH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   13
      Left            =   5760
      TabIndex        =   21
      Top             =   8640
      Width           =   1830
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PENJUALAN BERSIH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   12
      Left            =   5760
      TabIndex        =   20
      Top             =   8160
      Width           =   1830
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PENJUALAN BERSIH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   11
      Left            =   5760
      TabIndex        =   19
      Top             =   7800
      Width           =   1830
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PENJUALAN BERSIH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Index           =   10
      Left            =   5760
      TabIndex        =   18
      Top             =   7320
      Width           =   1830
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PENJUALAN BERSIH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Index           =   9
      Left            =   5760
      TabIndex        =   17
      Top             =   6960
      Width           =   1830
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7680
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7680
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABA BERSIH SEBELUM PAJAK"
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
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   8640
      Width           =   3000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH PENDAPATAN DAN BIAYA LAIN-LAIN"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   8160
      Width           =   4470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABA BERSIH USAHA"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   7800
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH BIAYA"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   7320
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PENJUALAN BERSIH"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   1950
   End
   Begin MSForms.ListBox ListBox3 
      Height          =   5175
      Left            =   5280
      TabIndex        =   11
      Top             =   1560
      Width           =   2415
      VariousPropertyBits=   746588185
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "4260;9128"
      MatchEntry      =   0
      BorderColor     =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox2 
      Height          =   5175
      Left            =   1200
      TabIndex        =   10
      Top             =   1560
      Width           =   4095
      VariousPropertyBits=   746588185
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7223;9128"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   5175
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
      VariousPropertyBits=   746588185
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;9128"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   5280
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   1200
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   120
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PERHITUNGAN"
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
      Top             =   1200
      Width           =   2160
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
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   1080
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
      Left            =   120
      TabIndex        =   4
      Top             =   1200
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
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LAPORAN LABA-RUGI"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2670
   End
End
Attribute VB_Name = "LabaRugi_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub Combo1_Click()
Data1.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "'"
Data1.Refresh
isi_data
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
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
    LabaRugi_frm.PrintForm
    Command1(0).Visible = True
    Command1(1).Visible = True
Case 1
    Akuntansi_frm.Enabled = True
    Unload Me
    Akuntansi_frm.Show
End Select
End Sub

Private Sub Form_Activate()
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

Sub isi_data()
Dim a As Boolean
Dim b, c As Double
Dim d, e As Double
ListBox1.Clear
ListBox2.Clear
ListBox3.Clear
Label2(9).Caption = ""
Label2(10).Caption = ""
Label2(11).Caption = ""
Label2(12).Caption = ""
Label2(13).Caption = ""
b = 0
c = 0
d = 0
e = 0
With Data2.Recordset
    If Not .BOF And Not Data1.Recordset.BOF Then
        .MoveFirst
        Do While Not .EOF
            If Trim(Mid(.Fields("no_akun"), 1, 1)) = "4" Then
                a = False
                ListBox1.AddItem (!no_akun)
                ListBox2.AddItem (!nm_akun)
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        ListBox3.AddItem Format(Data1.Recordset!saldo_akhir, "###,###.##")
                        b = b + Val(Data1.Recordset!saldo_akhir)
                        a = True
                        Data1.Recordset.MoveLast
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    ListBox3.AddItem ""
                End If
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 1)) = "5" Then
                a = False
                ListBox1.AddItem (!no_akun)
                ListBox2.AddItem (!nm_akun)
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        ListBox3.AddItem Format(Data1.Recordset!saldo_akhir, "###,###.##")
                        a = True
                        c = c + Data1.Recordset!saldo_akhir
                        Data1.Recordset.MoveLast
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    ListBox3.AddItem ""
                End If
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 1)) = "6" Then
                a = False
                ListBox1.AddItem (!no_akun)
                ListBox2.AddItem (!nm_akun)
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        ListBox3.AddItem Format(Data1.Recordset!saldo_akhir, "###,###.##")
                        a = True
                        d = d + Data1.Recordset!saldo_akhir
                        Data1.Recordset.MoveLast
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    ListBox3.AddItem ""
                End If
            End If
            If Trim(Mid(.Fields("no_akun"), 1, 1)) = "7" Then
                a = False
                ListBox1.AddItem (!no_akun)
                ListBox2.AddItem (!nm_akun)
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                    If !no_akun = Data1.Recordset!no_akun Then
                        ListBox3.AddItem Format(Data1.Recordset!saldo_akhir, "###,###.##")
                        a = True
                        e = e + Data1.Recordset!saldo_akhir
                        Data1.Recordset.MoveLast
                    End If
                    Data1.Recordset.MoveNext
                Loop
                If a = False Then
                    ListBox3.AddItem ""
                End If
            End If
            .MoveNext
        Loop
        Label2(9).Caption = Format(b - c, "###,###.00")
        Label2(10).Caption = Format(d, "###,###.00")
        Label2(11).Caption = Format(b - c - d, "###,###.00")
        Label2(12).Caption = Format(e, "###,###.00")
        Label2(13).Caption = Format(b - c - d + e, "###,###.00")
    End If
End With
End Sub
