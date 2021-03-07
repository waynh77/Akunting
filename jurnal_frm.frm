VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Begin VB.Form jurnal_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6735
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3360
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3840
      Top             =   5400
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2535
      Left            =   1560
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
      _Version        =   524288
      _ExtentX        =   6165
      _ExtentY        =   4471
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2008
      Month           =   1
      Day             =   27
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "KELUAR"
      DownPicture     =   "jurnal_frm.frx":0000
      Height          =   735
      Index           =   4
      Left            =   9480
      MouseIcon       =   "jurnal_frm.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "jurnal_frm.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   2760
      TabIndex        =   27
      Text            =   "Text7"
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Text            =   "Text6"
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "TAMBAH"
      DownPicture     =   "jurnal_frm.frx":173E
      Height          =   735
      Index           =   0
      Left            =   120
      MouseIcon       =   "jurnal_frm.frx":2408
      MousePointer    =   99  'Custom
      Picture         =   "jurnal_frm.frx":2712
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "EDIT"
      DownPicture     =   "jurnal_frm.frx":4094
      Height          =   735
      Index           =   1
      Left            =   1440
      MouseIcon       =   "jurnal_frm.frx":4D5E
      MousePointer    =   99  'Custom
      Picture         =   "jurnal_frm.frx":5068
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "HAPUS"
      DownPicture     =   "jurnal_frm.frx":69EA
      Height          =   735
      Index           =   2
      Left            =   2760
      MouseIcon       =   "jurnal_frm.frx":76B4
      MousePointer    =   99  'Custom
      Picture         =   "jurnal_frm.frx":79BE
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CETAK"
      DownPicture     =   "jurnal_frm.frx":9340
      Height          =   735
      Index           =   3
      Left            =   4080
      MouseIcon       =   "jurnal_frm.frx":A00A
      MousePointer    =   99  'Custom
      Picture         =   "jurnal_frm.frx":A314
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5880
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "jurnal_frm.frx":AFDE
      Height          =   5535
      Left            =   5280
      OleObjectBlob   =   "jurnal_frm.frx":AFF2
      TabIndex        =   19
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Data Data1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "DATA JURNAL"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2160
      TabIndex        =   18
      Text            =   "Text5"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Text            =   "Text4"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   4320
      Width           =   2295
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   2640
      TabIndex        =   8
      Text            =   "Combo4"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Text            =   "Combo3"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DEBIT"
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
      Index           =   12
      Left            =   960
      TabIndex        =   32
      Top             =   5520
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   11
      Left            =   120
      TabIndex        =   31
      Top             =   5520
      Width           =   750
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL KREDIT"
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
      Index           =   10
      Left            =   2760
      TabIndex        =   25
      Top             =   4800
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DEBIT"
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
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   1365
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URAIAN"
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
      Index           =   8
      Left            =   1200
      TabIndex        =   17
      Top             =   2880
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUKTI TRANSAKSI"
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
      Index           =   7
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA AKUN"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   6
      Left            =   1560
      TabIndex        =   14
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   2280
      Y2              =   2280
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
      Height          =   225
      Index           =   5
      Left            =   2760
      TabIndex        =   11
      Top             =   4080
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
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO AKUN"
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
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1005
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
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JURNAL"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "jurnal_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sd, sk As Double

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

Sub cmd_awal()
Command1(0).Caption = "TAMBAH"
Command1(1).Caption = "EDIT"
Command1(2).Visible = True
Command1(3).Visible = True
Command1(4).Visible = True
DBGrid1.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
End Sub

Sub cmd_simpan()
Combo1.Enabled = False
Combo2.Enabled = False
Command1(0).Caption = "SIMPAN"
Command1(1).Caption = "BATAL"
Command1(2).Visible = False
Command1(3).Visible = False
Command1(4).Visible = False
DBGrid1.Enabled = False
End Sub


Sub kosong()
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
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

Private Sub Calendar1_Click()
Text1 = Format(Calendar1.Value, "d mmmm yyyy")
Calendar1.Visible = False
End Sub

Private Sub Combo1_Click()
Data1.RecordSource = "select * from jurnal where bulan= '" & Combo1 & "' and tahun = '" & Combo2 & "' order by tgl_jurnal asc"
Data1.Refresh
If Not Data1.Recordset.BOF Then
Data1.Recordset.MoveLast
End If
cek_bb
hit_jml
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Sub tutup()
Text1.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
End Sub

Sub buka()
Text1.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
End Sub

Sub cek_bb()
With bb_frm
.Data1.RecordSource = "select * from buku_besar where bulan= '" & Combo1 & "' and tahun = '" & Combo2 & "'"
.Data1.Refresh
If .Data1.Recordset.BOF Then
    Command1(0).Visible = False
    Command1(1).Visible = False
    Command1(2).Visible = False
    Command1(3).Visible = False
Else
    Command1(0).Visible = True
    Command1(1).Visible = True
    Command1(2).Visible = True
    Command1(3).Visible = True
End If
End With
    bb_frm.Data1.RecordSource = "select * from buku_besar"
    bb_frm.Data1.Refresh

End Sub

Private Sub Combo2_Click()
Data1.RecordSource = "select * from jurnal where bulan= '" & Combo1 & "' and tahun = '" & Combo2 & "'order by tgl_jurnal asc"
Data1.Refresh
If Not Data1.Recordset.BOF Then
Data1.Recordset.MoveLast
End If
cek_bb
hit_jml

End Sub

Private Sub Combo3_Click()
isi_cmb4
isi_akun
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        cmd_simpan
        buka
        kosong
        Text1 = Format(Date, "d mmmm yyyy")
        isi_cmb3
        isi_cmb4
        Label3.Caption = "t"
        sk = 0
        sd = 0
    Else
        simpan
        cmd_awal
        tutup
        If Not Data1.Recordset.BOF Then
        Data1.Recordset.MoveLast
        End If
        hit_jml
        
    End If
Case 1
    If Command1(1).Caption = "EDIT" Then
        cmd_simpan
        isi_cmb4
        Combo4 = Data1.Recordset!nm_bantu
        buka
        Label3.Caption = "e"
        sd = Data1.Recordset!debit
        sk = Data1.Recordset!kredit
    Else
        cmd_awal
        tutup
        If Not Data1.Recordset.BOF Then
            Data1.Recordset.MoveLast
        End If
        isi
        hit_jml
        
    End If
Case 2

Case 3
    Call uncons
Case 4
    bb_frm.Data1.Refresh
    BantuPiutang_frm.Data1.Refresh
    BantuHutang_frm.Data1.Refresh
    Data1.Refresh
    Unload Me
    Akuntansi_frm.Enabled = True
    Akuntansi_frm.Show
End Select
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
isi
End Sub

Private Sub Form_Activate()
Data1.DatabaseName = App.Path + "\dbakuntansi.mdb"
Data1.RecordSource = "select * from jurnal where bulan= '" & Combo1 & "' and tahun = '" & Combo2 & "'order by tgl_jurnal asc "
Data1.Refresh
If Not Data1.Recordset.BOF Then
Data1.Recordset.MoveLast
End If
isi_cmb1
isi_cmb2
isi_cmb3
isi_akun
isi
hit_jml
cek_bb
End Sub


Private Sub Form_Load()
kosong
Text1 = Format(Date, "d mmmm yyyy")
Calendar1 = Date
tutup
cek_bb
End Sub

Sub isi_cmb3()
Combo3.Clear
With AkunTbl_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If !dk <> "-" Then
            Combo3.AddItem !no_akun
        End If
        .MoveNext
    Loop
    Combo3.ListIndex = 0
End If
End With
End Sub

Sub simpan()
If Text5 = "" Then
    Call msgValDat
    Text5.SetFocus
Else
    With Data1.Recordset
    If Label3.Caption = "t" Then
        .AddNew
        trans
        Data1.Refresh
    Else
        .Edit
        trans
    End If
    End With
End If
End Sub

Sub hit_jml()
Dim a, b As Double
a = 0
b = 0
With Data1.Recordset
Data1.Refresh
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
           a = a + !debit
           b = b + !kredit
           .MoveNext
        Loop
    End If
End With
Data1.Refresh
Text6 = Format(a, "###,###.##")
Text7 = Format(b, "###,###.##")
    If a - b = 0 Then
        Label2(12).Caption = "SALDO BALANCE"
        Timer1.Enabled = False
        Timer2.Enabled = False
        Label2(12).Visible = True
    Else
        Label2(12).Caption = "SALDO BELUM BALANCE"
        Timer1.Enabled = True
    End If

End Sub

Sub trans()
Dim a As String
With Data1.Recordset
    !bulan = Combo1
    !tahun = Combo2
    !tgl_jurnal = Text1
    !no_akun = Combo3
    !nm_bantu = Combo4
    !bukti = Text4
    !uraian = Text5
    !debit = Val(Text2)
    !kredit = Val(Text3)
    .Update
End With
With bb_frm.Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !bulan = Combo1 And !tahun = Combo2 And !no_akun = Combo3 Then
                .Edit
                !saldo_debit = !saldo_debit + Val(Text2) - sd
                !saldo_kredit = !saldo_kredit + Val(Text3) - sk
                With AkunTbl_frm.Data1.Recordset
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = Combo3 Then
                            a = !dk
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                End With
                If a = "d" Then
                    !saldo_akhir = !saldo_awal - !saldo_debit + !saldo_kredit
                Else
                    !saldo_akhir = !saldo_awal - !saldo_kredit + !saldo_debit
                End If
                .Update
                .MoveLast
            End If
            .MoveNext
        Loop
    End If
End With
'If Combo3 = "1-130" Or Combo3 = "2-110" Then
    If Combo3 = "1-130" Then
        With BantuPiutang_frm.Data1.Recordset
            If Not .BOF Then
                .MoveFirst
                Do While Not .EOF
                    If !bulan = Combo1 And !tahun = Combo2 And !nm_piutang = Combo4 Then
                        .Edit
                        !saldo_akhir = !saldo_akhir + Val(Text2) - Val(Text3) - sk + sd
                        .Update
                        .MoveLast
                    End If
                    .MoveNext
                Loop
            End If
        End With
    ElseIf Combo3 = "2-110" Then
        With BantuHutang_frm.Data1.Recordset
            If Not .BOF Then
                .MoveFirst
                Do While Not .EOF
                    If !bulan = Combo1 And !tahun = Combo2 And !nm_hutang = Combo4 Then
                        .Edit
                        !saldo_akhir = !saldo_akhir - Val(Text2) + Val(Text3) + sd - sk
                        .Update
                        .MoveLast
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
'End If
    bb_frm.Data1.Refresh
    BantuPiutang_frm.Data1.Refresh
    BantuHutang_frm.Data1.Refresh
    Data1.Refresh
End Sub

Sub isi()
With Data1.Recordset
If Not .BOF Then
    Text1 = Format(!tgl_jurnal, "d mmmm yyyy")
    Combo3 = !no_akun
    If Combo3 = "1-130" Or Combo3 = "2-110" Then
        Combo4.Visible = True
        Combo4 = !nm_bantu
    Else
        Combo4.Visible = False
    End If
    Text4 = !bukti
    Text5 = !uraian
    Text2 = !debit
    Text3 = !kredit
    isi_akun
Else
    kosong
End If
End With
End Sub

Sub isi_cmb4()
Combo4.Clear
If Combo3 = "1-130" Then
    Combo4.Visible = True
    With BantuPiutang_frm.Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo4.AddItem !nm_piutang
            .MoveNext
        Loop
    End If
    Combo4.ListIndex = 0
    End With
ElseIf Combo3 = "2-110" Then
    Combo4.Visible = True
    With BantuHutang_frm.Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo4.AddItem !nm_hutang
            .MoveNext
        Loop
    End If
    Combo4.ListIndex = 0
    End With
ElseIf Combo3 = "2-120" Then
    Combo4.Visible = True
    With BantuHutangLancar_frm.Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo4.AddItem !nm_hutang
            .MoveNext
        Loop
    End If
    Combo4.ListIndex = 0
    End With
ElseIf Combo3 = "2-210" Then
    Combo4.Visible = True
    With BantuHutangPanjang_frm.Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo4.AddItem !nm_hutang
            .MoveNext
        Loop
    End If
    Combo4.ListIndex = 0
    End With
Else
    Combo4.Visible = False
End If
End Sub

Sub isi_akun()
With AkunTbl_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
    If !no_akun = Combo3 Then
        Label2(6).Caption = !nm_akun
        .MoveLast
    End If
    .MoveNext
    Loop
    AkunTbl_frm.Data1.Refresh
End If
End With
End Sub

Private Sub Text1_GotFocus()
Calendar1.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_LostFocus()
Calendar1.Visible = False
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii <= Asc("-") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii <= Asc("-") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
Label2(12).Visible = True
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Label2(12).Visible = False
Timer2.Enabled = False
Timer1.Enabled = True
End Sub
