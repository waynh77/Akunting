VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form AkunTbl_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4980
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10965
   ControlBox      =   0   'False
   Icon            =   "AkunTbl_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CETAK"
      DownPicture     =   "AkunTbl_frm.frx":1CCA
      Height          =   735
      Index           =   4
      Left            =   3120
      MouseIcon       =   "AkunTbl_frm.frx":2994
      MousePointer    =   99  'Custom
      Picture         =   "AkunTbl_frm.frx":2C9E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "KELUAR"
      DownPicture     =   "AkunTbl_frm.frx":3968
      Height          =   735
      Index           =   3
      Left            =   1680
      MouseIcon       =   "AkunTbl_frm.frx":4632
      MousePointer    =   99  'Custom
      Picture         =   "AkunTbl_frm.frx":493C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "AkunTbl_frm.frx":5606
      Height          =   4695
      Left            =   4680
      OleObjectBlob   =   "AkunTbl_frm.frx":561A
      TabIndex        =   12
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "HAPUS"
      DownPicture     =   "AkunTbl_frm.frx":5FF5
      Height          =   735
      Index           =   2
      Left            =   240
      MouseIcon       =   "AkunTbl_frm.frx":6CBF
      MousePointer    =   99  'Custom
      Picture         =   "AkunTbl_frm.frx":6FC9
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "EDIT"
      DownPicture     =   "AkunTbl_frm.frx":894B
      Height          =   735
      Index           =   1
      Left            =   1680
      MouseIcon       =   "AkunTbl_frm.frx":9615
      MousePointer    =   99  'Custom
      Picture         =   "AkunTbl_frm.frx":991F
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "TAMBAH"
      DownPicture     =   "AkunTbl_frm.frx":B2A1
      Height          =   735
      Index           =   0
      Left            =   240
      MouseIcon       =   "AkunTbl_frm.frx":BF6B
      MousePointer    =   99  'Custom
      Picture         =   "AkunTbl_frm.frx":C275
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Data data1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "DATA TABEL AKUNTANSI"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Width           =   4215
   End
   Begin VB.ComboBox combo2 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox combo1 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox text2 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox text1 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DAFTAR AKUN"
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
      TabIndex        =   9
      Top             =   120
      Width           =   4440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Akun NR/LR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Akun D/K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Akun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Akun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   945
   End
End
Attribute VB_Name = "AkunTbl_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub kosong()
Text1 = ""
Text2 = ""
Combo1 = ""
Combo2 = ""
End Sub

Sub isi()
With Data1.Recordset
If Not .BOF Then
    Text1 = !no_akun
    Text2 = !nm_akun
    Combo1 = !dk
    Combo2 = !nrlr
End If
End With
End Sub

Sub tutup()
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
End Sub

Sub buka()
Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
End Sub

Sub cmd_awal()
Command1(0).Caption = "TAMBAH"
Command1(1).Caption = "EDIT"
Command1(2).Enabled = True
End Sub

Sub cmd_proses()
Command1(0).Caption = "SIMPAN"
Command1(1).Caption = "BATAL"
Command1(2).Enabled = False
End Sub

Sub limiter()
Text1.MaxLength = 6
Text2.MaxLength = 50
End Sub

Sub simpan()
Dim a As Boolean
If Text1 = "" Or Text2 = "" Or Combo1 = "" Or Combo2 = "" Then
    pesan_validasi
    If Text1 = "" Then
        Text1.SetFocus
    ElseIf Text2 = "" Then
        Text2.SetFocus
    ElseIf Combo1 = "" Then
        Combo1.SetFocus
    ElseIf Combo2 = "" Then
        Combo2.SetFocus
    End If
Else
    With Data1.Recordset
    If Label2.Caption = "t" Then
        a = False
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                If !no_akun = Text1 Then
                    a = True
                    .MoveLast
                End If
                .MoveNext
            Loop
        End If
        If a = False Then
            .AddNew
            !no_akun = Text1
            !nm_akun = Text2
            !dk = Combo1
            !nrlr = Combo2
            .Update
            Data1.Refresh
            cmd_awal
            tutup
            isi
        Else
            MsgBox "Data Sudah Ada, Silahkan Masukan Data yg lainnya...", vbOKOnly, "Validasi Data"
            Text1.SetFocus
        End If
    Else
        .Edit
        !no_akun = Text1
        !nm_akun = Text2
        !dk = Combo1
        !nrlr = Combo2
        .Update
        Data1.Refresh
        cmd_awal
        tutup
        isi
    End If
    End With
End If
End Sub

Sub isi_cmb1()
Combo1.Clear
Combo1.AddItem ("D")
Combo1.AddItem ("K")
Combo1.AddItem ("-")
Combo1.ListIndex = 0
End Sub

Sub isi_cmb2()
Combo2.Clear
Combo2.AddItem ("NR")
Combo2.AddItem ("LR")
Combo2.AddItem ("-")
Combo2.ListIndex = 0
End Sub

Sub pesan_validasi()
MsgBox "Data belum lengkap", vbOKOnly, "Validasi Data"
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        Label2.Caption = "t"
        buka
        kosong
        cmd_proses
        Text1.SetFocus
    Else
        simpan
    End If
Case 1
    If Command1(1).Caption = "EDIT" Then
        Label2.Caption = "e"
        buka
        Text1.Enabled = False
        cmd_proses
        Text2.SetFocus
    Else
        cmd_awal
        kosong
        tutup
        isi
    End If
Case 2
    If Not Data1.Recordset.BOF Then
    x = MsgBox("Apakah anda yakin?", vbOKCancel, "Hapus Data")
    If x = vbOK Then
        Data1.Recordset.Delete
    End If
    End If
Case 3
    Me.Hide
    Akuntansi_frm.Enabled = True
    Akuntansi_frm.Show
Case 4
    Me.Enabled = False
    cetakAkun_frm.Show
End Select
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Text1.Enabled = False Then
    isi
End If
End Sub

Private Sub Form_Activate()
limiter
kosong
cmd_awal
tutup
isi_cmb1
isi_cmb2
isi
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data1.RecordSource = "select * from akuntbl order by no_akun asc"
Label2 = "t"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Akuntansi_frm.Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc("-") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
