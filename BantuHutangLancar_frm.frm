VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form BantuHutangLancar_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7980
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "BantuHutangLancar_frm.frx":0000
      Height          =   3735
      Left            =   240
      OleObjectBlob   =   "BantuHutangLancar_frm.frx":0014
      TabIndex        =   25
      Top             =   4080
      Width           =   9375
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh Data"
      DownPicture     =   "BantuHutangLancar_frm.frx":09E7
      Height          =   735
      Left            =   3120
      MouseIcon       =   "BantuHutangLancar_frm.frx":16B1
      MousePointer    =   99  'Custom
      Picture         =   "BantuHutangLancar_frm.frx":19BB
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   3600
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Data Data1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "TAMBAH"
      DownPicture     =   "BantuHutangLancar_frm.frx":333D
      Height          =   735
      Index           =   0
      Left            =   120
      MouseIcon       =   "BantuHutangLancar_frm.frx":4007
      MousePointer    =   99  'Custom
      Picture         =   "BantuHutangLancar_frm.frx":4311
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "EDIT"
      DownPicture     =   "BantuHutangLancar_frm.frx":5C93
      Height          =   735
      Index           =   1
      Left            =   1200
      MouseIcon       =   "BantuHutangLancar_frm.frx":695D
      MousePointer    =   99  'Custom
      Picture         =   "BantuHutangLancar_frm.frx":6C67
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "HAPUS"
      DownPicture     =   "BantuHutangLancar_frm.frx":85E9
      Height          =   735
      Index           =   2
      Left            =   2280
      MouseIcon       =   "BantuHutangLancar_frm.frx":92B3
      MousePointer    =   99  'Custom
      Picture         =   "BantuHutangLancar_frm.frx":95BD
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "KELUAR"
      DownPicture     =   "BantuHutangLancar_frm.frx":AF3F
      Height          =   735
      Index           =   3
      Left            =   3360
      MouseIcon       =   "BantuHutangLancar_frm.frx":BC09
      MousePointer    =   99  'Custom
      Picture         =   "BantuHutangLancar_frm.frx":BF13
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "BantuHutangLancar_frm.frx":C67D
      Height          =   2895
      Left            =   4560
      OleObjectBlob   =   "BantuHutangLancar_frm.frx":C691
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lancar Lainnya"
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
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   240
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buku Bantu Hutang"
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
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   2745
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
      TabIndex        =   22
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Akun"
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
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2-120"
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
      Index           =   2
      Left            =   960
      TabIndex        =   20
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Hutang "
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
      TabIndex        =   19
      Top             =   1560
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Awal"
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
      TabIndex        =   18
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Akhir"
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
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1110
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Index           =   6
      Left            =   4680
      TabIndex        =   15
      Top             =   3120
      Width           =   510
   End
   Begin VB.Label awal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nominal"
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
      Left            =   8670
      TabIndex        =   14
      Top             =   3120
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Akhir"
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
      Left            =   5400
      TabIndex        =   13
      Top             =   3360
      Width           =   1110
   End
   Begin VB.Label akhir 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nominal"
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
      Left            =   8670
      TabIndex        =   12
      Top             =   3360
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Awal"
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
      Index           =   7
      Left            =   5400
      TabIndex        =   11
      Top             =   3120
      Width           =   1065
   End
   Begin VB.Line Line2 
      X1              =   5400
      X2              =   9720
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nominal"
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
      Left            =   8670
      TabIndex        =   10
      Top             =   3600
      Width           =   810
   End
End
Attribute VB_Name = "BantuHutangLancar_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sid As Double

Private Sub Combo1_Click()
Data1.RecordSource = "select nm_Hutang,saldo_awal,saldo_akhir,no_akun,bulan,tahun from bantuHutangLANCAR where bulan = '" & Combo1 & "'"
Data1.Refresh
hit_jml
End Sub

Private Sub Combo2_Click()
Data1.RecordSource = "select nm_Hutang,saldo_awal,saldo_akhir,no_akun,bulan,tahun from bantuHutangLANCAR where tahun = '" & Combo2 & "'"
Data1.Refresh
hit_jml
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        kosong
        buka
        cmd_simpan
        Text1.SetFocus
        Label3.Caption = "t"
        sid = 0
    Else
        simpan_data
    End If
Case 1
    If Command1(1).Caption = "EDIT" Then
        buka
        cmd_simpan
        Text1.SetFocus
        Label3.Caption = "e"
        Text2 = Format(Text2, "###")
        sid = Data1.Recordset!saldo_awal
    Else
        tutup
        cmd_awal
        Data1.Refresh
    End If
Case 2
    If Not Data1.Recordset.BOF Then
    x = MsgBox("Apakah anda yakin?", vbOKCancel, "Hapus Data")
        If x = vbOK Then
            Data1.Recordset.Delete
            Data1.Refresh
            hit_jml
        End If
    End If
Case 3
Unload Me
Akuntansi_frm.Enabled = True
Akuntansi_frm.Show
End Select
End Sub

Sub simpan_data()
Dim a As Boolean
If Text1 = "" Or Text2 = "" Then
    Call msgValDat
    If Text1 = "" Then
        Text1.SetFocus
    ElseIf Text2 = "" Then
        Text2.SetFocus
    End If
Else
'    Text3 = Val(Text2)
    With Data1.Recordset
    If Label3.Caption = "t" Then
        If Not .BOF Then
            .MoveFirst
            a = False
            Do While Not .EOF
                If !nm_hutang = Text1 Then
                    .MoveLast
                    a = True
                End If
                .MoveNext
            Loop
        End If
        If a = True Then
            MsgBox "Data sudah ada... silahkan masukan data lainnya", vbCritical, "Validasi Data"
            Text1.SetFocus
        Else
            .AddNew
            !saldo_akhir = Val(Text2)
            simpan
            .Update
            hit_jml
        End If
    Else
        .Edit
        simpan
        .Update
        hit_jml
    End If
    End With
    Data1.Refresh
    tutup
    cmd_awal
    hit_jml
End If
End Sub

Sub cmd_awal()
Command1(0).Caption = "TAMBAH"
Command1(1).Caption = "EDIT"
Command1(2).Enabled = True
Command1(3).Enabled = True
End Sub

Sub cmd_simpan()
Command1(0).Caption = "SIMPAN"
Command1(1).Caption = "BATAL"
Command1(2).Enabled = False
Command1(3).Enabled = False
End Sub

Sub simpan()
With Data1.Recordset
!no_akun = Label2(2).Caption
!nm_hutang = Text1
!bulan = Combo1
!tahun = Combo2
!saldo_awal = Val(Text2)
End With
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

Private Sub Command2_Click()
Dim a, b As Double
Data2.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data2.RecordSource = "select* from jurnal where bulan='" & Combo1 & "'and tahun='" & Combo2 & "'and no_akun='" & Label2(2).Caption & "'"
Data2.Refresh
If Not Data2.Recordset.BOF Then
    With Data1.Recordset
        .MoveFirst
        Do While Not .EOF
            a = 0
            b = 0
            Data2.Recordset.MoveFirst
            Do While Not Data2.Recordset.EOF
                If Data2.Recordset!nm_bantu = !nm_hutang Then
                    a = a + Data2.Recordset!debit
                    b = b + Data2.Recordset!kredit
                End If
                Data2.Recordset.MoveNext
            Loop
            .Edit
            !saldo_akhir = !saldo_awal - a + b
            .Update
            .MoveNext
        Loop
    End With
    hit_jml
    Data1.Refresh
    Data2.Refresh
End If
End Sub


Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Command1(0).Caption <> "SIMPAN" Then
    isi
End If
End Sub

Private Sub Form_Activate()
Data1.RecordSource = "select nm_Hutang,saldo_awal,saldo_akhir,no_akun,bulan,tahun from bantuHutangLANCAR where bulan = '" & Combo1 & "'"
Data1.Refresh
isi
hit_jml
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data1.RecordSource = "bantuHutangLancar"
kosong
isi_cmb1
isi_cmb2
Data1.Caption = "Data Buku Bantu Hutang"
tutup
End Sub

Sub isi()
With Data1.Recordset
If Not .BOF Then
    Text1 = !nm_hutang
    Text2 = Format(!saldo_awal, "###,###,###.##")
    Text3 = Format(!saldo_akhir, "###,###,###.##")
Else
    kosong
End If
End With
End Sub

Sub hit_jml()
Dim a, b As Double
Dim c As Double
With Data1.Recordset
a = 0
b = 0
c = 0
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        a = a + !saldo_awal
        b = b + !saldo_akhir
        .MoveNext
    Loop
End If
c = b - a
awal.Caption = Format(a, "###,###.##")
akhir.Caption = Format(b, "###,###.##")
Label4.Caption = Format(c, "###,###.##")
End With

With bb_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If !no_akun = Label2(2).Caption And !bulan = Combo1 And !tahun = Combo2 Then
            .Edit
            !saldo_awal = a
            !saldo_akhir = b
            .Update
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With

Data1.Refresh
End Sub

Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
End Sub

Sub tutup()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
End Sub

Sub buka()
Text1.Enabled = True
Text2.Enabled = True
'Text3.Enabled = True
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

