VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form bb_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7620
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CETAK"
      DownPicture     =   "bb_frm.frx":0000
      Height          =   735
      Index           =   1
      Left            =   7440
      MouseIcon       =   "bb_frm.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "bb_frm.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh Data"
      DownPicture     =   "bb_frm.frx":1C9E
      Height          =   735
      Left            =   5760
      MouseIcon       =   "bb_frm.frx":2968
      MousePointer    =   99  'Custom
      Picture         =   "bb_frm.frx":2C72
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   1455
   End
   Begin VB.Data data1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "DBBUKUBESAR"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   7080
      TabIndex        =   19
      Text            =   "Text3"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   18
      Text            =   "Text3"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   17
      Text            =   "Text3"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      Text            =   "xxx"
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Text            =   "Combo3"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "EDIT"
      DownPicture     =   "bb_frm.frx":45F4
      Height          =   735
      Index           =   0
      Left            =   5760
      MouseIcon       =   "bb_frm.frx":52BE
      MousePointer    =   99  'Custom
      Picture         =   "bb_frm.frx":55C8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "KELUAR"
      DownPicture     =   "bb_frm.frx":6F4A
      Height          =   735
      Index           =   2
      Left            =   7440
      MouseIcon       =   "bb_frm.frx":7C14
      MousePointer    =   99  'Custom
      Picture         =   "bb_frm.frx":7F1E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "bb_frm.frx":8688
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "bb_frm.frx":869C
      TabIndex        =   20
      Top             =   3600
      Width           =   8895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit"
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
      Left            =   2520
      TabIndex        =   14
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kredit"
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
      Left            =   4800
      TabIndex        =   13
      Top             =   2400
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Index           =   9
      Left            =   7080
      TabIndex        =   12
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   8880
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Akun DK"
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
      Left            =   3000
      TabIndex        =   11
      Top             =   1320
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUKU BESAR"
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
      TabIndex        =   9
      Top             =   120
      Width           =   1845
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
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   8880
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Akun"
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
      TabIndex        =   7
      Top             =   1320
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Akun"
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
      TabIndex        =   6
      Top             =   1680
      Width           =   1125
   End
End
Attribute VB_Name = "bb_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Data1.RecordSource = "select no_akun,saldo_awal,saldo_debit,saldo_kredit,saldo_akhir,tahun,bulan from buku_besar where bulan= '" & Combo1 & "' and tahun = '" & Combo2 & "'"
Data1.Refresh
cek_data
End Sub

Private Sub Combo2_Click()
Data1.RecordSource = "select no_akun,saldo_awal,saldo_debit,saldo_kredit,saldo_akhir,tahun,bulan from buku_besar where bulan= '" & Combo1 & "' and tahun = '" & Combo2 & "'"
Data1.Refresh
cek_data
End Sub

Private Sub Combo3_Click()
'isi_akun
pindah_grid
End Sub

Sub pindah_grid()
If Combo3 <> "" Then
    If Not Data1.Recordset.BOF Then
        Data1.Recordset.MoveFirst
        Do Until Data1.Recordset!no_akun = Combo3
            Data1.Recordset.MoveNext
        Loop
    End If
End If
End Sub

Sub isi_saldo()
With Data1.Recordset
If Not .BOF Then
    Text3(0) = !saldo_awal
    Text3(1) = !saldo_debit
    Text3(2) = !saldo_kredit
    Text3(3) = !saldo_akhir
End If
End With
frmt_saldo
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "EDIT" Then
        If Combo3 <> "1-130" And Combo3 <> "2-110" Then
            cmd_simpan
            DBGrid1.Enabled = False
            frmt_awal
            Text3(0).SetFocus
        Else
            MsgBox "Transaksi Hutang dan Piutang dilakukan di Buku Bantu", vbInformation, "Validasi Input"
        End If
    Else
        simpan
        isi_saldo
        frmt_saldo
    End If
Case 1
    If Command1(1).Caption = "BATAL" Then
        cmd_awal
        DBGrid1.Enabled = True
        frmt_saldo
    Else
        MsgBox "Silahkan anda mencetak Neraca Lajur untuk mengetahui Saldo Terakhir", vbInformation, "Maaf..."
        'Call uncons
    End If
Case 2
    Unload Me
    Akuntansi_frm.Enabled = True
    Akuntansi_frm.Show
End Select
End Sub

Sub frmt_saldo()
Text3(0) = Format(Text3(0), "###,###.##")
Text3(1) = Format(Text3(1), "###,###.##")
Text3(2) = Format(Text3(2), "###,###.##")
Text3(3) = Format(Text3(3), "###,###.##")
End Sub

Sub frmt_awal()
Text3(0) = Format(Text3(0), "#")
Text3(1) = Format(Text3(1), "#")
Text3(2) = Format(Text3(2), "#")
Text3(3) = Format(Text3(3), "#")
End Sub

Sub simpan()

With Data1.Recordset
    .Edit
    !saldo_awal = Val(Text3(0))
    If Text2 = "D" Then
        !saldo_akhir = !saldo_awal + !saldo_debit - !saldo_kredit
    Else
        !saldo_akhir = !saldo_awal - !saldo_debit + !saldo_kredit
    End If
    .Update
End With
'Data1.Refresh
cmd_awal
'Text3(3) = !saldo_akhir
DBGrid1.Enabled = True
End Sub

Sub cmd_awal()
Command1(0).Caption = "EDIT"
Command1(1).Caption = "CETAK"
Command1(2).Enabled = True
Text3(0).Enabled = False
End Sub

Sub cmd_simpan()
Command1(0).Caption = "SIMPAN"
Command1(1).Caption = "BATAL"
Command1(2).Enabled = False
Text3(0).Enabled = True
End Sub

Private Sub Command2_Click()
Dim a, b As Double
Dim c As String
Data2.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data2.RecordSource = "select* from jurnal where bulan='" & Combo1 & "'and tahun='" & Combo2 & "'"
Data2.Refresh
With Data2.Recordset
If Not .BOF Then
    Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
        a = 0
        b = 0
        .MoveFirst
        Do While Not .EOF
            If !no_akun = Data1.Recordset!no_akun Then
                a = a + !debit
                b = b + !kredit
            End If
        .MoveNext
        Loop
        Data1.Recordset.Edit
        Data1.Recordset!saldo_debit = a
        Data1.Recordset!saldo_kredit = b
        c = "D"
        With AkunTbl_frm.Data1.Recordset
            .MoveFirst
            Do While Not .EOF
                If !no_akun = Data1.Recordset!no_akun Then
                    c = !dk
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
        If c = "D" Then
            Data1.Recordset!saldo_akhir = Data1.Recordset!saldo_awal + Data1.Recordset!saldo_debit - Data1.Recordset!saldo_kredit
        ElseIf c = "K" Then
            Data1.Recordset!saldo_akhir = Data1.Recordset!saldo_awal - Data1.Recordset!saldo_debit + Data1.Recordset!saldo_kredit
        End If
        Data1.Recordset.Update
        Data1.Recordset.MoveNext
    Loop
End If
Data1.Refresh
End With
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Data1.Recordset.BOF Then
    kosong
Else
    isi_saldo
    isi_grid
    isi_akun
End If
End Sub

Private Sub Form_Activate()
Data1.RecordSource = "select no_akun,saldo_awal,saldo_debit,saldo_kredit,saldo_akhir,tahun,bulan from buku_besar where bulan= '" & Combo1 & "' and tahun = '" & Combo2 & "'"
Data1.Refresh
cek_data
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path + "\dbakuntansi.mdb"
Data1.RecordSource = "select no_akun,saldo_awal,saldo_debit,saldo_kredit,saldo_akhir,tahun,bulan from buku_besar where bulan= '" & Combo1 & "' and tahun = '" & Combo2 & "'"
Data1.Refresh
isi_cmb1
isi_cmb2
isi_cmb3
isi_akun
cek_data
End Sub

Sub cek_data()
If Data1.Recordset.BOF Then
    Command1(0).Visible = False
    Command1(1).Visible = False
Else
    Command1(0).Visible = True
    Command1(1).Visible = True
End If
End Sub

Sub isi_akun()
With AkunTbl_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
    If !no_akun = Combo3 Then
        Text1 = !nm_akun
        Text2 = !dk
        .MoveLast
    End If
    .MoveNext
    Loop
    AkunTbl_frm.Data1.Refresh
End If
End With
End Sub

Sub isi_grid()
If Not Data1.Recordset.BOF Then
Combo3 = Data1.Recordset!no_akun
End If
End Sub

Sub kosong()
Combo3 = ""
Text1 = ""
Text2 = ""
Text3(0) = ""
Text3(1) = ""
Text3(2) = ""
Text3(3) = ""
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

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii <= Asc("-") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Command1_Click (0)
    End If
End Sub
