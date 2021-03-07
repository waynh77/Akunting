VERSION 5.00
Begin VB.Form Period_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2085
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "BATAL"
      DownPicture     =   "Period_frm.frx":0000
      Height          =   855
      Index           =   1
      Left            =   2520
      Picture         =   "Period_frm.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "PROSES"
      DownPicture     =   "Period_frm.frx":264C
      Height          =   855
      Index           =   0
      Left            =   120
      Picture         =   "Period_frm.frx":3316
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT PERIODE"
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
      TabIndex        =   5
      Top             =   120
      Width           =   2010
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
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Period_frm"
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
Combo1.ListIndex = 0
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
Combo2.ListIndex = 0
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Dim a As Boolean
Select Case Index
Case 0
    AkunTbl_frm.data1.Refresh
    bb_frm.data1.Refresh
    With bb_frm.data1.Recordset
    If Not .BOF Then
        .MoveFirst
        a = False
        Do While Not .EOF
            If !bulan = Combo1.ListIndex And !tahun = Combo2 Then
                a = True
                .MoveLast
            End If
            .MoveNext
        Loop
        If a = False Then
            AkunTbl_frm.data1.Recordset.MoveFirst
            Do While Not AkunTbl_frm.data1.Recordset.EOF
            If AkunTbl_frm.data1.Recordset!dk <> "-" Then
                .AddNew
                !bulan = Combo1 '.ListIndex + 1
                !tahun = Combo2
                !no_akun = AkunTbl_frm.data1.Recordset!no_akun
                !saldo_awal = 0
                !saldo_debit = 0
                !saldo_kredit = 0
                !saldo_akhir = 0
                .Update
            End If
                AkunTbl_frm.data1.Recordset.MoveNext
            Loop
            bb_frm.data1.Refresh
            Me.Hide
            Akuntansi_frm.Enabled = True
            Akuntansi_frm.Show
        Else
            MsgBox "Data sudah ada,silahkan masukan yg lain", vbCritical, "Validasi Data"
            Combo1.SetFocus
        End If
    Else
        AkunTbl_frm.data1.Recordset.MoveFirst
        Do While Not AkunTbl_frm.data1.Recordset.EOF
            If AkunTbl_frm.data1.Recordset!dk <> "-" Then
                .AddNew
                !bulan = Combo1 '.ListIndex + 1
                !tahun = Combo2
                !no_akun = AkunTbl_frm.data1.Recordset!no_akun
                !saldo_awal = 0
                !saldo_debit = 0
                !saldo_kredit = 0
                !saldo_akhir = 0
                .Update
            End If
            AkunTbl_frm.data1.Recordset.MoveNext
        Loop
        bb_frm.data1.Refresh
        Me.Hide
        Akuntansi_frm.Enabled = True
        Akuntansi_frm.Show
    End If
    End With
Case 1
    Me.Hide
    Akuntansi_frm.Enabled = True
    Akuntansi_frm.Show
End Select
End Sub

Private Sub Form_Activate()
With bb_frm.data1.Recordset
If Not .BOF Then
    .MoveLast
    Combo1 = !bulan
    Combo2 = !tahun
End If
End With
End Sub

Private Sub Form_Load()
isi_cmb1
isi_cmb2
End Sub
