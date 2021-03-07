VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form LabaDitahan_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3525
   ClientLeft      =   4650
   ClientTop       =   3165
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6105
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CETAK"
      DownPicture     =   "LabaDitahan_frm.frx":0000
      Height          =   735
      Index           =   0
      Left            =   2520
      MouseIcon       =   "LabaDitahan_frm.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "LabaDitahan_frm.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "KELUAR"
      DownPicture     =   "LabaDitahan_frm.frx":1C9E
      Height          =   735
      Index           =   1
      Left            =   4200
      MouseIcon       =   "LabaDitahan_frm.frx":2968
      MousePointer    =   99  'Custom
      Picture         =   "LabaDitahan_frm.frx":2C72
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Text            =   "xxx"
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
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABA DITAHAN AWAL"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   2280
      Width           =   2130
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABA DITAHAN AWAL"
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
      Left            =   3720
      TabIndex        =   10
      Top             =   1800
      Width           =   2130
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABA DITAHAN AWAL"
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
      Left            =   3720
      TabIndex        =   9
      Top             =   1560
      Width           =   2130
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5880
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH LABA DITAHAN"
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
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   2370
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABA/RUGI PEIODE BERJALAN"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1800
      Width           =   3045
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABA DITAHAN AWAL"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   2130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3-200"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   480
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
      ForeColor       =   &H00008000&
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   5880
      Y1              =   1080
      Y2              =   1080
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
      Caption         =   "LAPORAN LABA DITAHAN"
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
      Width           =   3735
   End
End
Attribute VB_Name = "LabaDitahan_frm"
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
Data3.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "' and no_akun='3-200'"
Data3.Refresh
isi_data
End Sub

Sub isi_data()
Dim a As Double
Dim b, c As Double
Dim d, e As Double
a = 0
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
        Label2(7).Caption = Format(b - c - d + e, "###,###.00")
        With Data3.Recordset
            If Not .BOF Then
                a = !saldo_awal
                Label2(6).Caption = Format(a, "###,###.00")
            End If
        End With
        Label2(8).Caption = Format(a + b - c - d + e, "###,###.00")
    End If
End With
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub Combo2_Click()
Data1.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "'"
Data1.Refresh
Data3.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "' and no_akun='3-200'"
Data3.Refresh
isi_data
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Command1(0).Visible = False
    Command1(1).Visible = False
    CommonDialog1.ShowPrinter
    LabaDitahan_frm.PrintForm
    Command1(0).Visible = True
    Command1(1).Visible = True
Case 1
    Akuntansi_frm.Enabled = True
    Unload Me
    Akuntansi_frm.Show
End Select
End Sub

Private Sub Form_Activate()
Data1.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "'"
Data1.Refresh
Data2.RecordSource = "select * from akuntbl order by no_akun asc"
Data2.Refresh
Data3.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "' and no_akun='3-200'"
Data3.Refresh
isi_cmb1
isi_cmb2
isi_data
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data2.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data3.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data1.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "'"
Data1.Refresh
Data2.RecordSource = "select * from akuntbl order by no_akun asc"
Data2.Refresh
Data3.RecordSource = "select * from buku_besar where bulan = '" & Combo1 & " ' and tahun= '" & Combo2 & "' and no_akun='3-200'"
Data3.Refresh
isi_cmb1
isi_cmb2
isi_data
End Sub
