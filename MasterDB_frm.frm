VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form MasterDB_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7860
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "MasterDB_frm.frx":0000
      Height          =   5895
      Left            =   120
      OleObjectBlob   =   "MasterDB_frm.frx":0014
      TabIndex        =   6
      Top             =   1800
      Width           =   9135
   End
   Begin VB.Data Data1 
      BackColor       =   &H0080FF80&
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
      Top             =   1320
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "KELUAR"
      DownPicture     =   "MasterDB_frm.frx":09FF
      Height          =   855
      Index           =   2
      Left            =   7680
      MouseIcon       =   "MasterDB_frm.frx":16C9
      MousePointer    =   99  'Custom
      Picture         =   "MasterDB_frm.frx":19D3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "HAPUS SEMUA"
      DownPicture     =   "MasterDB_frm.frx":269D
      Height          =   855
      Index           =   1
      Left            =   6000
      MouseIcon       =   "MasterDB_frm.frx":3367
      MousePointer    =   99  'Custom
      Picture         =   "MasterDB_frm.frx":3671
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "HAPUS DATA"
      DownPicture     =   "MasterDB_frm.frx":4FF3
      Height          =   855
      Index           =   0
      Left            =   4320
      MouseIcon       =   "MasterDB_frm.frx":5CBD
      MousePointer    =   99  'Custom
      Picture         =   "MasterDB_frm.frx":5FC7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080FF80&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH DATA =>"
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
      Left            =   6240
      TabIndex        =   8
      Top             =   1440
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH DATA =>"
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
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   1440
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT TABLE"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER DATABASE AKUNTANSI"
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
      Width           =   3975
   End
End
Attribute VB_Name = "MasterDB_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Data1.Caption = Combo1
Data1.RecordSource = Combo1
Data1.Refresh
Label2(2).Caption = Data1.Recordset.RecordCount
If Data1.Recordset.BOF Then
    Command1(0).Visible = False
    Command1(1).Visible = False
Else
    Command1(0).Visible = True
    Command1(1).Visible = True
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Not Data1.Recordset.BOF Then
        x = MsgBox("Apakah anda yakin ?", vbYesNo, "Hapus Data")
        If x = vbYes Then
            Data1.Recordset.Delete
            Data1.Recordset.MoveNext
        End If
        Data1.Refresh
    End If
Case 1
    If Not Data1.Recordset.BOF Then
        x = MsgBox("Apakah anda yakin ?", vbYesNo, "Hapus Semua Data")
        If x = vbYes Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
                Data1.Recordset.Delete
                Data1.Recordset.MoveNext
            Loop
            Data1.Refresh
        End If
    End If
Case 2
    Unload Me
    Akuntansi_frm.Enabled = True
    Akuntansi_frm.Show
End Select
End Sub

Private Sub Form_Activate()
Label2(2).Caption = Data1.Recordset.RecordCount

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbakuntansi.mdb"
isi_cmb1

End Sub

Sub isi_cmb1()
Combo1.Clear
Combo1.AddItem "Buku_Besar"
Combo1.AddItem "temp_Buku_Besar"
Combo1.AddItem "jurnal"
Combo1.AddItem "bantupiutang"
Combo1.AddItem "bantuhutang"
Combo1.AddItem "bantuhutanglancar"
Combo1.AddItem "bantuhutangpanjang"
Combo1.AddItem "temppiutang"
Combo1.AddItem "temphutang"
Combo1.AddItem "temphutanglancar"
Combo1.AddItem "temphutangpanjang"
Combo1.AddItem "Akuntbl"
Combo1.ListIndex = 0
End Sub
