VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Akuntansi_frm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Waynh Accounting System"
   ClientHeight    =   10830
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15270
   ClipControls    =   0   'False
   Icon            =   "Akuntansi_frm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   10680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
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
      Height          =   495
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2640
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":2F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":41CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":5450
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":66D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":7954
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":8BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":9E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":B0DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":C35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":D5DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":E860
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":FAE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":10D64
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":11FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":13268
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":144EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":1576C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":169EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":17C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":18EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":1A174
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":1B3F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   1535
      ButtonWidth     =   2593
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tabel Akun"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buku Besar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bantu Piutang"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bantu Hutang"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hutang Lancar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hutang Jk.Panjang"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Jurnal"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Neraca Lajur"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Keluar"
            ImageIndex      =   10
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Akuntansi_frm.frx":1C678
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":1C992
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":1D66C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":1E346
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":1F020
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":1FCFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":209D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":216AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":22388
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":23062
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":23D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":24A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Akuntansi_frm.frx":256F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   10455
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Picture         =   "Akuntansi_frm.frx":263CA
            Text            =   "WaynhSoft"
            TextSave        =   "WaynhSoft"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2858
            Text            =   "Telp : 021-9389 4481"
            TextSave        =   "Telp : 021-9389 4481"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4710
            Text            =   "Email : WahyuNHidayat@Gmail.com"
            TextSave        =   "Email : WahyuNHidayat@Gmail.com"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4763
            Text            =   "YM : Wahyu_NHidayat@Yahoo,com"
            TextSave        =   "YM : Wahyu_NHidayat@Yahoo,com"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6773
            Text            =   "www.Pak-Professor.com"
            TextSave        =   "www.Pak-Professor.com"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "3:38 AM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "5/17/2008"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "WaynhSoft"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   14160
      MouseIcon       =   "Akuntansi_frm.frx":2C792
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   9720
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tutup Siklus Akuntansi"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   12
      Left            =   11520
      MouseIcon       =   "Akuntansi_frm.frx":2CA9C
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   8280
      Width           =   3570
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   10920
      X2              =   11640
      Y1              =   6240
      Y2              =   6720
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   10920
      X2              =   11520
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   10920
      X2              =   11520
      Y1              =   6000
      Y2              =   5520
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   8040
      X2              =   8760
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   4920
      X2              =   11760
      Y1              =   6120
      Y2              =   7320
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   4920
      X2              =   6480
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008000&
      BorderStyle     =   3  'Dot
      X1              =   2760
      X2              =   3960
      Y1              =   7320
      Y2              =   6120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00008000&
      BorderStyle     =   3  'Dot
      X1              =   2760
      X2              =   3960
      Y1              =   6720
      Y2              =   6120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00008000&
      BorderStyle     =   3  'Dot
      X1              =   2880
      X2              =   3960
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00008000&
      BorderStyle     =   3  'Dot
      X1              =   2880
      X2              =   3960
      Y1              =   5520
      Y2              =   6120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      BorderStyle     =   3  'Dot
      X1              =   2640
      X2              =   3960
      Y1              =   4920
      Y2              =   6120
   End
   Begin VB.Shape Shape7 
      Height          =   735
      Left            =   11400
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buku Pembantu"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   18
      Left            =   11640
      MouseIcon       =   "Akuntansi_frm.frx":2CDA6
      TabIndex        =   18
      Top             =   7080
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Neraca"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   17
      Left            =   11640
      MouseIcon       =   "Akuntansi_frm.frx":2D0B0
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   6480
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Laporan Saldo Laba"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   16
      Left            =   11640
      MouseIcon       =   "Akuntansi_frm.frx":2D3BA
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Laporan Rugi Laba"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   15
      Left            =   11640
      MouseIcon       =   "Akuntansi_frm.frx":2D6C4
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   5280
      Width           =   2910
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Neraca Lajur"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   14
      Left            =   8760
      MouseIcon       =   "Akuntansi_frm.frx":2D9CE
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   5880
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buku Besar"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   13
      Left            =   6240
      MouseIcon       =   "Akuntansi_frm.frx":2DCD8
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5880
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jurnal"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   11
      Left            =   3960
      MouseIcon       =   "Akuntansi_frm.frx":2DFE2
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5880
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bukti Memorial"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Top             =   7080
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bukti Penjualan"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   10
      Top             =   6480
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bukti Pembelian"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   9
      Top             =   5880
      Width           =   2580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bukti Kas (Bank)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   5280
      Width           =   2595
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bukti Kas"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Otomatis"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   5
      Left            =   6480
      TabIndex        =   6
      Top             =   3600
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manual"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   4
      Left            =   2160
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OUTPUT"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   3
      Left            =   12360
      TabIndex        =   4
      Top             =   2880
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROSES"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   2
      Left            =   6360
      TabIndex        =   3
      Top             =   2880
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SIKLUS AKUNTANSI"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   14895
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Top             =   2760
      Width           =   15135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   3975
      Left            =   11280
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   3975
      Left            =   8400
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   3975
      Left            =   5640
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   3975
      Left            =   3000
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   3975
      Left            =   120
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   5640
      Top             =   3480
      Width           =   9615
   End
   Begin VB.Menu msd_mnu 
      Caption         =   "Master Data"
   End
   Begin VB.Menu period_mnu 
      Caption         =   "New Period"
      Visible         =   0   'False
   End
   Begin VB.Menu akuntbl_mnu 
      Caption         =   "Tabel Akun"
   End
   Begin VB.Menu BB_mnu 
      Caption         =   "Buku Besar"
   End
   Begin VB.Menu bbantu_mnu 
      Caption         =   "Buku Pembantu"
      Begin VB.Menu bpiutang_mnu 
         Caption         =   "Piutang"
      End
      Begin VB.Menu bhutang_mnu 
         Caption         =   "Hutang"
         Begin VB.Menu htgusaha_mnu 
            Caption         =   "Hutang Usaha"
         End
         Begin VB.Menu htgLancar_mnu 
            Caption         =   "Hutang Lancar Lainnya"
         End
         Begin VB.Menu htgPanjang_mnu 
            Caption         =   "Hutang Jangka Panjang"
         End
      End
   End
   Begin VB.Menu Jurnal_mnu 
      Caption         =   "Jurnal"
   End
   Begin VB.Menu NLajur_mnu 
      Caption         =   "Neraca Lajur"
   End
   Begin VB.Menu lap_mnu 
      Caption         =   "Laporan"
      Begin VB.Menu lapLR_mnu 
         Caption         =   "Laba-Rugi"
      End
      Begin VB.Menu laplabat_mnu 
         Caption         =   "Laba ditahan"
      End
      Begin VB.Menu lneraca_mnu 
         Caption         =   "Neraca"
      End
   End
   Begin VB.Menu bantu_mnu 
      Caption         =   "Bantuan"
      Begin VB.Menu about_mnu 
         Caption         =   "About"
      End
      Begin VB.Menu kal_mnu 
         Caption         =   "Kalkulator"
      End
      Begin VB.Menu SA_mnu 
         Caption         =   "Siklus Akuntansi"
      End
      Begin VB.Menu manual_mnu 
         Caption         =   "User Manual"
      End
   End
   Begin VB.Menu toolmnu 
      Caption         =   "Toolbar"
      Begin VB.Menu onTool_mnu 
         Caption         =   "On"
      End
      Begin VB.Menu offtool_mnu 
         Caption         =   "Off"
      End
   End
   Begin VB.Menu keluar_mnu 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Akuntansi_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnakhir As Byte
Dim thnakhir As Single

Sub umpetSA()
Dim a As Byte
a = 0
Do Until a = 19
    Label1(a).Visible = False
    a = a + 1
Loop
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = False
    Shape5.Visible = False
    Shape6.Visible = False
    Shape7.Visible = False
    Shape8.Visible = False
    Shape9.Visible = False
    Shape10.Visible = False
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
Line11.Visible = False
End Sub

Sub munculSA()
Dim a As Byte
a = 0
Do Until a = 19
    Label1(a).Visible = True
    a = a + 1
Loop
    Shape1.Visible = True
    Shape2.Visible = True
    Shape3.Visible = True
    Shape4.Visible = True
    Shape5.Visible = True
    Shape6.Visible = True
    Shape7.Visible = True
    Shape8.Visible = True
    Shape9.Visible = True
    Shape10.Visible = True
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Line5.Visible = True
Line6.Visible = True
Line7.Visible = True
Line8.Visible = True
Line9.Visible = True
Line10.Visible = True
Line11.Visible = True
End Sub

Private Sub about_mnu_Click()
WaynhSoft_frm.Show
End Sub

Private Sub akuntbl_mnu_Click()
Me.Enabled = False
AkunTbl_frm.Show
End Sub

Private Sub BB_mnu_Click()
bb_frm.Show
Me.Enabled = False
End Sub

Private Sub bpiutang_mnu_Click()
Me.Enabled = False
BantuPiutang_frm.Show
End Sub

Private Sub Form_Activate()
bb_frm.Data1.RecordSource = "select*from buku_besar"
bb_frm.Data1.Refresh
If bb_frm.Data1.Recordset.BOF Then
    MsgBox "Data masih kosong... (auto create data)", vbInformation, "Create New Data"
    buat_data
'    Period_frm.Show
Else
    cek_periode
End If
'Toolbar1.Visible = False
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh
Data5.Refresh
End Sub


Sub cek_periode()
Dim a, b As Byte
Dim blnskrng As Byte
Dim thnskrng As Single
'ambil bln saat ini
blnskrng = Month(Date)
thnskrng = Year(Date)
'get bln data terakhir
With bb_frm.Data1.Recordset
    .MoveLast
    b = 1
    a = 15
    Do Until b = 13
        If !bulan = MonthName(Month(a)) Then
            blnakhir = Month(a)
        End If
        b = b + 1
        a = a + 31
    Loop
    thnakhir = !tahun
    If blnakhir < blnskrng Or thnakhir < thnskrng Then
        MsgBox "Anda sudah masuk periode baru akuntansi... (auto create data)", vbInformation, "New Period"
        isi_datatemp
        trans_data
        hapus_temp
        bb_frm.Data1.Refresh
    End If
End With
End Sub

Sub isi_datatemp()
Dim a As String
With bb_frm.Data1.Recordset
    .MoveFirst
    Do While Not .EOF
        If !bulan = MonthName(blnakhir) And !tahun = thnakhir Then
                Data1.Recordset.AddNew
                Data1.Recordset!no_akun = !no_akun
                With AkunTbl_frm.Data1.Recordset
                    If Not .BOF Then
                        .MoveFirst
                        Do While Not .EOF
                            If !no_akun = bb_frm.Data1.Recordset!no_akun Then
                                a = !nrlr
                                .MoveLast
                            End If
                            .MoveNext
                        Loop
                    End If
                End With
                If a = "LR" Then
                    Data1.Recordset!saldo_awal = !saldo_akhir
                    Data1.Recordset!saldo_akhir = !saldo_akhir
                    Data1.Recordset.Update
                Else
                    Data1.Recordset!saldo_awal = !saldo_akhir
                    Data1.Recordset!saldo_akhir = !saldo_akhir
                    Data1.Recordset.Update
                End If
        End If
        .MoveNext
    Loop
    Data1.Refresh
End With

With BantuPiutang_frm.Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !bulan = MonthName(blnakhir) And !tahun = thnakhir Then
                Data2.Recordset.AddNew
                Data2.Recordset!no_akun = !no_akun
                Data2.Recordset!nm_piutang = !nm_piutang
                Data2.Recordset!saldo_awal = !saldo_akhir
                Data2.Recordset!saldo_akhir = !saldo_akhir
                Data2.Recordset.Update
            End If
            .MoveNext
        Loop
        Data2.Refresh
    End If
End With

With BantuHutang_frm.Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !bulan = MonthName(blnakhir) And !tahun = thnakhir Then
                Data3.Recordset.AddNew
                Data3.Recordset!no_akun = !no_akun
                Data3.Recordset!nm_hutang = !nm_hutang
                Data3.Recordset!saldo_awal = !saldo_akhir
                Data3.Recordset!saldo_akhir = !saldo_akhir
                Data3.Recordset.Update
            End If
            .MoveNext
        Loop
        Data3.Refresh
    End If
End With
With BantuHutangLancar_frm.Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !bulan = MonthName(blnakhir) And !tahun = thnakhir Then
                Data4.Recordset.AddNew
                Data4.Recordset!no_akun = !no_akun
                Data4.Recordset!nm_hutang = !nm_hutang
                Data4.Recordset!saldo_awal = !saldo_akhir
                Data4.Recordset!saldo_akhir = !saldo_akhir
                Data4.Recordset.Update
            End If
            .MoveNext
        Loop
        Data4.Refresh
    End If
End With
With BantuHutangPanjang_frm.Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !bulan = MonthName(blnakhir) And !tahun = thnakhir Then
                Data5.Recordset.AddNew
                Data5.Recordset!no_akun = !no_akun
                Data5.Recordset!nm_hutang = !nm_hutang
                Data5.Recordset!saldo_awal = !saldo_akhir
                Data5.Recordset!saldo_akhir = !saldo_akhir
                Data5.Recordset.Update
            End If
            .MoveNext
        Loop
        Data3.Refresh
    End If
End With
End Sub

Sub trans_data()
Dim bln As Byte
bln = (Month(Date))
With bb_frm.Data1.Recordset
If Not Data1.Recordset.BOF Then
    Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
        .AddNew
        !bulan = MonthName(bln)
        !tahun = Format(Date, "yyyy")
        !no_akun = Data1.Recordset!no_akun
        !saldo_awal = Data1.Recordset!saldo_awal
        !saldo_debit = 0
        !saldo_kredit = 0
        !saldo_akhir = Data1.Recordset!saldo_akhir
        .Update
        Data1.Recordset.MoveNext
    Loop
End If
End With
With BantuPiutang_frm.Data1.Recordset
If Not Data2.Recordset.BOF Then
    Data2.Recordset.MoveFirst
    Do While Not Data2.Recordset.EOF
        .AddNew
        !bulan = MonthName(bln)
        !tahun = Format(Date, "yyyy")
        !no_akun = Data2.Recordset!no_akun
        !nm_piutang = Data2.Recordset!nm_piutang
        !saldo_awal = Data2.Recordset!saldo_awal
        !saldo_akhir = Data2.Recordset!saldo_akhir
        .Update
        Data2.Recordset.MoveNext
    Loop
End If
End With
With BantuHutang_frm.Data1.Recordset
If Not Data3.Recordset.BOF Then
    Data3.Recordset.MoveFirst
    Do While Not Data3.Recordset.EOF
        .AddNew
        !bulan = MonthName(bln)
        !tahun = Format(Date, "yyyy")
        !no_akun = Data3.Recordset!no_akun
        !nm_hutang = Data3.Recordset!nm_hutang
        !saldo_awal = Data3.Recordset!saldo_awal
        !saldo_akhir = Data3.Recordset!saldo_akhir
        .Update
        Data3.Recordset.MoveNext
    Loop
End If
End With
With BantuHutangLancar_frm.Data1.Recordset
If Not Data4.Recordset.BOF Then
    Data4.Recordset.MoveFirst
    Do While Not Data4.Recordset.EOF
        .AddNew
        !bulan = MonthName(bln)
        !tahun = Format(Date, "yyyy")
        !no_akun = Data4.Recordset!no_akun
        !nm_hutang = Data4.Recordset!nm_hutang
        !saldo_awal = Data4.Recordset!saldo_awal
        !saldo_akhir = Data4.Recordset!saldo_akhir
        .Update
        Data4.Recordset.MoveNext
    Loop
End If
End With
With BantuHutangPanjang_frm.Data1.Recordset
If Not Data5.Recordset.BOF Then
    Data5.Recordset.MoveFirst
    Do While Not Data5.Recordset.EOF
        .AddNew
        !bulan = MonthName(bln)
        !tahun = Format(Date, "yyyy")
        !no_akun = Data5.Recordset!no_akun
        !nm_hutang = Data5.Recordset!nm_hutang
        !saldo_awal = Data5.Recordset!saldo_awal
        !saldo_akhir = Data5.Recordset!saldo_akhir
        .Update
        Data5.Recordset.MoveNext
    Loop
End If
End With
BantuPiutang_frm.Data1.Refresh
BantuHutang_frm.Data1.Refresh
BantuHutangLancar_frm.Data1.Refresh
BantuHutangPanjang_frm.Data1.Refresh
bb_frm.Data1.Refresh
Me.Hide
Akuntansi_frm.Enabled = True
Akuntansi_frm.Show
End Sub

Sub hapus_temp()
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        .Delete
        .MoveNext
    Loop
End If
End With
Data1.Refresh

With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        .Delete
        .MoveNext
    Loop
End If
End With
Data2.Refresh

With Data3.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        .Delete
        .MoveNext
    Loop
End If
End With
Data3.Refresh

With Data4.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        .Delete
        .MoveNext
    Loop
End If
End With
Data4.Refresh

With Data5.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        .Delete
        .MoveNext
    Loop
End If
End With
Data5.Refresh

End Sub

Sub buat_data()
Dim bln As Byte
bln = (Month(Date))
    With bb_frm.Data1.Recordset
        AkunTbl_frm.Data1.Recordset.MoveFirst
        Do While Not AkunTbl_frm.Data1.Recordset.EOF
            If AkunTbl_frm.Data1.Recordset!dk <> "-" Then
                .AddNew
                !bulan = MonthName(bln)
                !tahun = Format(Date, "yyyy")
                !no_akun = AkunTbl_frm.Data1.Recordset!no_akun
                !saldo_awal = 0
                !saldo_debit = 0
                !saldo_kredit = 0
                !saldo_akhir = 0
                .Update
            End If
            AkunTbl_frm.Data1.Recordset.MoveNext
        Loop
        bb_frm.Data1.Refresh
        Me.Hide
        Akuntansi_frm.Enabled = True
        Akuntansi_frm.Show
    End With

End Sub


Private Sub Form_Click()
WaynhSoft_frm.Hide
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data1.RecordSource = "temp_buku_besar"
Data2.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data2.RecordSource = "temppiutang"
Data3.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data3.RecordSource = "temphutang"
Data4.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data4.RecordSource = "temphutanglancar"
Data5.DatabaseName = App.Path & "\dbakuntansi.mdb"
Data5.RecordSource = "temphutangpanjang"
'umpetSA
onTool_mnu.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub htgLancar_mnu_Click()
Me.Enabled = False
BantuHutangLancar_frm.Show
End Sub

Private Sub htgPanjang_mnu_Click()
Me.Enabled = False
BantuHutangPanjang_frm.Show
End Sub

Private Sub htgusaha_mnu_Click()
Me.Enabled = False
BantuHutang_frm.Show
End Sub

Private Sub Jurnal_mnu_Click()
Me.Enabled = False
jurnal_frm.Show
End Sub

Private Sub kal_mnu_Click()
    AppActivate Shell("calc.exe", 1)
End Sub

Private Sub keluar_mnu_Click()
Dim x
x = MsgBox("Apakah anda yakin ingin keluar dari aplikasi?", vbYesNo, "KELUAR...")
If x = vbYes Then
    End
End If
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 11
    Jurnal_mnu_Click
Case 13
    BB_mnu_Click
Case 14
    NLajur_mnu_Click
Case 15
    lapLR_mnu_Click
Case 16
    laplabat_mnu_Click
Case 17
    lneraca_mnu_Click
Case 12
    umpetSA
End Select
End Sub

Private Sub Label2_Click()
Form2.Show
End Sub

Private Sub laplabat_mnu_Click()
LabaDitahan_frm.Show
Me.Enabled = False
End Sub

Private Sub lapLR_mnu_Click()
LabaRugi_frm.Show
Akuntansi_frm.Enabled = False
End Sub

Private Sub lneraca_mnu_Click()
Neraca_frm.Show
Me.Enabled = False
End Sub

Private Sub manual_mnu_Click()
Call uncons
End Sub

Private Sub msd_mnu_Click()
MasterDB_frm.Show
Me.Enabled = False
End Sub

Private Sub NLajur_mnu_Click()
Me.Enabled = False
NeracaLAjur_frm.Show
End Sub

Private Sub offtool_mnu_Click()
Toolbar1.Visible = False
onTool_mnu.Enabled = True
offtool_mnu.Enabled = False
End Sub

Private Sub onTool_mnu_Click()
Toolbar1.Visible = True
onTool_mnu.Enabled = False
offtool_mnu.Enabled = True
End Sub


Private Sub SA_mnu_Click()
munculSA
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    akuntbl_mnu_Click
Case 2
    BB_mnu_Click
Case 3
    bpiutang_mnu_Click
Case 4
    htgusaha_mnu_Click
Case 5
    htgLancar_mnu_Click
Case 6
    htgPanjang_mnu_Click
Case 7
    Jurnal_mnu_Click
Case 8
    NLajur_mnu_Click
Case 9
    keluar_mnu_Click
End Select
End Sub
