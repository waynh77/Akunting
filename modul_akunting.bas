Attribute VB_Name = "modul_akunting"
Option Explicit
Dim db As String

Public Sub bukadb()
db = App.Path + "\dbakuntansi.mdb"
End Sub

Public Sub uncons()
    MsgBox "Maaf fitur yg anda pilih msh dalam proses pengembangan...", vbOKOnly, "Under Construction"
End Sub

Public Sub msgkosong()
    MsgBox "Data masih kosong, silahkan diisi terlebih dahulu", vbOKOnly, "Data Kosong"
End Sub

Public Sub msgValDat()
    MsgBox "Data belum lengkap...", vbOKOnly, "Validasi Data"
End Sub

Public Sub Main()
Akuntansi_frm.Show
WaynhSoft_frm.Show
'cetakAkun_frm.Show
'Neraca_frm.Show
'LabaDitahan_frm.Show
'BantuPiutang_frm.Show
'jurnal_frm.Show
'AkunTbl_frm.Show
'bb_frm.Show
'NeracaLAjur_frm.Show
'LabaRugi_frm.Show
End Sub

