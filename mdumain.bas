Attribute VB_Name = "mdumain"


Public fMainform As MDIAkuntansi
Public Sub Main()
'End
Set fMainform = New MDIAkuntansi
Load fMainform
fMainform.Show
End Sub

