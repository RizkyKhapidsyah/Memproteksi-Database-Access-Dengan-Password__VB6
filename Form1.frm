VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memproteksi Database Access dengan Password"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub SetDatabasePassword(DBFile As String, _
NewPassword$)
On Error GoTo salah
    Dim db As Database
    'Buka file database
    Set db = OpenDatabase(DBFile, True)
    'Set password baru
    db.NewPassword "", NewPassword$
    'db.NewPassword "", ""
    'Tutup file database
    db.Close
    Exit Sub
salah:
Select Case Err.Number
       Case 3024
            MsgBox "File tidak ditemukan atau path file salah!", vbCritical, "File Tidak Ditemukan"
            End
       Case 3031
            MsgBox "File sudah dipassword!", _
                   vbCritical, "File sudah dipassword"
            End
       Case 3044
            MsgBox "Nama direktori/path salah!", _
                   vbCritical, "Direktori Salah"
            End
       Case Else
            MsgBox Err.Number & vbCrLf & _
                   Err.Description & vbCrLf & _
                   "Hubungi programmer Anda !", _
                   vbInformation, "Peringatan"
            End
End Select
End Sub

Private Sub Command1_Click()
    NewPassword$ = InputBox("Masukkan password: ", "Set Password Baru")
    If NewPassword$ = "" Then Exit Sub
    Call SetDatabasePassword(App.Path & "\Akademik.mdb", NewPassword$)
    MsgBox "File berhasil dipassword!", _
            vbInformation, "Sukses Password"
End Sub

Public Sub ClearDatabasePassword(DBFile As String, OldPassword$)
On Error GoTo salah
    Dim db As Database
    'Buka file database
    Set db = OpenDatabase(DBFile, True, False, ";pwd=" & OldPassword$)
    'Hapus password jika berhasil membuka file tsb
    'db.NewPassword OldPassword$, ""
    'Tutup database
    db.Close
    Exit Sub
salah:
Select Case Err.Number
       Case 3024
            MsgBox "File tidak ditemukan atau path file salah!", vbCritical, "File Tidak Ditemukan"
            End
       Case 3031
            MsgBox "Password salah!", vbCritical, _
                   "Password Salah"
            End
       Case 3044
            MsgBox "Nama direktori/path salah!", _
                   vbCritical, "Direktori Salah"
            End
       Case Else   'Kasus lainnya, silahkan
                   'diterjemahkan sendiri
            MsgBox Err.Number & vbCrLf & _
                   Err.Description & vbCrLf & _
                   "Hubungi programmer Anda !", _
                   vbInformation, "Peringatan"
            End
End Select
End Sub

Private Sub Command2_Click()
    OldPassword$ = InputBox("Masukkan password lama: ", "Hapus Password")
    Call ClearDatabasePassword(App.Path & "\Akademik.mdb", OldPassword$)
    MsgBox "Password berhasil dihapus!", _
            vbInformation, "Sukses Hapus Password"
End Sub


