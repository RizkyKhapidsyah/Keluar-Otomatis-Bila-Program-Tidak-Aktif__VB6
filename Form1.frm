VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Keluar Otomatis Bila Program Tidak Aktif"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim awal As Date
Dim Gerak As Boolean
Dim Aksi As Boolean

Private Sub Form_Load()
  'Inisialisasi semua variabel dan Timer
  Gerak = False
  Aksi = False
  Timer1.Interval = 500
  Timer1.Enabled = True
  awal = Time
End Sub

Private Sub Form_MouseMove(Button As Integer, _
Shift As Integer, X As Single, Y As Single)
'Jika ada pergerakan mouse di form, set waktu mulai
'untuk perhitungan durasi dengan waktu saat itu
   awal = Time
   'Update status...
   Aksi = True
End Sub

Private Sub Timer1_Timer()
Dim durasi As Date
  Aksi = False
  'Periksa...
  If Aksi = False Then
     Gerak = False
     Timer1.Enabled = True
  Else 'Jika ada perubahan di Mouse_Move
     Gerak = True
     Timer1.Enabled = False
  End If
  Text1.Text = awal
  Text2.Text = Time
  'Jika tidak ada pergerakan, aktifkan perhitungan
  'durasi
  If Gerak = False Then
    durasi = Time - awal
    'Dalam contoh ini, jika 5 detik aplikasi tidak
    'mengalami kegiatan, maka langsung keluar...
    If Format(durasi, "hh:mm:ss") = "00:00:05" Then
       'Sebelum keluar, bebaskan semua variabel di form
       'ini
       Set Form1 = Nothing
       Unload Me
    End If
  End If
End Sub


