VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghitung Selisih Dua Buah Tanggal (3)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   1800
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function SelisihHariJam(ByVal Awal As Date, _
ByVal Akhir As Date) As String
Dim Detik As Long, Hari As Long, Jam As Long
Dim JamLengkap As String
   
If Awal > Akhir Then
   MsgBox "Tanggal dan waktu awal harus lebih kecil " _
   & vbCrLf & _
            "dari pada tanggal dan waktu akhir", _
            vbCritical, "Peringatan"
     Exit Function
  End If
  
  'Tampung dalam durasi satuan terkecil, yaitu: DETIK
  Detik = DateDiff("s", Awal, Akhir)
  
  'Hitung jumlah jam dgn cara membagi 3600
  '(backslash ("\") supaya menghasilkan nilai Integer
  'tanpa pembulatan ke atas)
  Jam = Detik \ 3600
  
  'Jika jumlah jam lebih besar dari 23 artinya: lebih
  'dari 1 hari
  If Jam > 23 Then
     'Hitung jumlah hari dgn car membagi 24
     '(backslash ("\") supaya menghasilkan nilai
     'integer tanpa pembulatan ke atas)
     Hari = Jam \ 24
     
     'Hitung Durasi Jam dalam hh:mm:ss
     JamLengkap = Format((Akhir - Awal), "hh:mm:ss")
    Else 'Jika jumlah jam <= 23
     Hari = 0   'maka jumlah hari = nol
     'Hitung Durasi Jam dalam hh:mm:ss
     JamLengkap = Format((Akhir - Awal), "hh:mm:ss")
  End If
  
  If Hari = 0 Then  'Jika jumlah hari = 0
     'Tampung hasil akhirnya
     SelisihHariJam = JamLengkap
  
  Else  'Jika jumlah hari > 0, tampilkan jumlah harinya
     'Tampung hasil akhirnya
     SelisihHariJam = Hari & " hari, " & JamLengkap
  End If
  Exit Function

End Function

Private Sub Form_Load()
  Timer1.Interval = 500
  Timer1.Enabled = True
  Text1.Text = "01/03/2002 07:18:00"
  'Text2.Text = "01/09/2002 09:42:30"
  Text2.Text = Now
End Sub

Private Sub Timer1_Timer()
On Error GoTo Pesan
  Text2.Text = Now
  Label1.Caption = SelisihHariJam(CDate(Text1.Text), _
                      CDate(Text2.Text))
  Exit Sub
Pesan:
  MsgBox "Tanggal atau format-nya salah!", _
         vbCritical, "Error Tanggal"
End Sub



