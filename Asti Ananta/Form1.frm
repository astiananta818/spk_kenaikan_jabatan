VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.Image a 
      Height          =   6690
      Left            =   -720
      Picture         =   "Form1.frx":0000
      Top             =   -480
      Width           =   10545
   End
   Begin VB.Menu FILE 
      Caption         =   "FILE"
      Begin VB.Menu LOGIN 
         Caption         =   "LOGIN"
      End
   End
   Begin VB.Menu PROSES 
      Caption         =   "PROSES"
      Begin VB.Menu NILAI_KRITERIA 
         Caption         =   "NILAI KRITERIA"
      End
      Begin VB.Menu NILAI_AKHIR 
         Caption         =   "NILAI AKHIR"
      End
   End
   Begin VB.Menu LAPORAN 
      Caption         =   "LAPORAN"
      Begin VB.Menu REPORT 
         Caption         =   "REPORT"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LOGIN_Click()
Form2.Show
End Sub

Private Sub NILAI_AKHIR_Click()
Form3.Show
End Sub

Private Sub NILAI_KRITERIA_Click()
Form4.Show
End Sub

