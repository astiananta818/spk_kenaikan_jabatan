VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6750
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Comexit 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   4680
      TabIndex        =   8
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Comlogin 
      Caption         =   "LOGIN"
      Height          =   735
      Left            =   2160
      TabIndex        =   6
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox txtjabatan 
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtpasword 
      Height          =   735
      Left            =   3600
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtusername 
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   " DATA LOGIN"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "JABATAN"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "PASWORD"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "USER NAME"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Top             =   480
      Width           =   10575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsLOGIN As New ADODB.Recordset
  
Private Sub Comexit_Click()
x = MsgBox("Yakin Keluar?", vbQuestion + vbYesNo, "informasi")
If x = vbYes Then End
End Sub

Private Sub Comlogin_Click()
user = txtusername.Text
pasword = txtpasword.Text
jabatan = txtjabatan.Text
If user = "asti" And pasword = "1001" And jabatan = "staff" Then
MsgBox "selamat datang"
Form3.Show
Form2.Hide
Else
LOGIN = LOGIN + 1
MsgBox "anda salah memasukkan pasword" & LOGIN & " kali "
If LOGIN = 2 Then
MsgBox "kesempatan anda satu kali lagi", vbExclamation
End If
If LOGIN = 3 Then
MsgBox "anda sudah salah memasukkan pasword 3 kali, maka program kali ini akan kami tutup!", vbCritical
End
End If
End If
End Sub


