VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form4"
   ScaleHeight     =   7695
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Comkeluar 
      Caption         =   "KELUAR"
      Height          =   495
      Left            =   7200
      TabIndex        =   17
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Comhapus 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   7200
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Comedit 
      Caption         =   "EDIT"
      Height          =   615
      Left            =   7200
      TabIndex        =   15
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Comsimpan 
      Caption         =   "SIMPAN"
      Height          =   615
      Left            =   7200
      TabIndex        =   14
      Top             =   1320
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   600
      TabIndex        =   13
      Top             =   5760
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   3201
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtperilaku 
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox txtpenilaiankinerja 
      Height          =   615
      Left            =   2880
      TabIndex        =   10
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox txtmasakerja 
      Height          =   615
      Left            =   2880
      TabIndex        =   9
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox txtmasukkankriteria 
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtjabatan 
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtnamakaryawan 
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "DATA NILAI KRITERIA"
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "PERILAKU"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "PENILAIAN KINERJA"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "MASA KERJA"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "MASUKKAN KRITERIA"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "JABATAN"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "NAMA KARYAWAN"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   6690
      Left            =   -360
      Picture         =   "Form4.frx":0000
      Top             =   0
      Width           =   10545
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsnilaikriteria As New ADODB.Recordset

Private Sub Comedit_Click()
koneksidb.Execute "update tabel_nilaikriteria set namakaryawan='" & txtnamakaryawan.Text & "',jabatan='" & txtjabatan & "',masukkankriteria='" & txtmasukkankriteria & "',masakerja='" & txtmasakerja & "',penilaiankinerja='" & txtpenilaiankinerja & "',perilaku='" & txtperilaku & "'"
Call update
Call edit_grid
Call kosong
End Sub

Private Sub Comhapus_Click()
koneksidb.Execute "delete from tabel_nilaikriteria where namakaryawan='" & txtnamakaryawan & "'"
Call refreshh
Call kosong
txtnamakaryawan.SetFocus
End Sub

Private Sub Comkeluar_Click()
x = MsgBox("Yakin Keluar?", vbQuestion + vbYesNo, "informasi")
If x = vbYes Then End
End Sub

Private Sub Comsimpan_Click()
If txtnamakaryawan = "" Then
MsgBox "Nama Karyawan Kosong", vbExclamation, "pesan"
txtnamakaryawan.SetFocus
Exit Sub
End If
    If txtjabatan = "" Then
    MsgBox "Jabatan Kosong", vbExclamation, "pesan"
    txtjabatan.SetFocus
    Exit Sub
    End If
If txtmasukkankriteria = "" Then
MsgBox "Masukka Kritria Kosong", vbExclamation, "pesan"
txtmasukkankriteria.SetFocus
Exit Sub
End If
    If txtmasakerja = "" Then
    MsgBox "Masa Kerja Kosong", vbExclamation, "pesan"
    txtmasakerja.SetFocus
    Exit Sub
    End If
Set rsnilaikriteria = New ADODB.Recordset
rsnilaikriteria.Open "select*from tabel_nilaikriteria where namakaryawan='" & txtnamakaryawan & "'", koneksidb
If Not rsnilaikriteria.EOF Then
MsgBox "Nama Karyawan sudah ada", vbCritical, "pesan"
txtnamakaryawan = ""
txtnamakaryawan.SetFocus
Exit Sub
Else
koneksidb.Execute "insert into tabel_nilaikriteria(namakaryawan,jabatan,masukkankriteria,masakerja,penilaiankinerja,perilaku) value ('" & txtnamakaryawan & "','" & txtjabatan & "','" & txtmasukkankriteria & "','" & txtmasakerja & "','" & txtpenilaiankinerja & "','" & txtperilaku & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = rsnilaikriteria
With DataGrid1
End With
Call edit_grid
End If
End Sub

Private Sub DataGrid1_Click()
txtnamakaryawan.Text = rsnilaikriteria!namakaryawan
txtjabatan.Text = rsnilaikriteria!jabatan
txtmasukkankriteria.Text = rsnilaikriteria!masukkankriteria
txtmasakerja.Text = rsnilaikriteria!masakerja
txtpenilaiankinerja.Text = rsnilaikriteria!penilaiankinerja
txtperilaku.Text = rsnilaikriteria!perilaku
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rsnilaikriteria
With rsnilaikriteria
End With
Call edit_grid
End Sub
Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "Nama Karyawan"
    .Columns(1).Caption = "Jabatan"
    .Columns(2).Caption = "Masukkan Kriteria"
    .Columns(3).Caption = "Masa Kerja"
    .Columns(4).Caption = "Penilaian Kinerja"
    .Columns(5).Caption = "Perilaku"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
    .Columns(3).Width = 1200
    .Columns(4).Width = 1200
    .Columns(5).Width = 1200
End With
End Sub

Sub tampil_data()
Set rsnilaikriteria = New ADODB.Recordset
rsnilaikriteria.ActiveConnection = koneksidb
rsnilaikriteria.CursorLocation = adUseClient
rsnilaikriteria.LockType = adLockOptimistic
rsnilaikriteria.Source = "select * from tabel_nilaikriteria"
rsnilaikriteria.Open
End Sub

Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rsnilaikriteria
With DataGrid1
End With
End Sub

Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rsnilaikriteria
With DataGrid1
End With
Call edit_grid
End Sub

Sub kosong()
txtnamakaryawan = ""
txtjabatan = ""
txtmasukkankriteria = ""
txtmasakerja = ""
txtpenilaiankinerja = ""
txtperilaku = ""
txtnamakaryawan.SetFocus
End Sub

Private Sub Text3_Change()

End Sub
