VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form3"
   ScaleHeight     =   7665
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   600
      TabIndex        =   17
      Top             =   5280
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   3836
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
   Begin VB.CommandButton Comkeluar 
      Caption         =   "KELUAR"
      Height          =   495
      Left            =   7800
      TabIndex        =   16
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Comhapus 
      Caption         =   "HAPUS"
      Height          =   615
      Left            =   7800
      TabIndex        =   15
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Comedit 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   7800
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Comsimpan 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   7680
      TabIndex        =   13
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ComboBox Comjeniskelamin 
      Height          =   315
      Left            =   4200
      TabIndex        =   12
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtalamat 
      Height          =   405
      Left            =   4200
      TabIndex        =   11
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtbagian 
      Height          =   405
      Left            =   4200
      TabIndex        =   10
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtjabatan 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtnamakaryawan 
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtnik 
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "JENIS KELAMIN"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "ALAMAT"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "BAGIAN"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "JABATAN"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "NAMA KARYAWAN"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "NIK"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DATA NILAI AKHIR"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   6690
      Left            =   120
      Picture         =   "Form3.frx":0000
      Top             =   120
      Width           =   10545
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nilaiakhir As New ADODB.Recordset
Private Sub Comedit_Click()
koneksidb.Execute "update tabel_nilaiakhir set nik='" & txtnik.Text & "',namakaryawan='" & txtnamakaryawan & "',jabatan='" & txtjabatan & "',bagian='" & txtbagian & "',alamat='" & txtalamat & "',jeniskelamin='" & Comjeniskelamin & "'"
Call update
Call edit_grid
Call kosong
End Sub

Private Sub Comhapus_Click()
koneksidb.Execute "delete from tabel_nilaiakhir where nik='" & txtnik & "'"
Call refreshh
Call kosong
txtnik.SetFocus
End Sub

Private Sub Comkeluar_Click()
x = MsgBox("Yakin Keluar?", vbQuestion + vbYesNo, "informasi")
If x = vbYes Then End
End Sub

Private Sub Comsimpan_Click()
If txtnik = "" Then
MsgBox "Nik Kosong", vbExclamation, "pesan"
txtnik.SetFocus
Exit Sub
End If
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
    If txtbagian = "" Then
    MsgBox "Bagian Kosong", vbExclamation, "pesan"
    txtbagian.SetFocus
    Exit Sub
    End If
If txtalamat = "" Then
MsgBox "Alamat Kosong", vbExclamation, "pesan"
txtalamat.SetFocus
Exit Sub
End If
    If Comjeniskelamin = "" Then
    MsgBox "Jenis Kelamin Kosong", vbExclamation, "pesan"
    Comjeniskelamin.SetFocus
    Exit Sub
    End If
Set nilaiakhir = New ADODB.Recordset
nilaiakhir.Open "select*from tabel_nilaiakhir where nik='" & txtnik & "'", koneksidb
If Not nilaiakhir.EOF Then
MsgBox "Nik sudah ada", vbCritical, "pesan"
txtnik = ""
txtnik.SetFocus
Exit Sub
Else
koneksidb.Execute "insert into tabel_nilaiakhir(nik,namakaryawan,jabatan,bagian,alamat,jeniskelamin) value ('" & txtnik & "','" & txtnamakaryawan & "','" & txtjabatan & "','" & txtbagian & "','" & txtalamat & "','" & Comjeniskelamin & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = nilaiakhir
With DataGrid1
End With
Call edit_grid
End If
End Sub

Private Sub DataGrid1_Click()
txtnik.Text = nilaiakhir!nik
txtnamakaryawan.Text = nilaiakhir!namakaryawan
txtjabatan.Text = nilaiakhir!jabatan
txtbagian.Text = nilaiakhir!bagian
txtalamat.Text = nilaiakhir!alamat
Comjeniskelamin.Text = nilaiakhir!jeniskelamin
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = nilaiakhir
With nilaiakhir
With Comjeniskelamin
    .AddItem " Laki-Laki"
    .AddItem " Perempuan "
End With
Call edit_grid
End With
End Sub

Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "Nik"
    .Columns(1).Caption = "Nama Karyawan"
    .Columns(2).Caption = "Jabatan"
    .Columns(3).Caption = "Bagian"
    .Columns(4).Caption = "Alamat"
    .Columns(5).Caption = "Jenis Kelamin"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
    .Columns(3).Width = 1200
    .Columns(4).Width = 1200
    .Columns(5).Width = 1200
End With
End Sub

Sub tampil_data()
Set nilaiakhir = New ADODB.Recordset
nilaiakhir.ActiveConnection = koneksidb
nilaiakhir.CursorLocation = adUseClient
nilaiakhir.LockType = adLockOptimistic
nilaiakhir.Source = "select * from tabel_nilaiakhir"
nilaiakhir.Open
End Sub

Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = nilaiakhir
With DataGrid1
End With
End Sub

Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = nilaiakhir
With DataGrid1
End With
Call edit_grid
End Sub

Sub kosong()
txtnik = ""
txtnamakaryawan = ""
txtjabatan = ""
txtbagian = ""
txtalamat = ""
Comjeniskelamin = ""
txtnik.SetFocus
End Sub
