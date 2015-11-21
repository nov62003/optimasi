VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHasil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimasi Pendapatan Maksimal dengan menggunakan Algoritma Genetika"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7620
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid msHasil 
      Height          =   3150
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Individu Baru yang dibangkitkan secara acak"
      Top             =   120
      Width           =   7400
      _ExtentX        =   13044
      _ExtentY        =   5556
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kesimpulan :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.Maksimal = Pendapatan Maksimal Angkutan Umum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   3600
      Width           =   3795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Catatan : K = Kromosom, KT = Kecepatan Tempuh, JP = Jumlah Penumpang, BP = Banyaknya Putaran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   7320
   End
End
Attribute VB_Name = "frmHasil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    With msHasil
        .TextMatrix(0, 0) = "K"
        .ColAlignment(0) = 3
        .ColWidth(0) = 600
        .TextMatrix(0, 1) = "BINER"
        .ColAlignment(1) = 3
        .ColWidth(1) = 1400
        .TextMatrix(0, 2) = "KT"
        .ColAlignment(2) = 3
        .ColWidth(2) = 600
        .TextMatrix(0, 3) = "JP"
        .ColAlignment(3) = 3
        .ColWidth(3) = 600
        .TextMatrix(0, 4) = "BP"
        .ColAlignment(4) = 3
        .ColWidth(4) = 600
        .TextMatrix(0, 5) = "P. MAKSIMAL"
        .ColAlignment(5) = 3
        .ColWidth(5) = 1600
        .TextMatrix(0, 6) = "KETERANGAN"
        .ColAlignment(6) = 3
        .ColWidth(6) = 2000
        
        
        .Rows = 11
    End With
    
    Label3.Left = (Me.Width - Label3.Width) / 2
   
End Sub
