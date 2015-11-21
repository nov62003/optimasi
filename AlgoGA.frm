VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimasi Pendapatan Maksimal dengan menggunakan Algoritma Genetika"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Proses Genetika"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      TabIndex        =   6
      Top             =   7680
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   12120
      TabIndex        =   20
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdMut 
      Caption         =   "Mutasi"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9240
      TabIndex        =   5
      ToolTipText     =   "Proses pembangkitan individu baru dan mengevaluasi kelayakan individu"
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton cmdCross 
      Caption         =   "Crossover"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      ToolTipText     =   "Proses pembangkitan individu baru dan mengevaluasi kelayakan individu"
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hasil Seleksi dengan metode Roulette Whell :"
      Height          =   3495
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   8950
      Begin MSFlexGridLib.MSFlexGrid msSeleksi 
         Height          =   2655
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Individu Baru yang dibangkitkan secara acak"
         Top             =   720
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollBars      =   0
         Appearance      =   0
      End
      Begin VB.Label lblTF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fitness :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdProses 
      Caption         =   "Proses Seleksi"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      ToolTipText     =   "Proses pembangkitan individu baru dan mengevaluasi kelayakan individu"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   4560
      TabIndex        =   11
      Top             =   1230
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid msGen 
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Individu Baru yang dibangkitkan secara acak"
      Top             =   1230
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   0
      Appearance      =   0
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "Generate Kromosom Awal"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Proses pembangkitan individu baru dan mengevaluasi kelayakan individu"
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   350
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "2000"
      ToolTipText     =   "Tarif Angkutan Umum untuk per orang yaitu 2000 / orang"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   350
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "5000"
      ToolTipText     =   "Jarak Tempuh Angkutan Umum dalam satuan meter"
      Top             =   120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid msBiner 
      Height          =   2655
      Left            =   2520
      TabIndex        =   13
      ToolTipText     =   "Individu Baru yang dibangkitkan secara acak"
      Top             =   1230
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid msCross 
      Height          =   2655
      Left            =   6960
      TabIndex        =   18
      ToolTipText     =   "Individu Baru yang dibangkitkan secara acak"
      Top             =   120
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollBars      =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid msMutasi 
      Height          =   2655
      Left            =   9180
      TabIndex        =   19
      ToolTipText     =   "Individu Baru yang dibangkitkan secara acak"
      Top             =   2880
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   0
      Appearance      =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kromosom Biner :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   14
      Top             =   960
      Width           =   1485
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Evaluasi Kromosom :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4560
      TabIndex        =   12
      Top             =   960
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kromosom Awal :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tarif Angkutan :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   558
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jarak Tempuh :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   198
      Width           =   1335
   End
End
Attribute VB_Name = "frmProses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim NPop As Byte, BP As Long, Ef() As Long, Iterasi As Long, KInt1() As Long
Dim i As Integer, j As Integer, k() As Long, V As Long, KBiner() As Long, KInt2() As Long
Dim teksLayak() As String, IPop As Integer, Tf As Long, P() As Double
Dim c() As Double, Range() As Double, KSeleksi() As Long, jebak As Byte
Dim RC1() As Double, R1() As Double, KCross() As Long, Kounter As Byte
Dim KCrossTemp() As Long, KTemp() As Long, RMutasi() As Double, KMutasi() As Long
Dim IndLayak() As Long

Sub evaluasiIndividu(ByVal j, ByVal ite) 'Prosedur untuk evaluasi kromosom
    ReDim Preserve Ef(NPop) As Long, teksLayak(NPop) As String
    
    Select Case ite
    Case 0
    V = k(j, 1) * 1000
    BP = Int(60 / ((Val(txtData(0).Text) / V) * 60))
    Ef(j) = BP * k(j, 2) * txtData(1)
    If Ef(j) >= 100000 Then
        teksLayak(j) = "LAYAK"
        List1.AddItem BP & "  | " & Ef(j) & " | " & teksLayak(j)
        cmdGen.Enabled = False
        cmdProses.Enabled = False
    Else
        teksLayak(j) = "T. LAYAK"
        If teksLayak(j) = "T. LAYAK" Then
            generateIndividu
        End If
    End If
    Case Is > 0
        Call Bin2Int

        List1.Clear
        'Me.Caption = Iterasi
        For i = 1 To NPop
            If k(i, 1) > 60 Or k(i, 1) < 40 Then
                k(i, 1) = Int(Rnd(1) * ((60 - 40) + 1) + 40) 'KT
                msGen.TextMatrix(i, 1) = k(i, 1)
            End If
            
            If k(i, 2) > 12 Or k(i, 2) = 0 Then
                k(i, 2) = Int(Rnd(1) * ((12 - 1) + 1) + 1)   'JP
                msGen.TextMatrix(i, 2) = k(i, 2)
            End If
        Next i
        
        Call Int2Bin
    End Select
End Sub

Sub IndividuLayak(ByVal IL) 'Prosedur untuk kromosom Optimal
    ReDim Preserve IndLayak(NPop) As Long, teksLayak(NPop) As String
    
    V = k(IL, 1) * 1000
    BP = Int(60 / ((Val(txtData(0).Text) / V) * 60))
    IndLayak(IL) = BP * k(IL, 2) * txtData(1)
    If IndLayak(IL) >= 100000 Then
        teksLayak(IL) = "Optimal"
        List1.AddItem IL & " | " & BP & "  | " & IndLayak(IL) & " | " & teksLayak(IL)
        cmdGen.Enabled = False
        cmdProses.Enabled = False
        
        With frmHasil.msHasil
            .TextMatrix(IL, 0) = IL
            .TextMatrix(IL, 1) = KBiner(IL)
            .TextMatrix(IL, 2) = k(IL, 1)
            .TextMatrix(IL, 3) = k(IL, 2)
            .TextMatrix(IL, 4) = BP
            .TextMatrix(IL, 5) = IndLayak(IL)
            .TextMatrix(IL, 6) = teksLayak(IL)
        End With
        Kounter = Kounter + 1
    Else
        teksLayak(IL) = "T. Optimal"
        If teksLayak(IL) = "T. Optimal" Then
            'generateIndividu
        End If
    End If
    
End Sub

Sub generateIndividu()  'Prosedur untuk membangkitan kromosom
    Randomize
    ReDim k(NPop, 2) As Long
    List1.Clear
    
    For i = 1 To NPop
        k(i, 1) = Int(Rnd(1) * ((60 - 40) + 1) + 40) 'KT
        k(i, 2) = Int(Rnd(1) * ((12 - 1) + 1) + 1)   'JP
        
        msGen.TextMatrix(i, 0) = i
        msGen.TextMatrix(i, 1) = k(i, 1)
        msGen.TextMatrix(i, 2) = k(i, 2)
        Call evaluasiIndividu(i, Iterasi)
    Next i
    Call Int2Bin
End Sub

Sub Bin2Int()
    ReDim k(NPop, 2) As Long
    For i = 1 To NPop
        msGen.TextMatrix(i, 0) = i
        Call GetDecimalNumber(Left(KMutasi(i), 6))
        k(i, 1) = DecimalNumber
        msGen.TextMatrix(i, 1) = k(i, 1)
        Call GetDecimalNumber(Right(KMutasi(i), 4))
        k(i, 2) = DecimalNumber
        msGen.TextMatrix(i, 2) = k(i, 2)
    Next i
End Sub

Sub Int2Bin()
    ReDim KBiner(NPop) As Long
    For i = 1 To NPop
        msBiner.TextMatrix(i, 0) = i
        Call GetBinaryNumber(k(i, 1), 6)
        KBiner(i) = FinalBinNum
        Call GetBinaryNumber(k(i, 2), 4)
        KBiner(i) = KBiner(i) & FinalBinNum
        msBiner.TextMatrix(i, 1) = KBiner(i)
    Next i
End Sub

Sub seleksidenganRW() 'Prosedur untuk memilih kromosom dengan metode Roulette Whell
    ReDim P(NPop) As Double, c(NPop) As Double
    Tf = 0
    For i = 1 To NPop
        Tf = Tf + Ef(i)
    Next i
    lblTF.Caption = "Total Fitness : " & Tf
    
    For i = 1 To NPop
        P(i) = Ef(i) / Tf
        msSeleksi.TextMatrix(i, 0) = Round(P(i), 10)
    Next i
    
    c(1) = P(1)
    msSeleksi.TextMatrix(1, 1) = Round(c(1), 10)
    For i = 2 To NPop
        c(i) = c(i - 1) + P(i)
        msSeleksi.TextMatrix(i, 1) = Round(c(i), 10)
    Next i
    
    Randomize
    ReDim r(NPop) As Double
    
    For i = 1 To NPop
        r(i) = Round((Rnd(1) * ((10 - 1) + 1) + 1) / 12, 8) 'Nilai Acak
        msSeleksi.TextMatrix(i, 2) = Round(r(i), 10)
    Next i
    
    ReDim KSeleksi(NPop) As Long
    For i = 1 To NPop 'R
        For j = 1 To NPop 'C
            If Round(c(j), 10) > Round(r(i), 10) Then
                msSeleksi.TextMatrix(i, 3) = "Kromosom " & j & " terpilih karena C(" & j & ") > R(" & i & ")"
                KSeleksi(i) = KBiner(j)
                msSeleksi.TextMatrix(i, 4) = KSeleksi(i)
                Exit For
            End If
        Next j
    Next i
End Sub

Sub CrossOver()
    Dim PC As Single, temp(10) As Long, k As Byte
    PC = 0.5
    
    Dim PanjangKromosom As Byte, PK As Byte
    PanjangKromosom = Len(CStr(KSeleksi(1)))
    PK = PanjangKromosom - 1
    
    Randomize
    ReDim r(NPop) As Double, KCross(NPop) As Long, R1(NPop) As Double
    
    Kounter = 0
    For i = 1 To NPop
        msCross.TextMatrix(i, 0) = KBiner(i)
        msCross.TextMatrix(i, 1) = KSeleksi(i)
       
        r(i) = Round((Rnd(1) * ((10 - 1) + 1) + 1) / 12, 8) 'Nilai Acak
        msCross.TextMatrix(i, 2) = Round(r(i), 10)
       
        If PC > r(i) Then
            msCross.TextMatrix(i, 3) = "Crossing"
            Kounter = Kounter + 1
            KCross(i) = KSeleksi(i)
            temp(i) = i '+ 1
            R1(i) = Int(Rnd(1) * ((PK - 1) + 1) + 1)   'Nilai Acak
            'msCross.TextMatrix(i, 4) = temp(i)
        Else
            'KCross(i) = KSeleksi(i)
            msCross.TextMatrix(i, 3) = "No Crossing"
            'msCross.TextMatrix(i, 4) = ""
        End If
    Next i
    
    Dim n As Byte, tes, jum
    ReDim KCrossTemp(NPop) As Long, KTemp(NPop) As Long
    
    For k = 1 To 10
        KCrossTemp(k) = KCross(k)
        msCross.TextMatrix(k, 4) = KCrossTemp(k)
        
        If KCrossTemp(k) <> 0 Then
            msCross.TextMatrix(k, 4) = k & " " & KCrossTemp(k)
            jum = jum + 1
            If jum = 1 Then
                tes = KCrossTemp(k)
            End If
            
            If jum = Kounter Then
                KTemp(k) = Left(KCrossTemp(k), R1(k)) & Right(tes, 10 - R1(k))
            End If
            
            For n = k To 10
                If KCrossTemp(k) <> KCross(n) And KCross(n) <> 0 Then
                    R1(n) = Int(Rnd(1) * ((PK - 1) + 1) + 1)   'Nilai Acak
                    KTemp(k) = Left(KCrossTemp(k), R1(n)) & Right(KCross(n), 10 - R1(n))
                    Exit For
                End If
            Next n
        Else
            KTemp(k) = KSeleksi(k)
        End If
        msCross.TextMatrix(k, 4) = KTemp(k)
    Next k
End Sub

Sub Mutasi()
    Dim PM As Single, Cari, CariIf, Ganti As String, PDigit As Byte, CountJ
    PM = 0.1

    Randomize
    ReDim RMutasi(10, 10) As Double, KMutasi(NPop) As Long
    
    List2.Clear
    For i = 1 To NPop
        msMutasi.TextMatrix(i, 0) = KTemp(i)
        msMutasi.TextMatrix(i, 1) = "No. Mutasi"
        KMutasi(i) = KTemp(i)
        msMutasi.TextMatrix(i, 2) = KMutasi(i)
        For j = 1 To NPop
            CountJ = CountJ + 1
            
            RMutasi(i, j) = Round((Rnd(1) * ((10 - 1) + 1) + 1) / 20, 5) 'Nilai Acak
            If PM > RMutasi(i, j) Then
                Kounter = Kounter + 1
                Cari = Mid(CStr(KTemp(i)), j, 1)
                CariIf = IIf(Cari = "1", "0", "1")
                Ganti = Replace(CStr(KMutasi(i)), Cari, CariIf, j, 1, vbBinaryCompare)
                PDigit = Len(Ganti)
                KMutasi(i) = Left(CStr(KMutasi(i)), 10 - PDigit) & Ganti
                msMutasi.TextMatrix(i, 1) = "Mutasi"
                List2.AddItem i & " | " & j & " | " & Round(RMutasi(i, j), 3) & " | " & Cari & " | " & CountJ
                msMutasi.TextMatrix(i, 2) = KMutasi(i)
            End If
        Next j
    Next i
End Sub

Sub Kosong()
    For i = 1 To NPop
        'With msGen
         '   .TextMatrix(i, 0) = ""
         '   .TextMatrix(i, 1) = ""
         '   .TextMatrix(i, 2) = ""
        'End With
    
        'With msBiner
         '   .TextMatrix(i, 0) = ""
         '   .TextMatrix(i, 1) = ""
        'End With
        
        'jebak = 0
        If jebak = 0 Then
        With msSeleksi
            .TextMatrix(i, 0) = ""
            .TextMatrix(i, 1) = ""
            .TextMatrix(i, 2) = ""
            .TextMatrix(i, 3) = ""
            .TextMatrix(i, 4) = ""
        End With
        End If
        
        If jebak = 1 Then
        With msCross
            .TextMatrix(i, 0) = ""
            .TextMatrix(i, 1) = ""
            .TextMatrix(i, 2) = ""
            .TextMatrix(i, 3) = ""
            '.TextMatrix(i, 4) = ""
        End With
        End If
        
        If jebak = 2 Then
        With msMutasi
            .TextMatrix(i, 0) = ""
            .TextMatrix(i, 1) = ""
            .TextMatrix(i, 2) = ""
        End With
        End If
        
    Next i
End Sub

Private Sub cmdCross_Click()
    jebak = 1
    Call CrossOver
    
    For i = 1 To NPop
        If KTemp(i) = 0 Then
            Call CrossOver
        End If
    Next i
    
    If Kounter < 5 Then Call CrossOver
    Kounter = 0
    cmdCross.Enabled = False
    cmdMut.Enabled = True
    cmdMut.SetFocus
End Sub

Private Sub cmdGen_Click()
    generateIndividu
    cmdProses.Enabled = True
    cmdProses.SetFocus
End Sub

Private Sub cmdMut_Click()
    jebak = 2
    Call Kosong
    Call Mutasi
    Kounter = 0
    
    If Kounter < 10 Then Call Mutasi
    Kounter = 0
    
    Iterasi = Iterasi + 1
    cmdMut.Enabled = False
    Command1.Enabled = True
    Command1.SetFocus
End Sub

Private Sub cmdProses_Click()
    jebak = 0
    Call Kosong
    Call seleksidenganRW
    cmdProses.Enabled = False
    cmdCross.Enabled = True
    cmdCross.SetFocus
End Sub

Private Sub Command1_Click()
    Do While Iterasi <= 99
        Call evaluasiIndividu(0, Iterasi)
        Call cmdProses_Click
        Call cmdCross_Click
        Call cmdMut_Click
    Loop
    
    If Iterasi = 100 Then
       For i = 1 To NPop
            Call IndividuLayak(i)
        Next i
        Command1.Enabled = False
        cmdGen.Enabled = True
        cmdGen.SetFocus
        
        
        frmHasil.Label3.Caption = "Setelah dilakukan iterasi sebanyak " & Iterasi & " kali " & _
        Chr(13) & "maka ada " & Kounter & " Individu yang Optimal. " _
        & Chr(13) & "Seperti yang terlihat pada tabel di atas"
        frmHasil.Show 1
        
        Iterasi = 0
    End If
End Sub

Private Sub Form_Load()
    Iterasi = 0
    NPop = 10
    
    msGen.TextMatrix(0, 0) = "KROMOSOM"
    msGen.ColAlignment(0) = 3
    msGen.ColWidth(0) = 1060
    msGen.TextMatrix(0, 1) = "KT"
    msGen.ColAlignment(1) = 3
    msGen.ColWidth(1) = 600
    msGen.TextMatrix(0, 2) = "JP"
    msGen.ColAlignment(2) = 3
    msGen.ColWidth(2) = 600
    
    msGen.Rows = NPop + 1
    
    msBiner.TextMatrix(0, 0) = "K"
    msBiner.ColAlignment(0) = 3
    msBiner.ColWidth(0) = 505
    msBiner.TextMatrix(0, 1) = "-"
    msBiner.ColAlignment(1) = 3
    msBiner.ColWidth(1) = 1400
    
    msBiner.Rows = NPop + 1
    
    
    msSeleksi.TextMatrix(0, 0) = "Probabilitas"
    msSeleksi.ColAlignment(0) = 3
    msSeleksi.ColWidth(0) = 1400
    msSeleksi.TextMatrix(0, 1) = "Komulatif"
    msSeleksi.ColAlignment(1) = 3
    msSeleksi.ColWidth(1) = 1400
    msSeleksi.TextMatrix(0, 2) = "Nilai Acak"
    msSeleksi.ColAlignment(2) = 3
    msSeleksi.ColWidth(2) = 1400
    msSeleksi.TextMatrix(0, 3) = "Perbadingan C dan R"
    msSeleksi.ColAlignment(3) = 3
    msSeleksi.ColWidth(3) = 3100
    msSeleksi.TextMatrix(0, 4) = "K. Baru"
    msSeleksi.ColAlignment(4) = 3
    msSeleksi.ColWidth(4) = 1400
    
    msSeleksi.Rows = NPop + 1
    
    msCross.TextMatrix(0, 0) = "K->Awal"
    msCross.ColAlignment(0) = 3
    msCross.ColWidth(0) = 1400
    msCross.TextMatrix(0, 1) = "K-Hasil (Seleksi)"
    msCross.ColAlignment(1) = 3
    msCross.ColWidth(1) = 1400
    msCross.TextMatrix(0, 2) = "Nilai Acak"
    msCross.ColAlignment(2) = 3
    msCross.ColWidth(2) = 1400
    msCross.TextMatrix(0, 3) = "PC(0.5) > R"
    msCross.ColAlignment(3) = 3
    msCross.ColWidth(3) = 1400
    msCross.TextMatrix(0, 4) = "K-Hasil (Crossing)"
    msCross.ColAlignment(4) = 3
    msCross.ColWidth(4) = 1400
    
    msCross.Rows = NPop + 1
    
    msMutasi.TextMatrix(0, 0) = "K-Hasil (Crossing)"
    msMutasi.ColAlignment(0) = 3
    msMutasi.ColWidth(0) = 1400
    msMutasi.TextMatrix(0, 1) = "R"
    msMutasi.ColAlignment(1) = 3
    msMutasi.ColWidth(1) = 2000
    msMutasi.TextMatrix(0, 2) = "K-Hasil (Mutasi)"
    msMutasi.ColAlignment(2) = 3
    msMutasi.ColWidth(2) = 1400
    
    msMutasi.Rows = NPop + 1
End Sub

Private Sub txtData_GotFocus(Index As Integer)
    For i = 0 To txtData().Count - 1
        txtData(i).SelStart = 0
        txtData(i).SelLength = Len(txtData(i).Text)
    Next i
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        SendKeys vbTab
    Case 8        'Backspace
    Case 48 To 57 'angka 0-9
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtData_LostFocus(Index As Integer)
    If Index = 0 And txtData(0).Text < 2000 Then
        MsgBox "Jarak Tempuh tidak boleh kosong dan " _
                & Chr(13) & "Tidak boleh kurang dari 2000 meter", vbOKOnly, "-- Perhatian"
        txtData(0).Text = 0
        txtData(0).SetFocus
    End If
End Sub
