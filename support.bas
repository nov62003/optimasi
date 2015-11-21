Attribute VB_Name = "support"
Option Explicit
Dim Digit(30) As Double             'atur variabel array untuk nilai integer
Dim BinaryDigit(30) As Integer      'atur variabel array untuk pengali biner
Public FinalBinNum As String        'variabel yang mengembalikan nilai biner
Public DecimalNumber As Double      'variabel yang mengembalikan nilai decimal

Public Function GetBinaryNumber(ByVal Number As Double, ByVal JumDigit)
Dim i As Integer, Locater As Integer, DecimalCheck As Integer

On Error GoTo ERROR_ROUTINE

For i = 0 To 30
    BinaryDigit(i) = 0                          'atur digit-digit biner ke nol
Next i

i = 0

Digit(0) = 1                                    'atur digit pengali pertama ke nilai 1

For i = 1 To 30                                 'secara berturut-turut setiap digit dikali dengan 2
    Digit(i) = Digit(i - 1) * 2
Next i


FinalBinNum = 0                                 'mengembalikan nilai variabel FinalBinNum ke 0

If Number > 1073741824 Then                     'cek apakah nilai desimal yang diinputkan tidak terlalu besar
    MsgBox "This Number is too Large"
    Exit Function
End If

DecimalCheck = InStr(1, Number, ".", vbTextCompare)     'cek nilai desimal apakah berupa bilangan pecahan
If DecimalCheck > 0 Then                        'jika ya maka tidak bisa dijalankan
    MsgBox "please no decimals"
    Exit Function
End If

If Number < 0 Then                              'cek apakah Number berupa bilangan negatif
    MsgBox "please no negative numbers"
    Exit Function
End If

Do While Number > 0                             'kalkulasi perulangan biner
    Select Case Number                          'langkah untuk merubah bilangan desimal
        Case Is >= Digit(30)
           Number = Number - Digit(30)
           BinaryDigit(0) = 1
           
        Case Is >= Digit(29)
           Number = Number - Digit(29)
           BinaryDigit(1) = 1
           
        Case Is >= Digit(28)
           Number = Number - Digit(28)
           BinaryDigit(2) = 1
           
        Case Is >= Digit(27)
           Number = Number - Digit(27)
           BinaryDigit(3) = 1
           
        Case Is >= Digit(26)
           Number = Number - Digit(26)
           BinaryDigit(4) = 1
           
        Case Is >= Digit(25)
           Number = Number - Digit(25)
           BinaryDigit(5) = 1
           
        Case Is >= Digit(24)
           Number = Number - Digit(24)
           BinaryDigit(6) = 1
           
        Case Is >= Digit(23)
           Number = Number - Digit(23)
           BinaryDigit(7) = 1
           
        Case Is >= Digit(22)
           Number = Number - Digit(22)
           BinaryDigit(8) = 1
           
        Case Is >= Digit(21)
           Number = Number - Digit(21)
           BinaryDigit(9) = 1
           
        Case Is >= Digit(20)
           Number = Number - Digit(20)
           BinaryDigit(10) = 1
           
        Case Is >= Digit(19)
           Number = Number - Digit(19)
           BinaryDigit(11) = 1
           
        Case Is >= Digit(18)
           Number = Number - Digit(18)
           BinaryDigit(12) = 1
           
        Case Is >= Digit(17)
           Number = Number - Digit(17)
           BinaryDigit(13) = 1
           
        Case Is >= Digit(16)
           Number = Number - Digit(16)
           BinaryDigit(14) = 1
           
        Case Is >= Digit(15)
           Number = Number - Digit(15)
           BinaryDigit(15) = 1
           
        Case Is >= Digit(14)
           Number = Number - Digit(14)
           BinaryDigit(16) = 1
           
        Case Is >= Digit(13)
           Number = Number - Digit(13)
           BinaryDigit(17) = 1
           
        Case Is >= Digit(12)
           Number = Number - Digit(12)
           BinaryDigit(18) = 1
           
        Case Is >= Digit(11)
           Number = Number - Digit(11)
           BinaryDigit(19) = 1
           
        Case Is >= Digit(10)
           Number = Number - Digit(10)
           BinaryDigit(20) = 1
           
        Case Is >= Digit(9)
           Number = Number - Digit(9)
           BinaryDigit(21) = 1
           
        Case Is >= Digit(8)
           Number = Number - Digit(8)
           BinaryDigit(22) = 1
           
        Case Is >= Digit(7)
           Number = Number - Digit(7)
           BinaryDigit(23) = 1
           
        Case Is >= Digit(6)
           Number = Number - Digit(6)
           BinaryDigit(24) = 1
           
        Case Is >= Digit(5)
           Number = Number - Digit(5)
           BinaryDigit(25) = 1
           
        Case Is >= Digit(4)
           Number = Number - Digit(4)
           BinaryDigit(26) = 1
           
        Case Is >= Digit(3)
           Number = Number - Digit(3)
           BinaryDigit(27) = 1
           
        Case Is >= Digit(2)
           Number = Number - Digit(2)
           BinaryDigit(28) = 1
           
        Case Is >= Digit(1)
           Number = Number - Digit(1)
           BinaryDigit(29) = 1
           
        Case Is = Digit(0)
           Number = Number - Digit(0)
           BinaryDigit(30) = 1
    End Select
Loop


i = 0                                                   'reset counter

For i = 0 To 30                                         'letakkan perulangan biner secara bersamaan
    FinalBinNum = FinalBinNum & BinaryDigit(i)
Next i

'Locater = InStr(1, FinalBinNum, 1, vbTextCompare)   'cari angka 1 dalam biner
                                                    
                                                    
'FinalBinNum = Mid(FinalBinNum, Locater, 30)         'remove leading zeros

If JumDigit = 6 Then
    FinalBinNum = Right(FinalBinNum, 6)
Else
    FinalBinNum = Right(FinalBinNum, 4)
End If

'/////////////////////////////////////////////////////////////////////////////////////////
EXIT_ROUTINE:
    Exit Function
ERROR_ROUTINE:
    MsgBox Err.Description & Err.Number
    Resume EXIT_ROUTINE
    
End Function

Public Function GetDecimalNumber(ByVal BinaryNumber As String)
Dim i As Integer
Dim Locater As Integer
Dim StrLength As Integer

On Error GoTo ERROR_ROUTINE
'//////////////////////////////////////////////////////////////////////////////

Digit(0) = 1                                    'setup binary multipliers

For i = 1 To 30 Step -1
    Digit(i) = Digit(i - 1) * 2
Next i


StrLength = Len(BinaryNumber)                   'get length of binary string
BinaryNumber = StrReverse(BinaryNumber)         'reverse binary string

DecimalNumber = 0                               'reset decimal number

For i = 1 To StrLength
    Locater = Mid(BinaryNumber, i, 1)           'get digit from binary string
    DecimalNumber = DecimalNumber + (Val(Digit(i - 1)) * Locater)    'multiply digit by binary power
Next i                                                               'and add to total


'/////////////////////////////////////////////////////////////////////////////////////////
EXIT_ROUTINE:
    Exit Function
ERROR_ROUTINE:
    MsgBox Err.Description & Err.Number
    Resume EXIT_ROUTINE

End Function

