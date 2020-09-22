Attribute VB_Name = "Convert"
Option Explicit
Private Function Biner(Bilangan) As String
Dim Basis As Integer
Dim Hsltemp As Variant
Dim sisa As Variant
Dim HslBagi As Variant
    Hsltemp = ""
    sisa = ""
    Basis = 2
    Do
        Hsltemp = sisa & Hsltemp
        HslBagi = Bilangan \ Basis
        sisa = Bilangan Mod Basis
        Bilangan = HslBagi
        
    Loop Until HslBagi <= 1
    Biner = HslBagi & sisa & Hsltemp
    Biner = Right("0000000" & Biner, 7)
End Function

'Create 7 bit / 8 bit
'charhex(str,7)-->7 bit for receiving SMS
'charhex(str,8)-->8 bit for send
Public Function CharHex(ByVal Txt As String, ByVal bit As Integer)
    Dim i As Integer, bin As String, nbin As String, n As String
    Dim bil As Integer, sisa As Integer, lbin As Integer, nol As String
    bin = ""
    nbin = ""
 
    If bit = 7 Then
        For i = 1 To Len(Txt) Step 2
            n = Mid(Txt, i, 2)
            bin = HexToBin(n) & bin
        Next
        bil = Len(bin) \ bit
        sisa = Len(bin) Mod bit
        For i = 1 To (Len(bin) - sisa) Step bit
           ' MsgBox Chr$(HexToDec(BinToHex(Mid(bin, i + Sisa, bit))))
            nbin = Chr$(HexToDec(BinToHex(Mid(bin, i + sisa, bit)))) & nbin
        Next
    Else
        For i = 1 To Len(Txt)
            n = Mid(Txt, i, 1)
            bin = Biner(Asc(n)) & bin
        Next
        sisa = Len(bin) Mod bit
    
        If sisa > 0 Then
            For i = 1 To bit - sisa
                nol = nol & "0"
            Next
        End If
    
        bin = nol & bin
        bil = Len(bin) \ bit
        For i = 1 To bil
            nbin = nbin & BinToHex(Mid(bin, Len(bin) + 1 - bit * i, bit))
        Next
    End If
    CharHex = nbin
End Function

Public Function BinToHex(ByVal Biner As String) As String
    Dim bin As String, n As String, nil As String, i As Integer
    bin = ""
    Biner = Right("00000000" & Biner, 8)
    For i = 1 To 2
        bin = Mid(Biner, Len(Biner) + 1 - 4 * i, 4)
        Select Case bin
            Case "0000": n = "0"
            Case "0001": n = "1"
            Case "0010": n = "2"
            Case "0011": n = "3"
            Case "0100": n = "4"
            Case "0101": n = "5"
            Case "0110": n = "6"
            Case "0111": n = "7"
            Case "1000": n = "8"
            Case "1001": n = "9"
            Case "1010": n = "A"
            Case "1011": n = "B"
            Case "1100": n = "C"
            Case "1101": n = "D"
            Case "1110": n = "E"
            Case "1111": n = "F"
        End Select
        nil = n & nil
    Next
    BinToHex = nil
End Function
Public Function HexToBin(ByVal Biner As String) As String
    Dim bin As String, n As String, nil As String, i As Integer
    bin = ""
    For i = 1 To Len(Biner)
        bin = Mid(Biner, i, 1)
        Select Case bin
            Case "0": n = "0000"
            Case "1": n = "0001"
            Case "2": n = "0010"
            Case "3": n = "0011"
            Case "4": n = "0100"
            Case "5": n = "0101"
            Case "6": n = "0110"
            Case "7": n = "0111"
            Case "8": n = "1000"
            Case "9": n = "1001"
            Case "A": n = "1010"
            Case "B": n = "1011"
            Case "C": n = "1100"
            Case "D": n = "1101"
            Case "E": n = "1110"
            Case "F": n = "1111"
        End Select
        nil = nil & n
    Next
    HexToBin = nil
End Function
Public Function ConvToChar(ByVal hx As String) As String
    Dim i As Integer, tx As String
    For i = 1 To Len(hx) Step 2
        tx = tx & Chr(HexToDec(Mid(hx, i, 2)))
    Next
    ConvToChar = tx
End Function
Public Function HexToDec(ByVal x As String) As Integer
    Dim m As String, i As Byte, nil As Integer, n As Integer
    For i = 1 To 2
        m = Mid(x, i, 1)
        Select Case UCase(m)
            Case "A": n = 10
            Case "B": n = 11
            Case "C": n = 12
            Case "D": n = 13
            Case "E": n = 14
            Case "F": n = 15
            Case Else: n = CInt(m)
        End Select
        If i = 1 Then
            nil = n * 16
        Else
            nil = nil + n
        End If
    Next
    HexToDec = nil
End Function
Public Function DecToHex(ByVal x As Integer) As String
    Dim nil As String
    nil = Hex(x)
    If Len(nil) = 1 Then
        nil = "0" & nil
    End If
    DecToHex = nil
End Function
'Reverse number
Public Function RevNum(ByVal numb As String) As String
    Dim s As Integer, ma As String, b As String, a As String
    Dim ta As String
     s = 1
     ma = ""
     While (s <= Len(numb))
       ta = Mid(numb, s, 2)
       a = Mid(ta, 1, 1)
       b = Mid(ta, 2, 1)
       If b = "" Then b = "F"
       ma = ma & b & a
       s = s + 2
     Wend
     RevNum = ma
End Function
'Delay time in second
Sub Tunda(ByVal dtk As Single)
    Dim awal As Variant
    awal = Timer
    Do While Timer < awal + dtk
      DoEvents
    Loop
End Sub

