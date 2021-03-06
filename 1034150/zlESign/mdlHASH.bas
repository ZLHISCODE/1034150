Attribute VB_Name = "mdlHASH"
'hashֵ???㺯??
Private Type Word
B0 As Byte
B1 As Byte
B2 As Byte
B3 As Byte
End Type

Private Function AndW(w1 As Word, w2 As Word) As Word
AndW.B0 = w1.B0 And w2.B0
AndW.B1 = w1.B1 And w2.B1
AndW.B2 = w1.B2 And w2.B2
AndW.B3 = w1.B3 And w2.B3
End Function

Private Function OrW(w1 As Word, w2 As Word) As Word
OrW.B0 = w1.B0 Or w2.B0
OrW.B1 = w1.B1 Or w2.B1
OrW.B2 = w1.B2 Or w2.B2
OrW.B3 = w1.B3 Or w2.B3
End Function

Private Function XorW(w1 As Word, w2 As Word) As Word
XorW.B0 = w1.B0 Xor w2.B0
XorW.B1 = w1.B1 Xor w2.B1
XorW.B2 = w1.B2 Xor w2.B2
XorW.B3 = w1.B3 Xor w2.B3
End Function

Private Function NotW(w As Word) As Word
NotW.B0 = Not w.B0
NotW.B1 = Not w.B1
NotW.B2 = Not w.B2
NotW.B3 = Not w.B3
End Function

Private Function AddW(w1 As Word, w2 As Word) As Word
Dim i As Long, w As Word

i = CLng(w1.B3) + w2.B3
w.B3 = i Mod 256
i = CLng(w1.B2) + w2.B2 + (i \ 256)
w.B2 = i Mod 256
i = CLng(w1.B1) + w2.B1 + (i \ 256)
w.B1 = i Mod 256
i = CLng(w1.B0) + w2.B0 + (i \ 256)
w.B0 = i Mod 256

AddW = w
End Function

Private Function CircShiftLeftW(w As Word, n As Long) As Word
Dim d1 As Double, d2 As Double

d1 = WordToDouble(w)
d2 = d1
d1 = d1 * (2 ^ n)
d2 = d2 / (2 ^ (32 - n))
CircShiftLeftW = OrW(DoubleToWord(d1), DoubleToWord(d2))
End Function

Private Function WordToHex(w As Word) As String
WordToHex = Right$("0" & Hex$(w.B0), 2) & Right$("0" & Hex$(w.B1), 2) _
& Right$("0" & Hex$(w.B2), 2) & Right$("0" & Hex$(w.B3), 2)
End Function

Private Function HexToWord(H As String) As Word
HexToWord = DoubleToWord(Val("&H" & H & "#"))
End Function

Private Function DoubleToWord(n As Double) As Word
DoubleToWord.B0 = Int(DMod(n, 2 ^ 32) / (2 ^ 24))
DoubleToWord.B1 = Int(DMod(n, 2 ^ 24) / (2 ^ 16))
DoubleToWord.B2 = Int(DMod(n, 2 ^ 16) / (2 ^ 8))
DoubleToWord.B3 = Int(DMod(n, 2 ^ 8))
End Function

Private Function WordToDouble(w As Word) As Double
WordToDouble = (w.B0 * (2 ^ 24)) + (w.B1 * (2 ^ 16)) + (w.B2 * (2 ^ 8)) _
+ w.B3
End Function

Private Function DMod(value As Double, divisor As Double) As Double
DMod = value - (Int(value / divisor) * divisor)
If DMod < 0 Then DMod = DMod + divisor
End Function

Private Function F(t As Long, B As Word, C As Word, D As Word) As Word
Select Case t
Case Is <= 19
F = OrW(AndW(B, C), AndW(NotW(B), D))
Case Is <= 39
F = XorW(XorW(B, C), D)
Case Is <= 59
F = OrW(OrW(AndW(B, C), AndW(B, D)), AndW(C, D))
Case Else
F = XorW(XorW(B, C), D)
End Select
End Function
Public Function StringSHA1(inMessage As String) As String
' ?????ַ?????SHA1ժҪ
Dim inLen As Long
Dim inLenW As Word
Dim padMessage As String
Dim numBlocks As Long
Dim w(0 To 79) As Word
Dim blockText As String
Dim wordText As String
Dim i As Long, t As Long
Dim temp As Word
Dim K(0 To 3) As Word
Dim H0 As Word
Dim H1 As Word
Dim H2 As Word
Dim H3 As Word
Dim H4 As Word
Dim A As Word
Dim B As Word
Dim C As Word
Dim D As Word
Dim E As Word

inMessage = StrConv(inMessage, vbFromUnicode)

inLen = LenB(inMessage)
inLenW = DoubleToWord(CDbl(inLen) * 8)

padMessage = inMessage & ChrB(128) _
& StrConv(String((128 - (inLen Mod 64) - 9) Mod 64 + 4, Chr(0)), 128) _
& ChrB(inLenW.B0) & ChrB(inLenW.B1) & ChrB(inLenW.B2) & ChrB(inLenW.B3)

numBlocks = LenB(padMessage) / 64

' initialize constants
K(0) = HexToWord("5A827999")
K(1) = HexToWord("6ED9EBA1")
K(2) = HexToWord("8F1BBCDC")
K(3) = HexToWord("CA62C1D6")

' initialize 160-bit (5 words) buffer
H0 = HexToWord("67452301")
H1 = HexToWord("EFCDAB89")
H2 = HexToWord("98BADCFE")
H3 = HexToWord("10325476")
H4 = HexToWord("C3D2E1F0")

' each 512 byte message block consists of 16 words (W) but W is expanded
For i = 0 To numBlocks - 1
blockText = MidB$(padMessage, (i * 64) + 1, 64)
' initialize a message block
For t = 0 To 15
wordText = MidB$(blockText, (t * 4) + 1, 4)
w(t).B0 = AscB(MidB$(wordText, 1, 1))
w(t).B1 = AscB(MidB$(wordText, 2, 1))
w(t).B2 = AscB(MidB$(wordText, 3, 1))
w(t).B3 = AscB(MidB$(wordText, 4, 1))
Next

' create extra words from the message block
For t = 16 To 79
' W(t) = S^1 (W(t-3) XOR W(t-8) XOR W(t-14) XOR W(t-16))
w(t) = CircShiftLeftW(XorW(XorW(XorW(w(t - 3), w(t - 8)), _
w(t - 14)), w(t - 16)), 1)
Next

' make initial assignments to the buffer
A = H0
B = H1
C = H2
D = H3
E = H4

' process the block
For t = 0 To 79
temp = AddW(AddW(AddW(AddW(CircShiftLeftW(A, 5), _
F(t, B, C, D)), E), w(t)), K(t \ 20))
E = D
D = C
C = CircShiftLeftW(B, 30)
B = A
A = temp
Next

H0 = AddW(H0, A)
H1 = AddW(H1, B)
H2 = AddW(H2, C)
H3 = AddW(H3, D)
H4 = AddW(H4, E)
Next

StringSHA1 = WordToHex(H0) & WordToHex(H1) & WordToHex(H2) _
& WordToHex(H3) & WordToHex(H4)

End Function

Public Function SHA1(inMessage() As Byte) As String
' ?????ֽ???????SHA1ժҪ
Dim inLen As Long
Dim inLenW As Word
Dim numBlocks As Long
Dim w(0 To 79) As Word
Dim blockText As String
Dim wordText As String
Dim t As Long
Dim temp As Word
Dim K(0 To 3) As Word
Dim H0 As Word
Dim H1 As Word
Dim H2 As Word
Dim H3 As Word
Dim H4 As Word
Dim A As Word
Dim B As Word
Dim C As Word
Dim D As Word
Dim E As Word
Dim i As Long
Dim lngPos As Long
Dim lngPadMessageLen As Long
Dim padMessage() As Byte

inLen = UBound(inMessage) + 1
inLenW = DoubleToWord(CDbl(inLen) * 8)

lngPadMessageLen = inLen + 1 + (128 - (inLen Mod 64) - 9) Mod 64 + 8
ReDim padMessage(lngPadMessageLen - 1) As Byte
For i = 0 To inLen - 1
padMessage(i) = inMessage(i)
Next i
padMessage(inLen) = 128
padMessage(lngPadMessageLen - 4) = inLenW.B0
padMessage(lngPadMessageLen - 3) = inLenW.B1
padMessage(lngPadMessageLen - 2) = inLenW.B2
padMessage(lngPadMessageLen - 1) = inLenW.B3

numBlocks = lngPadMessageLen / 64

' initialize constants
K(0) = HexToWord("5A827999")
K(1) = HexToWord("6ED9EBA1")
K(2) = HexToWord("8F1BBCDC")
K(3) = HexToWord("CA62C1D6")

' initialize 160-bit (5 words) buffer
H0 = HexToWord("67452301")
H1 = HexToWord("EFCDAB89")
H2 = HexToWord("98BADCFE")
H3 = HexToWord("10325476")
H4 = HexToWord("C3D2E1F0")

' each 512 byte message block consists of 16 words (W) but W is expanded
' to 80 words
For i = 0 To numBlocks - 1
' initialize a message block
For t = 0 To 15
w(t).B0 = padMessage(lngPos)
w(t).B1 = padMessage(lngPos + 1)
w(t).B2 = padMessage(lngPos + 2)
w(t).B3 = padMessage(lngPos + 3)
lngPos = lngPos + 4
Next

' create extra words from the message block
For t = 16 To 79
' W(t) = S^1 (W(t-3) XOR W(t-8) XOR W(t-14) XOR W(t-16))
w(t) = CircShiftLeftW(XorW(XorW(XorW(w(t - 3), w(t - 8)), _
w(t - 14)), w(t - 16)), 1)
Next

' make initial assignments to the buffer
A = H0
B = H1
C = H2
D = H3
E = H4

' process the block
For t = 0 To 79
temp = AddW(AddW(AddW(AddW(CircShiftLeftW(A, 5), _
F(t, B, C, D)), E), w(t)), K(t \ 20))
E = D
D = C
C = CircShiftLeftW(B, 30)
B = A
A = temp
Next

H0 = AddW(H0, A)
H1 = AddW(H1, B)
H2 = AddW(H2, C)
H3 = AddW(H3, D)
H4 = AddW(H4, E)
Next

SHA1 = WordToHex(H0) & WordToHex(H1) & WordToHex(H2) _
& WordToHex(H3) & WordToHex(H4)

End Function

Public Function FileSHA1(strFilename As String) As String
' ?????ļ???SHA1ժҪ
Dim lngFileNo As Long
Dim bytData() As Byte

If Dir(strFilename) = "" Then
GoTo PROC_EXIT
End If

lngFileNo = FreeFile

On Error GoTo PROC_ERR

' ?????ļ?
Open strFilename For Binary As lngFileNo

' ??ȡ?ļ?????
ReDim bytData(LOF(lngFileNo) - 1) As Byte
Get #lngFileNo, 1, bytData

' ?ر??ļ?
Close lngFileNo

' ?????ļ???SHA1ժҪ
FileSHA1 = SHA1(bytData)

PROC_EXIT:
Erase bytData
Exit Function

PROC_ERR:
Close
GoTo PROC_EXIT

End Function





