VERSION 5.00
Begin VB.Form frmLisGraph 
   BorderStyle     =   0  'None
   Caption         =   "Graph"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2500
      Left            =   0
      ScaleHeight     =   2445
      ScaleWidth      =   2445
      TabIndex        =   0
      Top             =   -15
      Width           =   2500
   End
End
Attribute VB_Name = "frmLisGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Diff
    'Diff AL 作图 要用
    NoL = 0: NoN: NoE: ln: RN: LL: AL: LMU: LMD: LMN: MN: RM: NL: NE: RMN: FNE: FMN: FLN
End Enum

Public Function Draw_HMX_DF1(ByVal str_Line As String, ByVal str_Data As String) As String
    '画DF1图
    '入参
    '   str_Line:画线的坐标，用，分隔，一共5个
    '   str_Data:散点图数据
    '出参
    '   绘图成功，返回图形文件名。
    
    Picture1.Scale (0, 0)-(256, 256)
    Picture1.BackColor = vbWhite
    Dim x As Integer, Y As Integer
    Dim i_L1 As Integer, i_L2 As Integer, i_L3 As Integer, i_L4 As Integer, i_L5 As Integer
    Dim str_Img As String
    str_Img = str_Data
    i_L1 = Split(str_Line, ",")(0)
    i_L2 = Split(str_Line, ",")(1)
    i_L3 = Split(str_Line, ",")(2)
    i_L4 = Split(str_Line, ",")(3)
    i_L5 = Split(str_Line, ",")(4)
    Picture1.Line (i_L2, 0)-(i_L2, 256 - i_L4), vbBlack, BF
    Picture1.Line (i_L1, 0)-(i_L1, 256 - i_L5), vbBlack, BF
    Picture1.Line (0, 256 - i_L3)-(i_L1, 256 - i_L3), vbBlack, BF
    Picture1.Line (i_L1, 256 - i_L4)-(256, 256 - i_L4), vbBlack, BF
    Picture1.Line (0, 256 - i_L5)-(256, 256 - i_L5), vbBlack, BF
    
    
    For x = 1 To 64
        For Y = 64 To 1 Step -1
            If Mid(str_Img, 1, 1) <> "0" Then
                Call DrawPoint(Mid(str_Img, 1, 1), x, Y)
            End If
            str_Img = Mid(str_Img, 2)
        Next
    Next
    If Dir(App.Path & "\DF1_Tmp.Bmp") <> "" Then
        Kill App.Path & "\DF1_Tmp.Bmp"
    End If
    Draw_HMX_DF1 = App.Path & "\DF1_Tmp.Bmp"
    SavePicture Picture1.Image, Draw_HMX_DF1
    
End Function

Public Function Draw_HMX_DF2(ByVal str_Data As String) As String
    '画DF2图
    '入参
    '   str_Data:散点图数据
    '出参
    '   绘图成功，返回图形文件名。
    
    
    Picture1.Scale (0, 0)-(256, 256)
    Picture1.BackColor = vbWhite
    Dim x As Integer, Y As Integer
    Dim str_Line As String
    
    str_Line = str_Data
    For x = 1 To 64
        For Y = 64 To 1 Step -1
            If Mid(str_Line, 1, 1) <> "0" Then
                Call DrawPoint(Mid(str_Line, 1, 1), x, Y)
            End If
            str_Line = Mid(str_Line, 2)
        Next
    Next
    If Dir(App.Path & "\DF2_Tmp.Bmp") <> "" Then
        Kill App.Path & "\DF2_Tmp.Bmp"
    End If
    Draw_HMX_DF2 = App.Path & "\DF2_Tmp.Bmp"
    SavePicture Picture1.Image, Draw_HMX_DF2
    
End Function

Private Function DrawPoint(ByVal str_in As String, ByVal x As Integer, ByVal Y As Integer)
    '画点函数
    Dim lng_x As Long, lng_y As Long
    Select Case str_in
    Case "1"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If lng_x = 1 And lng_y = 1 Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "2"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x = 1 And lng_y <= 2) Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "3"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x = 1 And lng_y <= 3) Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "4"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x = 1 And lng_y <= 4) Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    
    Case "5"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x <= 2 And lng_y = 1) Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "6"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x <= 2 And lng_y <= 2) Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "7"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x <= 2 And lng_y <= 3) Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "8"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x <= 2 And lng_y >= 2 And lng_y <= 4) Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "9"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x >= 2 And lng_x <= 3 And lng_y = 1) Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "A"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x >= 2 And lng_x <= 3 And lng_y >= 2 And lng_y <= 2) Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    ', "C", "D", "E", "F"
    '问题：29348
    '修改颜色
    Case "B"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x >= 2 And lng_x <= 3 And lng_y >= 2 And lng_y <= 3) Then
                    Picture1.PSet ((x - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbYellow
                End If
            Next
        Next
    
    Case "C"

    Case "D"
    Case "E"
    Case "F"
    End Select
End Function

Private Sub Form_Load()
   Me.Hide
End Sub

Public Function Draw_Bc5500(ByVal str_bin As String, ByVal strFileName As String, ByVal strColor) As Boolean
    
    Dim lngX As Long, lngY As Long, lngV As Long, i As Integer
    Dim strByte As String, strV As String
    Dim strData As String, lngCount As Long, lngDawPoint As Long
    Dim strColorPoint As String, lngPointColor As Long
    Dim strInColor As String, lngMaxType As Long
    
    strData = str_bin
    strInColor = strColor
    Picture1.Scale (0, 0)-(256, 256)
    Picture1.BackColor = vbWhite
    
    Picture1.Line (0, 0)-(0, 255)
    Picture1.Line (0, 255)-(255, 255)
    
    Do While Len(strInColor) > 0
        For i = 0 To 1
            strByte = Mid(Left(strInColor, 3), 2)
            strInColor = Mid(strInColor, 4)
            If i = 0 Then
                strV = strByte
            Else
                strColorPoint = strColorPoint & "," & Val("&H" & strV & strByte)
            End If
        Next
    Loop
    
    If strColorPoint <> "" Then
        strColorPoint = Mid(strColorPoint, 2)
        lngMaxType = UBound(Split(strColorPoint, ","))
        
    End If

     
    Do While Len(strData) > 0

        
        strByte = Mid(Left(strData, 3), 2)
        lngX = Val("&H" & strByte)
        strData = Mid(strData, 4)

        strByte = Mid(Left(strData, 3), 2)
        lngY = Val("&H" & strByte)

        strData = Mid(strData, 10)
        
        lngCount = lngCount + 1

        If lngCount > lngDawPoint Then
            '换色
            If InStr(strFileName, "BASO") > 0 Then
                If UBound(Split(strColorPoint, ",")) = lngMaxType Then
                    lngPointColor = vbBlue
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 1 Then
                    lngPointColor = vbGreen
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 2 Then
                    lngPointColor = vbCyan
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 3 Then
                    lngPointColor = vbRed
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 4 Then
                    lngPointColor = vbMagenta
                End If
            Else
                If UBound(Split(strColorPoint, ",")) = lngMaxType Then
                    lngPointColor = vbBlue
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 1 Then
                    lngPointColor = vbGreen
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 2 Then
                    lngPointColor = vbMagenta
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 3 Then
                    lngPointColor = vbRed
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 4 Then
                    lngPointColor = vbCyan
                End If
            
            End If

            If strColorPoint <> "" Then
                If InStr(strColorPoint, ",") > 0 Then
                    lngDawPoint = lngDawPoint + Mid(strColorPoint, 1, InStr(strColorPoint, ",") - 1)
                    strColorPoint = Mid(strColorPoint, InStr(strColorPoint, ",") + 1)
                Else
                    lngDawPoint = lngDawPoint + strColorPoint
                    strColorPoint = ""
                End If
            End If
        End If
        Picture1.PSet (lngX, 256 - lngY), RGB(lngPointColor Mod 256, lngPointColor / 256 Mod 256, lngPointColor / 256 / 256)
    Loop
    
    If Dir(strFileName) <> "" Then
        Kill strFileName
    End If
    SavePicture Picture1.Image, strFileName
    Draw_Bc5500 = True
End Function

Public Function DrawP60(ByVal str_in As String, ByVal strFileName As String, ByVal strFlag As String) As Boolean
    Dim str_Line As String, x As Integer, Y As Integer
    Picture1.Scale (0, 0)-(128, 128)
    Picture1.BackColor = vbWhite
    str_Line = str_in
    For Y = 0 To 127
        For x = 0 To 127
            If Val(Replace(Mid(str_Line, 1, 3), ",", "")) <> 0 Then
                Picture1.PSet (x, Y), vbBlack
            End If
            str_Line = Mid(str_Line, 4)
            If str_Line = "" Then Exit For
        Next
        If str_Line = "" Then Exit For
    Next
    '---
    Dim strA As String, intloop As Integer
    Dim intA(18) As Integer
    Dim X1 As Currency, X2 As Currency, Y1 As Currency, Y2 As Currency
    strA = strFlag ' "022,025,048,035,118,030,068,078,090,070,090,118,029,071,051,002,002,002"
                     '022 025 048 035 118 030 068 078 090 070 090 118 029 063 038 002 002 002
    For intloop = LBound(Split(strA, ",")) To UBound(Split(strA, ","))
        intA(intloop) = Split(strA, ",")(intloop)
    Next
    
    '第一块
    X1 = intA(Diff.NoL): Y1 = 127: X2 = intA(Diff.NoL): Y2 = 127 - intA(Diff.NL)
    Picture1.Line (X1, Y1)-(X2, Y2), vbRed
    
    X1 = intA(Diff.NoL): Y1 = 127 - intA(Diff.NL): X2 = intA(Diff.LMU): Y2 = 127 - intA(Diff.NL)
    Picture1.Line (X1, Y1)-(X2, Y2), vbRed
    
    X1 = intA(Diff.LMU): Y1 = 127 - intA(Diff.NL): X2 = intA(Diff.LMD): Y2 = 127
    Picture1.Line (X1, Y1)-(X2, Y2), vbRed
    
    X1 = intA(Diff.LL): Y1 = 127: X2 = intA(Diff.LL): Y2 = 127 - intA(Diff.NL)
    Picture1.Line (X1, Y1)-(X2, Y2), vbMagenta
    
    X1 = intA(Diff.AL): Y1 = 127: X2 = intA(Diff.AL): Y2 = 127 - intA(Diff.NL)
    Picture1.Line (X1, Y1)-(X2, Y2), vbMagenta
    '第二块
    Picture1.Line (intA(Diff.RM), 127)-(intA(Diff.RM), 127 - intA(Diff.RMN)), vbRed
    Picture1.Line (intA(Diff.LMN), 127 - intA(Diff.NL))-(intA(Diff.MN), 127 - intA(Diff.RMN)), vbRed
    Picture1.Line (intA(Diff.MN), 127 - intA(Diff.RMN))-(127, 127 - intA(Diff.RMN)), vbRed
    
    '第三块
    Picture1.Line (intA(Diff.NoN), 127 - intA(Diff.NL))-(intA(Diff.NoN), 127 - intA(Diff.NE)), vbRed
    Picture1.Line (intA(Diff.NoN), 127 - intA(Diff.NE))-(127, 127 - intA(Diff.NE)), vbRed
    Picture1.Line (intA(Diff.ln), 127 - intA(Diff.NL))-(intA(Diff.ln), 127 - intA(Diff.NE)), vbMagenta
    Picture1.Line (intA(Diff.RN), 127 - intA(Diff.RMN))-(intA(Diff.RN), 127 - intA(Diff.NE)), vbRed
    
    '第四块
    Picture1.Line (intA(Diff.NoE), 127 - intA(Diff.NE))-(intA(Diff.NoE), 127 - 127), vbRed
    '虚线
    Picture1.Line (intA(Diff.NoE), 127 - (intA(Diff.NE) + intA(Diff.FNE)))-(127, 127 - (intA(Diff.NE) + intA(Diff.FNE))), vbBlue
    Picture1.Line (intA(Diff.NoN), 127 - (intA(Diff.NL) + intA(Diff.FLN)))-(intA(Diff.LMN), 127 - (intA(Diff.NL) + intA(Diff.FLN))), vbBlue
    Picture1.Line (intA(Diff.LMN), 127 - (intA(Diff.NL) + intA(Diff.FLN)))-(intA(Diff.MN), 127 - (intA(Diff.RMN) + intA(Diff.FMN))), vbBlue

    SavePicture Picture1.Image, strFileName
    DrawP60 = True
End Function


Public Function DrawDiff5AL(ByVal strCode As String, ByVal strFileName As String, ByVal strFlag As String) As Boolean
    Dim x As Integer, Y As Integer, str_in As String
    Dim strBit As String
    
    str_in = strCode
    
    Picture1.Scale (0, 0)-(128, 128)
    Picture1.BackColor = vbWhite
    
    For Y = 0 To 127
        For x = 0 To 127
            strBit = Left(str_in, 1)
            If Val(strBit) <> 0 Then
                Picture1.PSet (x, Y), vbBlack
            End If
            str_in = Mid(str_in, 2)
            If str_in = "" Then Exit For
        Next
        If str_in = "" Then Exit For
    Next
    
    '---
    Dim strA As String, intloop As Integer
    Dim intA(18) As Integer
    Dim X1 As Currency, X2 As Currency, Y1 As Currency, Y2 As Currency
    strA = strFlag ' "022,025,048,035,118,030,068,078,090,070,090,118,029,071,051,002,002,002"
    For intloop = LBound(Split(strA, ",")) To UBound(Split(strA, ","))
        intA(intloop) = Split(strA, ",")(intloop)
    Next
    Picture1.DrawMode = vbCopyPen
    Picture1.DrawStyle = vbSolid
    Picture1.DrawWidth = 1.5
    '第一块
    X1 = 0: Y1 = 127 - intA(Diff.NoL): X2 = intA(Diff.NL): Y2 = 127 - intA(Diff.NoL)
    Picture1.Line (X1, Y1)-(X2, Y2), vbBlue
    
    X1 = intA(Diff.NL): Y1 = 127 - intA(Diff.NoL): X2 = intA(Diff.NL): Y2 = 127 - intA(Diff.LMU)
    Picture1.Line (X1, Y1)-(X2, Y2), vbBlue
    
    X1 = intA(Diff.NL): Y1 = 127 - intA(Diff.LMU): X2 = 0: Y2 = 127 - intA(Diff.LMD)
    Picture1.Line (X1, Y1)-(X2, Y2), vbBlue
    

    '第二块
    Picture1.Line (0, 127 - intA(Diff.RM))-(intA(Diff.RMN), 127 - intA(Diff.RM)), vbBlue
    Picture1.Line (intA(Diff.NL), 127 - intA(Diff.LMN))-(intA(Diff.RMN), 127 - intA(Diff.MN)), vbBlue
    Picture1.Line (intA(Diff.RMN), 127 - intA(Diff.MN))-(intA(Diff.RMN), 127 - 127), vbBlue
    
    '第三块
    Picture1.Line (intA(Diff.NL), 127 - intA(Diff.NoN))-(intA(Diff.NE), 127 - intA(Diff.NoN)), vbBlue
    Picture1.Line (intA(Diff.NE), 127 - intA(Diff.NoN))-(intA(Diff.NE), 127 - 127), vbBlue
   
    Picture1.Line (intA(Diff.RMN), 127 - intA(Diff.RN))-(intA(Diff.NE), 127 - intA(Diff.RN)), vbBlue
    
    '第四块
    Picture1.Line (intA(Diff.NE), 127 - intA(Diff.NoE))-(127, 127 - intA(Diff.NoE)), vbBlue
    
    Picture1.DrawWidth = 1
    '外框
    Picture1.Line (0, 0)-(128, 0), vbBlack
    Picture1.Line (0, 0)-(0, 128), vbBlack
    Picture1.Line (127, 0)-(127, 128), vbBlack
    Picture1.Line (0, 127)-(127, 127), vbBlack
    '虚线
    Picture1.DrawStyle = vbDot
    
    X1 = 0: Y1 = 127 - intA(Diff.LL): X2 = intA(Diff.NL): Y2 = 127 - intA(Diff.LL)
    Picture1.Line (X1, Y1)-(X2, Y2), vbBlack
    
    X1 = 0: Y1 = 127 - intA(Diff.AL): X2 = intA(Diff.NL): Y2 = 127 - intA(Diff.AL)
    Picture1.Line (X1, Y1)-(X2, Y2), vbBlack
    
    Picture1.Line (intA(Diff.NL), 127 - intA(Diff.ln))-(intA(Diff.NE), 127 - intA(Diff.ln)), vbBlack
    
    Picture1.Line ((intA(Diff.NE) + intA(Diff.FNE)), 127 - intA(Diff.NoE))-((intA(Diff.NE) + intA(Diff.FNE)), 127 - 127), vbBlack
    Picture1.Line ((intA(Diff.NL) + intA(Diff.FLN)), 127 - intA(Diff.NoN))-((intA(Diff.NL) + intA(Diff.FLN)), 127 - intA(Diff.LMN)), vbBlack
    Picture1.Line ((intA(Diff.NL) + intA(Diff.FLN)), 127 - intA(Diff.LMN))-((intA(Diff.RMN) + intA(Diff.FMN)), 127 - intA(Diff.MN)), vbBlack
    

    SavePicture Picture1.Image, strFileName
    DrawDiff5AL = True
End Function


Public Function Draw_YDA_111(arrHigh() As Double, arrVAL() As Double, arrLow() As Double, strImgPath As String, str标本号 As String) As String
    Dim x As Integer, Y As Double
    Const int_左边距 = 20, int_下边距 = 2
    Picture1.Width = 5115: Picture1.Height = 3795
    Picture1.Scale (0, 18)-(250, 0)
    Picture1.BackColor = vbWhite

'    For x = 30 To 210
'        picDraw.PSet (x + int_左边距 / 2, (arrHigh(0) / x ^ 2 + arrHigh(1) / x + arrHigh(2)) + int_下边距), vbRed
'        picDraw.PSet (x + int_左边距 / 2, (arrVAL(0) / x ^ 2 + arrVAL(1) / x + arrVAL(2)) + int_下边距), vbBlack
'        picDraw.PSet (x + int_左边距 / 2, (arrLow(0) / x ^ 2 + arrLow(1) / x + arrLow(2)) + int_下边距), vbGreen
'    Next
    For x = 30 To 210
        '高限曲线
        Picture1.Line (x + int_左边距, (arrHigh(0) / x ^ 2 + arrHigh(1) / x + arrHigh(2)) + int_下边距)- _
                    (x - 1 / 2 + int_左边距, (arrHigh(0) / (x - 1 / 2) ^ 2 + arrHigh(1) / (x - 1 / 2) + _
                    arrHigh(2)) + int_下边距), vbRed
        Picture1.Line (x + int_左边距, (arrHigh(0) / x ^ 2 + arrHigh(1) / x + arrHigh(2)) + int_下边距)- _
                    (x + 1 / 2 + int_左边距, (arrHigh(0) / (x + 1 / 2) ^ 2 + arrHigh(1) / (x + 1 / 2) + _
                    arrHigh(2)) + int_下边距), vbRed
        '检验结果曲线
        Picture1.DrawWidth = 2
        Picture1.Line (x + int_左边距, (arrVAL(0) / x ^ 2 + arrVAL(1) / x + arrVAL(2)) + int_下边距)- _
                    (x - 1 / 2 + int_左边距, (arrVAL(0) / (x - 1 / 2) ^ 2 + arrVAL(1) / (x - 1 / 2) + _
                    arrVAL(2)) + int_下边距), vbBlack
        Picture1.Line (x + int_左边距, (arrVAL(0) / x ^ 2 + arrVAL(1) / x + arrVAL(2)) + int_下边距)- _
                    (x + 1 / 2 + int_左边距, (arrVAL(0) / (x + 1 / 2) ^ 2 + arrVAL(1) / (x + 1 / 2) + _
                    arrVAL(2)) + int_下边距), vbBlack
        Picture1.DrawWidth = 1
        
        '低限曲线
        Picture1.Line (x + int_左边距, (arrLow(0) / x ^ 2 + arrLow(1) / x + arrLow(2)) + int_下边距)- _
                    (x - 1 / 2 + int_左边距, (arrLow(0) / (x - 1 / 2) ^ 2 + arrLow(1) / (x - 1 / 2) + _
                    arrLow(2)) + int_下边距), vbGreen
        Picture1.Line (x + int_左边距, (arrLow(0) / x ^ 2 + arrLow(1) / x + arrLow(2)) + int_下边距)- _
                    (x + 1 / 2 + int_左边距, (arrLow(0) / (x + 1 / 2) ^ 2 + arrLow(1) / (x + 1 / 2) + _
                    arrLow(2)) + int_下边距), vbGreen
    Next
    
    'X 轴线
    Picture1.Line (int_左边距, int_下边距)-(int_左边距 + 220, int_下边距)
    Picture1.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 + 0.3)
    Picture1.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 - 0.3)
    'X 轴刻度
    Picture1.Line (int_左边距 + 10, int_下边距)-(int_左边距 + 10, int_下边距 + 0.3)
    Picture1.Line (int_左边距 + 30, int_下边距)-(int_左边距 + 30, int_下边距 + 0.3)
    Picture1.Line (int_左边距 + 70, int_下边距)-(int_左边距 + 70, int_下边距 + 0.3)
    Picture1.Line (int_左边距 + 120, int_下边距)-(int_左边距 + 120, int_下边距 + 0.3)
    Picture1.Line (int_左边距 + 150, int_下边距)-(int_左边距 + 150, int_下边距 + 0.3)
    Picture1.Line (int_左边距 + 200, int_下边距)-(int_左边距 + 200, int_下边距 + 0.3)
    
    'Y 轴线
    Picture1.Line (int_左边距, int_下边距)-(int_左边距, int_下边距 + 14)
    Picture1.Line (int_左边距, int_下边距 + 14)-(int_左边距 - 3, int_下边距 + 14 - 0.5)
    Picture1.Line (int_左边距, int_下边距 + 14)-(int_左边距 + 3, int_下边距 + 14 - 0.5)
    'Y 轴刻度
    Picture1.Line (int_左边距, int_下边距 + 5)-(int_左边距 + 3, int_下边距 + 5)
    Picture1.Line (int_左边距, int_下边距 + 10)-(int_左边距 + 3, int_下边距 + 10)
    Picture1.Line (int_左边距, int_下边距 + 12)-(int_左边距 + 3, int_下边距 + 12)
    
'    '标题
'    With Picture1
'        .CurrentX = int_左边距 + 130
'        .CurrentY = int_下边距 + 15
'        .FontSize = 12
'        .FontBold = True
'    End With
'    Picture1.Print "血液粘度曲线"
    
    
    'X 周标签
    Picture1.FontBold = False
    With Picture1
        .CurrentX = int_左边距 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 0
    
    With Picture1
        .CurrentX = int_左边距 + 10 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 10
    
    With Picture1
        .CurrentX = int_左边距 + 30 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 30
    
    With Picture1
        .CurrentX = int_左边距 + 70 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 70
    
    With Picture1
        .CurrentX = int_左边距 + 120 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 120
    
    With Picture1
        .CurrentX = int_左边距 + 150 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 150
    
    With Picture1
        .CurrentX = int_左边距 + 200 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 200
    
    With Picture1
        .CurrentX = int_左边距 + 230 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print "V"
    
    
    'Y 轴标签
    With Picture1
        .CurrentX = int_左边距 - 17
        .CurrentY = int_下边距 + 5 + 0.5
        .FontSize = 10
    End With
    Picture1.Print 5
    
    With Picture1
        .CurrentX = int_左边距 - 17
        .CurrentY = int_下边距 + 10 + 0.5
        .FontSize = 10
    End With
    Picture1.Print 10
    
    With Picture1
        .CurrentX = int_左边距 - 17
        .CurrentY = int_下边距 + 12 + 0.5
        .FontSize = 10
    End With
    Picture1.Print 12
    
    With Picture1
        .CurrentX = 0
        .CurrentY = int_下边距 + 15 + 0.5
        .FontSize = 10
    End With
    Picture1.Print "mpa·s"
    
    If Dir(strImgPath & "\YDA-111_" & str标本号 & ".JPG") <> "" Then
        Kill strImgPath & "\YDA-111_" & str标本号 & ".JPG"
    End If
    Draw_YDA_111 = strImgPath & "\YDA-111_" & str标本号 & ".JPG"
    
    SavePic Picture1.Image, Draw_YDA_111, "JPG"
    
    'Call SavePicture(Picture1.Image, Draw_YDA_111)
End Function

'clsLISDev_File_Fascow   2010D 血流变曲线图形
Public Function Draw_2010D(arrHigh() As Double, arrVAL() As Double, arrLow() As Double, arrNianDu() As Double, strImgPath As String, str标本号 As String) As String
    Dim intloop As Integer
    Dim x As Integer
   
    Dim dblAA As Double, dblBB As Double, dblc As Double
    Dim dblAA1 As Double, dblBB1 As Double, dblC1 As Double
    Dim dblAA2 As Double, dblBB2 As Double, dblC2 As Double
    
    Const int_左边距 = 20, int_下边距 = 2
    
    Call ImageCalc(arrLow(0), arrVAL(0), arrHigh(0), dblAA2, dblBB2, dblC2)
    Call ImageCalc(arrLow(1), arrVAL(1), arrHigh(1), dblAA, dblBB, dblc)
    Call ImageCalc(arrLow(2), arrVAL(2), arrHigh(2), dblAA1, dblBB1, dblC1)
    
    Picture1.Width = 5115: Picture1.Height = 3795
    Picture1.Scale (0, 36)-(250, 0)
    Picture1.BackColor = vbWhite
    
    For intloop = 15 To 200
        x = intloop
        
        '高限曲线
        Picture1.Line (x + int_左边距, (dblAA1 * Exp(dblBB1 * x ^ dblC1)) + int_下边距)- _
                    (x - 1 / 2 + int_左边距, (dblAA1 * Exp(dblBB1 * (x - 1 / 2) ^ dblC1)) + int_下边距), vbRed
        Picture1.Line (x + int_左边距, (dblAA1 * Exp(dblBB1 * x ^ dblC1)) + int_下边距)- _
                    (x + 1 / 2 + int_左边距, (dblAA1 * Exp(dblBB1 * (x + 1 / 2) ^ dblC1)) + int_下边距), vbRed
        
       '检验结果曲线
        Picture1.DrawWidth = 2
        Picture1.Line (x + int_左边距, (dblAA * Exp(dblBB * x ^ dblc)) + int_下边距)- _
                    (x - 1 + int_左边距, (dblAA * Exp(dblBB * (x - 1) ^ dblc)) + int_下边距), vbBlack
        Picture1.Line (x + int_左边距, (dblAA * Exp(dblBB * x ^ dblc)) + int_下边距)- _
                    (x + 1 + int_左边距, (dblAA * Exp(dblBB * (x + 1) ^ dblc)) + int_下边距), vbBlack
        Picture1.DrawWidth = 1
        
        '低限曲线
        Picture1.Line (x + int_左边距, (dblAA2 * Exp(dblBB2 * x ^ dblC2)) + int_下边距)- _
                    (x - 1 / 2 + int_左边距, (dblAA2 * Exp(dblBB2 * (x - 1 / 2) ^ dblC2)) + int_下边距), vbGreen
        Picture1.Line (x + int_左边距, (dblAA2 * Exp(dblBB2 * x ^ dblC2)) + int_下边距)- _
                    (x + 1 / 2 + int_左边距, (dblAA2 * Exp(dblBB2 * (x + 1 / 2) ^ dblC2)) + int_下边距), vbGreen
                    
        '血浆粘度
        Picture1.Line (int_左边距, arrNianDu(1) + int_下边距)-(int_左边距 + 200, arrNianDu(1) + int_下边距)
        If intloop Mod 20 = 0 Then
            Picture1.Line (x + int_左边距, arrNianDu(0) + int_下边距)-(x + int_左边距 - 2.5, arrNianDu(0) + int_下边距)
            Picture1.Line (x + int_左边距, arrNianDu(0) + int_下边距)-(x + int_左边距 + 2.5, arrNianDu(0) + int_下边距)
            
            Picture1.Line (x + int_左边距, arrNianDu(2) + int_下边距)-(x + int_左边距 - 2.5, arrNianDu(2) + int_下边距)
            Picture1.Line (x + int_左边距, arrNianDu(2) + int_下边距)-(x + int_左边距 + 2.5, arrNianDu(2) + int_下边距)
            
        End If
    Next
    
     'X 轴线
    Picture1.Line (int_左边距, int_下边距)-(int_左边距 + 220, int_下边距)
    Picture1.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 + 0.3)
    Picture1.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 - 0.3)
    
    'Y 轴线
    Picture1.Line (int_左边距, int_下边距)-(int_左边距, int_下边距 + 34)
    Picture1.Line (int_左边距, int_下边距 + 34)-(int_左边距 - 4, int_下边距 + 34 - 0.5)
    Picture1.Line (int_左边距, int_下边距 + 34)-(int_左边距 + 4, int_下边距 + 34 - 0.5)


    'X 周标签，刻度线
    For intloop = 0 To 200 Step 20
        Picture1.Line (int_左边距 + intloop, int_下边距)-(int_左边距 + intloop, int_下边距 + 0.5)
        Picture1.FontBold = False
        With Picture1
            .CurrentX = int_左边距 - 8 + intloop
            .CurrentY = int_下边距 - 0.3
            .FontSize = 10
        End With
        Picture1.Print intloop
    Next
    
    With Picture1
        .CurrentX = int_左边距 + 210 - 8
        .CurrentY = int_下边距 + 2
        .FontSize = 10
    End With
    Picture1.Print "(1/S)"
    
    
    'Y 轴标签
    For intloop = 6 To 30 Step 6
        Picture1.Line (int_左边距, int_下边距 + intloop)-(int_左边距 + 3, int_下边距 + intloop)
        With Picture1
            .CurrentX = int_左边距 - 17
            .CurrentY = int_下边距 + intloop + 0.5
            .FontSize = 10
        End With
        Picture1.Print intloop
    Next
    
    With Picture1
        .CurrentX = int_左边距 + 1
        .CurrentY = int_下边距 + 32 + 0.5
        .FontSize = 10
    End With
    Picture1.Print "(mpas)"
    
    If Dir(strImgPath & "\2010D_" & str标本号 & ".JPG") <> "" Then
        Kill strImgPath & "\2010D_" & str标本号 & ".JPG"
    End If
    Draw_2010D = strImgPath & "\2010D_" & str标本号 & ".JPG"
    
    SavePic Picture1.Image, Draw_2010D, "JPG"

End Function

Private Sub ImageCalc(dblQlow As Double, dblQmid As Double, dblQhigh As Double, dblAA As Double, dblBB As Double, dblc As Double)
    Dim dblE As Double
    Dim dblC1 As Double
    Dim dblC2 As Double
    Dim dblD As Double
    Dim dblY As Double
    Dim dblY1 As Double
    Dim dblY2 As Double
 

    dblE = 0.0000001
    dblC1 = 1
    dblC2 = -5
    
    dblD = Log(dblQlow / dblQmid) / Log(dblQlow / dblQhigh)
    dblY1 = (1 - (30 / 3) ^ dblC1) / (1 - (200 / 3) ^ dblC1) - dblD
    dblY2 = (1 - (30 / 3) ^ dblC2) / (1 - (200 / 3) ^ dblC2) - dblD
    
    While Abs(dblY2 - dblY1) > dblE
        dblc = (dblC1 + dblC2) / 2
        dblY = (1 - (30 / 3) ^ dblc) / (1 - (200 / 3) ^ dblc) - dblD
        
        If dblY * dblY1 > 0 Then
            dblY1 = dblY
            dblC1 = dblc
        Else
            dblY2 = dblY
            dblC2 = dblc
        End If
    Wend
    
    dblBB = Log(dblQlow / dblQmid) / (3 ^ dblc - 30 ^ dblc)
    dblAA = dblQhigh / Exp(dblBB * 200 ^ dblc)
End Sub

Public Function Draw_SA6000(strLow As String, strVal As String, strHigh As String, strImgPath As String, str标本号 As String) As String
    Dim intloop As Integer
    Dim x As Single
    Dim sngX1 As Single, sngY1 As Single, sngX2 As Single, sngY2 As Single
    Dim varTmp As Variant
    Const int_左边距 = 20, int_下边距 = 3

    Picture1.Width = 5115: Picture1.Height = 3795
    Picture1.Scale (0, 42)-(255, 0)
    Picture1.BackColor = vbWhite

    For intloop = 1 To 220
        x = intloop
        '高限曲线
        sngX1 = 1: sngY1 = Split(Split(strLow, ";")(1), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(1), ",")(0)
    
        Picture1.Line (x + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(x + 3 - 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x + 3 - 7 / 50 / 2) + int_下边距), vbRed
        Picture1.Line (x + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(x + 3 + 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x + 3 + 7 / 50 / 2) + int_下边距), vbRed

       '检验结果曲线
        sngX1 = 1: sngY1 = Split(Split(strVal, ",")(1), "-")(1): sngX2 = 200: sngY2 = Split(Split(strVal, ",")(4), "-")(1)
        
        Picture1.DrawWidth = 1
        Picture1.Line (x + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(x + 3 - 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x + 3 - 7 / 50 / 2) + int_下边距), vbGreen
        Picture1.Line (x + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(x + 3 + 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x + 3 + 7 / 50 / 2) + int_下边距), vbGreen
        Picture1.DrawWidth = 1

        '低限曲线
        sngX1 = 1: sngY1 = Split(Split(strLow, ";")(0), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(0), ",")(0)
        
        Picture1.Line (x + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(x + 3 - 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x + 3 - 7 / 50 / 2) + int_下边距), vbMagenta
        Picture1.Line (x + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(x + 3 + 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, x + 3 + 7 / 50 / 2) + int_下边距), vbMagenta
    Next

     'X 轴线
    Picture1.Line (int_左边距, int_下边距)-(int_左边距 + 220, int_下边距)
    Picture1.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 + 0.3)
    Picture1.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 - 0.3)

    'Y 轴线
    Picture1.Line (int_左边距, int_下边距)-(int_左边距, int_下边距 + 38)
    Picture1.Line (int_左边距, int_下边距 + 38)-(int_左边距 - 4, int_下边距 + 38 - 0.5)
    Picture1.Line (int_左边距, int_下边距 + 38)-(int_左边距 + 4, int_下边距 + 38 - 0.5)


    'X 轴刻度
    Picture1.Line (int_左边距 + 3, int_下边距)-(int_左边距 + 3, int_下边距 + 0.5)
    Picture1.Line (int_左边距 + 10, int_下边距)-(int_左边距 + 10, int_下边距 + 0.5)
    Picture1.Line (int_左边距 + 30, int_下边距)-(int_左边距 + 30, int_下边距 + 0.5)
    Picture1.Line (int_左边距 + 100, int_下边距)-(int_左边距 + 100, int_下边距 + 0.5)
    Picture1.Line (int_左边距 + 200, int_下边距)-(int_左边距 + 200, int_下边距 + 0.5)
    
    Picture1.FontBold = False
    With Picture1
        .CurrentX = int_左边距 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 1
    
    With Picture1
        .CurrentX = int_左边距 + 3 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 3
    
    With Picture1
        .CurrentX = int_左边距 + 10 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 10
    
    With Picture1
        .CurrentX = int_左边距 + 30 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 30
    
    With Picture1
        .CurrentX = int_左边距 + 100 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 100
    
    With Picture1
        .CurrentX = int_左边距 + 200 - 18
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 200
    
    With Picture1
        .CurrentX = int_左边距 + 220 - 18
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print "(1/s)"

    'Y 轴标签
    For intloop = 0 To 36 Step 2
        Picture1.Line (int_左边距, int_下边距 + intloop)-(int_左边距 + 3, int_下边距 + intloop)
        With Picture1
            .CurrentX = int_左边距 - 17
            .CurrentY = int_下边距 + intloop + 0.5
            .FontSize = 10
        End With
        Picture1.Print intloop
    Next

    With Picture1
        .CurrentX = int_左边距 + 1
        .CurrentY = int_下边距 + 32 + 0.5
        .FontSize = 10
    End With
    Picture1.Print "(mpas)"

    If Dir(strImgPath & "\SA6000_" & str标本号 & ".jpg") <> "" Then
        Kill strImgPath & "\SA6000_" & str标本号 & ".jpg"
    End If
    Draw_SA6000 = strImgPath & "\SA6000_" & str标本号 & ".jpg"

    SavePic Picture1.Image, Draw_SA6000, "jpg"
End Function

Private Function GetY_SA6000(sngX1 As Single, sngY1 As Single, sngX2 As Single, sngY2 As Single, sngX As Single) As Single
    Dim dblA As Double, dblB As Double, sngY As Single
    
    dblB = (Sqr(sngY1) - Sqr(sngY2)) / (Sqr(1 / sngX1) - Sqr(1 / sngX2))
    dblA = Sqr(sngY1) - dblB * Sqr(1 / sngX1)
    sngY = dblA ^ 2 + dblB ^ 2 / sngX + 2 * dblA * dblB * Sqr(1 / sngX)
    GetY_SA6000 = sngY
End Function

Public Function Draw_ZL6000C(strLow As String, strVal As String, strHigh As String, strImgPath As String, str标本号 As String) As String
    Dim intloop As Integer
    Dim x As Single
    Dim sngX1 As Single, sngY1 As Single, sngX2 As Single, sngY2 As Single
    Dim varTmp As Variant
    Dim Z As Single
    Dim U As Single
    Const int_左边距 = 20, int_下边距 = 3

    Picture1.Width = 5115: Picture1.Height = 3795
    Picture1.Scale (0, 42)-(330, 0)
    Picture1.BackColor = vbWhite

     'X 轴线
    Picture1.Line (int_左边距, int_下边距)-(int_左边距 + 310, int_下边距), vbWhite
    Picture1.Line (int_左边距 + 310, int_下边距)-(int_左边距 + 305, int_下边距 + 0.3), vbBlue
    Picture1.Line (int_左边距 + 310, int_下边距)-(int_左边距 + 305, int_下边距 - 0.3), vbBlue

    'Y 轴线
    Picture1.Line (int_左边距, int_下边距)-(int_左边距, int_下边距 + 38), vbWhite
    Picture1.Line (int_左边距, int_下边距 + 38)-(int_左边距 - 4, int_下边距 + 38 - 0.5), vbBlue
    Picture1.Line (int_左边距, int_下边距 + 38)-(int_左边距 + 4, int_下边距 + 38 - 0.5), vbBlue
    'X 轴刻度
    
    Picture1.Line (int_左边距 + 0, int_下边距)-(int_左边距 + 0, int_下边距 + 40), vbBlue
    Picture1.Line (int_左边距 + 60, int_下边距)-(int_左边距 + 60, int_下边距 + 40), vbBlue
    Picture1.Line (int_左边距 + 75, int_下边距)-(int_左边距 + 75, int_下边距 + 40), vbBlue
    Picture1.Line (int_左边距 + 130, int_下边距)-(int_左边距 + 130, int_下边距 + 40), vbBlue
    Picture1.Line (int_左边距 + 180, int_下边距)-(int_左边距 + 180, int_下边距 + 40), vbBlue
    Picture1.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 220, int_下边距 + 40), vbBlue
    Picture1.Line (int_左边距 + 260, int_下边距)-(int_左边距 + 260, int_下边距 + 40), vbBlue
    Picture1.ForeColor = vbBlack
        'Y 轴标签
    For intloop = 0 To 35 Step 5
        Picture1.Line (int_左边距, int_下边距 + intloop)-(int_左边距 + 320, int_下边距 + intloop), vbBlue
        With Picture1
            .CurrentX = int_左边距 - 22
            .CurrentY = int_下边距 + intloop + 0.5
            .FontSize = 10
        End With
        Picture1.Print intloop
    Next
    
    Picture1.FontBold = False
    Picture1.ForeColor = vbBlack
    With Picture1
        .CurrentX = int_左边距 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 1
    
    With Picture1
        .CurrentX = int_左边距 + 60 - 10
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 3
    
    With Picture1
        .CurrentX = int_左边距 + 130 - 10
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 10
    
    With Picture1
        .CurrentX = int_左边距 + 180 - 10
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 30
    
    With Picture1
        .CurrentX = int_左边距 + 220 - 10
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 100
    
    With Picture1
        .CurrentX = int_左边距 + 260 - 10
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print 200
    
    With Picture1
        .CurrentX = int_左边距 + 290 - 12
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    Picture1.Print "(1/s)"



    With Picture1
        .CurrentX = int_左边距 + 1
        .CurrentY = int_下边距 + 32 + 0.5
        .FontSize = 10
    End With
    Picture1.Print "(mpas)"

    For intloop = 20 To 95
        x = intloop
'       低限曲线
        sngX1 = 1: sngY1 = Split(Split(strLow, ";")(1), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(1), ",")(0)
        Picture1.Line (x + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x - 2) + int_下边距)-(x + 7 / 8 + 0.2 + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x - 2 + 7 / 8 + 0.2) + int_下边距), vbRed
'        检验结果曲线
         sngX1 = 1: sngY1 = Split(Split(strVal, ",")(1), "-")(1): sngX2 = 5: sngY2 = Split(Split(strVal, ",")(2), "-")(1)
         Picture1.Line (x + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x - 2) + int_下边距 - 0.1415926)-(x + 7 / 8 + 0.2 + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x + 7 / 8 + 0.2 - 2) + int_下边距 - 0.1415926), vbGreen
     
        '高限曲线
        sngX1 = 1: sngY1 = Split(Split(strLow, ";")(0), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(0), ",")(0)
        Picture1.Line (x + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x - 2) + int_下边距)-(x + 7 / 8 + 0.2 + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x - 2 + 7 / 8 + 0.2) + int_下边距), vbMagenta
    
    Next
    For intloop = 95 To 200
        x = intloop
        '低限曲线
        sngX1 = 1: sngY1 = Split(Split(strLow, ";")(1), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(1), ",")(0)
        Picture1.Line (x + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(x + 7 / 8 + 0.2 + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x + 7 / 8 + 0.2) + int_下边距), vbRed
        '检验结果曲线
         sngX1 = 5: sngY1 = Split(Split(strVal, ",")(2), "-")(1): sngX2 = 30: sngY2 = Split(Split(strVal, ",")(3), "-")(1)
         Picture1.Line (x + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x) + int_下边距 - 0.1415926)-(x + 7 / 8 + 0.2 + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x + 7 / 8 + 0.2) + int_下边距 - 0.1415926), vbGreen

        '高限曲线
        sngX1 = 1: sngY1 = Split(Split(strLow, ";")(0), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(0), ",")(0)
        Picture1.Line (x + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(x + 7 / 8 + 0.2 + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x + 7 / 8 + 0.2) + int_下边距), vbMagenta

    Next

    For intloop = 200 To 330
        x = intloop
        '低限曲线
        sngX1 = 1: sngY1 = Split(Split(strLow, ";")(1), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(1), ",")(0)
        Picture1.Line (x + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(x + 7 / 8 + 0.2 + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x + 7 / 8 + 0.2) + int_下边距), vbRed
        '检验结果曲线
         sngX1 = 30: sngY1 = Split(Split(strVal, ",")(3), "-")(1): sngX2 = 200: sngY2 = Split(Split(strVal, ",")(4), "-")(1)
         Picture1.Line (x + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x) + int_下边距 - 0.1415926)-(x + 7 / 8 + 0.2 + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x + 7 / 8 + 0.2) + int_下边距 - 0.1415926), vbGreen

        '高限曲线
        sngX1 = 1: sngY1 = Split(Split(strLow, ";")(0), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(0), ",")(0)
        Picture1.Line (x + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(x + 7 / 8 + 0.2 + int_左边距 - 20, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, x + 7 / 8 + 0.2) + int_下边距), vbMagenta

    Next

    If Dir(strImgPath & "\ZL6000C_" & str标本号 & ".jpg") <> "" Then
        Kill strImgPath & "\ZL6000C_" & str标本号 & ".jpg"
    End If
    Draw_ZL6000C = strImgPath & "\ZL6000C_" & str标本号 & ".jpg"

    SavePic Picture1.Image, Draw_ZL6000C, "jpg"
End Function

Private Function GetY_ZL6000C(sngX1 As Single, sngY1 As Single, sngX2 As Single, sngY2 As Single, sngX As Single) As Single
    Dim dblA As Double, dblB As Double, sngY As Single
    Dim dblc As Double
    dblB = (Sqr(sngY1) - Sqr(sngY2)) / (Sqr(1 / sngX1) - Sqr(1 / sngX2))
    dblA = Sqr(sngY1) - dblB * Sqr(1 / sngX1)
    sngY = (dblA ^ 2 + dblB ^ 2) / sngX + 18 * dblA * dblB * Sqr(1 / sngX) - 1
    GetY_ZL6000C = sngY
End Function

Public Function Draw_mvis(arrHigh() As Double, arrVAL() As Double, arrLow() As Double, arrNianDu() As Double, strImgPath As String, str标本号 As String) As String
    Dim intloop As Integer
    Dim x As Integer
   
    Dim dblAA As Double, dblBB As Double, dblc As Double
    Dim dblAA1 As Double, dblBB1 As Double, dblC1 As Double
    Dim dblAA2 As Double, dblBB2 As Double, dblC2 As Double
    
    Const int_左边距 = 20, int_下边距 = 2
    
    Call ImageCalc(arrLow(0), arrVAL(0), arrHigh(0), dblAA2, dblBB2, dblC2)
    Call ImageCalc(arrLow(1), arrVAL(1), arrHigh(1), dblAA, dblBB, dblc)
    Call ImageCalc(arrLow(2), arrVAL(2), arrHigh(2), dblAA1, dblBB1, dblC1)
        Picture1.Width = 5115: Picture1.Height = 3795
        Picture1.Scale (0, 26)-(250, 0)
        Picture1.BackColor = vbWhite
        
        For intloop = 2 To 200
            x = intloop
            Picture1.DrawWidth = 2
            '高限曲线
            Picture1.Line (x + int_左边距, (dblAA1 * Exp(dblBB1 * x ^ dblC1)) + int_下边距)- _
                        (x - 1 + int_左边距, (dblAA1 * Exp(dblBB1 * (x - 1) ^ dblC1)) + int_下边距), &H808080
            Picture1.Line (x + int_左边距, (dblAA1 * Exp(dblBB1 * x ^ dblC1)) + int_下边距)- _
                        (x + 1 + int_左边距, (dblAA1 * Exp(dblBB1 * (x + 1) ^ dblC1)) + int_下边距), &H808080
            
           '检验结果曲线

            Picture1.Line (x + int_左边距, (dblAA * Exp(dblBB * x ^ dblc)) + int_下边距)- _
                        (x - 1 + int_左边距, (dblAA * Exp(dblBB * (x - 1) ^ dblc)) + int_下边距), vbRed
            Picture1.Line (x + int_左边距, (dblAA * Exp(dblBB * x ^ dblc)) + int_下边距)- _
                        (x + 1 + int_左边距, (dblAA * Exp(dblBB * (x + 1) ^ dblc)) + int_下边距), vbRed
'            Picture1.DrawWidth = 1
            
            '低限曲线
            Picture1.Line (x + int_左边距, (dblAA2 * Exp(dblBB2 * x ^ dblC2)) + int_下边距)- _
                        (x - 1 + int_左边距, (dblAA2 * Exp(dblBB2 * (x - 1) ^ dblC2)) + int_下边距), &H808080
            Picture1.Line (x + int_左边距, (dblAA2 * Exp(dblBB2 * x ^ dblC2)) + int_下边距)- _
                        (x + 1 + int_左边距, (dblAA2 * Exp(dblBB2 * (x + 1) ^ dblC2)) + int_下边距), &H808080
                        
            '血浆粘度
            Picture1.Line (int_左边距, arrNianDu(1) + int_下边距)-(int_左边距 + 200, arrNianDu(1) + int_下边距), vbGreen
        Next
        
         'X 轴线
        Picture1.Line (int_左边距, int_下边距)-(int_左边距 + 220, int_下边距)
        Picture1.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 + 0.3)
        Picture1.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 - 0.3)
        
        'Y 轴线
        Picture1.Line (int_左边距, int_下边距)-(int_左边距, int_下边距 + 34)
        Picture1.Line (int_左边距, int_下边距 + 34)-(int_左边距 - 4, int_下边距 + 34 - 0.5)
        Picture1.Line (int_左边距, int_下边距 + 34)-(int_左边距 + 4, int_下边距 + 34 - 0.5)
    
    
        'X 周标签，刻度线
        For intloop = 0 To 200 Step 20
            Picture1.Line (int_左边距 + intloop, int_下边距)-(int_左边距 + intloop, int_下边距 + 0.5)
            Picture1.FontBold = False
            With Picture1
                .CurrentX = int_左边距 - 8 + intloop
                .CurrentY = int_下边距 - 0.3
                .FontSize = 10
            End With
            Picture1.Print intloop
        Next
        
        With Picture1
            .CurrentX = int_左边距 + 210 - 8
            .CurrentY = int_下边距 + 2
            .FontSize = 10
        End With
        Picture1.Print "切变率(1/S)"
        
        
        'Y 轴标签
        For intloop = 2 To 20 Step 2
            Picture1.Line (int_左边距, int_下边距 + intloop)-(int_左边距 + 3, int_下边距 + intloop)
            With Picture1
                .CurrentX = int_左边距 - 17
                .CurrentY = int_下边距 + intloop + 0.5
                .FontSize = 10
            End With
            Picture1.Print intloop
        Next
        
        With Picture1
            .CurrentX = int_左边距 + 1
            .CurrentY = int_下边距 + 22 + 0.5
            .FontSize = 10
        End With
        Picture1.Print "粘度(mpas)"
        
        If Dir(strImgPath & "\mvis_" & str标本号 & ".JPG") <> "" Then
            Kill strImgPath & "\mvis_" & str标本号 & ".JPG"
        End If
        Draw_mvis = strImgPath & "\mvis_" & str标本号 & ".JPG"
        
        SavePic Picture1.Image, Draw_mvis, "JPG"
End Function



Public Sub Draw()
    Dim i  As Integer
    Dim j As Integer
    Picture1.Width = 7000
    Picture1.Height = 4000
    Picture1.Cls
    Picture1.BackColor = RGB(255, 255, 255)
    Picture1.Line (800, 400)-(6600, 3100), RGB(200, 200, 255), BF
    Picture1.Line (800, 400)-(800, 3100), vbBlack, BF
    Picture1.Line (800, 400)-(6600, 400), vbBlack, BF
    Picture1.Line (6600, 400)-(6600, 3100), vbBlack, BF
    Picture1.Line (800, 3100)-(6600, 3100), vbBlack, BF
    Picture1.Line (750, 2800)-(6600, 2800), vbBlack, BF
    Picture1.CurrentX = Picture1.CurrentX - 6600 + 500
    Picture1.CurrentY = Picture1.CurrentY - 80
    Picture1.Print 0
    Picture1.Line (780, 3000)-(820, 3000), vbBlack, BF
    Picture1.Line (780, 2900)-(820, 2900), vbBlack, BF
    Picture1.Line (6580, 3000)-(6600, 3000), vbBlack, BF
    Picture1.Line (6580, 2900)-(6600, 2900), vbBlack, BF
    'X轴
    For i = 1 To 5
        Picture1.Line (750, 2900 - i * 500)-(850, 2900 - i * 500), vbBlack, BF
        Picture1.CurrentX = Picture1.CurrentX - 500
        Picture1.CurrentY = Picture1.CurrentY - 80
        Picture1.Print i * 100
        Picture1.Line (6550, 2900 - i * 500)-(6600, 2900 - i * 500), vbBlack, BF
        Picture1.DrawStyle = 2
        Picture1.Line (850, 2900 - i * 500)-(6550, 2900 - i * 500), vbGrayText
        For j = 1 To 5
            Picture1.Line (780, 2900 - (j * 100 + (i - 1) * 500))-(820, 2900 - (j * 100 + (i - 1) * 500)), vbBlack, BF
            Picture1.Line (6580, 2900 - (j * 100 + (i - 1) * 500))-(6600, 2900 - (j * 100 + (i - 1) * 500)), vbBlack, BF
        Next j
    Next i
    'Y轴
    Picture1.CurrentY = 3100
    For i = 1 To 29
        Picture1.Line (200 * i + 800, 400)-(200 * i + 800, 450), vbBlack, BF
        Picture1.DrawStyle = 2
        Picture1.Line (200 * i + 800, 450)-(200 * i + 800, 3050), vbGrayText
        Picture1.Line (200 * i + 800, 3050)-(200 * i + 800, 3150), vbBlack, BF
        If i < 10 Then
            Picture1.CurrentX = Picture1.CurrentX - 80
            Picture1.Print i + 1
        Else
            Picture1.CurrentX = Picture1.CurrentX - 120
            Picture1.Print i + 1
        End If
        i = i + 1
        If i < 29 Then
            Picture1.Line (200 * i + 800, 400)-(200 * i + 800, 420), vbBlack, BF
            Picture1.Line (200 * i + 800, 3080)-(200 * i + 800, 3120), vbBlack, BF
        End If
    Next i

End Sub



Public Sub DrawGraph(LinesArray() As String, ColorArray() As Long, RowCaption() As String)
    Dim i As Integer, j As Integer, RowSize As Integer, ColSize As Integer
    Dim StepSize As Single, ArrayIndex As Integer, lineDimensions() As String
    Dim FirstPoint As Integer, SecondPoint As Integer, lineColor As Long
    Dim str项目 As String

    ColSize = 5
    RowSize = 200
    Picture1.DrawStyle = vbSolid
    Picture1.Line (800, 140)-(800, 350)
    Picture1.Line (800, 140)-(800 + 1200 * UBound(LinesArray) + 1200, 140)
    Picture1.Line (800, 350)-(800 + 1200 * UBound(LinesArray) + 1200, 350)

    For ArrayIndex = LBound(LinesArray) To UBound(LinesArray)
        lineColor = ColorArray(ArrayIndex)
        Picture1.Line (900 + ArrayIndex * 1200, 240)-(900 + 1200 * ArrayIndex + 300, 240), lineColor
        Picture1.Line (800 + 1200 * UBound(LinesArray) + 1200, 140)-(800 + 1200 * UBound(LinesArray) + 1200, 360)
        Picture1.CurrentX = Picture1.CurrentX - 700 - (UBound(LinesArray) - ArrayIndex) * 1200
        Picture1.CurrentY = Picture1.CurrentY - 200
        Picture1.Print RowCaption(ArrayIndex)
        lineDimensions = Split(LinesArray(ArrayIndex), ",")
        Picture1.CurrentX = 800
        Picture1.CurrentY = 2900

        For i = LBound(lineDimensions) To UBound(lineDimensions) - 1
            FirstPoint = 2800 - (CInt(lineDimensions(i)) * ColSize)
            SecondPoint = 2800 - (CInt(lineDimensions(i + 1)) * ColSize)
            Picture1.CurrentY = FirstPoint + 100
            Picture1.CurrentX = 0 + (i * RowSize) + (RowSize / 2) - TextWidth(CInt(lineDimensions(i)))
            Picture1.Line (800 + i * RowSize, FirstPoint)-(800 + (i + 1) * RowSize, SecondPoint), lineColor
        Next i
    Next ArrayIndex

End Sub

