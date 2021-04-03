VERSION 5.00
Begin VB.Form frmAirBubbleMessage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmAirBubbleMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As Object
Private mblnStartUp As Boolean
Private mlngMaxHeight As Long
Private mintPauseTime As Integer
Private mintWaitTime As Integer
Private mintTransparentGrade As Integer

Public Function SetMeLine(ByVal lngBackColor1 As Long, ByVal lngBackColor2 As Long)
    Dim lngR As Long, arrTxt As Variant, i As Long, lngWidth As Long
    '背景
    Call DrawColorToColor(Me, lngBackColor1, lngBackColor2)
    '边框：API=RoundRect
    Me.Line (Screen.TwipsPerPixelX, 0)-(Me.Width - Screen.TwipsPerPixelX, 0), RGB(118, 118, 118)
    Me.Line (Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY)-(Me.Width - Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    Me.Line (0, Screen.TwipsPerPixelY)-(0, Me.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    Me.Line (Me.Width - Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(Me.Width - Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    Me.PSet (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    Me.PSet (Me.Width - Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    Me.PSet (Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    Me.PSet (Me.Width - Screen.TwipsPerPixelX * 2, Me.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)

    '形状
    lngR = CreateRoundRectRgn(0, 0, Me.ScaleX(Me.Width, Me.ScaleMode, vbPixels) + 1, Me.ScaleY(Me.Height, Me.ScaleMode, vbPixels) + 1, 4, 4)
    Call SetWindowRgn(Me.hwnd, lngR, False)
End Function

Public Property Let TransparentGrade(vData As Single)
    mintTransparentGrade = Int(255 * ((100 - vData) / 100))
    Call transparent
End Property

Private Sub transparent()
    Dim ret As Long
    
    ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
    
    SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
    SetLayeredWindowAttributes Me.hwnd, 0, mintTransparentGrade, LWA_ALPHA
End Sub

Public Function ShowContent(ByVal strText As String, ByVal objFont As StdFont, Optional lngLeftGap As Long, Optional lngRightGap As Long, Optional lngRowGap As Long)
    If Not (objFont Is Nothing) Then
        Me.FontName = objFont.Name
        Me.FontSize = objFont.Size
        Me.FontBold = objFont.Bold
        Me.FontItalic = objFont.Italic
        Me.FontStrikethru = objFont.Strikethrough
        Me.FontUnderline = objFont.Underline
    End If
    Call PrintContent(Me, strText, lngLeftGap, lngRightGap, lngRowGap)
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub
