VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAirBubble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrShowText As String
Private mfrmParent As Object
Private mlngMaxHeight As Long
Private mbytShowStyle As Byte
Private mbytShowPostion As Byte
Private mbytShowMode As Byte
Private msglTransparentGrade As Single
Private mlngBackColor As Long
Private mlngBackColor1 As Long
Private mlngBackColor2 As Long
Private mlngForeColor As Long
Private mintWaitTimer As Integer
Private mfrmAirBubbleMessage As frmAirBubbleMessage
Private mfrmAirBubbleTransparence As frmAirBubbleTransparence
Private mfrmAirBubbleTransparenceMessage As frmAirBubbleTransparenceMessage
Private mintGrade As Integer
Private mintPouse As Integer
Private mblnAutoSize As Boolean
Private mlngRowGap As Long
Private mobjFont As StdFont
Private mlngInitHeight As Long
Private mstrSoundTmp As String

Private Const SND_ASYNC = &H1   ' Play asynchronously (see note below).����ʱ���Խ��д��������
Private Const SND_NODEFAULT = &H2 ' Do not use default sound.

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Enum SHOWPOSTION
    ScreenCenter = 0
    OwnerCenter = 1
    RightBottom = 2
End Enum

Public Enum SHOWMODE
    ֱ�ӵ��� = 1
    ���뽥�� = 2
    �������� = 3
End Enum

Private Const GAPLEFT = 80
Private Const GAPHEIGHT = 80
Private Const GAPROW = 80

Private Const WM_PAINT = &HF

Public Enum SoundType
    �� = 0
    ���� = 1
    ����ռ� = 2
    �绰���� = 3
    ������ = 4
    ������ = 5
    ������ = 6
    ��ʾ�� = 7
    ����Ϣ = 8
End Enum

Private WithEvents tmrOpen As timer
Attribute tmrOpen.VB_VarHelpID = -1
Private WithEvents tmrClose As timer
Attribute tmrClose.VB_VarHelpID = -1
Private WithEvents tmrAlert As timer
Attribute tmrAlert.VB_VarHelpID = -1

Private Function InitForm(ByVal objMain As Object)
    On Error Resume Next
    objMain.Controls.Remove "Timer1"
    objMain.Controls.Remove "Timer2"
    objMain.Controls.Remove "Timer3"
    
    Set tmrClose = objMain.Controls.Add("VB.Timer", "Timer1")
    tmrClose.Enabled = False
    tmrClose.Interval = 1
    
    Set tmrOpen = objMain.Controls.Add("VB.Timer", "Timer2")
    tmrOpen.Enabled = False
    tmrOpen.Interval = 1
    
    Set tmrAlert = objMain.Controls.Add("VB.Timer", "Timer3")
    tmrAlert.Enabled = False
    tmrAlert.Interval = 1000
End Function

Public Function OpenTransparentAirBubble(ByVal frmParent As Object, _
                                        ByVal strShowText As String, _
                                        Optional bytShowPostion As Byte = SHOWPOSTION.ScreenCenter, _
                                        Optional bytShowMode As Byte = SHOWMODE.���뽥��, _
                                        Optional sglTransparentGrade As Single = 50, _
                                        Optional lngBackColor As Long = 0, _
                                        Optional lngForeColor As Long = 16777215, _
                                        Optional objFont As StdFont, _
                                        Optional intWaitTimer As Integer = 5, _
                                        Optional lngWidth As Long = 4200, _
                                        Optional lngHeight As Long = 1500, _
                                        Optional bytSound As SoundType = ����, _
                                        Optional blnAutoSize As Boolean)
    '******************************************************************************************************************
    '����:�����ݴ���
    '����:frmParent ���ݵ������߻��ϼ�����
    '****bytShowPostion ���ݵ���ʾλ�ã�0-��Ļ����;1-�����߾���;2-���½ǣ�ȱʡΪ0-��Ļ����
    '****sglTransparentGrade ����͸���̶ȣ�0��100,0��ʾ��͸��;100��ʾ��ȫ͸����ȱʡΪ50��ֻ������͸������
    '****bytShowMode  ��ʾ��ʽ  1��ֱ�ӵ���  2�����뽥��  3���������ϵ���
    '****lngBackColor    ��ѡ�������ɫֵ��ȱʡΪ0(��ɫ)
    '****lngForeColor    ��ѡ����ݵ��ı�������ɫ��ȱʡΪ16777215����ɫ��
    '****objFont ��ѡ����ݵ��ı�������ʾ���壬ȱʡΪ���塢9�Ŵ�С
    '****intWaitTimer    ��ѡ�������ʾʱ�䣬������ʱ����Զ��ر����ݣ���λ���룬���Ϊ0��ʾһֱ��ʾ��ȱʡΪ5��
    '****lngWidth    ��ѡ����ݴ���Ŀ��ȣ�ȱʡΪ4200�
    '****lngHeight   ��ѡ����ݴ���ĸ߶ȣ�ȱʡΪ1500�
    '****bytShowStyle    ��ѡ����ݷ��0-͸������;1-��͸������,ȱʡΪ0-͸�����ݡ�
    '****bytSound    ��ѡ�������ʾʱ����ʾ������ĿǰΪ0��8��0��ʾû��������1��8Ϊǰ���ᵽ��������ȱʡΪ1(��������)
    '����:
    '******************************************************************************************************************

    Set mfrmParent = frmParent
    mstrShowText = strShowText
    mbytShowStyle = 0
    mbytShowPostion = bytShowPostion
    mbytShowMode = bytShowMode
    msglTransparentGrade = sglTransparentGrade
    mlngBackColor = lngBackColor
    mlngForeColor = lngForeColor
    mintWaitTimer = intWaitTimer
    mblnAutoSize = blnAutoSize
    Set mobjFont = objFont
    
    If msglTransparentGrade > 100 Then
        msglTransparentGrade = 100
    ElseIf msglTransparentGrade < 0 Then
        msglTransparentGrade = 0
    End If
    
    Call InitForm(mfrmParent)

    '������ʾ��
    Call PlayWave(bytSound)
    
    mfrmAirBubbleTransparence.Width = lngWidth
    mfrmAirBubbleTransparence.Height = lngHeight
    
    mfrmAirBubbleTransparenceMessage.Width = lngWidth
    mfrmAirBubbleTransparenceMessage.Height = lngHeight

    mfrmAirBubbleTransparence.ForeColor = lngForeColor
    mfrmAirBubbleTransparenceMessage.ForeColor = lngForeColor
    mfrmAirBubbleTransparence.BackColor = lngBackColor
    '��͸������
    Call ShowMessageTransparence

End Function

Public Function OpenUnTransparentAirBubble(ByVal frmParent As Object, _
                                ByVal strShowText As String, _
                                Optional bytShowPostion As Byte = SHOWPOSTION.ScreenCenter, _
                                Optional bytShowMode As Byte = SHOWMODE.���뽥��, _
                                Optional lngBackColor1 As Long = &HFFFFFF, _
                                Optional lngBackColor2 As Long = &HF7F9FA, _
                                Optional lngForeColor As Long = 0, _
                                Optional objFont As StdFont, _
                                Optional intWaitTimer As Integer = 5, _
                                Optional lngWidth As Long = 4200, _
                                Optional lngHeight As Long = 1500, _
                                Optional bytSound As SoundType = ����, _
                                Optional blnAutoSize As Boolean)
    '******************************************************************************************************************
    '����:�����ݴ���
    '����:frmParent ���ݵ������߻��ϼ�����
    '****bytShowPostion ���ݵ���ʾλ�ã�0-��Ļ����;1-�����߾���;2-���½ǣ�ȱʡΪ0-��Ļ����
    '****sglTransparentGrade ����͸���̶ȣ�0��100,0��ʾ��͸��;100��ʾ��ȫ͸����ȱʡΪ50��ֻ������͸������
    '****bytShowMode  ��ʾ��ʽ  1��ֱ�ӵ���  2�����뽥��  3���������ϵ���
    '****lngBackColor    ��ѡ�������ɫֵ��ȱʡΪ0(��ɫ)
    '****lngForeColor    ��ѡ����ݵ��ı�������ɫ��ȱʡΪ16777215����ɫ��
    '****objFont ��ѡ����ݵ��ı�������ʾ���壬ȱʡΪ���塢9�Ŵ�С
    '****intWaitTimer    ��ѡ�������ʾʱ�䣬������ʱ����Զ��ر����ݣ���λ���룬���Ϊ0��ʾһֱ��ʾ��ȱʡΪ5��
    '****lngWidth    ��ѡ����ݴ���Ŀ��ȣ�ȱʡΪ4200�
    '****lngHeight   ��ѡ����ݴ���ĸ߶ȣ�ȱʡΪ1500�
    '****bytShowStyle    ��ѡ����ݷ��0-͸������;1-��͸������,ȱʡΪ0-͸�����ݡ�
    '****bytSound    ��ѡ�������ʾʱ����ʾ������ĿǰΪ0��8��0��ʾû��������1��8Ϊǰ���ᵽ��������ȱʡΪ1(��������)
    '����:
    '******************************************************************************************************************

    Set mfrmParent = frmParent
    mstrShowText = strShowText
    mbytShowStyle = 1
    mbytShowPostion = bytShowPostion
    mbytShowMode = bytShowMode
    mlngBackColor1 = lngBackColor1
    mlngBackColor2 = lngBackColor2
    mlngForeColor = lngForeColor
    mintWaitTimer = intWaitTimer
    mblnAutoSize = blnAutoSize
    Set mobjFont = objFont

    Call InitForm(mfrmParent)
    
    '������ʾ��
    Call PlayWave(bytSound)
    
    mfrmAirBubbleMessage.Width = lngWidth
    mfrmAirBubbleMessage.Height = lngHeight
    
    mfrmAirBubbleMessage.ForeColor = lngForeColor
    
    '�򿪲�͸������
    Call ShowMessageOpaque

End Function

Private Function ShowMessageOpaque()
'�򿪲�͸�����ݴ���
    
    Dim lngScreenX As Long
    Dim lngScreenY As Long
    Dim lngCX As Long
    Dim lngCY As Long
    Dim lngScaleX As Long
    Dim lngScaleY As Long
    Dim lngTextWidth As Long
    Dim lngTextHeight As Long
    Dim lngTextHeightFlag
    Dim objPoint As POINTAPI
    Dim objRect As RECT
    Dim lngHwnd As Long
    Dim lngR As Long
    
    If Not mobjFont Is Nothing Then
        Set mfrmAirBubbleMessage.Font = mobjFont
    End If
    Call SeParantSize(mfrmAirBubbleMessage, mstrShowText)
    mlngMaxHeight = mfrmAirBubbleMessage.Height + 15
    Select Case mbytShowPostion
    Case SHOWPOSTION.OwnerCenter
        
        Call ClientToScreen(mfrmParent.hwnd, objPoint)
        
        lngScreenX = objPoint.X * Screen.TwipsPerPixelX + (mfrmParent.Width - mfrmAirBubbleMessage.Width) / 2
        lngScreenY = objPoint.Y * Screen.TwipsPerPixelY + (mfrmParent.Height - mfrmAirBubbleMessage.Height) / 2
        
        If mbytShowMode = SHOWMODE.�������� Then
            lngScreenY = objPoint.Y * Screen.TwipsPerPixelY + (mfrmParent.Height - mfrmAirBubbleMessage.Height) / 2 + mfrmAirBubbleMessage.Height
            mlngInitHeight = 0
        End If
    Case SHOWPOSTION.RightBottom
        lngHwnd = FindWindow("Shell_TrayWnd", vbNullString)
        Call GetClientRect(lngHwnd, objRect)

        lngScreenX = Screen.Width - mfrmAirBubbleMessage.Width
        lngScreenY = Screen.Height - objRect.Bottom * Screen.TwipsPerPixelY - mfrmAirBubbleMessage.Height
        
        If mbytShowMode = SHOWMODE.�������� Then
            lngScreenY = Screen.Height - objRect.Bottom * Screen.TwipsPerPixelY
            mlngInitHeight = 0
        End If
    Case SHOWPOSTION.ScreenCenter
        lngScreenX = (Screen.Width - mfrmAirBubbleMessage.Width) / 2
        lngScreenY = (Screen.Height - mfrmAirBubbleMessage.Height) / 2
        
        If mbytShowMode = SHOWMODE.�������� Then
            lngScreenY = (Screen.Height - mfrmAirBubbleMessage.Height) / 2 + mfrmAirBubbleMessage.Height
            mlngInitHeight = 0
        End If
    End Select

    lngCX = mfrmAirBubbleMessage.Width
    lngCY = mlngInitHeight
    On Error Resume Next
    
    ShowWindow mfrmAirBubbleMessage.hwnd, 1
    SetWindowPos mfrmAirBubbleMessage.hwnd, HWND_TOPMOST, lngScreenX / 15, lngScreenY / 15, lngCX / 15, lngCY / 15, &H10 Or &H1

    mfrmParent.SetFocus
    Call mfrmAirBubbleMessage.SetMeLine(mlngBackColor1, mlngBackColor2)
    Select Case mbytShowMode
    Case SHOWMODE.ֱ�ӵ���
        lngR = GetWindowLong(mfrmAirBubbleMessage.hwnd, GWL_EXSTYLE)
        Call mfrmAirBubbleMessage.ShowContent(mstrShowText, mobjFont, GAPLEFT, GAPLEFT, GAPROW)
        Call SetWindowLong(mfrmAirBubbleMessage.hwnd, GWL_EXSTYLE, lngR Or WS_EX_LAYERED)
        SetLayeredWindowAttributes mfrmAirBubbleMessage.hwnd, 0, 255, LWA_ALPHA
    Case SHOWMODE.���뽥��
        lngR = GetWindowLong(mfrmAirBubbleMessage.hwnd, GWL_EXSTYLE)
        Call SetWindowLong(mfrmAirBubbleMessage.hwnd, GWL_EXSTYLE, lngR Or WS_EX_LAYERED)
        SetLayeredWindowAttributes mfrmAirBubbleMessage.hwnd, 0, 0, LWA_ALPHA
        Call mfrmAirBubbleMessage.ShowContent(mstrShowText, mobjFont, GAPLEFT, GAPLEFT, GAPROW)
    Case SHOWMODE.��������
        lngR = GetWindowLong(mfrmAirBubbleMessage.hwnd, GWL_EXSTYLE)
        Call SetWindowLong(mfrmAirBubbleMessage.hwnd, GWL_EXSTYLE, lngR Or WS_EX_LAYERED)
        SetLayeredWindowAttributes mfrmAirBubbleMessage.hwnd, 0, 255, LWA_ALPHA
        Call mfrmAirBubbleMessage.ShowContent(mstrShowText, mobjFont, GAPLEFT, GAPLEFT, GAPROW)
    End Select
    '��ʼ͸����
    mintPouse = 0
    tmrOpen.Enabled = True
End Function

Private Function ShowMessageTransparence()
'��͸�����ݴ���
    
    Dim lngScreenX As Long
    Dim lngScreenY As Long
    Dim lngCX As Long
    Dim lngCY As Long
    Dim lngScaleX As Long
    Dim lngScaleY As Long
    Dim lngTextWidth As Long
    Dim lngTextHeight As Long
    Dim lngTextHeightFlag
    Dim objPoint As POINTAPI
    Dim objRect As RECT
    Dim lngHwnd As Long
    
    If Not mobjFont Is Nothing Then
        Set mfrmAirBubbleTransparence.Font = mobjFont
        Set mfrmAirBubbleTransparenceMessage.Font = mobjFont
    End If
    Call SeParantSize(mfrmAirBubbleTransparence, mstrShowText)
    Call SeParantSize(mfrmAirBubbleTransparenceMessage, mstrShowText)
    mlngMaxHeight = mfrmAirBubbleTransparence.Height + 15
    Select Case mbytShowPostion
    Case SHOWPOSTION.OwnerCenter
        
        Call ClientToScreen(mfrmParent.hwnd, objPoint)
        
        lngScreenX = objPoint.X * Screen.TwipsPerPixelX + (mfrmParent.Width - mfrmAirBubbleTransparence.Width) / 2
        lngScreenY = objPoint.Y * Screen.TwipsPerPixelY + (mfrmParent.Height - mfrmAirBubbleTransparence.Height) / 2

        If mbytShowMode = SHOWMODE.�������� Then
            lngScreenY = objPoint.Y * Screen.TwipsPerPixelY + (mfrmParent.Height - mfrmAirBubbleTransparence.Height) / 2 + mfrmAirBubbleTransparence.Height
            mfrmAirBubbleTransparence.Height = 0
        End If
    Case SHOWPOSTION.RightBottom
    
        lngHwnd = FindWindow("Shell_TrayWnd", vbNullString)
        Call GetClientRect(lngHwnd, objRect)

        lngScreenX = Screen.Width - mfrmAirBubbleTransparence.Width
        lngScreenY = Screen.Height - objRect.Bottom * Screen.TwipsPerPixelY - mfrmAirBubbleTransparence.Height
        
        If mbytShowMode = SHOWMODE.�������� Then
            lngScreenY = Screen.Height - objRect.Bottom * Screen.TwipsPerPixelY
            mfrmAirBubbleTransparence.Height = 0
        End If
    Case SHOWPOSTION.ScreenCenter
        lngScreenX = (Screen.Width - mfrmAirBubbleTransparence.Width) / 2
        lngScreenY = (Screen.Height - mfrmAirBubbleTransparence.Height) / 2
        
        If mbytShowMode = SHOWMODE.�������� Then
            lngScreenY = (Screen.Height - mfrmAirBubbleTransparence.Height) / 2 + mfrmAirBubbleTransparence.Height
            mfrmAirBubbleTransparence.Height = 0
        End If
    End Select

    lngCX = mfrmAirBubbleTransparence.Width
    lngCY = mfrmAirBubbleTransparence.Height

    On Error Resume Next
    
    ShowWindow mfrmAirBubbleTransparence.hwnd, 1
    SetWindowPos mfrmAirBubbleTransparence.hwnd, HWND_TOPMOST, lngScreenX / 15, lngScreenY / 15, lngCX / 15, lngCY / 15, &H10 Or &H1

    ShowWindow mfrmAirBubbleTransparenceMessage.hwnd, 1
    SetWindowPos mfrmAirBubbleTransparenceMessage.hwnd, HWND_TOPMOST, lngScreenX / 15, lngScreenY / 15, lngCX / 15, lngCY / 15, &H10 Or &H1
    
    mfrmParent.SetFocus

    Select Case mbytShowMode
    Case SHOWMODE.ֱ�ӵ���
        Call mfrmAirBubbleTransparence.SetMeLine
        mfrmAirBubbleTransparence.TransparentGrade = msglTransparentGrade
        mfrmAirBubbleTransparenceMessage.TransparentGrade = 100
        Call mfrmAirBubbleTransparenceMessage.ShowContent(mstrShowText, mobjFont, GAPLEFT, GAPLEFT, GAPROW)
        tmrOpen.Enabled = True
    Case SHOWMODE.���뽥��
        Call mfrmAirBubbleTransparence.SetMeLine
        Call mfrmAirBubbleTransparenceMessage.SetMeShape
        mfrmAirBubbleTransparence.TransparentGrade = 100
        mfrmAirBubbleTransparenceMessage.TransparentGrade = 100
        Call mfrmAirBubbleTransparence.ShowContent(mstrShowText, mobjFont, GAPLEFT, GAPLEFT, GAPROW)
        mintPouse = 0
        tmrOpen.Enabled = True
    Case SHOWMODE.��������
        mfrmAirBubbleTransparence.TransparentGrade = msglTransparentGrade
        mfrmAirBubbleTransparenceMessage.TransparentGrade = 100
        mfrmAirBubbleTransparenceMessage.Cls
        Call mfrmAirBubbleTransparenceMessage.ShowContent(mstrShowText, mobjFont, GAPLEFT, GAPLEFT, GAPROW)
        mintPouse = 0
        tmrOpen.Enabled = True
    End Select
End Function

Public Function CloseAirBubble()
    On Error Resume Next
    mintGrade = 0
    If Not (mfrmAirBubbleMessage Is Nothing) Then
        Unload mfrmAirBubbleMessage
    End If
    
    If Not (mfrmAirBubbleTransparence Is Nothing) Then
        Unload mfrmAirBubbleTransparence
    End If
    
    If Not (mfrmAirBubbleTransparenceMessage Is Nothing) Then
        Unload mfrmAirBubbleTransparenceMessage
    End If
    If Dir(mstrSoundTmp) <> "" Then
        Kill mstrSoundTmp
    End If
End Function

Private Function SeParantSize(ByVal objMain As Object, ByVal strText As String)
    Dim lngR As Long
    Dim arrTxt As Variant
    Dim i As Integer
    Dim j As Integer
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngRow As Long
    
    If mblnAutoSize And strText <> "" Then
        arrTxt = Split(strText, vbCrLf)
        lngWidth = objMain.TextWidth(arrTxt(0))
        If UBound(arrTxt) > 0 Then
            For i = 0 To UBound(arrTxt)
                If objMain.TextWidth(arrTxt(i)) > lngWidth Then lngWidth = objMain.TextWidth(arrTxt(i))
                If objMain.TextWidth(arrTxt(i)) > objMain.Width - GAPLEFT * 2 Then
                    lngRow = lngRow + (0 - (Int(objMain.TextWidth(arrTxt(i)) / (objMain.Width - GAPLEFT * 2) * (-1))))
                Else
                    lngRow = lngRow + 1
                End If
            Next
        Else
            lngRow = (0 - (Int(objMain.TextWidth(arrTxt(i)) / (objMain.Width - GAPLEFT * 2) * (-1))))
        End If
        If lngWidth + GAPLEFT * 2 <= objMain.Width Then
            objMain.Width = lngWidth + GAPLEFT * 2
        End If
        objMain.Height = (objMain.TextHeight(arrTxt(0)) + GAPROW) * lngRow - GAPROW + GAPHEIGHT * 2
    Else
        arrTxt = Split(strText, vbCrLf)
        If UBound(arrTxt) > 0 Then
            For i = 0 To UBound(arrTxt)
                If objMain.TextWidth(arrTxt(i)) > objMain.Width - GAPLEFT * 2 Then
                    lngRow = lngRow + (0 - (Int(objMain.TextWidth(arrTxt(i)) / (objMain.Width - GAPLEFT * 2) * (-1))))
                Else
                    lngRow = lngRow + 1
                End If
            Next
        Else
            lngRow = (0 - (Int(objMain.TextWidth(arrTxt(i)) / (objMain.Width - GAPLEFT * 2) * (-1))))
        End If
        If (objMain.TextHeight(arrTxt(0)) + GAPROW) * lngRow - GAPROW + GAPHEIGHT * 2 > objMain.Height Then
            objMain.Height = (objMain.TextHeight(arrTxt(0)) + GAPROW) * lngRow - GAPROW + GAPHEIGHT * 2
        End If
    End If
End Function

Private Sub Class_Initialize()
    Set mfrmAirBubbleMessage = New frmAirBubbleMessage
    Set mfrmAirBubbleTransparence = New frmAirBubbleTransparence
    Set mfrmAirBubbleTransparenceMessage = New frmAirBubbleTransparenceMessage
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    If Not (mfrmAirBubbleMessage Is Nothing) Then
        Unload mfrmAirBubbleMessage
    End If
    If Not (mfrmAirBubbleTransparence Is Nothing) Then
        Unload mfrmAirBubbleTransparence
    End If
    If Not (mfrmAirBubbleTransparenceMessage Is Nothing) Then
        Unload mfrmAirBubbleTransparenceMessage
    End If
    If Dir(mstrSoundTmp) <> "" Then
        Kill mstrSoundTmp
    End If
End Sub

Private Sub tmrAlert_Timer()
    If mintWaitTimer = 0 Then
        tmrAlert.Enabled = False
        Exit Sub
    End If
    mintPouse = mintPouse + 1
    If mintPouse > mintWaitTimer Then
        tmrAlert.Enabled = False
        tmrClose.Enabled = True
        mintPouse = 0
    End If
    
End Sub

Private Sub tmrClose_Timer()
    Dim lngHeight As Long
    Dim lngNewHeight As Long
    Dim lngScaleY As Long
    Dim lngH As Long
    Dim lngTop As Long
    
    Select Case mbytShowMode
    Case SHOWMODE.ֱ�ӵ���
        tmrClose.Enabled = False
        Unload mfrmAirBubbleTransparence
        Unload mfrmAirBubbleTransparenceMessage
        Unload mfrmAirBubbleMessage
    Case SHOWMODE.���뽥��
        Select Case mbytShowStyle
        Case 0  '͸������
            If mintGrade < 1 Then
                tmrClose.Enabled = False
                Unload mfrmAirBubbleTransparence
                Unload mfrmAirBubbleTransparenceMessage
            Else
                mfrmAirBubbleTransparenceMessage.Cls
                SetLayeredWindowAttributes mfrmAirBubbleTransparence.hwnd, 0, mintGrade, LWA_ALPHA
                mintGrade = mintGrade - 1
            End If
        Case Else  '��͸������
            If mintGrade < 1 Then
                tmrClose.Enabled = False
                Unload mfrmAirBubbleMessage
            Else
                SetLayeredWindowAttributes mfrmAirBubbleMessage.hwnd, 0, mintGrade, LWA_ALPHA
                mintGrade = mintGrade - 3
            End If
        End Select
    Case SHOWMODE.��������
        If mbytShowStyle = 0 Then
            lngHeight = mfrmAirBubbleTransparence.Height
            If lngHeight > 60 Then

                lngHeight = lngHeight - 30
                lngTop = mfrmAirBubbleTransparence.Top + 30
    
                On Error Resume Next
    
                MoveWindow mfrmAirBubbleTransparence.hwnd, mfrmAirBubbleTransparence.Left / 15, lngTop / 15, mfrmAirBubbleTransparence.Width / 15, lngHeight / 15, 1
                SetWindowPos mfrmAirBubbleTransparence.hwnd, -1, mfrmAirBubbleTransparence.Left / 15, lngTop / 15, mfrmAirBubbleTransparence.Width / 15, lngHeight / 15, &H10 Or &H1
        
                MoveWindow mfrmAirBubbleTransparenceMessage.hwnd, mfrmAirBubbleTransparenceMessage.Left / 15, lngTop / 15, mfrmAirBubbleTransparenceMessage.Width / 15, lngHeight / 15, 1
                SetWindowPos mfrmAirBubbleTransparenceMessage.hwnd, -1, mfrmAirBubbleTransparenceMessage.Left / 15, lngTop / 15, mfrmAirBubbleTransparenceMessage.Width / 15, lngHeight / 15, &H10 Or &H1

            Else
                tmrClose.Enabled = False
                Unload mfrmAirBubbleTransparence
                Unload mfrmAirBubbleTransparenceMessage
            End If
        Else
            lngHeight = mfrmAirBubbleMessage.Height
            If lngHeight > 60 Then

                lngHeight = lngHeight - 30
                lngTop = mfrmAirBubbleMessage.Top + 30
    
                On Error Resume Next
    
                MoveWindow mfrmAirBubbleMessage.hwnd, mfrmAirBubbleMessage.Left / 15, lngTop / 15, mfrmAirBubbleMessage.Width / 15, lngHeight / 15, 1
                SetWindowPos mfrmAirBubbleMessage.hwnd, -1, mfrmAirBubbleMessage.Left / 15, lngTop / 15, mfrmAirBubbleMessage.Width / 15, lngHeight / 15, &H10 Or &H1
            Else
                tmrClose.Enabled = False
                Unload mfrmAirBubbleMessage
            End If
        End If
        mintGrade = 0
    End Select
End Sub

Private Sub tmrOpen_Timer()
   Dim lngScaleY As Long
    Dim lngNewHeight As Long
     Dim lngHeight As Long
    
    
    Dim i As Integer

    Dim lngH As Long
    Dim lngTop As Long

    Select Case mbytShowMode
    Case SHOWMODE.ֱ�ӵ���
        tmrOpen.Enabled = False
        tmrAlert.Enabled = True
    Case SHOWMODE.���뽥��
        Select Case mbytShowStyle
        Case 0 '͸������
            If mintGrade > (100 - msglTransparentGrade) * (255 / 100) Then
                If mintGrade > 255 Then mintGrade = 255
                Call mfrmAirBubbleTransparenceMessage.ShowContent(mstrShowText, mobjFont, GAPLEFT, GAPLEFT, GAPROW)
                SetLayeredWindowAttributes mfrmAirBubbleTransparenceMessage.hwnd, mfrmAirBubbleTransparenceMessage.BackColor, mintGrade, LWA_COLORKEY
                tmrOpen.Enabled = False
                tmrAlert.Enabled = True
                Exit Sub
            End If
            SetLayeredWindowAttributes mfrmAirBubbleTransparence.hwnd, 0, mintGrade, LWA_ALPHA
            mintGrade = mintGrade + 1
        Case 1 '��͸������
            If mintGrade > 255 Then
                mintGrade = 255
                SetLayeredWindowAttributes mfrmAirBubbleMessage.hwnd, 0, 255, LWA_ALPHA
                tmrOpen.Enabled = False
                tmrAlert.Enabled = True
                Exit Sub
            End If
            SetLayeredWindowAttributes mfrmAirBubbleMessage.hwnd, 0, mintGrade, LWA_ALPHA
            Call SendMessage(mfrmAirBubbleMessage.hwnd, WM_PAINT, 0, 0) '��һ����ʾʱ������Ч��
            mintGrade = mintGrade + 3
        End Select
    Case 3
        If mbytShowStyle = 0 Then
            lngScaleY = mfrmAirBubbleTransparence.Height - mfrmAirBubbleTransparence.ScaleHeight
            lngHeight = mfrmAirBubbleTransparence.Height

            If mfrmAirBubbleTransparence.Height < mlngMaxHeight + lngScaleY Then
                lngNewHeight = lngHeight + 30

                If lngNewHeight > mlngMaxHeight + lngScaleY Then lngNewHeight = mlngMaxHeight + lngScaleY

                lngH = mfrmAirBubbleTransparence.Height + (lngNewHeight - lngHeight)
                lngTop = mfrmAirBubbleTransparence.Top - (lngNewHeight - lngHeight)

                On Error Resume Next

                MoveWindow mfrmAirBubbleTransparence.hwnd, mfrmAirBubbleTransparence.Left / 15, lngTop / 15, mfrmAirBubbleTransparence.Width / 15, lngH / 15, 1
                SetWindowPos mfrmAirBubbleTransparence.hwnd, -1, mfrmAirBubbleTransparence.Left / 15, lngTop / 15, mfrmAirBubbleTransparence.Width / 15, lngH / 15, &H10 Or &H1

                MoveWindow mfrmAirBubbleTransparenceMessage.hwnd, mfrmAirBubbleTransparenceMessage.Left / 15, lngTop / 15, mfrmAirBubbleTransparenceMessage.Width / 15, lngH / 15, 1
                SetWindowPos mfrmAirBubbleTransparenceMessage.hwnd, -1, mfrmAirBubbleTransparenceMessage.Left / 15, lngTop / 15, mfrmAirBubbleTransparenceMessage.Width / 15, lngH / 15, &H10 Or &H1

            Else
                tmrOpen.Enabled = False
                tmrAlert.Enabled = True
            End If
        Else
            lngScaleY = 0
            lngHeight = mlngInitHeight

            If mlngInitHeight < mlngMaxHeight + lngScaleY Then
                lngNewHeight = lngHeight + 30
                
                If lngNewHeight > mlngMaxHeight + lngScaleY Then lngNewHeight = mlngMaxHeight + lngScaleY

                lngH = mlngInitHeight + (lngNewHeight - lngHeight)
                lngTop = mfrmAirBubbleMessage.Top - (lngNewHeight - lngHeight)
                mlngInitHeight = lngNewHeight
                On Error Resume Next

                MoveWindow mfrmAirBubbleMessage.hwnd, mfrmAirBubbleMessage.Left / 15, lngTop / 15, mfrmAirBubbleMessage.Width / 15, lngH / 15, 1
                SetWindowPos mfrmAirBubbleMessage.hwnd, -1, mfrmAirBubbleMessage.Left / 15, lngTop / 15, mfrmAirBubbleMessage.Width / 15, lngH / 15, &H10 Or &H1
            Else
                tmrOpen.Enabled = False
                tmrAlert.Enabled = True
            End If
        End If
    End Select
    
End Sub

Private Function MoveAirWindow(ByVal objMain As Object)
    Dim lngScaleY As Long
    Dim lngNewHeight As Long
     Dim lngHeight As Long
    Dim lngH As Long
    Dim lngTop As Long
    lngScaleY = objMain.Height - objMain.ScaleHeight
    lngHeight = objMain.Height
    
    If objMain.Height < mlngMaxHeight + lngScaleY Then
        lngNewHeight = lngHeight + 30

        If lngNewHeight > mlngMaxHeight + lngScaleY Then lngNewHeight = mlngMaxHeight + lngScaleY

        lngH = objMain.Height + (lngNewHeight - lngHeight)
        lngTop = objMain.Top - (lngNewHeight - lngHeight)
        
        On Error Resume Next
        
        MoveWindow objMain.hwnd, objMain.Left / 15, lngTop / 15, objMain.Width / 15, lngH / 15, 1
        SetWindowPos objMain.hwnd, -1, objMain.Left / 15, lngTop / 15, objMain.Width / 15, lngH / 15, &H10 Or &H1

    End If
End Function

Private Function PlayWave(lngKey As Long) As String
    '����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
    '����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
    '����:�����ļ���
    
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255
    Dim strR As String
    
    On Error Resume Next
    
    arrData = LoadResData(lngKey, "WAVE")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    
    Call sndPlaySound(strR, SND_NODEFAULT Or SND_ASYNC)
        If Dir(mstrSoundTmp) <> "" Then
        Kill mstrSoundTmp
    End If
    Kill strR
    mstrSoundTmp = strR
End Function