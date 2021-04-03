VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVfwCaptureParameterCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'��Ƶ��ʾ����
Public Enum TShowWay
  swNormal = 0 'ԭ��С��ʾ
  swStretch = 1 '������ʾ
  swFit = 2  '�Զ���Ӧ
  swAutoFitCut = 3 '���ü���Χ��ʾ
  swWindowAutoFit = 4 '����Ӧ��Ƶ��С
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'�ɼ���ز�������
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mlngCaptureDeviceIndex            '�ɼ��豸����
Private mlngCapBit As Long                '�ɼ�λ������
Private mlngCapVideoWidth As Long         '�ɼ���Ƶ����
Private mlngCapVideoHeight As Long        '�ɼ���Ƶ�߶�
Private mlngCapPreviewRate As Long        '�ɼ�Ԥ���ٶ�
Private mblnIsScale As Boolean            '�Ƿ��Զ���Ӧ
Private mblnCaptureAudio As Boolean       '�Ƿ�ɼ���Ƶ
Private mblnAbortLeftMouse As Boolean     '�Ƿ������ֹ�ɼ�
Private mblnAbortRightMouse As Boolean    '�Ƿ��Ҽ���ֹ�ɼ�
Private mblnLimitEnabled As Boolean      '�Ƿ�����ʱ������
Private mlngAVStreamMaster As Long        '��Ƶͬ����ʽ
Private mlngRequestMicroSecPerFrame As Long '�ɼ�֡Ƶ��
Private mlngVideoShowWay As TShowWay        '��Ƶ��ʾ��ʽ (0�� ԭ��С��ʾ,1��������ʾ,2�����������ţ�3�����ü���Χ���ţ�4������Ӧ��Ƶ��С)

Private mdblLeftCutRate As Double        '��ü�����
Private mdblTopCutRate As Double       '
Private mdblWidthCutRate As Double       '���Ȳü�����
Private mdblHeightCutRate As Double      '

Private Const M_LNG_DEFAULT_MAX_CAPDEV_COUNT As Long = 20  'Ĭ�����ɼ��豸����
Private Const M_LNG_DEFAULT_CAPBIT As Long = 24            'Ĭ�ϲɼ�λ��

Private Const M_STR_DEFAULT_CONFIG_NAME As String = "VfwCapture.ini" 'Ĭ�ϲɼ������������ļ�����




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'����Ĭ�ϲ�������
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadDefaultCfg()
  mlngCaptureDeviceIndex = 0
  mlngCapBit = M_LNG_DEFAULT_CAPBIT
  mlngCapVideoWidth = 320
  mlngCapVideoHeight = 240
  mlngCapPreviewRate = 66
  mblnIsScale = True
  mblnCaptureAudio = False
  mblnAbortLeftMouse = True
  mblnAbortRightMouse = True
  mblnLimitEnabled = True
  mlngAVStreamMaster = AVSTREAMMASTER_NONE
  mlngRequestMicroSecPerFrame = 15
  mlngVideoShowWay = 0
  
  mdblHeightCutRate = 0
  mdblWidthCutRate = 0
  mdblTopCutRate = 0
  mdblLeftCutRate = 0
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'��¡����
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CloneTo(ByRef curCaptureParameter As clsVfwCaptureParameterCfg)
  curCaptureParameter.AbortLeftMouse = Me.AbortLeftMouse
  curCaptureParameter.AbortRightMouse = Me.AbortRightMouse
  curCaptureParameter.AVStreamMaster = Me.AVStreamMaster
  curCaptureParameter.CapBit = Me.CapBit
  curCaptureParameter.CapPreviewRate = Me.CapPreviewRate
  curCaptureParameter.CaptureAudio = Me.CaptureAudio
  curCaptureParameter.CaptureDeviceIndex = Me.CaptureDeviceIndex
  curCaptureParameter.CapVideoHeight = Me.CapVideoHeight
  curCaptureParameter.CapVideoWidth = Me.CapVideoWidth
  curCaptureParameter.IsScale = Me.IsScale
  curCaptureParameter.LimitEnabled = Me.LimitEnabled
  curCaptureParameter.RequestMicroSecPerFrame = Me.RequestMicroSecPerFrame
  curCaptureParameter.VideoShowWay = Me.VideoShowWay
  
  curCaptureParameter.LeftCutRate = Me.LeftCutRate
  curCaptureParameter.TopCutRate = Me.TopCutRate
  curCaptureParameter.WidthCutRate = Me.WidthCutRate
  curCaptureParameter.HeightCutRate = Me.HeightCutRate
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'�����������
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveParameterCfg(ByVal sFileName As String)
  Dim strCurrentFile As String
  Dim iniFile As clsIniFile
  
  strCurrentFile = sFileName
  
  If Trim(strCurrentFile) = "" Then
    strCurrentFile = App.Path & "\" & M_STR_DEFAULT_CONFIG_NAME
  End If
  
  Set iniFile = New clsIniFile
  
  Call iniFile.SetIniFile(strCurrentFile)
  
  Call iniFile.WriteValue("VfwCaptureConfig", "AbortLeftMouse", CStr(Me.AbortLeftMouse))
  Call iniFile.WriteValue("VfwCaptureConfig", "AbortRightMouse", CStr(Me.AbortRightMouse))
  Call iniFile.WriteValue("VfwCaptureConfig", "AVStreamMaster", CStr(Me.AVStreamMaster))
  Call iniFile.WriteValue("VfwCaptureConfig", "CapBit", CStr(Me.CapBit))
  Call iniFile.WriteValue("VfwCaptureConfig", "CapPreviewRate", CStr(Me.CapPreviewRate))
  Call iniFile.WriteValue("VfwCaptureConfig", "CaptureAudio", CStr(Me.CaptureAudio))
  Call iniFile.WriteValue("VfwCaptureConfig", "CaptureDeviceIndex", CStr(Me.CaptureDeviceIndex))
  Call iniFile.WriteValue("VfwCaptureConfig", "CapVideoHeight", CStr(Me.CapVideoHeight))
  Call iniFile.WriteValue("VfwCaptureConfig", "CapVideoWidth", CStr(Me.CapVideoWidth))
  Call iniFile.WriteValue("VfwCaptureConfig", "IsScale", CStr(Me.IsScale))
  Call iniFile.WriteValue("VfwCaptureConfig", "LimitEnabled", CStr(Me.LimitEnabled))
  Call iniFile.WriteValue("VfwCaptureConfig", "RequestMicroSecPerFrame", CStr(Me.RequestMicroSecPerFrame))
  Call iniFile.WriteValue("VfwCaptureConfig", "VideoShowWay", CStr(Me.VideoShowWay))
  
  Call iniFile.WriteValue("VfwCaptureConfig", "LeftCutRate", CStr(Me.LeftCutRate))
  Call iniFile.WriteValue("VfwCaptureConfig", "TopCutRate", CStr(Me.TopCutRate))
  Call iniFile.WriteValue("VfwCaptureConfig", "WidthCutRate", CStr(Me.WidthCutRate))
  Call iniFile.WriteValue("VfwCaptureConfig", "HeightCutRate", CStr(Me.HeightCutRate))
  
  Set iniFile = Nothing
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'��ȡ��������
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReadParameterCfg(ByVal sFileName As String)
  Dim strCurrentFile As String
  Dim iniFile As clsIniFile
  
  strCurrentFile = sFileName
  
  If Trim(strCurrentFile) = "" Then
    strCurrentFile = App.Path & "\" & M_STR_DEFAULT_CONFIG_NAME
  End If
  
  Set iniFile = New clsIniFile
  
  Call iniFile.SetIniFile(strCurrentFile)
  
  Me.AbortLeftMouse = CBool(iniFile.ReadValue("VfwCaptureConfig", "AbortLeftMouse", CStr(True)))
  Me.AbortRightMouse = CBool(iniFile.ReadValue("VfwCaptureConfig", "AbortRightMouse", CStr(True)))
  Me.AVStreamMaster = CLng(iniFile.ReadValue("VfwCaptureConfig", "AVStreamMaster", CStr(0)))
  Me.CapBit = CLng(iniFile.ReadValue("VfwCaptureConfig", "CapBit", CStr(24)))
  Me.CapPreviewRate = CLng(iniFile.ReadValue("VfwCaptureConfig", "CapPreviewRate", CStr(66)))
  Me.CaptureAudio = CBool(iniFile.ReadValue("VfwCaptureConfig", "CaptureAudio", CStr(False)))
  Me.CaptureDeviceIndex = CLng(iniFile.ReadValue("VfwCaptureConfig", "CaptureDeviceIndex", CStr(0)))
  Me.CapVideoHeight = CLng(iniFile.ReadValue("VfwCaptureConfig", "CapVideoHeight", CStr(240)))
  Me.CapVideoWidth = CLng(iniFile.ReadValue("VfwCaptureConfig", "CapVideoWidth", CStr(320)))
  Me.IsScale = CBool(iniFile.ReadValue("VfwCaptureConfig", "IsScale", CStr(True)))
  Me.LimitEnabled = CBool(iniFile.ReadValue("VfwCaptureConfig", "LimitEnabled", CStr(False)))
  Me.RequestMicroSecPerFrame = CLng(iniFile.ReadValue("VfwCaptureConfig", "RequestMicroSecPerFrame", CStr(15)))
  Me.VideoShowWay = CLng(iniFile.ReadValue("VfwCaptureConfig", "VideoShowWay", CStr(0)))
  
  Me.LeftCutRate = CDbl(iniFile.ReadValue("VfwCaptureConfig", "LeftCutRate", CStr(0)))
  Me.TopCutRate = CDbl(iniFile.ReadValue("VfwCaptureConfig", "TopCutRate", CStr(0)))
  Me.WidthCutRate = CDbl(iniFile.ReadValue("VfwCaptureConfig", "WidthCutRate", CStr(0)))
  Me.HeightCutRate = CDbl(iniFile.ReadValue("VfwCaptureConfig", "HeightCutRate", CStr(0)))
  
  Set iniFile = Nothing
End Sub

'//�ɼ��豸����
Property Get CaptureDeviceIndex() As Long
  CaptureDeviceIndex = mlngCaptureDeviceIndex
End Property

Property Let CaptureDeviceIndex(value As Long)
  mlngCaptureDeviceIndex = value
End Property



'//��ɫ���
Property Get CapBit() As Long
  CapBit = mlngCapBit
End Property

Property Let CapBit(value As Long)
  mlngCapBit = value
End Property


'//ȡ����Ƶ����
Property Get CapVideoWidth() As Long
  CapVideoWidth = mlngCapVideoWidth
End Property

Property Let CapVideoWidth(value As Long)
  mlngCapVideoWidth = value
End Property


'//ȡ����Ƶ�߶�
Property Get CapVideoHeight() As Long
  CapVideoHeight = mlngCapVideoHeight
End Property

Property Let CapVideoHeight(value As Long)
  mlngCapVideoHeight = value
End Property


'//ȡ��Ԥ���ٶ�
Property Get CapPreviewRate() As Long
  CapPreviewRate = mlngCapPreviewRate
End Property

Property Let CapPreviewRate(value As Long)
  mlngCapPreviewRate = value
End Property


'//�Ƿ��Զ���Ӧ
Property Get IsScale() As Boolean
  IsScale = mblnIsScale
End Property

Property Let IsScale(value As Boolean)
  mblnIsScale = value
End Property


'//�Ƿ�ɼ���Ƶ
Property Get CaptureAudio() As Boolean
  CaptureAudio = mblnCaptureAudio
End Property

Property Let CaptureAudio(value As Boolean)
  mblnCaptureAudio = value
End Property


'//����Ƿ���ֹ�ɼ�
Property Get AbortLeftMouse() As Boolean
  AbortLeftMouse = mblnAbortLeftMouse
End Property

Property Let AbortLeftMouse(value As Boolean)
  mblnAbortLeftMouse = value
End Property


'//�Ҽ��Ƿ���ֹ�ɼ�
Property Get AbortRightMouse() As Boolean
  AbortRightMouse = mblnAbortRightMouse
End Property

Property Let AbortRightMouse(value As Boolean)
  mblnAbortRightMouse = value
End Property
  
  
'//�Ƿ���ʱ������
Property Get LimitEnabled() As Boolean
  LimitEnabled = mblnLimitEnabled
End Property

Property Let LimitEnabled(value As Boolean)
  mblnLimitEnabled = value
End Property

  
'//��Ƶͬ����ʽ
Property Get AVStreamMaster() As Long
  AVStreamMaster = mlngAVStreamMaster
End Property

Property Let AVStreamMaster(value As Long)
  mlngAVStreamMaster = value
End Property

  
'//�ɼ�֡Ƶ��
Property Get RequestMicroSecPerFrame() As Long
  RequestMicroSecPerFrame = mlngRequestMicroSecPerFrame
End Property

Property Let RequestMicroSecPerFrame(value As Long)
  mlngRequestMicroSecPerFrame = value
End Property


'��Ƶ��ʾ��ʽ
Property Get VideoShowWay() As Long
  VideoShowWay = mlngVideoShowWay
End Property

Property Let VideoShowWay(value As Long)
  mlngVideoShowWay = value
End Property


'ͼ��ü�����
Property Get HeightCutRate() As Double
  HeightCutRate = mdblHeightCutRate
End Property
   
Property Let HeightCutRate(value As Double)
  mdblHeightCutRate = value
End Property

   
Property Get WidthCutRate() As Double
  WidthCutRate = mdblWidthCutRate
End Property

Property Let WidthCutRate(value As Double)
  mdblWidthCutRate = value
End Property


Property Get TopCutRate() As Double
  TopCutRate = mdblTopCutRate
End Property

Property Let TopCutRate(value As Double)
  mdblTopCutRate = value
End Property


Property Get LeftCutRate() As Double
  LeftCutRate = mdblLeftCutRate
End Property

Property Let LeftCutRate(value As Double)
  mdblLeftCutRate = value
End Property