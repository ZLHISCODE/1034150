VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDicomPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''DICOM��ӡʹ�õı�����
''''�ƽ�
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public strname As String
Public strIPAddress As String
Public lngPort As Long
Public strAETitle As String
Public strFormat As String
Public strPriority As String
Public lngCopies As Long
Public strMedium As String
Public strOrientation As String
Public strFilmSize As String
Public strFilmBox As String
Public strResolution As String
Public strMagnification As String
Public strSmooth As String
Public strTrim As String
Public strMinDensity As String
Public strMaxDensity As String
Public strEmptyDensity As String
Public strBorderDensity As String
Public strPolarity As String
Public lngID As Long
Public lngBitDepth As Long
Public strSCUAETitle As String  'SCU��AE����
Public lngImageBorderWidth As Long  'ͼ��ı߿�
Public intImageResolution           'ͼ��ķֱ��ʣ���λ��PPI���Ƽ�ֵ 300

