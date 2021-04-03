VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPictureEditor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mfMain As fMain
Attribute mfMain.VB_VarHelpID = -1
Public lngKeyOfPic As Long                      'ͼƬ��Keyֵ

Public Event pOK(ByRef FinalPicture As StdPicture, ByVal lngWidth As Long, ByVal lngHeight As Long)    '���棬�����޸ĺ����ʱͼƬ·����JPEGͼƬ��
Public Event pCancel()                          'ȡ�����˳�

'################################################################################################################
'## ���ܣ�  ��ʾ�༭������
'##
'## ������  lngSys      :ϵͳ��
'##         cnMain      :���ݿ�����
'##         srcPic      :ԴͼƬ StcPicture
'##         lngKey      :ͼƬKeyֵ
'##         frmParent   :������
'##         bln����     :��ͼƬ�����Ƿ���������ǣ��������༭ʱ������ͼƬ
'##         blnModeless :�Ƿ��Ƿ�ģ̬��Ĭ��Ϊ��ģ̬
'################################################################################################################
Public Sub ShowPicEditor(ByVal lngSys As Long, _
    ByRef cnMain As ADODB.Connection, _
    ByRef srcPic As StdPicture, _
    Optional lngKey As Long = 0, _
    Optional bln���� As Boolean, _
    Optional ByRef frmParent As Object = Nothing, _
    Optional ByVal blnModeless As Boolean = True)
    
    Call InitCommon(cnMain)
    glngSys = lngSys
    lngKeyOfPic = lngKey
    gbln���� = bln����
    
    If mfMain Is Nothing Then Set mfMain = New fMain
    Set gfrmMain = mfMain
    
    If gfDialogEx Is Nothing Then Set gfDialogEx = New fDialogEx
    If gfFilter Is Nothing Then Set gfFilter = New fFilter
    If gfOrientation Is Nothing Then Set gfOrientation = New fOrientation
    If gfPanView Is Nothing Then Set gfPanView = New fPanView
    If gfPrint Is Nothing Then Set gfPrint = New fPrint
    If gfProperties Is Nothing Then Set gfProperties = New fProperties
    If gfResize Is Nothing Then Set gfResize = New fResize
    If gfTexturize Is Nothing Then Set gfTexturize = New fTexturize
    
    Call gfrmMain.ShowMe(srcPic, frmParent, blnModeless)
End Sub

Private Sub Class_Initialize()
    If mfMain Is Nothing Then Set mfMain = New fMain
    Set gfrmMain = mfMain
End Sub

'################################################################################################################
'## ���ܣ�  �ͷ���Դ
'################################################################################################################
Private Sub Class_Terminate()
    On Error Resume Next
    Unload gfrmMain
    Unload gfDialogEx
    Unload gfFilter
    Unload gfOrientation
    Unload gfPanView
    Unload gfPrint
    Unload gfProperties
    Unload gfResize
    Unload gfTexturize
    Unload mfMain
    
    Set gfrmMain = Nothing
    Set gfDialogEx = Nothing
    Set gfFilter = Nothing
    Set gfOrientation = Nothing
    Set gfPanView = Nothing
    Set gfPrint = Nothing
    Set gfProperties = Nothing
    Set gfResize = Nothing
    Set gfTexturize = Nothing
    Set mfMain = Nothing
End Sub

Private Sub mfMain_pCancel()
    RaiseEvent pCancel
End Sub

Private Sub mfMain_pOK(ByRef FinalPicture As StdPicture, ByVal lngWidth As Long, ByVal lngHeight As Long)
    RaiseEvent pOK(FinalPicture, lngWidth, lngHeight)
End Sub