VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRFileDefineInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�����ļ�����
Public ID As Long               'ID
Public ���� As EPRDocTypeEnum   '���ࣺ1-���ﲡ��;2-סԺ����;3-������¼;4-��������;5-�������;6-֪���ļ�;7-���Ʊ���;8-��������
Public ��� As Long             '�ļ����
Public ���� As String           '�ļ�����
Public ˵�� As String           '�ļ�˵��
Public ҳ�� As String           '�ļ�ҳ�棺ֻ��סԺ�����ͻ�����¼��������ҳ��,�������͵Ĳ���������ҳ,ҳ���ź��ļ����һ��
Public ���� As Long             '�Ƿ�����0-�û�����Ĳ����ļ�;1-ϵͳ��������ͨ�����ļ�,����ɾ��;<0,ָ�������Ƴ���ʵ�ֵ����ⲡ��,����ɾ��
Public ͨ�� As Long             'ͨ���ԣ�0-δ��;1-ͨ��;2-ר��
Public ҳ������ As String       '����ҳ������

Public ���� As Long             '�����Ƶ���:0-�ɱ༭,�����༭��ʽ���;1-�ޱ༭,���Զ��屨�����(�紦����);2-�ɱ༭,���Զ��屨�����(������鱨��)���Ի�����¼:0-���˼�¼;1-������¼��������Ч��
Public ��ʽ As String           'PaperKind;PaperOrient;PaperHeight;PaperWidth;MarginLeft;MarginRight;MarginTop;MarginBottom;BackColor;PaperColor;ShowPageNumber
Public ҳü As String
Public ҳ�� As String

Public lngModule As Long

Private HeadFontName As String
Private HeadFontSize As Long
Private HeadFontBold As Boolean
Private HeadFontItalic As Boolean
Private HeadFontUnderline As Boolean
Private HeadFontStrikethrough As Boolean
Private HeadFontColor As OLE_COLOR

Private FootFontName As String
Private FootFontSize As Long
Private FootFontBold As Boolean
Private FootFontItalic As Boolean
Private FootFontUnderline As Boolean
Private FootFontStrikethrough As Boolean
Private FootFontColor As OLE_COLOR

Public Property Let HeadFont(vData As String)
    On Error Resume Next
    Dim T As Variant
    T = Split(vData, "|")
    HeadFontName = T(0)
    HeadFontSize = T(1)
    HeadFontBold = T(2)
    HeadFontItalic = T(3)
    HeadFontUnderline = T(4)
    HeadFontStrikethrough = T(5)
    HeadFontColor = T(6)
End Property

Public Property Get HeadFont() As String
    HeadFont = HeadFontName & "|" & HeadFontSize & "|" & HeadFontBold & "|" & HeadFontItalic & "|" & HeadFontUnderline & "|" & HeadFontStrikethrough & "|" & HeadFontColor
End Property

Public Property Let FootFont(vData As String)
    On Error Resume Next
    Dim T As Variant
    T = Split(vData, "|")
    FootFontName = T(0)
    FootFontSize = T(1)
    FootFontBold = T(2)
    FootFontItalic = T(3)
    FootFontUnderline = T(4)
    FootFontStrikethrough = T(5)
    FootFontColor = T(6)
End Property

Public Property Get FootFont() As String
    FootFont = FootFontName & "|" & FootFontSize & "|" & FootFontBold & "|" & FootFontItalic & "|" & FootFontUnderline & "|" & FootFontStrikethrough & "|" & FootFontColor
End Property

Public Function GetFormat(edtThis As Object) As String
    Dim strF As String
    With edtThis
        strF = .PaperKind                               '0
        strF = strF & ";" & .PaperOrient                '1
        strF = strF & ";" & .PaperHeight                '2
        strF = strF & ";" & .PaperWidth                 '3
        strF = strF & ";" & .MarginLeft                 '4
        strF = strF & ";" & .MarginRight                '5
        strF = strF & ";" & .MarginTop                  '6
        strF = strF & ";" & .MarginBottom               '7
        strF = strF & ";" & .BackColor                  '8
        strF = strF & ";" & .PaperColor                 '9
        strF = strF & ";" & IIf(.ShowPageNumber, 1, 0)  '10
        strF = strF & ";" & .HeadFontName & "|" & .HeadFontSize & "|" & .HeadFontBold & "|" & .HeadFontItalic & "|" & .HeadFontUnderline & "|" & .HeadFontStrikethrough & "|" & .HeadFontColor
        strF = strF & ";" & .FootFontName & "|" & .FootFontSize & "|" & .FootFontBold & "|" & .FootFontItalic & "|" & .FootFontUnderline & "|" & .FootFontStrikethrough & "|" & .FootFontColor
    End With
    GetFormat = strF
End Function

Public Sub SetFormat(edtThis As Object, strFmt As String)
    Dim T As Variant, T2 As Variant
    T = Split(strFmt, ";")
    On Error Resume Next
    With edtThis
        .PaperHeight = T(2)
        .PaperWidth = T(3)
        If T(0) > 0 Then
            .PaperKind = T(0)
        End If
        .PaperOrient = T(1)
        .MarginLeft = T(4)
        .MarginRight = T(5)
        .MarginTop = T(6)
        .MarginBottom = T(7)
        .BackColor = T(8)
        .PaperColor = T(9)
        .ShowPageNumber = (T(10) = 1)
        T2 = Split(T(11), "|")
        .HeadFontName = T2(0)
        .HeadFontSize = T2(1)
        .HeadFontBold = T2(2)
        .HeadFontItalic = T2(3)
        .HeadFontUnderline = T2(4)
        .HeadFontStrikethrough = T2(5)
        .HeadFontColor = T2(6)
        T2 = Split(T(12), "|")
        .FootFontName = T2(0)
        .FootFontSize = T2(1)
        .FootFontBold = T2(2)
        .HeadFontItalic = T2(3)
        .FootFontUnderline = T2(4)
        .HeadFontStrikethrough = T2(5)
        .FootFontColor = T2(6)
    End With
End Sub
Public Function ReadPagePicture() As StdPicture
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  intKind-��������; strPageNo-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strPicFile As String
    Set ReadPagePicture = Nothing
    strPicFile = zlBlobRead(7, Me.���� & "-" & Me.ҳ��)
    If gobjFSO.FileExists(strPicFile) Then
        Err = 0: On Error Resume Next
        Set ReadPagePicture = LoadPicture(strPicFile)
        gobjFSO.DeleteFile strPicFile, True      'ɾ����ʱ�ļ�
    End If
End Function
Public Sub ReadPageHead(edtThis As Object)
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  intKind-��������; strPageNo-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strFile As String
    strFile = zlBlobRead(12, Me.���� & "-" & Me.ҳ��, App.Path & "\Head.rtf")
    If gobjFSO.FileExists(strFile) Then
        edtThis.HeadFile = strFile           '��ȡ�ļ�
        gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        If Trim(edtThis.HeadFileText) = "" Then GoTo Headtxt
    Else
Headtxt:
        edtThis.Head = Me.ҳü
        edtThis.HeadTextToFile '�����ֶ���Rtf�ؼ���
    End If
End Sub
Public Sub ReadPageFoot(edtThis As Object)
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  intKind-��������; strPageNo-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strFile As String, strFoot As String
    strFile = zlBlobRead(13, Me.���� & "-" & Me.ҳ��, App.Path & "\Foot.rtf")
    If gobjFSO.FileExists(strFile) Then
        edtThis.FootFile = strFile            '��ȡ�ļ�
        gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        If Trim(edtThis.FootFileText) = "" Then GoTo Foottxt
    Else
Foottxt:
        edtThis.Foot = Me.ҳ��
        edtThis.FootTextToFile '�����ֶ���Rtf�ؼ���
    End If
End Sub