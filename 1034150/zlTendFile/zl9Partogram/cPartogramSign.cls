VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPartogramSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'###############################################################################################
'##ģ �� ����cPartogramSign.cls
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��11��24��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ��������ǩ�����󣨶������ͣ�8��
'##��    ����
'###############################################################################################

Option Explicit
'ǩ��״̬
Public Enum EPRSignLevelEnum
    cprSL_�հ� = 0              'δǩ��
    cprSL_���� = 1              '����ҽʦǩ��
    cprSL_���� = 2              '����ҽʦǩ��
    cprSL_���� = 3              '����ҽʦǩ��
    cprSL_���� = 4              '���ߣ�ǩ�����𲻰�����ֻ��ʾ��Ա��������ְ�ƣ��Ա���������ҽʦ
End Enum

Private mvarKey As Long                     '�ؼ��֣��洢�ڡ������ǡ���
Private mvarID As Long                      '���ڱ����IDֵ
Private mvar�ļ�ID As Long                  '�ļ�ID
Private mvar��ID As Long                    '��ID
Private mvar��ʼ�� As Long                  '��ʼ��
Private mvar��ֹ�� As Long                  '��ֹ��
Private mvar������� As Long                'ÿ��ǩ�������ı�˳���и�Ψһ��ţ�ȫ��Ψһ��
Private mvar���� As String                  'ǩ���û����������洢�ڡ������ı�����
Private mvarǩ����Ϣ As String              '����ǩ���ַ������洢�ڡ�Ҫ��ֵ����
Private mvarǰ������ As String              '�ƺ����߶ٺŵ�ǰ�����֣��洢�ڡ�Ҫ�����ơ���
Private mvarǩ������ As EPRSignLevelEnum    '0,1,2,3���洢�ڡ�Ҫ�ر�ʾ����
Private mvarǩ����ID As Long                '��¼ǩ����ID
Private mblnǩ��ͼƬ As Boolean             '�Ƿ�ʹ��ǩ��ͼƬ,�洢��"������̬"��

'����5��������ɡ��������ԡ����洢�ڡ��������ԡ��ֶ���
Private mvarǩ����ʽ As Long                '1������ǩ����2������ǩ��
Private mvarǩ������ As Long                '��1��ʼ��ţ��ɳ����Զ����亭�壨��ʾ��ͬ�汾��ǩ��������
                                            'Ŀǰ����ʹ�õ�ǩ������1
Private mvar֤��ID As Long                  '����֤��ID��
Private mvar��ʾ��ǩ As Boolean             '�Ƿ���ʾ��ǩ��־
Private mvarǩ��ʱ�� As String              'ǩ��ʱ�䣬���ڻ���ʱ�ָ����ʱ��
Private mvar��ʾʱ�� As String              '�Ƿ���ʾǩ��ʱ�估��ʾ�ĸ�ʽ��,�մ���ʾ����ʾ
Private mvarǩ��Ҫ�� As String             '�Ƿ�Ϊǩ��Ҫ��Keyֵ
Private mvarʱ��� As String                'ʱ������洢�ڡ�Ҫ�ص�λ����
Private mvarʱ�����Ϣ As String            '����ǩ�����ص�ʱ�������֤������Ϣ(Base64)

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Long)
    mvarID = vData
End Property

Public Property Get ID() As Long
    ID = mvarID
End Property

Public Property Let �ļ�ID(ByVal vData As Long)
    mvar�ļ�ID = vData
End Property

Public Property Get �ļ�ID() As Long
    �ļ�ID = mvar�ļ�ID
End Property

Public Property Let ��ID(ByVal vData As Long)
    mvar��ID = vData
End Property

Public Property Get ��ID() As Long
    ��ID = mvar��ID
End Property

Public Property Let ��ʼ��(ByVal vData As Long)
    mvar��ʼ�� = vData
End Property

Public Property Get ��ʼ��() As Long
    ��ʼ�� = mvar��ʼ��
End Property

Public Property Let ��ֹ��(ByVal vData As Long)
    mvar��ֹ�� = vData
End Property

Public Property Get ��ֹ��() As Long
    ��ֹ�� = mvar��ֹ��
End Property

Public Property Let �������(ByVal vData As Long)
    mvar������� = vData
End Property

Public Property Get �������() As Long
    ������� = mvar�������
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
Attribute ����.VB_UserMemId = 0
    ���� = mvar����
End Property
Public Property Let ǩ����ID(ByVal vData As Long)
    mvarǩ����ID = vData
End Property

Public Property Get ǩ����ID() As Long
    ǩ����ID = mvarǩ����ID
End Property

Public Property Let ǩ����Ϣ(ByVal vData As String)
    mvarǩ����Ϣ = vData
End Property

Public Property Get ǩ����Ϣ() As String
    ǩ����Ϣ = mvarǩ����Ϣ
End Property

Public Property Let ǩ��Ҫ��(ByVal vData As String)
    mvarǩ��Ҫ�� = vData
End Property

Public Property Get ǩ��Ҫ��() As String
    ǩ��Ҫ�� = mvarǩ��Ҫ��
End Property

Public Property Let ǰ������(ByVal vData As String)
    mvarǰ������ = vData
End Property

Public Property Get ǰ������() As String
    ǰ������ = mvarǰ������
End Property

Public Property Let ǩ������(ByVal vData As EPRSignLevelEnum)
    mvarǩ������ = vData
End Property

Public Property Get ǩ������() As EPRSignLevelEnum
    ǩ������ = mvarǩ������
End Property

Public Property Let ǩ����ʽ(ByVal vData As Long)
    mvarǩ����ʽ = vData
End Property

Public Property Get ǩ����ʽ() As Long
    ǩ����ʽ = mvarǩ����ʽ
End Property

Public Property Let ǩ������(ByVal vData As Long)
    mvarǩ������ = vData
End Property

Public Property Get ǩ������() As Long
    ǩ������ = mvarǩ������
End Property

Public Property Let ֤��ID(ByVal vData As Long)
    mvar֤��ID = vData
End Property

Public Property Get ֤��ID() As Long
    ֤��ID = mvar֤��ID
End Property

Public Property Let ǩ��ʱ��(ByVal vData As String)
    mvarǩ��ʱ�� = vData
End Property

Public Property Get ǩ��ʱ��() As String
    ǩ��ʱ�� = mvarǩ��ʱ��
End Property

Public Property Let ��ʾʱ��(ByVal vData As String)
    mvar��ʾʱ�� = vData
End Property

Public Property Get ��ʾʱ��() As String
    ��ʾʱ�� = mvar��ʾʱ��
End Property

Public Property Let ʱ���(ByVal vData As String)
    mvarʱ��� = vData
End Property

Public Property Get ʱ���() As String
    ʱ��� = mvarʱ���
End Property

Public Property Let ʱ�����Ϣ(ByVal vData As String)
    mvarʱ�����Ϣ = vData
End Property

Public Property Get ʱ�����Ϣ() As String
    ʱ�����Ϣ = mvarʱ�����Ϣ
End Property

Public Property Let ��ʾ��ǩ(ByVal vData As Boolean)
    mvar��ʾ��ǩ = vData
End Property

Public Property Get ��ʾ��ǩ() As Boolean
    ��ʾ��ǩ = mvar��ʾ��ǩ
End Property

Public Property Let ��������(ByVal vData As String)
    Dim T As Variant
    T = Split(vData, ";")
    On Error Resume Next
    mvarǩ����ʽ = Val(T(0))
    mvarǩ������ = Val(T(1))
    mvar֤��ID = Val(T(2))
    mvar��ʾ��ǩ = (Val(T(3)) = 1)
    mvarǩ��ʱ�� = Format(T(4), "yyyy-mm-dd hh:mm:ss")
    mvar��ʾʱ�� = T(5)
    If UBound(T) >= 6 Then
        mvarǩ��Ҫ�� = CStr(T(6))
    Else
        mvarǩ��Ҫ�� = ""
    End If
End Property

Public Property Get ��������() As String
    �������� = CStr(mvarǩ����ʽ) & ";" & CStr(mvarǩ������) & ";" & _
        CStr(mvar֤��ID) & ";" & IIf(mvar��ʾ��ǩ, 1, 0) & ";" & _
        Format(mvarǩ��ʱ��, "yyyy-mm-dd hh:mm:ss") & ";" & mvar��ʾʱ�� & ";" & mvarǩ��Ҫ��
End Property