VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPROutLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarDBID As Long          '���ݿ�ID
Private mvar�ļ�ID As Long
Private mvar���� As String
Private mvar˵�� As String
Private mvar���� As Boolean
Private mvarԤ��ID As Long
Private mvar���� As Boolean
Private mvar��� As Long
Private mvar�ı� As String

Private mvar��ˮ�� As Long      '��Ӧ�ڼ��ϵ�Key��

Public Property Let ��ˮ��(ByVal vData As Long)
    mvar��ˮ�� = vData
End Property

Public Property Get ��ˮ��() As Long
    ��ˮ�� = mvar��ˮ��
End Property

Public Property Let �ı�(ByVal vData As String)
    mvar�ı� = vData
End Property

Public Property Get �ı�() As String
    �ı� = mvar�ı�
End Property

Public Property Let ���(ByVal vData As Long)
    mvar��� = vData
End Property

Public Property Get ���() As Long
    ��� = mvar���
End Property

Public Property Let ����(ByVal vData As Boolean)
    mvar���� = vData
End Property

Public Property Get ����() As Boolean
    ���� = mvar����
End Property

Public Property Let Ԥ��ID(ByVal vData As Long)
    mvarԤ��ID = vData
End Property

Public Property Get Ԥ��ID() As Long
    Ԥ��ID = mvarԤ��ID
End Property

Public Property Let ����(ByVal vData As Boolean)
    mvar���� = vData
End Property

Public Property Get ����() As Boolean
    ���� = mvar����
End Property

Public Property Let ˵��(ByVal vData As String)
    mvar˵�� = vData
End Property

Public Property Get ˵��() As String
    ˵�� = mvar˵��
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
    ���� = mvar����
End Property

Public Property Let �ļ�ID(ByVal vData As Long)
    mvar�ļ�ID = vData
End Property

Public Property Get �ļ�ID() As Long
    �ļ�ID = mvar�ļ�ID
End Property

Public Property Let DBID(ByVal vData As Long)
    mvarDBID = vData
End Property

Public Property Get DBID() As Long
    DBID = mvarDBID
End Property



