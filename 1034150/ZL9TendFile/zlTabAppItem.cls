VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlTFTabAppItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarText As String '���ظ���
Public Property Let Text(ByVal vData As String)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.Text = 5
    mvarText = vData
End Property


Public Property Get Text() As String
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.Text
    Text = mvarText
End Property


