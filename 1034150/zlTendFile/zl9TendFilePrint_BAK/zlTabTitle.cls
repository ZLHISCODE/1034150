VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlTabTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'�ֲ�������������ֵ
Private mvarText As String '���ظ���
Private mvarColor As Long   '���ظ���
Private mvarFont As StdFont '���ظ���

Public Property Set Font(ByVal vData As StdFont)
    Set mvarFont = vData
End Property


Public Property Get Font() As StdFont
    Set Font = mvarFont
End Property

Public Property Let Color(ByVal vData As Long)
    mvarColor = vData
End Property


Public Property Get Color() As Long
    Color = mvarColor
End Property


Public Property Let Text(ByVal vData As String)
    mvarText = vData
End Property


Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = mvarText
End Property


Private Sub Class_Initialize()
    mvarText = ""
    mvarColor = 0
    Set mvarFont = New StdFont
    mvarFont.Name = "����"
    mvarFont.Size = 18
    mvarFont.Bold = False
    mvarFont.Italic = False

End Sub

Private Sub Class_Terminate()
    Set mvarFont = Nothing
End Sub