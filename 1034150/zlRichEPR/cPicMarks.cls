VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPicMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cPicMark"
Attribute VB_Ext_KEY = "Member0" ,"cPicMark"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlngKey As Long

Public Property Get Item(vntIndexKey As Variant) As cPicMark
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
    mlngKey = 0
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'################################################################################################################
'## ���ܣ�  ����������ӵ����ϵ� cPicMark ����
'################################################################################################################
Public Function LastPicMark() As cPicMark
    Set LastPicMark = mCol(mCol.Count)
End Function

'################################################################################################################
'## ���ܣ�  �����¶���
'##
'## ������  lngKey :���ڻָ��ɵļ��϶��󣬲��������Ҫ���������ˮ�š�
'################################################################################################################
Public Function Add(Optional lngKey As Long = 0) As Long
    Dim objNewMember As cPicMark
    Set objNewMember = New cPicMark
    Dim sKey As String

    mlngKey = mlngKey + 1
    sKey = "K" & mlngKey
    mCol.Add objNewMember, sKey
    mCol(mCol.Count).Key = "K" & mlngKey        '�����Լ��ڸ������еĹؼ��֣�����ɾ���ڵ㣡
    
    Add = mlngKey
    Set objNewMember = Nothing
End Function

'################################################################################################################
'## ���ܣ�  ���Ƽ���
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cPicMarks
    Dim cNodes As New cPicMarks
    Dim i As Long
    For i = 1 To mCol.Count
        cNodes.Add
        If blnWithKeys Then cNodes.LastPicMark.Key = mCol(i).Key
        cNodes.LastPicMark.Selected = mCol(i).Selected
        cNodes.LastPicMark.X1 = mCol(i).X1
        cNodes.LastPicMark.Y1 = mCol(i).Y1
        cNodes.LastPicMark.X2 = mCol(i).X2
        cNodes.LastPicMark.Y2 = mCol(i).Y2
        cNodes.LastPicMark.�㼯 = mCol(i).�㼯
        cNodes.LastPicMark.���� = mCol(i).����
        cNodes.LastPicMark.���� = mCol(i).����
        cNodes.LastPicMark.��䷽ʽ = mCol(i).��䷽ʽ
        cNodes.LastPicMark.���ɫ = mCol(i).���ɫ
        cNodes.LastPicMark.�߿� = mCol(i).�߿�
        cNodes.LastPicMark.����ɫ = mCol(i).����ɫ
        cNodes.LastPicMark.���� = mCol(i).����
        cNodes.LastPicMark.���� = mCol(i).����
        cNodes.LastPicMark.����ɫ = mCol(i).����ɫ
    Next i
    Set Clone = cNodes
End Function

'################################################################################################################
'## ���ܣ�  ��ȡ��ǵ������ַ���
'##
'## ˵����  ���á�||����������
'################################################################################################################
Public Property Get ��������() As String
    Dim strR As String, i As Long
    For i = 1 To mCol.Count
        If i = 1 Then
            strR = strR & mCol(i).��������
        Else
            strR = strR & "||" & mCol(i).��������
        End If
    Next
    �������� = strR
End Property

'################################################################################################################
'## ���ܣ�  ͨ��һ���ַ��������ñ�Ǽ�������
'##
'## ˵����  ���á�||����������
'################################################################################################################
Public Property Let ��������(vData As String)
    Dim T As Variant, i As Long
    T = Split(vData, "||")
    mlngKey = 0
    Set mCol = New Collection
    For i = 0 To UBound(T)
        Call Add
        Me.LastPicMark.�������� = T(i)
    Next
End Property
