VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPROutLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cCPROutLine"
Attribute VB_Ext_KEY = "Member0" ,"cCPROutLine"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlng��ˮ�� As Long

Public Function LastOutline() As cCPROutLine
    Set LastOutline = mCol(mCol.Count)
End Function

Public Function Add() As cCPROutLine
    '�����¶���
    Dim objNewMember As cCPROutLine
    Dim sKey As String
    
    Set objNewMember = New cCPROutLine
    
    mlng��ˮ�� = mlng��ˮ�� + 1
    sKey = "K" & mlng��ˮ��
    objNewMember.��ˮ�� = mlng��ˮ��
    
    mCol.Add objNewMember, sKey
    
    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As cCPROutLine
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
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub

Public Sub FillTree(ByVal objTree As Object)
    '���ڵ���䵽����
    Dim i As Long, j As Long, k As Long
    objTree.Nodes.Clear
    
    If mCol.Count = 0 Then
        Exit Sub
    End If
    LockWindowUpdate objTree.hWnd
    mCol(1).��� = 1        '���±���һ���Ǳ���1
    
    Dim Root As Node    '���ڵ�
    Dim NodeOutline(1 To 9) As Node    '�ӽڵ�
    Dim Node As Node
    
    Set Root = objTree.Nodes.Add(Text:="�ļ����", Image:=1)
    Root.Bold = True
    Root.Expanded = True
    
    For k = 1 To mCol.Count
        For i = 1 To mCol.Count
            If mCol(i).��ˮ�� = mCol(k).��ˮ�� Then
                If mCol(i).��� > j + 1 Then mCol(i).��� = j + 1
                Select Case mCol(i).���
                Case 1
                    Set Node = objTree.Nodes.Add(Root, tvwChild, "K" & mCol(i).��ˮ��, IIf(Len(mCol(i).����) > 18, Left(mCol(i).����, 18) & "...", mCol(i).����), 2)
                Case Else
                    Set Node = objTree.Nodes.Add(NodeOutline(mCol(i).��� - 1), tvwChild, "K" & mCol(i).��ˮ��, IIf(Len(mCol(i).����) > 18, Left(mCol(i).����, 18) & "...", mCol(i).����), 2)
                End Select
                Set NodeOutline(mCol(i).���) = Node
                j = mCol(i).���
                Node.Expanded = True
                Exit For
            End If
        Next i
    Next k
    
    LockWindowUpdate 0
End Sub






