VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'��Ҫ�ر༭ʱ���ı�
Option Explicit
Private mvarKey As Long                     '�����еĹؼ���
Private mvarID As Long                      '
Private mvar��ID As Long
Private mvar���� As String
Private mvar�������� As Boolean
Private mvar�����д� As Long
Private mvar�����ı� As String
Private mvar��ʼ�� As Long
Private mvar��ֹ�� As Long
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
Public Property Let ��ID(ByVal vData As Long)
    mvar��ID = vData
End Property

Public Property Get ��ID() As Long
    ��ID = mvar��ID
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
    ���� = mvar����
End Property
Public Property Let ��������(ByVal vData As Boolean)
    mvar�������� = vData
End Property

Public Property Get ��������() As Boolean
    �������� = mvar��������
End Property

Public Property Let �����д�(ByVal vData As Long)
    mvar�����д� = vData
End Property

Public Property Get �����д�() As Long
    �����д� = mvar�����д�
End Property
Public Property Let �����ı�(ByVal vData As String)
    mvar�����ı� = vData
End Property

Public Property Get �����ı�() As String
    �����ı� = mvar�����ı�
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
Public Sub Clone(cNode As cTabText)
    With cNode
        .�������� = .��������
        .�����д� = Me.�����д�
        .�����ı� = Me.�����ı�
        .��ʼ�� = Me.��ʼ��
        .��ֹ�� = Me.��ֹ��
    End With
End Sub
Public Function GetSaveSQL(ByVal editType As TabET, ByVal CellTmp As cTabCell, arrSQL As Variant) As Boolean
'����:����̶��ı�,���ı�����,������������
    On Error GoTo errHand
    With CellTmp
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        Select Case editType
        Case TabET_�����ļ�����
            arrSQL(UBound(arrSQL)) = "Zl_�����ļ��ṹ_Update(" & mvarID & "," & .�ļ�ID & "," & .ID & _
                                        "," & .������� & "," & cprCTText & ",Null," & IIf(mvar��������, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|'," & mvar�����д� & ",'" & Replace(mvar�����ı�, "'", "'" & "||Chr(39)||" & "'") & "')"
        Case TabET_ȫ��ʾ���༭
            arrSQL(UBound(arrSQL)) = "Zl_������������_Update(" & mvarID & "," & .�ļ�ID & "," & .ID & _
                                        "," & .������� & "," & cprCTText & ",Null," & IIf(mvar��������, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|'," & mvar�����д� & ",'" & Replace(mvar�����ı�, "'", "'" & "||Chr(39)||" & "'") & "')"
        Case TabET_�������༭, TabET_���������
            arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�������_Update(" & mvarID & "," & .�ļ�ID & "," & mvar��ʼ�� & "," & mvar��ֹ�� & "," & .ID & _
                                        "," & .������� & "," & cprCTText & ",Null," & IIf(mvar��������, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|'," & mvar�����д� & ",'" & Replace(mvar�����ı�, "'", "'" & "||Chr(39)||" & "'") & "')"
        End Select
    End With
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function InsertIntoEditor(ByRef edtThis As Object, Optional ByVal lngStartPos As Long = -1) As Boolean
'******************************************************************************************************************
' ���ܣ�  ��������Ҫ�ص�Editor��
'
' ������  edtThis         :��ǰ�ı༭���ؼ�
'         lngStartPos     :��ǰλ��
'         blnFirstIns     :�Ƿ��ǵ�һ�β��루Ĭ�ϵ�һ�β������ĩβ�Ӷ��ţ�
'******************************************************************************************************************
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '��֤���ܲ���ؼ����ڲ�
    
    If edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then Exit Function

    Dim p As Long, lngKey As Long, blnForce As Boolean, lngLen As Long, lngLen2 As Long, strTmp As String
    p = lngStartPos
    blnForce = edtThis.ForceEdit

    With edtThis
        .Freeze
        .ForceEdit = True
        .Range(lngStartPos, lngStartPos).Text = Me.�����ı�
        If Me.�������� Then
            .Range(lngStartPos, Len(Me.�����ı�)).Font.Protected = True '����Ҫ�ؼ���������
            .Range(lngStartPos, Len(Me.�����ı�)).Font.BackColor = &HFFC0C0
        End If
        lngStartPos = lngStartPos + Len(Me.�����ı�)                               'λ�ú���
        .Range(lngStartPos, lngStartPos).Selected                                           '��궨λ
        .ForceEdit = blnForce
        .UnFreeze
    End With
    InsertIntoEditor = True
End Function
'################################################################################################################
'## ���ܣ�  ������ж����ID�Ƿ���Ч
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š���ID=0������ǵ�Ԫ����CellȡID
'################################################################################################################
Public Sub CheckValidID(eEditType As TabET)
    Dim i As Long, strTable As String
    Select Case eEditType
    Case TabET_�����ļ�����
        strTable = "�����ļ��ṹ"
    Case TabET_ȫ��ʾ���༭
        strTable = "������������"
    Case TabET_�������༭, TabET_���������
        strTable = "���Ӳ�������"
    End Select
    
    If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId(strTable)
End Sub
