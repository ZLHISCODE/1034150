VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRCompend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const COM_BACKCOLOR = &HD5FEFF              '������ٵı�����ɫ '&HDCDCDC
Private Const COM_FORECOLOR = vbBlack               '������ٵ�ǰ����ɫ
Private Const COM_UNDERLINE = cprWave               '������ٵ��»���

Private mvarKey As Long             '��Ӧ�ڼ��ϵ�Key
Private mvar��Key As Long           '��Key
Private mvarID As Double              'ID
Private mvar�ļ�ID As Long          '�ļ�ID
Private mvar��ID As Double            '��ID
Private mvar������� As Double        '�������
Private mvar�������� As Boolean     '�����������
Private mvar���� As String          '���ƣ������ı�
Private mvar˵�� As String          '˵������������
Private mvar�������ID As Long      '�������ID
Private mvarԤ�����ID As Long      'Ԥ�����ID
Private mvar������� As Boolean     '����������   '����ʾ����û�С�������١��ĸ���
Private mvarʹ��ʱ�� As String      '���ʹ��ʹ��ʱ��   '����ʾ����û�С�ʹ��ʱ�����ĸ���
Private mvar��ʼ�� As Long          '��ʼ�棬ר���ڡ����Ӳ�����¼��
Private mvar��ֹ�� As Long          '��ֹ�棬ר���ڡ����Ӳ�����¼��
Private mvarLevel As Long           'Level��Level 1��2��3��...

Private mvar�ڲ���� As Long        '�ڲ���ţ�������ٿ��������������Ų�ͬ������������ļ���ͳһ����

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ��Key(ByVal vData As Long)
    mvar��Key = vData
End Property

Public Property Get ��Key() As Long
    ��Key = mvar��Key
End Property

Public Property Let ID(ByVal vData As Double)
    mvarID = vData
End Property

Public Property Get ID() As Double
    ID = mvarID
End Property

Public Property Let �ļ�ID(ByVal vData As Long)
    mvar�ļ�ID = vData
End Property

Public Property Get �ļ�ID() As Long
    �ļ�ID = mvar�ļ�ID
End Property

Public Property Let ��ID(ByVal vData As Double)
    mvar��ID = vData
End Property

Public Property Get ��ID() As Double
    ��ID = mvar��ID
End Property

Public Property Let �������(ByVal vData As Double)
    mvar������� = vData
End Property

Public Property Get �������() As Double
    ������� = mvar�������
End Property

Public Property Let ��������(ByVal vData As Boolean)
    mvar�������� = vData
End Property

Public Property Get ��������() As Boolean
    �������� = mvar��������
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
    ���� = Replace(mvar����, "'", "' || chr(39) || '")
End Property

Public Property Let ˵��(ByVal vData As String)
    mvar˵�� = vData
End Property

Public Property Get ˵��() As String
    ˵�� = Replace(mvar˵��, "'", "' || chr(39) || '")
End Property

Public Property Let �������ID(ByVal vData As Long)
    mvar�������ID = vData
End Property

Public Property Get �������ID() As Long
    �������ID = mvar�������ID
End Property

Public Property Let Ԥ�����ID(ByVal vData As Long)
    mvarԤ�����ID = vData
End Property

Public Property Get Ԥ�����ID() As Long
    Ԥ�����ID = mvarԤ�����ID
End Property

Public Property Let �������(ByVal vData As Boolean)
    mvar������� = vData
End Property

Public Property Get �������() As Boolean
    ������� = mvar�������
End Property

Public Property Let ʹ��ʱ��(ByVal vData As String)
    mvarʹ��ʱ�� = vData
End Property

Public Property Get ʹ��ʱ��() As String
    ʹ��ʱ�� = mvarʹ��ʱ��
End Property

Public Property Let ��ʼ��(ByVal vData As Long)
    mvar��ʼ�� = vData
End Property

Public Property Get ��ʼ��() As Long
    ��ʼ�� = IIf(mvar��ʼ�� = 0, 1, mvar��ʼ��)
End Property

Public Property Let ��ֹ��(ByVal vData As Long)
    mvar��ֹ�� = vData
End Property

Public Property Get ��ֹ��() As Long
'    ��ֹ�� = IIf(mvar��ֹ�� = 0, 0, mvar��ֹ��)
    ��ֹ�� = 0
End Property

Public Property Let Level(ByVal vData As Long)
    mvarLevel = vData
End Property

Public Property Get Level() As Long
    Level = mvarLevel
End Property

Public Property Let �ڲ����(ByVal vData As Long)
    mvar�ڲ���� = vData
End Property

Public Property Get �ڲ����() As Long
    �ڲ���� = mvar�ڲ����
End Property

'################################################################################################################
'## ���ܣ�  ������ٶ���
'##
'## ������  blnWithKey  :�Ƿ���ͬKeyֵһ���ƣ�
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cEPRCompend
    Dim cNode As New cEPRCompend
    If blnWithKey Then
        cNode.Key = Me.Key
    End If
    cNode.ID = Me.ID
    cNode.�ļ�ID = Me.�ļ�ID
    cNode.��ʼ�� = Me.��ʼ��
    cNode.��ֹ�� = Me.��ֹ��
    cNode.��ID = Me.��ID
    cNode.��Key = Me.��Key
    cNode.������� = Me.�������
    cNode.�������� = Me.��������
    cNode.���� = Me.����
    cNode.˵�� = Me.˵��
    cNode.�������ID = Me.�������ID
    cNode.Ԥ�����ID = Me.Ԥ�����ID
    cNode.������� = Me.�������
    cNode.ʹ��ʱ�� = Me.ʹ��ʱ��
    cNode.Level = Me.Level
    cNode.�ڲ���� = Me.�ڲ����
    Set Clone = cNode
End Function

Public Function FillCompendMember(ByVal rs As ADODB.Recordset, ByVal strTableName As String) As Boolean
'���ܣ������ݼ�¼������Ա
    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    
    If NVL(rs!������, 0) <> 0 Then Me.Key = NVL(rs!������, 0)
    Me.Level = NVL(rs!Level, 0)
    Me.ID = rs!ID
    Me.�ļ�ID = NVL(rs!�ļ�ID, 0)
    Me.��ID = NVL(rs!��ID, 0)
    Me.������� = NVL(rs!�������, 0)
    Me.�������� = IIf(NVL(rs!��������, 0) = 0, False, True)
    Me.���� = NVL(rs!�����ı�)
    Me.˵�� = NVL(rs!��������)
    Me.�������ID = NVL(rs!�������ID, 0)
    Me.Ԥ�����ID = NVL(rs!Ԥ�����ID, 0)
    Me.������� = IIf(NVL(rs!�������, 0) = 0, False, True)
    Me.ʹ��ʱ�� = NVL(rs!ʹ��ʱ��)
    Me.�ڲ���� = Me.�������
    
    If InStr(strTableName, "���Ӳ�������") > 0 Then '�����ļ��ṹ�Ͳ�����������û���������ֶ�,���������������
        Me.��ʼ�� = NVL(rs!��ʼ��, 1)
        Me.��ֹ�� = NVL(rs!��ֹ��, 0)
    End If
    
    FillCompendMember = True
End Function
'################################################################################################################
'## ���ܣ�  �������ID�Ƿ���Ч
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š�
'################################################################################################################
Public Sub CheckValidID(eEdtType As EditTypeEnum)
    Select Case eEdtType
    Case cprET_�����ļ�����
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("�����ļ��ṹ")
    Case cprET_ȫ��ʾ���༭
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("������������")
    Case cprET_�������༭, cprET_���������
        If Me.ID = 0 Then Me.ID = GetEPRContentNextId
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  ������ٵ�Editor��
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'##         lngStartPos     :��ǰλ��
'##         blnFirstInsert  :�ǲ��ǵ�һ�β��루��ʱ��ĩβ�Զ���������һ�д������ʽ���ı������������������ͬ��
'##         objDocument     :Document�������ڴ�������Ԥ�����ʱ�����������ȵ����⴦����
'################################################################################################################
Public Function InsertIntoEditor(ByRef edtThis As Object, _
    Optional ByVal lngStartPos As Long = -1, _
    Optional ByVal blnFirstInsert As Boolean = False, _
    Optional ByRef objDocument As cEPRDocument) As Boolean
        
    If edtThis.ViewMode <> cprNormal Then InsertIntoEditor = False: Exit Function
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '��֤���ܲ���ؼ����ڲ�

    Dim p As Long, lngKey As Long, blnForce As Boolean, strTmp As String, lngLen As Long
    blnForce = edtThis.ForceEdit

    With edtThis
        .Freeze
        .ForceEdit = True
        
        strTmp = Format(Me.Key, "00000000") & "," & IIf(Me.��������, 1, 0) & ",0)"
        
        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        If lngStartPos = 0 Then
            .Range(lngStartPos, lngStartPos) = "OS(" & strTmp & "OE(" & strTmp
            .Range(lngStartPos, lngStartPos + 32).Font.Hidden = True
            .Range(lngStartPos, lngStartPos + 32).Font.Protected = True
        Else
            If .Range(lngStartPos - 2, lngStartPos) = vbCrLf Then
                .Range(lngStartPos, lngStartPos) = "OS(" & strTmp & "OE(" & strTmp
                .Range(lngStartPos, lngStartPos + 32).Font.Hidden = True
                .Range(lngStartPos, lngStartPos + 32).Font.Protected = True
            Else
                .Range(lngStartPos, lngStartPos) = vbCrLf & "OS(" & strTmp & "OE(" & strTmp
                .Range(lngStartPos, lngStartPos + 2).Font.Protected = False
                .Range(lngStartPos, lngStartPos + 2).Font.Hidden = False
                lngStartPos = lngStartPos + 2
                .Range(lngStartPos, lngStartPos + 32).Font.Hidden = True
                .Range(lngStartPos, lngStartPos + 32).Font.Protected = True
            End If
        End If
        If Me.Ԥ�����ID < 0 Then
            'Ԥ����ٵ����⴦��
            If blnFirstInsert Then
'               SetCommonStyle edtThis, Me.Level, lngStartPos, lngStartPos + 32, True
                lngStartPos = lngStartPos + 32
                .Range(lngStartPos, lngStartPos) = Me.���� & vbCrLf
                lEE = lngStartPos + Len(Me.����)
'               SetCommonStyle edtThis, Me.Level, lngStartPos, lEE, True
                .Range(lngStartPos, lEE + 2).Font.Protected = False
                .Range(lngStartPos, lEE + 2).Font.Hidden = False
                lngStartPos = lEE + 2
                
            End If
        Else
            lngStartPos = lngStartPos + 32
            If blnFirstInsert Then
'                SetCommonStyle edtThis, Me.Level, lngStartPos - 32, lngStartPos, True
                .Range(lngStartPos, lngStartPos) = Me.���� & vbCrLf
                lEE = lngStartPos + Len(Me.����)
'                SetCommonStyle edtThis, Me.Level, lngStartPos, lEE, True
                .Range(lngStartPos, lEE + 2).Font.Protected = False
                .Range(lngStartPos, lEE + 2).Font.Hidden = False
                lngStartPos = lEE + 2
'                SetCommonStyle edtThis, "����", lngStartPos, lngStartPos, True
            End If
        End If
        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        .Range(lngStartPos, lngStartPos).Selected
                
        .ForceEdit = blnForce
        .UnFreeze
    End With
    InsertIntoEditor = True
End Function

'################################################################################################################
'## ���ܣ�  ��ȡ��ٱ��浽���ݿ��SQL���
'##
'## ������  eEdtType        :IN��       ��ǰ�༭ģʽ
'##         ArraySQL()      :IN/OUT��   SQL����
'################################################################################################################
Public Function GetSaveSQL(ByVal eEdtType As EditTypeEnum, ByRef ArraySQL() As String) As Boolean
    Dim lngCount As Long
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String
    
    Call CheckValidID(eEdtType)         '���ID��Ч�ԣ�
    Select Case eEdtType
    Case cprET_�����ļ�����
        gstrSQL = "Zl_�����ļ��ṹ_Update(" & Me.ID & "," & Me.�ļ�ID & "," & IIf(Me.��ID = 0, "NULL", Me.��ID) & "," & _
            Me.������� & ",1," & Me.Key & "," & IIf(Me.��������, 1, 0) & ",'" & Me.˵�� & "',NULL,'" & Me.���� & "',NULL," & _
            IIf(Me.Ԥ�����ID = 0, "NULL", Me.Ԥ�����ID) & "," & IIf(Me.�������, 1, 0) & ",'" & Me.ʹ��ʱ�� & "')"
    Case cprET_ȫ��ʾ���༭
        gstrSQL = "Zl_������������_Update(" & Me.ID & "," & Me.�ļ�ID & "," & IIf(Me.��ID = 0, "NULL", Me.��ID) & "," & _
            Me.������� & ",1," & Me.Key & "," & IIf(Me.��������, 1, 0) & ",'" & Me.˵�� & "',NULL,'" & Me.���� & "',NULL," & _
            IIf(Me.�������ID = 0, "NULL", Me.�������ID) & "," & IIf(Me.Ԥ�����ID = 0, "NULL", Me.Ԥ�����ID) & "," & IIf(Me.�������, 1, 0) & ",'" & Me.ʹ��ʱ�� & "')"
    Case cprET_�������༭, cprET_���������
        '���� ��ʼ�棬��ֹ�� �ֶΣ�������ͬ
        gstrSQL = "ZL_���Ӳ�������_Update(" & Me.ID & "," & Me.�ļ�ID & "," & Me.��ʼ�� & "," & Me.��ֹ�� & "," & IIf(Me.��ID = 0, "NULL", Me.��ID) & "," & _
            Me.������� & ",1," & Me.Key & "," & IIf(Me.��������, 1, 0) & ",'" & Me.˵�� & "',NULL,'" & Me.���� & "',NULL," & _
            IIf(Me.�������ID = 0, "NULL", Me.�������ID) & "," & IIf(Me.Ԥ�����ID = 0, "NULL", Me.Ԥ�����ID) & "," & IIf(Me.�������, 1, 0) & ",'" & Me.ʹ��ʱ�� & "')"
    End Select
    ArraySQL(lngCount) = gstrSQL
    GetSaveSQL = True
End Function

'################################################################################################################
'## ���ܣ�  ��λ���ı��е���ٿ�ͷλ��
'##
'## ������  edtThis     :Editor�༭��
'################################################################################################################
Public Sub GotoStartPos(ByRef edtThis As Object)
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean
    bFinded = FindKey(edtThis, "O", Me.Key, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        edtThis.TOM.TextDocument.Range(lEE, lEE).Select
'        edtThis.Range(lEE, lEE).ScrollIntoView cprSPStart
        If edtThis.Visible And edtThis.Enabled Then edtThis.SetFocus
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ��ȡ������ݵ���ʼ����ֹλ��
'##
'## ������  edtThis     :Editor�༭��
'##         lngStartPos :������ʼλ��
'##         lngEndPos   :���ؽ���λ��
'##
'## ˵����  ���ĩβ��һ�����ز������� vbCRLF��
'################################################################################################################
Public Function GetPosition(ByRef edtThis As Object, ByRef lngStartPos As Long, ByRef lngEndPos As Long) As Boolean
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean, lKey As Long
    bFinded = FindKey(edtThis, "O", Me.Key, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        lngStartPos = lEE - 32   '����һ��vbCrLf��λ��
    Else
        GetPosition = False
        Exit Function
    End If
    
    bFinded = FindNextKey(edtThis, lEE, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        lngEndPos = lSS
    Else
        lngEndPos = Len(edtThis.Text)
    End If
    GetPosition = True
End Function