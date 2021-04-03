VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediMasters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MediMaster"
Attribute VB_Ext_KEY = "Member0" ,"MediMaster"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'�ֲ����������漯��
Private mCol As Collection

Private Function AddMaster(ByVal objMaster As MediMaster, Optional sKey As String) As Boolean
    '�����¶���
    Dim objNewMember As MediMaster
    
    Set objNewMember = objMaster
    '���ô��뷽��������
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As MediMaster
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    ' x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)
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

Public Function GetMediMasters(ByVal dateBeging As Date, ByVal DateEnd As Date, ByVal lng����ID As Long, Optional ByVal lng����ID As Long)
    'ȡ�ݴ�ҩƷ��Ϣ,�浽MediMasters��
    Dim strSQL As String, rsMain As New ADODB.Recordset, rsBill As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim objMediMaster As MediMaster, objMediBill As MediBill
    Dim i As Integer
    On Error GoTo errHandle
    For i = 1 To mCol.Count
        Remove 1
    Next
    '�����ʱ��ȡ����¼,��ʾ����������,��ϸ�а��������¼
    strSQL = "Select C.���� as ������,C.���� as ������,B.����ID, A.NO, B.�����, B.����, B.�Ա�, B.����, A.����Ա, A.�Ǽ�ʱ��, A.����ʱ��, Sum(A.��� * ���ϵ��) As �ϼ�, A.ժҪ" & vbNewLine & _
            "From ���ű� C,������Ϣ B, �ݴ�ҩƷ��¼ A" & vbNewLine & _
            "Where A.����ID=C.ID And A.����id = B.����id And A.���ϵ��=1 And A.�Ǽ�ʱ�� Between [1] And [2] And A.����id = [3] " & IIf(lng����ID = 0, "", "And A.����ID=[4]") & vbNewLine & _
            "Group By C.���� ,C.���� ,B.����ID, A.NO, B.�����, B.����, B.�Ա�, B.����, A.����Ա, A.�Ǽ�ʱ��, A.����ʱ��, A.ժҪ"
    Set rsMain = zlDatabase.OpenSQLRecord(strSQL, "GetMediBills", dateBeging, DateEnd, lng����ID, lng����ID)
    Do Until rsMain.EOF
        Set objMediMaster = New MediMaster
        objMediMaster.NO = IIf(IsNull(rsMain.Fields("NO")), "", rsMain.Fields("NO"))
        objMediMaster.����� = IIf(IsNull(rsMain.Fields("�����")), 0, rsMain.Fields("�����"))
        objMediMaster.���� = IIf(IsNull(rsMain.Fields("����")), "", rsMain.Fields("����"))
        objMediMaster.�Ա� = IIf(IsNull(rsMain.Fields("�Ա�")), "", rsMain.Fields("�Ա�"))
        objMediMaster.���� = IIf(IsNull(rsMain.Fields("����")), "", rsMain.Fields("����"))
        objMediMaster.����Ա = IIf(IsNull(rsMain.Fields("����Ա")), "", rsMain.Fields("����Ա"))
        objMediMaster.�Ǽ�ʱ�� = IIf(IsNull(rsMain.Fields("�Ǽ�ʱ��")), CDate(0), rsMain.Fields("�Ǽ�ʱ��"))
        objMediMaster.�ϼ� = IIf(IsNull(rsMain.Fields("�ϼ�")), 0, Val(rsMain.Fields("�ϼ�")))
        objMediMaster.ժҪ = IIf(IsNull(rsMain.Fields("ժҪ")), "", rsMain.Fields("ժҪ"))
        objMediMaster.����ID = IIf(IsNull(rsMain.Fields("����ID")), 0, rsMain.Fields("����ID"))
        objMediMaster.����ID = lng����ID
        objMediMaster.�������� = IIf(IsNull(rsMain.Fields("������")), "", "[" & rsMain.Fields("������") & "]" & rsMain.Fields("������"))
        objMediMaster.����ʱ�� = IIf(IsNull(rsMain.Fields("����ʱ��")), CDate(0), rsMain.Fields("����ʱ��"))
        
        strSQL = "Select a.�Ǽ�ʱ��, a.����, a.���, a.���, a.��λ, a.����," & vbNewLine & _
                "            a.���ϵ��, a.ʹ��״̬, a.����, Nvl(b.��������, 0) As ��������, a.���, a.ҩƷid, a.ҩƷ����, a.ҽ��id, a.���ͺ�," & vbNewLine & _
                "            a.ִ�з���,a.ժҪ,a.����Ա" & vbNewLine & _
                "From (Select �Ǽ�ʱ��, ����, ���, ���, ��λ,  ����, ���ϵ��, ʹ��״̬, ����," & vbNewLine & _
                "                           ���, ҩƷid, ҩƷ����, Nvl(ҽ��id, 0) as ҽ��id, Nvl(���ͺ�, 0) as ���ͺ�, ִ�з���, ժҪ, ����Ա" & vbNewLine & _
                "            From �ݴ�ҩƷ��¼" & vbNewLine & _
                "            Where No = [1] And ���ϵ�� = 1 And ����ʱ�� " & IIf(objMediMaster.����ʱ�� = CDate(0), " is null", "=to_date('" & Format(objMediMaster.����ʱ��, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')") & ")  a," & vbNewLine & _
                "        (Select Sum(Nvl(����, 0)) As ��������, ���, ҩƷid, ҩƷ����, Nvl(ҽ��id,0) as ҽ��id, nvl(���ͺ�,0) as ���ͺ�, ִ�з���" & vbNewLine & _
                "            From �ݴ�ҩƷ��¼" & vbNewLine & _
                "            Where No = [1] And ���ϵ�� = -1 And ����ʱ�� " & IIf(objMediMaster.����ʱ�� = CDate(0), " is null", "=to_date('" & Format(objMediMaster.����ʱ��, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')") & vbNewLine & _
                "            Group By ���, ҩƷid, ҩƷ����, ҽ��id, ���ͺ�, ִ�з���) b" & vbNewLine & _
                "Where a.ҩƷid = b.ҩƷid(+) And a.ҩƷ���� = b.ҩƷ����(+) And a.ҽ��id = b.ҽ��id(+) And a.ִ�з��� = b.ִ�з���(+)" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select �Ǽ�ʱ��, ����, ���, ���, ��λ, ����, ���ϵ��, ʹ��״̬, ����," & vbNewLine & _
                "            0 As ��������, ���, ҩƷid, ҩƷ����, nvl(ҽ��id,0) as ҽ��id,nvl(���ͺ�,0) as ���ͺ�, ִ�з���, ժҪ, ����Ա" & vbNewLine & _
                "From �ݴ�ҩƷ��¼" & vbNewLine & _
                "Where No = [1] And ���ϵ�� = -1 And ����ʱ�� " & IIf(objMediMaster.����ʱ�� = CDate(0), " is null", "=to_date('" & Format(objMediMaster.����ʱ��, "yyyy-MM-dd hh:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')")

        Set rsBill = zlDatabase.OpenSQLRecord(strSQL, "GetMediBills", objMediMaster.NO)
        Do Until rsBill.EOF
            Set objMediBill = New MediBill
            objMediBill.���� = IIf(IsNull(rsBill.Fields("����")), 0, rsBill.Fields("����"))
            objMediBill.��� = IIf(IsNull(rsBill.Fields("���")), "", rsBill.Fields("���"))
            objMediBill.��� = IIf(IsNull(rsBill.Fields("���")), 0, rsBill.Fields("���"))
            objMediBill.���㵥λ = IIf(IsNull(rsBill.Fields("��λ")), "", rsBill.Fields("��λ"))
            objMediBill.���� = IIf(IsNull(rsBill.Fields("����")), 0, rsBill.Fields("����"))
            objMediBill.���ϵ�� = IIf(IsNull(rsBill.Fields("���ϵ��")), 0, rsBill.Fields("���ϵ��"))
            objMediBill.ʹ��״̬ = IIf(IsNull(rsBill.Fields("ʹ��״̬")), "", rsBill.Fields("ʹ��״̬"))
            objMediBill.���� = IIf(IsNull(rsBill.Fields("����")), 0, rsBill.Fields("����"))
            objMediBill.�������� = IIf(IsNull(rsBill.Fields("��������")), 0, rsBill.Fields("��������"))
            objMediBill.��� = IIf(IsNull(rsBill.Fields("���")), 0, rsBill.Fields("���"))
            objMediBill.ҩƷID = IIf(IsNull(rsBill.Fields("ҩƷID")), 0, rsBill.Fields("ҩƷID"))
            objMediBill.ҩƷ���� = IIf(IsNull(rsBill.Fields("ҩƷ����")), "", rsBill.Fields("ҩƷ����"))
            objMediBill.ҽ��ID = IIf(IsNull(rsBill.Fields("ҽ��ID")), 0, rsBill.Fields("ҽ��ID"))
            objMediBill.���ͺ� = IIf(IsNull(rsBill.Fields("���ͺ�")), 0, rsBill.Fields("���ͺ�"))
            objMediBill.ִ�з��� = IIf(IsNull(rsBill.Fields("ִ�з���")), 0, rsBill.Fields("ִ�з���"))
            objMediBill.�Ǽ�ʱ�� = IIf(IsNull(rsBill.Fields("�Ǽ�ʱ��")), CDate(0), rsBill.Fields("�Ǽ�ʱ��"))
            objMediBill.ʹ��ժҪ = IIf(IsNull(rsBill.Fields("ժҪ")), "", rsBill.Fields("ժҪ"))
            objMediBill.������ = IIf(IsNull(rsBill.Fields("����Ա")), "", rsBill.Fields("����Ա"))
            
            If objMediBill.ҩƷID > 0 Then
                strSQL = "Select b.������λ, a.����ϵ��, a.���ﵥλ, a.�����װ" & vbNewLine & _
                        "From ҩƷ��Ϣ b, ҩƷ��� a" & vbNewLine & _
                        "Where a.ҩ��id = b.ҩ��id And a.ҩƷid =[1] "
                        
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetҩƷ���", objMediBill.ҩƷID)
                If rsTmp.RecordCount > 0 Then
                    objMediBill.������λ = IIf(IsNull(rsTmp.Fields("������λ")), "", rsTmp.Fields("������λ"))
                    objMediBill.����ϵ�� = IIf(IsNull(rsTmp.Fields("����ϵ��")), 0, rsTmp.Fields("����ϵ��"))
                    objMediBill.���ﵥλ = IIf(IsNull(rsTmp.Fields("���ﵥλ")), "", rsTmp.Fields("���ﵥλ"))
                    objMediBill.�����װ = IIf(IsNull(rsTmp.Fields("�����װ")), "", rsTmp.Fields("�����װ"))
                End If
            Else
                objMediBill.������λ = objMediBill.���㵥λ
                objMediBill.����ϵ�� = 1
                objMediBill.���ﵥλ = objMediBill.���㵥λ
                objMediBill.�����װ = 1
            End If
            Call objMediMaster.AddBill(objMediBill, objMediBill.��� & "_" & objMediBill.���ϵ�� & "_" & Format(objMediBill.�Ǽ�ʱ��, "yyMMddhhmmss"))
            rsBill.MoveNext
        Loop
        Call AddMaster(objMediMaster, objMediMaster.NO & "_" & IIf(objMediMaster.����ʱ�� = CDate(0), "0", Format(objMediMaster.����ʱ��, "yyMMddhhmmss")))
        rsMain.MoveNext
    Loop
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function