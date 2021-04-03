VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Enum NumberChkEnum
    ChkBlank = 0  '���Ϊ��
    ChkBig = 1    '���> 0        �����ֵ,���Դ����ֵΪ׼
    ChkBigEq = 2  '���>=0        �����ֵ,���Դ����ֵΪ׼
    ChkNum = 3    '����Ƿ�Ϊ����
    ChkSml = 4    '���< 0        �����ֵ,���Դ����ֵΪ׼
    ChkSmlEq = 5  '���<=0        �����ֵ,���Դ����ֵΪ׼
End Enum

'ADO����
Private mclsFormObjects()      As New clsFormObject

Dim strLanguage             As String
Dim strSystem               As String
Dim StrKey                  As String
Dim strSQL                  As String

'==================================================================================================
'=PUBLIC
'=============================================================
'=����:����(But_Lock)
'=��ڲ���:
'=1.)����Button��������butName��������           ����:Form
'=���ڲ���:(��)
'=����:��մ������ı��������
'=����:2004-08-16
'=���:л��
'==============================================================
Public Sub But_Lock(objFrm As Object, Optional butName As String)
    Dim mCtl As Control
On Error GoTo ErrH
    For Each mCtl In objFrm.Controls
        Debug.Print mCtl.Name
        If TypeOf mCtl Is CommandButton Then                      'VB���ı���
            If mCtl.Name <> butName Then mCtl.Enabled = False
        End If
    Next
    Exit Sub
ErrH:
    Err.Clear
End Sub

'=============================================================
'=����:����(But_Open)
'=��ڲ���:
'=1.)����Button��������butName��������           ����:Form
'=���ڲ���:(��)
'=����:��մ������ı��������
'=����:2004-08-16
'=���:л��
'==============================================================
Public Sub But_Open(objFrm As Object, Optional butName As String)
    Dim mCtl As Control
On Error GoTo ErrH
    For Each mCtl In objFrm.Controls
        If TypeOf mCtl Is CommandButton Then                    'VB���ı���
            If mCtl.Name <> butName Then mCtl.Enabled = True
        End If
    Next
    Exit Sub
ErrH:
    Err.Clear
End Sub

'========================================================================================
'=���Q:���(ChkRsState)
'=��ڲ���:Rs               ����:ADODB.Recordset
'=���ڲ���:ChkRsState       ����:Boolean
'=����:����¼����״̬
'=����:2004-07-08
'=����:л��
'========================================================================================
Function Connection_ChkRsState(rs As ADODB.Recordset) As Boolean
On Error GoTo ErrH:
    With rs
        If rs Is Nothing Then
            Connection_ChkRsState = True
            Exit Function
        Else
            Connection_ChkRsState = False
        End If
        If rs.State = 0 Then
            Connection_ChkRsState = True
            Exit Function
        Else
            Connection_ChkRsState = False
        End If
        If .RecordCount < 1 Then
            Connection_ChkRsState = True
        Else
            Connection_ChkRsState = False
        End If
        If .EOF Or .BOF Then
            Connection_ChkRsState = True
        Else
            Connection_ChkRsState = False
        End If
    End With
    Exit Function
ErrH:
    Err.Clear
End Function

'==================================================================================================
'=��¼��λ
'==================================================================================================
Function Connection_GetBookMark(RsT As ADODB.Recordset, sFilter As String) As Variant
On Error GoTo ErrH
    If Connection_ChkRsState(RsT) Then
        Connection_GetBookMark = Null
    Else
        With RsT.Clone
            .Filter = sFilter
            If .EOF Or .BOF Then
                Connection_GetBookMark = Null
            Else
                Connection_GetBookMark = .Bookmark
            End If
        End With
    End If
    Exit Function
ErrH:
    Connection_GetBookMark = Null
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=ɾ����λ
'==================================================================================================
Function Connection_DelLocate(rs As ADODB.Recordset, MPOS As Variant) As Variant
    Dim lrow As Variant
On Error GoTo ErrH
    Connection_DelLocate = Null
    With rs
        If Connection_ChkRsState(rs) Then
            Connection_DelLocate = Null
        Else
            .MoveLast
            If .AbsolutePosition <= 0 Then
                lrow = .RecordCount
            Else
                lrow = .AbsolutePosition
            End If
            If CInt(lrow) >= CInt(MPOS) Then
                Connection_DelLocate = MPOS
            Else
                If CInt(MPOS) > CInt(lrow) Then
                    Connection_DelLocate = .Bookmark
                End If
            End If
        End If
    End With
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=��    ��:����ؼ��е�������Ϊ��
'=�� �� ֵ:�����Ϊ��,���ؿ��ַ���,���Ϊ�շ���ʾ��Ϣ[�����Դ������ȵ�ֵ]
'=��    ��:л��
'=��    ��:2005-07-16
'==================================================================================================
Public Function Chk_CheckTxtNull(mDescription As String, objText As Object, Optional ctrSetF As Object, Optional strMsg As String) As String
On Error GoTo ErrH
    Chk_CheckTxtNull = strMsg
    Select Case TypeName(objText)
        Case "TextBox", "BetryNote", "BetrySelect", "BetryDate", "BetryNum", "ComboBox"
            If Trim(objText.Text) = "" Then
                If ctrSetF Is Nothing Then Set ctrSetF = objText
                Chk_CheckTxtNull = strMsg & mDescription & "����Ϊ��!" & vbCrLf
            End If
    End Select
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'=============================================================
'=��    ��:����ؼ��е������Ƿ�Ϊ��ֵ
'=�� �� ֵ:�����,���ؿ��ַ���,�������,����ʾ��Ϣ
'=��    ��:����
'=��    ��:2004-06-24
'=�޸Ĺ���:���붨λ�ؼ�
'=�� �� ��:л��
'=��    ��:2004-07-20
'=============================================================
Public Function Chk_CheckTxtNum(mDescription As String, objText As Object, Optional ctrSetF As Object, Optional strMsg As String, Optional mType As NumberChkEnum, Optional mChkVal As Double) As String
    Dim strTmp  As String
On Error GoTo ErrH
    If IsNull(mType) = True Then mType = ChkBlank   '���Ĭ��Ϊ���Ϊ��
    If IsNull(mChkVal) = True Then mChkVal = 0      '���ֵĬ��Ϊ0
    strTmp = Trim(objText.Text)
    
    Chk_CheckTxtNum = ""
    Select Case mType
        Case ChkBlank  '���Ϊ��
            If Trim(strTmp) = "" Then Chk_CheckTxtNum = mDescription & "����Ϊ��!"
        Case ChkBig     '���> 0        �����ֵ,���Դ����ֵΪ׼
            If strTmp = "" Then
                Chk_CheckTxtNum = mDescription & "����Ϊ��!"
            ElseIf IsNumeric(strTmp) = False Then
                Chk_CheckTxtNum = mDescription & "����Ϊ����ֵ!"
            Else
                If Val(strTmp) <= mChkVal Then Chk_CheckTxtNum = mDescription & "����>" & mChkVal & "!"
            End If
        Case ChkBigEq   '���>=0        �����ֵ,���Դ����ֵΪ׼
            If strTmp = "" Then
                Chk_CheckTxtNum = mDescription & "����Ϊ��!"
            ElseIf IsNumeric(strTmp) = False Then
                Chk_CheckTxtNum = mDescription & "����Ϊ����ֵ!"
            Else
                If Val(strTmp) < mChkVal Then Chk_CheckTxtNum = mDescription & "����>=" & mChkVal & "!"
            End If
        Case ChkNum     '����Ƿ�Ϊ����
            If strTmp = "" Then
                Chk_CheckTxtNum = mDescription & "����Ϊ��!"
            Else
                If IsNumeric(strTmp) = False Then Chk_CheckTxtNum = mDescription & "����Ϊ����ֵ!"
            End If
        Case ChkSml     '���< 0        �����ֵ,���Դ����ֵΪ׼
            If strTmp = "" Then
                Chk_CheckTxtNum = mDescription & "����Ϊ��!"
            ElseIf IsNumeric(strTmp) = False Then
                Chk_CheckTxtNum = mDescription & "����Ϊ����ֵ!"
            Else
                If Val(strTmp) >= mChkVal Then Chk_CheckTxtNum = mDescription & "����<" & mChkVal & "!"
            End If
        Case ChkSmlEq   '���<=0        �����ֵ,���Դ����ֵΪ׼
            If strTmp = "" Then
                Chk_CheckTxtNum = mDescription & "����Ϊ��!"
            ElseIf IsNumeric(strTmp) = False Then
                Chk_CheckTxtNum = mDescription & "����Ϊ��!"
            Else
                If Val(strTmp) > mChkVal Then Chk_CheckTxtNum = mDescription & "����<=" & mChkVal & "!"
            End If
    End Select
    If Chk_CheckTxtNum <> "" Then
        If ctrSetF Is Nothing Then Set ctrSetF = objText
        Chk_CheckTxtNum = Chk_CheckTxtNum & vbCrLf
        Chk_CheckTxtNum = strMsg & Chk_CheckTxtNum
    Else
        Chk_CheckTxtNum = strMsg
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=��    ��:������������Ƿ����Ҫ��
'=�� �� ֵ:BOOLEAN[Y.�˳�,N.����]
'=��ڲ���:
'=       1:����ַ���
'=       2:��λ�Ŀؼ�
'=��    ��:л��
'=��    ��:2005-07-17
'==================================================================================================
Public Function Chk_CheckMsg(strMsg As String, Optional ctrSetF As Object) As Boolean
On Error GoTo ErrH
    Chk_CheckMsg = True
    If strMsg = "" Then
        Chk_CheckMsg = False
    Else
        MsgBox strMsg, vbInformation, "������ʾ"
        If (Not ctrSetF Is Nothing) Then
            If (ctrSetF.TabStop = True And ctrSetF.Enabled = True) Then ctrSetF.SetFocus
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=��    ��:����ؼ��е�������Ϊ������
'=�� �� ֵ:�����Ϊ��,���ؿ��ַ���,���Ϊ�շ���ʾ��Ϣ
'=��    ��:����
'=��    ��:2004-06-24
'=��    ��:�����ַ�����,����Ƿ�Ϊ��ȷ������������
'=�� �� ��:л��
'=����:2004-01-12
'==================================================================================================
Function Chk_Date(strDate As String, Optional mTile As String) As String
    Dim strMsg    As String
On Error GoTo ErrH
    strMsg = ""
    Chk_Date = ""
    If strDate <> "" Then
        If Not IsDate(strDate) Then
            strMsg = mTile & "����Ϊ������!" & vbCrLf
        ElseIf (Format(strDate, "yyyy-mm-dd")) < "1900" Then
            strMsg = mTile & "�������1900��!" & vbCrLf
        ElseIf (Format(strDate, "yyyy-mm-dd")) > "2099" Then
            strMsg = mTile & "����С��2099��!" & vbCrLf
        End If
    End If
    Chk_Date = strMsg
    Exit Function
ErrH:
    Err.Clear
    Chk_Date = "����Ϊ������!"
    Exit Function
End Function

'==================================================================================================
'=��  ��:Chk_SrhNum
'=��  ��:������ݲ�ѯ����ʼ����
'=��  ��:strTitle     ��Ϣ��
'=       objTextQ   ��ʼ�Ŀؼ�
'=       objTextH   ��ֹ�Ŀؼ�
'=����ֵ:
'==================================================================================================
Public Function Chk_SrhNum(ByVal strTitle As String, objTextQ As Object, objTextH As Object, Optional ctrSetF As Object, Optional strMsg As String) As String
    Dim num1    As Double
    Dim num2    As Double
    Dim strQ    As String
    Dim strH    As String
On Error GoTo ErrH
    strQ = Trim(objTextQ.Text)
    strH = Trim(objTextH.Text)
    Chk_SrhNum = ""
    If strQ = "" Then
        num1 = 0
    Else
        If Not IsNumeric(strQ) Then
            Chk_SrhNum = Chk_SrhNum & strTitle & "����ʼ������ʽ����ȷ!" & vbCrLf
            objTextQ.Text = ""
        Else
            num1 = CDbl(objTextQ.Text)
        End If
    End If
    If strH = "" Then
        num2 = 0
    Else
        If Not IsNumeric(strH) Then
            Chk_SrhNum = Chk_SrhNum & strTitle & "����ֹ������ʽ����ȷ!" & vbCrLf
            objTextH.Text = ""
        Else
            num2 = CDbl(objTextH.Text)
        End If
    End If
    If num1 > num2 And Len(strH) > 0 Then
        Chk_SrhNum = Chk_SrhNum & strTitle & "��Χ������ֹ�����ܴ�����ʼ����" & vbCrLf
        objTextQ.Text = ""
    End If
    If Chk_SrhNum <> "" Then
        If ctrSetF Is Nothing Then Set ctrSetF = objTextQ
    End If
    Chk_SrhNum = Chk_SrhNum & strMsg
    Exit Function
ErrH:
    Chk_SrhNum = ""
    Err.Clear
    Resume Next
End Function

'==================================================================================================
'=����:�õ������б��������(Cmb_List)
'=��ڲ���:
'=1).�����б���ؼ�         ����:Control
'=2).ԭ�����ID             ����:String
'=3).���(btyEditMode)      ����:Byte
'=      1.�޲���:��׼���ݼ���
'=      2.����Ϊ2:��׼���ݲ�ѯ
'=���ڲ���:��
'=����:��SSQL���ҵ����ݼ��ص������б�����ȥ
'=����:2004-12-11
'=���:л��
'=˵��:��ԭ�����ID�е����ݲ��ܴ�"-"
'==================================================================================================
Sub Cmb_List(Combo As Object, rs As ADODB.Recordset, Optional btyEditMode As Byte, Optional btyOne As Byte)
    Dim lngStep     As Long
On Error GoTo ErrH
    Combo.Clear
 
    If Connection_ChkRsState(rs) Then Exit Sub
    lngStep = 0
    With rs
        .MoveFirst
        Do While Not (.EOF Or .BOF)
            If btyOne = 1 Then
                Combo.List(lngStep) = Trim(rs.Fields(0))
                lngStep = lngStep + 1
                rs.MoveNext
            Else
                Combo.List(lngStep) = Trim(rs.Fields(0)) + strSplitCmb + Trim(rs.Fields(1))
                lngStep = lngStep + 1
                rs.MoveNext
            End If
        Loop
        Combo.ListIndex = 0
    End With
    If btyEditMode = 2 Then
        Combo.List(lngStep) = ""
        Combo.ListIndex = -1
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'==================================================================================================
'=����:�õ������б����Text������ȡ��ID(Cmb_ID)
'=��ڲ���:
'=1).�����б���ؼ�         ����:Control
'=���ڲ���:��
'=����:�õ������б����Text������ȡ��ID
'=����:2004-12-11
'=���:л��
'=˵��:��ԭ�����ID�е����ݲ��ܴ�"-"
'==================================================================================================
Function Cmb_ID(Combo As Object, Optional Index As Byte = 1) As String
    Dim xx          As Variant
On Error GoTo ErrH
    If Combo.Text = "" Then
        Cmb_ID = ""
    Else
        xx = Split(Combo.Text, strSplitCmb)
        If Index - 1 <= UBound(xx) Then '����±�ֵС������ֵ[֤���н�ȡֵ]
            Cmb_ID = xx(Index - 1)
        Else                        '����±�ֵ���ڵ�������ֵ[֤�����޽�ȡֵ]������
            Cmb_ID = "[��]"
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=����:�õ������б����Text������ȡ��ID(Cmb_EditIndex)
'=��ڲ���:
'=1).�����б���ؼ�         ����:Control
'=���ڲ���:��
'=����:�õ������б����Text������ȡ��ID
'=����:2004-12-11
'=���:л��
'=˵��:��ԭ�����ID�е����ݲ��ܴ�"-"
'==================================================================================================
Function Cmb_EditIndex(Combo As Object, sID As String) As Long
    Dim lngCount    As Long
    Dim lngStep     As Long
    Dim xx          As Variant
On Error GoTo ErrH
    lngCount = Combo.ListCount - 1
    For lngStep = 0 To lngCount
        xx = Split(Combo.List(lngStep), strSplitCmb)
        If sID = xx(0) Then
            Cmb_EditIndex = lngStep
            Exit For
        End If
    Next
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=����:���ı���ֵ(Frm_GetControl)
'=��ڲ���:
'=1.)objFrm������յĴ�����             ����:Form
'=2.)Rs��¼��                           ADODB.Recordset
'=���ڲ���:(��)
'=����:��մ������ı��������
'=����:2004-08-16
'=���:л��
'=˵��:���ı���DataField���踳�϶�Ӧ���ֶε�ֵ
'=�޸�:2004-09-13   �����,������CheckBox��ѡ������
'=���ÿؼ�:1.TextBox--VB���ı���    2.BetryNote--��ע����ؼ�   3.BetrySelect--����ѡ��ؼ�
'=4.BetryDate--��������ؼ�      5.BetryNum--��������ؼ�    6.CheckBox--��ѡ������
'==================================================================================================
Public Sub Frm_GetControl(objFrm As Object, rs As ADODB.Recordset)
    Dim mCtl As Control
On Error GoTo ErrH
    With rs.Clone
        For Each mCtl In objFrm.Controls
            Select Case TypeName(mCtl)
                Case "TextBox"
                    If Len(Trim(mCtl.DataField)) > 0 Then
                        If mCtl.Tag <> "" Then
                            If UCase(mCtl.Tag) <> "MAX" Or UCase(mCtl.Tag) <> "MIN" Then
                                mCtl.Text = "" & .Fields(mCtl.DataField)
                            Else
                                mCtl.Text = Format("" & .Fields(mCtl.DataField), mCtl.Tag)
                            End If
                        Else
                            mCtl.Text = "" & .Fields(mCtl.DataField)
                        End If
                    End If
                Case "ComboBox"
                    If Len(Trim(mCtl.DataField)) > 0 Then
                        If mCtl.STYLE <> 2 Then
                            mCtl.Text = "" & .Fields(mCtl.DataField)
                        Else
                            mCtl.ListIndex = Cmb_EditIndex(mCtl, "" & rs.Fields(mCtl.DataField))
                        End If
                    End If
                Case "CheckBox"
                    If Len(Trim(mCtl.DataField)) > 0 Then mCtl.Value = IIf("" & .Fields(mCtl.DataField) = "1", 1, 0)
            End Select
        Next
    End With
    Exit Sub
ErrH:
    Debug.Print mCtl.DataField
    Err.Clear
    Resume Next
End Sub

'========================================================================================
'=����:���(System)
'=��ڲ���:
'=1.)objFrm������յĴ�����           ����:Form
'=2.)SetMode                          ����:Byte
'=    (1.�޲���,��׼�������;2.����Ϊ2,��׼�༭,��ѯ����)
'=���ڲ���:(��)
'=����:��մ������ı��������(��ϵͳ���ݴ����ʽ��)
'=����:2004-08-16
'=���:л��
'========================================================================================
Public Sub Sys_System(objFrm As Object)
    Dim mCtl                As Control
    Dim i                   As Long
    Dim intSetp             As Integer
    Dim strTabName          As String
    
On Error GoTo ErrH
    i = objFrm.hWnd
    ReDim Preserve mclsFormObjects(i + objFrm.Controls.count)
    
    For Each mCtl In objFrm.Controls
        With mCtl
            Select Case TypeName(mCtl)
                Case "SkinLabel"
                    .Caption = (.Caption)
                Case "TextBox"
                    Set mclsFormObjects(i).txt = mCtl
                    .Text = ""
                    .Appearance = 1
                    If .Locked Then .BackColor = &H80000000 Else .BackColor = vbWhite
                Case "CheckBox"
                    Set mclsFormObjects(i).chk = mCtl
                    .Value = 0
                    
                Case "ComboBox"
                    Set mclsFormObjects(i).Cmb = mCtl
                    .ListIndex = -1
                    .Appearance = 1
                    If .Locked Then .BackColor = &H80000000 Else .BackColor = vbWhite
                Case "CommandButton"
                    .Caption = (.Caption)
                    .ToolTipText = (.ToolTipText)
                Case "Label"
                    .Caption = (.Caption)
                Case "Frame"
                    .Caption = (.Caption)
                Case "VSFlexGrid"
                    RestoreFlexState mCtl, objFrm.Caption
            End Select
        End With
        i = i + 1
    Next
    Exit Sub
ErrH:
    Err.Clear
    Resume Next
End Sub

'========================================================================================
'=�����������,�õ����ҵĴ�д����
'========================================================================================
Function Sys_HBTran(money As String) As String '
    Dim X               As String
    Dim Y               As String
    Dim i               As Long
    Const zimu = ".sbqwsbqysbqwsbq" '����λ�ô���
    Const letter = "0123456789sbqwy.zjf" '���庺����д
    Const upcase = "��Ҽ��������½��ƾ�ʰ��Ǫ�f�|Բ���Ƿ�" '�����д����
    Dim temp As String
On Error GoTo ErrH
    temp = money
    If InStr(temp, ".") > 0 Then temp = Left(temp, InStr(temp, ".") - 1)
    If Len(temp) > 16 Then MsgBox "��Ŀ̫���޷����㣡������һ�������µ�����", 64, "������ʾ": Exit Function 'ֻ��ת��һ����Ԫ������Ŀ�Ļ��ң�
    X = Format(money, "0.00")                                           '��ʽ������
    Y = ""
    For i = 1 To Len(X) - 3
        Y = Y & Mid(X, i, 1) & Mid(zimu, Len(X) - 2 - i, 1)
    Next
    If Right(X, 3) = ".00" Then
        Y = Y & "z"                                                     '***Ԫ��
    Else
        Y = Y & Left(Right(X, 2), 1) & "j" & Right(X, 1) & "f"          '*Ԫ*��*��
    End If
    Y = Replace(Y, "0q", "0")                                           '������ǧ(�磺40200���f��ǧ�㷡��)
    Y = Replace(Y, "0b", "0")                                           '�������(�磺41000���fҼǧ���)
    Y = Replace(Y, "0s", "0")                                           '������ʮ(�磺204������ʰ����)
    Do While Y <> Replace(Y, "00", "0")
        Y = Replace(Y, "00", "0")                                       '����˫��(�磺1004ҼǪ������)
    Loop
    Y = Replace(Y, "0y", "y")                                           '������|(�磺210�|     ����Ҽʮ��|)
    Y = Replace(Y, "0w", "w")                                           '�������f(�磺210�f     ����Ҽʮ���f)
    Y = IIf(Len(X) = 5 And Left(Y, 1) = "1", Right(Y, Len(Y) - 1), Y)   '����Ҽʮ(�磺14Ҽʰ����10Ҽʰ)
    Y = IIf(Len(X) = 4, Replace(Y, "0.", ""), Replace(Y, "0.", "."))    '������Ԫ(�磺20.00��ʰ��Բ��0.12��ԲҼ�Ƿ���)
    For i = 1 To 19
        Y = Replace(Y, Mid(letter, i, 1), Mid(upcase, i, 1)) '��д����
    Next
    Sys_HBTran = Y
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=��  ��:Find_Txt
'=��  ��:��ѯ�������͵ķ���SQL����
'=��  ��:mField     �ֶ���
'=       objTextQ   ��ʼ���ı���
'=       objTextH   ��ֹ���ı���
'=����ֵ:
'==================================================================================================
Public Function Find_Txt(ByVal strField As String, objText As Object, sSql As String) As String
    Dim strSQL  As String
On Error GoTo ErrH
    Find_Txt = sSql
    If Trim(objText.Text) = "" Then Exit Function
    strSQL = strField & " like '" & GetPPFS & Trim(objText.Text) & "%'"
    If Find_Txt = "" Then
        Find_Txt = Trim(strSQL) & vbCrLf
    Else
        Find_Txt = Trim(Find_Txt & " AND " & strSQL) & vbCrLf
    End If
    Exit Function
ErrH:
    Err.Clear
    Resume Next
End Function

'==================================================================================================
'=��  ��:Find_Cmb
'=��  ��:��ѯ�����б��е����ݵķ���SQL����
'=��  ��:mField     �ֶ���
'=       objTextQ   ��ʼ���ı���
'=       objTextH   ��ֹ���ı���
'=����ֵ:
'==================================================================================================
Public Function Find_Cmb(ByVal strField As String, objCmb As Object, sSql As String, Optional BtyNoSP As Byte) As String
    Dim strSQL  As String
On Error GoTo ErrH
    Find_Cmb = sSql
    If Trim(objCmb.Text) = "" Then Exit Function
    If BtyNoSP = 0 Then
        strSQL = " nvl(" & strField & ",'0') = '" & Trim(Cmb_ID(objCmb)) & "'"
    Else
        strSQL = " nvl(" & strField & ",'') = '" & Trim(Cmb_ID(objCmb)) & "'"
    End If
    If Find_Cmb = "" Then
        Find_Cmb = Trim(strSQL) & vbCrLf
    Else
        Find_Cmb = Trim(Find_Cmb & " AND " & strSQL) & vbCrLf
    End If
    Exit Function
ErrH:
    Err.Clear
    Resume Next
End Function

'==================================================================================================
'=�õ���ѯ���ڵ�����[1.��ѯ��������,2.��ѯ��ϸ����]
'==================================================================================================
Function Frm_GetFilter(ByRef objFrm As Object, Optional FindMode As Byte = 0) As String
    Dim mCtl        As Control
    Dim QCnt        As Control 'ǰֵ
    Dim HCnt        As Control '��ֵ
    Dim sTmp        As String
On Error GoTo ErrH
    sTmp = "1=1 "
    Select Case FindMode
        Case 0
            For Each mCtl In objFrm.Controls
                '����2Ϊ��,���´�ѭ��
                If Right(mCtl.Name, 1) <> "2" Then
                    If mCtl.Tag = "" Or mCtl.Tag = "M" Then
                        Select Case TypeName(mCtl)
                            Case "TextBox"
                                If mCtl.DataField <> "" Then sTmp = Find_Txt(mCtl.DataField, mCtl, sTmp)
                             
                            Case "ComboBox"
                                If mCtl.DataField <> "" Then sTmp = Find_Cmb(mCtl.DataField, mCtl, sTmp)
                        End Select
                    End If
                End If
            Next
        Case 1
            For Each mCtl In objFrm.Controls
                '����2Ϊ��,���´�ѭ��
                If Right(mCtl.Name, 1) <> "2" Then
                    If mCtl.Tag = "D" Then
                        Select Case TypeName(mCtl)
                            Case "TextBox"
                                If mCtl.DataField <> "" Then sTmp = Find_Txt(mCtl.DataField, mCtl, sTmp)
                            Case "ComboBox"
                                If mCtl.DataField <> "" Then sTmp = Find_Cmb(mCtl.DataField, mCtl, sTmp)
                        End Select
                    End If
                End If
            Next
    End Select
    If Trim(sTmp) = "1=1" Then Frm_GetFilter = "" Else Frm_GetFilter = UCase(sTmp)
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function
'==================================================================================================
'ȷ��/ȡ����Ϣ
'==================================================================================================
Public Function Msg_OKC(StrMessage As String, Optional MsgType As VbMsgBoxStyle = vbOKCancel + vbDefaultButton2 + vbQuestion, Optional strTitle As String = "������ʾ") As Boolean
On Error GoTo ErrH
    Msg_OKC = IIf(MsgBox(StrMessage, MsgType, strTitle) <> 1, True, False)
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'ȷ����Ϣ
'==================================================================================================
Public Sub Msg_OK(StrOKMessage As String, Optional MsgType As VbMsgBoxStyle = vbOKOnly + vbInformation, Optional strTitle As String = "������ʾ")
On Error GoTo ErrH
    MsgBox StrOKMessage, MsgType, strTitle
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'==================================================================================================
'ORACLE ���ṹ
'==================================================================================================
Function GetRsFieldWidth(strTableName) As ADODB.Recordset
    On Error GoTo ErrH
    
    gstrSQL = "" & vbCrLf & _
        "SELECT USER_TAB_COLS.TABLE_NAME   as ����," & vbCrLf & _
        "       USER_TAB_COLS.COLUMN_NAME  as ����," & vbCrLf & _
        "       USER_TAB_COLS.DATA_TYPE    as ��������," & vbCrLf & _
        "       USER_TAB_COLS.DATA_LENGTH  as ����," & vbCrLf & _
        "       USER_TAB_COLS.NULLABLE     as �Ƿ�Ϊ��," & vbCrLf & _
        "       USER_TAB_COLS.COLUMN_ID    as �����," & vbCrLf & _
        "       user_col_comments.comments as ��ע" & vbCrLf & _
        "  From USER_TAB_COLS" & vbCrLf & _
        " INNER join user_col_comments" & vbCrLf & _
        "    ON user_col_comments.TABLE_NAME = USER_TAB_COLS.TABLE_NAME" & vbCrLf & _
        "   AND user_col_comments.COLUMN_NAME = USER_TAB_COLS.COLUMN_NAME" & vbCrLf & _
        " WHERE USER_TAB_COLS.Table_Name = [1]"
    
    Set GetRsFieldWidth = zlDatabase.OpenSQLRecord(gstrSQL, "clsCheck", strTableName)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Function
End Function