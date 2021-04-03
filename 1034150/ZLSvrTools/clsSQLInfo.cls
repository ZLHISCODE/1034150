VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��������ֵ�ľֲ�����
Private mvarSQL As String '�ֲ�����
Private mvarTip As String '�ֲ�����
Private mvarBlock As Boolean '�ֲ�����
Private mvarBlockType As String '�ֲ�����
Private mvarBlockName As String '�ֲ�����
Private mvarFile As String '�ֲ�����
Private mvarFileLine As Long '�ֲ�����
Private mvarFileLines As Long '�ֲ�����
Private mvarBlockWrap As Boolean '�ֲ�����
Private mvarLobDDL As Boolean '�ֲ�����
'-----------------------------------------------------------------------
'������������Ӧ�Գ����Ŀ�����SQL(����SQL���ַ����滻�ȴ���ʱ�������������)
'------------------------------------------------------------------------
'�ж�SQLInfo���Ƿ���SQL
Private mvarHaveSQL   As Boolean '�ֲ�����
'��ЧSQL��ǰ�����У�һ����Ա�ʶSQL����
Private mvarPartSQL As String '�ֲ�����
Private mvarForeign As Boolean '�Ƿ������
Private mvarLackOnline As Boolean '�Ƿ��������Online�ؼ���
Private mblnInit As Boolean

Public Sub ClearSQL()
    mblnInit = True
    Me.SQL = ""
    Me.Tip = ""
    Me.Block = False
    Me.BlockWrap = False
    Me.BlockType = ""
    Me.BlockName = ""
    Me.FileLine = 0
    Me.HaveSQL = False
    Me.PartSQL = ""
    Me.Foreign = False
    Me.LackOnline = False
    Me.LobDDL = False
    mblnInit = False
End Sub

Public Function CopyMe() As clsSQLInfo
    Dim objMe As New clsSQLInfo
    mblnInit = True
    objMe.SQL = Me.SQL
    objMe.Tip = Me.Tip
    objMe.Block = Me.Block
    objMe.BlockWrap = Me.BlockWrap
    objMe.BlockType = Me.BlockType
    objMe.BlockName = Me.BlockName
    objMe.File = Me.File
    objMe.FileLine = Me.FileLine
    objMe.FileLines = Me.FileLines
    objMe.HaveSQL = Me.HaveSQL
    objMe.PartSQL = Me.PartSQL
    objMe.Foreign = Me.Foreign
    objMe.LackOnline = Me.LackOnline
    objMe.LobDDL = mvarLobDDL
    mblnInit = False
    Set CopyMe = objMe
End Function

Public Sub CopySQL(ByVal objSQL As clsSQLInfo)
    Me.SQL = objSQL.SQL
    Me.Tip = objSQL.Tip
    Me.Block = objSQL.Block
    Me.BlockWrap = objSQL.BlockWrap
    Me.BlockType = objSQL.BlockType
    Me.BlockName = objSQL.BlockName
    Me.File = objSQL.File
    Me.FileLine = objSQL.FileLine
    Me.FileLines = objSQL.FileLines
    Me.HaveSQL = objSQL.HaveSQL
    Me.PartSQL = objSQL.PartSQL
    Me.Foreign = objSQL.Foreign
    Me.LackOnline = objSQL.LackOnline
End Sub

Public Function IsSameTo(ByVal objSQL As clsSQLInfo) As Boolean
'���ܣ���ǰSQL��������һ��SQL�Աȣ��ж������Ƿ���ͬһ�����󣬷��ضԱȽ��
    If objSQL Is Nothing Then Exit Function
    If objSQL.SQL <> Me.SQL Then Exit Function
    If objSQL.Tip <> Me.Tip Then Exit Function
    If objSQL.Block <> Me.Block Then Exit Function
    If objSQL.BlockWrap <> Me.BlockWrap Then Exit Function
    If objSQL.BlockType <> Me.BlockType Then Exit Function
    If objSQL.BlockName <> Me.BlockName Then Exit Function
    If objSQL.File <> Me.File Then Exit Function
    If objSQL.FileLine <> Me.FileLine Then Exit Function
    If objSQL.FileLines <> Me.FileLines Then Exit Function
    If objSQL.HaveSQL <> Me.HaveSQL Then Exit Function
    If objSQL.PartSQL <> Me.PartSQL Then Exit Function
    If objSQL.Foreign <> Me.Foreign Then Exit Function
    If objSQL.LackOnline <> Me.LackOnline Then Exit Function
    IsSameTo = True
End Function

Public Property Let BlockWrap(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BlockWrap = 5
    mvarBlockWrap = vData
End Property

Public Property Get BlockWrap() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BlockWrap
    BlockWrap = mvarBlockWrap
End Property

Public Property Let FileLines(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FileLines = 5
    mvarFileLines = vData
End Property

Public Property Get FileLines() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FileLines
    FileLines = mvarFileLines
End Property

Public Property Let FileLine(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FileLine = 5
    mvarFileLine = vData
End Property

Public Property Get FileLine() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FileLine
    FileLine = mvarFileLine
End Property

Public Property Let File(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.File = 5
    mvarFile = vData
End Property

Public Property Get File() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.File
    File = mvarFile
End Property

Public Property Let BlockName(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BlockName = 5
    mvarBlockName = vData
End Property

Public Property Get BlockName() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BlockName
    BlockName = mvarBlockName
End Property

Public Property Let BlockType(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BlockType = 5
    mvarBlockType = vData
End Property

Public Property Get BlockType() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BlockType
    BlockType = mvarBlockType
End Property

Public Property Let Block(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Block = 5
    mvarBlock = vData
End Property

Public Property Get Block() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Block
    Block = mvarBlock
End Property

Public Property Let Tip(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Tip = 5
    mvarTip = vData
End Property

Public Property Get Tip() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Tip
    Tip = mvarTip
End Property

Public Property Let SQL(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SQL = 5
    mvarSQL = vData
End Property

Public Property Get SQL() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SQL
    SQL = mvarSQL
End Property

Public Property Let HaveSQL(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SQL = 5
    mvarHaveSQL = vData
End Property

Public Property Get HaveSQL() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SQL
    HaveSQL = mvarHaveSQL
End Property

Public Property Let LobDDL(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SQL = 5
    mvarLobDDL = vData
End Property

Public Property Get LobDDL() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SQL
    LobDDL = mvarLobDDL
End Property

Public Property Let PartSQL(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SQL = 5
    mvarPartSQL = vData
    If mblnInit Then Exit Property
    If Not Me.Block Then
        If Me.PartSQL Like "ALTER *" Then
            If Me.PartSQL Like "ALTER TABLE * ADD CONSTRAINT * FOREIGN KEY*" Then
                Me.Foreign = True
            Else
                If Not Me.PartSQL Like "* ONLINE*" Then 'û��OnLIne�ؼ���
                    If Me.PartSQL Like "ALTER TABLE * ADD CONSTRAINT * PRIMARY KEY* USING INDEX*" Then
                        Me.LackOnline = True
                    ElseIf Me.PartSQL Like "ALTER TABLE * ADD CONSTRAINT * UNIQUE* USING INDEX*" Then
                        Me.LackOnline = True
                    ElseIf Me.PartSQL Like "ALTER INDEX * REBUILD*" Then
                        Me.LackOnline = True
                    ElseIf Me.PartSQL Like "ALTER TABLE * MODIFY[( ]* [CB]LOB*" Then
                        Me.LobDDL = True
                    End If
                ElseIf Me.PartSQL Like "ALTER TABLE * MODIFY[( ]* [CB]LOB*" Then
                    Me.LobDDL = True
                End If
            End If
        ElseIf Me.PartSQL Like "CREATE INDEX *" Then
            If Not Me.PartSQL Like "* ONLINE*" Then 'û��OnLIne�ؼ���
                Me.LackOnline = True
            End If
        End If
    End If
End Property

Public Property Get PartSQL() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SQL
    PartSQL = mvarPartSQL
End Property

Public Property Let Foreign(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SQL = 5
    mvarForeign = vData
End Property

Public Property Get Foreign() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SQL
    Foreign = mvarForeign
End Property

Public Property Let LackOnline(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SQL = 5
    mvarLackOnline = vData
End Property

Public Property Get LackOnline() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SQL
    LackOnline = mvarLackOnline
End Property

Public Function AnsySQL() As ADODB.Recordset
'���ܣ��Դ洢���̡��������в�������
    If Me.Block Then
        If Me.BlockType Like "*PROCEDURE*" Or Me.BlockType Like "*FUNCTION*" Then
            Set AnsySQL = AnsyProcedurePara(True)
        End If
    End If
End Function


Private Function AnsyProcedurePara(Optional ByVal blnTruncOwner As Boolean) As ADODB.Recordset
  '******************************************************************************************************************
    '���ܣ����
    '������blnTruncOwner=�Ƿ�ȥ��������
    '���أ��ɹ�����true�����򷵻�false
    '******************************************************************************************************************
    Dim strFMTSQL As String, strType As String, strName As String, strProcPara As String, strReturn As String
    Dim strParName As String, strParDirection  As String, strParType As String, strParTypeMask As String, strParDefault As String
    Dim cllStrs As Collection
    Dim arrTmp As Variant, arrTmp1 As Variant
    Dim arrLine As Variant, rsReturn As ADODB.Recordset, intIndex As Integer, i As Long, j As Long
    '��ȡ��׼ͷ
    strFMTSQL = GetFMTSQLStr(TrimCommentLossless(Me.SQL), cllStrs)
    If strFMTSQL = "" Then Exit Function
    '��ȡ���̺���ͷ
    strFMTSQL = Split(strFMTSQL, " PIPELINED AS ")(0)
    If strFMTSQL = "" Then Exit Function
    strFMTSQL = Split(strFMTSQL, " PIPELINED;")(0)
    If strFMTSQL = "" Then Exit Function
    strFMTSQL = Split(strFMTSQL, " AS ")(0)
    If strFMTSQL = "" Then Exit Function
    strFMTSQL = Split(strFMTSQL, " IS ")(0)
    If strFMTSQL = "" Then Exit Function
    arrTmp = Split(Replace(strFMTSQL, "(", ")"), ")")
    If arrTmp(0) Like "* FUNCTION *" Then
        strType = "FUNCTION"
    ElseIf arrTmp(0) Like "* PROCEDURE *" Then
        strType = "PROCEDURE"
    End If
    If strType = "" Then Exit Function
    arrTmp1 = Split(arrTmp(0), " " & strType & " ")
    If UBound(arrTmp1) < 1 Then Exit Function
    arrTmp1(1) = Trim(TrimEx(arrTmp1(1), True))
    strName = Split(arrTmp1(1), " ")(0)
    If UBound(arrTmp) = 0 Then '�����ڲ���
        If strType = "FUNCTION" Then
            strReturn = Trim(Split(arrTmp(0), " RETURN ")(1))
        End If
    Else
        If strType = "FUNCTION" And UBound(arrTmp) = 2 Then
            strReturn = Trim(Split(arrTmp(2), "RETURN ")(1))
        End If
        strProcPara = arrTmp(1)
    End If
    
    If blnTruncOwner Then
        If InStr(strName, ".") > 0 Then '��������ǰ׺
            strName = Split(strName, ".")(1)
        End If
    End If
    '��� IN \ OUT  ���� . �ֶ��� % Type : = 'A' ���д���
    '��� IN  OUT   IN \ OUT   Ĭ�� ���д���
    '���Ƚ��ַ���������������������ռλ��ռλ,��ֹ�ַ����������ֽ��
    On Error GoTo errH
    Set rsReturn = CopyNewRec(Nothing, True, , _
                        Array("λ��", adInteger, 3, 0, "����", adVarChar, 50, Empty, _
                              "����", adVarChar, 20, Empty, "��������", adVarChar, 100, Empty, _
                              "����", adVarChar, 20, Empty, "Ĭ��ֵ", adVarChar, 100, Empty))

    rsReturn.AddNew Array("λ��", "����", "����", "��������", "����", "Ĭ��ֵ"), _
                    Array(-1, strName, "", "", strType, "")
    If strReturn <> "" Then
        rsReturn.AddNew Array("λ��", "����", "����", "��������", "����", "Ĭ��ֵ"), _
                        Array(0, "����ֵ", "", "", strReturn, "")
    End If
    arrLine = Split(strProcPara, ","): intIndex = 0
    For i = LBound(arrLine) To UBound(arrLine)
        arrLine(i) = Trim(arrLine(i))
        If arrLine(i) <> "" Then
            '������������
            arrTmp = Split(Trim(arrLine(i)), " ")
            strParName = arrTmp(0): strParDirection = "": strParType = arrTmp(UBound(arrTmp)): strParDefault = "": strParTypeMask = ""
            For j = LBound(arrTmp) + 1 To UBound(arrTmp) - 1
                strParDirection = strParDirection & " " & arrTmp(j)
            Next
            strParDirection = Trim(strParDirection)
            If strParDirection = "IN OUT" Then
                strParDirection = "IN\OUT"
            ElseIf strParDirection = "" Then
                strParDirection = "IN"
            End If
            '��������:=A
            arrTmp = Split(strParType, ":=")
            strParType = arrTmp(0)
            If UBound(arrTmp) > 0 Then
                strParDefault = arrTmp(1)
                If strParDefault Like "[[]*[]]" Then
                    strParDefault = "'" & cllStrs(Mid(strParDefault, 2, Len(strParDefault) - 2)) & "'"
                End If
                If strParDefault = "''" Or strParDefault = "" Then strParDefault = "NULL"
            End If
            '��������.�ֶ���%Type
            If strParType Like "*%TYPE" Then
                strParTypeMask = Mid(strParType, 1, Len(strParType) - 5)
                strParType = ""
            End If
            intIndex = intIndex + 1
            rsReturn.AddNew Array("λ��", "����", "����", "��������", "����", "Ĭ��ֵ"), _
                            Array(intIndex, strParName, strParDirection, strParTypeMask, strParType, strParDefault)
        End If
    Next
    Set AnsyProcedurePara = rsReturn
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description
End Function

Public Function LoadSQL(ByVal strSQL As String, Optional ByVal strLineSplit As String = vbNewLine) As Boolean
'����SQL,����SQL����
'�������塢�������洢���̡�Type��������Ĵ���SQL�Ե���"/"������������������SQL��";"(�ֺ�)����������
    Dim objScript As New clsRunScript
    If objScript.AnalysisSQLString(strSQL, , strLineSplit) Then
        If Not objScript.EOF Then
            Call Me.CopySQL(objScript.SQLInfo)
            LoadSQL = True
        End If
    End If
End Function

