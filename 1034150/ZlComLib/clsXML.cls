VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'######################################################################################################################

Public Enum enumXmlType
    xsString = 1
    xsNumber = 2
    xsDate = 3
    xsTime = 4
    xsDateTime = 5
End Enum

Private mstrXmlText As String                                                       'XML�����¼
Private mintNodeLayer As Integer
Private mobjFso As FileSystemObject
Private mobjTextStream As TextStream
Private mobjXML As Object
Private mstrXmlVersion As String

'######################################################################################################################
'������������

Public Sub ClearXmlText()
    '******************************************************************************************************************
    '���ܣ���ջ����XML�ַ���
    '******************************************************************************************************************
    mintNodeLayer = 1
    mstrXmlText = ""
End Sub

Public Sub AppendNode(ByVal strName As String, Optional ByVal blnCloseNode As Boolean)
    '******************************************************************************************************************
    '���ܣ����������ݵ�XML�ڵ㣬���ϼ��ڵ�
    '������strName - �ڵ�����blnCloseNode - �Ƿ�Ϊ�����ڵ�
    '���أ�
    '******************************************************************************************************************
    
    If blnCloseNode = False Then
        mstrXmlText = mstrXmlText & Space(mintNodeLayer) & "<" & strName & ">" & vbCrLf
        mintNodeLayer = mintNodeLayer + 1
    Else
        mintNodeLayer = mintNodeLayer - 1
        mstrXmlText = mstrXmlText & Space(mintNodeLayer) & "</" & strName & ">" & vbCrLf
    End If
    
End Sub

Public Sub AppendData(ByVal strName As String, ByVal strData As String, Optional ByVal enumType As enumXmlType = xsString)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    strData = ChangeSpecialChar(strData)
    
    Select Case enumType
    Case enumXmlType.xsString
    
    Case enumXmlType.xsNumber
    
    Case enumXmlType.xsDate
        If strData <> "" Then strData = Format(strData, "YYYY-MM-DD")
    Case enumXmlType.xsTime
        If strData <> "" Then strData = Format(strData, "hh:mm:ss")
    Case enumXmlType.xsDateTime
        If strData <> "" Then strData = Format(strData, "YYYY-MM-DDThh:mm:ss")
    End Select
    
    mstrXmlText = mstrXmlText & Space(mintNodeLayer) & "<" & strName & ">" & strData & "</" & strName & ">" & vbCrLf
End Sub

Public Property Get XmlText() As String
    '******************************************************************************************************************
    '���ܣ���ȡ�����XML�ַ������
    '���أ�XML�ַ���
    '******************************************************************************************************************
    XmlText = mstrXmlText
End Property

Public Function OpenXMLDocument(ByVal strXml As String) As Boolean
    '******************************************************************************************************************
    '���ܣ���XML�ĵ�
    '������strXML-XML�ַ���
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    On Error GoTo errHand
    
    mstrXmlVersion = GetXMLVersion
    
    Set mobjXML = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
    
    OpenXMLDocument = mobjXML.loadXML(strXml)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    Set mobjXML = Nothing
    OpenXMLDocument = False
End Function

Public Function GetSingleNodeValue(ByVal strNode As String, ByRef strValue As String, Optional ByVal enumType As enumXmlType = xsString) As Boolean
    '******************************************************************************************************************
    '���ܣ��Ӵ򿪵�XML�ĵ��л�ȡָ���ڵ�ֵ����һ�ڵ㣩
    '������strNode-�ڵ����ƣ���Meta/Release��strValue-���ؽڵ�ֵ
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    Dim strData As String
    Dim objNode As Object
    
    On Error GoTo errHand
        
    If mobjXML Is Nothing Then Exit Function
    
    strNode = ".//" & Replace(strNode, "/", "//")
    
    Set objNode = mobjXML.selectSingleNode(strNode)
    strData = objNode.Text
    
    strValue = RestoreSpecialChar(strData)
    
    Select Case enumType
    Case enumXmlType.xsDateTime
        If strValue <> "" Then strValue = Replace(strValue, "T", " ")
    End Select
    
    GetSingleNodeValue = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    GetSingleNodeValue = False
End Function

Public Function GetMultiNodeRecord(ByVal strNode As String, ByRef rsValue As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ��Ӵ򿪵�XML�ĵ��л�ȡָ���ڵ㼰�ڵ����µ��������ݣ��Լ�¼����ʽ����
    '������strNode-�ڵ����ƣ���Meta/Release��rsValue-���ؽڵ��¼��
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    Dim objNodeList As Object
    Dim lngRow As Long
    
    On Error GoTo errHand
        
    If mobjXML Is Nothing Then Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
    Set rsValue = New ADODB.Recordset
    With rsValue
        .Fields.Append "id", adBigInt, , adFldKeyColumn
        .Fields.Append "parent_id", adBigInt
        .Fields.Append "node_name", adVarChar, 50
        .Fields.Append "node_value", adVarChar, 4000
        .Open
    End With
    
    strNode = ".//" & Replace(strNode, "/", "//")
    
    Set objNodeList = mobjXML.selectNodes(strNode)
    
    lngRow = 0
    If GetXMLNode(rsValue, objNodeList, lngRow, 0) = False Then Exit Function
                       
    If rsValue.RecordCount > 0 Then rsValue.MoveFirst
    
    GetMultiNodeRecord = True
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    GetMultiNodeRecord = False
End Function

Public Function GetAllNodeRecord(ByRef rsValue As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ��Ӵ򿪵�XML�ĵ��л�ȡָ���ڵ㼰�ڵ����µ��������ݣ��Լ�¼����ʽ����
    '������strNode-�ڵ����ƣ���Meta/Release��rsValue-���ؽڵ��¼��
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    Dim objNodeList As Object
    Dim lngRow As Long
    
    On Error GoTo errHand
        
    If mobjXML Is Nothing Then Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
    Set rsValue = New ADODB.Recordset
    With rsValue
        .Fields.Append "id", adBigInt, , adFldKeyColumn
        .Fields.Append "parent_id", adBigInt
        .Fields.Append "node_name", adVarChar, 50
        .Fields.Append "node_value", adVarChar, 4000
        .Open
    End With
    
    Set objNodeList = mobjXML.selectNodes(".//message")
    
    lngRow = 0
    If GetXMLNode(rsValue, objNodeList, lngRow, 0) = False Then Exit Function
                       
    If rsValue.RecordCount > 0 Then rsValue.MoveFirst
    
    GetAllNodeRecord = True
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    GetAllNodeRecord = False
End Function

Public Function CloseXMLDocument() As Boolean
    '******************************************************************************************************************
    '���ܣ��ر�XML�ĵ�
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Set mobjXML = Nothing
    CloseXMLDocument = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    CloseXMLDocument = False
End Function

'######################################################################################################################
Private Function ChangeSpecialChar(ByVal strXmlText As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strXml As String
    
'    &lt;    <   С��
'    &gt;    >   ����
'    &amp;   &   �ͺ�
'    &apos;  '   ������
'    &quot;  "   ����
    
    strXml = strXmlText
    If InStr(strXml, "<") > 0 Then strXml = Replace(strXml, "<", "&lt;")
    If InStr(strXml, ">") > 0 Then strXml = Replace(strXml, ">", "&gt;")
    If InStr(strXml, "&") > 0 Then strXml = Replace(strXml, "&", "&amp;")
    If InStr(strXml, "'") > 0 Then strXml = Replace(strXml, "'", "&apos;")
    If InStr(strXml, """") > 0 Then strXml = Replace(strXml, """", "&quot;")
    
    ChangeSpecialChar = strXml
End Function

Private Function RestoreSpecialChar(ByVal strXmlText As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strXml As String
    
    strXml = strXmlText
    If InStr(strXml, "&lt;") > 0 Then strXml = Replace(strXml, "&lt;", "<")
    If InStr(strXml, "&gt;") > 0 Then strXml = Replace(strXml, "&gt;", ">")
    If InStr(strXml, "&amp;") > 0 Then strXml = Replace(strXml, "&amp;", "&")
    If InStr(strXml, "&apos;") > 0 Then strXml = Replace(strXml, "&apos;", "'")
    If InStr(strXml, "&quot;") > 0 Then strXml = Replace(strXml, "&quot;", """")
    
    RestoreSpecialChar = strXml
End Function

Private Function GetXMLNode(ByRef rsXML As ADODB.Recordset, ByVal objNodeList As Object, ByRef lngRow As Long, ByVal lngParentRow As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objChildNodeList As Object
    Dim objNode As Object
    Dim blnExistChildNode As Boolean
    
    On Error GoTo errHand
    
    For Each objNode In objNodeList
        
        lngRow = lngRow + 1
                                
        rsXML.AddNew
        rsXML("id").value = lngRow
        If lngParentRow <> 0 Then
            rsXML("parent_id").value = lngParentRow
        End If
        rsXML("node_name").value = objNode.nodeName
        
        blnExistChildNode = True
        If objNode.childNodes.Length = 1 Then
            If objNode.childNodes.Item(0).hasChildNodes = False Then
                blnExistChildNode = False
            End If
        End If
        
        If blnExistChildNode = False Then
            rsXML("node_value").value = objNode.Text
        Else
            rsXML("node_value").value = ""
            Set objChildNodeList = objNode.childNodes
            Call GetXMLNode(rsXML, objChildNodeList, lngRow, lngRow)
        End If
    Next
    
    GetXMLNode = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    GetXMLNode = False
End Function


Private Function GetXMLVersion() As String
    
    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXML As Object
    
    On Error GoTo errHand
        
    varXMLVersion = Split(".6.0,.4.0", ",")
    
    On Error Resume Next
    If gobjComLib.OS.IsDesinMode = True Or gobjComLib.zlRegInfo("��Ȩ����") <> "1" Then
        For intLoop = 0 To UBound(varXMLVersion)
            Err = 0
            Set objXML = CreateObject("MSXML2.DOMDocument" & varXMLVersion(intLoop))
            If Err = 0 Then
                strXMLVer = varXMLVersion(intLoop)
                Exit For
            End If
        Next
        On Error GoTo errHand
        
        If strXMLVer = "" Then
            MsgBox "����MSXML2.DOMDocument����ʧ��"
            Exit Function
        End If
    Else
        strXMLVer = ""
    End If
    GetXMLVersion = strXMLVer
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function