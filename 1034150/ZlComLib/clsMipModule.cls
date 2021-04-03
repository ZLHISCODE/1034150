VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'######################################################################################################################

'API
'----------------------------------------------------------------------------------------------------------------------
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'����
'----------------------------------------------------------------------------------------------------------------------
Private mlngSystem As Long                                                          'ϵͳ���
Private mlngModule As Long                                                          'ģ����
Private mstrPrivs As String
Private mstrXmlSchemaFile As String
Private mstrComputer As String
Private mclsMipClient As clsMipClient
Private mstrSendInstance As String
Private mstrMsgKey As String
Private mobjFso As FileSystemObject
Private mobjTextStream As TextStream
Private mobjReceiveObject As Object
Private mstrProgram As String
Private mstrXmlVersion As String
Private mobjCommonMessageSchema As Object
Private WithEvents mfrmMipModule As frmMipModule
Attribute mfrmMipModule.VB_VarHelpID = -1

'�¼�
'----------------------------------------------------------------------------------------------------------------------
Public Event OpenReport(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
Public Event OpenModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
Public Event OpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)
Public Event ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
Public Event ConnectStateChanged(ByVal IsConnected As Boolean)

'######################################################################################################################
'������������
Public Function InitMessage(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPrivs As String, Optional ByVal strProgram As String) As Boolean
'    '******************************************************************************************************************
'    '���ܣ���ʼ��ģ����Ϣ����ʼ��ģ���ڱ�������¼ģ�����ߵȣ�
'    '������lngSystem-ϵͳ��ţ�lngModule-ģ����
'    '���أ���ʼ���ɹ�����True�����򷵻�False
'    '******************************************************************************************************************
'    Dim strSchema As String
'    Dim strSQL As String
'    Dim rsTmp As New ADODB.Recordset
'
'    On Error GoTo errHand
'
'    '------------------------------------------------------------------------------------------------------------------
'    '��ʼ��
'    If gclsMipClient.IsConnect = False Then Exit Function
'    mstrXmlVersion = GetXMLVersion
'    Set mobjReceiveObject = Nothing
'
'    mlngSystem = lngSystem
'    mlngModule = lngModule
'    mstrPrivs = strPrivs
'    mstrProgram = strProgram
'
'    Set mfrmMipModule = New frmMipModule
'    Set mobjFso = New FileSystemObject
'    If mlngModule = 0 Then
'
'        Set mobjCommonMessageSchema = CreateObject("MSXML2.XMLSchemaCache" & mstrXmlVersion)
''        Set mobjCommonMessageSchema = New MSXML2.XMLSchemaCache60
'        mstrXmlSchemaFile = App.Path & "\zlMipCommonMesage.xsd"
'        Set mobjTextStream = mobjFso.CreateTextFile(mstrXmlSchemaFile, True)
'        Call mobjTextStream.Write(GetCommonSchema)
'        mobjTextStream.Close
'    End If
'
'    mstrComputer = AnalyseComputer
'
'    InitMessage = LoginModule(lngSystem, lngModule, mstrSendInstance)
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    InitMessage = False
End Function

Public Function CloseMessage() As Boolean
    '******************************************************************************************************************
    '���ܣ��ر�ģ����Ϣ���ж��ģ���ڱ�������¼ģ�����ߵȣ�
    '���أ��رճɹ�����True�����򷵻�False
    '******************************************************************************************************************
    
'    CloseMessage = LogoutModule(mlngSystem, mlngModule, mstrSendInstance)
'
'    If Not (mfrmMipModule Is Nothing) Then
'        Unload mfrmMipModule
'        Set mfrmMipModule = Nothing
'    End If
End Function

Public Property Get IsConnect() As Boolean
'    '******************************************************************************************************************
'    '���ܣ��ж���Ϣ����ƽ̨�Ƿ�����
'    '���أ������ӷ���True�����򷵻�False
'    '******************************************************************************************************************
'    If Not (mclsMipClient Is Nothing) Then
'        IsConnect = mclsMipClient.IsConnect
'    End If
End Property

Public Function CommitMessage(ByVal strMsgIdentity As String, ByVal strMsgContent As String, Optional ByRef strMsgResult As String) As Boolean
'    '******************************************************************************************************************
'    '���ܣ��ύ��Ϣ��ZLHIS����̨��Ϣ�շ�����
'    '������strMsgIdentity - ��Ϣ��ʶ��strMsgContent - ��Ϣ����(XML)��strMsgResult - ����XML���
'    '���أ��ɹ�����True�����򷵻�False
'    '*****************************************************************************************************************
'    If Not (mclsMipClient Is Nothing) Then
'        CommitMessage = mclsMipClient.SendMessage(strMsgIdentity, strMsgContent, strMsgResult, mlngSystem, mlngModule, mstrSendInstance, mstrProgram)
'    End If
End Function

Public Function ShowMessage(ByVal strMsgKey As String, ByVal strShowContent As String, Optional ByVal strShowTitle As String, Optional ByVal strShowLinkTitle As String, Optional ByVal strShowLinkPara As String) As Boolean
'    '******************************************************************************************************************
'    '���ܣ��ڲ���ϵͳ��������ð��������Ϣ
'    '������strMsgKey-��Ϣid;strShowContent-��ʾ���ݣ�strShowTitle-ð�����ѱ��⣻strShowLinkTitle-������ʾ���⣻strShowLinkPara-������ز���
'    '���أ���ʾ�ɹ�����True�����򷵻�False
'    '*****************************************************************************************************************
'    Dim strXml As String
'
'    If Not (mclsMipClient Is Nothing) Then
'
'        strXml = "<?xml version=""1.0"" encoding=""gbk""?>" & vbCrLf
'        strXml = strXml & "<message>" & vbCrLf
'
'        strXml = strXml & "<topic>" & strShowTitle & "</topic>"
'        strXml = strXml & "<text>" & strShowContent & "</text>"
'        strXml = strXml & "<link>"
'        strXml = strXml & "     <type>99</type>"
'        strXml = strXml & "     <title>" & strShowLinkTitle & "</title>"
'        strXml = strXml & "     <para>" & strShowLinkPara & "</para>"
'        strXml = strXml & "</link>"
'
'        strXml = strXml & "</message>"
'
'        ShowMessage = mclsMipClient.ShowMessage(strMsgKey, strXml, mstrSendInstance)
'    End If
    
End Function

Public Property Set ReceiveObject(objReceiveObject As Object)
'    '******************************************************************************************************************
'    '���ܣ�
'    '���أ�
'    '******************************************************************************************************************
'    Set mobjReceiveObject = objReceiveObject
End Property

Public Property Get Instance() As String
'    '******************************************************************************************************************
'    '���ܣ���ȡʵ��������ʱ��ʵ��ֵ
'    '���أ�
'    '******************************************************************************************************************
'    Instance = mstrSendInstance
End Property

Public Property Get Program() As String
'    '******************************************************************************************************************
'    '���ܣ�
'    '���أ�
'    '******************************************************************************************************************
'    Program = mstrProgram
End Property

'######################################################################################################################
'��Ԫ���������ԣ���ͬһ�������ڿ��Է���
Friend Function ValidCommonMessageSchema(ByVal strXmlMessage As String, Optional ByRef strErrorReason As String) As Boolean
'    '******************************************************************************************************************
'    '���ܣ�У��XML��ʽ
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim strFile As String
'    Dim objXmlSchema As Object
'    Dim objXmlMessage As Object
'
'    Set objXmlSchema = CreateObject("MSXML2.XMLSchemaCache" & mstrXmlVersion)
''    Set objXmlSchema = New MSXML2.XMLSchemaCache60
'    objXmlSchema.Add "", mstrXmlSchemaFile
'
'    Set objXmlMessage = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
''    Set objXmlMessage = New MSXML2.DOMDocument60
'    objXmlMessage.async = False
'    objXmlMessage.validateOnParse = True
'    objXmlMessage.resolveExternals = False
'    Set objXmlMessage.schemas = objXmlSchema
'
''    strFile = "zlCommonMessage.xml"
'    strFile = App.Path & "\" & mobjFso.GetTempName & ".xml"
'
'    strXmlMessage = Replace(LCase(strXmlMessage), "encoding=""utf-8""?", "encoding=""gbk""?")
'
'    Set mobjTextStream = mobjFso.CreateTextFile(strFile, True)
'    mobjTextStream.Write strXmlMessage
'    mobjTextStream.Close
'
'    Call objXmlMessage.Load(strFile)
'    Call objXmlMessage.Validate
'    If objXmlMessage.parseError.errorCode <> 0 Then
'        strErrorReason = objXmlMessage.parseError.reason
'        ValidCommonMessageSchema = False
'    Else
'        ValidCommonMessageSchema = True
'    End If
'
'    Call mobjFso.DeleteFile(strFile, True)
End Function

Friend Sub ReceiveMessage(ByVal strMessageItemKey As String, ByVal strMsgContent As String)
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim objEventPara As Collection
'
'    On Error GoTo errHand
'
'    If Not (mobjReceiveObject Is Nothing) Then
'        Set objEventPara = New Collection
'        Call objEventPara.Add(strMessageItemKey)
'        Call objEventPara.Add(strMsgContent)
'
'        On Error Resume Next
'        Call mobjReceiveObject.ActionEvent("ReceiveMessage", objEventPara)
'
'        Call mobjReceiveObject.ReceiveMessage(strMessageItemKey, strMsgContent)
'    Else
'        RaiseEvent ReceiveMessage(strMessageItemKey, strMsgContent)
'    End If
'
'    Exit Sub
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
    
End Sub

Friend Sub OpenLink(ByVal strMsgKey As String, ByVal bytLinkType As Byte, ByVal strLinkPara As String)
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    mstrMsgKey = strMsgKey
'    If Not (mfrmMipModule Is Nothing) Then
'        Call mfrmMipModule.OpenLink(bytLinkType, strLinkPara)
'    End If
End Sub

Friend Sub ConnectStateChanged(ByVal IsConnected As Boolean)
'    '******************************************************************************************************************
'    '���ܣ���Ϣ����ƽ̨����״̬�仯
'    '������IsConnected-�Ƿ��Ѿ�����
'    '******************************************************************************************************************
'    Dim objEventPara As Collection
'
'    If Not (mobjReceiveObject Is Nothing) Then
'        Set objEventPara = New Collection
'        Call objEventPara.Add(IsConnected)
'        Call mobjReceiveObject.ActionEvent("ConnectStateChanged", objEventPara)
'    Else
'        RaiseEvent ConnectStateChanged(IsConnected)
'    End If
    
End Sub

Friend Property Get System() As Long
    '����:ϵͳ���
'    System = mlngSystem
End Property

Friend Property Get Module() As Long
    '����:ģ����
'    Module = mlngModule
End Property

Friend Property Get Privs() As String
    '����:ģ��Ȩ��
'    Privs = mstrPrivs
End Property

Friend Property Set MipClient(objData As Object)
'    Set mclsMipClient = objData
End Property

Friend Property Get MipClient() As Object
'    Set MipClient = mclsMipClient
End Property

''######################################################################################################################
''˽�з���������
'
'Private Function LoginModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strInstance As String) As Boolean
'    '******************************************************************************************************************
'    '���ܣ�������Ϣ����ƽ̨
'    '��������
'    '˵�����ɹ�����True,���򷵻�False
'    '******************************************************************************************************************
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    If gclsMipClient.IsConnect = False Then Exit Function
'
'    strSQL = "zltools.zlMip_OnLineModule_Insert('" & mstrComputer & "'," & lngSystem & "," & lngModule & ",'" & strInstance & "')"
'
'    Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "LoginModule")
'
'    LoginModule = True
'
'    Exit Function
'errHand:
'
'End Function
'
'Private Function LogoutModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strInstance As String) As Boolean
'    '******************************************************************************************************************
'    '���ܣ�������Ϣ����ƽ̨
'    '��������
'    '˵�����ɹ�����True,���򷵻�False
'    '******************************************************************************************************************
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    If gclsMipClient.IsConnect = False Then Exit Function
'
'    strSQL = "zltools.zlMip_OnLineModule_Delete('" & mstrComputer & "'," & lngSystem & "," & lngModule & ",'" & strInstance & "')"
'
'    Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "LogoutModule")
'
'    LogoutModule = True
'
'    Exit Function
'errHand:
'
'End Function
'
'Private Function GetCommonSchema() As String
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'
'    Dim strSchema As String
'
'    '"<?xml version=""1.0"" encoding=""gbk""?>" & vbNewLine & _
'
'    strSchema = _
'        "<?xml version=""1.0"" encoding=""utf-8""?>" & vbNewLine & _
'        "<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema""  elementFormDefault=""qualified"" attributeFormDefault=""unqualified"">" & vbNewLine & _
'            "<xs:element name=""message"">" & vbNewLine & _
'                "<xs:complexType>" & vbNewLine & _
'                    "<xs:sequence>"
'
'    strSchema = strSchema & _
'                        "<xs:element name=""head"" minOccurs=""1"" maxOccurs=""1"">" & vbNewLine & _
'                            "<xs:complexType>" & vbNewLine & _
'                                "<xs:sequence>" & vbNewLine & _
'                                    "<xs:element name=""zlhis_pub_004"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""msg_id"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""msg_item_code"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_station_name"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_station_ip"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_program"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_instance"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_system_code"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_module_code"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_mipuser"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_hisuser"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_time"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                "</xs:sequence>" & vbNewLine & _
'                            "</xs:complexType>" & vbNewLine & _
'                        "</xs:element>" & vbNewLine
'
'    strSchema = strSchema & _
'                        "<xs:element name=""topic"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1""/>" & vbNewLine & _
'                        "<xs:element name=""text"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine
'
'    strSchema = strSchema & _
'                        "<xs:element name=""link"" minOccurs=""0"" maxOccurs=""1"">" & vbNewLine & _
'                            "<xs:complexType>" & vbNewLine & _
'                                "<xs:sequence>" & vbNewLine & _
'                                    "<xs:element name=""type"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""title"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""para"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1""/>" & vbNewLine & _
'                                "</xs:sequence>" & vbNewLine & _
'                            "</xs:complexType>" & vbNewLine & _
'                        "</xs:element>" & vbNewLine
'
'    strSchema = strSchema & _
'                    "</xs:sequence>" & vbNewLine & _
'                "</xs:complexType>" & vbNewLine & _
'            "</xs:element>" & vbNewLine & _
'        "</xs:schema>"
'
'    GetCommonSchema = strSchema
'End Function
'
'Private Function AnalyseComputer() As String
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim strComputer As String * 256
'
'    Call GetComputerName(strComputer, 255)
'    AnalyseComputer = strComputer
'    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
'End Function
'
'Private Sub Class_Initialize()
'    mstrSendInstance = GetGUID
'End Sub
'
'Private Sub Class_Terminate()
'
'    If Not (mobjFso Is Nothing) Then Set mobjFso = Nothing
'    If Not (mobjTextStream Is Nothing) Then Set mobjTextStream = Nothing
'    If Not (mobjCommonMessageSchema Is Nothing) Then Set mobjCommonMessageSchema = Nothing
'    If Not (mfrmMipModule Is Nothing) Then
'        Unload mfrmMipModule
'        Set mfrmMipModule = Nothing
'    End If
'    If Not (mobjReceiveObject Is Nothing) Then Set mobjReceiveObject = Nothing
'
'End Sub

'Private Sub mfrmMipModule_OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)
'
'    Dim rsTmp As ADODB.Recordset
'    Dim strSQL As String
'    Dim varLinkPara As Variant
'    Dim strLink As String
'    Dim objEventPara As Collection
'
'    Select Case bytLinkType
'    '------------------------------------------------------------------------------------------------------------------
'    Case 1
'        '����
'        strSQL = "Select ϵͳ,����id From zlreports a Where a.���=[1]"
'        varLinkPara = Split(strLinkPara, "|")
'        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, gstrSysName, CStr(varLinkPara(0)))
'        If rsTmp.BOF = False Then
'            If Not (mobjReceiveObject Is Nothing) Then
'                Set objEventPara = New Collection
'                Call objEventPara.Add(bytLinkType)
'                Call objEventPara.Add(gobjComLib.zlCommFun.NVL(rsTmp("ϵͳ").value, 0))
'                Call objEventPara.Add(gobjComLib.zlCommFun.NVL(rsTmp("����id").value, 0))
'                If UBound(varLinkPara) >= 1 Then
'                    strLinkPara = Mid(strLinkPara, InStr(strLinkPara, "|") + 1)
'                    Call objEventPara.Add(strLinkPara)
'                Else
'                    Call objEventPara.Add("")
'                End If
'                Call mobjReceiveObject.ActionEvent("OpenReport", objEventPara)
'            Else
'                If UBound(varLinkPara) >= 1 Then
'                    strLinkPara = Mid(strLinkPara, InStr(strLinkPara, "|") + 1)
'                    RaiseEvent OpenReport(gobjComLib.zlCommFun.NVL(rsTmp("ϵͳ").value, 0), gobjComLib.zlCommFun.NVL(rsTmp("����id").value, 0), strLinkPara)
'                Else
'                    RaiseEvent OpenReport(gobjComLib.zlCommFun.NVL(rsTmp("ϵͳ").value, 0), gobjComLib.zlCommFun.NVL(rsTmp("����id").value, 0), "")
'                End If
'            End If
'        End If
'    '------------------------------------------------------------------------------------------------------------------
'    Case 2
'        'ģ��
'        varLinkPara = Split(strLinkPara, ";")
'        If UBound(varLinkPara) = 1 Then
'            If Not (mobjReceiveObject Is Nothing) Then
'                Set objEventPara = New Collection
'                Call objEventPara.Add(bytLinkType)
'                Call objEventPara.Add(Val(varLinkPara(0)))
'                Call objEventPara.Add(Val(varLinkPara(1)))
'                Call mobjReceiveObject.ActionEvent("OpenModule", objEventPara)
'            Else
'                RaiseEvent OpenModule(Val(varLinkPara(0)), Val(varLinkPara(1)), "")
'            End If
'        End If
'    '------------------------------------------------------------------------------------------------------------------
'    Case 99             '�Զ���
'        If Not (mobjReceiveObject Is Nothing) Then
'            Set objEventPara = New Collection
'            Call objEventPara.Add(mstrMsgKey)
'            Call objEventPara.Add(strLinkPara)
'            Call mobjReceiveObject.ActionEvent("OpenLink", objEventPara)
'        Else
'            RaiseEvent OpenLink(mstrMsgKey, strLinkPara)
'        End If
'    End Select
'End Sub

'Private Function GetXMLNode(ByRef rsXML As ADODB.Recordset, ByVal objNodeList As Object, ByRef lngRow As Long, ByVal lngParentRow As Long) As Boolean
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim objChildNodeList As Object
'    Dim objNode As Object
'
'    On Error GoTo errHand
'
'    For Each objNode In objNodeList
'
'        lngRow = lngRow + 1
'
'        rsXML.AddNew
'        rsXML("id").value = lngRow
'        If lngParentRow > 0 Then rsXML("parent_id").value = lngParentRow
'        rsXML("node_name").value = objNode.nodeName
'
'        If objNode.childNodes.Length = 1 Then
'            rsXML("node_value").value = objNode.Text
'        Else
'            rsXML("node_value").value = ""
'            Set objChildNodeList = objNode.childNodes
'            Call GetXMLNode(rsXML, objChildNodeList, lngRow, lngRow)
'        End If
'    Next
'
'    GetXMLNode = True
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    GetXMLNode = False
'End Function

'Private Function GetXMLVersion() As String
'
'    Dim varXMLVersion As Variant
'    Dim strXMLVer As String
'    Dim intLoop As Integer
'    Dim objXML As Object
'
'    On Error GoTo errHand
'
'    varXMLVersion = Split(".6.0,.4.0", ",")
'
'    On Error Resume Next
'    If gobjComLib.OS.IsDesinMode = True Or gobjComLib.zlRegInfo("��Ȩ����") <> "1" Then
'        For intLoop = 0 To UBound(varXMLVersion)
'            Err = 0
'            Set objXML = CreateObject("MSXML2.DOMDocument" & varXMLVersion(intLoop))
'            If Err = 0 Then
'                strXMLVer = varXMLVersion(intLoop)
'                Exit For
'            End If
'        Next
'        On Error GoTo errHand
'
'        If strXMLVer = "" Then
'            MsgBox "����MSXML2.DOMDocument����ʧ��"
'            Exit Function
'        End If
'    Else
'        strXMLVer = ""
'    End If
'    GetXMLVersion = strXMLVer
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    MsgBox Err.Description
'End Function