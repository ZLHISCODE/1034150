VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function OpenLoginWindow(Optional ByVal strRegPath As String = "") As ADODB.Connection
    Dim intCount As Integer
    Dim strUnitName As String
    Dim blnShowFlash As Boolean
    Dim strStation As String
    
    Set OpenLoginWindow = Nothing

    blnShowFlash = False
    Load SplashObj
    
    strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    
    If strUnitName <> "" Then
        With SplashObj
            .Show
            
            .lblGrant = strUnitName
            strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
            If Trim(strUnitName) = "" Then
                .Label3.Visible = False
                .lbl������.Visible = False
            Else
                .lbl������.Caption = ""
                For intCount = 0 To UBound(Split(strUnitName, ";"))
                    .lbl������.Caption = .lbl������.Caption & Split(strUnitName, ";")(intCount) & vbCrLf
                Next
            End If
            
            .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
            .lbl����֧���� = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
        End With
        
        blnShowFlash = True
        DoEvents
    End If
    
    strStation = Space(200)
    strStation = Trim(strStation)
    If Len(strStation) > 1 Then
        strStation = Left(strStation, Len(strStation) - 1)
    Else
        strStation = "..."
    End If
    
    '�û�ע��
    frmUserLogin.zlShowMe strRegPath
    
    Set OpenLoginWindow = frmUserLogin.mcnOracle
    
    '�ͷ�
    Set frmUserLogin.mcnOracle = Nothing
End Function

Public Function GetDBConnection(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As ADODB.Connection
    Set GetDBConnection = mdlPublic.OraDataOpen(strServerName, strUserName, strUserPwd)
End Function

'�û���
Public Property Get getUserName() As String
    getUserName = gstrUserName
End Property

'ϵͳ��ʾ��Ϣ
Public Property Get getSysName() As String
    getSysName = gstrSysName
End Property