VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
'˵������־��
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'---------------------------------------------------------------

Private mobjFSO As FileSystemObject
Private mstrContent As String

Private mstrPath As String
Private mstrFileName As String
Private mintSaveDays As Integer
Private mblnAction As Boolean
Private mbytType As Byte

'��־���0-��Ҫ��־��1-��ϸ��־
Public Property Get Type_() As Byte
    Type_ = mbytType
End Property
Public Property Let Type_(ByVal BytValue As Byte)
    mbytType = BytValue
End Property

'������־���
Public Property Get Action() As Boolean
    Action = mblnAction
End Property
Public Property Let Action(ByVal blnValue As Boolean)
    mblnAction = blnValue
End Property

'�洢Ч��
Public Property Get SaveDays() As Integer
    SaveDays = mintSaveDays
End Property
Public Property Let SaveDays(ByVal intValue As Integer)
    mintSaveDays = intValue
End Property

'��־�ļ������޺�׺��
Public Property Get FileName() As String
    FileName = mstrFileName
End Property
Public Property Let FileName(ByVal strValue As String)
    mstrFileName = strValue
End Property

'�洢��־·��
Public Property Get Path() As String
    Path = mstrPath
End Property
Public Property Let Path(ByVal strValue As String)
    Dim i As Integer
    Dim strDir As String
    
    mstrPath = strValue
    
    '��־�ļ���Ŀ¼
    For i = 1 To Len(mstrPath)
        If Mid(mstrPath, i, 1) = "\" Or i = Len(mstrPath) Then
            If i = Len(mstrPath) Then
                strDir = mstrPath
            Else
                strDir = Left(mstrPath, i - 1)
            End If
            On Error Resume Next
            If Dir(strDir, vbDirectory) = "" Then
                '����Ŀ¼
                MkDir strDir
            End If
            On Error GoTo 0
        End If
    Next
End Property

Public Sub AddSave(ByVal strContent As String)
'--------------------------------------------------------------------------------
'���ܣ�����������־����
'������
'  strContent����־����
'--------------------------------------------------------------------------------
    If mstrContent = "" Then
        mstrContent = strContent
    Else
        mstrContent = mstrContent & vbNewLine & strContent
    End If
    Call Save
End Sub

Public Sub Add(ByVal strContent As String, Optional ByVal bytLevel As Byte = 0, Optional ByVal bytType As Byte = 0)
'���ܣ���ʱ����־���ݱ���������
'������
'  strContent����־����
'  bytLevel���㼶�ı�����
'  bytType����־���0-��Ҫ��־��1-��ϸ��־

    If mblnAction = False Then Exit Sub

    If mbytType = 0 Then
        If bytType = 0 Then GoSub marAdd
    Else
        GoSub marAdd
    End If

    Exit Sub

marAdd:
    strContent = Replace(strContent, vbNewLine, vbNewLine & String(bytLevel * 2, vbKeySpace))
    mstrContent = mstrContent & String(bytLevel * 2, vbKeySpace) & strContent & vbNewLine
    Return
End Sub

Public Sub AddTime(Optional ByVal bytLevel As Byte = 0, Optional ByVal bytType As Byte = 0)
    Call Me.Add(vbNewLine & "" & Now, bytLevel, bytType)
End Sub

Public Sub Save()
'���ܣ���ʽ����־����������д����־�ļ�
    
    Dim objTS As TextStream
    Dim objFolder As Folder
    Dim objFile As File
    Dim strDate As String, strFileName As String
    Dim i As Integer

    If mblnAction = False Then Exit Sub
    If Trim(mstrContent) = "" Then Exit Sub
    If mstrPath = "" Then Exit Sub
'    If InDesign Then Exit Sub
    
    On Error GoTo hErr
    
    '�Զ�������־�ļ�
    strFileName = mstrFileName & Format(Date, "_yyyymmdd") & ".log"

    '�ռ�����
    Set objTS = mobjFSO.OpenTextFile(mstrPath & "\" & strFileName, ForAppending, True)
    If Right(mstrContent, 1) = vbNewLine Then
        objTS.WriteLine Left(mstrContent, Len(mstrContent) - 1)
    Else
        objTS.WriteLine mstrContent
    End If
    objTS.Close
    
    '���x�������־�ļ�����ɾ��
    Set objFolder = mobjFSO.GetFolder(mstrPath)
    For Each objFile In objFolder.Files
        If LCase(objFile.name) Like LCase(mstrFileName) & "_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].log" Then
            strDate = Split(objFile.name, "_")(1)
            strDate = Split(strDate, ".")(0)
            strDate = Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2)
            If Abs(Date - CDate(strDate)) >= mintSaveDays Then
                On Error Resume Next
                objFile.Delete True
                Err.Clear: On Error GoTo hErr
            End If
        End If
    Next
    
    mstrContent = ""
    Exit Sub
    
hErr:
    mstrContent = ""
End Sub

Private Sub Class_Initialize()
    Set mobjFSO = New FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set mobjFSO = Nothing
End Sub
