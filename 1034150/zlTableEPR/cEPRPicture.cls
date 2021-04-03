VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarKey As Long                     '�����еĹؼ���
Private mvarPicID As Long                   '����ͼ��ID,���Ӳ���ͼ��(����ID)-���Ӳ�������ID
Private mvarOrigPic As StdPicture           'ԭʼͼƬ
Private mvarOrigWidth As Long               'ͼƬԭʼ����
Private mvarOrigHeight As Long              'ͼƬԭʼ�߶�
Private mvarDesWidth As Long                'Ŀ�����
Private mvarDesHeight As Long               'Ŀ��߶�

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let PicID(ByVal vData As Long)
    mvarPicID = vData
End Property

Public Property Get PicID() As Long
    PicID = mvarPicID
End Property

Public Property Set OrigPic(ByRef vData As StdPicture)
    Set mvarOrigPic = vData
    If mvarOrigPic.Handle <> 0 Then
        mvarOrigHeight = mvarOrigPic.Height
        mvarOrigWidth = mvarOrigPic.Width
    End If
End Property

Public Property Let OrigPic(ByRef vData As StdPicture)
    Set mvarOrigPic = vData
End Property

Public Property Get OrigPic() As StdPicture
    Set OrigPic = mvarOrigPic
End Property
Public Property Let DesWidth(ByVal vData As Long)
    mvarDesWidth = vData
End Property

Public Property Get DesWidth() As Long
    DesWidth = mvarDesWidth
End Property

Public Property Let DesHeight(ByVal vData As Long)
    mvarDesHeight = vData
End Property

Public Property Get DesHeight() As Long
    DesHeight = mvarDesHeight
End Property
Public Property Let OrigWidth(ByVal vData As Long)
    mvarOrigWidth = vData
End Property

Public Property Get OrigWidth() As Long
    OrigWidth = mvarOrigWidth
End Property

Public Property Let OrigHeight(ByVal vData As Long)
    mvarOrigHeight = vData
End Property

Public Property Get OrigHeight() As Long
    OrigHeight = mvarOrigHeight
End Property
'################################################################################################################
'## ���ܣ�  ����ͼƬ����
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Sub Clone(cNode As cTabPicture)
    With cNode
        Set .OrigPic = Me.OrigPic
        .OrigWidth = Me.OrigWidth
        .OrigHeight = Me.OrigHeight
    End With
End Sub
'################################################################################################################
'## ���ܣ�  ��ȡͼƬ�����ַ���
'################################################################################################################
Public Property Get ��������() As String
    Dim strR As String
    strR = Me.OrigWidth & ";" & Me.OrigHeight & ";" & Me.DesWidth & ";" & Me.DesHeight
    �������� = strR
End Property

'################################################################################################################
'## ���ܣ�  ͨ��һ���ַ���������ͼƬ����
'################################################################################################################
Public Property Let ��������(vData As String)
On Error Resume Next
    Dim T As Variant, i As Long
    T = Split(vData, ";")
    If UBound(T) > 0 Then
        Me.OrigWidth = IIf(T(0) = "", 0, T(0))
        Me.OrigHeight = IIf(T(1) = "", 0, T(1))
        Me.DesWidth = IIf(T(0) = "", 0, T(2))
        Me.DesHeight = IIf(T(0) = "", 0, T(3))
    End If
End Property


'################################################################################################################
'## ���ܣ�  �����ݿ�ֱ�ӻ�ȡͼƬ����
'##
'## ������  eEdtType    :��ǰ�༭��ʽ
'################################################################################################################
Public Function GetPictureFromDB(ByVal EM As TabEM, ByVal editType As TabET, ByVal blnMoved As Boolean, ByRef tmpMark As cTabPicMarks, ByRef tmpMarkKeys As String) As StdPicture
Dim strFile As String, strZipFile As String, rs As New ADODB.Recordset
    Set GetPictureFromDB = New StdPicture
    If EM = TabEm_���� Then
        strZipFile = zlBlobRead(2, Me.PicID)
    Else
        Select Case editType
        Case TabET_�����ļ�����
            strZipFile = zlBlobRead(2, Me.PicID)
        Case TabET_ȫ��ʾ���༭
            gstrSQL = "Select ID from ������������ Where ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "", Me.PicID)
            If rs.EOF Then
                strZipFile = zlBlobRead(2, Me.PicID)
            Else
                strZipFile = zlBlobRead(4, Me.PicID)
            End If
        Case TabET_�������༭, TabET_���������
            strZipFile = zlBlobRead(6, Me.PicID, , blnMoved)  '����������ʷ��ʱ,ͼ������Ҳ��ת��
        End Select
    End If
    If strZipFile = "" Then Exit Function
    If gobjFSO.FileExists(strZipFile) Then
        strFile = zlFileUnzip(strZipFile, "Jpg")
        gobjFSO.DeleteFile strZipFile, True
        If strFile = "" Then Exit Function
        
        If gobjFSO.FileExists(strFile) Then
            Set Me.OrigPic = LoadPicture(strFile)
            gobjFSO.DeleteFile strFile, True
        End If
    End If


    Dim strTableName As String, strPicMark As String, i As Integer, strMarkKey As String, lngKey As Long, int���� As Integer, int��ʼ�� As Integer, int��ֹ�� As Integer
    If EM = TabEm_���� Then
        strTableName = "�����ļ��ṹ"
    Else
        Select Case editType
        Case TabET_�����ļ�����
            strTableName = "�����ļ��ṹ"
        Case TabET_ȫ��ʾ���༭
            gstrSQL = "Select ID from ������������ Where ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "", Me.PicID)
            If rs.EOF Then
                strTableName = "�����ļ��ṹ"
            Else
                strTableName = "������������"
            End If
        Case TabET_�������༭, TabET_���������
            strTableName = IIf(blnMoved, "H���Ӳ�������", "���Ӳ�������")
        End Select
    End If
    
    If InStr(strTableName, "���Ӳ�������") > 0 Then '��ȡ��ͬ��Ԫ����ֹ��=0,�����õ�Ԫ����Ŀǰ���°�
        gstrSQL = "Select *" & vbNewLine & _
                "From " & strTableName & vbNewLine & _
                "Where ��id = [1]" & vbNewLine & _
                "Order By �������, �����д�"
    Else
        gstrSQL = "Select * From " & strTableName & " Where ��id = [1] Order By �������, �����д�"
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cTableEPR", Me.PicID)
    Do Until rs.EOF
        strPicMark = strPicMark & rs!�����ı�
        int���� = Nvl(rs!��������, 0)
        If InStr(strTableName, "���Ӳ�������") > 0 Then
            int��ʼ�� = Nvl(rs!��ʼ��, 1)
            int��ֹ�� = Nvl(rs!��ֹ��, 0)
        End If
        rs.MoveNext
    Loop
    If strPicMark <> "" Then
        For i = 0 To UBound(Split(strPicMark, "||"))
            lngKey = tmpMark.Add
            tmpMarkKeys = tmpMarkKeys & "|" & lngKey
            With tmpMark("K" & lngKey)
                .��ID = Me.PicID
                .ѡ�� = Split(Split(strPicMark, "||")(i), "|")(0)
                .���� = Split(Split(strPicMark, "||")(i), "|")(1)
                .���� = Split(Split(strPicMark, "||")(i), "|")(2)
                .�㼯 = Split(Split(strPicMark, "||")(i), "|")(3)
                .X1 = Split(Split(strPicMark, "||")(i), "|")(4)
                .Y1 = Split(Split(strPicMark, "||")(i), "|")(5)
                .X2 = Split(Split(strPicMark, "||")(i), "|")(6)
                .Y2 = Split(Split(strPicMark, "||")(i), "|")(7)
                .���ɫ = Split(Split(strPicMark, "||")(i), "|")(8)
                .��䷽ʽ = Split(Split(strPicMark, "||")(i), "|")(9)
                .����ɫ = Split(Split(strPicMark, "||")(i), "|")(10)
                .����ɫ = Split(Split(strPicMark, "||")(i), "|")(11)
                .���� = Split(Split(strPicMark, "||")(i), "|")(12)
                .�߿� = Split(Split(strPicMark, "||")(i), "|")(13)
                .���� = Split(Split(strPicMark, "||")(i), "|")(14)
                .�������� = IIf(int���� = 1, True, False)
                .��ʼ�� = int��ʼ��
                .��ֹ�� = int��ֹ��
            End With
        Next
    End If
    Set GetPictureFromDB = mvarOrigPic
End Function
Public Function GetSaveSQL(ByVal editType As TabET, ByVal CellTmp As cTabCell, ByVal PicMarksTmp As cTabPicMarks, arrSQL As Variant) As Boolean
'����:����ͼƬ��Ԫ��,ͼƬ����,ͼƬ��Ǳ���SQL
    On Error GoTo errHand
    If editType = TabET_��������� Then GetSaveSQL = True: Exit Function '��˲��ܱ༭ͼƬ
    With CellTmp
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        Select Case editType
        Case TabET_�����ļ�����
            arrSQL(UBound(arrSQL)) = "Zl_�����ļ��ṹ_Update(" & .ID & "," & .�ļ�ID & _
                                        ",Null," & .������� & "," & .�������� & ",Null," & IIf(.��������, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.�������� & "',Null,'" & .�����ı� & "')"
        Case TabET_ȫ��ʾ���༭
            arrSQL(UBound(arrSQL)) = "Zl_������������_Update(" & .ID & "," & .�ļ�ID & _
                                        ",Null," & .������� & "," & .�������� & ",Null," & IIf(.��������, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.�������� & "',Null,'" & .�����ı� & "')"
        Case TabET_�������༭, TabET_���������
            arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�������_Update(" & .ID & "," & .�ļ�ID & "," & .��ʼ�� & "," & .��ֹ�� & _
                                        ",Null," & .������� & "," & .�������� & ",Null," & IIf(.��������, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.�������� & "',Null,'" & .�����ı� & "')"
        End Select
        
        If mvarOrigPic.Handle <> 0 Then '����ͼƬ'��˲��ܱ༭ͼƬ
            Dim strFile As String, strZipFile As String
            strFile = App.Path & "\Tmp.Jpg"
            If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True

            Dim GpInput As GdiplusStartupInput
            Dim m_GDIpToken         As Long         ' ���ڹر� GDI+
            Dim oDIB As New cDIB
            Dim DIBDither As cDIBDither
            Dim DIBPal As New cDIBPal
            '-- ���� GDI+ Dll
            GpInput.GdiplusVersion = 1
            If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
                '����BMP��ʽ���棡������ͼƬ���
                SavePicture mvarOrigPic, strFile       '�����ʽΪBMP��ʽ
            Else
                '����JPEGѹ����ʽ����
                Call oDIB.CreateFromStdPicture(mvarOrigPic, DIBPal, DIBDither)
                'ѹ���洢
                Call mGdIpEx.SaveDIB(oDIB, strFile, [ImageJPEG], 100)           '90%��JPEGͼƬѹ������
            End If
            Call mGdIpEx.GdiplusShutdown(m_GDIpToken)
            
            strZipFile = zlFileZip(strFile)     'ѹ���ļ�
            gobjFSO.DeleteFile strFile, True 'ɾ��ԭʼ�ļ�
            If strZipFile <> "" Then
                If gobjFSO.FileExists(strZipFile) Then
                    Select Case editType
                        Case TabET_�����ļ�����
                            Call zlBlobSql(2, .ID, strZipFile, arrSQL)
                        Case TabET_ȫ��ʾ���༭
                            Call zlBlobSql(4, .ID, strZipFile, arrSQL)
                        Case TabET_�������༭, TabET_���������
                            Call zlBlobSql(6, .ID, strZipFile, arrSQL)
                    End Select
                    gobjFSO.DeleteFile strZipFile, True      'ɾ����ʱ�ļ�
                End If
            End If
        End If
        
        If .PicMarkKey <> "" Then '������
            
            Dim i As Long, strPicMark As String, strSub As String, lngLen As Long, lngSub�д� As Long, lngID As Long
            For i = 1 To UBound(Split(.PicMarkKey, "|"))
                With PicMarksTmp("K" & Split(.PicMarkKey, "|")(i))
                    strPicMark = strPicMark & "||" & .��������
                End With
            Next
            
            strPicMark = Mid(strPicMark, 3)
            lngLen = Len(strPicMark)
            i = 0    '����4000Ϊ��ֶδ洢��
            lngSub�д� = 1
            Do While (i * 4000 + 1 <= lngLen)
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                strSub = Mid(strPicMark, i * 4000 + 1, 4000)
                Select Case editType
                Case TabET_�����ļ�����
                    lngID = zlDatabase.GetNextId("�����ļ��ṹ")
                    arrSQL(UBound(arrSQL)) = "Zl_�����ļ��ṹ_Update(" & lngID & "," & CellTmp.�ļ�ID & "," & .ID & "," & _
                                            CellTmp.������� & "," & cprCTPicMark & ",Null," & IIf(.��������, 1, 0) & ",'" & _
                                            CellTmp.Row & "," & CellTmp.Col & "|'," & lngSub�д� & ",'" & strSub & "')"
                Case TabET_ȫ��ʾ���༭
                    lngID = zlDatabase.GetNextId("������������")
                    arrSQL(UBound(arrSQL)) = "Zl_������������_Update(" & lngID & "," & CellTmp.�ļ�ID & "," & .ID & "," & _
                                            CellTmp.������� & "," & cprCTPicMark & ",Null," & IIf(.��������, 1, 0) & ",'" & _
                                            CellTmp.Row & "," & CellTmp.Col & "|'," & lngSub�д� & ",'" & strSub & "')"
                Case TabET_�������༭, TabET_���������
                    lngID = zlDatabase.GetNextId("���Ӳ�������")
                    arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�������_Update(" & lngID & "," & CellTmp.�ļ�ID & "," & .��ʼ�� & "," & .��ֹ�� & "," & .ID & "," & _
                                            CellTmp.������� & "," & cprCTPicMark & ",Null," & IIf(.��������, 1, 0) & ",'" & _
                                            CellTmp.Row & "," & CellTmp.Col & "|'," & lngSub�д� & ",'" & strSub & "')"
                End Select
                lngSub�д� = lngSub�д� + 1
                i = i + 1
            Loop
        End If
    End With
    
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub Class_Initialize()
    Set mvarOrigPic = New StdPicture
End Sub

Private Sub Class_Terminate()
    Set mvarOrigPic = Nothing
End Sub