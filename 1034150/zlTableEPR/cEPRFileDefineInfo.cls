VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabFileDefineInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�����ļ�����
Public ID As Long               'ID
Public ���� As EPRDocType   '���ࣺ1-���ﲡ��;2-סԺ����;3-������¼;4-��������;5-�������;6-֪���ļ�;7-���Ʊ���;8-��������
Public ��� As String             '�ļ����
Public ���� As String           '�ļ�����
Public ˵�� As String           '�ļ�˵��
Public ҳ�� As String           '�ļ�ҳ�棺ֻ��סԺ�����ͻ�����¼��������ҳ��,�������͵Ĳ���������ҳ,ҳ���ź��ļ����һ��
Public ���� As Long             '�Ƿ�����0-�û�����Ĳ����ļ�;1-ϵͳ��������ͨ�����ļ�,����ɾ��;<0,ָ�������Ƴ���ʵ�ֵ����ⲡ��,����ɾ��
Public ͨ�� As Long             'ͨ���ԣ�0-δ��;1-ͨ��;2-ר��
Public ҳ������ As String       '����ҳ������
Public ���� As Long             '�����Ƶ���:0-�ɱ༭,�����༭��ʽ���;1-�ޱ༭,���Զ��屨�����(�紦����);2-�ɱ༭,���Զ��屨�����(������鱨��)���Ի�����¼:0-���˼�¼;1-������¼��������Ч��

Public lngModule As Long        'ģ���

'��ʽ As String           'PaperKind;PaperOrient;PaperHeight;PaperWidth;MarginTop;MarginBottom;MarginLeft;MarginRight;BackColor;PaperColor;ShowPageNumber
'                            ;ҳüHeadFontName|HeadFontSize|HeadFontBold|HeadFontItalic|HeadFontUnderline|HeadFontStrikethrough|HeadFontColor|HeadMargin
'                            ;ҳ��FootFontName|FootFontSize|FootFontBold|FootFontItalic|FootFontUnderline|FootFontStrikethrough|FootFontColor|FootMargin
'                            ;HCenter;VCenter
Private mvarPaperKind As Integer
Private mvarPaperOrient As Integer
Private mvarPaperHeight As Long
Private mvarPaperWidth  As Long
Private mvarMarginLeft As Long
Private mvarMarginRight As Long
Private mvarMarginTop As Long
Private mvarMarginBottom As Long
Private mvarHCenter As Byte '�������
Private mvarVCenter As Byte '�������

'ҳü                   '
Private mvarHeadConText As String          '����(Ҫ����[]����),ҳü����
Private mvarHeadFontName As String
Private mvarHeadFontSize As Long
Private mvarHeadFontBold As Boolean
Private mvarHeadFontItalic As Boolean
Private mvarHeadFontUnderline As Boolean
Private mvarHeadFontStrikethrough As Boolean
Private mvarHeadFontColor As OLE_COLOR
Private mvarHeadMargin As Long
Private mvarHeadPic As StdPicture

'ҳ��                  '
Private mvarFootConText As String               '����(Ҫ����[]����),ҳ������
Private mvarFootFontName As String
Private mvarFootFontSize As Long
Private mvarFootFontBold As Boolean
Private mvarFootFontItalic As Boolean
Private mvarFootFontUnderline As Boolean
Private mvarFootFontStrikethrough As Boolean
Private mvarFootFontColor As OLE_COLOR
Private mvarFootMargin As Long
Public Sub GetFileDefineInfo(ByVal lngFileID As Long)
Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
    gstrSQL = "Select a.ID, a.����, a.���, a.����, a.˵��, a.ҳ��, a.����, a.ͨ��, b.���� As ҳ������, b.����, b.��ʽ, b.ҳü, b.ҳ�� " & _
            " From �����ļ��б� a, ����ҳ���ʽ b " & _
            " Where a.ҳ�� = b.��� And a.���� = b.���� And a.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cTableEPR", lngFileID)
    With Me
        .ID = lngFileID
        If Not rsTemp.EOF Then
            .���� = Nvl(rsTemp!����, 1)  '1-���ﲡ��;2-סԺ����;3-������¼;4-��������;5-����֤������;6-֪���ļ�;7-���Ʊ���;8-��������
            .��� = Nvl(rsTemp!���)
            .���� = Nvl(rsTemp!����)
            .˵�� = Nvl(rsTemp!˵��)
            .ҳ�� = Nvl(rsTemp!ҳ��)
            .���� = Nvl(rsTemp!����, 0)  '��������2-����ʽ����
            .ͨ�� = Nvl(rsTemp!ͨ��, 0)  '�����Ƶ�������Ч:0-δ��;1-ͨ��;2-ר��
            .���� = Nvl(rsTemp!����, 0)
            .ҳ������ = Nvl(rsTemp!ҳ������)
            .��ʽ = Nvl(rsTemp!��ʽ, "9;1;16840;11907;482;805;350;350;16777215;16777215;1;����|9|False|False|False|False|vbBlack|350;����|9|False|False|False|False|vbBlack|350;0;0") 'ȱʡ A4 ���� 297 210
            .ҳü = Nvl(rsTemp!ҳü, "")
            .ҳ�� = Nvl(rsTemp!ҳ��, "")
            
            Dim strFile As String, strZipFile As String
            strZipFile = zlBlobRead(7, Me.���� & "-" & Me.ҳ��, strZipFile)
            If strZipFile = "" Then Exit Sub
            If gobjFSO.FileExists(strZipFile) Then
                strFile = zlFileUnzip(strZipFile, "Jpg")
                gobjFSO.DeleteFile strZipFile, True
                If strFile = "" Then Exit Sub
                
                If gobjFSO.FileExists(strFile) Then
                    Set Me.HeadPic = LoadPicture(strFile)
                    gobjFSO.DeleteFile strFile, True
                End If
            End If
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Property Let PaperKind(vData As Integer)
    mvarPaperKind = vData
End Property
Public Property Get PaperKind() As Integer
    PaperKind = mvarPaperKind
End Property
Public Property Let PaperOrient(vData As Integer)
    mvarPaperOrient = vData
End Property
Public Property Get PaperOrient() As Integer
    PaperOrient = mvarPaperOrient
End Property

Public Property Let PaperHeight(vData As Long)
    mvarPaperHeight = vData
End Property
Public Property Get PaperHeight() As Long
    PaperHeight = mvarPaperHeight
End Property
Public Property Let PaperWidth(vData As Long)
    mvarPaperWidth = vData
End Property
Public Property Get PaperWidth() As Long
    PaperWidth = mvarPaperWidth
End Property
Public Property Let MarginTop(vData As Long)
    mvarMarginTop = vData
End Property
Public Property Get MarginTop() As Long
    MarginTop = mvarMarginTop
End Property
Public Property Let MarginBottom(vData As Long)
    mvarMarginBottom = vData
End Property
Public Property Get MarginBottom() As Long
    MarginBottom = mvarMarginBottom
End Property
Public Property Let MarginLeft(vData As Long)
    mvarMarginLeft = vData
End Property
Public Property Get MarginLeft() As Long
    MarginLeft = mvarMarginLeft
End Property
Public Property Let MarginRight(vData As Long)
    mvarMarginRight = vData
End Property
Public Property Get MarginRight() As Long
    MarginRight = mvarMarginRight
End Property
Public Property Let HCenter(vData As Byte)
    mvarHCenter = vData
End Property
Public Property Get HCenter() As Byte
    HCenter = mvarHCenter
End Property
Public Property Let VCenter(vData As Byte)
    mvarVCenter = vData
End Property
Public Property Get VCenter() As Byte
    VCenter = mvarVCenter
End Property
Public Property Let ��ʽ(vData As String)
On Error Resume Next
    Dim T As Variant, THead As Variant, TFoot As Variant
    T = Split(vData, ";")
    mvarPaperKind = T(0)
    mvarPaperOrient = T(1)
    mvarPaperHeight = T(2)
    mvarPaperWidth = T(3)
    mvarMarginTop = T(4)
    mvarMarginBottom = T(5)
    mvarMarginLeft = T(6)
    mvarMarginRight = T(7)
    
    THead = Split(T(11), "|")
    mvarHeadFontName = UsableFont(THead(0))
    mvarHeadFontSize = THead(1)
    mvarHeadFontBold = THead(2)
    mvarHeadFontItalic = THead(3)
    mvarHeadFontUnderline = THead(4)
    mvarHeadFontStrikethrough = THead(5)
    mvarHeadFontColor = THead(6)
    mvarHeadMargin = THead(7)
    If mvarHeadMargin = 0 Then
        mvarHeadMargin = mvarMarginTop
    End If
    
    TFoot = Split(T(12), "|")
    mvarFootFontName = UsableFont(TFoot(0))
    mvarFootFontSize = TFoot(1)
    mvarFootFontBold = TFoot(2)
    mvarFootFontItalic = TFoot(3)
    mvarFootFontUnderline = TFoot(4)
    mvarFootFontStrikethrough = TFoot(5)
    mvarFootFontColor = TFoot(6)
    mvarFootMargin = TFoot(7)
    If mvarFootMargin = 0 Then
        mvarFootMargin = mvarMarginBottom
    End If
    mvarHCenter = T(13)
    mvarVCenter = T(14)
    Err.Clear
End Property
Public Property Get ��ʽ() As String
    ��ʽ = mvarPaperKind & ";" & mvarPaperOrient & ";" & mvarPaperHeight & ";" & mvarPaperWidth & ";" & mvarMarginTop & ";" & mvarMarginBottom & ";" & mvarMarginLeft & ";" & mvarMarginRight & ";16777215;16777215;1" & _
        ";" & mvarHeadFontName & "|" & mvarHeadFontSize & "|" & mvarHeadFontBold & "|" & mvarHeadFontItalic & "|" & mvarHeadFontUnderline & "|" & mvarHeadFontStrikethrough & "|" & mvarHeadFontColor & "|" & mvarHeadMargin & _
        ";" & mvarFootFontName & "|" & mvarFootFontSize & "|" & mvarFootFontBold & "|" & mvarFootFontItalic & "|" & mvarFootFontUnderline & "|" & mvarFootFontStrikethrough & "|" & mvarFootFontColor & "|" & mvarFootMargin & ";" & mvarHCenter & ";" & mvarVCenter
End Property
Public Property Get HeadConText() As String
    HeadConText = mvarHeadConText
End Property
Public Property Let HeadConText(vData As String)
    mvarHeadConText = vData
End Property
Public Property Get HeadFontName() As String
    HeadFontName = mvarHeadFontName
End Property
Public Property Let HeadFontName(vData As String)
    mvarHeadFontName = vData
End Property
Public Property Get HeadFontSize() As Long
    HeadFontSize = mvarHeadFontSize
End Property
Public Property Let HeadFontSize(vData As Long)
    mvarHeadFontSize = vData
End Property
Public Property Get HeadFontBold() As Boolean
    HeadFontBold = mvarHeadFontBold
End Property
Public Property Let HeadFontBold(vData As Boolean)
    mvarHeadFontBold = vData
End Property
Public Property Get HeadFontItalic() As Boolean
    HeadFontItalic = mvarHeadFontItalic
End Property
Public Property Let HeadFontItalic(vData As Boolean)
    mvarHeadFontItalic = vData
End Property
Public Property Get HeadFontUnderline() As Boolean
    HeadFontUnderline = mvarHeadFontUnderline
End Property
Public Property Let HeadFontUnderline(vData As Boolean)
    mvarHeadFontUnderline = vData
End Property
Public Property Get HeadFontStrikethrough() As Boolean
    HeadFontStrikethrough = mvarHeadFontStrikethrough
End Property
Public Property Let HeadFontStrikethrough(vData As Boolean)
    mvarHeadFontStrikethrough = vData
End Property
Public Property Get HeadFontColor() As OLE_COLOR
    HeadFontColor = mvarHeadFontColor
End Property
Public Property Let HeadFontColor(vData As OLE_COLOR)
    mvarHeadFontColor = vData
End Property
Public Property Get HeadMargin() As Long
    HeadMargin = mvarHeadMargin
End Property
Public Property Let HeadMargin(vData As Long)
    mvarHeadMargin = vData
    If mvarHeadMargin = 0 Then
        If mvarPaperOrient = 1 Then mvarHeadMargin = mvarMarginTop Else mvarHeadMargin = mvarMarginLeft
    End If
End Property
Public Property Get ҳü() As String
    ҳü = mvarHeadConText
End Property
Public Property Let ҳü(vData As String)
    mvarHeadConText = vData
End Property
Public Property Get FootConText() As String
    FootConText = mvarFootConText
End Property
Public Property Let FootConText(vData As String)
    mvarFootConText = vData
End Property
Public Property Get FootFontName() As String
    FootFontName = mvarFootFontName
End Property
Public Property Let FootFontName(vData As String)
    mvarFootFontName = vData
End Property
Public Property Get FootFontSize() As Long
    FootFontSize = mvarFootFontSize
End Property
Public Property Let FootFontSize(vData As Long)
    mvarFootFontSize = vData
End Property
Public Property Get FootFontBold() As Boolean
    FootFontBold = mvarFootFontBold
End Property
Public Property Let FootFontBold(vData As Boolean)
    mvarFootFontBold = vData
End Property
Public Property Get FootFontItalic() As Boolean
    FootFontItalic = mvarFootFontItalic
End Property
Public Property Let FootFontItalic(vData As Boolean)
    mvarFootFontItalic = vData
End Property
Public Property Get FootFontUnderline() As Boolean
    FootFontUnderline = mvarFootFontUnderline
End Property
Public Property Let FootFontUnderline(vData As Boolean)
    mvarFootFontUnderline = vData
End Property
Public Property Get FootFontStrikethrough() As Boolean
    FootFontStrikethrough = mvarFootFontStrikethrough
End Property
Public Property Let FootFontStrikethrough(vData As Boolean)
    mvarFootFontStrikethrough = vData
End Property
Public Property Get FootFontColor() As OLE_COLOR
    FootFontColor = mvarFootFontColor
End Property
Public Property Let FootFontColor(vData As OLE_COLOR)
    mvarFootFontColor = vData
End Property
Public Property Get FootMargin() As Long
    FootMargin = mvarFootMargin
End Property
Public Property Let FootMargin(vData As Long)
    mvarFootMargin = vData
    If mvarFootMargin = 0 Then
        If mvarPaperOrient = 1 Then mvarFootMargin = mvarMarginBottom Else mvarFootMargin = mvarMarginRight
    End If
End Property
Public Property Get ҳ��() As String
    ҳ�� = mvarFootConText
End Property
Public Property Let ҳ��(vData As String)
    mvarFootConText = vData
End Property
Public Property Get HeadPic() As StdPicture
    Set HeadPic = mvarHeadPic
End Property
Public Property Set HeadPic(ByVal vData As StdPicture)
    Set mvarHeadPic = vData
End Property
Public Property Let HeadPic(ByVal vData As StdPicture)
    Set mvarHeadPic = vData
End Property
Public Function GetSaveSQL(arrSQL As Variant) As Boolean
    On Error GoTo errHand
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_����ҳ���ʽ_Update(" & Me.���� & ",'" & Me.��� & "','" & Me.ҳ������ & "'," & _
                                            Me.���� & ",'" & Me.��ʽ & "','" & Me.ҳü & "','" & Me.ҳ�� & "')"
                                            
    If mvarHeadPic.Handle <> 0 Then '��ҳüͼ
        Dim strFile As String, strZipFile As String
        strFile = App.Path & "\Tmp.Jpg"
        If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True
        SavePicture mvarHeadPic, strFile
        strZipFile = zlFileZip(strFile)
        If strZipFile <> "" Then gobjFSO.DeleteFile strFile, True
        If gobjFSO.FileExists(strZipFile) Then
            Call zlBlobSql(7, Me.���� & "-" & Me.ҳ��, strZipFile, arrSQL)
            gobjFSO.DeleteFile strZipFile, True      'ɾ����ʱ�ļ�
        End If
    End If

    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_�����ļ��ṹ_Commit(" & Me.ID & ")"
    
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Class_Initialize()
    Set mvarHeadPic = New StdPicture
     mvarHeadFontName = "����"
     mvarHeadFontSize = 9
     mvarFootFontName = "����"
     mvarFootFontSize = 9
End Sub

Private Sub Class_Terminate()
    Set mvarHeadPic = Nothing
End Sub