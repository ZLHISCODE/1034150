VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmClientCopy 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Զ�����"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   ControlBox      =   0   'False
   Icon            =   "frmClientCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6900
   StartUpPosition =   2  '��Ļ����
   Begin zlHisCrust.UsrProgressBar prgPross 
      Height          =   300
      Left            =   45
      TabIndex        =   4
      Top             =   1245
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   529
      Color           =   12937777
      Value           =   100
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "�鿴��־(&C)"
      Height          =   375
      Left            =   3615
      TabIndex        =   3
      Top             =   4545
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   450
      Left            =   5145
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "���(&O)"
      Height          =   375
      Left            =   5220
      TabIndex        =   1
      Top             =   4545
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwMan 
      Height          =   2430
      Left            =   60
      TabIndex        =   0
      Top             =   2040
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   4286
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img2"
      SmallIcons      =   "img2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "������Ϣ"
         Object.Width           =   7585
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�ְ汾��"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ԭ�汾��"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "���޸�����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ԭ�޸�����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ҵ�񲿼�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "��װ·��"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "MD5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "�Զ�����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ǿ�Ƹ���"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList img2 
      Left            =   5535
      Top             =   1815
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":030A
            Key             =   "Ok"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":08A4
            Key             =   "Err"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":0E3E
            Key             =   "List"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   4725
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":13D8
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":16F2
            Key             =   "Err"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":1A0C
            Key             =   "List"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ�����������,���Ժ�..."
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1815
      TabIndex        =   6
      Top             =   495
      Width           =   4020
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ע�Ჿ��"
      Height          =   180
      Left            =   60
      TabIndex        =   5
      Top             =   1710
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   315
      Picture         =   "frmClientCopy.frx":1B66
      Top             =   255
      Width           =   720
   End
End
Attribute VB_Name = "frmClientCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnFirst As Boolean
Dim lngCount As Long
Dim mintColumn As Integer
Dim blnOk As Boolean
Dim blnAutoRun As Boolean '�Ƿ������������ļ�
Dim strAutoRun As String '�������ļ�·��
Dim strAutoRunBat As String


Private Sub cmdLog_Click()
    On Error Resume Next
    Dim ret As Long
    Dim strLogFile As String
    Dim strNotPad As String
    strNotPad = GetWinSystemPath & "\notepad.exe"
    If mobjFile.FileExists(strNotPad) Then
        strLogFile = gstrAppPath & "\ZLUpGradeList.lst"
        ret = ShellExecute(0&, "open", strNotPad, strLogFile, strLogFile, 5)    'SW_SHOW
        If ret = 31 Then
           MsgBox "û���ҵ��ʵ��ĳ���������,�밲װ��Ч�ĳ���!", vbInformation, "�ͻ����Զ�����"
        End If
    Else
        MsgBox "����û�а�װ���±�����,���ܴ���־�ļ�!" & vbCrLf & "���ֹ������������,���±�·��Ϊ:" & vbCrLf & strLogFile, vbInformation, "�ͻ����Զ�����"
    End If
End Sub

Private Sub cmdOK_Click()
    'ȷ���Ƿ񻹴�������
    
    
    
    If IsUpgrade = True Then
        '--------------------------------------------------------
        '�������ɹ�
        WriteTxtLog ""
        WriteTxtLog ""
        WriteTxtLog "--" & Format(Now(), "yyyy-mm-dd HH:MM:SS")
        WriteTxtLog "=============================================������һ�������������ռ����ɹ�======================================================================================================================================="
        Call SaveClientLog("������һ�������������ռ����ɹ�")
        Call UpdateCondition(2)
        Call CallHISEXE(False)
    Else
        WriteTxtLog ""
        WriteTxtLog ""
        WriteTxtLog "--" & Format(Now(), "yyyy-mm-dd HH:MM:SS")
        WriteTxtLog "=============================================�������ռ��ɹ�======================================================================================================================================="
        'ʵ���Զ�����������
        'ִ��HIS����
        Call SaveClientLog("�������ռ��ɹ�")
        Call UpdateCondition(1)
        Call CallHISEXE
    End If
    CloseLogFile
    End
End Sub

Private Sub CallHISEXE(Optional bln�û������� As Boolean = True)
    '����HIS
    Dim strUserName As String, strPassWord As String, mError As String
    Dim strFile As String
    
    '�����ZLBH�ں����������ٻص�
    If UCase(gstrAppEXE) = UCase("zlActMain.exe") Then
        MsgBox "�Զ��������,������ִ��ģ��!", vbInformation, "�Զ�����"
        Exit Sub
    End If
    If gblnPreUpgrade Then Exit Sub
    
    If bln�û������� Then
        Call AnalyseUserNameAndPassWord(strUserName, strPassWord)
    End If
    
    'ȷ���ļ��Ƿ����
    Err = 0: On Error Resume Next
    If gstrAppEXE <> "" Then
        strFile = gstrAppPath & "\" & gstrAppEXE
    Else
        strFile = gstrAppPath & "\ZLHIS90.exe"
    End If
    If FindFile(strFile) = False Then
        strFile = gstrAppPath & "\ZLHIS+.exe"
        If FindFile(strFile) = False Then
            If gstrAppEXE <> "" Then
                strFile = gstrAppPath & "\ZLHIS90.exe"
            End If
        End If
    End If
    
    If bln�û������� Then
        mError = Shell(strFile & " " & IIf(gstrHisCommand <> "", gstrHisCommand, strUserName & "/" & strPassWord), vbNormalFocus)
    Else
        mError = Shell(strFile, vbNormalFocus)
    End If
End Sub


Private Sub Form_Load()
    Dim strTxtFile As String, mError As String
    
'    Dim strSourceFile As String, strDescFile As String
    blnOk = False
    
    Call SetWindowPos(Me.hwnd, HWND_TOP, ((Screen.Width - Me.Width) / 2) / 15, ((Screen.Height - Me.Height) / 2) / 15, 0, 0, SWP_NOSIZE)
    Me.cmdOK.Caption = "ȡ��(&C)"
    blnFirst = True
    If gblnPreUpgrade Then
        Me.Hide
    Else
        Me.Show
    End If
    lblInfor.Caption = "�����������ݿ�..."
    Me.Refresh
    DoEvents
    
    '�������ݿ�
    If OpenOracle = False Then End: Exit Sub
    lblInfor.Caption = "��ʼ����..."
    
    '��ʼ������ʽ
    Call InitUpType
    '��ʼ�ռ���ʽ
    Call iniGatherTYpe
    '��ʼ������
    Call InintVar
    
    
    
    '�ж��Ƿ�ΪUSER��Ȩ������
    If GetAdmin = False Then
        '���û�й���Ȩ��
        If GetAdministrator = False Then '��ȡ����ԱȨ��
            Unload Me
        End If
        End 'ǿ���˳�����
    End If
    
    
    If gblnPreUpgrade Then
        'Ԥ����
        Call OpenLogFile(True)
        WriteTxtLog ""
        WriteTxtLog ""
        WriteTxtLog "=============================================��ʼ����Ԥ����============================================="
        WriteTxtLog "--" & Format(Now(), "yyyy-mm-dd HH:MM:SS")
    Else
        '    ȷ���Ƿ��������
        If IsUpgrade = False Then End: Exit Sub
        
        Call OpenLogFile(False)
        WriteTxtLog ""
        WriteTxtLog ""
        WriteTxtLog "=============================================��ʼ�������ռ�============================================="
        WriteTxtLog "--" & Format(Now(), "yyyy-mm-dd HH:MM:SS")
    End If
    
    lblInfor.Caption = "���������ļ�������..."
    
    '�����Ƿ���ͨ
    If IIf(gbln�ռ� = True, gintGatherTYpe = 0, gintUpType = 0) Then
        If IsNetServer = False Then
            MsgBox "�޷����ӵ�������:" & gstrServerPath & "��,��ȷ�������Ƿ�ͨ," & vbCrLf _
            & "������������ļ�" & IIf(gbln�ռ� = True, "�ռ�", "����") & "�����Ƿ�������ȷ!", vbInformation + vbDefaultButton1, "�ͻ����Զ�" & IIf(gbln�ռ� = True, "�ռ�", "����")
            
            Call SaveClientLog("�޷����ӵ�������:" & gstrServerPath & "��,��ȷ�������Ƿ�ͨ��")
            Call UpdateCondition(2)
            End:
            Exit Sub   '���ӹ��������
        End If
    Else
        If IsFtpServer = False Then
            MsgBox "�޷����ӵ�������:" & gstrServerPath & "��,��ȷ��FTP�������Ƿ���," & vbCrLf _
            & "������������ļ�" & IIf(gbln�ռ� = True, "�ռ�", "����") & "�����Ƿ�������ȷ!", vbInformation + vbDefaultButton1, "�ͻ����Զ�" & IIf(gbln�ռ� = True, "�ռ�", "����")
            
            Call SaveClientLog("�޷����ӵ�������:" & gstrServerPath & "��,��ȷFTP�������Ƿ�ͨ��")
            Call UpdateCondition(2)
            End:
            Exit Sub   '����FTP������
        End If
    End If
    
    '���ȼ���Ƿ���MD5���ò����Ƿ���Ҫ����
    Call isMD5UpGrade
    
    
    'ȷ���Ƿ�������������
    If gBlnHisCrustCompare Then
        lblInfor.Caption = "������������..."
        If InStrRev(UCase(App.Path), UCase("\Apply"), -1) = 0 Then
            If isHisCurstUpGrade = True Then
                Err = 0: On Error Resume Next
                mError = Shell(gstrAppPath & "\Apply\zlHisCrust.exe" & " " & gcnnOracle.ConnectionString & "||1" & "||" & gstrAppEXE & "||||" & gstrHisCommand, vbNormalFocus)
                '������ǳ���
                If mError <> 0 Then
                    End
                    Exit Sub
                End If
            Else
            End If
        End If
    End If

    '���⴦��7Z���ļ�
    Call is7zUpGrade
    
    
    '��ȡ��ʱ���Ŀ¼
    gstrTempPath = GetTmpPath
    If gstrTempPath <> "" Then
        gstrTempPath = gstrTempPath & "ZLTEMP\"
    Else
        gstrTempPath = GetWinPath & "\ZLTEMP\"
    End If
    
    '��ȡԤ������ʱĿ¼
    gstrPerTempPath = GetTmpPath
    If gstrPerTempPath <> "" Then
        gstrPerTempPath = gstrPerTempPath & "ZLPERTEMP\"
    Else
        gstrPerTempPath = GetWinPath & "\ZLPERTEMP\"
    End If
    
    '�Ƿ�Ϊ��ʱ��ʽ����
    If gblnOfficialUpgrade Then
        gblnԤ����� = GetPreUpgrad(gstrComputerName)
    End If
    
    '������������
    If gbln�ռ� Then
        lblInfor.Caption = "�����ϴ��ļ�����..."
        If GetClientFiles = False Then End: Exit Sub
    Else
        lblInfor.Caption = "������������..."
        If getSeverFiles = False Then End: Exit Sub
    End If
    lngCount = 0
    Timer1.Enabled = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Form_Resize()
        Err = 0
        On Error Resume Next
        With Image1
            .Width = Me.ScaleWidth
        End With
        With cmdOK
            .Top = ScaleHeight - .Height - 50
            .Left = ScaleWidth - .Width - 100
        End With
        
        With cmdLog
            .Top = ScaleHeight - .Height - 50
            .Left = ScaleWidth - .Width - cmdOK.Width - 200
        End With
        
        If gblnPreUpgrade Then
            Me.Hide
        Else
            Me.Show
        End If
'        With prgPross
'            .Top = cmdOK.Top - prgPross.Height - 50
'            .Width = ScaleWidth - 100
'        End With
'
'        With lvwMan
'            .Width = ScaleWidth
'            .Height = prgPross.Top - .Top - 100
'        End With
'        With lblInfor
'            .Left = ScaleLeft + 50
'            .Top = cmdOK.Top + (cmdOK.Height - .Height) / 2
'            .Width = cmdOK.Left
'        End With
End Sub

Private Function FileUpgrade() As Boolean
    '����:�ļ���������
    Dim lst As ListItem
    Dim n As Integer
    Dim strSourceFile As String
    
    Dim strTargetFile As String
    Dim strSourceVer As String, strSourceDate As String
    Dim strTargetVer As String, strTargetDate As String
    Dim strErrMsg As String
    Dim blnCommon As Boolean
    Dim strUpgradeInfor As String
    Dim strLocalPath As String '���ؽ�ѹ�ļ�·��
    Dim objFile As New FileSystemObject
    Dim strTempFlag As String
    Dim lngDefeated As Long 'ʧ�ܸ���
    Dim blnע�� As Boolean
    Dim bln���� As Boolean
    Dim blnSysFile As Boolean
    Dim strSysFile As String

    
    FileUpgrade = False
    lngDefeated = 0
    prgPross.Min = 0
    prgPross.Value = 0
    prgPross.Max = IIf(Me.lvwMan.ListItems.Count = 0, 2, Me.lvwMan.ListItems.Count)
    
    strUpgradeInfor = ""
    Me.lblInfor.Caption = "����ע���ļ�"
    
'    For Each lst In Me.lvwMan.ListItems
    For n = 1 To Me.lvwMan.ListItems.Count
        Set lst = Me.lvwMan.ListItems(n)
        
        blnCommon = False
        prgPross.Value = prgPross.Value + 1
        lst.Selected = True
        
        If gbln�ռ� Then
            strSourceFile = gstrAppPath & "\" & lst.Text
            strTargetFile = gstrServerPath & "\" & GetMyCompterName & "_" & lst.Text
            
            If gintGatherTYpe = 0 Then
                '�Ƚ��ļ�
                If CompareFile(strSourceFile, strTargetFile, strSourceVer, strSourceDate, strTargetVer, strTargetDate) Then
                    GetCopyAndReg strSourceFile, strTargetFile, strErrMsg, True
                    
                    'д��ע����Ϣ
                    If strErrMsg <> "��������!" And strErrMsg <> "δװ�˲���!" And strErrMsg <> "�ͻ��˲����ڴ˲���,�����ǽ���������!" Then
                        If strUpgradeInfor = "" Then
                            strUpgradeInfor = "��" & Format(Date, "yyyy-mm-dd hh:mm:ss") & "������������, " & vbCrLf & " ���ٴ�һ��������������,��:" & lst.Text
                        End If
                        lst.SmallIcon = "Err"
                    Else
                        lst.SmallIcon = "Ok"
                    End If
                Else
                    lst.SmallIcon = "Ok"
                    strErrMsg = "�ļ���ͬ,û�б�Ҫ����!"
                End If
            Else
                strTargetFile = GetMyCompterName & "_" & lst.Text
                If FtpupFile(strSourceFile, strTargetFile) Then
                    lst.SmallIcon = "Ok"
                    strErrMsg = "�ռ����!"
                Else
                    lst.SmallIcon = "Err"
                    strErrMsg = "���ش���!"
                End If
            End If
        Else
            '0.��ȡ:��ȡ�ļ�����·��:���ҵ�񲿼��Ƿ�װ
            If gintUpType = 0 Then
                strSourceFile = gstrServerPath & "\" & lst.Text & ".7z"
                strTargetFile = GetSetupPath(lst.Text, NVL(lst.SubItems(7), ""), NVL(lst.Tag, ""), gstrAppPath, NVL(lst.SubItems(6), ""))
            Else
                strSourceFile = lst.Text & ".7z"
                strTargetFile = GetSetupPath(lst.Text, NVL(lst.SubItems(7), ""), NVL(lst.Tag, ""), gstrAppPath, NVL(lst.SubItems(6), ""))
            End If
            
            '�����:68569,ɾ��PUBLIC��Ӧ��SYSTEM32�µĹ�������
            If UCase(NVL(lst.SubItems(7), "")) = "[PUBLIC]" Then
                strSysFile = GetWinSystemPath & "\" & lst.Text
                If objFile.FileExists(strSysFile) Then
                    On Error Resume Next
                    Call objFile.DeleteFile(strSysFile)
                    Sleep 50
                    If objFile.FileExists(strSysFile) Then
                        WriteTxtLog "ɾ��PUBLIC��Ӧ��SYSTEM32�µĹ�������:" & strSysFile & "ʧ��!"
                    Else
                        WriteTxtLog "ɾ��PUBLIC��Ӧ��SYSTEM32�µĹ�������:" & strSysFile & "�ɹ�!"
                    End If
                End If
            End If
            
            strSourceVer = GetFileListValue(lst.Text, 1) '��lst.SubItems(1)
            strSourceDate = GetFileListValue(lst.Text, 2) '��lst.SubItems(3)
            
            
            'strTargetFile="" ��ʾδ��װ��ҵ�񲿼�
            If strTargetFile = "" And NVL(lst.Tag, "") = "1" Then
                strTargetVer = ""
                strTargetDate = ""
                strErrMsg = "����û�а�װ�ò���,��������"
                lst.SmallIcon = "Ok"
                GoTo zt
            End If
            
            
            '1.���:����Ƿ���Ҫ���ظ��ļ�,�Ƚ��ļ���MD5ֵ
            lblInfor.Caption = "���ڼ�鲿��:" & lst.Text
            If CompareMD5Down(strTargetFile, lst.Text) = False Then
                strTargetVer = GetCommpentVersion(strTargetFile)
                strTargetDate = Format(FileDateTime(strTargetFile), "yyyy-MM-DD hh:mm:ss")
                strErrMsg = "�ļ�MD5��ͬ,��������!"
                lst.SmallIcon = "Ok"
                GoTo zt
            Else
'����:
                '����ļ����ڻ�ȡ�ְ汾���޸�����
                If mobjFile.FileExists(strTargetFile) Then
                    strTargetVer = GetCommpentVersion(strTargetFile)
                    strTargetDate = Format(FileDateTime(strTargetFile), "yyyy-MM-DD hh:mm:ss")
                Else
                    strTargetVer = ""
                    strTargetDate = ""
                End If
                
                '2.�����ļ�
                strLocalPath = lst.Text & ".7z"
                lblInfor.Caption = "�������ز���:" & lst.Text
                If FileTempDown(strSourceFile, strLocalPath, strErrMsg) = False Then
                    If strErrMsg <> "������ɣ�" Then
                        If mobjFile.FileExists(strTargetFile) Then
                            strErrMsg = "�ļ��ڷ�����Ŀ¼������!"
                        Else
                        
                        
                            If strErrMsg = "�ļ��ڷ�����Ŀ¼������!" Then
                                lst.SmallIcon = "Err"
                                GoTo zt
                            Else
                                strErrMsg = "�ļ���������!"
                                lst.SmallIcon = "Ok"
                                GoTo zt
                            End If
                        End If
                    Else
                        strErrMsg = "�����ļ�ʧ��!"
                    End If
                    lst.SmallIcon = "Err"
                    GoTo zt
                End If
                
                
                '�����Ԥ�������������
                If gblnPreUpgrade Then
                    lst.SmallIcon = "Ok"
                    GoTo zt
                End If
                
                
                '3.��ѹ�ļ�
                lblInfor.Caption = "���ڽ�ѹ����:" & lst.Text
                If FileDeCompression(strLocalPath, strErrMsg) = False Then
                    strErrMsg = "��ѹ���ļ�ʧ��!"
                    lst.SmallIcon = "Err"
                    GoTo zt
                End If
                
                '4.������ע���ļ�
'                strTargetFile = "C:\Temp\" & lst.Text
                lblInfor.Caption = "����ע�Ჿ��:" & lst.Text
                If lst.SubItems(9) = "" Or lst.SubItems(9) = "0" Then
                    blnע�� = False
                Else
                    blnע�� = True
                End If

                If lst.SubItems(10) = "" Or lst.SubItems(10) = "0" Then
                    bln���� = False
                Else
                    bln���� = True
                End If
                
                If NVL(lst.Tag, "") = "5" Then
                    blnSysFile = True
                Else
                    blnSysFile = False
                End If
                
                If GetCopyAndReg(strLocalPath, strTargetFile, strErrMsg, blnע��, blnSysFile, bln����) = False Then
                    strErrMsg = "�滻�ļ�ʧ�ܿ����ѱ����������ռ!"
                    lst.SmallIcon = "Err"
                    On Error Resume Next
                    Call Kill(strLocalPath)
                    GoTo zt
                Else
                    On Error Resume Next
                    lst.SmallIcon = "Ok"
                    Call Kill(strLocalPath)
                End If
                
                
                '5.���MD5ֵ�Ƿ���ȷ
                If strErrMsg = "���Ա���������" Then GoTo zt
                If strErrMsg = "�������ռ" Then GoTo zt
                If CheckSysFile(strTargetFile) Then GoTo zt
                If blnSysFile = True And bln���� = False Then GoTo zt
                lblInfor.Caption = "���ڱȽϲ���:" & lst.Text
                If CompareMD5Down(strTargetFile, lst.Text, strErrMsg) Then
                   If strErrMsg = "������û�и��ļ�MD5��Ϣ!" Then
                     strErrMsg = "������û�и��ļ�MD5��Ϣ!"
                   Else
                     strErrMsg = "�ļ���������,MD5��ֵ��ԭֵ��һ��!"
                   End If
                   lst.SmallIcon = "Err"
                   GoTo zt
                End If
                lst.SmallIcon = "Ok"
            End If
        End If
zt:
        If lst.SmallIcon = "Err" Then
            If strUpgradeInfor = "" Then
                strUpgradeInfor = "��" & Format(Date, "yyyy-mm-dd hh:mm:ss") & "������������, " & vbCrLf & " ���ٴ�һ��������������,��:" & lst.Text
            Else
                strUpgradeInfor = strUpgradeInfor & "," & lst.Text
            End If
            strTempFlag = "[ʧ��]:"
            lngDefeated = lngDefeated + 1
        Else
            strTempFlag = "[�ɹ�]:"
        End If
        
        strErrMsg = IIf(gbln�ռ�, Replace(strErrMsg, "����", "�ռ�"), strErrMsg)
        lst.SubItems(3) = strTargetVer
        lst.SubItems(5) = strTargetDate
        lst.SubItems(1) = strErrMsg
        lst.EnsureVisible
        WriteTxtLog strTempFlag & strSourceFile & "(�汾:" & strSourceVer & "   �޸�����:" & strSourceDate & ")    ====>    " & vbCrLf & _
                        strTargetFile & "(�汾:" & strTargetVer & "   �޸�����:" & strTargetDate & ")        ������Ϣ:" & strErrMsg & vbCrLf
        
        DoEvents
    Next
    
    'ִ���������ļ�
    strAutoRun = gstrAppPath & "\zlAutoRun.ini"
    strAutoRunBat = gstrAppPath & "\zlAutoRun.bat"
    If mobjFile.FileExists(strAutoRun) Or mobjFile.FileExists(strAutoRunBat) Then
        Dim ret As Long
        Name strAutoRun As gstrAppPath & "\zlAutoRun.bat"
        On Error Resume Next
        Call Kill(strAutoRun)
        
        ret = ShellExecute(0&, "open", gstrAppPath & "\zlAutoRun.bat", "", gstrAppPath & "\zlAutoRun.bat", 5) 'SW_SHOW
        If ret = 31 Then
            strErrMsg = "������ִ��ʧ��!"
'           MsgBox "û���ҵ��ʵ��ĳ���������,�밲װ��Ч�ĳ���!", vbInformation, "��ʾ"
            WriteTxtLog "�������ļ�ִ��ʧ��!"
        Else
            WriteTxtLog "�������ļ�ִ�гɹ�!"
        End If
        
        blnAutoRun = True
    Else
        blnAutoRun = False
    End If
    
    '���7z.exe����ϵͳ����
    Call fun_KillProcess(PROAPPCTION)
    
    If InStr(1, UCase(App.Path), UCase("APPLY")) <> 0 Then  '�����ϼ�Ŀ¼
        GetCopyAndReg App.Path & "\zlHisCrust.exe", Replace(App.Path, "\Apply", "") & "\zlHisCrust.exe", strErrMsg
    End If
    
   '���ռ����������ļ�
    strSourceFile = gstrAppPath & "\ZLUpGradeList.Lst"
    strTargetFile = gstrServerPath & "\" & GetMyCompterName & "_ZLUpGradeList.LOG"
    If InStr(1, gstr�ռ�����, "LOG") <> 0 And gbln�ռ� Then
        '�ռ�������־
        If objFile.FileExists(strSourceFile) Then
            GetCopyAndReg strSourceFile, strTargetFile, strErrMsg
        End If
    End If
    
    If lngDefeated = 0 Then
        If gblnPreUpgrade Then
            WriteTxtLog "����Ԥ�������!"
            Call SaveClientLog("����Ԥ�������")
            Call UpdateCondition(1)
        Else
            WriteTxtLog "�����������!"
            Call SaveClientLog("�����������")
            Call UpdateCondition(1)
        End If
        Me.lblInfor.Caption = IIf(gbln�ռ�, "�ռ�", "����") & "�ɹ�"
        cmdLog.Visible = False
        gblnOk = True
    Else
        '��¼������־
        If GetErrParameter(3) = "1" Then
            Dim i As Long
            With lvwMan
            For i = 1 To .ListItems.Count
                If .ListItems(i).SmallIcon <> "Ok" Then
                    Call SaveErrLog(.ListItems(i).Text & "-" & .ListItems(i).SubItems(1))
                    Call SaveClientLog(.ListItems(i).Text & "-" & .ListItems(i).SubItems(1))
                End If
            Next
            End With
        End If
        

        '������ʾ�����б�
        With lvwMan
            For n = 1 To .ListItems.Count
                If n > .ListItems.Count Then
                    Exit For
                End If
                If .ListItems(n).SmallIcon = "Ok" Then
                    .ListItems.Remove n
                    n = n - 1
                End If
            Next
        End With
        
        Me.Height = 5445
        lblInfor.Caption = "�ͻ��˲����������"
        WriteTxtLog "�ܹ�:" & lngDefeated & "�ļ�����ʧ��!"
        Call SaveClientLog("�ܹ�:" & lngDefeated & "�ļ�����ʧ��!")
        Me.lblInfor.Caption = "��" & lngDefeated & "�ļ�����ʧ��,��˲�!"
        cmdLog.Visible = True
        gblnOk = False
    End If
    
    Dim strSQL As String
    Err = 0
    
    '��������˵��,�����������ռ����
     If gblnPreUpgrade = False Then
         If strUpgradeInfor <> "" Then
             If LenB(StrConv(strUpgradeInfor, vbFromUnicode)) > 200 Then
                 strUpgradeInfor = Mid(strUpgradeInfor, 1, 200)
             End If
             'strSQL = "Update zltools.zlclients set ˵��='" & strUpgradeInfor & "' where upper(����վ)='" & UCase(gstrComputerName) & "'"
             strSQL = "Zl_Zlclients_Control(9,'" & gstrComputerName & "',Null,Null,Null,Null,Null,Null,Null,Null,'" & strUpgradeInfor & "')"
                  
         Else
             strUpgradeInfor = "��" & Format(Now, "yyyy-mm-dd HH:mm:ss") & "�����˲���"
             If gbln�ռ� Then
                 'strSQL = "Update zltools.zlclients set ˵��='" & strUpgradeInfor & "' ,�ռ���־=0 where upper(trim(����վ))='" & UCase(gstrComputerName) & "'"
                strSQL = "Zl_Zlclients_Control(10,'" & gstrComputerName & "',Null,Null,Null,Null,Null,Null,Null,Null,'" & strUpgradeInfor & "')"
             Else
                ' strSQL = "Update zltools.zlclients set ˵��='" & strUpgradeInfor & "' ,������־=0 where upper(trim(����վ))='" & UCase(gstrComputerName) & "'"
                strSQL = "Zl_Zlclients_Control(11,'" & gstrComputerName & "',Null,Null,Null,Null,Null,Null,Null,Null,'" & strUpgradeInfor & "')"
                 '������Ԥ����Ŀ¼����,�ͽ���ɾ���ļ�Ŀ¼
                 If gblnOfficialUpgrade And gblnԤ����� Then
                    If mobjFile.FolderExists(gstrPerTempPath) Then
                       On Error Resume Next
                       Call mobjFile.DeleteFolder(Left(gstrPerTempPath, Len(gstrPerTempPath) - 1))
                    End If
                 End If
             End If
        End If
    Else
        '����վ���Ԥ�������״̬
        If strUpgradeInfor <> "" Then
            If LenB(StrConv(strUpgradeInfor, vbFromUnicode)) > 200 Then
                 strUpgradeInfor = Mid(strUpgradeInfor, 1, 200)
            End If
            strSQL = "Zl_Zlclients_Control(12,'" & gstrComputerName & "',Null,Null,Null,Null,Null,Null,Null,Null,'" & "Ԥ��������:" & strUpgradeInfor & "')"
           ' strSQL = "Update zltools.zlclients set Ԥ�����=0,˵��='" & "Ԥ��������:" & strUpgradeInfor & "' where upper(trim(����վ))='" & UCase(gstrComputerName) & "'"
        Else
            strSQL = "Zl_Zlclients_Control(13,'" & gstrComputerName & "')"
           ' strSQL = "Update zltools.zlclients set Ԥ�����=1 where upper(trim(����վ))='" & UCase(gstrComputerName) & "'"
        End If
    End If
   
    gcnnOracle.Execute strSQL
    
    '���Ϊ��ʱ����
    If gblnOfficialUpgrade Then
        'strSQL = "Update zltools.zlclients set Ԥ��ʱ��=Null ,Ԥ�����=Null where upper(trim(����վ))='" & UCase(gstrComputerName) & "'"
        strSQL = "Zl_Zlclients_Control(14,'" & gstrComputerName & "')"
        gcnnOracle.Execute strSQL
    End If
    
    Me.cmdOK.Caption = "���(&O)"
    Me.cmdOK.Visible = True
    blnOk = True
    FileUpgrade = True

    '�Ͽ���������
    
    If IIf(gbln�ռ� = True, gintGatherTYpe = 0, gintUpType = 0) Then
        '�ر�Share����
        CancelNetServer
    Else
        '�ر�FTP����
        CancelFtpServer
    End If
'    End
End Function

Private Function FindHisBrow() As Boolean
    '����:���Ҳ�����HIS�����ڵ���ؽ���
    '�ɹ�:�����ɹ�,����true,���򷵻�false
    Dim lngHwnd As Long
    Dim lngZlhisHwnd As Long
    Dim lngVBHwnd As Long
    Dim lngPid As Long
    Dim lngProcess As Long
    Err = 0: On Error GoTo ErrHand:
    
    '���Ԥ����,���˳���
    If gblnPreUpgrade Then
        Exit Function
    End If
    
    Do While True
         lngHwnd = FindWindow(vbNullString, "����̨")
         If lngHwnd = 0 Then
            lngHwnd = FindWindow(vbNullString, "ҽԺ��Ϣϵͳ")
            If lngHwnd = 0 Then
                Exit Do
            End If
         End If
         If lngHwnd <> 0 Then
            '�����Ƿ���VB�ڵ��õ���̨���ǳ���ֱ��ִ�е���̨
            lngZlhisHwnd = fun_ExitsProcess("zlhis+.exe")
            If lngZlhisHwnd <> 0 Then
                Call TerminateProcess(lngZlhisHwnd, 1&)
            Else
                lngVBHwnd = fun_ExitsProcess("vb6.exe")
                If lngVBHwnd <> 0 Then
                    If MsgBox("���������⵽VB6�����˿��ܻ������Ĳ���." & vbCrLf & "Ϊ�˱�֤ϵͳ��������,�Ƿ�ر�VB6����!", vbQuestion + vbYesNo, "�ͻ����Զ�����") = vbYes Then
                        Call GetWindowThreadProcessId(lngHwnd, lngPid)
                        lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
                        Call TerminateProcess(lngProcess, 1&)
                    Else
                        GoTo NoClose
                    End If
                Else
                    Call GetWindowThreadProcessId(lngHwnd, lngPid)
                    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
                    Call TerminateProcess(lngProcess, 1&)
                End If
            End If
         End If
     Loop
     
     
    Do While True
        '�Զ��ر�zlSvrStudio����
        lngZlhisHwnd = fun_ExitsProcess("zlSvrStudio.exe")
        If lngZlhisHwnd <> 0 Then
            Call TerminateProcess(lngZlhisHwnd, 1&)
        End If
        
        If lngZlhisHwnd = 0 Then
            Exit Do
        End If
    Loop
    
    Do While True
        '�Զ��ر�Zl9LISComm����
        lngZlhisHwnd = fun_ExitsProcess("Zl9LISComm.exe")
        If lngZlhisHwnd <> 0 Then
            Call TerminateProcess(lngZlhisHwnd, 1&)
        End If
        
        If lngZlhisHwnd = 0 Then
            Exit Do
        End If
    Loop
    
    Do While True
        '�Զ��ر�zlLisReceiveSend����
        lngZlhisHwnd = fun_ExitsProcess("zlLisReceiveSend.exe")
        If lngZlhisHwnd <> 0 Then
            Call TerminateProcess(lngZlhisHwnd, 1&)
        End If
        
        If lngZlhisHwnd = 0 Then
            Exit Do
        End If
    Loop
    
    Do While True
        '�Զ��ر�ZlPacsSrv����
        lngZlhisHwnd = fun_ExitsProcess("ZlPacsSrv.exe")
        If lngZlhisHwnd <> 0 Then
            Call TerminateProcess(lngZlhisHwnd, 1&)
        End If
        
        If lngZlhisHwnd = 0 Then
            Exit Do
        End If
    Loop
    
NoClose:
    FindHisBrow = False
    Exit Function
ErrHand:
End Function

'�жϴ����Ƿ����Ҫ��
Function TaskWindow(hwcurr As Long) As Long
    Dim lngStyle As Long, IsTask As Long
    '��ȡ���ڷ�񣬲��ж��Ƿ����Ҫ��
    lngStyle = GetWindowLong(hwcurr, GWL_STYLE)
    If (lngStyle And IsTask) = IsTask Then
     TaskWindow = True
    End If
End Function

Public Sub CloseWindow(app_name As String)
    Dim app_hwnd As Long
    app_hwnd = FindWindow(vbNullString, app_name)
    SendMessage app_hwnd, WM_CLOSE, 0, 0
End Sub

Private Sub lvwMan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If blnOk = False Then Exit Sub
    Err = 0
    On Error Resume Next
    If mintColumn = ColumnHeader.Index - 1 Then
        lvwMan.SortOrder = IIf(lvwMan.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMan.SortKey = mintColumn
        lvwMan.SortOrder = lvwAscending
    End If

End Sub

Private Sub Timer1_Timer()
        Dim strSQL As String
        Timer1.Enabled = False
        If FindHisBrow = False Then
            MousePointer = 11
            If FileUpgrade = False Then
                Me.cmdOK.Caption = "ȡ��(&C)"
                Timer1.Enabled = False
                MousePointer = 0
                Exit Sub
            Else
                If gbln�ռ� = False Then
                    'ȷ���Ƿ��ռ�����
                    If Is�ռ��ļ� Then
                        '�����ռ��ļ�
                        Call InintVar
                        lblInfor.Caption = "�������������ϴ�������..."

                        '�����Ƿ���ͨ
                        If IIf(gbln�ռ� = True, gintGatherTYpe = 0, gintUpType = 0) Then
                            If IsNetServer = False Then
                                Timer1.Enabled = False
                                MousePointer = 0
                                MsgBox "�޷����ӵ�������:" & gstrServerPath & "��,��ȷ�������Ƿ�ͨ," & vbCrLf _
                                & "������������ļ�" & IIf(gbln�ռ� = True, "�ռ�", "����") & "�����Ƿ�������ȷ!", vbInformation + vbDefaultButton1, "�ͻ����Զ�" & IIf(gbln�ռ� = True, "�ռ�", "����")
                                               
                                Call SaveClientLog("�޷����ӵ�������:" & gstrServerPath & "��,��ȷ����������Ƿ�ͨ��")
                                Call UpdateCondition(2)
                                Exit Sub   '���ӹ��������
                            End If
                        Else
                            If IsFtpServer = False Then
                                Timer1.Enabled = False
                                MousePointer = 0
                                MsgBox "�޷����ӵ�������:" & gstrServerPath & "��,��ȷ��FTP�������Ƿ���," & vbCrLf _
                                & "������������ļ�" & IIf(gbln�ռ� = True, "�ռ�", "����") & "�����Ƿ�������ȷ!", vbInformation + vbDefaultButton1, "�ͻ����Զ�" & IIf(gbln�ռ� = True, "�ռ�", "����")
                                
                                Call SaveClientLog("�޷����ӵ�������:" & gstrServerPath & "��,��ȷFTP�������Ƿ�ͨ��")
                                Call UpdateCondition(2)
                                Exit Sub   '����FTP������
                            End If
                        End If

                        lblInfor.Caption = "�����ϴ��ļ�����..."
                        If GetClientFiles = True Then
                            If FileUpgrade = False Then
                                Me.cmdOK.Caption = "ȡ��(&C)"
                                Timer1.Enabled = False
                                MousePointer = 0
                                Exit Sub
                            End If
                        Else
                            lblInfor.Caption = "û�б��ռ����ļ�..."
                            Me.cmdOK.Caption = "���(&C)"
                        End If
                    End If
                End If
            End If
            MousePointer = 0
            Timer1.Enabled = False
            
            '����������˳�
            If gblnOk Then '�޴����˳�
                Call cmdOK_Click
            End If
        Else
            Timer1.Enabled = True
        End If
End Sub
Private Function Is�ռ��ļ�() As Boolean
    '   ���ܣ�ȷ���Ƿ��ռ��ļ�
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo ErrHand:
    strSQL = "Select * From zlClients Where  �ռ���־=1 and upper(����վ)='" & gstrComputerName & "'"
    With rsTemp
        .Open strSQL, gcnnOracle
        gbln�ռ� = .RecordCount <> 0
        Is�ռ��ļ� = gbln�ռ�
        .Close
    End With
    Exit Function
ErrHand:
    Is�ռ��ļ� = False
End Function

Private Function GetClientFiles() As Boolean
     '�ռ��ļ�
     Dim i As Long, lngfile As Long, j As Long
     Dim strFileName  As String, strFile As String, str�汾�� As String
     Dim strArr, strArr1
     Dim lst As ListItem
     Dim objFile As New FileSystemObject
     
     strArr = Split(gstr�ռ�����, ";")
     
     strFileName = ""
     
     For i = 0 To UBound(strArr)
        strArr1 = Split(strArr(i), ",")
        For j = 0 To UBound(strArr1)
            If InStr(1, strArr1(j), ".") <> 0 Then
                '�������С����,���ʾ�ض��ļ�
                strFileName = strFileName & ";" & strArr(i)
            Else
                strFileName = strFileName & ";*." & strArr(i)
            End If
        Next
     Next
     If strFileName <> "" Then
        strFileName = Mid(strFileName, 2)
     Else
        strFileName = "*.DLSDKSKS"
     End If
     Err = 0
     On Error GoTo ErrHand:
     i = 1
   
    With File1
        .Path = gstrAppPath
        .FileName = strFileName
        lvwMan.ListItems.Clear
        For lngfile = 0 To .ListCount - 1
            '����SQLNET.Log�ļ�
            If InStr(1, UCase(.List(lngfile)), "SQLNET.LOG") = 0 Then
                strFile = gstrAppPath & "\" & .List(lngfile)
                str�汾�� = GetCommpentVersion(strFile)
                Set lst = lvwMan.ListItems.Add(, "K" & i, .List(lngfile), "List", "List")
                    lst.Tag = objFile.GetExtensionName(.List(lngfile))
                    lst.SubItems(2) = str�汾��
                    lst.SubItems(4) = Format(FileDateTime(strFile), "yyyy-MM-DD hh:mm:ss") 'Format(FileDateTime(strFile), "yyyy-MM-dd hh:mm:ss")
            End If
            i = i + 1
        Next
    End With
    GetClientFiles = True
    Exit Function
ErrHand:
    GetClientFiles = False
End Function

Private Function getSeverFiles() As Boolean
    '----------------------------------------------------------------------------------------
    '����:��ȡ�����������µ���������
    '����:
    '����:��д�ɹ�,����true,���򷵻�False
    '----------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lst As ListItem
    Dim rsTmp As New ADODB.Recordset
    getSeverFiles = False
    
    strSQL = "Select ���, �ļ���, �汾��, �޸�����,�ļ�����,ҵ�񲿼�,��װ·��,MD5,�Զ�ע��,ǿ�Ƹ��� From zlfilesupgrade where upper(�ļ���) not in('ZLHISCRUST.EXE' , '7Z.EXE','7Z.DLL') and MD5 is not null order by ���"

    lvwMan.ListItems.Clear
    With rsTmp
        .CursorLocation = adUseClient
        .Open strSQL, gcnnOracle
        If .RecordCount = 0 Then
            getSeverFiles = True
            Exit Function
        End If
        Err = 0
        On Error GoTo ErrHand:
        Do While Not .EOF
            If UCase(NVL(!�ļ���)) = "ZLHISCRUST.EXE" Or UCase(NVL(!�ļ���)) = "7Z.EXE" Or UCase(NVL(!�ļ���)) = "7Z.DLL" Or UCase(NVL(!�ļ���)) = "ZLRUNAS.EXE" Then
                GoTo NotAddFile
            End If
            Set lst = lvwMan.ListItems.Add(, "K" & .AbsolutePosition, !�ļ���, "List", "List")
            lst.Tag = IIf(IsNull(!�ļ�����), 0, !�ļ�����)
            lst.SubItems(2) = IIf(IsNull(!�汾��), "", GetVersion(IIf(IsNull(!�汾��), 0, !�汾��)))
            lst.SubItems(4) = Format(!�޸�����, "yyyy-MM-DD HH:mm:ss")
            lst.SubItems(6) = NVL(!ҵ�񲿼�, "")
            lst.SubItems(7) = NVL(!��װ·��, "")
            lst.SubItems(8) = NVL(!MD5, "")
            lst.SubItems(9) = NVL(!�Զ�ע��, "")
            lst.SubItems(10) = NVL(!ǿ�Ƹ���)
NotAddFile:
            .MoveNext
        Loop
    End With
    
    getSeverFiles = True
    Exit Function
ErrHand:
    getSeverFiles = False
End Function

Private Function GetPreUpgrad(ByVal strComputerName As String) As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    strSQL = "Select * From zlClients Where  Ԥ�����=1 and upper(����վ)='" & strComputerName & "'"
    
     With rsTmp
        .CursorLocation = adUseClient
        .Open strSQL, gcnnOracle
        If .RecordCount = 1 Then
            GetPreUpgrad = True
            Exit Function
        Else
            GetPreUpgrad = False
            Exit Function
        End If
    End With
    Exit Function
ErrHand:
    GetPreUpgrad = False
End Function

Private Function CompareMD5Down(ByVal strLocalFile As String, ByVal strName As String, Optional strMsg As String) As Boolean
    '����:�Ƚ��ļ���MD5�Ƿ���ͬ
    '���:strLocalFile ��Ҫ�Ƚ��ļ�����·��,strName�ļ�����,intOption ������ֵ��2�ַ��ط���
    '����:True ��Ҫ���� Flase ����Ҫ����
    '����:ף��
    '����:2010/12/15
    On Error GoTo errH
    Dim objFile As New FileSystemObject
    Dim strFileMD5 As String
    Dim strListFileMD5 As String
    If objFile.FileExists(strLocalFile) Then
        strFileMD5 = HashFile(strLocalFile, 2 ^ 27)
        strListFileMD5 = GetFileListValue(strName, 0)
        If strListFileMD5 = "" Then
            strMsg = "������û�и��ļ�MD5��Ϣ!"
        End If
        
        If strFileMD5 = strListFileMD5 Then
            CompareMD5Down = False
        Else
            CompareMD5Down = True
        End If
    Else
        CompareMD5Down = True
    End If
    Exit Function
errH:
    If Err Then
        CompareMD5Down = True
    End If
End Function

Private Function GetFileListValue(ByVal strFileName As String, ByVal intOption As Integer) As String
    '���ܴӷ������б��ȡ�ļ�����Ϣ
    '��� strFileName ��Ҫ��ȡ��Ϣ���ļ���
    '0:��ȡMD5ֵ
    '1:��ȡ�汾��
    '2:��ȡ�޸�����
    On Error GoTo errH
    Dim i As Integer
    Dim lngCurFileIndex As Long
    lngCurFileIndex = -1
    With lvwMan
        For i = 1 To .ListItems.Count
            If UCase(.ListItems(i).Text) = UCase(strFileName) Then
                lngCurFileIndex = i
                Exit For
            End If
        Next
        
        If lngCurFileIndex >= 0 Then
            Select Case intOption
            Case 0
                GetFileListValue = .ListItems(i).SubItems(8) '�ļ�MD5ֵ
            Case 1
                GetFileListValue = IIf(.ListItems(i).SubItems(2) = "0", "", .ListItems(i).SubItems(2)) '�汾��
            Case 2
                GetFileListValue = .ListItems(i).SubItems(4) '�޸�����
            Case 3
                GetFileListValue = NVL(.ListItems(i).SubItems(10), 0)
            End Select
        Else
            GetFileListValue = ""
        End If
    End With
    Exit Function
errH:
    If Err Then
        GetFileListValue = ""
    End If
End Function

'����USERȨ�޵�Ȩ������
Private Function GetAdministrator() As Boolean
     Dim strAppRunas As String
     Dim strUser As String
     Dim strPass As String
     Dim strMsg As String
     Dim strSQL As String
     Dim rsTmpUser As New ADODB.Recordset
     Dim rsTmpPass As New ADODB.Recordset
 
        '1.��鱾���Ƿ���zlRunas�ļ�.
        strAppRunas = App.Path & "\zlRunas.exe"
        If mobjFile.FileExists(strAppRunas) = False Then
            '�����ھ��ط�����������
            
            
            '�����Ƿ���ͨ
            If IIf(gbln�ռ� = True, gintGatherTYpe = 0, gintUpType = 0) Then
                If IsNetServer = False Then
                    MsgBox "�޷����ӵ�������:" & gstrServerPath & "��,��ȷ�������Ƿ�ͨ," & vbCrLf _
                    & "������������ļ�" & IIf(gbln�ռ� = True, "�ռ�", "����") & "�����Ƿ�������ȷ!", vbInformation + vbDefaultButton1, "�ͻ����Զ�" & IIf(gbln�ռ� = True, "�ռ�", "����")
                    
                    Call SaveClientLog("�޷����ӵ�������:" & gstrServerPath & "��,��ȷ����������Ƿ�ͨ��")
                    Call UpdateCondition(2)
                    End:
                    Exit Function '���ӹ��������
                End If
            Else
                If IsFtpServer = False Then
                    MsgBox "�޷����ӵ�������:" & gstrServerPath & "��,��ȷ��FTP�������Ƿ���," & vbCrLf _
                    & "������������ļ�" & IIf(gbln�ռ� = True, "�ռ�", "����") & "�����Ƿ�������ȷ!", vbInformation + vbDefaultButton1, "�ͻ����Զ�" & IIf(gbln�ռ� = True, "�ռ�", "����")
                    
                    Call SaveClientLog("�޷����ӵ�������:" & gstrServerPath & "��,��ȷFTP�������Ƿ�ͨ��")
                    Call UpdateCondition(2)
                    End:
                    Exit Function '����FTP������
                End If
            End If
            
            Call isRunasUpGrade '����Runas�ļ�
            
            Sleep 100
        End If
        
        
        
         '2.��ȡ���ͻ��˹���Ա�û�������
        '�ж��Ƿ����سɹ�,�ɹ��ż���ִ��
        If mobjFile.FileExists(strAppRunas) = False Then
            MsgBox "δ����������ZLRUNAS,USERȨ��ִ�й���" & vbNewLine & "���������Ŀ¼���Ƿ���ڸ��ļ�.", vbInformation + vbDefaultButton1, "�ͻ����Զ�����"
            
            
            Call SaveClientLog("δ����������ZLRUNAS,USERȨ��ִ�й���,���������Ŀ¼���Ƿ���ڸ��ļ�.")
            GetAdministrator = False
            Exit Function
        End If
          

        strSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ like '����Ա�˺�'"
        With rsTmpUser
            .Open strSQL, gcnnOracle
            If rsTmpUser.RecordCount = 1 Then
                strUser = Trim(NVL(rsTmpUser!����))
                
                
                strSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ like '����Ա����'"
                rsTmpPass.Open strSQL, gcnnOracle
                If rsTmpPass.RecordCount = 1 Then
                    strPass = decipher(Trim(NVL(rsTmpPass!����)))
                Else
                    strPass = ""
                End If
                
            Else
                strUser = "Administrator"
                strPass = ""
            End If
        End With
        
        '3.ִ��zlRunas , ʹ�ù���ԱȨ�޵�¼
        strMsg = RunasShell(strUser, strPass)
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation + vbDefaultButton1, "�ͻ����Զ�����"
            
            Call SaveClientLog(strMsg & "[ZLRUNAS]")
            
            GetAdministrator = False
            Exit Function
        End If
        
        '4.�˳�����
        Unload Me
End Function

Private Function RunasShell(ByVal strUser As String, ByVal strPass As String) As String
    On Error GoTo errH
    Dim strRunas As String
    Dim strApp As String
    Dim strMsg As String
    Dim strShellTxt  As String
    strRunas = App.Path & "\ZLRUNAS.EXE"
    strApp = App.Path & "\ZLHISCRUST.EXE"
    '·���в��������ģ�����ִ�в��ɹ�
    strShellTxt = strRunas & " -u " & strUser & " -p " & strPass & " -ex """ & strApp & """" & " -lwp"
    strMsg = GetCmdTxt(strShellTxt)
    
    If InStr(strMsg, (1326)) > 0 Then
        RunasShell = "��¼ʧ��: δ֪���û�����������롣"
        Exit Function
    End If
    
    If InStr(strMsg, (1058)) > 0 Then
        RunasShell = "�޷���������ԭ�������SecLogon���񱻽��á�"
        Exit Function
    End If
    
    If InStr(strMsg, (1717)) > 0 Then
        RunasShell = "'·���в��������ģ�����ִ�в��ɹ�"
        Exit Function
    End If
    
    RunasShell = ""
    Exit Function
errH:
    If Err Then
        RunasShell = ""
    End If
End Function
