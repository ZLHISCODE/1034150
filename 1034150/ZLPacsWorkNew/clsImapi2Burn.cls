VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImapi2Burn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


  Private Type TGUID
    D1 As Long
    D2 As Long
    D3 As Long
    D4(7) As Byte
  End Type
  
  Private Type tagSTAT
    pwcsName As String
    type As Long
    cbSize As Double
    mtime As Date
    ctime As Date
    atime As Date
    grfMode As Long
    grfLocksSupported As Long
    clsid As TGUID
    grfStateBits As Long
    reserved As Long
  End Type
  

Private Const C_STR_BURN_CLIENTNAME As String = "ZLSoftBurn"
Private Const C_STR_DISC_DIR_SPLIT  As String = ">"


Private WithEvents mobjDiscs       As MsftDiscMaster2
Attribute mobjDiscs.VB_VarHelpID = -1
Private WithEvents mobjFileImage   As MsftFileSystemImage
Attribute mobjFileImage.VB_VarHelpID = -1
Private WithEvents mobjDiscFormat  As MsftDiscFormat2Data
Attribute mobjDiscFormat.VB_VarHelpID = -1

Private mobjCurRecord   As MsftDiscRecorder2
Private mobjDiscEarse   As MsftDiscFormat2Erase '���̲�������

Private mstrBurnDeviceName  As String   '��ǰ��Ҫʹ�õĿ�¼�豸����
Private mstrBurnVolumeName As String    '�̷�����
Private mlngWriteSpeed As Long          '��¼�ٶ�
Private mblnIsOverWrite As Boolean      '�Ƿ񸲸�
Private mlngVerificationLevel As TIntergrityVerificationLevel   '�����Լ�⼶��
Private mblnOnceMedia As Boolean        '�Ƿ��չ��̣��������ٴ�д��
Private mblnIncludeBaseDir As Boolean   '�Ƿ��������Ŀ¼
Private mblnIsCancelWrite As Boolean    '�Ƿ�����¼����
Private mlngReserveSize As Long         '����Ԥ����С����λK��

Private mobjBurnDirResources As Collection  '���ִ���¼��Ŀ¼
Private mobjBurnFileResources As Collection '�������¼���ļ�

Private mstrBurnPath As String

Public Event OnBurnEvent(ByVal strCurState As String, ByRef args As clsImapi2BurnArgs)

Public Event OnReadyFileEvent(ByVal strFile As String, ByVal strSourcePath As String, _
    ByVal copiedSectors As Long, ByVal totalSectors As Long)

Public Event OnBurnProcedureEvent(ByVal strState As String)





'***********************************************************************************************************
'
'��װʹ��imapi2���п�¼�Ĺ���
'
'***********************************************************************************************************




'�жϸ�ϵͳ���Ƿ���ڿ��õĿ�¼�豸
Property Get HasBurnDeviceInSystem() As Boolean
    HasBurnDeviceInSystem = mobjDiscs.IsSupportedEnvironment
End Property


'��ȡ��������
Property Get DeviceCount() As Long
    DeviceCount = mobjDiscs.Count
End Property


'��ȡ��¼�豸������
Private Function GetBurnDeviceName(objRecord As MsftDiscRecorder2) As String
    GetBurnDeviceName = TransformArrayToOneDimension(objRecord.VolumePathNames)(1) & objRecord.ProductId
End Function


'�����豸����(����+�豸��)
Property Get DeviceName(ByVal lngDeviceIndex As Long) As String
On Error Resume Next
    Dim i As Long
    Dim objRecord As New MsftDiscRecorder2
       
    Call objRecord.InitializeDiscRecorder(mobjDiscs.Item(lngDeviceIndex))
    
    DeviceName = GetBurnDeviceName(objRecord)
    
    Set objRecord = Nothing

End Property



'��ǰ��ʹ�õĿ�¼�豸����
Property Get CurBurnDevice() As String
    CurBurnDevice = mstrBurnDeviceName
End Property


Property Let CurBurnDevice(ByVal value As String)
    mstrBurnDeviceName = value

    Set mobjCurRecord = Nothing
    Set mobjCurRecord = GetDiscDevice(value)
    
    If mobjCurRecord Is Nothing Then Exit Property
    
    If Not mobjDiscFormat.IsRecorderSupported(mobjCurRecord) Then
        '��֧�ֿ�¼����
        Set mobjCurRecord = Nothing
        Exit Property
    End If
    
    If Not mobjDiscFormat.IsCurrentMediaSupported(mobjCurRecord) Then
        '�жϹ����Ƿ�׼������
        Set mobjCurRecord = Nothing
        Exit Property
    End If
    
    mobjDiscFormat.Recorder = mobjCurRecord
'    mobjDiscFormat.ClientName = C_STR_BURN_CLIENTNAME

End Property


'��¼ʹ�õľ�������
Property Get BurnVolumeName() As String
    BurnVolumeName = mstrBurnVolumeName
End Property

Property Let BurnVolumeName(value As String)
    If Trim(value) = "" Then Exit Property
    
    mstrBurnVolumeName = value
End Property


'����д���ٶ�
Property Get WriteSpeed() As Long
    WriteSpeed = mlngWriteSpeed
End Property


Property Let WriteSpeed(value As Long)
    mlngWriteSpeed = value
End Property


'�Ƿ��Դ�ļ����и���
Property Get IsOverWrite() As Boolean
    IsOverWrite = mblnIsOverWrite
End Property

Property Let IsOverWirte(value As Boolean)
    mblnIsOverWrite = value
End Property


'�Ƿ����������¼Ŀ¼
Property Get IsIncludeBaseDir() As Boolean
    IsIncludeBaseDir = mblnIncludeBaseDir
End Property

Property Let IsIncludeBaseDir(value As Boolean)
    mblnIncludeBaseDir = value
End Property


'�����Լ�⼶��
Property Get VerificationLevel() As TIntergrityVerificationLevel
    VerificationLevel = mlngVerificationLevel
End Property

Property Let VerificationLevel(value As TIntergrityVerificationLevel)
    mlngVerificationLevel = value
End Property


'������С(��λΪK)
Property Get ReserveKBSize() As Long
    ReserveKBSize = mlngReserveSize
End Property


Property Let ReserveKBSize(value As Long)
    mlngReserveSize = value
End Property


'�Ƿ�Ϊһ����ý��
Property Get OnceMedia() As Boolean
    OnceMedia = mblnOnceMedia
End Property

Property Let OnceMedia(value As Boolean)
    mblnOnceMedia = value
End Property



'**************************************************************************************************************************









'��ȡ��ǰ��¼�豸
Public Function GetCurDiscDevice() As MsftDiscRecorder2
    Set GetCurDiscDevice = mobjCurRecord
End Function


'�����豸��ȡ��ָ���Ŀ�¼�豸
Public Function GetDiscDevice(ByVal strDeviceName As String) As MsftDiscRecorder2
    Dim i As Integer
    Dim objRecord As MsftDiscRecorder2
    
    Set GetDiscDevice = Nothing
    For i = 0 To mobjDiscs.Count - 1
        Set objRecord = New MsftDiscRecorder2
        Call objRecord.InitializeDiscRecorder(mobjDiscs.Item(i))
        
        If UCase(GetBurnDeviceName(objRecord)) = UCase(strDeviceName) Then
            Set GetDiscDevice = objRecord
            Exit Function
        Else
            Set objRecord = Nothing
        End If
    Next i
End Function


'����豸�Ƿ�֧�ֿ�¼
Public Function CheckingDeviceIsBurn(ByVal strDeviceName As String) As Boolean
    Dim objRecord As MsftDiscRecorder2
    
    Set objRecord = GetDiscDevice(strDeviceName)

    '�ж��Ƿ�֧�ֿ�¼
    CheckingDeviceIsBurn = mobjDiscFormat.IsRecorderSupported(objRecord)
End Function


'�������Ƿ�׼������
Public Function CheckingDeviceIsReady(ByVal strDeviceName As String) As Boolean
    Dim objRecord As MsftDiscRecorder2
        
    Set objRecord = GetDiscDevice(strDeviceName)
      
    CheckingDeviceIsReady = mobjDiscFormat.IsCurrentMediaSupported(objRecord)
    
    Set objRecord = Nothing
End Function


'��ȡ��������
Public Function GetDiscName(ByVal strDeviceName As String)
    Dim strVolumnName  As String
    Dim FileSystemName As String
    Dim lngLength As Long
    Dim SerialNumber As Long

    
    strVolumnName = String(255, Chr(0))
    FileSystemName = String(255, Chr(0))


    Call GetVolumeInformation(Mid(strDeviceName, 1, 3), strVolumnName, 255, SerialNumber, 0, 0, FileSystemName, 255)
    
    GetDiscName = Replace(Trim(strVolumnName), Chr(0), "")
End Function


'��ȡָ��ý�����ͱ���
Private Function GetDiscMediaTypeAlias(ByVal lngMediaType As Long) As String
    GetDiscMediaTypeAlias = ""
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "GetCurDiscMediaTypeAlias", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Function
    End If
    
    Select Case lngMediaType
        Case 1
            GetDiscMediaTypeAlias = "CD-ROM"
        Case 2
            GetDiscMediaTypeAlias = "CD-R"
        Case 3
            GetDiscMediaTypeAlias = "CD-RW"
        Case 4
            GetDiscMediaTypeAlias = "DVD-ROM"
        Case 5
            GetDiscMediaTypeAlias = "DVD-RAM"
        Case 6
            GetDiscMediaTypeAlias = "DVD+R"
        Case 7
            GetDiscMediaTypeAlias = "DVD+RW"
        Case 8
            GetDiscMediaTypeAlias = "DVD+R DUAL LAYER"
        Case 9
            GetDiscMediaTypeAlias = "DVD-R"
        Case 10
            GetDiscMediaTypeAlias = "DVD-RW"
        Case 11
            GetDiscMediaTypeAlias = "DVD-R DUAL LAYER"
        Case 12
            GetDiscMediaTypeAlias = "RANDOMLY WRITABLE"
        Case 13
            GetDiscMediaTypeAlias = "DVD+RW DUAL LAYER"
        Case 14
            GetDiscMediaTypeAlias = "HD DVD-ROM"
        Case 15
            GetDiscMediaTypeAlias = "HD DVD-R"
        Case 16
            GetDiscMediaTypeAlias = "HD DVD-RAM"
        Case 17
            GetDiscMediaTypeAlias = "BD-ROM"
        Case 18
            GetDiscMediaTypeAlias = "BD-R"
        Case 19
            GetDiscMediaTypeAlias = "MAX"
        Case Else
            GetDiscMediaTypeAlias = "δ֪"
    End Select
End Function


'ȡ�õ�ǰý�����ͱ���
Public Function GetCurDiscMediaTypeAlias() As String
    GetCurDiscMediaTypeAlias = ""
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "GetCurDiscMediaType", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Function
    End If
    
    GetCurDiscMediaTypeAlias = GetDiscMediaTypeAlias(mobjDiscFormat.CurrentPhysicalMediaType)
End Function

'��ȡ��ǰ����ý����
Public Function GetCurDiscMediaType() As IMAPI_MEDIA_PHYSICAL_TYPE
    GetCurDiscMediaType = IMAPI_MEDIA_TYPE_UNKNOWN
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "GetCurDiscMediaType", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Function
    End If
    
    GetCurDiscMediaType = mobjDiscFormat.CurrentPhysicalMediaType
End Function




'ȡ�ù��̿�¼ʱ��֧�ֵ�������
Public Function GetCurSupportedSpeedCount() As Long
    GetCurSupportedSpeedCount = -1
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "EjectDisc", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Function
    End If

    GetCurSupportedSpeedCount = UBound(TransformArrayToOneDimension(mobjDiscFormat.SupportedWriteSpeeds))
End Function

'ȡ�ù��̿�¼��֧�ֵ�����
Public Function GetCurSupportedSpeed(ByVal lngSpeedIndex As Long) As Long
    GetCurSupportedSpeed = -1
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "EjectDisc", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Function
    End If
    
    GetCurSupportedSpeed = TransformArrayToOneDimension(mobjDiscFormat.SupportedWriteSpeeds)(lngSpeedIndex + 1)
End Function


'��ȡ֧�ֵ�ý���ʽ����
Public Function GetSupportedMediaTypeCount() As Long
    GetSupportedMediaTypeCount = -1
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "EjectDisc", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Function
    End If
    
    GetSupportedMediaTypeCount = UBound(TransformArrayToOneDimension(mobjDiscFormat.SupportedMediaTypes))
End Function

'ȡ�õ�ǰ�豸��֧�ֵĸ�ʽ����
Public Function GetCurSupportedMediaTypeAlias(ByVal lngSpeedIndex As Long) As String
    GetCurSupportedMediaTypeAlias = ""
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "EjectDisc", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Function
    End If
    
    GetCurSupportedMediaTypeAlias = GetDiscMediaTypeAlias(TransformArrayToOneDimension(mobjDiscFormat.SupportedMediaTypes)(lngSpeedIndex + 1))
End Function







'ȡ�ÿ��ù��̴�С
Public Function GetDiscFreeSize() As Double
    GetDiscFreeSize = 0
    
    If mobjCurRecord Is Nothing Then
        Exit Function
    End If
    
    GetDiscFreeSize = CDbl(mobjDiscFormat.FreeSectorsOnMedia) * CDbl(2048) - CDbl(mlngReserveSize * 1024)
End Function


'ȡ�ù����ܵĴ�С
Public Function GetDiscTotalSize() As Double
    GetDiscTotalSize = 0
    
    If mobjCurRecord Is Nothing Then
        Exit Function
    End If

    GetDiscTotalSize = CDbl(mobjDiscFormat.TotalSectorsOnMedia) * CDbl(2048)
End Function


'�����豸
Public Sub EjectDisc()
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "EjectDisc", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Sub
    End If
    
    mobjCurRecord.EjectMedia
End Sub


'�����������ݣ���ʱ�����ù��ܣ�����������.....��
Public Sub EraseDisc()

'    Call Err.Raise(-1, "EraseDisc", "δ��ʵ�ʲ��ԣ��Ĺ�����δ����......")
    
    
    If mobjCurRecord Is Nothing Then
        Call err.Raise(-1, "EraseDisc", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Sub
    End If

    Set mobjDiscEarse = Nothing
    Set mobjDiscEarse = New MsftDiscFormat2Erase

    '�жϴ����Ƿ�֧�����ݲ���
    If Not mobjDiscEarse.IsCurrentMediaSupported(mobjCurRecord) Then
        Call err.Raise(-1, "EraseDisc", "��ǰ���̲�֧�����ݲ���������")
        Exit Sub
    End If

    mobjDiscEarse.Recorder = mobjCurRecord
    mobjDiscEarse.ClientName = C_STR_BURN_CLIENTNAME


    Call mobjDiscEarse.EraseMedia
End Sub


'��ȡĿ¼�µ��ļ���С
Public Function GetPathSize(ByVal strBurnPath As String) As Double
On Error Resume Next
    Dim objFileSystem As New FileSystemObject
    Dim objDir As Folder
    
    Set objDir = objFileSystem.GetFolder(strBurnPath)
    
    GetPathSize = objDir.Size
    
    Set objDir = Nothing
    Set objFileSystem = Nothing
End Function


'������¼����
Public Sub CancelBurn()
    mblnIsCancelWrite = True
'    If mobjCurRecord Is Nothing Then
'        Exit Sub
'    End If
'
'    Call mobjDiscFormat.CancelWrite
End Sub


'�ж��Ƿ������ͬ����Ҫ��¼��Ŀ¼
Private Function HasBurnDirTree(ByVal strDir As String) As Boolean
    Dim i As Long
    
    HasBurnDirTree = False
    
    For i = 1 To mobjBurnDirResources.Count
        If mobjBurnDirResources.Item(i) = strDir Then
            HasBurnDirTree = True
            Exit Function
        End If
    Next i
End Function


'�ж��Ƿ������ͬ�Ŀ�¼�ļ�
Private Function HasBurnFile(ByVal strFile As String) As Boolean
    Dim i As Long
    
    HasBurnFile = False
    
    For i = 1 To mobjBurnFileResources.Count
        If InStr(mobjBurnFileResources.Item(i), strFile & C_STR_DISC_DIR_SPLIT) > 0 Then
            HasBurnFile = True
            Exit Function
        End If
    Next i
End Function


'���ӿ�¼Ŀ¼��
Public Sub AddBurnDirTree(ByVal strDir As String)
    '���������ͬ����Ҫ��¼��Ŀ¼�����˳�����
    If HasBurnDirTree(strDir) Then Exit Sub
    
    Call mobjBurnDirResources.Add(strDir)
End Sub


'���ӿ�¼�ļ�
Public Sub AddBurnFile(ByVal strFile As String, Optional ByVal strToDiscDir As String)
    '���������ͬ����Ҫ��¼���ļ������˳�����
    If HasBurnFile(strFile) Then Exit Sub
    
    Call mobjBurnFileResources.Add(strFile & C_STR_DISC_DIR_SPLIT & strToDiscDir)
End Sub



'�Ƴ���¼Ŀ¼��
Public Sub RemoveBurnDirTree(ByVal strDir As String)
    Dim i As Long
    
    For i = 1 To mobjBurnDirResources.Count
        If mobjBurnDirResources.Item(i) = strDir Then
            Call mobjBurnDirResources.Remove(i)
            Exit Sub
        End If
    Next i
End Sub

'�Ƴ�����Ҫ��¼���ļ�
Public Sub RemoveBurnFile(ByVal strFile As String)
    Dim i As Long
    
    For i = 1 To mobjBurnFileResources.Count
        If InStr(mobjBurnFileResources.Item(i), strFile & C_STR_DISC_DIR_SPLIT) > 0 Then
            Call mobjBurnFileResources.Remove(i)
            Exit Sub
        End If
    Next i
End Sub


'������д���¼��Ŀ¼
Public Sub ClearBurnDir()
    Dim i As Long
    
    For i = mobjBurnDirResources.Count To 1 Step -1
        Call mobjBurnDirResources.Remove(i)
    Next i
End Sub


'������д���¼���ļ�
Public Sub ClearBurnFile()
    Dim i As Long
    
    For i = mobjBurnFileResources.Count To 1 Step -1
        Call mobjBurnFileResources.Remove(i)
    Next i
End Sub


'��ȡ��¼��Դ���ܴ�С(��λΪ�ֽ�)
Public Function GetBurnResourceTotalSize() As Double
    Dim i As Long
    Dim lngDirSplitIndex As Long
    Dim strFilePath As String
    
    GetBurnResourceTotalSize = 0
    
    For i = 1 To mobjBurnDirResources.Count
        GetBurnResourceTotalSize = GetBurnResourceTotalSize + GetPathSize(mobjBurnDirResources.Item(i))
    Next i
    
    For i = 1 To mobjBurnFileResources.Count
        strFilePath = mobjBurnFileResources.Item(i)
        
        lngDirSplitIndex = InStr(strFilePath, C_STR_DISC_DIR_SPLIT)
        If lngDirSplitIndex > 0 Then strFilePath = Mid(strFilePath, 1, lngDirSplitIndex - 1)
        
        GetBurnResourceTotalSize = GetBurnResourceTotalSize + FileSystem.FileLen(strFilePath)
    Next i
    
End Function




'��ȡ��Ҫ��¼��Ŀ¼
Private Sub ReadBurnDir(objRootDir As IFsiDirectoryItem)
    Dim i As Long
        
    For i = 1 To mobjBurnDirResources.Count
        mstrBurnPath = mobjBurnDirResources.Item(i)
        
        Call DoBurnProcedureEvent("��ȡ��¼Ŀ¼��" & mstrBurnPath)
        
        Call objRootDir.AddTree(mobjBurnDirResources.Item(i), mblnIncludeBaseDir)
    Next i
End Sub


'��ȡ��Ҫ��¼���ļ�
Private Sub ReadBurnFile(objRootDir As IFsiDirectoryItem)
    Dim i As Long
    Dim j As Long
    Dim lngSplitIndex As Long
    
    Dim aryDir() As String
    Dim strFilePath As String
    Dim strDiscDir As String
    Dim strFileName As String
    Dim objFsiNewDirItem As IFsiDirectoryItem
    Dim objFsiSubDirItem As IFsiDirectoryItem
    Dim objFileStream As Stream
    
    For i = 1 To mobjBurnFileResources.Count
        strDiscDir = mobjBurnFileResources.Item(i)
        
        'ȡ�Ŀ�¼�ļ������Ŀ¼�ķָ������ַ�����
        lngSplitIndex = InStr(strDiscDir, C_STR_DISC_DIR_SPLIT)
        
        strFilePath = Mid(strDiscDir, 1, IIf(lngSplitIndex > 0, lngSplitIndex - 1, 255))
        strFileName = Mid(strFilePath, InStrRev(strFilePath, "\") + 1, 255)
        
        mstrBurnPath = strFilePath
        
        Call DoBurnProcedureEvent("��ȡ��¼�ļ���Ϣ��" & strFilePath)
        
        '�ж��Ƿ���Ҫ�ڹ����ϴ���Ŀ¼
        If lngSplitIndex > 0 Then
            strDiscDir = "\" & Mid(strDiscDir, lngSplitIndex + 1, 255) & "\"
            aryDir = Split(strDiscDir, "\")
            
            Call DoBurnProcedureEvent("��������Ŀ¼��" & Replace(strDiscDir, "\\", ""))
            
            Set objFsiNewDirItem = Nothing
            
            '�ڴ����ϴ���ָ��Ŀ¼
            For j = LBound(aryDir) To UBound(aryDir)
                If Trim(aryDir(j)) <> "" Then
                    On Error Resume Next
                    
                    If objFsiNewDirItem Is Nothing Then
'                        On Error GoTo errCreateDir
                        Set objFsiNewDirItem = objRootDir.Item(aryDir(j))
'errCreateDir:
                        
                        '���û�л�ȡ��ָ���Ĺ���Ŀ¼���򴴽�
                        If objFsiNewDirItem Is Nothing Then
                            Call objRootDir.AddDirectory(aryDir(j))
                            Set objFsiNewDirItem = objRootDir.Item(aryDir(j))
                        End If
                    Else
'                        On Error GoTo errCreateSubDir
                        Set objFsiSubDirItem = Nothing
                        Set objFsiSubDirItem = objFsiNewDirItem.Item(aryDir(j))
'errCreateSubDir:
                        If objFsiSubDirItem Is Nothing Then
                            Call objFsiNewDirItem.AddDirectory(aryDir(j))
                            Set objFsiNewDirItem = objFsiNewDirItem.Item(aryDir(j))
                        Else
                            Set objFsiNewDirItem = objFsiSubDirItem
                        End If
                    End If
                End If
            Next j
        End If
        
        If objFsiNewDirItem Is Nothing Then Set objFsiNewDirItem = objRootDir
        
        Call DoBurnProcedureEvent("��ȡ�ļ�����" & strFilePath)
        
        '��ȡ�ļ���
        Set objFileStream = New Stream
        
        '��������Ϊ���������ͣ����ܱ�֤��¼���ļ���ȷ
        objFileStream.type = adTypeBinary
        
        Call objFileStream.Open
        Call objFileStream.LoadFromFile(strFilePath)
        
        Call objFsiNewDirItem.AddFile(strFileName, objFileStream)
        
'        '����ط����ܶ������йرգ������ڴ�������ʱ�����������ܶ�ȡ������
'        Call DoBurnProcedureEvent("�ر��ļ�����" & strFilePath)
'        Call objFileStream.Close
        
        Set objFileStream = Nothing
        Set objFsiNewDirItem = Nothing
    Next i
End Sub


'׼����¼��Ҫ����Դ
Private Function ReadyBurnResource(ByRef objFileImage As MsftFileSystemImage, _
    ByVal lngMediaType As IMAPI_MEDIA_PHYSICAL_TYPE) As FsiStream
    
    Dim objRootDir As FsiDirectoryItem
    
    Set ReadyBurnResource = Nothing
    
    Call objFileImage.ChooseImageDefaultsForMediaType(lngMediaType)
    
    '���ÿ�¼���̺�����iso��ͨ������
On Error GoTo errHandle
    '����ͬһ���̽��ж�ο�¼ʱ�����ø����Խ��������ӦΪ�ö����Ѵӹ����ж�ȡ��ϵͳ����
    objFileImage.FileSystemsToCreate = FsiFileSystemISO9660 Or FsiFileSystemJoliet
errHandle:

On Error GoTo errRead
    objFileImage.VolumeName = mstrBurnVolumeName

    Set objRootDir = objFileImage.Root
    
    '������Ҫ��¼��Ŀ¼��
    Call DoBurnProcedureEvent("���ش���¼Ŀ¼...")
    Call ReadBurnDir(objRootDir)
    
    
    '������Ҫ��¼���ļ�
    Call DoBurnProcedureEvent("���ش���¼�ļ�...")
    Call ReadBurnFile(objRootDir)
    
    
    Call DoBurnProcedureEvent("������������...")
    
    Set ReadyBurnResource = objFileImage.CreateResultImage.ImageStream
    
    Exit Function
errRead:
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function


'��ʼ��¼
Public Sub StartBurn()
    Dim objBurnVerification As IBurnVerification
    Dim objFileStream As FsiStream
    Dim objIMultiSession    '�������������Ͷ��󣬲��ܶ�MultisessionInterfaces�ӿڸ�ֵ
    
    mblnIsCancelWrite = False
    
    If mobjCurRecord Is Nothing Then
        Call DoBurnProcedureEvent("��ǰ�豸��Ч������ִ�иò���...")
        Call err.Raise(-1, "StartBurn", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Sub
    End If
    
    If mobjBurnDirResources.Count <= 0 And mobjBurnFileResources.Count <= 0 Then
        Call DoBurnProcedureEvent("δ�ҵ���Ҫ��¼����Դ������ִ�иò�����")
        Call err.Raise(-1, "StartBurn", "δ�ҵ���Ҫ��¼����Դ������ִ�иò�����")
        Exit Sub
    End If
    
    Call DoBurnProcedureEvent("���ÿ�¼����...")
    
    Set mobjFileImage = Nothing
    Set mobjFileImage = New MsftFileSystemImage
    
    'ѡ���¼�豸
    Call mobjFileImage.ChooseImageDefaults(mobjCurRecord)
    
    '���ÿ�¼����
    If mlngWriteSpeed > 0 Then
        Call mobjDiscFormat.SetWriteSpeed(mlngWriteSpeed, True)
    End If
    
    '�����Ѻ�����
    mobjDiscFormat.ClientName = C_STR_BURN_CLIENTNAME
    
    mobjDiscFormat.ForceOverwrite = mblnIsOverWrite
    
    mobjDiscFormat.ForceMediaToBeClosed = mblnOnceMedia
    
    '���������Լ�鼶��
    Set objBurnVerification = mobjDiscFormat
    objBurnVerification.BurnVerificationLevel = mlngVerificationLevel
    
    mobjFileImage.FreeMediaBlocks = mobjDiscFormat.FreeSectorsOnMedia - Fix(mlngReserveSize * 1024 / 2048) '���Ӵ��̵�Ԥ����С
    
    '����������(��ȡ�������Ѵ��ڵ�����)
    If Not mobjDiscFormat.MediaHeuristicallyBlank Then
        Call DoBurnProcedureEvent("���ع�������...")
        
         Set objIMultiSession = mobjFileImage
        
        objIMultiSession.MultisessionInterfaces = mobjDiscFormat.MultisessionInterfaces
        

        mobjFileImage.ImportFileSystem
    End If

    '׼����¼�ļ�
On Error GoTo errHandle
    Call DoBurnProcedureEvent("׼����¼����...")
    
    Set objFileStream = ReadyBurnResource(mobjFileImage, IMAPI_MEDIA_TYPE_DISK) 'mobjDiscFormat.CurrentPhysicalMediaType
    
    If objFileStream Is Nothing Then
        Call DoBurnProcedureEvent("δ�ܳɹ���ȡ��Ҫ��¼����������������ֹ...")
        Call err.Raise(-1, "StartBurn", "δ�ܳɹ���ȡ��Ҫ��¼����������������ֹ��")
        Exit Sub
    End If
    
    DoEvents
    
    If mblnIsCancelWrite Then
        Call DoBurnProcedureEvent("��¼�����ѱ�����...")
        Call err.Raise(-1, "StartBurn", "��¼�����ѱ�������")
        Exit Sub
    End If
    
    '��ʼ��¼�ļ�
    Call DoBurnProcedureEvent("ִ�п�¼����...")
    
    Call mobjDiscFormat.Write(objFileStream)
    
    
    Call DoBurnProcedureEvent("��ɿ�¼����������...")
    Call mobjCurRecord.EjectMedia
    
    Set objBurnVerification = Nothing
    Set objFileStream = Nothing
    Set mobjFileImage = Nothing
    
    
    Call DoBurnProcedureEvent("��������ݿ�¼����...")
    
    Exit Sub
errHandle:
    Set objBurnVerification = Nothing
    Set objFileStream = Nothing

    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Sub



Private Sub DoBurnProcedureEvent(ByVal strState As String)
On Error Resume Next
'    DoEvents
    
    RaiseEvent OnBurnProcedureEvent(strState)
    
'    DoEvents
End Sub




'��¼ISO�ļ�
Public Sub BurnIso(ByVal strIsoFile As String)
    Dim objBurnVerification As IBurnVerification
    Dim objIsoStream As Stream
    
    mblnIsCancelWrite = False
    
    If mobjCurRecord Is Nothing Then
        Call DoBurnProcedureEvent("��ǰ�豸��Ч������ִ�иò���...")
        Call err.Raise(-1, "BurnIso", "��ǰ�豸��Ч������ִ�иò�����")
        Exit Sub
    End If
    
    Call DoBurnProcedureEvent("���ÿ�¼����...")
        
    '���ÿ�¼����
    If mlngWriteSpeed > 0 Then
        Call mobjDiscFormat.SetWriteSpeed(mlngWriteSpeed, True)
    End If
        
    '�����Ѻ�����
    mobjDiscFormat.ClientName = C_STR_BURN_CLIENTNAME
    
    mobjDiscFormat.ForceOverwrite = mblnIsOverWrite
    
    mobjDiscFormat.ForceMediaToBeClosed = mblnOnceMedia
    
    '���������Լ�鼶��
    Set objBurnVerification = mobjDiscFormat
    objBurnVerification.BurnVerificationLevel = mlngVerificationLevel
    
    '����������
    If Not mobjDiscFormat.MediaHeuristicallyBlank Then
        Call DoBurnProcedureEvent("��¼ISO����ʱ�������հ׹��̡�")
        Call err.Raise(-1, "BurnIso", "��¼ISO����ʱ�������հ׹��̡�")
        
        Exit Sub
    End If
    
    '׼����¼�ļ�
On Error GoTo errHandle
    Call DoBurnProcedureEvent("��ȡISO����...")
    
    Set objIsoStream = New Stream
    objIsoStream.type = adTypeBinary
    
    Call objIsoStream.Open
    Call objIsoStream.LoadFromFile(strIsoFile)
    
    
    If mblnIsCancelWrite Then
        Call DoBurnProcedureEvent("��¼�����ѱ�����...")
        Call err.Raise(-1, "StartBurn", "��¼�����ѱ�������")
        Exit Sub
    End If
    
    '��ʼ��¼�ļ�
    Call DoBurnProcedureEvent("ִ�п�¼����...")
    
    Call mobjDiscFormat.Write(objIsoStream)
    
    
    Call DoBurnProcedureEvent("��ɿ�¼����������...")
    Call mobjCurRecord.EjectMedia
    
    Call objIsoStream.Close
    
    Set objIsoStream = Nothing
    Set objBurnVerification = Nothing
    
    Exit Sub
errHandle:
    Set objIsoStream = Nothing
    Set objBurnVerification = Nothing

    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Sub

'����ΪISO
Public Sub SaveIso(ByVal strIsoPath As String)
    Dim objIsoStream As Stream
    Dim objFileStream As FsiStream
    Dim dblInSize As Double
    Dim dblOutSize As Double
    
    Dim tagS As tagSTAT
    Dim objVaraint
    Dim tagVaraint
    Dim bojImageStream As Stream
    
'    'vb����֧�ֲ��ã��в��ܱ���ΪISO�ļ�...
'    Exit Sub
    
    Set mobjFileImage = Nothing
    Set mobjFileImage = New MsftFileSystemImage
    
    Call DoBurnProcedureEvent("׼��ISO��������...")
    
    Set objFileStream = ReadyBurnResource(mobjFileImage, IMAPI_MEDIA_TYPE_DISK)
    
    If objFileStream Is Nothing Then
        Call DoBurnProcedureEvent("δ�ܳɹ���ȡ�����ISO��������������ֹ...")
        Call err.Raise(-1, "StartBurn", "δ�ܳɹ���ȡ�����ISO��������������ֹ��")
        Exit Sub
    End If
    
    Exit Sub
    
    Set objIsoStream = New Stream
    
    objIsoStream.type = adTypeBinary
    objIsoStream.Mode = adModeWrite
    
    
    objVaraint = objFileStream
'
'    Set bojImageStream = objVaraint
    
'    Call objVaraint.Stat(tagVaraint, 1)
    Call objVaraint.RemoteCopyTo(objIsoStream, 255, dblInSize, dblOutSize)
'    Call objFileStream.Clone(objIsoStream)
    
    Call objIsoStream.SaveToFile(strIsoPath, adSaveCreateOverWrite)
    
    objIsoStream.Close
    
    Set objIsoStream = Nothing
    Set objFileStream = Nothing
End Sub



Private Sub Class_Initialize()
    Set mobjDiscs = New MsftDiscMaster2
    Set mobjDiscFormat = New MsftDiscFormat2Data
    Set mobjFileImage = Nothing
    Set mobjCurRecord = Nothing
    Set mobjDiscEarse = Nothing
    Set mobjBurnDirResources = New Collection
    Set mobjBurnFileResources = New Collection
    
    mstrBurnDeviceName = ""
    mstrBurnVolumeName = Format(zlDatabase.Currentdate, "yyyymmdd")
    mblnIsOverWrite = True
    mlngVerificationLevel = ivlFull
    mblnOnceMedia = False
    mblnIncludeBaseDir = False
    mblnIsCancelWrite = False
    mstrBurnPath = ""
    mlngReserveSize = 0 '10 * 1024 '(10M)
End Sub

Private Sub Class_Terminate()
    Set mobjBurnFileResources = Nothing
    Set mobjBurnDirResources = Nothing
    Set mobjDiscEarse = Nothing
    Set mobjCurRecord = Nothing
    Set mobjFileImage = Nothing
    Set mobjDiscFormat = Nothing
    Set mobjDiscs = Nothing
End Sub


'��ȡд��״̬
Private Function GetWriteAction(iFormat2DataArgs As IDiscFormat2DataEventArgs)
    Select Case iFormat2DataArgs.CurrentAction
        Case 1
            GetWriteAction = "��ʽ��ý��..."
        Case 2
            GetWriteAction = "��ʼ���豸..."
        Case 3
            GetWriteAction = "�Ż�д�����..."
        Case 4
            GetWriteAction = "д�����ݵ�ý��..."
        Case 5
            GetWriteAction = "ͬ�����岢�رմŵ�..."
        Case 6
            GetWriteAction = "�ɹ�д������..."
        Case 7
            GetWriteAction = "���������..."
        Case Else
            GetWriteAction = "ȷ��ý��֧����..."
    End Select
End Function

Private Sub mobjDiscFormat_Update(ByVal Object As Object, ByVal Progress As Object)
On Error Resume Next
    Dim iFormat2DataArgs As IDiscFormat2DataEventArgs
    Dim strState As String
    Dim args As New clsImapi2BurnArgs
    
    If mblnIsCancelWrite Then
        Call Object.CancelWrite
    End If
    
    Set iFormat2DataArgs = Progress

    strState = GetWriteAction(iFormat2DataArgs)
    
    args.ElapsedTime = iFormat2DataArgs.ElapsedTime
    args.FreeSystemBuffer = iFormat2DataArgs.FreeSystemBuffer
    args.LastReadLba = iFormat2DataArgs.LastReadLba
    args.LastWrittenLba = iFormat2DataArgs.LastWrittenLba
    args.RemainingTime = iFormat2DataArgs.RemainingTime
    args.SectorCount = iFormat2DataArgs.SectorCount
    args.StartLba = iFormat2DataArgs.StartLba
    args.TotalSystemBuffer = iFormat2DataArgs.TotalSystemBuffer
    args.TotalTime = iFormat2DataArgs.TotalTime
    args.UsedSystemBuffer = iFormat2DataArgs.UsedSystemBuffer
    
'    DoEvents
    
    RaiseEvent OnBurnEvent(strState, args)
    
'    DoEvents
    
    Set args = Nothing
End Sub

Private Sub mobjDiscs_NotifyDeviceAdded(ByVal Object As Object, ByVal uniqueId As String)
    '......
End Sub

Private Sub mobjDiscs_NotifyDeviceRemoved(ByVal Object As Object, ByVal uniqueId As String)
    '......
End Sub

Private Sub mobjFileImage_Update(ByVal Object As Object, ByVal currentFile As String, ByVal copiedSectors As Long, ByVal totalSectors As Long)
On Error Resume Next
    Dim strCurBurnPath As String
'    Dim dblCurFileSize As Double
    
    strCurBurnPath = IIf(mblnIncludeBaseDir, Mid(mstrBurnPath, 1, InStrRev(mstrBurnPath, "\") - 1), strCurBurnPath)
'
'    dblCurFileSize = FileSystem.FileLen(strCurBurnPath & currentFile)
    
'    DoEvents
    
    RaiseEvent OnReadyFileEvent(currentFile, strCurBurnPath, copiedSectors, totalSectors)
    
'    DoEvents
End Sub