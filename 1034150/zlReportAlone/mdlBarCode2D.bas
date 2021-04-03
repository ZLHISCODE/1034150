Attribute VB_Name = "mdlBarCode2D"
Option Explicit

Public Function DrawBarCode2D(ByVal strText As String, picTemp As PictureBox, Optional lngSize As Long) As StdPicture
'���ܣ�����QR��ά����ͼƬ
'������lngSize=������TwipΪ��λ��ͼƬ���ʳߴ�
'���أ�QR��ά����ͼƬ�������ǷŴ��˵�BMPͼƬ
    Static objQRMaker As Object
    Static intInstall As Integer '0-δ���,1-�Ѱ�װ,-1-δ��װ
    
    Dim strFile As String
    Dim objPic As StdPicture
    
    If intInstall = 0 Then
        On Local Error Resume Next
        Set objQRMaker = CreateObject("QRMAKER.QRmakerCtrl.1")
        Err.Clear: On Local Error GoTo 0
        intInstall = IIF(objQRMaker Is Nothing, -1, 1)
        
        '��ʼ���ؼ�����
        If intInstall = 1 Then
            With objQRMaker
                .GapAdjust = 0 'GpAjOff
                .LanguageCode = 1
                
                .EccLevel = 1 'M
                .ModelNo = 2 'Model2
                .Rotate = 0 'D0
                
                .QuietZone = 1
                
                .ForeWColor = vbWhite
                .ForeBColor = vbBlack
            End With
        End If
    End If
    
    lngSize = 0
    
    If intInstall = -1 Then
        picTemp.AutoRedraw = True
        picTemp.BorderStyle = 0
        picTemp.ScaleMode = vbTwips
        picTemp.Cls
        
        lngSize = picTemp.ScaleX(50, vbPixels, vbTwips)
        picTemp.Width = lngSize: picTemp.Height = lngSize
        
        picTemp.DrawWidth = 1
        picTemp.Line (0, 0)-(picTemp.Width, picTemp.Height), vbBlack
        picTemp.Line (picTemp.Width, 0)-(0, picTemp.Height), vbBlack
        picTemp.DrawWidth = 2
        picTemp.Line (0, 0)-(picTemp.Width, picTemp.Height), vbBlack, B
        
        Set DrawBarCode2D = picTemp.Image
        picTemp.Cls
    Else
        If strText = "" Then strText = "����������Ϣ��ҵ���޹�˾"

        objQRMaker.InputData = strText
        objQRMaker.Refresh

        With picTemp
            .AutoRedraw = True
            .BorderStyle = vbBSNone
            .ScaleMode = vbTwips
            .Cls
            .Width = objQRMaker.Picture.Width
            .Height = objQRMaker.Picture.Height
        End With

        '�����QRMaker�����ʵ�ʴ�С
        'lngSize = objQRMaker.Picture.Width
        'Ϊ������ǰ������Ӧ���������С���������е����㷽ʽ
        lngSize = picTemp.ScaleX(2 * (objQRMaker.NumCell + objQRMaker.QuietZone * 2), vbPixels, vbTwips)

        picTemp.PaintPicture objQRMaker.Picture, 0, 0, picTemp.Width, picTemp.Height
        Set DrawBarCode2D = picTemp.Image
    End If
End Function
