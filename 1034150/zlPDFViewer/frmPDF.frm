VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmPDF 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin SHDocVwCtl.WebBrowser WebSub 
      Height          =   1635
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   2085
      ExtentX         =   3678
      ExtentY         =   2884
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrFileType As String           '文件类型
Private WithEvents mobjPDF As VBControlExtender
Attribute mobjPDF.VB_VarHelpID = -1

Public Function NewControl(objParent As Object, ByVal strControlClass As String, ByVal strName As String, Optional objPart As Object) As Object
    Dim objCrl As Object
    
    '加入协议，只能加一次，第二次会出错
    On Error Resume Next
    Call Licenses.Add(strControlClass)
    On Error GoTo errhand
    
    '产生动态控件
    If objPart Is Nothing Then
        Set objCrl = objParent.Controls.Add(strControlClass, strName)
    Else
        Set objCrl = objParent.Controls.Add(strControlClass, strName)
        Set objCrl.Container = objPart
        objCrl.Move 0, 0, objPart.Width, objPart.Height
        objCrl.ZOrder
        objCrl.Visible = False
    End If
    
    
    Set NewControl = objCrl
    Exit Function
errhand:


End Function

Private Sub Form_Load()
    Set mobjPDF = NewControl(Me, "mobjPDF.PDF", "PDF", Me)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Not mobjPDF Is Nothing Then mobjPDF.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    WebSub.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Function LoadFile(ByVal strFile As String) As Boolean
      '功能：加载文件
1         On Error GoTo LoadFile_Error

          'PDF文件
2         strFile = UCase(strFile)
3         If strFile Like "*.PDF" Then    'PDF文件
4             If Not mobjPDF Is Nothing Then
5                 LoadFile = mobjPDF.LoadFile(strFile)
6                 mobjPDF.Visible = True
7                 WebSub.Visible = False
8             Else
9                 Call WebSub.Navigate(strFile)
10                WebSub.Visible = True
11            End If
12            mstrFileType = "PDF"
13        ElseIf strFile Like "*.HTML" Or strFile Like "*.HTM" Then     'html文件
14            Call WebSub.Navigate(strFile)
15            If Not mobjPDF Is Nothing Then mobjPDF.Visible = False
16            WebSub.Visible = True
17            mstrFileType = "HTML"
18        ElseIf strFile Like "*.XPS" Or strFile Like "*.OXPS" Then       'XPS文件
19            Call WebSub.Navigate(strFile)
20            If Not mobjPDF Is Nothing Then mobjPDF.Visible = False
21            WebSub.Visible = True
22            mstrFileType = "XPS"
23        End If

24        Exit Function
LoadFile_Error:
25        MsgBox "zlPDFViewer frmPDF 执行(LoadFile)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl
26        Err.Clear

End Function

Public Function PrintFile(ByVal intType As Integer) As Boolean
    Dim RetVal As Long
    Dim strSQL As String

    '功能：打印
    '参数: intType 打印方式,0-直接打印,1-交互打印
    If mstrFileType = "PDF" And Not mobjPDF Is Nothing Then
        If intType = 0 Then
            If Not mobjPDF Is Nothing Then mobjPDF.printAll
        ElseIf intType = 1 Then
            If Not mobjPDF Is Nothing Then mobjPDF.printWithDialog
        End If
    Else
        If intType = 0 Then
            WebSub.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
        Else
            WebSub.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
        End If
    End If
End Function

Public Function WaitTime(ByVal lng序号 As Long, ByVal strFilePath As String, ByVal strName As String) As String
'功能:打印等待
'参数:strFilePath文件路径
'     strName 报告名称
    WaitTime = frmWait.ShowMe(Me, lng序号, strFilePath, strName)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mobjPDF = Nothing
End Sub
