VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4365
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4365
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '??Ļ????
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   1005
      TabIndex        =   8
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Label lbltag 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "????"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4755
      TabIndex        =   9
      Top             =   1815
      Width           =   180
   End
   Begin VB.Image imgPic 
      Height          =   2745
      Left            =   150
      Picture         =   "frmSplash.frx":5D0A2
      Top             =   420
      Width           =   1260
   End
   Begin VB.Label LblProductName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "????"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1590
      TabIndex        =   7
      Top             =   1350
      Width           =   4650
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ʹ??Ȩ???ڣ?"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   6
      Top             =   2205
      Width           =   1080
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "??Ʒ?????̣?"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   5
      Top             =   3030
      Width           =   1080
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "????֧???̣?"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   4
      Top             =   2610
      Width           =   1080
   End
   Begin VB.Label lblGrant 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   3
      Top             =   2205
      Width           =   90
   End
   Begin VB.Label lbl????֧???? 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   2
      Top             =   2610
      Width           =   90
   End
   Begin VB.Label lbl?????? 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   1
      Top             =   3030
      Width           =   90
   End
   Begin VB.Image ImgIndicate 
      Appearance      =   0  'Flat
      Height          =   780
      Left            =   165
      Picture         =   "frmSplash.frx":5D923
      Top             =   3390
      Width           =   780
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "???棺????????????????????????ʹ??????֤??????δ????Ȩ???ɣ??κ??˲??ø??ơ????ۼ????ܴ??????????򽫳е?ȫ?????????Ρ?"
      Height          =   465
      Left            =   1065
      TabIndex        =   0
      Top             =   3825
      Width           =   5490
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Relogin(ByVal FrmMainObj As Object)
    Unload FrmMainObj
    Call Main
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    gdtStart = 0
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    gdtStart = 0
End Sub

Private Sub LblProductName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    gdtStart = 0
End Sub

Private Sub lblGrant_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    gdtStart = 0
End Sub

Private Sub lblWarning_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    gdtStart = 0
End Sub

Private Sub lbl????֧????_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    gdtStart = 0
End Sub

Private Sub lbl??????_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    gdtStart = 0
End Sub

