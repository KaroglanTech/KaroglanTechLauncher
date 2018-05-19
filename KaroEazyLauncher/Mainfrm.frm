VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Mainfrm 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "LauncherMain"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13485
   Icon            =   "Mainfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   13485
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   7680
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5265
      ScaleWidth      =   9225
      TabIndex        =   13
      Top             =   2040
      Width           =   9255
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   5775
         Left            =   -240
         TabIndex        =   14
         Top             =   -240
         Width           =   9735
         ExtentX         =   17171
         ExtentY         =   10186
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
         Location        =   ""
      End
   End
   Begin KaroglanLauncher.N_Shape StartBtn 
      Height          =   855
      Left            =   11280
      TabIndex        =   3
      Top             =   7560
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1508
      Picture_Normal  =   "Mainfrm.frx":492A
      Picture_Down    =   "Mainfrm.frx":4946
      Picture_Hover   =   "Mainfrm.frx":4962
      Stretch         =   0   'False
      Caption         =   "开始游戏"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin VB.TextBox javapathtxt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   5520
      Width           =   3015
   End
   Begin VB.TextBox memorytxt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Text            =   "1024"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox usernametxt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   0
      Top             =   3120
      Width           =   3015
   End
   Begin KaroglanLauncher.N_Shape topBtn 
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      Picture_Normal  =   "Mainfrm.frx":497E
      Picture_Down    =   "Mainfrm.frx":499A
      Picture_Hover   =   "Mainfrm.frx":49B6
      Stretch         =   0   'False
      Caption         =   "按钮2"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape AutoMemoryBtn 
      Height          =   375
      Left            =   11880
      TabIndex        =   6
      Top             =   4320
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      Picture_Normal  =   "Mainfrm.frx":49D2
      Picture_Down    =   "Mainfrm.frx":49EE
      Picture_Hover   =   "Mainfrm.frx":4A0A
      Stretch         =   0   'False
      Caption         =   "自动内存"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape AutoJavaBtn 
      Height          =   375
      Left            =   10440
      TabIndex        =   7
      Top             =   6120
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      Picture_Normal  =   "Mainfrm.frx":4A26
      Picture_Down    =   "Mainfrm.frx":4A42
      Picture_Hover   =   "Mainfrm.frx":4A5E
      Stretch         =   0   'False
      Caption         =   "自动查找"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape ScrJavaBtn 
      Height          =   375
      Left            =   11880
      TabIndex        =   8
      Top             =   6120
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      Picture_Normal  =   "Mainfrm.frx":4A7A
      Picture_Down    =   "Mainfrm.frx":4A96
      Picture_Hover   =   "Mainfrm.frx":4AB2
      Stretch         =   0   'False
      Caption         =   "手动选择"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape topBtn 
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   9
      Top             =   1560
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      Picture_Normal  =   "Mainfrm.frx":4ACE
      Picture_Down    =   "Mainfrm.frx":4AEA
      Picture_Hover   =   "Mainfrm.frx":4B06
      Stretch         =   0   'False
      Caption         =   "按钮3"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape topBtn 
      Height          =   495
      Index           =   3
      Left            =   4800
      TabIndex        =   10
      Top             =   1560
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      Picture_Normal  =   "Mainfrm.frx":4B22
      Picture_Down    =   "Mainfrm.frx":4B3E
      Picture_Hover   =   "Mainfrm.frx":4B5A
      Stretch         =   0   'False
      Caption         =   "按钮4"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape buttomBtn 
      Height          =   495
      Index           =   0
      Left            =   6240
      TabIndex        =   11
      Top             =   7920
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      Picture_Normal  =   "Mainfrm.frx":4B76
      Picture_Down    =   "Mainfrm.frx":4B92
      Picture_Hover   =   "Mainfrm.frx":4BAE
      Stretch         =   0   'False
      Caption         =   "按钮1"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape topBtn 
      Height          =   495
      Index           =   4
      Left            =   6360
      TabIndex        =   12
      Top             =   1560
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      Picture_Normal  =   "Mainfrm.frx":4BCA
      Picture_Down    =   "Mainfrm.frx":4BE6
      Picture_Hover   =   "Mainfrm.frx":4C02
      Stretch         =   0   'False
      Caption         =   "按钮5"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape topBtn 
      Height          =   495
      Index           =   5
      Left            =   7920
      TabIndex        =   15
      Top             =   1560
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      Picture_Normal  =   "Mainfrm.frx":4C1E
      Picture_Down    =   "Mainfrm.frx":4C3A
      Picture_Hover   =   "Mainfrm.frx":4C56
      Stretch         =   0   'False
      Caption         =   "按钮6"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape topBtn 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      Picture_Normal  =   "Mainfrm.frx":4C72
      Picture_Down    =   "Mainfrm.frx":4C8E
      Picture_Hover   =   "Mainfrm.frx":4CAA
      Stretch         =   0   'False
      Caption         =   "按钮1"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape buttomBtn 
      Height          =   495
      Index           =   1
      Left            =   7920
      TabIndex        =   18
      Top             =   7920
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      Picture_Normal  =   "Mainfrm.frx":4CC6
      Picture_Down    =   "Mainfrm.frx":4CE2
      Picture_Hover   =   "Mainfrm.frx":4CFE
      Stretch         =   0   'False
      Caption         =   "按钮2"
      BackGround      =   15725042
      BackColorNormal =   33023
      BackColorDown   =   16576
      BorderColorNormal=   33023
      BorderColorHover=   8438015
      BorderColorDown =   16576
      FontSize        =   10
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin KaroglanLauncher.N_Shape buttomBtn 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   19
      Top             =   7920
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      Picture_Normal  =   "Mainfrm.frx":4D1A
      Picture_Down    =   "Mainfrm.frx":4D36
      Picture_Hover   =   "Mainfrm.frx":4D52
      Stretch         =   0   'False
      Caption         =   "按钮3"
      BackGround      =   15725042
      BackColorNormal =   1511338
      BackColorHover  =   7176406
      BackColorDown   =   788330
      BorderColorNormal=   1511338
      BorderColorHover=   7176406
      BorderColorDown =   788330
      FontSize        =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   16777215
      ForeColorHover  =   16777215
      ForeColorDown   =   16777215
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin VB.Label sTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "stitle"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   2640
      TabIndex        =   21
      Top             =   720
      Width           =   750
   End
   Begin VB.Label mainTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mainTitle"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   480
      TabIndex        =   20
      Top             =   480
      Width           =   2145
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户端版本：null"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   10560
      TabIndex        =   16
      Top             =   6720
      Width           =   1650
   End
   Begin VB.Label CloseL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   13200
      TabIndex        =   4
      Top             =   120
      Width           =   210
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FEUI As New FEUI
Dim LT As New LauncherTools
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Type topBtnT
    btnName As String
    btnLink As String
End Type
Private Type buttomBtnT
    btnVis As Boolean
    btnName As String
    btnLink As String
End Type

Dim KaroUIPath As String
Private Sub AutoJavaBtn_Click()
    javapathtxt.Text = LT.Findjava
End Sub

Private Sub AutoMemoryBtn_Click()
    memorytxt.Text = LT.RunMem
End Sub

Private Sub buttomBtn_Click(Index As Integer)
    Select Case Index
    Case 0
        LT.OpenURL Me, buttomBtn(Index).Tag
    Case 1
        LT.OpenURL Me, buttomBtn(Index).Tag
    Case 2
        Shell App.path & Right(buttomBtn(Index).Tag, Len(buttomBtn(Index).Tag) - 1), 1
        Unload Me
    End Select
End Sub

Private Sub CloseL_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Text1 = Dir("C:\WINDOWS\SysWOW64\javaw.exe")
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If Dir(App.path & "\setting.ini") <> "" Then
        usernametxt.Text = GetINI("setting", "username", App.path & "\setting.ini")
        memorytxt.Text = GetINI("setting", "memory", App.path & "\setting.ini")
        javapathtxt.Text = GetINI("setting", "javapath", App.path & "\setting.ini")
    Else
        Call AutoJavaBtn_Click
    End If
    KaroUIPath = App.path & "\KaroLauncher\ui.ini"
    Me.Picture = LoadPicture(App.path & "\KaroLauncher\kbg.jpg.jyp")
    WebBrowser1.Navigate GetINI("setting", "imageNoticeURL", KaroUIPath)
    Dim tmp As String
    Open App.path & "\ColorUpdate\version.ini" For Input As #1
    Line Input #1, tmp
    Close #1
    Label1.Caption = "客户端版本：" & tmp
    
    btnIniti
End Sub
Private Sub btnIniti()
    mainTitle.Caption = GetINI("setting", "title", KaroUIPath)
    sTitle.Caption = GetINI("setting", "stitle", KaroUIPath)
    sTitle.Left = mainTitle.Left + mainTitle.Width + 200
    Dim i As Integer
    For i = 0 To 5
        topBtn(i).Caption = GetINI("topbtn", "btn" & i + 1, KaroUIPath)
        topBtn(i).Tag = GetINI("topbtn", "btn" & i + 1 & "link", KaroUIPath)
    Next i
    For i = 0 To 2
        buttomBtn(i).Visible = GetINI("buttombtn", "btn" & i + 1 & "vis", KaroUIPath)
        buttomBtn(i).Caption = GetINI("buttombtn", "btn" & i + 1, KaroUIPath)
        buttomBtn(i).Tag = GetINI("buttombtn", "btn" & i + 1 & "link", KaroUIPath)
    Next i
End Sub

Private Sub buttomBtn2_Click()
    LT.OpenURL Me, "https://raw.githubusercontent.com/Tollainmear/PicRepo/master/%E8%B5%9E%E5%8A%A9%E5%88%97%E8%A1%A81.3.jpg"
End Sub

Private Sub buttomBtn3_Click()
    Shell App.path & "\hmclcore.exe", 1
    Unload Me
    End
End Sub


Private Sub ScrJavaBtn_Click()
    Dim ftmp As String
    ftmp = GetDialog("OPEN", "查找javaw.exe", "javaw.exe", "Javaw可执行应用程序", "exe")
    If ftmp <> "" Then javapathtxt.Text = ftmp
End Sub

Private Sub topBtn5_Click()
    '    If Dir(App.Path & "\PostInfo.exe") <> "" Then
    '    Shell "cmd /c " & App.Path & "\PostInfo.exe"
    'Else
    '    MsgBox "该组件已损坏"
    'End If
    
End Sub

Private Sub StartBtn_Click()
    If Dir(javapathtxt.Text) = "" Then
        MsgBox "Java不存在"
        Exit Sub
    End If
    If usernametxt.Text = "" Or javapathtxt.Text = "" Or memorytxt.Text = "" Then MsgBox "请填写设置!": Exit Sub
    
    If Dir(App.path & "\ColorProtect.exe") <> "" Then Shell App.path & "\ColorProtect.exe", 1
    
    Open App.path & "\Setting.ini" For Output As #1
    Print #1, "[setting]"
    Print #1, "username=" & usernametxt.Text
    Print #1, "memory=" & memorytxt.Text
    Print #1, "javapath=" & javapathtxt.Text
    Close #1
    
    
    'Shell App.Path & "\ColorLauncherCore.exe " & usernametxt & "," & memorytxt & "," _
    & javapathtxt & "," & "1.10.2-forge1.10.2-12.18.3.2185,,,,,,", 1
    
    launchGame usernametxt.Text, memorytxt.Text, javapathtxt.Text, "1.12.2-forge1.12.2-14.23.3.2655"
    WaitFrm.WaitFrm Me, "游戏正在准备 即将为您启动" & vbCrLf & "整个过程大概在3分-8分钟,请耐心等待..", 190, 20000
    
    If Dir(App.path & "\ColorProtect.exe") <> "" Then
        Shell App.path & "\ColorProtect.exe", 1
    End If
    
    'Unload Me
    'End
End Sub


Private Sub Timer1_Timer()
    'On Error Resume Next
    If Dir(App.path & "\refresh.jyc") <> "" Then
        Me.WebBrowser1.Refresh
        Kill App.path & "\refresh.jyc"
    End If
    Timer1.Enabled = False
End Sub

Private Sub topBtn_Click(Index As Integer)
    Select Case Index
    Case 0
        LT.OpenURL Me, topBtn(Index).Tag
    Case 1
        LT.OpenURL Me, topBtn(Index).Tag
    Case 2
        LT.OpenURL Me, topBtn(Index).Tag
    Case 3
        LT.OpenURL Me, topBtn(Index).Tag
    Case 4
        Shell pointPath(topBtn(Index).Tag), 1
    Case 5
        Shell "explorer.exe " & App.path & Right(topBtn(Index).Tag, Len(topBtn(Index).Tag) - 1), 1
    End Select
End Sub
Private Function pointPath(path As String) As String
    Dim tmp As String
    If Left(pointPath, 1) = "." Then
        pointPath = App.path & Right(path, Len(path) - 1)
    Else
        pointPath = path
    End If
End Function
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    WebBrowser1.Silent = True
End Sub
Private Sub WebBrowser1_DownloadBegin()
    WebBrowser1.Silent = True
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState <> 2 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
    End If
End Sub
