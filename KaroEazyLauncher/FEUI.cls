VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FEUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FlatEazyUI
'BY: JayshonYves
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Dim WithEvents frmclose As Label
Attribute frmclose.VB_VarHelpID = -1
Dim WithEvents SmallCt As Label
Attribute SmallCt.VB_VarHelpID = -1
Dim WithEvents frmbigs As Shape
Attribute frmbigs.VB_VarHelpID = -1
Dim WithEvents frmbig As Label
Attribute frmbig.VB_VarHelpID = -1
Dim WithEvents Frmad As Form
Attribute Frmad.VB_VarHelpID = -1
Dim WithEvents Frmdisi As Timer
Attribute Frmdisi.VB_VarHelpID = -1
Dim WithEvents Frmdiso As Timer
Attribute Frmdiso.VB_VarHelpID = -1
Dim FrmObj As Form
Dim QLineC
Dim wCtColor
Dim SmallB As Boolean, BigB As Boolean, CloseB As Boolean
Dim sMove As Boolean
Dim Frmalpha As Integer
Dim Frmdis As Boolean
Dim i As Integer
Public Function FEUI(frm As Form, SmallCtB As Boolean, BigCtB As Boolean, _
    CloseCtB As Boolean, FrmTitle As Boolean, FrmMove As Boolean, _
    Optional Mdisi As Boolean, Optional Mdiso As Boolean, _
    Optional FrmBC As String, Optional Linecolor As String, _
    Optional CtColor As String, Optional Alpha As Integer) As String
    'FEUI(窗体,最小化按钮,最大化按钮,关闭按钮,窗体标题,窗体移动,淡入,淡出,窗体颜色,边框颜色,控制按纽颜色,透明度)
    On Error Resume Next
    '设置默认参数
    
    
    If AlphaAm = True Then Alpha = 0
    If Linecolor = "" Then Linecolor = &HFFFF00
    If CtColor = "" Then CtColor = vbWhite
    If Alpha > 255 Then Alpha = 255
    If Alpha < 0 Then Alpha = 0
    
    SmallB = SmallCtB
    BigB = BigCtB
    CloseB = CloseCtB
    sTitle = FrmTitle
    sMove = FrmMove
    Set FrmObj = frm
    Set Frmad = frm
    QLineC = Linecolor
    wCtColor = CtColor
    Frmalpha = Alpha
    Frmdis = Mdiso
    
    With frm
        .AutoRedraw = True
        .BackColor = FrmBC
    End With
    
    '淡入
    If Mdisi Then
        Set Frmdisi = frm.Controls.Add("vb.timer", "Frmdisi1")
        With Frmdisi
            .Interval = 10
            .Enabled = True
        End With
    End If
    
    '淡出
    If Mdiso Then
        Set Frmdiso = frm.Controls.Add("vb.timer", "Frmdisi2")
        With Frmdiso
            .Interval = 10
            .Enabled = False
        End With
    End If
    
    
    '窗体标题
    If FrmTitle Then
        Set frmCaption = frm.Controls.Add("vb.label", "TitleCaption")
        With frmCaption
            .Visible = True
            .Font = "微软雅黑"
            .AutoSize = "true"
            .Caption = frm.Caption
            .Left = 200
            .Top = 120
            .BackStyle = 0
            .ForeColor = CtColor
            .FontSize = 10
        End With
    End If
    
    '设置关闭按钮
    If CloseCtB = True Then
        Set frmclose = frm.Controls.Add("vb.label", "Closebtn")
        With frmclose
            .Visible = True
            .AutoSize = False
            .Top = 150
            .Left = frm.Width - 150 - 240
            .Width = 240
            .Height = 240
            .BackStyle = 0
            .ZOrder 0
        End With
        frm.Line (frm.Width - 150 - 240, 150)-(frm.Width - 150, 150 + 240), CtColor
        frm.Line (frm.Width - 150, 150)-(frm.Width - 150 - 240, 150 + 240), CtColor
    End If
    
    '设置最小化按钮
    If SmallCtB = True Then
        Dim Mo As Integer
        If BigCtB = True Then Mo = 150 + 240 Else Mo = 0
        If CloseCtB = False Then Mo = Mo - (150 + 240) Else Mo = Mo
        frm.Line (frm.Width - 150 - 240 - 150 - 240 - Mo, 150 + 230)-(frm.Width - 150 - 240 - 150 - Mo, 150 + 230), CtColor
        Set SmallCt = frm.Controls.Add("vb.label", "Smallbtn")
        With SmallCt
            .Visible = True
            .AutoSize = False
            .Top = 150
            .Left = frm.Width - 150 - 240 - 150 - 240 - Mo
            .Width = 240
            .Height = 240
            .BackStyle = 0
            .ZOrder 0
        End With
    End If
    
    '设置最大化按钮
    If BigCtB = True Then
        Dim MoB As Integer
        If CloseCtB = True Then MoB = 150 + 240 Else MoB = 0
        Set frmbigs = frm.Controls.Add("vb.shape", "Bigbtn")
        With frmbigs
            .BorderColor = CtColor
            .Visible = True
            .Height = 255
            .Width = 255
            .Top = 150
            .Left = frm.Width - 150 - 240 - MoB
            .BorderStyle = 1
            .BackColor = 0
            .FillStyle = 1
            .Shape = 0
            .BorderWidth = 1
        End With
        Set frmbig = frm.Controls.Add("vb.label", "Bigbtns")
        With frmbig
            .Visible = True
            .AutoSize = False
            .Top = 150
            .Left = frm.Width - 150 - 240 - MoB
            .Width = 255
            .Height = 255
            .BackStyle = 0
            .ZOrder 0
        End With
        
    End If
    
    Call PrintLines
    
    Dim rtn As Long
    rtn = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong frm.hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes frm.hwnd, 0, Alpha, LWA_ALPHA
End Function
Public Sub RePrint()
    '重绘过程
    'FrmObj.Cls
    PrintLines
    'CloseButton
    If CloseB = True Then
        FrmObj.Line (FrmObj.Width - 150 - 240, 150)-(FrmObj.Width - 150, 150 + 240), wCtColor
        FrmObj.Line (FrmObj.Width - 150, 150)-(FrmObj.Width - 150 - 240, 150 + 240), wCtColor
        frmclose.Left = FrmObj.Width - 150 - 240
    End If
    'BigButton
    If BigB = True Then
        Dim MoB As Integer
        If CloseB = True Then MoB = 150 + 240 Else MoB = 0
        frmbigs.Left = FrmObj.Width - 150 - 240 - MoB
        frmbig.Left = FrmObj.Width - 150 - 240 - MoB
    End If
    'SmallButton
    If SmallB = True Then
        Dim Mo As Integer
        If BigB = True Then Mo = 150 + 240 Else Mo = 0
        If CloseB = False Then Mo = Mo - (150 + 240) Else Mo = Mo
        FrmObj.Line (FrmObj.Width - 150 - 240 - 150 - 240 - Mo, 150 + 230)-(FrmObj.Width - 150 - 240 - 150 - Mo, 150 + 230), wCtColor
        SmallCt.Left = FrmObj.Width - 150 - 240 - 150 - 240 - Mo
    End If
End Sub
Private Sub PrintLines()
    FrmObj.Line (0, 0)-(0, FrmObj.Height), QLineC
    FrmObj.Line (0, 0)-(FrmObj.Width, 0), QLineC
    FrmObj.Line (FrmObj.Width - 10, 0)-(FrmObj.Width - 10, FrmObj.Height - 10), QLineC
    FrmObj.Line (0, FrmObj.Height - 10)-(FrmObj.Width, FrmObj.Height - 10), QLineC
End Sub
Private Sub frmclose_Click()
    If Frmdis Then
        Frmdiso.Enabled = True
    Else
        Unload FrmObj
    End If
End Sub
Private Sub Smallct_Click()
    FrmObj.WindowState = 1
End Sub
Private Sub frmbig_Click()
    If FrmObj.WindowState <> 2 Then
        FrmObj.WindowState = 2
    Else
        FrmObj.WindowState = 0
    End If
    RePrint
End Sub
Private Sub Frmad_Resize()
    'FrmObj.Cls
    'Call PrintLines
End Sub
Private Sub Frmdisi_Timer()
    i = i + 5
    SetWindowLong FrmObj.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes FrmObj.hwnd, 0, i, LWA_ALPHA
    If i >= Frmalpha Then Frmdisi.Enabled = False
End Sub
Private Sub Frmdiso_Timer()
    i = i - 5
    SetWindowLong FrmObj.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes FrmObj.hwnd, 0, i, LWA_ALPHA
    If i <= 0 Then
        Frmdiso.Enabled = False
        Unload FrmObj
    End If
End Sub

Private Sub Frmad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If sMove Then
        If FrmObj.WindowState <> 2 Then
            ReleaseCapture
            SendMessage FrmObj.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
        End If
    End If
End Sub
