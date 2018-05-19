VERSION 5.00
Begin VB.Form WaitFrm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Waitting"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   2760
      Top             =   2520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   1365
   End
End
Attribute VB_Name = "WaitFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FEUI As New FEUI
Public Sub WaitFrm(obj As Form, info As String, al As Integer, Optional times As Integer)
    Me.Width = obj.Width
    Me.Height = obj.Height
    Me.Top = obj.Top
    Me.Left = obj.Left
    Label1.Caption = info
    Label1.Top = obj.Height / 2 - Label1.Height / 2
    Label1.Left = obj.Width / 2 - Label1.Width / 2
    FEUI.FEUI Me, False, False, False, True, False, False, False, , , , al
    If times <> 0 Then Timer1.Interval = times
    
    Me.Show
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    'MsgBox "µÈ´ý³¬Ê±"
    Unload Me
    Unload Mainfrm
    End
    Timer1.Enabled = False
End Sub
