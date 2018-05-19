VERSION 5.00
Begin VB.Form HelloFrm 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "欢迎回来"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   960
      Top             =   840
   End
End
Attribute VB_Name = "HelloFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FEUI As New FEUI
Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\KaroLauncher\hello.jpg.jyp")
    FEUI.FEUI Me, False, False, False, True, True, True, True, , , , 240
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Mainfrm.Show
End Sub

Private Sub Timer1_Timer()
    Unload Me
    
End Sub
