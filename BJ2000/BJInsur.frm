VERSION 5.00
Begin VB.Form BJInsur 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Insurance"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   Icon            =   "BJInsur.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.CommandButton CmdNo 
      Caption         =   "&No"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton CmdYes 
      Caption         =   "&Yes"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblInsurance 
      BackStyle       =   0  '³z©ú
      Caption         =   "Insurance?"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "BJInsur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public InsYes
Public InsNo

Private Sub CmdYes_Click()
    
    If BJMain.lblmoney.Caption < BJMain.lblBet.Caption / 2 Then
        MsgBox "You have not enough money!!", vbOKOnly, "Black Jack 2000 Alert"
    Else
        BJMain.lblmoney = BJMain.lblmoney - BJMain.lblBet / 2
        BJMain.CmdHit.Enabled = True
        BJMain.CmdStand.Enabled = True
    End If
    BJMain.CmdHit.Enabled = True
    BJMain.CmdStand.Enabled = True
    InsYes = 1
    Unload Me

End Sub

Private Sub CmdNo_Click()

    InsYes = 0
    BJMain.CmdHit.Enabled = True
    BJMain.CmdStand.Enabled = True
    Unload Me
    
End Sub

Private Sub Form_Load()

    IntRet = sndPlaySound(App.Path & "/insurance.wav", &H1)
    
End Sub
