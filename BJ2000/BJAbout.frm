VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form BJAbout 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "BJAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _cx             =   4196447
      _cy             =   4196447
      FlashVars       =   ""
      Movie           =   "\bj2000.swf"
      Src             =   "\bj2000.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.kyolinux.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   3240
      MouseIcon       =   "BJAbout.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "martin@kyolinux.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1365
      MouseIcon       =   "BJAbout.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3540
      Width           =   1455
   End
   Begin VB.Label lblIntro 
      BackColor       =   &H0000FFFF&
      Caption         =   $"BJAbout.frx":091E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Image imgCourse 
      Height          =   240
      Left            =   0
      Picture         =   "BJAbout.frx":0A2A
      Top             =   1920
      Width           =   6015
   End
   Begin VB.Image imgIVE 
      Height          =   1305
      Left            =   1800
      Picture         =   "BJAbout.frx":11CE
      Top             =   120
      Width           =   4140
   End
   Begin VB.Label lblDept 
      BackStyle       =   0  'Transparent
      Caption         =   "The Department of Electronics Engineering"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   5535
   End
End
Attribute VB_Name = "BJAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    Flash1.Movie = App.Path & "\bj2000.swf"
    
End Sub

Private Sub Label2_Click()

    xreturn = Shell("start.exe http://www.kyolinux.com", 0)

End Sub

Private Sub lblEmail_Click()

    xreturn = Shell("start.exe mailto:martin@kyolinux.com", 0)

End Sub

Private Sub txtSelfDetails_Change()

End Sub

