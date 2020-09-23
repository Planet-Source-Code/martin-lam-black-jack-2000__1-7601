VERSION 5.00
Object = "{6DE6E6DD-C656-11D2-B052-444553540000}#3.0#0"; "VBCARDS.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form BJMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  '¨S¦³®Ø½u
   Caption         =   "Black Jack 2000"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8385
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   6150
   ScaleWidth      =   8385
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.TextBox TxtPlay 
      Height          =   375
      Left            =   960
      TabIndex        =   32
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox TxtCom 
      Height          =   375
      Left            =   360
      TabIndex        =   31
      Top             =   6360
      Width           =   375
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   735
      Left            =   2520
      TabIndex        =   30
      Top             =   2640
      Width           =   3375
      _cx             =   4200257
      _cy             =   4195600
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "009A34"
      SWRemote        =   ""
   End
   Begin VB.PictureBox PicBack 
      Height          =   1455
      Left            =   120
      Picture         =   "Form1.frx":13A86
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   29
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDouble 
      Caption         =   "&Double"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7200
      TabIndex        =   19
      Top             =   5280
      Width           =   1095
   End
   Begin VB.PictureBox PicCom7 
      Height          =   1455
      Left            =   4680
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox PicCom6 
      Height          =   1455
      Left            =   4320
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   15
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox PicCom5 
      Height          =   1455
      Left            =   3960
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox PicCom4 
      Height          =   1455
      Left            =   3600
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox PicCom3 
      Height          =   1455
      Left            =   3240
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   12
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox PicCom2 
      Height          =   1455
      Left            =   2880
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton CmdStand 
      Caption         =   "&Stand"
      Height          =   615
      Left            =   5880
      TabIndex        =   10
      Top             =   5280
      Width           =   1095
   End
   Begin VB.PictureBox PicPlay7 
      Height          =   1455
      Left            =   4440
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox PicPlay6 
      Height          =   1455
      Left            =   4080
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdHit 
      Caption         =   "&Hit"
      Height          =   615
      Left            =   4560
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.PictureBox PicPlay5 
      Height          =   1455
      Left            =   3720
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox PicPlay4 
      Height          =   1455
      Left            =   3360
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox PicPlay3 
      Height          =   1455
      Left            =   3000
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "&New Game"
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin VB.PictureBox PicPlay2 
      Height          =   1455
      Left            =   2640
      Picture         =   "Form1.frx":18BC8
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox PicPlay1 
      Height          =   1455
      Left            =   2280
      Picture         =   "Form1.frx":1DD0A
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox PicCom1 
      Height          =   1455
      Left            =   2520
      Picture         =   "Form1.frx":22E4C
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VBCards.Deck Deck1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   1032
      Picture         =   "Form1.frx":27F8E
   End
   Begin VB.Label lblBet50 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BackStyle       =   0  '³z©ú
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   2040
      TabIndex        =   28
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblBet5 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BackStyle       =   0  '³z©ú
      ForeColor       =   &H80000008&
      Height          =   1150
      Left            =   1060
      TabIndex        =   27
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblBet10 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BackStyle       =   0  '³z©ú
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   90
      TabIndex        =   26
      Top             =   4770
      Width           =   975
   End
   Begin VB.Line Line8 
      X1              =   960
      X2              =   2040
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line7 
      X1              =   960
      X2              =   2040
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line6 
      X1              =   2040
      X2              =   2040
      Y1              =   4320
      Y2              =   4680
   End
   Begin VB.Line Line5 
      X1              =   960
      X2              =   960
      Y1              =   4320
      Y2              =   4680
   End
   Begin VB.Label Label4 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackColor       =   &H00FFFFFF&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   25
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label lblBet 
      BackColor       =   &H80000009&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   24
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblBettitle 
      BackColor       =   &H00FFFF80&
      Caption         =   "Bet"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape sp5 
      BorderStyle     =   0  '³z©ú
      Height          =   1150
      Left            =   1080
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblMoneytitle 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackColor       =   &H00FFFF80&
      Caption         =   "Your Money"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   22
      Top             =   240
      Width           =   1695
   End
   Begin VB.Line Line4 
      X1              =   8160
      X2              =   8160
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Line Line3 
      X1              =   6480
      X2              =   6480
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   8160
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   6480
      X2              =   8160
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackColor       =   &H00FFFFFF&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   21
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblmoney 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6840
      TabIndex        =   20
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblPlay 
      BackStyle       =   0  '³z©ú
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   615
      Left            =   6240
      TabIndex        =   18
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblCom 
      BackStyle       =   0  '³z©ú
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   615
      Left            =   1440
      TabIndex        =   17
      Top             =   1200
      Width           =   615
   End
   Begin VB.Menu File 
      Caption         =   "&File "
      Begin VB.Menu New_Game 
         Caption         =   "Re-Start &New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Sound1 
      Caption         =   "&Sound"
      Begin VB.Menu Sound 
         Caption         =   "&On"
         Checked         =   -1  'True
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu About 
      Caption         =   "&About"
      Begin VB.Menu Help 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu About1 
         Caption         =   "A&bout"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "BJMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, x, y, z, bj
Dim play1, com1, flag, comparewin, Number%, IntRet
Dim time As Integer
Dim Index As Integer
Dim val As Integer
Dim Start As Long
Public PlayBet
Public PlayMoney
Const Max = 52
Option Base 1
Dim num(52)
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub About1_Click()

    BJAbout.Visible = True
    
End Sub

Private Sub cmdDouble_Click()

    cdouble
    cmdDouble.Enabled = False
    
End Sub

Private Sub CmdNew_Click()

    If lblBet.Caption = 0 Then
        MsgBox "Please bet first!!", vbOKOnly, "Black Jack 2000 Alert"
    Else

        setNum
        a = 0
        b = 0
        c = 0
        d = 0
        e = 0
        f = 0
        g = 0
        h = 0
        i = 0
        j = 0
        k = 0
        l = 0
        m = 0
        n = 0
        k = 0
        l = 0
        m = 0
        n = 0
        o = 0
        p = 0
        q = 0
        r = 0
        s = 0
        cmdDouble.Enabled = False
        CmdStand.Enabled = True
        lblBet10.Enabled = False
        lblBet5.Enabled = False
        lblBet50.Enabled = False
        PicPlay3.Visible = False
        PicPlay4.Visible = False
        PicPlay5.Visible = False
        PicPlay6.Visible = False
        PicPlay7.Visible = False
        PicCom2.Visible = False
        PicCom3.Visible = False
        PicCom4.Visible = False
        PicCom5.Visible = False
        PicCom6.Visible = False
        PicCom7.Visible = False
        Flash.Visible = False
        Flash.Movie = App.Path & "/a.swf"
    
                While j < 3
                    openNumber
                    Index = Index + 1
                    j = j + 1
                Wend
    
            k = Int(Rnd * 52) + 1
            Deck1.ChangeCard = k
            PicCom1.Picture = Deck1.Picture
    
            If k > 39 Then
                k = k - 39
            End If
        
            If k > 26 Then
                k = k - 26
            End If
        
            If k > 13 Then
                k = k - 13
            End If
        
            If k > 10 Then
                k = 10
            End If
        
            lblCom.Caption = k
    
            a = Int(Rnd * 52) + 1
            Deck1.ChangeCard = a
            PicPlay1.Picture = Deck1.Picture
            b = Int(Rnd * 52) + 1
            Deck1.ChangeCard = b
            PicPlay2.Picture = Deck1.Picture
    
            If a > 39 Then
                a = a - 39
            End If
        
            If a > 26 Then
                a = a - 26
            End If
        
            If a > 13 Then
                a = a - 13
            End If
        
            If a > 10 Then
                a = 10
            End If
    
            If b > 39 Then
                b = b - 39
            End If
        
            If b > 26 Then
                b = b - 26
            End If
    
            If b > 13 Then
                b = b - 13
            End If
        
            If b > 10 Then
                b = 10
            End If
    
            CmdNew.Enabled = False
            CmdHit.Enabled = True
            
        lblPlay.Caption = a + b
    
        If a = 1 Or b = 1 Then
            z = 0
        Else
            z = 1
        End If
        
        If z = 0 Then
            lblPlay.Caption = lblPlay.Caption + 10
        End If
    
        If lblPlay.Caption > 21 And z = 1 Then
            lblPlay.Caption = lblPlay.Caption - 10
        End If
    
        If a = 1 And b = 10 Then
            s = 1
        Else
            s = 0
        End If
    
        If a = 10 And b = 1 Then
            t = 1
        Else
            t = 0
        End If
    
        If k = 1 Then
            Combj
        End If
        
        If t = 1 Or s = 1 Then
            blackjack
        End If
    
    End If
    
    If lblPlay.Caption = 11 Then
        cmdDouble.Enabled = True
    End If
    
End Sub

Private Sub bothbj()
    
    CmdHit.Enabled = False
    CmdStand.Enabled = False
    PicCom2.Visible = True
    PicCom2.Picture = PicBack.Picture
    CmdHit.Enabled = False
    CmdStand.Enabled = False
    BJInsur.Show 1
    stand
    
End Sub

Private Sub Combj()

    CmdHit.Enabled = False
    CmdStand.Enabled = False
    PicCom2.Visible = True
    PicCom2.Picture = PicBack.Picture
    CmdHit.Enabled = False
    CmdStand.Enabled = False
    BJInsur.Show 1

End Sub

Private Sub CmdHit_Click()

    Dim Card
        
        If PicPlay7.Visible = False Then
            Card = 1
        End If

        If PicPlay6.Visible = False Then
            Card = 2
        End If

        If PicPlay5.Visible = False Then
            Card = 3
        End If

        If PicPlay4.Visible = False Then
            Card = 4
        End If

        If PicPlay3.Visible = False Then
            Card = 5
        End If

Select Case Card

    Case 1
        If Sound.Checked = True Then
            IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
        End If
        
        PicPlay7.Visible = True
        h = Int(Rnd * 52 + 1)
        Deck1.ChangeCard = h
        PicPlay7.Picture = Deck1.Picture
        
            If h > 39 Then
                h = h - 39
            End If
            
            If h > 26 Then
                h = h - 26
            End If
            
            If h > 13 Then
                h = h - 13
            End If
            
            If h > 10 Then
                h = 10
            End If

    Case 2
        If Sound.Checked = True Then
            IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
        End If
        
        PicPlay6.Visible = True
        g = Int(Rnd * 52 + 1)
        Deck1.ChangeCard = g
        PicPlay6.Picture = Deck1.Picture
            
            If g > 39 Then
                g = g - 39
            End If
            
            If g > 26 Then
                g = g - 26
            End If
            
            If g > 13 Then
                g = g - 13
            End If
            
            If g > 10 Then
                g = 10
            End If

    Case 3
        If Sound.Checked = True Then
            IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
        End If
        
        PicPlay5.Visible = True
        f = Int(Rnd * 52 + 1)
        Deck1.ChangeCard = f
        PicPlay5.Picture = Deck1.Picture
        
            If f > 39 Then
                f = f - 39
            End If
        
            If f > 26 Then
                f = f - 26
            End If
        
            If f > 13 Then
                f = f - 13
            End If
            
            If f > 10 Then
                f = 10
            End If

    Case 4
        If Sound.Checked = True Then
            IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
        End If
        
        PicPlay4.Visible = True
        e = Int(Rnd * 52 + 1)
        Deck1.ChangeCard = e
        PicPlay4.Picture = Deck1.Picture
        
            If e > 39 Then
                e = e - 39
            End If
        
            If e > 26 Then
                e = e - 26
            End If
        
            If e > 13 Then
                e = e - 13
            End If
        
            If e > 10 Then
                e = 10
            End If

    Case 5
        If Sound.Checked = True Then
            IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
        End If
        
        PicPlay3.Visible = True
        c = Int(Rnd * 52 + 1)
        Deck1.ChangeCard = c
        PicPlay3.Picture = Deck1.Picture
        
            If c > 39 Then
                c = c - 39
            End If
        
            If c > 26 Then
                c = c - 26
            End If
        
            If c > 13 Then
                c = c - 13
            End If
        
            If c > 10 Then
                c = 10
            End If
        
End Select

    If a = 1 Or b = 1 Or c = 1 Or e = 1 Or f = 1 Or g = 1 Or h = 1 Then
        x = 0
    Else
        x = 1
    End If

    lblPlay.Caption = a + b + c + e + f + g + h
    
    If x = 0 Then
        lblPlay.Caption = lblPlay.Caption + 10
    End If
    
    If x = 0 And lblPlay.Caption > 21 Then
        lblPlay.Caption = lblPlay.Caption - 10
    End If

    If lblPlay.Caption > 21 Then
        loss
        CmdStand.Enabled = False
        CmdHit.Enabled = False
        CmdNew.Enabled = True
    End If
    
    u = lblPlay.Caption
    v = lblCom.Caption
    
End Sub

Private Sub cdouble()
    
    If lblmoney.Caption < lblBet.Caption / 1 Then
        MsgBox "You have not enough money!!", vbOKOnly, "Black Jack 2000 Alert"
    Else
        lblmoney.Caption = lblmoney.Caption - lblBet.Caption
        lblBet.Caption = lblBet.Caption * 2
        PicPlay3.Visible = True
        c = Int(52 * Rnd) + 1
        Deck1.ChangeCard = c
        PicPlay3.Picture = Deck1.Picture

            If c > 39 Then
                c = c - 39
            End If

            If c > 26 Then
                c = c - 26
            End If
                
            If c > 13 Then
                c = c - 13
            End If
            
            If c > 10 Then
                c = 10
            End If
            
        lblPlay.Caption = a + b + c
        stand
    End If
    
End Sub

Private Sub Combj1()

    Flash.Visible = True
    Flash.Movie = App.Path & "\loss.swf"
    lblmoney.Caption = lblmoney.Caption + lblBet.Caption * 1.5
    CmdNew.Enabled = True
    
End Sub

Private Sub draw()

    If Sound.Checked = True Then
        IntRet = sndPlaySound(App.Path & "/snore.wav", &H1)
    End If
    
    Flash.Visible = True
    Flash.Movie = App.Path & "\draw.swf"
    CmdHit.Enabled = False
    CmdStand.Enabled = False
    CmdNew.Enabled = True
    lblmoney.Caption = lblmoney.Caption + lblBet.Caption * 1
    CmdNew.Enabled = True
    
End Sub

Private Sub loss()

    If Sound.Checked = True Then
        IntRet = sndPlaySound(App.Path & "/hung-01.wav", &H1)
    End If
    
    Flash.Visible = True
    Flash.Movie = App.Path & "\loss.swf"
    lblBet.Caption = 0
    lblBet5.Enabled = True
    lblBet10.Enabled = True
    lblBet50.Enabled = True
    CmdNew.Enabled = True
    
End Sub

Private Sub win()
    
    If Sound.Checked = True Then
        If lblCom.Caption > 21 Then
            IntRet = sndPlaySound(App.Path & "/Explode.wav", &H1)
        Else
            IntRet = sndPlaySound(App.Path & "/haha.wav", &H1)
        End If
    End If
    
    Flash.Visible = True
    Flash.Movie = App.Path & "\win.swf"
    lblmoney.Caption = lblBet.Caption * 2 + lblmoney.Caption
    lblBet.Caption = 0
    lblBet5.Enabled = True
    lblBet10.Enabled = True
    lblBet50.Enabled = True
    CmdNew.Enabled = True
    
End Sub

Private Sub blackjack()

    If Sound.Checked = True Then
        IntRet = sndPlaySound(App.Path & "/clap.wav", &H1)
    End If
    
    Flash.Visible = True
    Flash.Movie = App.Path & "\bj.swf"
    lblmoney.Caption = lblBet.Caption * 2.5 + lblmoney.Caption
    lblBet.Caption = 0
    lblBet5.Enabled = True
    lblBet10.Enabled = True
    lblBet50.Enabled = True
    CmdHit.Enabled = False
    CmdStand.Enabled = False
    CmdNew.Enabled = True
        
End Sub

Private Sub CmdStand_Click()

    stand

End Sub

Private Sub stand()

    cmdDouble.Enabled = False
    CmdHit.Enabled = False
    CmdNew.Enabled = False
    CmdStand.Enabled = False
    
    While lblCom.Caption < 17
        Dim ComCard
            If PicCom7.Visible = False Then
                ComCard = 1
            End If

            If PicCom6.Visible = False Then
                ComCard = 2
            End If

            If PicCom5.Visible = False Then
                ComCard = 3
            End If

            If PicCom4.Visible = False Then
                ComCard = 4
            End If

            If PicCom3.Visible = False Then
                ComCard = 5
            End If

            If PicCom2.Picture = PicBack.Picture Then
                ComCard = 6
            End If

            If PicCom2.Visible = False Then
                ComCard = 7
            End If

    Select Case ComCard
        
        Case 1

            Start = Timer
            Do While Timer < Start + time
                DoEvents
            Loop
  
            If Sound.Checked = True Then
                IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
            End If
            
            PicCom7.Visible = True
            l = Int(Rnd * 52 + 1)
            Deck1.ChangeCard = l
            PicCom7.Picture = Deck1.Picture
            
            If l > 39 Then
                l = l - 39
            End If
                
            If l > 26 Then
                l = l - 26
            End If

            If l > 13 Then
                l = l - 13
            End If

            If l > 10 Then
                l = 10
            End If

        Case 2

            Start = Timer
            Do While Timer < Start + time
                DoEvents
            Loop
                
            If Sound.Checked = True Then
                IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
            End If
            
            PicCom6.Visible = True
            m = Int(Rnd * 52 + 1)
            Deck1.ChangeCard = m
            PicCom6.Picture = Deck1.Picture
            
            If m > 39 Then
                m = m - 39
            End If

            If m > 26 Then
                m = m - 26
            End If
            
            If m > 13 Then
                m = m - 13
            End If

            If m > 10 Then
                m = 10
            End If

        Case 3

            Start = Timer
            Do While Timer < Start + time
                DoEvents
            Loop
            
            If Sound.Checked = True Then
                IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
            End If
            
            PicCom5.Visible = True
            n = Int(Rnd * 52 + 1)
            Deck1.ChangeCard = n
            PicCom5.Picture = Deck1.Picture
            
            If n > 39 Then
                n = n - 39
            End If
            
            If n > 26 Then
                n = n - 26
            End If
            
            If n > 13 Then
                n = n - 13
            End If
            
            If n > 10 Then
                n = 10
            End If

        Case 4

            Start = Timer
            Do While Timer < Start + time
                DoEvents
            Loop
  
            If Sound.Checked = True Then
                IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
            End If
            
            PicCom4.Visible = True
            o = Int(Rnd * 52 + 1)
            Deck1.ChangeCard = o
            PicCom4.Picture = Deck1.Picture
            
            If o > 39 Then
                o = o - 39
            End If
    
            If o > 26 Then
                o = o - 26
            End If
    
            If o > 13 Then
                o = o - 13
            End If
    
            If o > 10 Then
                o = 10
            End If

        Case 5

            Start = Timer
            Do While Timer < Start + time
                DoEvents
            Loop
  
            If Sound.Checked = True Then
                IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
            End If
            
            PicCom3.Visible = True
            p = Int(Rnd * 52 + 1)
            Deck1.ChangeCard = p
            PicCom3.Picture = Deck1.Picture
            
            If p > 39 Then
                p = p - 39
            End If
            
            If p > 26 Then
                p = p - 26
            End If

            If p > 13 Then
                p = p - 13
            End If

            If p > 10 Then
                p = 10
            End If

        Case 6

            Start = Timer
            Do While Timer < Start + time
                DoEvents
            Loop
    
            If Sound.Checked = True Then
                IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
            End If
            
            PicCom2.Visible = True
            r = Int(Rnd * 52 + 1)
            Deck1.ChangeCard = r
            PicCom2.Picture = Deck1.Picture
    
            If r > 39 Then
                r = r - 39
            End If
            
            If r > 26 Then
                r = r - 26
            End If
            
            If r > 13 Then
                r = r - 13
            End If
            
            If r > 10 Then
                r = 10
            End If

        Case 7

            Start = Timer
            Do While Timer < Start + time
                DoEvents
            Loop
            
            If Sound.Checked = True Then
                IntRet = sndPlaySound(App.Path & "/whoosh.wav", &H1)
            End If
            
            PicCom2.Visible = True
            r = Int(Rnd * 52 + 1)
            Deck1.ChangeCard = r
            PicCom2.Picture = Deck1.Picture
            
            If r > 39 Then
                r = r - 39
            End If

            If r > 26 Then
                r = r - 26
            End If

            If r > 13 Then
                r = r - 13
            End If

            If r > 10 Then
                r = 10
            End If

    End Select

        If k = 1 Or l = 1 Or m = 1 Or n = 1 Or o = 1 Or p = 1 Or q = 1 Or r = 1 Then
            w = 0
        Else
            w = 1
        End If

        lblCom.Caption = k + l + m + n + o + p + q + r

        If w = 0 Then
            lblCom.Caption = lblCom.Caption + 10
        End If
    
        If w = 0 And lblCom.Caption > 21 Then
            lblCom.Caption = lblCom.Caption - 10
        End If
    
    Wend

    compare

End Sub

Private Sub compare()

    If k = 1 And r = 10 And BJInsur.InsYes = 1 Then
        Combj1
    End If

    If lblCom.Caption = lblPlay.Caption Then
        comparewin = 2
    End If

    If lblCom.Caption * 1 > lblPlay.Caption * 1 Then
        comparewin = 3
    End If

    If lblCom.Caption * 1 > lblPlay.Caption * 1 And lblCom.Caption * 1 > 21 Then
        comparewin = 4
    End If
    
    If lblCom.Caption * 1 < lblPlay.Caption * 1 Then
        comparewin = 4
    End If

    Select Case comparewin
    
        Case 1
            Combj1
        
        Case 2
            draw
        
        Case 3
            loss
        
        Case 4
            win
    
    End Select

    lblBet.Caption = 0
    lblBet5.Enabled = True
    lblBet10.Enabled = True
    lblBet50.Enabled = True

End Sub

Private Sub Command4_Click()
    
    Unload Me

End Sub

Private Sub Exit_Click()

    Unload BJInsur
    Unload Me

End Sub

Private Sub Form_Load()
            
    Sound.Checked = True
    time = 0.51
    Index = 0
    setNum
    CmdStand.Enabled = False
    CmdHit.Enabled = False
    PicCom2.Visible = False
    PicCom3.Visible = False
    PicCom4.Visible = False
    PicCom5.Visible = False
    PicCom6.Visible = False
    PicCom7.Visible = False
    BJInsur.InsYes = 0

End Sub

Private Sub openNumber()
    
    Randomize
    flag = True
    
    While flag
        Number% = Int(52 * Rnd) + 1
        
        If num(Number%) = True Then
            num(Number%) = False
            flag = False
        End If
    Wend
   
End Sub
Private Sub setNum()
    
    For i = 1 To Max
        num(i) = True
    Next i
    
    Index = 0
    j = 0
    
End Sub

Private Sub Help_Click()

    BJHelp.Visible = True

End Sub

Private Sub lblBet10_Click()

    If lblmoney.Caption <= 9 Then
        MsgBox "You have not enough money!!", vbOKOnly, "Black Jack 2000 Alert"
    Else
        lblBet.Caption = lblBet.Caption + 10
        lblmoney.Caption = lblmoney.Caption - 10
    End If
    
End Sub

Private Sub lblBet5_Click()

    If lblmoney.Caption <= 4 Then
        MsgBox "You have not enough money!!", vbOKOnly, "Black Jack 2000 Alert"
    Else
        lblBet.Caption = lblBet.Caption + 5
        lblmoney.Caption = lblmoney.Caption - 5
    End If
    
End Sub

Private Sub lblBet50_Click()

    If lblmoney.Caption <= 49 Then
        MsgBox "You have not enough money!!", vbOKOnly, "Black Jack 2000 Alert"
    Else
        lblBet.Caption = lblBet.Caption + 50
        lblmoney.Caption = lblmoney.Caption - 50
    End If
    
End Sub

Private Sub New_Game_Click()
    lblmoney.Caption = 500
    lblBet.Caption = 0
    lblPlay.Caption = 0
    lblCom.Caption = 0
    PicCom1.Visible = True
    PicCom1.Picture = PicBack.Picture
    PicCom2.Visible = False
    PicCom3.Visible = False
    PicCom4.Visible = False
    PicCom5.Visible = False
    PicCom6.Visible = False
    PicCom7.Visible = False
    PicPlay1.Visible = True
    PicPlay1.Picture = PicBack.Picture
    PicPlay2.Visible = True
    PicPlay2.Picture = PicBack.Picture
    PicPlay3.Visible = False
    PicPlay4.Visible = False
    PicPlay5.Visible = False
    PicPlay6.Visible = False
    PicPlay7.Visible = False
    CmdNew.Enabled = True
    CmdHit.Enabled = False
    CmdStand.Enabled = False
    cmdDouble.Enabled = False
    lblBet5.Enabled = True
    lblBet10.Enabled = True
    lblBet50.Enabled = True
    Flash.Visible = False

End Sub


Private Sub Sound_Click()
    
    If Sound.Checked = True Then
        Sound.Checked = False
    Else
        Sound.Checked = True
    End If
    
End Sub
