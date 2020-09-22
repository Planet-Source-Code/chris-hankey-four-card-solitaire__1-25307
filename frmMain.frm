VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   Caption         =   "4-Card Solitaire"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":0000
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   755
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6165
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "DealsLeft"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17223
            MinWidth        =   17223
            Key             =   "Messages"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGStart 
         Caption         =   "&Start"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuGDeal 
         Caption         =   "&Deal"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbMoveCard As Boolean
Private WithEvents mobjGame As CGame
Attribute mobjGame.VB_VarHelpID = -1

Private Sub Deal()

    If Not mobjGame Is Nothing Then
        mobjGame.Deal
        StatusBar1.Panels("DealsLeft").Text = "Deals Left: " & mobjGame.DealsLeft
    End If
End Sub

Private Sub Start()
    Dim iInitOK As Long
    Dim iDrawOK As Long
    
    Me.Cls
    
    Set mobjGame = New CGame
    mobjGame.Start Me
    
    StatusBar1.Panels("DealsLeft").Text = "Deals Left: " & mobjGame.DealsLeft

    
    
    Me.Refresh
End Sub



Private Sub Form_Load()
    Call ShowDefaultMessage
    Call Start
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        
        If Not mobjGame Is Nothing Then
            mbMoveCard = mobjGame.WasCardGrabbed(CLng(X), CLng(Y))
            If mbMoveCard = True Then
                MousePointer = vbCustom
            End If
        End If
    ElseIf vbRightButton Then
        Deal
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mbMoveCard = True Then
        mobjGame.MoveCard CLng(X), CLng(Y)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If mbMoveCard = True Then
        MousePointer = 0
        mobjGame.CardDropped X, Y
        mbMoveCard = False
    End If
    
    ShowDefaultMessage
End Sub

Private Sub mnuGDeal_Click()
    Call Deal
End Sub
Private Sub ShowDefaultMessage()
    StatusBar1.Panels("Messages").Text = "Start Game Ctl-S, Deal Ctl-D or Right Click"
End Sub

Private Sub mnuGStart_Click()
    Call Start
End Sub

Private Sub mnuHAbout_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mobjGame_GameOver()
    Dim iPlayAgain As Integer
    Dim sTemp As String
    
    Me.Refresh
    DoEvents
    
    If mobjGame.GameHasBeenWon = True Then
        sTemp = "YOU ARE A WINNER!  Do you want to play again?"
    Else
        sTemp = "No valid moves remain--YOU ARE A LOSER!  Do you want to play again?"
    End If
        
    
    iPlayAgain = MsgBox(sTemp, vbYesNo)
    
    If iPlayAgain = vbYes Then
        Start
        Deal
    End If
    
End Sub

Private Sub mobjGame_Msg(sMsg As Variant)
    StatusBar1.Panels("Messages").Text = sMsg
End Sub

Private Sub mnuHContents_Click()
HtmlHelp Me.hWnd, App.Path & "\" & "cards.chm", HH_DISPLAY_TOC, Null
End Sub
