VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Sir Tommy Patience"
   ClientHeight    =   5610
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   8865
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   591
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGNew 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFStats 
         Caption         =   "&Statistics"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHTopics 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu mnuHSep01 
         Caption         =   "-"
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

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim lngAnswer As Long
    Dim strPrompt As String

    gclsGame.DoMove x, y
    gclsGame.UpdateScreen Me.hDC, Me.ScaleWidth, Me.ScaleHeight
    Me.Refresh
    
    If Not gclsGame.AnyMovesLeft Then
        If gclsGame.CardsLeft = 0 Then
            strPrompt = "Congratulations! You have won!"
        Else
            strPrompt = "No more moves left."
        End If
            
        strPrompt = strPrompt & vbNewLine & vbNewLine & vbNewLine & _
                    "Do you wish to play again?"
                    
        gclsStats.AddGameStat gclsGame.CardsLeft
        
        If MsgBox(strPrompt, vbExclamation + vbYesNo, Me.Caption) = vbYes Then
            mnuGNew_Click
        Else
            EndApp
        End If
        
    End If
    
End Sub

Private Sub Form_Resize()

    gclsGame.UpdateScreen Me.hDC, Me.ScaleWidth, Me.ScaleHeight
    Me.Refresh

End Sub

Private Sub Form_Terminate()

    EndApp

End Sub

Private Sub mnuFStats_Click()

    Load frmStats
    frmStats.Show 1, Me

End Sub

Private Sub mnuGExit_Click()

    EndApp

End Sub

Private Sub mnuGNew_Click()

    gclsGame.NewGame
    gclsGame.UpdateScreen Me.hDC, Me.ScaleWidth, Me.ScaleHeight
    Me.Refresh

End Sub

Private Sub mnuHAbout_Click()

    Load frmAbout
    frmAbout.Show 1, Me

End Sub

Private Sub mnuHTopics_Click()

    DisplayHelp

End Sub
