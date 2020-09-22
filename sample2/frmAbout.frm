VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   2925
   ClientLeft      =   4245
   ClientTop       =   3285
   ClientWidth     =   6090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2018.887
   ScaleMode       =   0  'User
   ScaleWidth      =   5718.826
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4680
      TabIndex        =   0
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Line Line1 
      X1              =   901.49
      X2              =   5521.625
      Y1              =   1159.566
      Y2              =   1159.566
   End
   Begin VB.Label lblComments 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   960
      TabIndex        =   9
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label lblInformation 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label lblDistribution 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label lblMailto 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      MousePointer    =   10  'Up Arrow
      TabIndex        =   6
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label lblWebAddress 
      AutoSize        =   -1  'True
      Caption         =   "www.easicomm.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      MousePointer    =   10  'Up Arrow
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblRevision 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   195
      Left            =   3900
      TabIndex        =   4
      Top             =   782
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label lblCopyRight 
      AutoSize        =   -1  'True
      Caption         =   "C"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   105
   End
   Begin VB.Line Line2 
      X1              =   901.49
      X2              =   5521.625
      Y1              =   496.957
      Y2              =   496.957
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   2
      Top             =   780
      Width           =   825
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Monitior"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3105
   End
   Begin VB.Image imgAbout 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TEXT_COPYRIGHT As String = "Copyright © 2000, Sean Calvert."
Private Const TEXT_DISTRIBUTE As String = "Freely distributed. All rights reserved."
Private Const TEXT_INFO As String = "For further information:"
Private Const TEXT_WWW As String = "http://www.seanie.co.uk"
Private Const TEXT_COMMENTS As String = "Send questions and comments to:"
Private Const TEXT_SMTP As String = "secalvert@seanie.co.uk"

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOK_Click()

  Unload Me
  
End Sub

Private Sub Form_GotFocus()

    lblWebAddress.ForeColor = vbBlue
    lblMailto.ForeColor = vbBlue

End Sub

Private Sub Form_Load()

    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor
    If App.Revision <> 0 Then
        lblRevision.Visible = True
        lblRevision.Caption = "Revision " & App.Revision
'        lblRevision.Left = lblVersion.Left + lblVersion.Width
    End If
    lblTitle.Caption = App.Title
    lblDistribution.Caption = TEXT_DISTRIBUTE
    lblInformation.Caption = TEXT_INFO
    Me!lblCopyRight.Caption = TEXT_COPYRIGHT
    lblComments.Caption = TEXT_COMMENTS
    Me.lblWebAddress.Caption = TEXT_WWW
    Me.lblMailto.Caption = TEXT_SMTP
      
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    lblWebAddress.ForeColor = vbBlue
    lblMailto.ForeColor = vbBlue

End Sub

Private Sub lblMailto_Click()

    lblMailto.ForeColor = vbBlue
    ShellExecute 0, "Open", "mailto:" & TEXT_SMTP, "", "", vbNormalFocus

End Sub

Private Sub lblMailto_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    lblMailto.ForeColor = vbRed
    lblWebAddress.ForeColor = vbBlue

End Sub

Private Sub lblWebAddress_Click()

    lblWebAddress.ForeColor = vbBlue
    ShellExecute 0, "Open", TEXT_WWW, "", "", vbNormalFocus

End Sub

Private Sub lblWebAddress_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    lblMailto.ForeColor = vbBlue
    lblWebAddress.ForeColor = vbRed

End Sub
