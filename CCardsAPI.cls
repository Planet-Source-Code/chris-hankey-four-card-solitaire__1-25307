VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCardsAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum DrawModeEnum
    mdFaceUp = 0           '* Draw card face up, card to draw specified by cd *'
    mdFaceDown = 1         '* Draw card face down, back specified by cd (cdFaceDownFirst..cdFaceDownLast) *'
    mdHilite = 2           '* Same as FaceUp except drawn with NOTSRCCOPY mode *'
    mdGhost = 3            '* Draw a ghost card -- for ace piles *'
    mdRemove = 4           '* draw background specified by rgbBgnd *'
    mdInvisibleGhost = 5   '* ? *'
    mdDeckX = 6            '* Draw X *'
    mdDeckO = 7            '* Draw O *'
End Enum
Const conFaces = 0
Const conBacks = 1
Const conInvert = 2
Const conCrossHatch = 53   'This is the design for the discard pile marker
Const conPlaid = 54
Const conWeave = 55
Const conRobot = 56
Const conRoses = 57
Const conIvyBlack = 58
Const conIvyBlue = 59
Const conFishCyan = 60
Const conFishBlue = 61
Const conShell = 62
Const conCastle = 63
Const conBeach = 64
Const conCardHand = 65
Const conUnused = 66
Const conX = 67            'big red X
Const conO = 68            'big green O

Const suClub = 0
Const suDiamond = 1
Const suHeart = 2
Const suSpade = 3
Const suMax = 4
Const suFirst = suClub
Const raAce = 0
Const raTwo = 1
Const raThree = 2
Const raFour = 3
Const raFive = 4
Const raSix = 5
Const raSeven = 6
Const raEight = 7
Const raNine = 8
Const raTen = 9
Const raJack = 10
Const raQueen = 11
Const raKing = 12
Const raMax = 13
Const raNil = 15
Const raFirst = raAce

Public Enum CardsEnum
    cdAClubs = 0
    cd2Clubs = 4
    cd3Clubs = 8
    cd4Clubs = 12
    cd5Clubs = 16
    cd6Clubs = 20
    cd7Clubs = 24
    cd8Clubs = 28
    cd9Clubs = 32
    cdTClubs = 36
    cdJClubs = 40
    cdQClubs = 44
    cdKClubs = 48
    cdADiamonds = 1
    cd2Diamonds = 5
    cd3Diamonds = 9
    cd4Diamonds = 13
    cd5Diamonds = 17
    cd6Diamonds = 21
    cd7Diamonds = 25
    cd8Diamonds = 29
    cd9Diamonds = 33
    cdTDiamonds = 37
    cdJDiamonds = 41
    cdQDiamonds = 45
    cdKDiamonds = 49
    cdAHearts = 2
    cd2Hearts = 6
    cd3Hearts = 10
    cd4Hearts = 14
    cd5Hearts = 18
    cd6Hearts = 22
    cd7Hearts = 26
    cd8Hearts = 30
    cd9Hearts = 34
    cdTHearts = 38
    cdJHearts = 42
    cdQHearts = 46
    cdKHearts = 50
    cdASpades = 3
    cd2Spades = 7
    cd3Spades = 11
    cd4Spades = 15
    cd5Spades = 19
    cd6Spades = 23
    cd7Spades = 27
    cd8Spades = 31
    cd9Spades = 35
    cdTSpades = 39
    cdJSpades = 43
    cdQSpades = 47
    cdKSpades = 51
    cdFaceDown1 = 54
    cdFaceDown2 = 55
    cdFaceDown3 = 56
    cdFaceDown4 = 57
    cdFaceDown5 = 58
    cdFaceDown6 = 59
    cdFaceDown7 = 60
    cdFaceDown8 = 61
    cdFaceDown9 = 62
    cdFaceDown10 = 63
    cdFaceDown11 = 64
    cdFaceDown12 = 65
    cdFaceDownFirst = 54
    cdFaceDownLast = 65
End Enum

Private mlCardWidth As Long
Private mlCardHeight As Long

Private Declare Function cdtInit Lib "Cards32.Dll" (pdxCard As Long, pdyCard As Long) As Long

Private Declare Function cdtDraw Lib "Cards32.Dll" (hdc As Long, X As Long, Y As Long, _
cd As Long, md As Long, rgbBgnd As Long) As Long

Private Declare Function cdtDrawExt Lib "Cards32.Dll" (hdc As Long, X As Long, _
Y As Long, dX As Long, dY As Long, cd As Long, md As Long, rgbBgnd As Long) As Long

Private Declare Function cdtAnimate Lib "Cards32.Dll" (hdc As Long, cd As Long, X As Long, Y As Long, ispr As Long) As Long
Private Declare Function cdtTerm Lib "Cards32.Dll" () As Long

Private mudtRect As RECT
Public Function AnimateCard(hdc As Long, cd As CardsEnum, X As Long, Y As Long, _
ispr As Long) As Long

End Function

Public Function DrawCard(hdc As Long, X As Long, Y As Long, cd As CardsEnum, _
md As DrawModeEnum, rgbBand As Long) As CCard
    Dim objCard As CCard
    Dim lResult As Long
    
    On Error GoTo ErrorHandler
    
    Set objCard = New CCard
    
    SetRect mudtRect, X, Y, X + mlCardWidth, Y + mlCardHeight
    lResult = cdtDraw(ByVal hdc, ByVal X, ByVal Y, ByVal cd, ByVal md, ByVal rgbBand)
    If lResult = 0 Then
        Err.Raise 32000 & vbObjectError, , "can't draw card"
    End If
    objCard.SetCardData mudtRect.left, mudtRect.right, mudtRect.Top, mudtRect.bottom, _
    cd, md, rgbBand
      
    
    
    Set DrawCard = objCard
ExitSub:
    Exit Function
ErrorHandler:
    Set DrawCard = Nothing
End Function
Public Function Init(ByRef X As Long, ByRef Y As Long) As Long
    Init = cdtInit(mlCardWidth, mlCardHeight)
End Function
Public Function Terminate()

End Function

