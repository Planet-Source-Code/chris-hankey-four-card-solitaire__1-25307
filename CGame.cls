VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'this type is used to mark the start of the drag
Private Type DragPosType
    X As Long
    Y As Long
End Type

Private mobjCardsAPI As CCardsAPI
Private miDeals As Integer
Private mobjDeck As CDeck
Private mobjDiscard As CCards
Private mobjPile1 As CCards
Private mobjPile2 As CCards
Private mobjPile3 As CCards
Private mobjPile4 As CCards
Private mobjForm As Form

Private mobjDragSrcPile As CCards
Private mobjDragCard As CCard
Private mlCardWidth As Long
Private mlCardHeight As Long

Private mudtDragPos As DragPosType

'flags
Private mbPileSelected As Boolean
Private mbCardBeingDragged As Boolean

'events
Public Event Msg(sMsg)
Public Event GameOver()

'constants
Private Const ALL_CARDS_DEALT = 52
Private Const TOTAL_DEALS = 13

Private Sub AddDragCard()
    Dim objCard As CCard
    
    'adds the drag card to the pile, or simply removes it if the user
    'is trying to discard it.
    
    If mobjDiscard.Selected = True Then
        'erase the drag card
        Set objCard = mobjDiscard(1)
        
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, objCard.left, _
        objCard.Top, objCard.Card, mdFaceUp, objCard.Color)
        mobjDiscard.Selected = False
        
        mobjDiscard.Remove 1
        mobjDiscard.Add objCard
        
        Set mobjDragCard = Nothing
        GoTo ExitSub
    End If
    
    If mobjPile1.Selected = True Then
        Set objCard = mobjPile1(1)
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, objCard.left, _
        objCard.Top, objCard.Card, mdFaceDown, mobjForm.BackColor)
        mobjPile1.Selected = False
        
        mobjPile1.Remove 1
        mobjPile1.Add objCard
        
        
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, mobjPile1.GetNextCardLeft, _
        mobjPile1.GetNextCardTop, mobjDragCard.Card, mdFaceUp, mobjDragCard.Color)
        mobjPile1.Selected = False

        
        mobjPile1.Add objCard
    End If
    If mobjPile2.Selected = True Then
        Set objCard = mobjPile2(1)
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, objCard.left, _
        objCard.Top, objCard.Card, mdFaceDown, mobjForm.BackColor)
        mobjPile2.Selected = False
        
        mobjPile2.Remove 1
        mobjPile2.Add objCard
        
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, mobjPile2.GetNextCardLeft, _
        mobjPile2.GetNextCardTop, mobjDragCard.Card, mdFaceUp, mobjDragCard.Color)
        mobjPile2.Selected = False
        
        mobjPile2.Add objCard
    End If
    If mobjPile3.Selected = True Then
    
        Set objCard = mobjPile3(1)
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, objCard.left, _
        objCard.Top, objCard.Card, mdFaceDown, mobjForm.BackColor)
        mobjPile3.Selected = False
        
        mobjPile3.Remove 1
        mobjPile3.Add objCard
    
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, mobjPile3.GetNextCardLeft, _
        mobjPile3.GetNextCardTop, mobjDragCard.Card, mdFaceUp, mobjDragCard.Color)
        mobjPile3.Add objCard
    End If
    If mobjPile4.Selected = True Then
        Set objCard = mobjPile4(1)
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, objCard.left, _
        objCard.Top, objCard.Card, mdFaceDown, mobjForm.BackColor)
        mobjPile4.Selected = False
        
        mobjPile4.Remove 1
        mobjPile4.Add objCard
    
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, mobjPile4.GetNextCardLeft, _
        mobjPile4.GetNextCardTop, mobjDragCard.Card, mdFaceUp, mobjDragCard.Color)
        mobjPile4.Add objCard
    End If

    
ExitSub:
    Exit Sub
End Sub

Private Sub AddBackDragCard()
    Dim objCard As CCard
    
    'adds the drag card back to its original pile
    If mbCardBeingDragged = True Then
        mobjForm.Cls
        Me.RedrawBoard
        
        'erase the drag card
        mobjCardsAPI.DrawCard mobjForm.hdc, mobjDragCard.left, mobjDragCard.Top, _
        mobjDragCard.Card, mdRemove, mobjForm.BackColor
        
        
        'the user did not drop the card on a pile, so
        'add the card being dragged back to its original pile,
        mobjDragSrcPile.Add mobjDragCard
        
        're-draw the pile w the drag card
        For Each objCard In mobjDragSrcPile
            mobjCardsAPI.DrawCard mobjForm.hdc, objCard.left, objCard.Top, _
            objCard.Card, objCard.DrawMode, objCard.Color
        Next objCard
    
        mobjForm.Refresh
    End If
End Sub
Public Sub RedrawBoard()
    Dim objCard As CCard

    mobjForm.Cls
    'redraw all of the cards that were showing
    For Each objCard In mobjPile1
        With objCard
            mobjCardsAPI.DrawCard mobjForm.hdc, .left, .Top, .Card, .DrawMode, .Color
        End With
    Next objCard
    For Each objCard In mobjPile2
        With objCard
            mobjCardsAPI.DrawCard mobjForm.hdc, .left, .Top, .Card, .DrawMode, .Color
        End With
    Next objCard
    For Each objCard In mobjPile3
        With objCard
            mobjCardsAPI.DrawCard mobjForm.hdc, .left, .Top, .Card, .DrawMode, .Color
        End With
    Next objCard
    For Each objCard In mobjPile4
        With objCard
            mobjCardsAPI.DrawCard mobjForm.hdc, .left, .Top, .Card, .DrawMode, .Color
        End With
    Next objCard
    
    For Each objCard In mobjDiscard
        With objCard
            mobjCardsAPI.DrawCard mobjForm.hdc, .left, .Top, .Card, .DrawMode, .Color
        End With
    Next objCard

End Sub
Private Function RemoveDragCard(ByRef robjPile As CCards)
    Dim objCard As CCard
    
    'removes the card being dragged from the pile
    robjPile.Remove robjPile.Count
    
    're-draw the pile w/o the drag card
    For Each objCard In robjPile
        mobjCardsAPI.DrawCard mobjForm.hdc, objCard.left, objCard.Top, _
        objCard.Card, objCard.DrawMode, objCard.Color
    Next objCard
    Set mobjDragSrcPile = robjPile
    
End Function

Public Function Start(objForm As Form) As String
    Dim objCard As CCard
    Dim lResult As Long
    On Error GoTo ErrorHandler
    
    mobjCardsAPI.Init mlCardHeight, mlCardWidth
    
    
    'initialize the game variables/objects
    Set mobjForm = objForm
    miDeals = 0
    mobjDeck.NewDeck
    mobjDeck.Shuffle
    
    'set a ghost card for each of the piles
    Set objCard = mobjCardsAPI.DrawCard(objForm.hdc, 250, 50, 53, mdFaceDown, objForm.BackColor)
    objCard.Moveable = False
    mobjPile1.Add objCard

    
    Set objCard = mobjCardsAPI.DrawCard(objForm.hdc, 375, 50, 53, mdFaceDown, objForm.BackColor)
    objCard.Moveable = False
    mobjPile2.Add objCard
    
    Set objCard = mobjCardsAPI.DrawCard(objForm.hdc, 500, 50, 53, mdFaceDown, objForm.BackColor)
    objCard.Moveable = False
    mobjPile3.Add objCard
    
    Set objCard = mobjCardsAPI.DrawCard(objForm.hdc, 625, 50, 53, mdFaceDown, objForm.BackColor)
    objCard.Moveable = False
    mobjPile4.Add objCard
    
    Set objCard = mobjCardsAPI.DrawCard(objForm.hdc, 75, 50, 54, mdFaceUp, _
    objForm.BackColor)
    
    mobjDiscard.Add objCard
    
    
    mobjForm.Refresh
ExitSub:
    Exit Function
ErrorHandler:
    Start = Err.Description
    
End Function
Public Function CardDropped(X, Y)
    
    On Error Resume Next
    If mbCardBeingDragged = True Then
        
        If mbPileSelected = True Then
            'add the drag card to its new pile or "discard" it
            AddDragCard
            mbPileSelected = False
        Else
            'add the drag card back to its source pile
            AddBackDragCard
        End If
        
        mbCardBeingDragged = False
        RedrawBoard
        mobjForm.Refresh
    End If
    
    If Me.DealsLeft = 0 And Me.ValidMovesExist = False Then
        RaiseEvent GameOver
    End If
End Function

Private Sub CheckPileForDragOver(ByRef robjPile As CCards, X As Long, Y As _
Long)
    Dim objCard As CCard
    Dim udtRect As RECT

    'check the passed pile to see if a card is being dragged over, if it
    'is then highlight the pile
    If robjPile.PileID <> mobjDragSrcPile.PileID Then
        'don't highlight the source pile
        If robjPile.Count = 1 Then
            'only highlight "empty" piles
            Set objCard = robjPile(1)
            udtRect.bottom = objCard.bottom
            udtRect.Top = objCard.Top
            udtRect.left = objCard.left
            udtRect.right = objCard.right
        
            If PtInRect(udtRect, X, Y) = 1 And robjPile.Selected = False Then
                'highlight the selected pile
                Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, objCard.left, _
                objCard.Top, objCard.Card, mdHilite, mobjForm.BackColor)
                
                robjPile.Selected = True
                
                robjPile.Remove 1
                robjPile.Add objCard
                
                mbPileSelected = True
            ElseIf PtInRect(udtRect, X, Y) = 0 And robjPile.Selected = True Then
                'un-highlight the pile
                                    
                If objCard.Card = 53 Then
                    Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, objCard.left, _
                    objCard.Top, objCard.Card, mdFaceDown, mobjForm.BackColor)
                ElseIf objCard.Card = 54 Then
                    Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, objCard.left, _
                    objCard.Top, objCard.Card, mdFaceUp, mobjForm.BackColor)
                End If
                
                robjPile.Selected = False
                
                robjPile.Remove 1
                robjPile.Add objCard
                
                mbPileSelected = False
            End If
        End If
    End If
End Sub
Private Function CheckPileForGrab(ByRef robjPile As CCards, X As Long, Y As Long) _
As Boolean
    Dim objCard As CCard
    Dim udtRect As RECT

    'check the passed pile to see if a card was grabbed
    If robjPile.Count > 1 Then
        Set objCard = robjPile(robjPile.Count)
        
        
        
        udtRect.bottom = objCard.bottom
        udtRect.Top = objCard.Top
        udtRect.left = objCard.left
        udtRect.right = objCard.right
    
        If PtInRect(udtRect, X, Y) = 1 Then
            mudtDragPos.X = X
            mudtDragPos.Y = Y
            Set mobjDragCard = objCard
            RemoveDragCard robjPile
            mbCardBeingDragged = True
            CheckPileForGrab = True
            
            'Any top card can be grabbed, but not all top cards are
            'eligible for removal. let the user know if the card is not
            'eligible for removal
            If IsCardRemoveEligible(mobjDragCard, robjPile.PileID) = False Then
                RaiseEvent Msg("This card is not eligible for removal")
            End If
            
            
        End If
    End If
End Function
Public Function Deal()
    Dim objCard As CCard
    Dim lCardDealt As Long
    Dim iCnt As Integer
    
    If miDeals < TOTAL_DEALS Then
        miDeals = miDeals + 1
    
    
        'deal 4 cards
        lCardDealt = mobjDeck.GetTopCard
        If lCardDealt = ALL_CARDS_DEALT Then
            Exit Function
        End If
        
        'deal to pile 1
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, mobjPile1.GetNextCardLeft, _
        mobjPile1.GetNextCardTop, lCardDealt, mdFaceUp, vbWhite)
        
        mobjPile1.Add objCard
        
        'deal to pile 2
        lCardDealt = mobjDeck.GetTopCard
        If lCardDealt = ALL_CARDS_DEALT Then
            GoTo ExitSub
        End If
        
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, mobjPile2.GetNextCardLeft, _
        mobjPile2.GetNextCardTop, lCardDealt, mdFaceUp, vbWhite)
    
        mobjPile2.Add objCard
    
        'deal to pile 3
        lCardDealt = mobjDeck.GetTopCard
        If lCardDealt = ALL_CARDS_DEALT Then
            GoTo ExitSub
        End If
        
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, mobjPile3.GetNextCardLeft, _
        mobjPile3.GetNextCardTop, lCardDealt, mdFaceUp, vbWhite)
        
        mobjPile3.Add objCard
        
        
        'deal to pile 4
        lCardDealt = mobjDeck.GetTopCard
        If lCardDealt = ALL_CARDS_DEALT Then
            GoTo ExitSub
        End If
        
        Set objCard = mobjCardsAPI.DrawCard(mobjForm.hdc, mobjPile4.GetNextCardLeft, _
        mobjPile4.GetNextCardTop, lCardDealt, mdFaceUp, vbWhite)
        
        mobjPile4.Add objCard
    
    End If

    If Me.DealsLeft = 0 And Me.ValidMovesExist = False Then
        RaiseEvent GameOver
    End If
    
ExitSub:
    mobjForm.Refresh
End Function
Public Property Get DealsLeft() As Long
    DealsLeft = TOTAL_DEALS - miDeals
End Property
Public Function GameHasBeenWon() As Boolean
'determine whether the user won the game or not

On Error GoTo ErrorHandler
    
    If mobjPile1.Count = 2 And mobjPile1.Item(2).Card <= 3 Then
        If mobjPile2.Count = 2 And mobjPile2.Item(2).Card <= 3 Then
            If mobjPile3.Count = 2 And mobjPile3.Item(2).Card <= 3 Then
                If mobjPile4.Count = 2 And mobjPile4.Item(2).Card <= 3 Then
                    GameHasBeenWon = True
                End If
            End If
        End If
    End If

    Exit Function
ErrorHandler:
    GameHasBeenWon = False
End Function
Private Function IsCardRemoveEligible(ByRef robjCard As CCard, iPileID _
As Integer) As Boolean
'robjcard - card the user wants to grab or has grabbed
'This function checks to see if its eligible to be removed
    Dim objCard As CCard
    Dim bHigherCard As Boolean

    Set objCard = mobjPile1.Item(mobjPile1.Count)
    If mobjPile1.Count > 1 Then
        If ((robjCard.Card Mod 4) = (objCard.Card Mod 4)) And _
        (iPileID <> mobjPile1.PileID) And ((robjCard.Card < objCard.Card) _
        Or objCard.Card <= 3) Then
            bHigherCard = True
        End If
    End If
    
    Set objCard = mobjPile2.Item(mobjPile2.Count)
    If mobjPile2.Count > 1 Then
        If ((robjCard.Card Mod 4) = (objCard.Card Mod 4)) And _
        (iPileID <> mobjPile2.PileID) And ((robjCard.Card < objCard.Card) _
        Or objCard.Card <= 3) Then
            bHigherCard = True
        End If
    End If

    If mobjPile3.Count > 1 Then
        Set objCard = mobjPile3.Item(mobjPile3.Count)
        If ((robjCard.Card Mod 4) = (objCard.Card Mod 4)) And _
        (iPileID <> mobjPile3.PileID) And ((robjCard.Card < objCard.Card) _
        Or objCard.Card <= 3) Then
            bHigherCard = True
        End If
    End If

    Set objCard = mobjPile4.Item(mobjPile4.Count)
    
    If mobjPile4.Count > 1 Then
        If ((robjCard.Card Mod 4) = (objCard.Card Mod 4)) And _
        (iPileID <> mobjPile4.PileID) And ((robjCard.Card < objCard.Card) _
        Or objCard.Card <= 3) Then
            bHigherCard = True
        End If
    End If
    
    IsCardRemoveEligible = bHigherCard

End Function

Public Sub MoveCard(X As Long, Y As Long)
    
        
    If mbCardBeingDragged = True Then
        
        'check all the pile to see if the user is dragged a card
        'over the pile.  If the pile is "Empty" the user can
        'drop a card on the pile.  Note: the discard pile
        'should always be "empty"
        CheckPileForDragOver mobjPile1, X, Y
        CheckPileForDragOver mobjPile2, X, Y
        CheckPileForDragOver mobjPile3, X, Y
        CheckPileForDragOver mobjPile4, X, Y
        
        'only highlight the discard pile, if the card being dragged is
        'eligible for removal
        If IsCardRemoveEligible(mobjDragCard, mobjDragSrcPile.PileID) = True Then
            'the user can remove the card
            CheckPileForDragOver mobjDiscard, X, Y
        End If
        
        'redraw the board
        Me.RedrawBoard
        
        'draw the card being dragged, in its new location
        mobjCardsAPI.DrawCard mobjForm.hdc, mobjDragCard.left + (X - mudtDragPos.X), _
        mobjDragCard.Top + (Y - mudtDragPos.Y), mobjDragCard.Card, mdFaceUp, vbWhite
        mobjForm.Refresh
    End If
    
End Sub
Public Function WasCardGrabbed(X As Long, Y As Long) As Boolean
    'this should only be called when the user clicks the mouse
    Dim objCard As CCard
    Dim udtRect As RECT
    
    'check each pile to see if it was grabbed
    WasCardGrabbed = CheckPileForGrab(mobjPile1, X, Y)
    If WasCardGrabbed = False Then
        WasCardGrabbed = CheckPileForGrab(mobjPile2, X, Y)
        If WasCardGrabbed = False Then
            WasCardGrabbed = CheckPileForGrab(mobjPile3, X, Y)
            If WasCardGrabbed = False Then
                WasCardGrabbed = CheckPileForGrab(mobjPile4, X, Y)
            End If
        End If
    End If
    
    Me.MoveCard X, Y
ExitSub:
    Exit Function
End Function
Public Function ValidMovesExist() As Boolean
    'determines whether any valid moves exist
    On Error GoTo ErrorHandler
    
    
    
    'check each pile to see if the top card can be removed
    If IsCardRemoveEligible(mobjPile1(mobjPile1.Count), mobjPile1.PileID) = True Then
        ValidMovesExist = True
    ElseIf IsCardRemoveEligible(mobjPile2(mobjPile2.Count), mobjPile2.PileID) = True Then
        ValidMovesExist = True
    ElseIf IsCardRemoveEligible(mobjPile3(mobjPile3.Count), mobjPile3.PileID) = True Then
        ValidMovesExist = True
    ElseIf IsCardRemoveEligible(mobjPile4(mobjPile4.Count), mobjPile4.PileID) = True Then
        ValidMovesExist = True
    ElseIf (mobjPile1.Count > 2 Or mobjPile2.Count > 2 Or mobjPile3.Count > 2 Or _
    mobjPile4.Count > 2) And (mobjPile1.Count = 1 Or mobjPile2.Count = 1 Or _
    mobjPile3.Count = 1 Or mobjPile4.Count = 1) Then
        ValidMovesExist = True
    Else
        ValidMovesExist = False
    End If
    
    Exit Function
ErrorHandler:
    ValidMovesExist = False
End Function

Private Sub Class_Initialize()
    Set mobjCardsAPI = New CCardsAPI
    Set mobjPile1 = New CCards
    mobjPile1.PileID = 1
    Set mobjPile2 = New CCards
    mobjPile2.PileID = 2
    Set mobjPile3 = New CCards
    mobjPile3.PileID = 3
    Set mobjPile4 = New CCards
    mobjPile4.PileID = 4
    Set mobjDiscard = New CCards
    mobjDiscard.PileID = 0
    Set mobjDeck = New CDeck
    
End Sub
