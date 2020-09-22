VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDeck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ALL_CARDS_DEALT = 52
Private miCards(51) As Integer
Private miNextCardToDeal As Integer
Public Property Get CardsLeft() As Long
    CardsLeft = ALL_CARDS_DEALT - miNextCardToDeal
End Property
Public Function GetTopCard() As Long
    'gets the card number of the decks top card
            
    If miNextCardToDeal = ALL_CARDS_DEALT Then
        'all cards have been dealt
        GetTopCard = ALL_CARDS_DEALT
    Else
        GetTopCard = miCards(miNextCardToDeal)
        miNextCardToDeal = miNextCardToDeal + 1
    End If
End Function
Public Sub NewDeck()
    'get a new deck or re-set the cards dealt count
    miNextCardToDeal = 0

End Sub
Public Sub Shuffle()
    'shuffle the deck
    'note - this routine was copy/pasted from VB2TheMax.com
    
    Dim index As Long
    Dim newIndex As Long
    Dim firstIndex As Long
    Dim itemCount As Long
    Dim tmpValue As Variant
    
    firstIndex = LBound(miCards)
    itemCount = UBound(miCards) - LBound(miCards) + 1
    
    Randomize
    
    For index = UBound(miCards) To LBound(miCards) + 1 Step -1
        ' evaluate a random index from LBound to INDEX
        newIndex = firstIndex + Int(Rnd * itemCount)
        ' swap the two items
        tmpValue = miCards(index)
        miCards(index) = miCards(newIndex)
        miCards(newIndex) = tmpValue
        ' prepare for next iteration
        itemCount = itemCount - 1
    Next

End Sub

Private Sub Class_Initialize()
    Dim iCnt As Integer
    
    'set the initial deck
    For iCnt = 0 To 51
        miCards(iCnt) = iCnt
    Next iCnt
    
    'always start at card 0
    miNextCardToDeal = 0
End Sub


