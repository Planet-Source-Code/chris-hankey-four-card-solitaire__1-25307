VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CCard"
Attribute VB_Ext_KEY = "Member0" ,"CCard"
Option Explicit

'local variable to hold collection
Private mbSelected As Boolean
Private mCol As Collection

Private mlID As Long
Public Property Get PileID() As Long
    PileID = mlID
End Property
Public Property Let PileID(ByVal vlData As Long)
    mlID = vlData
End Property
Public Sub Add(ByVal objNewMember As CCard)
    'create a new object

    'we shouldn't need to define any Keys to this collection
    mCol.Add objNewMember




End Sub
Public Function GetNextCardLeft() As Long
'All lefts in the pile should be the same so just get any card and return
'the left
    Dim objCard As CCard
    
    
    Set objCard = mCol(1)
    GetNextCardLeft = objCard.left
End Function
Public Function GetNextCardTop() As Long
'The user needs to see all the cards in the pile so cards should be "staggered" when _
dealt or moved to a pile.  This returns the next Top.


    Dim objCard As CCard
    
    
    Set objCard = mCol(mCol.Count)
    'objCard.Top = objCard.Top + 20
    GetNextCardTop = objCard.Top + 20
End Function
Public Property Get Item(vntIndexKey As Variant) As CCard
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub
Public Property Get Selected() As Boolean
    Selected = mbSelected
End Property
Public Property Let Selected(ByVal vbData As Boolean)
    mbSelected = vbData
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

