VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CCards"
Option Explicit

Private mudtRect As RECT
Private mvarCCards As CCards
Private mlCard As Long
Private mlDrawMode As Long
Private mlColor As Long
Private mbMoveable As Boolean
Public Property Let Color(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Color = 5
    mlColor = vData
End Property
Public Property Get Color() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Color
    Color = mlColor
End Property
Public Property Let DrawMode(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DrawMode = 5
    mlDrawMode = vData
End Property
Public Property Get DrawMode() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DrawMode
    DrawMode = mlDrawMode
End Property
Public Property Let Card(ByVal vData As Long)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Card = Form1
     mlCard = vData
End Property
Public Property Get Card() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Card
     Card = mlCard
End Property
Public Sub SetCardData(ByVal vlLeft As Long, ByVal vlRight As Long, ByVal vlTop As Long, _
ByVal vlBottom As Long, ByVal vlCard As Long, ByVal vlDrawMode As Long, ByVal _
vlColor As Long)
    
    mudtRect.right = vlRight
    mudtRect.left = vlLeft
    mudtRect.Top = vlTop
    mudtRect.bottom = vlBottom
    mlCard = vlCard
    mlDrawMode = vlDrawMode
    mlColor = vlColor
End Sub
Public Property Get Moveable() As Boolean
    Moveable = mbMoveable
End Property
Public Property Let Moveable(vbData As Boolean)
    mbMoveable = vbData
End Property
Public Property Get left() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Left
    left = mudtRect.left
End Property
Public Property Let left(ByVal vlData As Long)
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Left
    mudtRect.left = vlData
End Property

Public Property Get right() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Right
    right = mudtRect.right
End Property
Public Property Get Top() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Top
    Top = mudtRect.Top
End Property
Public Property Let Top(lData As Long)
    mudtRect.Top = lData
End Property

Public Property Get bottom() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Bottom
    bottom = mudtRect.bottom
End Property



Private Sub Class_Initialize()
    mbMoveable = True
End Sub
