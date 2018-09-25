VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------
'|        _       _         _              _                                       |
'|       | |     (_)       | |            | |                                      |
'|       | |      _  _ __  | | __ ___   __| |                                      |
'|       | |     | || '_ \ | |/ // _ \ / _` |                                      |
'|       | |____ | || | | ||   <|  __/| (_| |                                      |
'|       |______||_||_| |_||_|\_\\___| \__,_|                                      |
'|        _       _       _                                                        |
'|       | |     (_)     | |                                                       |
'|       | |      _  ___ | |_                                                      |
'|       | |     | |/ __|| __|                                                     |
'|       | |____ | |\__ \| |_                                                      |
'|       |______||_||___/ \__|                                                     |
'|                                                                                 |
'|---------------------------------------------------------------------------------|
'| A simple linked list which supports adding ,deleting ,and inserting nodes       |
'|                                                                                 |
'| TODO: Add Documentation                                                         |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|---------------------------------------------------------------------------------|
'| Enjoy...                                                                        |
'| Created By Ryan Hoover                                                          |
'|                                                                                 |
'| Feel free to use or modify this code  but please provide attribution            |
'| If you like my work be sure to say thanks.                                      |
'-----------------------------------------------------------------------------------
Private LIHead As ClsListItem
Private ItemCount As Long

Private Sub Class_Initialize()
  ItemCount = 0
End Sub

Public Function AddFront(ByVal Value_To_Add As Variant) As ClsList
  Set AddFront = Me
  Dim Temp As ClsListItem
  Set Temp = LIHead
  Set LIHead = New ClsListItem
  Set LIHead.NextItem = Temp
  LIHead.Value = Value_To_Add
  ItemCount = ItemCount + 1
End Function

Public Function AddEnd(ByVal Value_To_Add As Variant) As ClsList
  Set AddEnd = Me
  Dim LI As ClsListItem
  Set LI = LIHead
  Do Until LI.NextItem Is Nothing
    Set LI = LI.NextItem
  Loop
  Set LI.NextItem = New ClsListItem
  LI.NextItem.Value = Value_To_Add
  ItemCount = ItemCount + 1
End Function

Public Function AddMiddle(ByVal Position As Long, ByVal Value_To_Add As Variant) As ClsList
  Dim LI, LInext As ClsListItem, I As Long
  Set AddMiddle = Me
  If Position = 1 Then AddFront Value_To_Add: Exit Function
  Set LI = LIHead
  For I = 2 To Position - 1
    If Not LI.NextItem Is Nothing Then Set LI = LI.NextItem
  Next I
  Set LInext = LI.NextItem
  Set LI.NextItem = New ClsListItem
  Set LI.NextItem.NextItem = LInext
  LI.NextItem.Value = Value_To_Add
  ItemCount = ItemCount + 1
End Function

Private Function GetItem(ByVal Position As Long) As ClsListItem
  Dim I As Long
  Set GetItem = LIHead
  For I = 2 To Position
    If Not GetItem.NextItem Is Nothing Then Set GetItem = GetItem.NextItem
  Next I
End Function

Public Property Get Value(ByVal Position As Long) As Variant
  If Position > ItemCount Then Exit Property
  Dim LI As ClsListItem
  Set LI = GetItem(Position)
  Value = LI.Value
End Property

Public Property Let Value(ByVal Position As Long, ByVal Val As Variant)
  If Position > ItemCount Then Exit Property
  Dim LI As ClsListItem
  Set LI = GetItem(Position)
  LI.Value = Val
End Property

Public Function DeleteNode(ByVal Position As Long) As ClsList
  If Position > ItemCount Then Exit Function
  Set DeleteNode = Me
  Dim LI As ClsListItem
  If Position = 1 Then
    Set LIHead = LIHead.NextItem
  Else
    Set LI = GetItem(Position - 1)
    Set LI.NextItem = LI.NextItem.NextItem
  End If
  ItemCount = ItemCount - 1
End Function

Public Sub ReCount()
  Dim LI As ClsListItem
  ItemCount = 0
  Set LI = LIHead
  If Not LI Is Nothing Then ItemCount = ItemCount + 1
  Do Until LI.NextItem Is Nothing
    Set LI = LI.NextItem
    ItemCount = ItemCount + 1
  Loop
End Sub

Public Property Get Count() As Long
  Count = ItemCount
End Property
