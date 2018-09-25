VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------
'|       _____  _                _                                   |
'|      / ____|| |              | |                                  |
'|     | (___  | |_  __ _   ___ | | __                               |
'|      \___ \ | __|/ _` | / __|| |/ /                               |
'|      ____) || |_| (_| || (__ |   <                                |
'|     |_____/  \__|\__,_| \___||_|\_\                               |
'|                                                                   |
'|                                                                   |
'|-------------- Lightweight Stack Implementation -------------------|
'|                                                                   |
'| A Stack returns items added to it in a LIFO manner                |
'|                                                                   |
'| Supported Operations: (All values of type Variant)                |
'|   .Push(Var) - Puts a value on the top of the stack               |
'|   .Pop()     - Removes and returns the top value from the stack   |
'|   .Peek()    - Returns the top value from the stack               |
'|                                                                   |
'|---------------Advanced Operations:--------------------------------|
'|                                                                   |
'|   .Clone()   - Returns a deep copy of the stack object            |
'|                   Implementation of interface ICloneable          |
'|                                                                   |
'|                                                                   |
'|-------------------------------------------------------------------|
'|               Created By Ryan Hoover                              |
'---------------------------------------------------------------------

Private STK As Collection
Private Sub Class_Initialize(): Set STK = New Collection: End Sub
Private Sub Class_Terminate(): Set STK = Nothing: End Sub
Public Function Push(ByVal Value As Variant) As Variant
    STK.Add Value: Push = STK.Item(STK.Count)
End Function
Public Function Pop() As Variant
    If STK.Count > 0 Then Pop = STK.Item(STK.Count): STK.Remove STK.Count
End Function
Public Function Peek() As Variant
Attribute Peek.VB_UserMemId = 0
    If STK.Count > 0 Then Peek = STK.Item(STK.Count)
End Function
Public Function Clone() As ClsStack
    Dim I As Long:  Set Clone = New ClsStack
    For I = 1 To STK.Count: Clone.Push STK.Item(I): Next I
End Function


