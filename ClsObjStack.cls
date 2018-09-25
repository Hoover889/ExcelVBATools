VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsObjStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------
'|     ____   _      _              _                                |
'|    / __ \ | |    (_)            | |                               |
'|   | |  | || |__   _   ___   ___ | |_                              |
'|   | |  | || '_ \ | | / _ \ / __|| __|                             |
'|   | |__| || |_) || ||  __/| (__ | |_                              |
'|    \____/ |_.__/ | | \___| \___| \__|                             |
'|     _____  _    _/ |        _                                     |
'|    / ____|| |  |__/        | |                                    |
'|   | (___  | |_  __ _   ___ | | __                                 |
'|    \___ \ | __|/ _` | / __|| |/ /                                 |
'|    ____) || |_| (_| || (__ |   <                                  |
'|   |_____/  \__|\__,_| \___||_|\_\                                 |
'|                                                                   |
'|                                                                   |
'|-------------- Lightweight Stack Implementation -------------------|
'|                                                                   |
'| A Stack returns items added to it in a LIFO manner                |
'|                                                                   |
'| Supported Operations: (All values of type Variant)                |
'|   .Push(Var) - Puts an object on the top of the stack             |
'|   .Pop()     - Removes and returns the top object from the stack  |
'|   .Peek()    - Returns the top Object from the stack              |
'|                                                                   |
'|---------------Advanced Operations:--------------------------------|
'|                                                                   |
'|   .Clone()   - Returns a deep copy of the stack object            |
'|                Does not copy objects stored in the stack          |
'|                   Implementation of interface ICloneable          |
'|                                                                   |
'|                                                                   |
'|-------------------------------------------------------------------|
'|               Created By Ryan Hoover                              |
'---------------------------------------------------------------------

Option Explicit
Private STK As Collection
Private Sub Class_Initialize(): Set STK = New Collection: End Sub
Private Sub Class_Terminate(): Set STK = Nothing: End Sub
Public Function Push(ByRef Value As Object) As Object
    STK.Add Value: Push = STK.Item(STK.Count)
End Function
Public Function Pop() As Object
    If STK.Count > 0 Then Set Pop = STK.Item(STK.Count): STK.Remove STK.Count
End Function
Public Function Peek() As Object
    If STK.Count > 0 Then Set Peek = STK.Item(STK.Count)
End Function
Public Function Clone() As ClsObjStack
    Dim I As Long:  Set Clone = New ClsStack
    For I = 1 To STK.Count: Clone.Push STK.Item(I): Next I
End Function

