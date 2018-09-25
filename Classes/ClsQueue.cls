VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------
'|       ____                                                        |
'|      / __ \                                                       |
'|     | |  | | _   _   ___  _   _   ___                             |
'|     | |  | || | | | / _ \| | | | / _ \                            |
'|     | |__| || |_| ||  __/| |_| ||  __/                            |
'|      \___\_\ \__,_| \___| \__,_| \___|                            |
'|                                                                   |
'|                                                                   |
'|-------------- Lightweight Queue Implementation -------------------|
'|                                                                   |
'| A Queue returns items added to it in a FIFO manner                |
'|                                                                   |
'| Supported Operations: (All values of type Variant)                |
'|   .Push(Var) - Puts a value on the bottom of the queue            |
'|   .Pop()     - Removes and returns the top value from the queue   |
'|   .Peek()    - Returns the top value from the queue               |
'|                                                                   |
'|-------------------------------------------------------------------|
'|               Created By Ryan Hoover                              |
'---------------------------------------------------------------------

Private QUE As Collection
Private Sub Class_Initialize(): Set QUE = New Collection: End Sub
Private Sub Class_Terminate(): Set QUE = Nothing: End Sub
Public Function Push(ByVal Value As Variant) As Variant:
    QUE.Add Value: Push = QUE.Item(QUE.Count)
End Function
Public Function Pop() As Variant
    If QUE.Count > 0 Then
      Pop = QUE.Item(1): QUE.Remove 1
    End If
End Function
Public Function Peek() As Variant
Attribute Peek.VB_UserMemId = 0
    If QUE.Count > 0 Then Peek = QUE.Item(0)
End Function
Public Function Clone() As ClsQueue
    Dim I As Long:  Set Clone = New ClsQueue
    For I = 1 To QUE.Count: Clone.Push QUE.Item(I): Next I
End Function

