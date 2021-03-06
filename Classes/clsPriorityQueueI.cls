VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPriorityQueueI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
'---------------------------------------------------------------------
'|    _____        _               _  _                              |
'|   |  __ \      (_)             (_)| |                             |
'|   | |__) |_ __  _   ___   _ __  _ | |_  _   _                     |
'|   |  ___/| '__|| | / _ \ | '__|| || __|| | | |                    |
'|   | |    | |   | || (_) || |   | || |_ | |_| |                    |
'|   |_|    |_|   |_| \___/ |_|   |_| \__| \__, |                    |
'|     ____                                 __/ |                    |
'|    / __ \                               |___/                     |
'|   | |  | | _   _   ___  _   _   ___                               |
'|   | |  | || | | | / _ \| | | | / _ \                              |
'|   | |__| || |_| ||  __/| |_| ||  __/                              |
'|    \___\_\ \__,_| \___| \__,_| \___|                              |
'|                                                                   |
'|                                                                   |
'|---------------- Priority Queue Implementation --------------------|
'|                                                                   |
'| A Queue that accepts QueueItem structs (Key + Value)              |
'| this Priority Queue uses an 'Insersion sort' type implementation  |
'| which allows items to be addeed in O(n) time while removing       |
'| items requires O(1) time                                          |
'|                                                                   |
'| Supported Operations:                                             |
'|   .Push() - Puts an object on the top of the queue                |
'|   .Pop()  - Removes and returns the top object from the queue     |
'|                                                                   |
'|-------------------------------------------------------------------|
'|               Created By Ryan Hoover                              |
'---------------------------------------------------------------------

Private A() As QueueItem
Private N   As Long
Private U   As Long

Private Sub Class_Initialize()
  N = 0: U = 16: ReDim A(0 To U)
End Sub

Private Sub Class_Terminate()
  Erase A
End Sub

Private Sub Resize()
  If U - N < 1 Then U = U + U: ReDim Preserve A(0 To U)
End Sub

Public Sub Push(ByRef V As QueueItem)
  Dim I As Long: Call Resize
  For I = 1 To N
    If V.Priority < A(I).Priority Then A(0) = A(I): A(I) = V: V = A(0)
  Next I
  N = N + 1: A(N) = V
End Sub

Public Function Pop() As QueueItem
  Pop = A(N): N = N - 1
End Function
