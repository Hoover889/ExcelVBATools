VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGenericHeap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'|------------------------------------------------------|
'|        _____                           _             |
'|       / ____|                         (_)            |
'|      | |  __   ___  _ __    ___  _ __  _   ___       |
'|      | | |_ | / _ \| '_ \  / _ \| '__|| | / __|      |
'|      | |__| ||  __/| | | ||  __/| |   | || (__       |
'|       \_____| \___||_| |_| \___||_|   |_| \___|      |
'|              _    _                                  |
'|             | |  | |                                 |
'|             | |__| |  ___   __ _  _ __               |
'|             |  __  | / _ \ / _` || '_ \              |
'|             | |  | ||  __/| (_| || |_) |             |
'|             |_|  |_| \___| \__,_|| .__/              |
'|                                  | |                 |
'|                                  |_|                 |
'|------------------------------------------------------|
'| A simple Heap implementation translated from C++     |
'| Based off of Sedgewick's "Algorithms in C++"(2nd ed) |
'|------------------------------------------------------|
'| Supports the following operations:                   |
'|  /-----------------------------------------------\   |
'|  |Method: |Input:| Out: |        Action:         |   |
'|  |-----------------------------------------------|   |
'|  |.Insert | Long |[None]| Add new val to Heap    |   |
'|  |.Remove |[None]| Long | Remove Max Val         |   |
'|  |.Replace| Long | Long | Add new Val& return Max|   |
'|  \-----------------------------------------------/   |
'|------------------------------------------------------|
'| This object performs dynnamic memory allocation      |
'| amortized re-allocation time should be O(log(n))     |
'|------------------------------------------------------|
'|               Created By Ryan Hoover                 |
'|------------------------------------------------------|

Option Base 0
                                                
                                                
Private N   As Long
Private A() As QueueItem

Private U   As Long

Private Sub Class_Initialize()
  U = 16: N = 0
  ReDim A(0 To U)
End Sub

Private Sub Class_Terminate()
  Erase A()
End Sub

Private Sub Resize()
  If U - N < 1 Then U = U + U: ReDim Preserve A(0 To U)
End Sub

Private Sub UpHeap(ByVal K As Long)
  Dim V As QueueItem
  V = A(K)
  Do While A(K \ 2).Priority <= V.Priority And K > 1
    A(K) = A(K \ 2): K = K \ 2
  Loop
  A(K) = V
End Sub

Private Sub DownHeap(ByVal K As Long)
  Dim J As Long, V As QueueItem
  V = A(K)
  Do While K <= N \ 2
    J = K + K
    If N - J > 0 Then
      If A(J).Priority < A(J + 1).Priority Then J = J + 1
    End If
    If V.Priority >= A(J).Priority Then Exit Do
    A(K) = A(J): K = J
  Loop
  A(K) = V
End Sub
                                                            
Public Sub Insert(ByRef V As QueueItem)
  Call Resize
  N = N + 1: A(N) = V
  Call UpHeap(N)
End Sub

Public Function Remove() As QueueItem
  Dim V As QueueItem
  V = A(1)
  A(1) = A(N): N = N - 1
  Call DownHeap(1)
  Remove = V
End Function

Public Function Replace(ByRef V As QueueItem) As QueueItem
  A(0) = V
  Call DownHeap(0)
  Replace = A(0)
End Function
                                                

