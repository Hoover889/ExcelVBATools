VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDeQue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------
'|      _____                                                        |
'|     |  __ \                                                       |
'|     | |  | |  ___   __ _  _   _   ___                             |
'|     | |  | | / _ \ / _` || | | | / _ \                            |
'|     | |__| ||  __/| (_| || |_| ||  __/                            |
'|     |_____/  \___| \__, | \__,_| \___|                            |
'|                       | |                                         |
'|                       |_|                                         |
'|                                                                   |
'|-------------- Lightweight Deque Implementation -------------------|
'|                                                                   |
'| A Deque combines the functuionality of a Stack and Queue          |
'|                                                                   |
'| Supported Operations: (All values of type Variant)                |
'|   .Push_Back(Var) - Puts a value at the back of the Deque         |
'|   .Push_Front(Var)- Puts a value at the front of the Deque        |
'|   .Pop_Back()     - Removes and returns the back value            |
'|   .Pop_Front()    - Removes and returns the front value           |
'|   .Back()         - Returns the back value                        |
'|   .Front()        - Returns the front value                       |
'|                                                                   |
'|-------------------------------------------------------------------|
'|               Created By Ryan Hoover                              |
'---------------------------------------------------------------------
Private DQUE As Collection

Private Sub Class_Initialize(): Set DQUE = New Collection: End Sub
Private Sub Class_Terminate(): Set DQUE = Nothing: End Sub
Public Function Push_Back(ByVal Value As Variant) As Variant: DQUE.Add Value, , , DQUE.Count: End Function
Public Function Push_Front(ByVal Value As Variant) As Variant: DQUE.Add Value, , 1: End Function
Public Function Pop_Back() As Variant: Pop_Back = DQUE.Item(DQUE.Count): DQUE.Remove DQUE.Count: End Function
Public Function Pop_Front() As Variant: Pop_Front = DQUE.Item(1): DQUE.Remove 1: End Function
Public Function Back() As Variant: Back = DQUE.Item(DQUE.Count): End Function
Public Function Front() As Variant: Front = DQUE.Item(1): End Function
