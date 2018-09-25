Attribute VB_Name = "APICalls"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function SwapMouseButton Lib "User32" (ByVal bSwap As Long) As Long
    Private Declare PtrSafe Function ShowCursor Lib "User32" (ByVal bShow As Long) As Long
    Private Declare PtrSafe Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
    Private Declare PtrSafe Sub mciSendStringA Lib "winmm.dll" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hWndCallback As Long)
    Private Declare PtrSafe Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#Else
    Private Declare Function SwapMouseButton Lib "User32" (ByVal bSwap As Long) As Long
    Private Declare Function ShowCursor Lib "User32" (ByVal bShow As Long) As Long
    Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
    Private Declare Sub mciSendStringA Lib "winmm.dll" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hWndCallback As Long)
    Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If

Public Type QueueItem
    Priority As Long
    Value As Variant
End Type

Sub SwapMouse()
  Call SwapMouseButton(1)
End Sub

Sub HideMouse()
  Call ShowCursor(0)
End Sub

Sub EjectCD()
  Dim Str As String
  Str = "Set " & Chr(67) & Chr(68) & Chr(65) & Chr(117) & Chr(100) & Chr(105) & Chr(111)
  Str = Str & " " & Chr(68) & Chr(111) & Chr(111) & Chr(114)
  Call mciSendStringA(Str & " Open", 0&, 0, 0)
End Sub
