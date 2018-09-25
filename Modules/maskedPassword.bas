Attribute VB_Name = "maskedPassword"
Option Explicit
'////////////////////////////////////////////////////////////////////
'Password masked inputbox
'Allows you to hide characters entered in a VBA Inputbox.
'
'Code written by Daniel Klann
'March 2003
'Adopted and modified by Trevor Hempel
'64bit compliance by Ryan Hoover
'////////////////////////////////////////////////////////////////////


'API functions to be used
#If VBA7 Then
'  #If VBA64 Then
    Private Declare PtrSafe Function CallNextHookEx Lib "User32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function SetWindowsHookEx Lib "User32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "User32" (ByVal hHook As LongPtr) As Long
    Private Declare PtrSafe Function SendDlgItemMessage Lib "User32" Alias "SendDlgItemMessageA" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
'  #Else
'    Private Declare PtrSafe Function CallNextHookEx Lib "User32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
'    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'    Private Declare PtrSafe Function SetWindowsHookEx Lib "User32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "User32" (ByVal hHook As Long) As Long
'    Private Declare PtrSafe Function SendDlgItemMessage Lib "User32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Private Declare PtrSafe Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
'  #End If
#Else
  Private Declare Function CallNextHookEx Lib "User32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
  Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
  Private Declare Function SetWindowsHookEx Lib "User32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
  Private Declare Function UnhookWindowsHookEx Lib "User32" (ByVal hHook As Long) As Long
  Private Declare Function SendDlgItemMessage Lib "User32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
#End If
#If VBA64 Then
  Private hHook As LongPtr
#Else
  Private hHook As Long
#End If

'#If VBA64 Then
  Private Function NewProc64(ByVal lngCode As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim RetVal As Long, lngBuffer As Long, strCN As String
    If lngCode < 0 Then NewProc64 = CallNextHookEx(hHook, lngCode, wParam, lParam): Exit Function
    strCN = Space$(256):   lngBuffer = 255
    If lngCode = 5 Then    'A window has been activated
      RetVal = GetClassName(wParam, strCN, lngBuffer)
      If Left$(strCN, RetVal) = "#32770" Then SendDlgItemMessage wParam, &H1324, &HCC, Asc("*"), &H0
    End If
    CallNextHookEx hHook, lngCode, wParam, lParam
  End Function
'#Else
'  Private Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Dim RetVal As Long, lngBuffer As Long, strCN As String
'    If lngCode < 0 Then NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam): Exit Function
'    strCN = Space$(256):   lngBuffer = 255
'    If lngCode = 5 Then    'A window has been activated
'      RetVal = GetClassName(wParam, strCN, lngBuffer)
'      If Left$(strCN, RetVal) = "#32770" Then SendDlgItemMessage wParam, &H1324, &HCC, Asc("*"), &H0
'    End If
'    CallNextHookEx hHook, lngCode, wParam, lParam
'  End Function
'#End If

Public Function InputBoxDK(Prompt, Optional Title, Optional Default, Optional XPos, Optional YPos, Optional HelpFile, Optional Context) As String
  '#If VBA64 Then
    Dim lngModHwnd As LongPtr
    Dim lngThreadID As Long
'  #Else
'    Dim lngModHwnd As Long
'    Dim lngThreadID As Long
'  #End If
  
  lngThreadID = GetCurrentThreadId
'  #If VBA64 Then
    lngModHwnd = GetModuleHandle(vbNullString):  hHook = SetWindowsHookEx(5, AddressOf NewProc64, lngModHwnd, lngThreadID)
'  #Else
'    lngModHwnd = GetModuleHandle(vbNullString):  hHook = SetWindowsHookEx(5, AddressOf NewProc64, lngModHwnd, lngThreadID)
'  #End If
  InputBoxDK = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context):  UnhookWindowsHookEx hHook
End Function

