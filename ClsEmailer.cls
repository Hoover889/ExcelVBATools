VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsEmailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'/---------------------------------------------------------------------------------\
'|     ____          _    _                _                                       |
'|    / __ \        | |  | |              | |                                      |
'|   | |  | | _   _ | |_ | |  ___    ___  | | __                                   |
'|   | |  | || | | || __|| | / _ \  / _ \ | |/ /                                   |
'|   | |__| || |_| || |_ | || (_) || (_) ||   <                                    |
'|    \____/  \__,_| \__||_| \___/  \___/ |_|\_\                                   |
'|    ______                    _  _                                               |
'|   |  ____|                  (_)| |                                              |
'|   | |__    _ __ ___    __ _  _ | |  ___  _ __                                   |
'|   |  __|  | '_ ` _ \  / _` || || | / _ \| '__|                                  |
'|   | |____ | | | | | || (_| || || ||  __/| |                                     |
'|   |______||_| |_| |_| \__,_||_||_| \___||_|                                     |
'|                                                                                 |
'|                                                                                 |
'|---------------------------------------------------------------------------------|
'| Enjoy...                                                                        |
'| Created By Ryan Hoover                                                          |
'|            Ryan.Hoover@LOreal.com                                               |
'|                                                                                 |
'| Feel free to use or modify this code  but please provide attribution            |
'| If you like my work be sure to say thanks.                                      |
'\---------------------------------------------------------------------------------/

Private OLApp As Object
Private BStarted As Boolean

Private Sub Class_Initialize()
  On Error Resume Next
  Set OLApp = GetObject(, "Outlook.Application")
  If Err <> 0 Then
    Set OLApp = CreateObject("Outlook.Application")
    BStarted = True
    Err.Clear
  End If
  On Error GoTo 0
End Sub

Private Sub Class_Terminate()
  If BStarted Then OLApp.Quit
  Set OLApp = Nothing
End Sub

Public Sub Email(ByVal StrTo As String, ByVal StrSubject As String, ByVal StrBody As String, Optional ByVal StrAttach As String = vbNullString)
  With OLApp.CreateItem(0)
    .To = StrTo
    .Subject = StrSubject
    .Body = StrBody
    If Len(StrAttach) > 0 Then .Attachments.Add StrAttach
    .Send
  End With
End Sub

