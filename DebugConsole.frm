VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DebugConsole 
   Caption         =   "Please Wait..."
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9705
   OleObjectBlob   =   "DebugConsole.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DebugConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'-------------------------------------------------------------------------------------------------------------------------
'|    _____         _                                 1                                                                   |
'|   |  __ \       | |                                2                                                                   |
'|   | |  | |  ___ | |__   _   _   __ _               3                                                                   |
'|   | |  | | / _ \| '_ \ | | | | / _` |              4                                                                   |
'|   | |__| ||  __/| |_) || |_| || (_| |              5                                                                   |
'|   |_____/  \___||_.__/  \__,_| \__, |              6                                                                   |
'|        _____                    __/ |  _           7                                                                   |
'|       / ____|                  |___/  | |          8                                                                   |
'|      | |      ___   _ __   ___   ___  | |  ___     9                                                                   |
'|      | |     / _ \ | '_ \ / __| / _ \ | | / _ \    0                                                                   |
'|      | |____| (_) || | | |\__ \| (_) || ||  __/    1                                                                   |
'|       \_____|\___/ |_| |_||___/ \___/ |_| \___|    2                                                                   |
'|                                                    3                                                                   |
'| Created By Ryan Hoover                                                                                                |
'|            Ryan.Hoover@Loreal.com                                                                                     |
'-------------------------------------------------------------------------------------------------------------------------
'|                                                                                                                       |
'|                                                                                                                       |
'|                                                                                                                       |
'|                                                                                                                       |
'|                                                                                                                       |
'|                                                                                                                       |
'|                                                                                                                       |
'|                                                                                                                       |
'|                                                                                                                       |
'|                                                                                                                       |
'|                                                                                                                       |
'| Enjoy...                                                                                                              |
'| Created By Ryan Hoover                                                                                                |
'|            Ryan.Hoover@LOreal.com                                                                                     |
'|                                                                                                                       |
'| Feel free to use or modify this code however you like but please provide attribution                                  |
'| If you like my work be sure to say thanks.                                                                            |
'-------------------------------------------------------------------------------------------------------------------------

Public MaxLines As Long
Public MaxLen   As Long
Public strText  As Long

Private Sub UserForm_Activate():
  MaxLines = 19
  MaxLen = 2048
  DoEvents
  Me.Repaint
End Sub

Public Sub ChangeText(ByVal Txt As String)
  strText = TrimConsole(Txt, Me.MaxLen, Me.MaxLines)
  Me.CustomLabel.Caption = strText
  DoEvents
  Me.Repaint
End Sub

Public Sub AppendLine(ByVal Txt As String)
  strText = strText & vbNewLine & Txt
  strText = TrimConsole(strText, Me.MaxLen, Me.MaxLines)
  Me.CustomLabel.Caption = strText
  DoEvents
  Me.Repaint
End Sub

Private Function TrimConsole(ByVal Msg As String, Optional ByVal MaxChars As Long = 2048, Optional ByVal MaxLines As Long = 19) As String
  Dim I As Long
  Dim Numlines As Long
  Dim Length As Long
  Dim SecondLinePos As Long
  Numlines = CountLines(Msg)
  If Numlines > MaxLines Then
    For I = MaxLines + 1 To Numlines
      Msg = RemoveLine(Msg)
    Next I
  End If
  Do While Len(Msg) >= MaxChars
    SecondLinePos = InStr(1, Msg, vbNewLine) + Len(vbNewLine)
    Msg = Mid(Msg, SecondLinePos, (Len(Msg) - SecondLinePos) + 1)
  Loop
  TrimConsole = Msg
End Function

Private Function CountLines(ByVal Msg As String, Optional ByVal lineDelim As String = vbNewLine) As Long
  Dim curPos As Long
  If Len(Msg & vbNullString) = 0 Then CountLines = 0: Exit Function
  CountLines = 1
  curPos = 1
  Do
    curPos = InStr(curPos, Msg, lineDelim)
    If curPos < 1 Then Exit Function
    CountLines = CountLines + 1
    curPos = curPos + Len(lineDelim)
  Loop
End Function

Private Function RemoveLine(ByVal Msg As String, Optional ByVal lineDelim As String = vbNewLine) As String
  Dim SecondLinePos As Long
  SecondLinePos = InStr(1, Msg, lineDelim) + Len(lineDelim)
  If SecondLinePos > 0 Then
    RemoveLine = Mid(Msg, SecondLinePos, (Len(Msg) - SecondLinePos) + 1)
  Else
    RemoveLine = Msg
  End If
End Function



Private Function Regex(ByVal Str As String, ByVal pattern As String, Optional IgnoreCase As Boolean = True) As Long
  With CreateObject("vbscript.regexp")
    .pattern = pattern
    .IgnoreCase = IgnoreCase
    .Global = True
    Regex = .Execute(Str).Count
  End With
End Function
