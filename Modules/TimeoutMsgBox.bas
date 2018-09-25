Attribute VB_Name = "TimeoutMsgBox"
Option Explicit

Function MsgBoxTimer(ByVal Prompt As String, _
            Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
            Optional ByVal Title As String = vbNullString, _
            Optional ByVal TimeOut As Integer = 10) As Long
            
    '/---------------------------------------------------\
    '|                                                   |
    '|   \  \            This                            |
    '|   |__/ _   .-.     code                           |
    '|  (o_o)(_`>(   )     has                           |
    '|   { }//||\\`-'       bugs                         |
    '|                                                   |
    '|   This function no longer works in Excel 2016+    |
    '|   due to an unresolved bug in Windows where the   |
    '|   timeout parameter is ignored                    |
    '|                                                   |
    '\---------------------------------------------------/
    
    Dim InfoBox As Object, Result As Long
    If Title = vbNullString Then Title = Application.Name
    Set InfoBox = CreateObject("WScript.Shell")
    Result = InfoBox.Popup(Prompt, TimeOut, Title, Buttons)
    MsgBoxTimer = Result
    Set InfoBox = Nothing
End Function


