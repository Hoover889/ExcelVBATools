Attribute VB_Name = "MissingShortcuts"
Option Explicit
Public WB As ClsAppEvents
Sub LastSheet()         ' Keyboard Shortcut: Ctrl+g
Attribute LastSheet.VB_Description = "Activates the last used worksheet"
Attribute LastSheet.VB_ProcData.VB_Invoke_Func = "g\n14"
  WB.GoBack
End Sub

Sub PasteValues()       ' Keyboard Shortcut: Ctrl+Shift+V
Attribute PasteValues.VB_ProcData.VB_Invoke_Func = "V\n14"
  If Application.CutCopyMode Then Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

Sub MakeText()          ' Keyboard Shortcut: Ctrl+t
Attribute MakeText.VB_ProcData.VB_Invoke_Func = "t\n14"
  Selection.TextToColumns Destination:=Selection, Tab:=True, FieldInfo:=Array(Array(0, xlTextFormat))
End Sub

Sub RefreshAllPivots()  ' No Keyboard Shortcut, but on Custom Menu-Bar
  Dim PC As PivotCache
  For Each PC In ActiveWorkbook.PivotCaches
    PC.Refresh
  Next PC
End Sub

Sub MacroTimer()
Attribute MacroTimer.VB_ProcData.VB_Invoke_Func = "k\n14"
  ' Test of 2 stage progress bar
  ProgressBar2.Show
End Sub

Sub UnHideAllCells()
  ActiveSheet.UsedRange.Hidden = False
End Sub

Sub UnHideAllSheets()
  Dim WS As Worksheet
  For Each WS In ActiveWorkbook.Worksheets
    WS.Visible = xlSheetVisible
  Next WS
End Sub

