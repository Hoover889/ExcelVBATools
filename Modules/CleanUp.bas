Attribute VB_Name = "CleanUp"
Option Explicit

Public Sub CleanupStyles()
  'This subroutine removes all custom Styles in a workbook leaving only the excel Default Styles, this can drastically reduce workbook size
  Dim St As Style
  Dim Count As Long, I As Long
On Error Resume Next
  Count = ActiveWorkbook.Styles.Count
  I = 1
  For Each St In ActiveWorkbook.Styles
    Debug.Print I & " of " & Count & " - " & St.Name
    If Not St.BuiltIn Then St.Delete
    DoEvents
    I = I + 1
  Next
End Sub
