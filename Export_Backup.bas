Attribute VB_Name = "Export_Backup"
Option Explicit

Private Const Module As Long = 1
Private Const ClsMdl As Long = 2
Private Const Form   As Long = 3
Private Const Doc    As Long = 100
  
Public Function ExportVisualBasicCode(ByVal Directory As String, Optional ByRef WB As Workbook = Nothing) As Long
    'directory = "C:\Test\"
  Dim VBComp As VBIDE.VBComponent
  Dim ext    As String
  
  If WB Is Nothing Then Set WB = ThisWorkbook
  For Each VBComp In WB.VBProject.VBComponents
    Select Case VBComp.Type
      Case ClsMdl, Doc: ext = ".cls"
      Case Form:        ext = ".frm"
      Case Module:      ext = ".bas"
      Case Else:        ext = ".txt"
    End Select
    Call VBComp.Export(Directory & "\" & VBComp.Name & ext)
    ExportVisualBasicCode = ExportVisualBasicCode + 1
  Next
End Function

