Attribute VB_Name = "Misc_Functions"
Option Explicit
Private Const RegPath As String = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Excel\Security\Trusted Locations\"
Private Const DW As String = "REG_DWORD"



' Two-argument Ackermann–Peter function
Public Function Ackermann(ByVal M As Long, ByVal N As Long) As Long
    Ackermann = IIf(M = 0, N + 1, Ackermann(M - 1, IIf(N = 0, 1, Ackermann(M, N - 1))))
End Function

Private Function DOS(ByVal Start As Double, ByRef Demand As Range) As Double
  Dim Tot As Double, D As Double, I As Long, M As Long
  M = Demand.Columns.Count
  For I = 1 To M
    D = Demand.Cells(1, I).Value2
    If Start >= D Then
      Start = Start - D
      Tot = Tot + D
      DOS = DOS + 30
    Else
      DOS = DOS + (Start / D) * 30
      Exit Function
    End If
  Next I
  If Tot <= 0 Then
    DOS = -1
  Else
    DOS = DOS + ((Start * M * 30) / Tot)
  End If
End Function

Public Function MOC(ByVal Start As Double, ByRef Demand As Range, Optional ByRef AvgMo As Double = -1) As Double
  Dim Tot As Double, D As Double, I As Long, M As Long
  If Start < 0 Then MOC = 0: Exit Function
  M = Demand.Columns.Count
  For I = 1 To M
    D = Demand.Cells(1, I).Value2
    If Start >= D Then
      Start = Start - D
      Tot = Tot + D
      MOC = MOC + 1
    Else
      MOC = MOC + (Start / D)
      Exit Function
    End If
  Next I
  If AvgMo = -1 Then
    If Tot <= 0 Then
      MOC = -1
    Else
      MOC = MOC + ((Start * M) / Tot)
    End If
  Else
    If AvgMo <= 0 Then
      MOC = -1
    Else
      MOC = MOC + (Start / AvgMo)
    End If
  End If
End Function

'Sub RegKeySave():   Dim WS As WshShell: Set WS = New WshShell
'    WS.RegWrite RegPath & "AllowNetworkLocations", 1, DW: WS.RegWrite RegPath & "Winshuttle\AllowSubfolders", 1, DW
'    WS.RegWrite RegPath & "Winshuttle\Path", "\\usmfgpwytecfil1\logistics$\Winshuttle\", "REG_SZ": Set WS = Nothing
'End Sub

Sub PasteSelectedColumns(ByRef SourceWS As Worksheet, _
                         ByRef TargetRng As Range, _
                         ByRef ColHeaders() As Variant, _
                Optional ByVal Values As Boolean = False, _
                Optional ByVal DebugMode As Boolean = False)
                
    'Put in a Source workbook, Destination Range, and an array of column headers
    'this sub will paste only the selected columns (in the order listed) at the specified destination
    'Warning, if the source or header list contains duplicates you may not get the column you were looking for.
  Dim Upper As Long, Lower As Long, Found As Long, I As Long, J As Long
  Lower = LBound(ColHeaders): Upper = UBound(ColHeaders)
  With SourceWS
    If DebugMode Then
      For I = Lower To Upper: For J = I + 1 To Upper
        If ColHeaders(I) = ColHeaders(J) Then MsgBox "Duplicate String detected in Column Headers": Exit Sub
      Next J: Next I
      For I = 1 To 255
        If Len(.Cells(1, I).Value2 & "") = 0 Then Exit For
        For J = I + 1 To 255
          If Len(.Cells(1, J).Value2 & "") = 0 Then Exit For
          If .Cells(1, I).Value2 = .Cells(1, J).Value2 Then MsgBox "Duplicate String detected in SourceWorksheet Headers": Exit Sub
        Next J
        For J = Lower To Upper
          If .Cells(1, I).Value2 = ColHeaders(J) Then Found = Found + 1
        Next J
      Next I
      If Found <> (Upper - Lower) Then MsgBox "Of " & (Lower - Upper) & " Column Headers, " & Found & " Were Found.": Exit Sub
    End If
    Found = 0
    For I = 1 To 255
      For J = Lower To Upper
        If .Cells(1, I).Value2 = ColHeaders(J) Then
          .Range(.Cells(1, I), .Cells(.UsedRange.Rows.Count, I)).Copy
          If Values Then
            TargetRng.Offset(0, J - Lower).PasteSpecial xlPasteValues
          Else
            TargetRng.Offset(0, J - Lower).Paste
          End If
          Application.CutCopyMode = False
          Found = Found + 1
          Exit For
        End If
      Next J
      If Found >= Upper - Lower Then Exit For
    Next I
  End With
End Sub


Function RangetoHTML(ByRef Rng As Range)
  Dim FSO As Scripting.FileSystemObject
  Dim TS  As Scripting.TextStream
  Dim TempFile As String
  Dim TempWB As Workbook

  TempFile = Environ$("TEMP") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

  'Copy the range and create a new workbook to past the data in
  Rng.Copy
  Set TempWB = Workbooks.Add(1)
  With TempWB.Sheets(1)
    .Cells(1).PasteSpecial Paste:=8
    .Cells(1).PasteSpecial xlPasteValues, , False, False
    .Cells(1).PasteSpecial xlPasteFormats, , False, False
    .Cells(1).Select
    Application.CutCopyMode = False
    On Error Resume Next
    .DrawingObjects.Visible = True
    .DrawingObjects.Delete
    On Error GoTo 0
  End With

  'Publish the sheet to a htm file
  With TempWB.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        Filename:=TempFile, _
        Sheet:=TempWB.Sheets(1).Name, _
        Source:=TempWB.Sheets(1).UsedRange.Address, _
        HtmlType:=xlHtmlStatic)
    .Publish True
  End With

  'Read all data from the htm file into RangetoHTML
  Set FSO = New Scripting.FileSystemObject
  Set TS = FSO.GetFile(TempFile).OpenAsTextStream(1, -2)
  RangetoHTML = TS.ReadAll
  TS.Close
  RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                        "align=left x:publishsource=")

  'Close TempWB
  TempWB.Close savechanges:=False

  'Delete the htm file we used in this function
  Kill TempFile

  Set TS = Nothing
  Set FSO = Nothing
  Set TempWB = Nothing
End Function
