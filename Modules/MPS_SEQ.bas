Attribute VB_Name = "MPS_SEQ"
Option Explicit

Sub FixMPSKickout()
  Dim R As Long
  With ActiveSheet
    R = .UsedRange.Rows.Count
    'Fix Col A Dates
    .Columns("A:A").NumberFormat = "m/d/yy h:mm;@"
    .Range("K2:K" & R).FormulaR1C1 = "=DATE(MID(RC1,7,4)*1,MID(RC1,4,2)*1,1)+TIMEVALUE(MID(RC1,12,8))"
    .Range("A2:A" & R).Value = .Range("K2:K" & R).Value2
    
    'Fix Col C Dates
    .Columns("C:C").NumberFormat = "m/d/yy h:mm;@"
    .Range("K2:K" & R).FormulaR1C1 = "=DATE(MID(RC3,7,4)*1,MID(RC3,4,2)*1,1)+TIMEVALUE(MID(RC3,12,8))"
    .Range("C2:C" & R).Value = .Range("K2:K" & R).Value2
    
    'Fix Col F Dates
    .Columns("F:F").NumberFormat = "[$-409]mmm-yy;@"
    .Range("K2:K" & R).FormulaR1C1 = "=DATE(MID(RC6,7,4)*1,MID(RC6,4,2)*1,1)+TIMEVALUE(MID(RC6,12,8))"
    .Range("F2:F" & R).Value = .Range("K2:K" & R).Value2
    .Columns("K:K").Delete Shift:=xlToLeft
  End With
End Sub

Sub Digest_MPlan()
  Dim WSin As Worksheet, WSout As Worksheet
  Dim I As Long, O As Long, Rows As Long
  Dim MatCode As String, MatDesc As String, RowType As String
  
  Set WSin = ActiveWorkbook.Sheets("M Plan")
  WSin.Cells.Replace "=+inf", "13", xlPart, xlByRows, False, SearchFormat:=False, ReplaceFormat:=False
  Set WSout = ActiveWorkbook.Sheets.Add
  WSout.Name = "FOCST"
  WSin.Range("A1:S1").Copy WSout.Range("A1:S1")
  Rows = WSin.Range("A" & WSin.Cells.Rows.Count).End(xlUp).Row
  O = 1:  RowType = vbNullString
  For I = 2 To Rows
    If WSin.Cells(I, 1) <> " " Then MatCode = WSin.Cells(I, 1): MatDesc = WSin.Cells(I, 2)
    If WSin.Cells(I, 3).Value = " Stock" Then
      RowType = "Stock"
    ElseIf Application.WorksheetFunction.Sum(WSin.Range("D" & I & ":S" & I)) > 0 Then
      Select Case WSin.Cells(I, 3).Value
        Case " D - Prévision":    RowType = "Forecasts M"
        Case " D - Commande":     RowType = "Orders M"
        Case "M Total Plan":      RowType = "M Total Plan"
        Case "Coverage":          RowType = "Coverage"
        Case " Target Stock":     RowType = "Target Stock"
        Case Else:                RowType = vbNullString
      End Select
    Else
      RowType = vbNullString
    End If
    If RowType <> vbNullString Then
      O = O + 1
      WSout.Cells(O, 1).Value = MatCode
      WSout.Cells(O, 2).Value = MatDesc
      WSout.Cells(O, 3).Value = RowType
      WSout.Range("D" & O & ":S" & O).Value = WSin.Range("D" & I & ":S" & I).Value
    End If
  Next I
  WSout.Range("U2").Value = 0
  WSout.Range("U2").Copy
  WSout.Range("D2:S" & O).PasteSpecial xlPasteValues, xlAdd, False, False
  WSout.Range("U2").ClearContents
  MsgBox "Macro Complete"
End Sub

Sub Import_ZSD13()
  Dim Field_Info As Variant
  Dim Path As String
  If Not GetTxtFile(Path) Then Exit Sub
  Field_Info = Array(Array(1, xlTextFormat), Array(2, xlTextFormat), Array(3, xlTextFormat), Array(4, xlTextFormat), Array(5, xlTextFormat), Array(6, xlTextFormat), Array(7, xlTextFormat), _
               Array(8, xlTextFormat), Array(9, xlTextFormat), Array(10, xlTextFormat), Array(11, 1), Array(12, 1), Array(13, xlTextFormat), Array(14, xlTextFormat), Array(15, xlGeneralFormat), _
               Array(16, xlGeneralFormat), Array(17, xlTextFormat), Array(18, xlTextFormat), Array(19, xlTextFormat), Array(20, xlTextFormat), Array(21, xlTextFormat))
               
  Workbooks.OpenText Filename:=Path, Origin:=437, StartRow:=13, DataType:=xlDelimited, TextQualifier:=xlSingleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Field_Info, TrailingMinusNumbers:=True
End Sub



Sub OpenFcst()
  Dim fd As FileDialog
  Dim strpath As String
  Set fd = Application.FileDialog(msoFileDialogFilePicker)
  fd.Filters.Clear
  fd.Filters.Add "Text Files *.txt", "*.txt"
  If Not fd.Show Then Exit Sub
  strpath = fd.SelectedItems(1)
    Workbooks.OpenText Filename:=strpath, Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", _
        FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 9), Array(4, 9), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1)), _
        TrailingMinusNumbers:=True
End Sub

