Attribute VB_Name = "MissingFunctions"
Option Explicit
'Returns the Color Value of a cell as a 6 digit hexidecimal
Public Function CellColor(ByRef Rng As Range, Optional ByVal Obj As String = "Interior") As String
    If Rng.Count <> 1 Then Exit Function
    Select Case Obj
      Case "Interior":    CellColor = WorksheetFunction.Dec2Hex(Rng.Interior.Color, 6)
      Case "Borders":     CellColor = WorksheetFunction.Dec2Hex(Rng.Borders.Color, 6)
      Case "Font":        CellColor = WorksheetFunction.Dec2Hex(Rng.Font.Color, 6)
    End Select
End Function

'---------- Simple Functions to return the name/ path of workbook/worksheet ---------
Public Function WBName():   WBName = ActiveWorkbook.Name:   End Function

Public Function WBPath():   WBPath = ActiveWorkbook.Path:   End Function

Public Function SHName():   SHName = ActiveSheet.Name:      End Function

'Concatenates all cells in a range and adds a delimiter between each value
Public Function ConcatWDelimiter(Rng As Range, Optional ByVal Delimiter As String = vbNullString)
Attribute ConcatWDelimiter.VB_Description = "This Function Concatenates the values of all cells in the selected range with the specified Delimiter"
Attribute ConcatWDelimiter.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim Cell As Range
  For Each Cell In Rng
    ConcatWDelimiter = ConcatWDelimiter & Cell.Text & Delimiter
  Next Cell
  ConcatWDelimiter = Left(ConcatWDelimiter, Len(ConcatWDelimiter) - Len(Delimiter))
End Function

Public Function ColLetter(ByRef Rng As Range) As String:
Attribute ColLetter.VB_Description = "This Function Returns the Column Letter of the Leftmost cell in the Specified Range"
Attribute ColLetter.VB_ProcData.VB_Invoke_Func = " \n14"
  ColLetter = Left(Rng.Cells(1, 1).Address(True, False), InStr(1, Rng.Cells(1, 1).Address(True, False), "$", 1) - 1)
End Function

Function VLookupArr(ByRef lookup_value As Range, ByRef tbl As Range, ByVal col_index_num As Long, Optional ByVal Vertical As Boolean = True)
Attribute VLookupArr.VB_Description = "[ARRAY FUNCTION USE CTRL+SHIFT+ENTER] Works like a regular Vlookup except it can return multiple results into multiple cells"
Attribute VLookupArr.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim R As Long, Max As Long, Temp() As Variant: ReDim Temp(0)
  Max = WorksheetFunction.Min(tbl.Rows.Count, tbl.Parent.UsedRange.Rows.Count)
  For R = 1 To Max
    If lookup_value.Value = tbl.Cells(R, 1) Then Temp(UBound(Temp)) = tbl.Cells(R, col_index_num).Value: ReDim Preserve Temp(UBound(Temp) + 1)
  Next R
  For R = UBound(Temp) To Range(Application.Caller.Address).Rows.Count: Temp(UBound(Temp)) = "": ReDim Preserve Temp(UBound(Temp) + 1):   Next R
  ReDim Preserve Temp(UBound(Temp) - 1): If Vertical Then VLookupArr = Application.Transpose(Temp) Else VLookupArr = Temp
End Function

Function VLookupAll(ByVal lookup_value As String, _
                    ByRef tbl As Range, _
                    ByVal col_index_num As Long, _
           Optional ByVal Seperator As String = ", ") As String
           
  Dim I As Long, Max As Long, Result As String
  Max = WorksheetFunction.Min(tbl.Rows.Count, tbl.Parent.UsedRange.Rows.Count)
  For I = 1 To Max
    If tbl.Cells(I, 1).Text = lookup_value Then
      Result = Result & (tbl.Cells(I, col_index_num).Text & Seperator)
    End If
  Next I
  If Len(Result) <> 0 Then
    VLookupAll = Left(Result, Len(Result) - Len(Seperator))
  End If
    
End Function


Public Function Regex(ByVal Str As String, ByVal pattern As String, Optional IgnoreCase As Boolean = True) As Long
Attribute Regex.VB_Description = "Evaluates a string using a specified regular expression, it returns the number of matches found within the search string."
Attribute Regex.VB_ProcData.VB_Invoke_Func = " \n14"
  With CreateObject("vbscript.regexp")
    .pattern = pattern
    .IgnoreCase = IgnoreCase
    .Global = True
    Regex = .Execute(Str).Count
  End With
End Function

Public Function RegexReplace(ByVal Str As String, ByVal pattern As String, ByVal NewVal As String, Optional IgnoreCase As Boolean = True) As String
  With CreateObject("vbscript.regexp")
    .pattern = pattern
    .IgnoreCase = IgnoreCase
    .Global = True
    RegexReplace = .Replace(Str, NewVal)
  End With
End Function

'Public Function MaxIf(ByRef SearchRange As Range, ByVal Target As Variant, ByRef ValRange As Range) As Variant
'  Dim I As Long, R As Long, Val As Variant, Max As Variant, Found As Boolean
'  R = WorksheetFunction.Min(SearchRange.Rows.Count, SearchRange.Parent.UsedRange.Rows.Count)
'  For I = 1 To R
'    If SearchRange(I, 1).Value2 = Target Then
'      Val = ValRange(I, 1).Value2
'      If (Found = False) Or (Val > Max) Then Max = Val: Found = True
'    End If
'  Next I
'  If Found Then MaxIf = Max Else MaxIf = CVErr(xlErrNA)
'End Function
'
'Public Function MinIf(ByRef SearchRange As Range, ByVal Target As Variant, ByRef ValRange As Range) As Variant
'  Dim I As Long, R As Long, Val As Variant, Min As Variant, Found As Boolean
'  R = WorksheetFunction.Min(SearchRange.Rows.Count, SearchRange.Parent.UsedRange.Rows.Count)
'  For I = 1 To R
'    If SearchRange(I, 1).Value2 = Target Then
'      Val = ValRange(I, 1).Value2
'      If (Found = False) Or (Val < Min) Then Min = Val: Found = True
'    End If
'  Next I
'  If Found Then MinIf = Min Else MinIf = CVErr(xlErrNA)
'End Function

Public Function Dictlookup(lookupRange As Range, refRange As Range, retRange As Range) As Variant
  Dim dict As Scripting.Dictionary
  Dim myRow As Range
  Dim I As Long, J As Long
  Dim vResults() As Variant

  ' 1. Build a dictionnary
  Set dict = New Scripting.Dictionary
  For Each myRow In refRange.Cells
    ' Append A : B to dictionnary
    dict.Add myRow.Value, retRange.Value
  Next myRow

  ' 2. Use it over all lookup data
  ReDim vResults(1 To lookupRange.Rows.Count, 1 To lookupRange.Columns.Count) As Variant
  For I = 1 To lookupRange.Rows.Count
    For J = 1 To lookupRange.Columns.Count
      If dict.Exists(lookupRange.Cells(I, J).Value) Then
        vResults(I, J) = dict(lookupRange.Cells(I, J).Value)
      End If
    Next J
  Next I

  Dictlookup = vResults
End Function

Private Sub DescribeFunction()
   Dim FuncName As String, FuncDesc As String, Category As String, ArgDesc(1 To 2) As String

   FuncName = "ConcatWDelimiter"
   FuncDesc = "This Function Concatenates the values of all cells in the selected range with the specified Delimiter"
   Category = 14
   ArgDesc(1) = "Range to Concatenate Values"
   ArgDesc(2) = "(Optional) Delimiter Between Values if omitted there will be no space between values"
   Application.MacroOptions _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:=Category, _
      ArgumentDescriptions:=ArgDesc
End Sub

Public Function ConvertToXML(ByRef InputRange As Range, _
                    Optional ByVal IncludeLineBreaks As Boolean = False, _
                    Optional ByVal DeclarationString As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no"" ?>") As String
' Takes a table of data and converts it to XML
' The table must have headers in the first row
  Dim R As Long, C As Long, I As Long, J As Long
  Dim Rng() As Variant, strOut As ClsFastString
  Set strOut = New ClsFastString
  Rng = InputRange:  R = InputRange.Rows.Count:  C = InputRange.Columns.Count
  strOut.Add DeclarationString & IIf(IncludeLineBreaks, vbNewLine, vbNullString)
  For I = 2 To R
    strOut.Add "<ListItem>" & IIf(IncludeLineBreaks, vbNewLine, vbNullString)
    For J = 1 To C
      strOut.Add "<" & Rng(1, J) & ">" & Rng(I, J) & "</" & Rng(1, J) & ">" & IIf(IncludeLineBreaks, vbNewLine, vbNullString)
    Next J
    strOut.Add "</ListItem>" & IIf(IncludeLineBreaks And I <> R, vbNewLine, vbNullString)
  Next I
  ConvertToXML = strOut.Value
  Set strOut = Nothing
End Function

Public Function ConvertToJSON(ByRef InputRange As Range, _
                     Optional ByVal IncludeLineBreaks As Boolean = False) As String
' Takes a table of data and converts it to JSON
' The table must have headers in the first row
  Dim R As Long, C As Long, I As Long, J As Long
  Dim Rng() As Variant, strOut As ClsFastString
  Set strOut = New ClsFastString
  Rng = InputRange:  R = InputRange.Rows.Count:  C = InputRange.Columns.Count
  For I = 2 To R
    strOut.Add "{" & IIf(IncludeLineBreaks, vbNewLine, vbNullString)
    For J = 1 To C
      Select Case VarType(Rng(I, J))
        Case vbString: Call strOut.Add("""" & Rng(1, J) & """:""" & Rng(I, J) & IIf(J = C, """", """,") & IIf(IncludeLineBreaks, vbNewLine, vbNullString))
        Case vbDate:   Call strOut.Add("""" & Rng(1, J) & """:""" & Format(Rng(I, J), "yyyy-mm-ddThh:mm:ss.000") & IIf(J = C, """", """,") & IIf(IncludeLineBreaks, vbNewLine, vbNullString))
        Case Else:     Call strOut.Add("""" & Rng(1, J) & """:" & Rng(I, J) & IIf(J = C, vbNullString, ",") & IIf(IncludeLineBreaks, vbNewLine, vbNullString))
      End Select
    Next J
    strOut.Add "}" & IIf(I = R, vbNullString, "," & IIf(IncludeLineBreaks, vbNewLine, vbNullString))
  Next I
  ConvertToJSON = strOut.Value
  Set strOut = Nothing
End Function


