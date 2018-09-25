Attribute VB_Name = "FileSystemFunctions"
Option Explicit

' This Function opens a file selector dialog to allow the user to select a text file.
' the Function returns true if the user selects a valid file.
' the Function requires a byref string as an input, the file that the user selects is stored in the string variable
Private Function GetTxtFile(ByRef Path As String) As Boolean
  Dim fd As FileDialog
  Set fd = Application.FileDialog(msoFileDialogFilePicker)
  With fd
    .Title = "Select the ZSD13 file"
    .AllowMultiSelect = False
    With .Filters
      .Clear
      .Add "Text Files (.txt)", "*.txt", 1
    End With
    GetTxtFile = .Show
    If GetTxtFile Then Path = .SelectedItems(1)
  End With
End Function

' This function scans a specified directory and returns a list of all files within,
'  you can specify a maximum scan depth as well as a regex pattern to filter for specifically named files.
Public Function ScanDir(ByVal Root As String, _
                        ByVal MaxDepth As Long, _
               Optional ByVal RegexFilter As String = vbNullString) As Collection
               
  Dim FSO As Scripting.FileSystemObject
  Dim FLD As Scripting.Folder
  Dim C   As Collection
  
  Set FSO = New FileSystemObject
  Set FLD = FSO.GetFolder(Root)
  Set C = New Collection
  Call ScanFolder(FLD, MaxDepth, RegexFilter, C)
  Set ScanDir = C
  Set FLD = Nothing
  Set FSO = Nothing
End Function

Private Sub ScanFolder(ByRef FLD As Folder, _
                       ByVal MaxDepth As Long, _
                       ByVal RegexFilter As String, _
                       ByRef C As Collection)
  Dim FIL    As Scripting.File
  Dim SubFld As Scripting.Folder
  For Each FIL In FLD.Files
    If RegexFilter = vbNullString Then
      C.Add FIL.Path
    ElseIf MissingFunctions.Regex(FIL.Name, RegexFilter, True) > 0 Then
      C.Add FIL.Path
    End If
  Next FIL
  If MaxDepth > 0 Then
    For Each SubFld In FLD.SubFolders
      Call ScanFolder(SubFld, MaxDepth - 1, RegexFilter, C)
    Next SubFld
  End If
End Sub

Private Function ExampleDelegate(ByRef FIL As Scripting.File) As Boolean
  ExampleDelegate = True
End Function


' This function scans a specified directory and returns a list of all files within, you can specify a maximum scan depth.
' you can also pass a function pointer to a delegate of the form: (bool(Scripting.File))* which will return true for files you wish to include in the output list
Public Function ScanDirDel(ByVal Root As String, _
                           ByVal MaxDepth As Long, _
                  Optional ByVal VaildateDel As String = "ExampleDelegate") As Collection
               
  Dim FSO As Scripting.FileSystemObject
  Dim FLD As Scripting.Folder
  Dim C   As Collection
  
  Set FSO = New FileSystemObject
  Set FLD = FSO.GetFolder(Root)
  Set C = New Collection
  Call ScanFolderDel(FLD, MaxDepth, VaildateDel, C)
  Set ScanDirDel = C
  Set FLD = Nothing
  Set FSO = Nothing
End Function

Private Sub ScanFolderDel(ByRef FLD As Folder, _
                       ByVal MaxDepth As Long, _
                       ByVal VaildateDel As String, _
                       ByRef C As Collection)
  Dim FIL    As Scripting.File
  Dim SubFld As Scripting.Folder
  For Each FIL In FLD.Files
    If Application.Run(VaildateDel, FIL) Then
      C.Add FIL.Path
    End If
  Next FIL
  If MaxDepth > 0 Then
    For Each SubFld In FLD.SubFolders
      Call ScanFolder(SubFld, MaxDepth - 1, VaildateDel, C)
    Next SubFld
  End If
End Sub


