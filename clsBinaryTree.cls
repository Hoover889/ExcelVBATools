VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBinaryTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------
'|     ____          _                               _                             |
'|    |  _ \        | |                             | |                            |
'|    | |_) |  __ _ | |  __ _  _ __    ___  ___   __| |                            |
'|    |  _ <  / _` || | / _` || '_ \  / __|/ _ \ / _` |                            |
'|    | |_) || (_| || || (_| || | | || (__|  __/| (_| |                            |
'|    |____/  \__,_||_| \__,_||_| |_| \___|\___| \__,_|                            |
'|     ____   _                              _______                               |
'|    |  _ \ (_)                            |__   __|                              |
'|    | |_) | _  _ __    __ _  _ __  _   _     | | _ __  ___   ___                 |
'|    |  _ < | || '_ \  / _` || '__|| | | |    | || '__|/ _ \ / _ \                |
'|    | |_) || || | | || (_| || |   | |_| |    | || |  |  __/|  __/                |
'|    |____/ |_||_| |_| \__,_||_|    \__, |    |_||_|   \___| \___|                |
'|                                    __/ |                                        |
'|                                   |___/                                         |
'|---------------------------------------------------------------------------------|
'| THIS CLASS IS DESIGNED FOR BOTH KEY & VALUE                                     |
'| Keys can be numeric and can be case-insensitive                                 |
'|   (see CaseSensitiveKeys property)                                              |
'| When passing Index to any function...                                           |
'|   - If Index is a node index, not a key, pass Index as a numeric data type      |
'|   - If Index is a Key, not a node index, pass Index as a string data type       |
'| Node indexes should never be cached.                                            |
'| They can change whenever any node is deleted                                    |
'|---------------------------------------------------------------------------------|
'| By default, all keys are case-sensitive; duplicate keys are not allowed.        |
'| Duplicate keys will raise an error if one is attempted.                         |
'| Steps for allowing duplicate keys(not reccomended):                             |
'|  - You can override this by providing a True value for the optional             |
'|      AllowDuplicateKeys parameter in the Add & ReKey functions                  |
'|  - Should you be willing to allow duplicate keys, you must handle some          |
'|      extra work when calling these functions:                                   |
'|      1. Delete.  Pass the numerical node Index, not the Key                     |
'|      2. ReKey. Pass the numerical node index, not the Key                       |
'|      3. Do not set a Key's value via the NodeKey property                       |
'|  - To get the correct index, when keys can be duplicated:                       |
'|      Index = .Find(Key,0&)                                                      |
'|      If .NodeValue(Index) = criteria Then you have the correct index            |
'|      Else call Find again & again, passing the Index+1 until criteria is found  |
'|      Index = Find(Key, Index + 1)                                               |
'| The CasesensitiveKeys Property will set whether Keys are case-sensitive         |
'|   This can only be set if the tree is empty                                     |
'|                                                                                 |
'|- Operations --------------------------------------------------------------------|
'|                                                                                 |
'| .Add(Key,Value,AllowDuplicates) As Boolean                                      |
'|    Returns true or false depending on success                                   |
'|    Key             = String Containing the Key for the new Node                 |
'|    Value           = Long Containing the value of the node                      |
'|    AllowDuplicates = Whether Duplicate keys Allowed (Optional)                  |
'|                                                                                 |
'| .Find(Key,FromIndex) As Long                                                    |
'|    Returns the Index of the node being searched for                             |
'|    Key             = String Containing the Key for the Node to Find             |
'|    FromIndex       = Index to start searching from (Optional)                   |
'|                                                                                 |
'| .ReKey(Index,NewKey,AllowDuplicates) As Long                                    |
'|    Returns the New Index of the node being Re-Keyed                             |
'|    Index           = Key or index Number of node to Re-Key                      |
'|    NewKey          = String of New Key                                          |
'|    AllowDuplicates = Whether Duplicate keys Allowed (Optional)                  |
'|                                                                                 |
'| .NodeChild(Index,Child) As Long                                                 |
'|    Returns the Index of one of a node's children                                |
'|    Index           = Key or index Number of node                                |
'|    Child           = True for Left False for Right child                        |
'|                                                                                 |
'| .Delete(Index) As Boolean                                                       |
'|    Returns true or false depending on success                                   |
'|    Index           = Key or index Number of node                                |
'|                                                                                 |
'| .Clear                                                                          |
'|    Erases the Tree                                                              |
'|                                                                                 |
'|- Properties --------------------------------------------------------------------|
'|                                                                                 |
'| .NodeValue(Index) As Long                                                       |
'|    Returns the Value of the specified Node                                      |
'| .NodeValue(Index,NewValue)                                                      |
'|    Sets the Value of the specified Node                                         |
'|                                                                                 |
'| .NodeKey(Index) As String                                                       |
'|    Returns the Key of the specified Node                                        |
'| .NodeKey(Index,NewKey)                                                          |
'|    Sets the Key of the specified Node (No duplicates, Use ReKey for duplicates) |
'|                                                                                 |
'| .NodeBalance(Index) as Long                                                     |
'|    Returns the Balance of a particular Node (Read Only)                         |
'|                                                                                 |
'| .NodeRootIndex As Long                                                          |
'|    Returns the Index of the Root Node (Read Only)                               |
'|                                                                                 |
'| .NodeCount As Long                                                              |
'|    Returns the number of Nodes in the tree (Read Only)                          |
'|                                                                                 |
'|- Traversal ---------------------------------------------------------------------|
'|  The Tree Supports 4 methods of traversal each returning a collection           |
'|     The Output collection is of the Tree Values indexed by Key                  |
'|     .InOrderAsc  - In Order Ascending                                           |
'|     .InOrderDesc - In Order Descending                                          |
'|     .PreOrder    - Pre-Order                                                    |
'|     .PostOrder   - Post-Order                                                   |
'|---------------------------------------------------------------------------------|
'| Enjoy...                                                                        |
'| Created By Ryan Hoover                                                          |
'|            Ryan.Hoover@LOreal.com                                               |
'|                                                                                 |
'| Feel free to use or modify this code  but please provide attribution            |
'| If you like my work be sure to say thanks.                                      |
'-----------------------------------------------------------------------------------

Private Type BTreeNode
  LeftChild As Long
  RightChild As Long
  Balance As Long
  Value As Long
  Key As String
End Type
Private Type BTreeKeyStruct
  Count As Long
  Node() As BTreeNode
End Type
Private BTreeKey         As BTreeKeyStruct
Private KeyCompareBinary As Boolean
Private TraversalResults As Collection

Public Property Let CaseSensitiveKeys(ByVal NewValue As Boolean)
  If NewValue Xor (KeyCompareBinary = vbBinaryCompare) Then
    If BTreeKey.Count = 0& Then
      If NewValue Then KeyCompareBinary = vbBinaryCompare Else KeyCompareBinary = vbTextCompare
    End If
  End If
End Property

Public Property Get CaseSensitiveKeys() As Boolean
  CaseSensitiveKeys = (KeyCompareBinary = vbBinaryCompare)
End Property

Public Sub Clear()
  Erase BTreeKey.Node()
  BTreeKey.Count = 0
End Sub

Public Function Add(ByVal Key As String, ByVal Value As Long, Optional ByVal AllowDuplicateKeys As Boolean = False) As Boolean
  If Key = vbNullString Then Exit Function
  If BTreeKey.Count = 0 Then
    Add = (CreateNode(0, True, Key, Value) > 0)
  Else
    Select Case LocalInsert(0&, Key, Value, AllowDuplicateKeys)
      Case -1:    Err.Raise 5, "cBinaryTree.Add", "System Error Preventing Tree Modification"
      Case 4:     Err.Raise 457, "cBinaryTree.Add"
      Case Else:  Add = True
    End Select
  End If
End Function

Public Function Find(ByVal Key As String, Optional ByVal FromIndex As Long = 0&) As Long
  Find = ValidateIndex(Key, FromIndex)
End Function

Public Function ReKey(ByVal Index As Variant, _
                      ByVal NewKey As String, _
             Optional ByVal AllowDuplicateKeys As Boolean = False) As Long
             
  Dim Value As Long, nodeIndex As Long
  If BTreeKey.Count = 0 Then Exit Function
  
  If NewKey = vbNullString Then
    Err.Raise 380, "cBinaryTree.ReKey": Exit Function
  End If
  
  nodeIndex = ValidateIndex(Index)
  If nodeIndex Then
    If StrComp(BTreeKey.Node(nodeIndex).Key, NewKey, KeyCompareBinary) = 0& Then
      ReKey = nodeIndex
    Else
      If Not AllowDuplicateKeys Then
        If LocalFind(NewKey, BTreeKey.Node(0).LeftChild + 0&) Then Err.Raise 457, "cBinaryTree.ReKey": Exit Function
      End If
      Value = BTreeKey.Node(nodeIndex).Value:     LocalDelete vbNullString, nodeIndex
      LocalInsert 0&, NewKey, Value, AllowDuplicateKeys:      ReKey = BTreeKey.Count
    End If
  Else
    Err.Raise 380, "cBinaryTree.ReKey"
  End If
End Function

Public Function NodeChild(ByVal Index As Variant, _
                          ByVal LeftChild As Boolean) As Long
                          
  Dim nodeIndex As Long:  nodeIndex = ValidateIndex(Index)
  If nodeIndex Then
    If LeftChild Then NodeChild = BTreeKey.Node(nodeIndex).LeftChild Else NodeChild = BTreeKey.Node(nodeIndex).RightChild
  End If
End Function
Public Property Get NodeValue(ByVal Index As Variant) As Long
  Dim nodeIndex As Long:  nodeIndex = ValidateIndex(Index)
  If nodeIndex Then NodeValue = BTreeKey.Node(nodeIndex).Value
End Property
Public Property Let NodeValue(ByVal Index As Variant, ByVal NewValue As Long)
  Dim nodeIndex As Long:  nodeIndex = ValidateIndex(Index)
  If nodeIndex Then BTreeKey.Node(nodeIndex).Value = NewValue
End Property
Public Property Get NodeKey(ByVal Index As Variant) As String
  Dim nodeIndex As Long:  nodeIndex = ValidateIndex(Index)
  If nodeIndex Then NodeKey = BTreeKey.Node(nodeIndex).Key
End Property
Public Property Let NodeKey(ByVal Index As Variant, ByVal NewKey As String)
  Dim nodeIndex As Long:  nodeIndex = ValidateIndex(Index)
  If nodeIndex Then Me.ReKey BTreeKey.Node(nodeIndex).Key, NewKey, False
End Property
Public Property Get NodeRootIndex() As Long
  If BTreeKey.Count Then NodeRootIndex = BTreeKey.Node(0).LeftChild
End Property
Public Property Get NodeCount() As Long
  NodeCount = BTreeKey.Count
End Property

Public Property Get NodeBalance(ByVal Index As Variant) As Long
  Dim nodeIndex As Long:  nodeIndex = ValidateIndex(Index)
  If nodeIndex Then NodeBalance = BTreeKey.Node(nodeIndex).Balance
End Property

Public Function Delete(ByVal Index As Variant) As Boolean
  If VarType(Index) = vbString Then
    Delete = LocalDelete(CStr(Index), 0&)
  Else
    Dim nodeIndex As Long:  nodeIndex = ValidateIndex(Index)
    If nodeIndex Then Delete = LocalDelete(vbNullString, nodeIndex)
  End If
End Function

Private Function LocalInsert(ByVal Parent As Long, _
                             ByVal Key As String, _
                             ByVal Value As Long, _
                             ByVal AllowDupKeys As Boolean) As Long
  On Error GoTo EH
  Dim Pivot, Result, dir, insParent, Child, cmpResult As Long
  
  If Parent = 0& Then
    dir = 1:    Pivot = BTreeKey.Node(0).LeftChild
  Else
    cmpResult = StrComp(Key, BTreeKey.Node(Parent).Key, KeyCompareBinary)
    If cmpResult = -1 Then
      dir = 1:    Pivot = BTreeKey.Node(Parent).LeftChild
    ElseIf cmpResult = 1 Then
      dir = 2:    Pivot = BTreeKey.Node(Parent).RightChild
    Else                ' match, allow only if AllowDupKeys is True
      If AllowDupKeys = False Then LocalInsert = 4&:   Exit Function
    End If
  End If
  
  If Pivot = 0& Then
    CreateNode Parent, (dir = 1), Key, Value
    LocalInsert = dir:      Exit Function
  End If
  
  Result = LocalInsert(Pivot, Key, Value, AllowDupKeys)
  
  If Result < 1& Or Result = 4& Then LocalInsert = Result:   Exit Function
  
  If (Result And &HF) = 1 Then ' moving up from the left else right
    BTreeKey.Node(Pivot).Balance = BTreeKey.Node(Pivot).Balance + 1
  Else
    BTreeKey.Node(Pivot).Balance = BTreeKey.Node(Pivot).Balance - 1
  End If
  
  Select Case BTreeKey.Node(Pivot).Balance
    Case 0:     Exit Function ' balanced suBTreeKey, no action needed
    Case 1, -1: LocalInsert = (Result And &HFFF) * &H10 Or dir: Exit Function ' shift Result and add this node's direction
  End Select
  
  With BTreeKey
    If (Result And &HF) = 1 Then insParent = .Node(Pivot).LeftChild Else insParent = .Node(Pivot).RightChild
    
    Select Case (Result And &HFF)
    
      Case &H11: ' left rotation
        .Node(Pivot).LeftChild = .Node(insParent).RightChild:   .Node(insParent).RightChild = Pivot
        If dir = 1 Then .Node(Parent).LeftChild = insParent Else .Node(Parent).RightChild = insParent
        .Node(Pivot).Balance = 0:                               .Node(insParent).Balance = 0
        
      Case &H12: ' right then left rotations
        Child = .Node(insParent).LeftChild: .Node(Pivot).RightChild = .Node(Child).LeftChild
        .Node(insParent).LeftChild = .Node(Child).RightChild
        If dir = 1 Then .Node(Parent).LeftChild = Child Else .Node(Parent).RightChild = Child
        .Node(Child).LeftChild = Pivot: .Node(Child).RightChild = insParent: .Node(Child).Balance = 0
        .Node(insParent).Balance = 0:   .Node(Pivot).Balance = 0:   Result = Result And &HF00
        If Result = &H100 Then
          .Node(insParent).Balance = -1
        ElseIf Result = &H200 Then
          .Node(Pivot).Balance = 1
        End If
        
      Case &H21: ' left then right rotations
        Child = .Node(insParent).RightChild:    .Node(insParent).RightChild = .Node(Child).LeftChild
        .Node(Pivot).LeftChild = .Node(Child).RightChild
        If dir = 1 Then .Node(Parent).LeftChild = Child Else .Node(Parent).RightChild = Child
        .Node(Child).LeftChild = insParent: .Node(Child).RightChild = Pivot:    .Node(Child).Balance = 0
        .Node(insParent).Balance = 0:   .Node(Pivot).Balance = 0:   Result = Result And &HF00
        If Result = &H100 Then
          .Node(Pivot).Balance = -1
        ElseIf Result = &H200 Then
          .Node(insParent).Balance = 1
        End If
        
      Case &H22: ' right rotation
        .Node(Pivot).RightChild = .Node(insParent).LeftChild:   .Node(insParent).LeftChild = Pivot
        If dir = 1 Then .Node(Parent).LeftChild = insParent Else .Node(Parent).RightChild = insParent
        .Node(Pivot).Balance = 0:   .Node(insParent).Balance = 0
    End Select
  End With
EH:
  If Err Then Err.Clear: LocalInsert = -1& ' possible error? Out of Memory maybe
End Function

Private Function LocalDelete(ByVal Key As String, _
                             ByVal Index As Long) As Boolean
                             
  Dim Parent    As Long, Child     As Long, GrandChild As Long
  Dim Successor As Long, pathCount As Long, cmpResult  As Long
  Dim cTraverseDir(1 To 35) As Byte ' used to cache path direction for balance factor updating
  Dim cTraversePtr(0 To 34) As Long ' used to cache node per path step for balance factor updating
  
  If BTreeKey.Count = 0 Then Exit Function      ' nothing to delete
  
  If BTreeKey.Count = 1 Then
    Me.Clear:   LocalDelete = True: Exit Function
  End If
  
  If Key = vbNullString Then Key = BTreeKey.Node(Index).Key
  Child = BTreeKey.Node(0).LeftChild             ' root index
  pathCount = 1&: cTraverseDir(pathCount) = 1 ' path begins by moving left
  
  Do Until Child = 0&                         ' find the node based on passed value
    cmpResult = StrComp(Key, BTreeKey.Node(Child).Key, KeyCompareBinary)
    If cmpResult = 0& Then Exit Do
    cTraversePtr(pathCount) = Child:    Parent = Child: pathCount = pathCount + 1&
    If cmpResult = -1& Then
      cTraverseDir(pathCount) = 1:    Child = BTreeKey.Node(Parent).LeftChild
    Else
      cTraverseDir(pathCount) = 2:    Child = BTreeKey.Node(Parent).RightChild
    End If
  Loop
  
  If Child = 0& Then Exit Function            ' requested value not found in tree; abort
  
  With BTreeKey
    If .Node(Child).RightChild = 0& Then        ' deleted node has no right child: Case 1
      If cTraverseDir(pathCount) = 1 Then     ' deleted node is to Parent's left
        .Node(Parent).LeftChild = .Node(Child).LeftChild
      Else
        .Node(Parent).RightChild = .Node(Child).LeftChild
      End If
    Else                                        ' deleted node has right child, does it's child have a right child?
      GrandChild = .Node(Child).RightChild:   cTraversePtr(pathCount) = Child:    cTraversePtr(0&) = pathCount
      pathCount = pathCount + 1&: cTraverseDir(pathCount) = 2 ' next potential path direction is right
      If .Node(GrandChild).LeftChild = 0& Then ' no left great grandchild: Case 2
      
        .Node(GrandChild).LeftChild = .Node(Child).LeftChild:   .Node(GrandChild).Balance = .Node(Child).Balance
        If cTraverseDir(cTraversePtr(0&)) = 1 Then                ' deleted node is to Parent's left
          .Node(Parent).LeftChild = GrandChild
        Else '
          .Node(Parent).RightChild = GrandChild
        End If
        
        cTraversePtr(cTraversePtr(0&)) = GrandChild ' in the recursion array, replace deleted node too
      Else    ' Case 3 deletion
        Do  ' find successor
          cTraversePtr(pathCount) = GrandChild:   pathCount = pathCount + 1&: cTraverseDir(pathCount) = 1
          Successor = BTreeKey.Node(GrandChild).LeftChild
          If BTreeKey.Node(Successor).LeftChild = 0& Then Exit Do
          GrandChild = Successor
        Loop
        
        .Node(GrandChild).LeftChild = .Node(Successor).RightChild:  .Node(Successor).LeftChild = .Node(Child).LeftChild
        .Node(Successor).RightChild = .Node(Child).RightChild:      .Node(Successor).Balance = .Node(Child).Balance
        
        If cTraverseDir(cTraversePtr(0&)) = 1 Then                ' deleted node is to Parent's left
          .Node(Parent).LeftChild = Successor
        Else
          .Node(Parent).RightChild = Successor
        End If
        
        cTraversePtr(cTraversePtr(0&)) = Successor                ' in the recursion array, replace deleted node too
      End If
      cTraversePtr(0&) = 0&                                         ' reset
    End If
    .Node(0).RightChild = Child:    .Count = .Count - 1&              ' decrement number of nodes in tree
  End With
  For pathCount = pathCount - 1 To 1 Step -1
    If cTraverseDir(pathCount + 1) = 1 Then                    ' deletion came from the left
      BTreeKey.Node(cTraversePtr(pathCount)).Balance = BTreeKey.Node(cTraversePtr(pathCount)).Balance - 1
    Else                                                       ' deletion came from the right
      BTreeKey.Node(cTraversePtr(pathCount)).Balance = BTreeKey.Node(cTraversePtr(pathCount)).Balance + 1
    End If
    
    Select Case BTreeKey.Node(cTraversePtr(pathCount)).Balance
      Case 1, -1: Exit For    ' no lower level was removed; tree still in balance; done
      
      Case 0&                 ' tree may have become unbalanced; keep adjusting balance up the tree
      
      Case Else               ' tree is now unbalanced
        With BTreeKey
          Parent = cTraversePtr(pathCount - 1&):      Child = cTraversePtr(pathCount)
          If cTraverseDir(pathCount + 1) = 1 Then
            GrandChild = .Node(Child).RightChild           ' cache deleted node's right child; will always have one
            If .Node(GrandChild).Balance = 1 Then          ' rebalance from GrandChild's Right to avoid additional rebalancing
              Successor = .Node(GrandChild).LeftChild:    .Node(GrandChild).LeftChild = .Node(Successor).RightChild
              .Node(Successor).RightChild = .Node(Child).RightChild:  .Node(Child).RightChild = .Node(Successor).LeftChild
              .Node(Successor).LeftChild = Child
              Select Case .Node(Successor).Balance       ' fix up the balance factors
                Case 0&:  .Node(Child).Balance = 0&:  .Node(GrandChild).Balance = 0&
                Case 1&:  .Node(Child).Balance = 1&:  .Node(GrandChild).Balance = 0&
                Case -1&: .Node(Child).Balance = 0&:  .Node(Successor).Balance = 0&:  .Node(GrandChild).Balance = -1&
              End Select
              GrandChild = Successor                     ' set for parent assignment & exit comparison
            Else                                           ' rebalance off the GrandChild vs GrandChild's Left
              If .Node(GrandChild).Balance = 0& Then     ' fix up the balance factors
                .Node(GrandChild).Balance = 1&: .Node(Child).Balance = -1
              Else
                .Node(GrandChild).Balance = 0&: .Node(Child).Balance = 0
              End If                                     ' do the shifting
              .Node(Child).RightChild = .Node(GrandChild).LeftChild:  .Node(GrandChild).LeftChild = Child
            End If
          Else                                               ' rebalance from right (opposite from above)
            GrandChild = .Node(Child).LeftChild
            If .Node(GrandChild).Balance = -1 Then
              Successor = .Node(GrandChild).RightChild:   .Node(GrandChild).RightChild = .Node(Successor).LeftChild
              .Node(Successor).LeftChild = .Node(Child).LeftChild:    .Node(Child).LeftChild = .Node(Successor).RightChild
              .Node(Successor).RightChild = Child
              Select Case .Node(Successor).Balance
                Case 0&:  .Node(Child).Balance = 0&:  .Node(GrandChild).Balance = 0&
                Case 1&:  .Node(Child).Balance = -1&: .Node(GrandChild).Balance = 0&
                Case -1&: .Node(Child).Balance = 0&:  .Node(Successor).Balance = 0&:  .Node(GrandChild).Balance = 1&
              End Select
              GrandChild = Successor
            Else
              If .Node(GrandChild).Balance = 0& Then
                .Node(GrandChild).Balance = -1&: .Node(Child).Balance = 1
              Else
                .Node(GrandChild).Balance = 0&:  .Node(Child).Balance = 0
              End If
              .Node(Child).LeftChild = .Node(GrandChild).RightChild:  .Node(GrandChild).RightChild = Child
            End If
          End If
          If cTraverseDir(pathCount) = 1 Then                ' assign parent the shifted node as its child
            .Node(Parent).LeftChild = GrandChild
          Else
            .Node(Parent).RightChild = GrandChild
          End If
        End With
        If BTreeKey.Node(GrandChild).Balance Then Exit For        ' if the shifted node lost a level, keep checking else abort
      End Select
  Next
  On Error Resume Next
  With BTreeKey   ' swap the last node in array with the one just deleted
    If .Node(0).RightChild <= .Count Then   ' need to find the last node in the array first
      LocalFind .Node(.Count + 1).Key, .Node(0).LeftChild + 0&, Parent
      .Node(.Node(0).RightChild) = .Node(.Count + 1)
      If .Node(Parent).RightChild > .Count Then
        .Node(Parent).RightChild = .Node(0).RightChild
      Else
        .Node(Parent).LeftChild = .Node(0).RightChild
      End If
      .Node(.Count + 1).Key = vbNullString
    Else
      .Node(.Node(0).RightChild).Key = vbNullString
    End If
  End With
  If UBound(BTreeKey.Node) > BTreeKey.Count + 20& Then ReDim Preserve BTreeKey.Node(0 To BTreeKey.Count + 5)
  If Err Then Err.Clear
  LocalDelete = True
End Function

Private Function LocalFind(ByVal Key As String, _
                           ByVal FromIndex As Long, _
                  Optional ByVal ParentIndex As Long) As Long
                  
  Dim cmpResult As Long:  ParentIndex = 0&
  Do
    cmpResult = StrComp(Key, BTreeKey.Node(FromIndex).Key, KeyCompareBinary)
    If cmpResult = 0& Then Exit Do
    ParentIndex = FromIndex
    If cmpResult = -1& Then
      FromIndex = BTreeKey.Node(FromIndex).LeftChild
    Else
      FromIndex = BTreeKey.Node(FromIndex).RightChild
    End If
  Loop Until FromIndex = 0&
  LocalFind = FromIndex
End Function


Private Function CreateNode(ByVal Parent As Long, _
                            ByVal LeftChild As Boolean, _
                            ByVal Key As String, _
                            ByVal Value As Long) As Long
                            
  If BTreeKey.Count = 0& Then
    ReDim BTreeKey.Node(0 To 10)
  ElseIf BTreeKey.Count = UBound(BTreeKey.Node) Then
    ReDim Preserve BTreeKey.Node(0 To BTreeKey.Count + 10)
  End If
  BTreeKey.Count = BTreeKey.Count + 1&
  If LeftChild Then
    BTreeKey.Node(Parent).LeftChild = BTreeKey.Count
  Else
    BTreeKey.Node(Parent).RightChild = BTreeKey.Count
  End If
  With BTreeKey.Node(BTreeKey.Count)
    .LeftChild = 0&:    .RightChild = 0&
    .Balance = 0&:  .Value = Value: .Key = Key
  End With
  CreateNode = BTreeKey.Count
End Function

Private Function ValidateIndex(ByVal Index As Variant, _
                      Optional ByVal FromIndex As Long) As Long
                      
  If BTreeKey.Count Then
    If VarType(Index) = vbString Then
      ValidateIndex = LocalFind(CStr(Index), Me.NodeRootIndex, FromIndex)
    Else
      On Error GoTo EH
      If CLng(Index) = 0& Then
        ValidateIndex = BTreeKey.Node(0).LeftChild
      Else
        If Index > 0& And Index <= BTreeKey.Count Then
          ValidateIndex = Index
        Else
           Err.Raise 380
        End If
      End If
    End If
  End If
EH:
  If Err Then Err.Raise Err.Number, "cBinaryTree"
End Function


'Traversal Functions
Public Function InOrderAsc() As Collection
  Set TraversalResults = New Collection
  Call PInOrderAsc(Me.NodeRootIndex)
  Set InOrderAsc = TraversalResults
End Function

Private Sub PInOrderAsc(ByVal Index As Long)
  Dim rNode As Long
  If Not Index Then Exit Sub
  rNode = Me.NodeChild(Index, True)
  Call PInOrderAsc(rNode)
  TraversalResults.Add Me.NodeValue(Index), Me.NodeKey(Index)
  rNode = Me.NodeChild(Index, False)
  Call PInOrderAsc(rNode)
End Sub

Public Function InOrderDesc() As Collection
  Set TraversalResults = New Collection
  Call PInOrderDesc(Me.NodeRootIndex)
  Set InOrderDesc = TraversalResults
End Function

Private Sub PInOrderDesc(ByVal Index As Long)
  Dim rNode As Long
  If Not Index Then Exit Sub
  rNode = Me.NodeChild(Index, False)
  Call PInOrderDesc(rNode)
  TraversalResults.Add Me.NodeValue(Index), Me.NodeKey(Index)
  rNode = Me.NodeChild(Index, True)
  Call PInOrderDesc(rNode)
End Sub

Public Function PreOrder() As Collection
  Set TraversalResults = New Collection
  Call PPreOrder(Me.NodeRootIndex)
  Set PreOrder = TraversalResults
End Function

Private Sub PPreOrder(ByVal Index As Long)
  Dim rNode As Long
  If Not Index Then Exit Sub
  TraversalResults.Add Me.NodeValue(Index), Me.NodeKey(Index)
  rNode = Me.NodeChild(Index, True)
  Call PPreOrder(rNode)
  rNode = Me.NodeChild(Index, False)
  Call PPreOrder(rNode)
End Sub
    
Public Function PostOrder() As Collection
  Set TraversalResults = New Collection
  Call PPostOrder(Me.NodeRootIndex)
  Set PostOrder = TraversalResults
End Function


Private Sub PPostOrder(ByVal Index As Long)
  Dim rNode As Long
  If Not Index Then Exit Sub
  rNode = Me.NodeChild(Index, True)
  Call PPostOrder(rNode)
  rNode = Me.NodeChild(Index, False)
  Call PPostOrder(rNode)
  TraversalResults.Add Me.NodeValue(Index), Me.NodeKey(Index)
End Sub


