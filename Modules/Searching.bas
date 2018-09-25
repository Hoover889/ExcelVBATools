Attribute VB_Name = "Searching"
Option Explicit

'/----------------------------------------------------------------------\
'|-------------------------- Search Functions --------------------------|
'|----------------------------------------------------------------------|
'|  The following functions search Through an array of type <Variant>   |
'|  to find a value of type <Variant>. If you wish to use any of these  |
'|  algoritms in an application I suggest changing <Variant> to the     |
'|  data type being searched for to significantly improve performance.  |
'|  All of these functions will follow the format:                      |
'|      int FunctionName( <T>* Arr, <T> Val, int L, int U)              |
'|  All algorithms will Raise an error (Err.Number 13)                  |
'\----------------------------------------------------------------------/


'/----------------------------------------------------------------------\
'|--------------------------- Unsorted Search --------------------------|
'|----------------------------------------------------------------------|
'| without any heuristics (as in an unsorted array) the only searching  |
'| method is brute force, going through every item in the list. This    |
'| method has a worst case efficiency of O(n) with an average case of   |
'| O(n/2) and a best case of O(1), with N being the number of items.    |
'\----------------------------------------------------------------------/

Public Function LinearSearch(ByRef Arr() As Variant, ByVal Val As Variant, ByVal L As Long, ByVal U As Long) As Long
  For LinearSearch = U To L Step -1
    If Arr(LinearSearch) = Val Then Exit Function
  Next LinearSearch
  Err.Raise 13
End Function


'/----------------------------------------------------------------------\
'|---------------------------- Sorted Search ---------------------------|
'|----------------------------------------------------------------------|
'| If an array is to be searched many times it is often beneficial to   |
'| sort it first, a sorted array allows several methods to be used that |
'| have much better average efficiency than a linear brute force search.|
'\----------------------------------------------------------------------/



'/---------------------- Binary Search (Recursive) ---------------------\
'| A binary search divides the array in half with every comparison,     |
'| this results in an average and worst case scenario of O(log(n))      |
'\----------------------------------------------------------------------/
Public Function BinarySearchR(ByRef Arr() As Variant, ByVal Val As Variant, ByVal L As Long, ByVal U As Long) As Long
  Dim P As Long
  If L < U Then BinarySearchR = LBound(Arr) - 1: Exit Function
  P = (L + U) \ 2
  Select Case Arr(P)
    Case Val:      BinarySearchR = P
    Case Is < Val: BinarySearchR = BinarySearchR(Arr, Val, P + 1, U)
    Case Is > Val: BinarySearchR = BinarySearchR(Arr, Val, L, P - 1)
    Case Else:     Err.Raise 13
  End Select
End Function

'/-------------------- Binary Search (Non-Recursive) -------------------\
'| Same method as above but avoids the overhead of recursion            |
'\----------------------------------------------------------------------/
Public Function BinarySearch(ByRef Arr() As Variant, ByVal Val As Variant, ByVal L As Long, ByVal U As Long) As Long
  Dim P As Long
  BinarySearch = L - 1
  Do
    If L < U Then Exit Function
    P = (L + U) \ 2
    Select Case Arr(P)
      Case Val:       BinarySearch = P: Exit Function
      Case Is < Val:  L = P + 1
      Case Is > Val:  U = P - 1
      Case Else:      Err.Raise 13
    End Select
  Loop
End Function

'/-------------------- Dictionary Search (Recursive) -------------------\
'| A Dictonary search attempts to partition close to where the target   |
'| value would be (Assuming an even distribution). This replicates how  |
'| a human would search a sorted document e.g. if you wanted to find    |
'| the definition of 'Aardvark' you wouldnt start in the middle of the  |
'| dictionary. This algorithm has an efficiency of O(log(log(n))) on    |
'| an evenly distributed list. (only works for numeric types)           |
'\----------------------------------------------------------------------/
Public Function DictionarySearchR(ByRef Arr() As Variant, ByVal Val As Variant, ByVal L As Long, ByVal U As Long) As Long
  Dim P As Long
  If L < U Then DictionarySearchR = LBound(Arr) - 1: Exit Function
  P = CLng((Val - Arr(L)) / (Arr(U) - Arr(L)) * (U - L)) + L
  Select Case Arr(P)
    Case Val:      DictionarySearchR = P
    Case Is < Val: DictionarySearchR = DictionarySearchR(Arr, Val, P + 1, U)
    Case Is > Val: DictionarySearchR = DictionarySearchR(Arr, Val, L, P - 1)
    Case Else:     Err.Raise 13
  End Select
End Function
