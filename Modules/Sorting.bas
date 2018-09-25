Attribute VB_Name = "Sorting"
Option Explicit
'------------------------------------------------------------------------
'--------------------------- Helper Functions ---------------------------
'------------------------------------------------------------------------

Private Sub Swap(ByRef A As Long, ByRef B As Long)
    '------------------ Swap ----------------------------------------------
    '| A Simple subroutine that swaps the values of the 2 inputs:         |
    '| A -> B and B -> A                                                  |
    '| This function is more costly then it should be because of the      |
    '| shortcomings of VBA, in almost every other language a simple       |
    '| operation like this would be recognised by the compiler and made   |
    '| into an XCHG instruction which is very fast, VBA can only interpret|
    '| what you code literally and doesnt give any way to indicate you    |
    '| want to do an XCHG instruction so this function must create a      |
    '| temporary variable and perform 3 assignment operations ( which     |
    '| takes at least 5 times as long as XCHG)                            |
    '----------------------------------------------------------------------
  Dim Temp As Long
  Temp = A
  A = B
  B = Temp
End Sub

Private Sub SwapO(ByRef A As Object, ByRef B As Object)
    '------------------ SwapO ---------------------------------------------
    '| Same as above swap function but it swaps objects                   |
    '----------------------------------------------------------------------
  Dim Temp As Object
  Set Temp = A
  Set A = B
  Set B = Temp
End Sub
Private Function DefaultCompare(A As Object, B As Object) As Boolean
    '------------------ Default Compare -----------------------------------
    '| In the functions that allow custom compare methods this is the     |
    '| default comparison. Because VBA does not support Lambda functions  |
    '| or function pointers, to pass a custom compare method use a string |
    '| containing the name of the compare function. it must follow the    |
    '| format:  bool FunctionName(type,type){};                           |
    '----------------------------------------------------------------------
  DefaultCompare = A < B
End Function

'/----------------------------------------------------------------------\
'|--------------------------- Simple Sorts -----------------------------|
'|----------------------------------------------------------------------|
'| All of these have an average efficiency of O(n^2) with n being the   |
'| number of elements in the array to be sorted. They are generally     |
'| easier to understand then more efficient sorts (merge,quick, etc.)   |
'| they also have lower overhead then the more efficient sorts, making  |
'| them practical for sorting smaller arrays (less than 100 items) and  |
'| even faster than the efficient sorts on arrays smaller than ~10      |
'| the simple sorts also have the advantage of smaller code size and    |
'| little to no extra memory required during the sort.                  |
'\----------------------------------------------------------------------/

Public Sub BubbleSort(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
    '/----------------- Bubble Sort --------------------------------------\
    '| One of the least Efficient sorts there is but easy to understand   |
    '| Start at the beginning of the list, if an item is less than the    |
    '| item before it, swap the values. Keep looping through the entire   |
    '| list until no swaps are made.                                      |
    '\--------------------------------------------------------------------/
  Dim Swapped As Boolean, I As Long
  Do
    Swapped = False
    For I = L + 1 To H
      If Arr(I) < Arr(I - 1) Then
        Swap Arr(I - 1), Arr(I)
        Swapped = True
      End If
    Next I
  Loop While Swapped
End Sub

Public Sub SelectionSort(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
    '/----------------- Selection Sort -----------------------------------\
    '| Generally slower than insertion sort but much faster than bubble.  |
    '| This searches through the entire array to select the lowest value. |
    '| The minimum is placed at the beginning of the list. This process   |
    '| is then repeated; items L+1 to H are searched and the minimum is   |
    '| placed into location L+1. Repeat on items L+2 to H and so on.      |
    '|                                                                    |
    '| Note: In VBA this sort is actually faster than insertion sort      |
    '|       because it minimizes the number of swaps, which are much more|
    '|       expensive then they should be due to the shortcomings of VBA.|
    '\--------------------------------------------------------------------/
    
  Dim I As Long, J As Long, Min As Long
  For I = L To H - 1
    Min = I
    For J = I + 1 To H
      If Arr(J) < Arr(Min) Then Min = J
    Next J
    If Min <> I Then Swap Arr(Min), Arr(I)
  Next I
End Sub

Public Sub InsertionSort(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
    '/----------------- Insertion Sort -----------------------------------\
    '| Widely considered the best of the simple sorts but still has O(n^2)|
    '| efficiency. It works by building a sorted array starting from the  |
    '| front. Starting from the front an item is taken from the unsorted  |
    '| part of the array and inserted into the correct position in the    |
    '| sorted part of the array which is built int he front of the array. |
    '|                                                                    |
    '| Note: In every programming language except VBA this algotrithm is  |
    '|       faster than selection sort, but because swapping values is so|
    '|       slow in VBA this algorithm performs worse.                   |
    '\--------------------------------------------------------------------/
  Dim I As Long, J As Long
  L = L + 1
  For I = L To H
    For J = I To L Step -1
      If Arr(J - 1) < Arr(J) Then Exit For
      Swap Arr(J), Arr(J - 1)
    Next J
  Next I
End Sub

Public Sub InsertionSortB(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
    '/----------------- Insertion Sort (Improved) ------------------------\
    '| A modified vesion of the insertion sort to reduce the number of    |
    '| movements, which are the slowest operation. The Logic is the same  |
    '| but, rather than making many swaps, which each require calling a   |
    '| function which creates a temporary variable each time it is called.|
    '| It saves the item to be inserted in a local varible Temp, it shifts|
    '| the array to accomodate the insertion requiring 1 assignment per   |
    '| shift ( "Arr(J) = Arr(J-1)" ) rather than the swap function which  |
    '| involves 3 assignments ( "A = Temp" , "A = B" , "B = Temp" ).      |
    '| for example take the following Array { 1 , 3 , 4 , 5 , 2 } and     |
    '| attempt to insert 2 into the proper location.                      |
    '| Swap Method:                                                       |
    '| Array                            Next Operation                    |
    '| { 1 , 3 , 4 , 5 , 2 }            Swap 2 and 5    ( 3 assignments ) |
    '| { 1 , 3 , 4 , 2 , 5 }            Swap 2 and 4    ( 3 assignments ) |
    '| { 1 , 3 , 2 , 4 , 5 }            Swap 2 and 3    ( 3 assignments ) |
    '| { 1 , 2 , 3 , 4 , 5 }            Complete: Total Assignments = 9   |
    '| Shift Method:                                                      |
    '| Array                    Temp    Next Operation                    |
    '| { 1 , 3 , 4 , 5 , 2 }    {0}     Store 2 to Temp  ( 1 assignment ) |
    '| { 1 , 3 , 4 , 5 , 2 }    {2}     set 2 equal to 5 ( 1 assignment ) |
    '| { 1 , 3 , 4 , 5 , 5 }    {2}     set 5 equal to 4 ( 1 assignment ) |
    '| { 1 , 3 , 4 , 4 , 5 }    {2}     set 4 equal to 3 ( 1 assignment ) |
    '| { 1 , 3 , 3 , 4 , 5 }    {2}     store Temp in 3  ( 1 assignment ) |
    '| { 1 , 2 , 3 , 4 , 5 }    {2}     Complete: Total Assignments = 5   |
    '\--------------------------------------------------------------------/
  Dim I As Long, J As Long, Temp As Long
  L = L + 1
  For I = L To H
    Temp = Arr(I)
    For J = I To L Step -1
      If Arr(J - 1) < Temp Then Exit For
      Arr(J) = Arr(J - 1)
    Next J
    Arr(J) = Temp
  Next I
End Sub

'/----------------------------------------------------------------------\
'|-------------------------- Intermediate Sorts ------------------------|
'|----------------------------------------------------------------------|
'| Shell Sort stands in its own category, it is extremely similar to the|
'| simple sorts above in how it works but is much more efficient, but it|
'| does not fit in with the efficient sorts as it does not utilize a    |
'| Divide and conquer approach like MergeSort and QuickSort             |
'\----------------------------------------------------------------------/

Public Sub ShellSort(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
    '/----------------- Shell Sort ---------------------------------------\
    '| Shell sort Bridges the gap between the simple sorts and efficient  |
    '| sorts. it fixes the flaw with insertion sort where only adjacent   |
    '| elements can be swapped. Shell sort steps through the array,       |
    '| performing an insertion sort on every nth item. the step interval  |
    '| is gradually reduced until the final pass has a step of 1. the goal|
    '| is that the first few passes get the array nearly sorted and the   |
    '| final pass only needs to make small adjustments. there is no clear |
    '| consensus as to what the optimal step size shoulf be and the       |
    '| performance is extremely variable, depending on array size, step   |
    '| incremant and how the array is shuffled. The example below uses a  |
    '| step of S = 3*S+1 Which results in steps of {1,4,13,40,121,...}    |
    '| This Step size is one of the few that ha been widel studied and    |
    '| results in an efficiency between O(n*(log(n))^2) and O(n^(3/2))    |
    '| which is significantly faster than the simple sorts, but still     |
    '| slower than the efficient sorts that we will review later. In      |
    '| practice Shell sort was widely used as it is faster than the simple|
    '| sorts when sorting more than 30 items and its low overhead made it |
    '| faster than the efficient sorts when sorting < ~7000 elements. In  |
    '| modern hardware it has fallen out of favor because processors are  |
    '| now better at predicting recursive functions and Cache misses cause|
    '| a larger slow-down then extra swaps in memory.                     |
    '----------------------------------------------------------------------
    '| Various step increments and their approximate efficiency           |
    '----------------------------------------------------------------------
    '| General Term (k>1)|      Gaps          |  Complexity (worst case)  |
    '| n/(2^k)           | {n/2,n/4,n/8,}     |     O(n^2)                |
    '| 2^k - 1           | {1,3,7,15,31,63,}  |     O(n^(3/2))            |
    '| 2^p*3^q           | {1,2,3,4,6,8,9,}   |     O(n*log^2(n))         |
    '| (3^k-1)/2         | {1,4,13,40,121,}   |     O(n^(3/2))            |
    '| 4^k*3*2^(k-1)+1   | {1,8,23,77,281,}   |     O(n^(4/3))            |
    '\--------------------------------------------------------------------/
  Dim I As Long, J As Long, S As Long, Temp As Long, Continue As Boolean
  Temp = (H - L) / 9: S = 1
  Do While S <= Temp
    S = 3 * S + 1
  Loop
  Do
    For I = S + L To H
      Temp = Arr(I): J = I
      Continue = J > S
      If Continue Then Continue = Arr(J - S) > Temp
      Do While Continue
        Arr(J) = Arr(J - S): J = J - S
        Continue = J > S
        If Continue Then Continue = Arr(J - S) > Temp
      Loop
      Arr(J) = Temp
    Next I
    S = S \ 3
  Loop While S > 0
End Sub

'/----------------------------------------------------------------------\
'|-------------------------- Efficient Sorts ---------------------------|
'|----------------------------------------------------------------------|
'| These sorts all have an avg case efficiency of O(n*log(n)) making    |
'| them significantly faster than the simple sorts. As a trade-off they |
'| are usually more complex, involve recursion, and have large overhead |
'\----------------------------------------------------------------------/


'/----------------- Merge Sort -----------------------------------------\
'| Merge sort is based on the following facts:                          |
'|    - Combining 2 sorted arrays can be done in linear time O(n)       |
'|    - An array of 1 element is always "sorted".                       |
'| Merge sort is different from the previous sorts in that it does not  |
'| modify the inputted array, but rather makes a new array containing   |
'| the elements of the first sorted. Because of this merge sort         |
'| requires O(n) additional memory for the new array to be stored in.   |
'| because it behaves differently I created a wrapper function to give  |
'| merge sort the same syntax as all the previous sorts. ("MergeSortH") |
'|                                                                      |
'| The process of merge sort breaks the inputted array in half          |
'| recursively until it is broken into n single element arrays and      |
'| combines them in sorted order. The number of splits required is      |
'| O(log(n)) and re-combining is O(n) for an  efficiency of O(n*log(n)) |
'|                                                                      |
'| Regardless of the ordering of the input Mergesort always takes       |
'| O(n*log(n)) time to run, even if the list is already sorted, in      |
'| certain situations it may be worthwile to check if the array is      |
'| already sorted before actually performing the sort.                  |
'\----------------------------------------------------------------------/

Public Sub MergeSortH(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
    '/----------------- Merge Sort Helper --------------------------------\
    '| This is only a wrapper for the real merge sort function which      |
    '| returns a sorted copy of the inputted array this wrapper modifies  |
    '| the inputted array. The copy is still made, but it is deleted after|
    '| the wrapper finishes, allowing you to sort part of an array        |
    '\--------------------------------------------------------------------/
  Dim Sort() As Long, I As Long
  ReDim Sort(H - L)
  For I = L To H
    Sort(I - L) = Arr(I)
  Next I
  Sort = Merge_Sort(Sort)
  For I = L To H
    Arr(I) = Sort(I - L)
  Next I
End Sub

Public Function Merge_Sort(ByRef Arr() As Long) As Long()
    '/----------------- Merge Sort ---------------------------------------\
    '| This is the main merge sort function.                              |
    '\--------------------------------------------------------------------/
  Dim P As Long, S As Long, I As Long
  Dim L() As Long, R() As Long
  S = UBound(Arr)
  If S = 0 Then
    Merge_Sort = Arr: Exit Function
  ElseIf S = 1 Then
    If Arr(0) > Arr(1) Then Swap Arr(0), Arr(1)
    Merge_Sort = Arr: Exit Function
  End If
  P = S \ 2
  ReDim L(P) As Long
  ReDim R(S - P - 1) As Long
  For I = 0 To P
    L(I) = Arr(I)
  Next I
  P = P + 1
  For I = P To S
    R(I - P) = Arr(I)
  Next I
  L = Merge_Sort(L)
  R = Merge_Sort(R)
  Merge_Sort = Combine(L, R)
End Function

Private Function Combine(ByRef Left() As Long, Right() As Long) As Long()
    '/----------------- Combine (Helper for MergeSort)--------------------\
    '| this function performs the re-combining step of the merge sort. It |
    '| takes 2 sorted arrays and combines them into a single sorted array |
    '\--------------------------------------------------------------------/
  Dim Arr() As Long
  Dim I As Long, R As Long, L As Long, UR As Long, UL As Long
  UR = UBound(Right)
  UL = UBound(Left)
  ReDim Arr(UR + UL + 1)
  R = 0:  L = 0
  For I = 0 To UR + UL + 1
    If Right(R) < Left(L) Then
      Arr(I) = Right(R): R = R + 1
      If R > UR Then GoTo AddRemainingLeft
    Else
      Arr(I) = Left(L): L = L + 1
      If L > UL Then GoTo AddRemainingRight
    End If
  Next I
  GoTo ReturnArray
AddRemainingLeft:
  For I = I + 1 To UR + UL + 1
    Arr(I) = Left(L): L = L + 1
  Next I
  GoTo ReturnArray
AddRemainingRight:
  For I = I + 1 To UR + UL + 1
    Arr(I) = Right(R): R = R + 1
  Next I
ReturnArray:
  Combine = Arr
End Function

'/----------------- Heap Sort --------------------------------------------\
'| HeapSort is useful in that a Heap is a common structure found in many  |
'| programs and if a heap has already been implemented then implementing  |
'| HeapSort is trivial. To increase performance you can use the input     |
'| array to contain the heap, but for simplicity the example below uses   |
'| a separate array to store the Heap.                                    |
'| Heapsort has an average performance of O(n*log(n)) but it generally    |
'| takes about 2x as long as quicksort because of the extra movement.     |
'\------------------------------------------------------------------------/
Public Sub HeapSort(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
  Dim I As Long, He As clsHeap
  Set He = New clsHeap
  For I = L To H: He.Insert Arr(I): Next I
  For I = H To L Step -1: Arr(I) = He.Remove: Next I
End Sub

'/----------------- Quick Sort -------------------------------------------\
'| QuickSort takes in an an array, and picks a 'pivot' value in the array,|
'| and moves every value greater than the pivot to the right of it and all|
'| values less than the pivot to the left after this partitioning         |
'| operation is complete the pivot is in its final position. The function |
'| then calls 2 copies of itself to perform the same operation on the     |
'| arrays to the left and right of the pivot.                             |
'|                                                                        |
'| Quicksort has an average run time of O(n*log(n)) with a worst case     |
'| scenario of O(n^2) but this is rare, given this implementations method |
'| for choosing a pivot. Only occurs on data deliberately made to break   |
'| quicksort's algorithm.                                                 |
'\------------------------------------------------------------------------/

Public Sub Quicksort(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
  Dim P As Long
  If H > L Then
    P = QSPartition(Arr, L, H)
    Call Quicksort(Arr, L, P - 1)
    Call Quicksort(Arr, P + 1, H)
  End If
End Sub

Private Function QSPartition(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long) As Long
    '/----------------- Partition (Helper for QuickSort)------------------\
    '| This function takes an array, chooses a partition and moves all    |
    '| elements greater than the partition to the right of it. and returns|
    '| the location of the partition.                                     |
    '\--------------------------------------------------------------------/
  Dim PI As Long, PV As Long, I As Long
  PI = Median(Arr, L, ((H - L) \ 2) + L, H)
  PV = Arr(PI)
  If PI <> H Then Swap Arr(PI), Arr(H)
  QSPartition = L
  For I = L To H - 1
    If Arr(I) <= PV Then
      If I <> QSPartition Then Swap Arr(I), Arr(QSPartition)
      QSPartition = QSPartition + 1
    End If
  Next I
  If H <> QSPartition Then Swap Arr(H), Arr(QSPartition)
End Function

Private Function Median(ByRef Arr() As Long, ByVal A As Long, ByVal B As Long, ByVal C As Long)
    '/----------------- Median (Helper for QuickSort)---------------------\
    '| This takes an array, and 3 indexes. It returs the index that       |
    '| coresponds to the median of the values at the indexes.             |
    '\--------------------------------------------------------------------/
  If Arr(A) > Arr(B) Then
    If Arr(A) > Arr(C) Then
      If Arr(B) > Arr(C) Then Median = B Else Median = C
    Else
      Median = A
    End If
  Else
    If Arr(B) > Arr(C) Then
      If Arr(A) > Arr(C) Then Median = A Else Median = C
    Else
      Median = B
    End If
  End If
End Function


'/----------------- Quick Sort (Repeats) ---------------------------------\
'| This is a variation of quicksort that uses a modified partitioning     |
'| function to give greater efficiency on arrays that have many repeated  |
'| elements, like the Dutch National Flag problem. Rather than placing a  |
'| single item in place as the pivot, it places all values equal to the   |
'| pivot. and returns an upper and lower bound. In general performs ~ 30% |
'| slower than standard Quicksort with no repeated keys but runs in half  |
'| the time of Standard Quicksort when there are many duplicate keys.     |
'| Note: This uses the same Median function as the standard quicksort.    |
'\------------------------------------------------------------------------/

Public Sub QuicksortR(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
  Dim Rt As Long, Le As Long
  If H > L Then
    Call QSPartitionR(Arr, L, H, Le, Rt)
    Call QuicksortR(Arr, L, Le)
    Call QuicksortR(Arr, Rt, H)
  End If
End Sub

Private Function QSPartitionR(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long, ByRef Le As Long, ByRef Rt As Long) As Long
    '/----------------- Partition (3-Way Partition) ----------------------\
    '| This function takes an array, chooses a pivot and moves elements   |
    '| greater than the pivot to the right, moves all elements less than  |
    '| the pivot to the left, elements equal to the pivot are stored in   |
    '| the middle, it returns the upper bound of the left section and the |
    '| lower bound of the right section.                                  |
    '\--------------------------------------------------------------------/
  Dim PI As Long, PV As Long, I As Long
  PI = Median(Arr, L, ((H - L) \ 2) + L, H)
  PV = Arr(PI)
  Le = L - 1: Rt = H + 1: I = L
  Do While I < Rt
    If Arr(I) < PV Then
      Le = Le + 1
      If I <> Le Then Swap Arr(I), Arr(Le)
      I = I + 1
    ElseIf Arr(I) > PV Then
      Rt = Rt - 1
      If I <> Rt Then Swap Arr(I), Arr(Rt)
    Else
      I = I + 1
    End If
  Loop
End Function

'/----------------------------------------------------------------------\
'|-------------------------- Advanced Sorts ----------------------------|
'|----------------------------------------------------------------------|
'| The 2 sorts below are more advanced than the others we have gone     |
'| through already. The first is an example of a hybrid sorting         |
'| algorithm that combines the advantages of multiple sorting algorithms|
'| while (hopefully) having the disadvantages of neither.               |
'| The second is a modification of quicksort that allows you to sort    |
'| objects, it takes in an array, lower bound, and upper bound, like    |
'| the other sorts but it also takes in a function pointer to a function|
'| that compares the 2 objects.                                         |
'\----------------------------------------------------------------------/

Public Sub HybridSort(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
    '/----------------- Hybrid Sort ------------------------------------------\
    '| This sort uses almost the exact same logic as quicksort but if the     |
    '| array is less than 10 items it uses selectionSort instead. this        |
    '| generally results in a 5% increase in speed.                           |
    '\------------------------------------------------------------------------/
  Dim P As Long
  Select Case H - L
    Case Is <= 0:
      'do nothing
    Case Is < 10:
      SelectionSort Arr, L, H
    Case Else:
      P = QSPartition(Arr, L, H)
      Call HybridSort(Arr, L, P - 1)
      Call HybridSort(Arr, P + 1, L)
  End Select
End Sub

    '/----------------- Intro Sort -------------------------------------------\
    '| Short for Introspective Sort, it is simply a quicksort that            |
    '| automatically changes to Merge sort as it approaches quadratic time.   |
    '\------------------------------------------------------------------------/
    
Public Sub IntroSort(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long)
  Call IntroSortH(Arr, L, H, 1, Int(Log(H - L)))
End Sub

Private Sub IntroSortH(ByRef Arr() As Long, ByVal L As Long, ByVal H As Long, ByVal Depth As Long, ByVal MaxDepth As Long)
  Dim P As Long
  If H > L Then
    If Depth >= MaxDepth Then
      Call MergeSortH(Arr, L, H)
    Else
      P = QSPartition(Arr, L, H)
      Call IntroSortH(Arr, L, P - 1, Depth + 1, MaxDepth)
      Call IntroSortH(Arr, P + 1, H, Depth + 1, MaxDepth)
    End If
  End If
End Sub

'/----------------- Quick Sort Objects -----------------------------------\
'| This is a version of the quicksort that sorts an array of objects. It  |
'| also takes in a function pointer to a comparison function. Because VBA |
'| does not have a true function pointer type this uses a string of the   |
'| name of the function to be called to compare objects.                  |
'| the function pointer must be of the following format:                  |
'|     bool (*CompareMethod)(object,object);                              |
'\------------------------------------------------------------------------/
Public Sub QuicksortO(ByRef Arr() As Object, ByVal L As Long, ByVal H As Long, _
                      Optional ByVal CompareMethod As String = "DefaultCompare")
  Dim P As Long
  If H > L Then
    P = QSPartitionO(Arr, L, H, CompareMethod)
    QuicksortO Arr, L, P - 1, CompareMethod
    QuicksortO Arr, P + 1, L, CompareMethod
  End If
End Sub

Private Function QSPartitionO(ByRef Arr() As Object, ByVal L As Long, ByVal H As Long, ByVal CompareMethod As String) As Long
    '------------------ PartitionO (Helper for QuickSortO)-----------------
    '----------------------------------------------------------------------
  Dim PI As Long, I As Long, PV As Object
  PI = MedianO(Arr, L, (H - L \ 2) + L, H, CompareMethod)
  Set PV = Arr(PI)
  SwapO Arr(H), Arr(PI)
  QSPartitionO = L
  For I = L To H - 1
    If Application.Run(CompareMethod, Arr(I), PV) Then
      SwapO Arr(I), Arr(QSPartitionO)
      QSPartitionO = QSPartitionO + 1
    End If
  Next I
  SwapO Arr(H), Arr(QSPartitionO)
End Function

Private Function MedianO(ByRef Arr() As Object, ByVal A As Long, ByVal B As Long, ByVal C As Long, ByVal CompareMethod As String) As Long
    '------------------ MedianO (Helper for QuickSortO)--------------------
    '----------------------------------------------------------------------
  With Application
    If .Run(CompareMethod, Arr(B), Arr(A)) Then
      If .Run(CompareMethod, Arr(C), Arr(A)) Then
        If .Run(CompareMethod, Arr(C), Arr(B)) Then MedianO = C Else MedianO = B
      Else
        MedianO = A
      End If
    Else
      If .Run(CompareMethod, Arr(C), Arr(B)) Then
        If .Run(CompareMethod, Arr(C), Arr(A)) Then MedianO = C Else MedianO = A
      Else
        MedianO = B
      End If
    End If
  End With
End Function
