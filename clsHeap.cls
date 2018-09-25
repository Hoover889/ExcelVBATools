VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHeap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'|------------------------------------------------------|
'|            _    _                                    |
'|           | |  | |                                   |
'|           | |__| |  ___   __ _  _ __                 |
'|           |  __  | / _ \ / _` || '_ \                |
'|           | |  | ||  __/| (_| || |_) |               |
'|           |_|  |_| \___| \__,_|| .__/                |
'|                                | |                   |
'|                                |_|                   |
'|------------------------------------------------------|
'| A simple Heap implementation translated from C++     |
'| Based off of Sedgewick's "Algorithms in C++"(2nd ed) |
'|------------------------------------------------------|
'| Supports the following operations:                   |
'|  /-----------------------------------------------\   |
'|  |Method: |Input:| Out: |        Action:         |   |
'|  |-----------------------------------------------|   |
'|  |.Insert | Long |[None]| Add new val to Heap    |   |
'|  |.Remove |[None]| Long | Remove Max Val         |   |
'|  |.Replace| Long | Long | Add new Val& return Max|   |
'|  \-----------------------------------------------/   |
'|------------------------------------------------------|
'| This object performs dynnamic memory allocation      |
'| amortized re-allocation time should be O(log(n))     |
'|------------------------------------------------------|
'|               Created By Ryan Hoover                 |
'|------------------------------------------------------|


Option Base 0                                   ' using namespace std;
                                                ' Class Heap {
                                                '     private:
Private N As Long                               '     int N;
Private A() As Long                             '     <Typ>* A;

Private U As Long
'// VBA does not support parameters in constructors so this heap starts at 16 items and resizes dynamically
Private Sub Class_Initialize()                  '     Heap(int Max)
  U = 16: N = 0                                 '     { N = 0;
  ReDim A(0 To U)                               '       A = new <Typ>[Max];
End Sub                                         '     }

Private Sub Class_Terminate()                   '     ~Heap()
  Erase A()                                     '     { delete A;
End Sub                                         '     }

Private Sub Resize()
  If U - N < 1 Then U = U + U: ReDim Preserve A(0 To U)
End Sub

Private Sub UpHeap(ByVal K As Long)             '     void Heap::UpHeap(int K)
  Dim V As Long                                 '     { <Typ> V;
  V = A(K)                                      '       V = A[K];
  Do While A(K \ 2) <= V And K > 0              '       while( A[K>>1] <= V) {
    A(K) = A(K \ 2): K = K \ 2                  '         A[K] = A[K>>1]; K = K>>1;
  Loop                                          '       };
  A(K) = V                                      '       A[K] = V;
End Sub                                         '     {

Private Sub DownHeap(ByVal K As Long)           '     void Heap::DownHeap(int K)
  Dim J As Long, V As Long                      '     { int J; <Typ> V;
  V = A(K)                                      '       V = A[K];
  Do While K <= N \ 2                           '       while (K <= N/2) {
    J = K + K                                   '         J = K<<1;
    If N - J > 0 Then                           '         /* Covered in line below */
      If A(J) < A(J + 1) Then J = J + 1         '         if (J<N && A[J] < A[J+1]) J++;
    End If                                      '          /* Covered in line above */
    If V >= A(J) Then Exit Do                   '         if (V >= A[J]) break;
    A(K) = A(J): K = J                          '         A[K] = A[J]; K = J;
  Loop                                          '        }
  A(K) = V                                      '       A[K] = V;
End Sub                                         '     }

                                                '     public:
Public Sub Insert(ByVal V As Long)              '     void Heap::Insert(<Typ> V)
  Call Resize                                   '     { /* No C++ Equivalent */
  N = N + 1: A(N) = V                           '       A[++N] = V;
  Call UpHeap(N)                                '       Heap::UpHeap(N);
End Sub                                         '     }

Public Function Remove() As Long                '     <Typ> Heap::Remove()
  Remove = A(1)                                 '     { <Typ> V = A[1];
  A(1) = A(N): N = N - 1                        '       A[1] = A[N--];
  Call DownHeap(1)                              '       Heap::DownHeap(1); return V;
End Function                                    '     }

Public Function Replace(ByVal V As Long) As Long '    int Heap::Replace(<Typ> V)
  A(0) = V                                      '     { A[0] = V;
  Call DownHeap(0)                              '       Heap::DownHeap(0);
  Replace = A(0)                                '       return A[0];
End Function                                    '     }
                                                '}

