VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFastString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------
'|      ______           _                                                         |
'|     |  ____|         | |                                                        |
'|     | |__  __ _  ___ | |_                                                       |
'|     |  __|/ _` |/ __|| __|                                                      |
'|     | |  | (_| |\__ \| |_                                                       |
'|     |_|   \__,_||___/ \__|                                                      |
'|       _____  _          _                                                       |
'|      / ____|| |        (_)                                                      |
'|     | (___  | |_  _ __  _  _ __    __ _                                         |
'|      \___ \ | __|| '__|| || '_ \  / _` |                                        |
'|      ____) || |_ | |   | || | | || (_| |                                        |
'|     |_____/  \__||_|   |_||_| |_| \__, |                                        |
'|                                    __/ |                                        |
'|                                   |___/                                         |
'|                                                                                 |
'| A More Efficient String that Pre-allocates memory to make Concatenetion faster  |
'|                                                                                 |
'|---- Intro ----------------------------------------------------------------------|
'| Standard string concatenation operates at an efficiency of O(n^2)               |
'| by pre allocating large blocks of memory (50kb) efficiency becomes O(n log n)   |
'| on large strings like HTML or XML prformance can be improved 20x or more        |
'|                                                                                 |
'| The fast string is an object rather than a basic Data type like standard string |
'| and therefore has more overhead and is not reccomended in situations where      |                                                                  |
'| strings will not exceed 20,000 characters as performance gains will be minimal  |
'|                                                                                 |
'|                                                                                 |
'|---- How To Use -----------------------------------------------------------------|
'| To use in a macro declare an object as ClsFastString and initialize             |
'|                                                                                 |
'|     Dim Str As ClsFastString                                                    |
'|     Set Str = New ClsFaststring                                                 |
'|                                                                                 |
'| To assign or get the string use the Value property                              |
'|                                                                                 |
'|     Str.Value = "Example Text"                                                  |
'|     Debug.Print Str.Value                                                       |
'|                                                                                 |
'| The Len(), Right(), Left(), and Mid() functions all work on Str.value but all   |
'| of these functions are also supported natively by the FastString Class.         |
'|                                                                                 |
'|    |-----------------------------------------|  The Fast String built in        |
'|    |FastString Function |VBA Built-in        |  functions are faster than the   |
'|    |--------------------|--------------------|  VBA built in functions too      |
'|    |Str.GetLeft(10)     |Left$(Str.Value,10) |                                  |
'|    |Str.GetRight(10)    |Right$(Str.Value,10)|                                  |
'|    |Str.GetMid(5,5)     |Mid$(Str.Value,5,5) |                                  |
'|    |Str.Length          |Len(Str.Value)      |                                  |
'|    |-----------------------------------------|                                  |
'|                                                                                 |
'| The biggest advantage of the FastString is the increased concatenation speed.   |
'| To concatenate another string to the FastStringuse the Add Method               |
'|                                                                                 |
'|     Str.Add " More Text"                                                        |
'|                                                                                 |
'| The Add method returns a FastString Object so it can be chained easily          |
'|                                                                                 |
'|     Str.Add(" More Text").Add " Even More Text"                                 |
'|                                                                                 |
'| The Clear method clears out the string faster than normal VBA strings           |
'|                                                                                 |
'|     Str.Clear                                                                   |
'|                                                                                 |
'|---------------------------------------------------------------------------------|
'| Enjoy...                                                                        |
'| Created By Ryan Hoover                                                          |
'|                                                                                 |
'| Feel free to use or modify this code  but please provide attribution            |
'| If you like my work be sure to say thanks.                                      |
'-----------------------------------------------------------------------------------


Private Const SizeIncrement = 50000
Private StrValue As String
Private StrLen As Long

Private Sub Class_Initialize()
  StrValue = Space$(SizeIncrement)
  StrLen = 0
End Sub

Public Function Add(Str As String) As ClsFastString
  Dim L As Long
  L = Len(Str)
  If (StrLen + L) >= Len(StrValue) Then
    If L > SizeIncrement Then StrValue = StrValue & Space$(L) Else StrValue = StrValue & Space$(SizeIncrement)
  End If
  Mid$(StrValue, StrLen + 1, L) = Str
  StrLen = StrLen + L
  Set Add = Me
End Function

Public Property Let Value(ByVal Str As String)
  Dim L As Long
  StrValue = Space$(SizeIncrement)
  StrLen = 0
  L = Len(Str)
  If (StrLen + L) >= Len(StrValue) Then
    If L > SizeIncrement Then StrValue = StrValue & Space$(L) Else StrValue = StrValue & Space$(SizeIncrement)
  End If
  Mid$(StrValue, StrLen + 1, L) = Str
  StrLen = StrLen + L
End Property

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
  Value = vbNullString
  If StrLen Then Value = Left(StrValue, StrLen)
End Property

Public Property Get Length() As Long
  Length = StrLen
End Property

Public Property Get GetLeft(ByVal NumChars As Long) As String
  GetLeft = Left$(StrValue, NumChars)
End Property

Public Property Get GetRight(ByVal NumChars As Long) As String
  GetRight = Mid$(StrValue, StrLen - NumChars, NumChars)
End Property

Public Property Get GetMid(ByVal StartPos As Long, ByVal NumChars As Long) As String
  GetMid = Mid$(StrValue, StartPos, NumChars)
End Property

Public Function Clear() As ClsFastString
  StrLen = 0
  Set Clear = Me
End Function

Public Function Delete(ByVal NumChars As Long) As ClsFastString
  If NumChars > StrLen Then StrLen = 0 Else StrLen = StrLen - NumChars
End Function
