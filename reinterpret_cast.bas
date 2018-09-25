Attribute VB_Name = "reinterpret_cast"
Option Explicit
'/-----------------------------------------\
'| A set of functions that replicate the   |
'| functionality of reinterpret_cast<T>    |
'| and other C pointer tricks in VBA       |
'|                                         |
'|  - C++ Code -                           |
'|  int i = 7;                             |
'|  float f = reinterpret_cast<float>(&i); |
'|                                         |
'|  - C Code   -                           |
'|  int i = 7;                             |
'|  float f = *(float *) &i;               |
'|                                         |
'|  - VBA Code -                           |
'|  Dim i as Long, f as Single             |
'|  i = 7                                  |
'|  f = LongToSingle(i)                    |
'\-----------------------------------------/

'--- 8 Byte UDTs ---
Private Type OneDouble:  Value As Double:          End Type
Private Type OneLongPtr: Value As LongPtr:         End Type
Private Type TwoLongs:   Value(0 To 1) As Long:    End Type
Private Type FourInts:   Value(0 To 3) As Integer: End Type
Private Type EightBytes: Value(0 To 7) As Byte:    End Type

'--- 4 Byte UDTs ---
Private Type OneLong:    Value As Long:            End Type
Private Type OneSingle:  Value As Single:          End Type
Private Type TwoInts:    Value(0 To 1) As Integer: End Type
Private Type FourBytes:  Value(0 To 3) As Byte:    End Type

'--- 2 Byte UDTs ---
Private Type OneInt:     Value As Integer:         End Type
Private Type TwoBytes:   Value(0 To 1) As Byte:    End Type


'--- 2 Byte Variable Conversions ---
'Integer to 2 Bytes conversions
Public Function IntToByteArr(ByVal Value As Integer) As Byte()
  Dim OI As OneInt, TB As TwoBytes: OI.Value = Value: LSet TB = OI: IntToByteArr = TB.Value
End Function
Public Function ByteArrToInt(ByRef Value() As Byte) As Integer
  Dim OI As OneInt, TB As TwoBytes: TB.Value(0) = Value(0): TB.Value(1) = Value(1): LSet OI = TB: ByteArrToInt = OI.Value
End Function
Public Function TwoBytesToInt(ByVal B0 As Byte, ByVal B1 As Byte) As Integer
  Dim OI As OneInt, TB As TwoBytes: TB.Value(0) = B0: TB.Value(1) = B1: LSet OI = TB: TwoBytesToInt = OI.Value
End Function

'--- 4 Byte Variable Conversions ---
'Long to 4 Bytes conversions
Public Function LongToByteArr(ByVal Value As Long) As Byte()
  Dim OL As OneLong, FB As FourBytes: OL.Value = Value: LSet FB = OL: LongToByteArr = FB.Value
End Function
Public Function ByteArrToLong(ByRef Value() As Byte) As Long
  Dim OL As OneLong, FB As FourBytes, I As Long: For I = 0 To 3: FB.Value(I) = Value(I): Next I: LSet OL = FB: ByteArrToLong = OL.Value
End Function
Public Function FourBytesToLong(ByVal B0 As Byte, ByVal B1 As Byte, ByVal B2 As Byte, ByVal B3 As Byte) As Long
  Dim OL As OneLong, FB As FourBytes: FB.Value(0) = B0: FB.Value(1) = B1: FB.Value(2) = B2: FB.Value(3) = B3: LSet OL = FB: FourBytesToLong = OL.Value
End Function

'Single to 4 Bytes conversions
Public Function SingleToByteArr(ByVal Value As Single) As Byte()
  Dim OS As OneSingle, FB As FourBytes: OS.Value = Value: LSet FB = OS: SingleToByteArr = FB.Value
End Function
Public Function ByteArrToSingle(ByRef Value() As Byte) As Single
  Dim OS As OneSingle, FB As FourBytes, I As Long: For I = 0 To 3: FB.Value(I) = Value(I): Next I: LSet OS = FB: ByteArrToSingle = OS.Value
End Function
Public Function FourBytesToSingle(ByVal B0 As Byte, ByVal B1 As Byte, ByVal B2 As Byte, ByVal B3 As Byte) As Single
  Dim OS As OneSingle, FB As FourBytes: FB.Value(0) = B0: FB.Value(1) = B1: FB.Value(2) = B2: FB.Value(3) = B3: LSet OS = FB: FourBytesToSingle = OS.Value
End Function

' long to single conversion for FInvSqrt (simulates casting void pointer between int and float in C++)
Public Function LongToSingle(ByVal Value As Long) As Single
  Dim OL As OneLong, OS As OneSingle: OL.Value = Value: LSet OS = OL: LongToSingle = OS.Value
End Function
Public Function SingleToLong(ByVal Value As Single) As Long
  Dim OL As OneLong, OS As OneSingle: OS.Value = Value: LSet OL = OS: SingleToLong = OL.Value
End Function

' Long to TwoInts
Public Function LongToTwoInts(ByVal Value As Long) As Integer()
  Dim OL As OneLong, TI As TwoInts: OL.Value = Value: LSet TI = OL: LongToTwoInts = TI.Value
End Function
Public Function IntArrToLong(ByRef Value() As Integer) As Long
  Dim OL As OneLong, TI As TwoInts: TI.Value(0) = Value(0): TI.Value(1) = Value(1): LSet OL = TI: IntArrToLong = OL.Value
End Function
Public Function TwoIntsToLong(ByVal I0 As Integer, ByVal I1 As Integer) As Long
  Dim OL As OneLong, TI As TwoInts: TI.Value(0) = I0: TI.Value(1) = I1: LSet OL = TI: TwoIntsToLong = OL.Value
End Function

'--- 8 Byte Variables --
'LongPtr to 8 Bytes conversions
Public Function LongPtrToByteArr(ByVal Value As LongPtr) As Byte()
  Dim OL As OneLongPtr, EB As EightBytes: OL.Value = Value: LSet EB = OL: LongPtrToByteArr = EB.Value
End Function
Public Function ByteArrToLongPtr(ByRef Value() As Byte) As LongPtr
  Dim OL As OneLongPtr, EB As EightBytes, I As Long: For I = 0 To 7: EB.Value(I) = Value(I): Next I: LSet OL = EB: ByteArrToLongPtr = OL.Value
End Function
Public Function EightBytesToLongPtr(ByVal B0 As Byte, ByVal B1 As Byte, ByVal B2 As Byte, ByVal B3 As Byte, _
                                    ByVal B4 As Byte, ByVal B5 As Byte, ByVal B6 As Byte, ByVal B7 As Byte) As LongPtr
  Dim OL As OneLongPtr, EB As EightBytes: EB.Value(0) = B0: EB.Value(1) = B1: EB.Value(2) = B2: EB.Value(3) = B3: EB.Value(4) = B4: EB.Value(5) = B5: EB.Value(6) = B6: EB.Value(7) = B7:
  LSet OL = EB: EightBytesToLongPtr = OL.Value
End Function

'Double to 8 Bytes conversions
Public Function DoubleToByteArr(ByVal Value As Double) As Byte()
  Dim OD As OneDouble, EB As EightBytes: OD.Value = Value: LSet EB = OD: DoubleToByteArr = EB.Value
End Function
Public Function ByteArrToDouble(ByRef Value() As Byte) As Double
  Dim OD As OneDouble, EB As EightBytes, I As Long: For I = 0 To 7: EB.Value(I) = Value(I): Next I: LSet OD = EB: ByteArrToDouble = OD.Value
End Function
Public Function EightBytesToDouble(ByVal B0 As Byte, ByVal B1 As Byte, ByVal B2 As Byte, ByVal B3 As Byte, _
                                   ByVal B4 As Byte, ByVal B5 As Byte, ByVal B6 As Byte, ByVal B7 As Byte) As Double
  Dim OD As OneDouble, EB As EightBytes: EB.Value(0) = B0: EB.Value(1) = B1: EB.Value(2) = B2: EB.Value(3) = B3:
  EB.Value(4) = B4: EB.Value(5) = B5: EB.Value(6) = B6: EB.Value(7) = B7:
  LSet OD = EB: EightBytesToDouble = OD.Value
End Function

' longptr to Double conversion for FInvSqrt (simulates casting void pointer between int and float in C++)
Public Function LongPtrToDouble(ByVal Value As LongPtr) As Double
  Dim OL As OneLongPtr, OD As OneDouble: OL.Value = Value: LSet OD = OL: LongPtrToDouble = OD.Value
End Function
Public Function DoubleToLongPtr(ByVal Value As Double) As LongPtr
  Dim OL As OneLongPtr, OD As OneDouble: OD.Value = Value: LSet OL = OD: DoubleToLongPtr = OL.Value
End Function

' functions to convert data types to strings containing binary representations of their values
Public Function DoubleToBinaryString(ByVal Value As Double, _
                            Optional ByVal AddSpaces As Boolean = False, _
                            Optional ByVal BigEndian As Boolean = False) As String
  Dim OD As OneDouble, EB As EightBytes, I As Long: DoubleToBinaryString = Space$(64 + IIf(AddSpaces, 7, 0))
  OD.Value = Value:  LSet EB = OD
  For I = 0 To 7
    Mid$(DoubleToBinaryString, IIf(AddSpaces, 9, 8) * I + 1, 8) = WorksheetFunction.Dec2Bin(EB.Value(IIf(BigEndian, 7 - I, I)), 8)
  Next I
End Function

Public Function SingleToBinaryString(ByVal Value As Single, _
                            Optional ByVal AddSpaces As Boolean = False, _
                            Optional ByVal BigEndian As Boolean = False) As String
  Dim OS As OneSingle, FB As FourBytes, I As Long: SingleToBinaryString = Space$(32 + IIf(AddSpaces, 3, 0))
  OS.Value = Value:  LSet FB = OS
  For I = 0 To 3
    Mid$(SingleToBinaryString, IIf(AddSpaces, 9, 8) * I + 1, 8) = WorksheetFunction.Dec2Bin(FB.Value(IIf(BigEndian, 3 - I, I)), 8)
  Next I
End Function

Public Function LongToBinaryString(ByVal Value As Long, _
                          Optional ByVal AddSpaces As Boolean = False, _
                          Optional ByVal BigEndian As Boolean = False) As String
  Dim OL As OneLong, FB As FourBytes, I As Long: LongToBinaryString = Space$(32 + IIf(AddSpaces, 3, 0))
  OL.Value = Value:  LSet FB = OL
  For I = 0 To 3
    Mid$(LongToBinaryString, IIf(AddSpaces, 9, 8) * I + 1, 8) = WorksheetFunction.Dec2Bin(FB.Value(IIf(BigEndian, 3 - I, I)), 8)
  Next I
End Function

