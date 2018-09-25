VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBitfield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'/-----------------------------------------------\
'|    ____   _  _    ______  _        _      _   |
'|   |  _ \ (_)| |  |  ____|(_)      | |    | |  |
'|   | |_) | _ | |_ | |__    _   ___ | |  __| |  |
'|   |  _ < | || __||  __|  | | / _ \| | / _` |  |
'|   | |_) || || |_ | |     | ||  __/| || (_| |  |
'|   |____/ |_| \__||_|     |_| \___||_| \__,_|  |
'|                                               |
'|-----------------------------------------------|
'| This is a generic Bitfield class, it fits 16  |
'| boolean variables into a single integer       |
'\-----------------------------------------------/
Private BitField As Long
Public Enum BitVal
  B_0 = &H1:    B_1 = &H2:    B_2 = &H4:    B_3 = &H8
  B_4 = &H10:   B_5 = &H20:   B_6 = &H40:   B_7 = &H80
  B_8 = &H100:  B_9 = &H200:  B_A = &H400:  B_B = &H800
  B_C = &H1000: B_D = &H2000: B_E = &H4000: B_F = &H8000
End Enum
' Generic Get / Let for any position
Public Property Get BitValue(ByVal FieldPos As BitVal) As Boolean
  BitValue = (BitField And (Not FieldPos))
End Property
Public Property Let BitValue(ByVal FieldPos As BitVal, ByVal Val As Boolean)
  If Val Then BitField = BitField Or FieldPos Else BitField = BitField And (Not FieldPos)
End Property

' Position Specific Get / Let
Public Property Get Val0() As Boolean:    Val0 = BitValue(BitVal.B_0): End Property
Public Property Let Val0(Val As Boolean): BitValue(B_0) = Val:         End Property
Public Property Get Val1() As Boolean:    Val1 = BitValue(BitVal.B_1): End Property
Public Property Let Val1(Val As Boolean): BitValue(B_1) = Val:         End Property
Public Property Get Val2() As Boolean:    Val2 = BitValue(BitVal.B_2): End Property
Public Property Let Val2(Val As Boolean): BitValue(B_2) = Val:         End Property
Public Property Get Val3() As Boolean:    Val3 = BitValue(BitVal.B_3): End Property
Public Property Let Val3(Val As Boolean): BitValue(B_3) = Val:         End Property
Public Property Get Val4() As Boolean:    Val4 = BitValue(BitVal.B_4): End Property
Public Property Let Val4(Val As Boolean): BitValue(B_4) = Val:         End Property
Public Property Get Val5() As Boolean:    Val5 = BitValue(BitVal.B_5): End Property
Public Property Let Val5(Val As Boolean): BitValue(B_5) = Val:         End Property
Public Property Get Val6() As Boolean:    Val6 = BitValue(BitVal.B_6): End Property
Public Property Let Val6(Val As Boolean): BitValue(B_6) = Val:         End Property
Public Property Get Val7() As Boolean:    Val7 = BitValue(BitVal.B_7): End Property
Public Property Let Val7(Val As Boolean): BitValue(B_7) = Val:         End Property
Public Property Get Val8() As Boolean:    Val8 = BitValue(BitVal.B_8): End Property
Public Property Let Val8(Val As Boolean): BitValue(B_8) = Val:         End Property
Public Property Get Val9() As Boolean:    Val9 = BitValue(BitVal.B_9): End Property
Public Property Let Val9(Val As Boolean): BitValue(B_9) = Val:         End Property
Public Property Get ValA() As Boolean:    ValA = BitValue(BitVal.B_A): End Property
Public Property Let ValA(Val As Boolean): BitValue(B_A) = Val:         End Property
Public Property Get ValB() As Boolean:    ValB = BitValue(BitVal.B_B): End Property
Public Property Let ValB(Val As Boolean): BitValue(B_B) = Val:         End Property
Public Property Get ValC() As Boolean:    ValC = BitValue(BitVal.B_C): End Property
Public Property Let ValC(Val As Boolean): BitValue(B_C) = Val:         End Property
Public Property Get ValD() As Boolean:    ValD = BitValue(BitVal.B_D): End Property
Public Property Let ValD(Val As Boolean): BitValue(B_D) = Val:         End Property
Public Property Get ValE() As Boolean:    ValE = BitValue(BitVal.B_E): End Property
Public Property Let ValE(Val As Boolean): BitValue(B_E) = Val:         End Property
Public Property Get ValF() As Boolean:    ValF = BitValue(BitVal.B_F): End Property
Public Property Let ValF(Val As Boolean): BitValue(B_F) = Val:         End Property
'Public Property Get Val() As Long:        Val = BitField:              End Property
'Public Property Set Val(NewVal As Long):  BitField = NewVal:           End Property
