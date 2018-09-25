VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private lValue As Long

Public Property Get Count() As Long
Attribute Count.VB_UserMemId = 0
    Count = lValue
    lValue = lValue + 1
End Property

Public Property Let Count(Num As Long)
    lValue = Num
End Property

Public Property Get CS() As Long
    CS = lValue
End Property
