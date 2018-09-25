VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function apiGetUName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function apiGetCName Lib "kernel32.dll" Alias "GetComputerNameExA" (ByVal NameTypeIndex As Long, ByVal lpBuffer As String, ByRef nSize As Long) As Long
#Else
    Private Declare Function apiGetUName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function apiGetCName Lib "kernel32.dll" Alias "GetComputerNameExA" (ByVal NameTypeIndex As Long, ByVal lpBuffer As String, ByRef nSize As Long) As Long
#End If

Public Property Get UserName() As String
    Dim X, Y As Long, UName As String: UName = Space(255)
    X = apiGetUName(UName, 255):    UserName = vbNullString
    Y = Len(Trim(UName)): If X > 0 Then UserName = Left$(UName, Y - 1)
End Property

Public Property Get UserNameProper() As String
    Dim X, Y As Long, UName As String: UName = Space(255)
    X = apiGetUName(UName, 255): UserNameProper = vbNullString: Y = Len(Trim(UName))
    If X > 0 Then UserNameProper = StrConv(Replace(Left$(UName, Y - 1), ".", " "), vbProperCase)
End Property

Public Property Get ComputerName() As String
    Dim X, Y As Long, CompName As String: CompName = Space(255)
    X = apiGetCName(0, CompName, 255):  ComputerName = vbNullString
    Y = Len(Trim(CompName)): If X > 0 Then ComputerName = Left$(CompName, Y - 1)
End Property

Public Property Get DNSDomain() As String
    Dim X, Y As Long, DNSName As String:  DNSName = Space(255)
    X = apiGetCName(2, DNSName, 255):   DNSDomain = vbNullString
    Y = Len(Trim(DNSName)): If X > 0 Then DNSDomain = Left$(DNSName, Y - 1)
End Property


