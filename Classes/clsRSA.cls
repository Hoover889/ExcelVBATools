VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' /---------------------------------------------------------------------\
' |                _____    _____                                       |
' |               |  __ \  / ____|   /\                                 |
' |               | |__) || (___    /  \                                |
' |               |  _  /  \___ \  / /\ \                               |
' |    ______     | | \ \  ____) |/ ____ \      _    _                  |
' |   |  ____|    |_|  \_\|_____//_/    \_\    | |  (_)                 |
' |   | |__    _ __    ___  _ __  _   _  _ __  | |_  _   ___   _ __     |
' |   |  __|  | '_ \  / __|| '__|| | | || '_ \ | __|| | / _ \ | '_ \    |
' |   | |____ | | | || (__ | |   | |_| || |_) || |_ | || (_) || | | |   |
' |   |______||_| |_| \___||_|    \__, || .__/  \__||_| \___/ |_| |_|   |
' |                                __/ || |                             |
' |                               |___/ |_|                             |
' |---------------------------------------------------------------------|
' | This is a wrapper class for Microsoft's built-in RSA encryption     |
' | service. It requires a reference to mscorlib.dll or changing the    |
' | constant 'UseLateBinding' to TRUE, but this will impact performance |
' | especially if you are creating many instances of the class.         |
' |---------------------------------------------------------------------|
' |------ Properties / Members -----------------------------------------|
' | clsRSA.MemberName      [ReturnType]      (Read-Only/Get/Set)        |
' |---------------------------------------------------------------------|
' | clsRSA.KeyState        [RSAKeyStates]      (Read-Only)              |
' |   Returns the following values to display if a Key is loaded        |
' |             0 = No Key Loaded   (can't encrypt or decrypt)          |
' |             1 = Public Key Only (can encrypt only)                  |
' |             2 = Private Key     (can encrypt & decrypt)             |
' |                                                                     |
' | clsRSA.Key_Private     [String]            (Read-Only)              |
' |   Returns an XML string containing the currently loaded Private key |
' |   if no key or a Public key are loaded this returns nothing         |
' |                                                                     |
' | clsRSA.Key             [String]            (Get)                    |
' |   Returns an XML string containing the currently loaded Public key  |
' |   if no key is loaded this returns nothing                          |
' |                                                                     |
' | clsRSA.Key             [String]            (Set)                    |
' |   Accepts an XML string containing either a public or private key   |
' |   check RSAKeyState to verify if the Key was loaded correctly       |
' |                                                                     |
' |---------------------------------------------------------------------|
' |------ Methods ------------------------------------------------------|
' | {ReturnType} MethodName(Argument,[OptionalArgument])                |
' |---------------------------------------------------------------------|
' | {void}       clsRSA.GenerateRandomKey()                             |
' |   Generates a new randomized private Key                            |
' |                                                                     |
' | {Boolean}    clsRSA.ImportKeyFromFile([FilePath])                   |
' |   Imports a key from a text file (XML Formatted)                    |
' |   If a file path is not provided then a dialog box will pop up      |
' |   prompting the user to select a file                               |
' |   Returns True if import was successful                             |
' |   use KeyState to verify the type of key loaded                     |
' |                                                                     |
' | {Boolean}    clsRSA.ExportKeyFromFile([KeyType],[FilePath])         |
' |   Exports the Currently loaded key to a text file (XML)             |
' |   The default Key type is Public but can be changed with [KeyType]  |
' |   If a file path is not provided then a dialog box will pop up      |
' |   note that you can not export a private key if a public is loaded  |
' |   Returns True if export was successful                             |
' |                                                                     |
' | {String}     clsRSA.EncryptString(Str)                              |
' |   Encrypts the provided string using the currently loaded Key       |
' |   If No Key is loaded then this function returns a null string      |
' |                                                                     |
' | {String}     clsRSA.DecryptString(Str)                              |
' |   Decrypts the provided string using the current loaded Private Key |
' |   If No Private Key is loaded this function returns a null string   |
' |---------------------------------------------------------------------|
' | Enjoy...                                                            |
' | Created By Ryan Hoover                                              |
' |                                                                     |
' | Feel free to use or modify this code however you like but please    |
' | provide attribution. If you like my work be sure to say thanks.     |
' \---------------------------------------------------------------------/

' /* Begin Code */
Option Explicit

Public Enum RSAKeyStates
  No_Key = 0
  Public_Key = 1
  Private_Key = 2
End Enum
' Constants
#Const UseLateBinding = False
Private Const PublicKeyLength    As Long = &HF3
Private Const PrivateKeyLength   As Long = &H393
Private Const PaddedInputLength  As Long = &H38
Private Const PaddedOutputLength As Long = &H40
' Private Members
#If UseLateBinding Then
  Private ObjRSA      As Object
#Else
  Private ObjRSA      As mscorlib.RSACryptoServiceProvider
#End If
Private P_RSAKeyState As RSAKeyStates

'Class Constructor Initializes RSACryptoServiceProvider and P_RSAKeyState
Private Sub Class_Initialize()
  #If UseLateBinding Then
    Set ObjRSA = CreateObject("RSACryptoServiceProvider")
  #Else
    Set ObjRSA = New mscorlib.RSACryptoServiceProvider
  #End If
  P_RSAKeyState = No_Key
End Sub

'Class Destructor releases resources used by RSACryptoServiceProvider for garbage collection
Private Sub Class_Terminate()
  ObjRSA.Clear
  Set ObjRSA = Nothing
End Sub

'Generates a new randomized private key
Public Sub GenerateRandomKey()
  On Error GoTo ErrHandler
  ObjRSA.Clear
  #If UseLateBinding Then
    Set ObjRSA = CreateObject("RSACryptoServiceProvider")
  #Else
    Set ObjRSA = New mscorlib.RSACryptoServiceProvider
  #End If
  P_RSAKeyState = Private_Key
  Exit Sub
ErrHandler:
  P_RSAKeyState = No_Key
End Sub

'Gets State of currently Loaded Key
Public Property Get KeyState() As RSAKeyStates
  KeyState = P_RSAKeyState
End Property
'If a private key is loaded, this function returns an XML string contining the currently loaded key
Public Property Get Key_Private() As String
  If P_RSAKeyState = Private_Key Then Key_Private = ObjRSA.ToXmlString(True)
End Property

'If any Key is loaded, this function returns an XML string contining the currently loaded key
Public Property Get Key() As String
  If P_RSAKeyState <> No_Key Then Key = ObjRSA.ToXmlString(False)
End Property

'Sets the RSA Key, check the RSAKeyState to check if this function succeeded
Public Property Let Key(Val As String)
  On Error GoTo ErrHandler
  Select Case Len(Val)
    Case PublicKeyLength:  P_RSAKeyState = Public_Key
    Case PrivateKeyLength: P_RSAKeyState = Private_Key
    Case Else:             P_RSAKeyState = No_Key:     Exit Property
  End Select
  ObjRSA.FromXmlString Val
  Exit Property
ErrHandler:
  P_RSAKeyState = No_Key
End Property

'Imports a RSA Ket from a text file
Public Function ImportKeyFromFile(Optional ByVal FilePath As String = vbNullString) As Boolean
  On Error GoTo ErrHandler
  Dim Key   As String
  Dim FileO As Long
  Dim F     As Office.FileDialog
  ImportKeyFromFile = False
  
  If Len(FilePath) = 0 Then
    Set F = Application.FileDialog(msoFileDialogFilePicker)
    With F
      .Title = "Please select a valid RSA Key file"
      .AllowMultiSelect = False
      With .Filters
        .Clear
        .Add "Text files (.txt)", "*.txt", 1
        .Add "All Files", "*.*", 2
      End With
      If Not .Show Then GoTo CleanUp
      FilePath = .SelectedItems(1)
    End With
  End If
  If StrComp(Right(FilePath, 4), ".txt", vbTextCompare) <> 0 Then GoTo CleanUp
  FileO = FreeFile
  Open FilePath For Input As #FileO
  Line Input #FileO, Key
  If StrComp(Left(Key, 13), "<RSAKeyValue>", vbTextCompare) <> 0 Then GoTo CleanUp
  Me.Key = Key
  ImportKeyFromFile = P_RSAKeyState <> No_Key
  
CleanUp:
  On Error Resume Next
  Close #FileO
  Set F = Nothing
  Exit Function
  
ErrHandler:
  ImportKeyFromFile = False
  Resume CleanUp
  
End Function

'Exports a RSA Ket to a text file
Public Function ExportKeyToFile(Optional ByVal KeyType As RSAKeyStates = RSAKeyStates.Public_Key, Optional ByVal FilePath As String = vbNullString) As Boolean
  On Error GoTo ErrHandler
  Dim F        As Office.FileDialog
  Dim Key      As String
  Dim FileO    As Long
  
  ExportKeyToFile = False
  Select Case KeyType
    Case RSAKeyStates.Private_Key: If P_RSAKeyState <> Private_Key Then GoTo CleanUp Else Key = Me.Key_Private
    Case RSAKeyStates.Public_Key:  If P_RSAKeyState = No_Key Then GoTo CleanUp Else Key = Me.Key
    Case Else: GoTo CleanUp
  End Select
  If Len(FilePath) = 0 Then
    Set F = Application.FileDialog(msoFileDialogSaveAs)
    With F
      .Title = "Please choose where to save the RSA Key file"
      .AllowMultiSelect = False
      .InitialFileName = "RSAKey.txt"
      If Not .Show Then GoTo CleanUp
      FilePath = .SelectedItems(1)
    End With
  End If
  If Right(FilePath, 4) <> ".txt" Then GoTo CleanUp
  FileO = FreeFile
  Open FilePath For Output As #FileO
  Write #FileO, Key
  ExportKeyToFile = True
  
CleanUp:
  On Error Resume Next
  Close #FileO
  Set F = Nothing
  Exit Function
  
ErrHandler:
  ExportKeyToFile = False
  Resume CleanUp
  
End Function

'Encrypts a string with the currently loaded Public or Private key
Public Function EncryptString(ByVal Str As String) As String
  On Error GoTo ErrHandler
  Dim InData()  As Byte, OutData() As Byte, L As Long, I As Long, tmpStr As String
  If P_RSAKeyState = No_Key Then GoTo ErrHandler
  L = Len(Str)
  For I = 1 To L Step PaddedInputLength
    InData = Mid(Str, I, PaddedInputLength)
    OutData = ObjRSA.Encrypt(InData, False)
    tmpStr = OutData
    EncryptString = EncryptString & tmpStr
    Erase InData, OutData
  Next I
CleanUp:
  On Error Resume Next
  Erase InData, OutData
  Exit Function
  
ErrHandler:
  EncryptString = vbNullString
  Resume CleanUp
End Function

'Decrypts a string with the currently loaded Private key
Public Function DecryptString(ByVal Str As String) As String
  On Error GoTo ErrHandler
  Dim InData()  As Byte, OutData() As Byte, L As Long, I As Long, tmpStr As String
  If P_RSAKeyState <> Private_Key Then GoTo ErrHandler
  L = Len(Str)
  For I = 1 To L Step PaddedOutputLength
    InData = Mid(Str, I, PaddedOutputLength)
    OutData = ObjRSA.Decrypt(InData, False)
    tmpStr = OutData
    DecryptString = DecryptString & tmpStr
    Erase InData, OutData
  Next I
CleanUp:
  On Error Resume Next
  Erase InData, OutData
  Exit Function
  
ErrHandler:
  DecryptString = vbNullString
  Resume CleanUp
End Function
