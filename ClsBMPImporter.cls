VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsBMPImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------
'|                                                                                 |
'|    ____   _  _                                                                  |
'|   |  _ \ (_)| |                                                                 |
'|   | |_) | _ | |_  _ __ ___    __ _  _ __                                        |
'|   |  _ < | || __|| '_ ` _ \  / _` || '_ \                                       |
'|   | |_) || || |_ | | | | | || (_| || |_) |                                      |
'|   |____/ |_| \__||_| |_| |_| \__,_|| .__/                                       |
'|    _____                           | |   _                                      |
'|   |_   _|                          |_|  | |                                     |
'|     | |   _ __ ___   _ __    ___   _ __ | |_  ___  _ __                         |
'|     | |  | '_ ` _ \ | '_ \  / _ \ | '__|| __|/ _ \| '__|                        |
'|    _| |_ | | | | | || |_) || (_) || |   | |_|  __/| |                           |
'|   |_____||_| |_| |_|| .__/  \___/ |_|    \__|\___||_|                           |
'|                     | |                                                         |
'|                     |_|                                                         |
'|                                                                                 |
'|---------------------------------------------------------------------------------|
'|                                                                                 |
'|   TODO: Add Documentation                                                       |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|                                                                                 |
'|---------------------------------------------------------------------------------|
'| Enjoy...                                                                        |
'| Created By Ryan Hoover                                                          |
'|            Ryan.Hoover@LOreal.com                                               |
'|                                                                                 |
'| Feel free to use or modify this code  but please provide attribution            |
'| If you like my work be sure to say thanks.                                      |
'-----------------------------------------------------------------------------------


' 14 Byte Header
Private Type BITMAPFILEHEADER
  strFileType     As String * 2 ' 2 Bytes
  lngFileSize     As Long       ' 4 Bytes
  bytReserved1    As Integer    ' 2 Bytes
  bytResrved2     As Integer    ' 2 Bytes
  lngBitmapOffset As Long       ' 4 Bytes
End Type

' 40 Byte Header
Private Type BITMAPINFOHEADER
  lngSize          As Long      ' 4 Bytes
  lngWidth         As Long      ' 4 Bytes
  lngHeight        As Long      ' 4 Bytes
  lngPlanes        As Integer   ' 2 Bytes
  intBitCount      As Integer   ' 2 Bytes
  lngCompression   As Long      ' 4 Bytes
  lngSizeImage     As Long      ' 4 Bytes
  lngXPelsPerMeter As Long      ' 4 Bytes
  lngYPelsPerMeter As Long      ' 4 Bytes
  lngClrUsed       As Long      ' 4 Bytes
  lngClrImportant  As Long
End Type

' Palette Type for 16Bit
Private Type PALETTE
  Blue    As Byte
  Green   As Byte
  Red     As Byte
  Reserve As Byte
End Type

' Palette Type for 24Bit
Private Type PALETTE24Bit
  Blue   As Byte
  Green  As Byte
  Red    As Byte
End Type

Public PixelWidth   As Long
Private DataArr()   As Long
Private IMGHeight   As Long
Private IMGWidth    As Long
Private StrFilePath As String

Private Sub Class_Initialize()
  PixelWidth = 1:  IMGHeight = 0:  IMGWidth = 0
End Sub
Property Get Width() As Double
  Width = IMGWidth
End Property
Property Get Height() As Double
  Height = IMGHeight
End Property
Property Get Path() As String
  Path = StrFilePath
End Property

Public Sub PrintBMP(ByVal strFileName As String, ByRef Dest As Range)
  Dim bmpFileHeader   As BITMAPFILEHEADER
  Dim bmpInfoHeader   As BITMAPINFOHEADER
  Dim XlPalette(255)  As PALETTE
  Dim Palette24       As PALETTE24Bit
  Dim R               As Long
  Dim C               As Long
  Dim I               As Long
  Dim Padding         As Byte
  Dim bytPixel        As Byte
    
  On Error GoTo CloseFile
  StrFilePath = strFileName
  Open strFileName For Binary As #1
        
  Get #1, , bmpFileHeader
  Get #1, , bmpInfoHeader
  IMGWidth = bmpInfoHeader.lngWidth + bmpInfoHeader.lngWidth Mod 4
  IMGHeight = bmpInfoHeader.lngHeight
        
  Set Dest = Dest.Resize(IMGHeight, IMGWidth)
    
  With Dest.Cells
    .ColumnWidth = PixelWidth
    .RowHeight = .Cells(1, 1).Width
    .Clear
    .Interior.Color = RGB(255, 255, 255)
  End With
    
  If bmpInfoHeader.intBitCount <= 8 Then
    For I = 0 To 255
      Get #1, , XlPalette(I)
    Next I
       
    For R = 1 To IMGHeight
      For C = 1 To IMGWidth
        If C <= bmpInfoHeader.lngWidth Then
          Get #1, , bytPixel
          Dest.Cells(IMGHeight + 1 - R, C).Interior.Color = RGB(XlPalette(bytPixel).Red, _
                                                                XlPalette(bytPixel).Green, _
                                                                XlPalette(bytPixel).Blue)
          DoEvents
        Else
          Get #1, , Padding
          Dest.Cells(IMGHeight + 1 - R, C).Interior.Color = RGB(255, 255, 255)
        End If
      Next C
    Next R
  Else
    For R = 1 To IMGHeight
      For C = 1 To IMGWidth
        If C <= bmpInfoHeader.lngWidth Then
          Get #1, , Palette24
          Dest.Cells(IMGHeight + 1 - R, C).Interior.Color = RGB(Palette24.Red, _
                                                                Palette24.Green, _
                                                                Palette24.Blue)
        Else
          Get #1, , Padding
          Dest.Cells(IMGHeight + 1 - R, C).Interior.Color = RGB(255, 255, 255)
        End If
        DoEvents
      Next C
    Next R
  End If
CloseFile:
  If Len(Err.Description) > 0 Then MsgBox Err.Description
  Close #1
End Sub

Public Function LoadBMP(ByVal strFileName As String) As Variant
  Dim bmpFileHeader   As BITMAPFILEHEADER
  Dim bmpInfoHeader   As BITMAPINFOHEADER
  Dim XlPalette(255)  As PALETTE
  Dim Palette24       As PALETTE24Bit
  Dim R               As Long
  Dim C               As Long
  Dim I               As Long
  Dim dAdjustedWidth  As Double
  Dim dAdjustedHeight As Double
  Dim Padding         As Byte
  Dim bytPixel        As Byte
  
  On Error GoTo CloseFile
  StrFilePath = strFileName
  Open strFileName For Binary As #1
  
  Get #1, , bmpFileHeader
  Get #1, , bmpInfoHeader
  
  If bmpInfoHeader.lngWidth Mod 4 Then
    IMGWidth = (((Int((bmpInfoHeader.lngWidth * bmpInfoHeader.intBitCount) / 32) + 1) * 4#)) / (bmpInfoHeader.intBitCount / 8#)
    If IMGWidth Mod 4 Then IMGWidth = Application.RoundUp(IMGWidth, 0)
  Else
    IMGWidth = bmpInfoHeader.lngWidth
  End If
  
  IMGHeight = bmpInfoHeader.lngHeight
  
  ReDim DataArr(1 To IMGHeight, 1 To IMGWidth)
  If bmpInfoHeader.intBitCount <= 8 Then
    For I = 0 To 255
      Get #1, , XlPalette(I)
    Next I
    For R = 1 To IMGHeight
      For C = 1 To IMGWidth
        If C <= bmpInfoHeader.lngWidth Then
          Get #1, , bytPixel
          DataArr(IMGHeight + 1 - R, C) = RGB(XlPalette(bytPixel).Red, _
                                              XlPalette(bytPixel).Green, _
                                              XlPalette(bytPixel).Blue)
        Else
          Get #1, , Padding
          DataArr(IMGHeight + 1 - R, C) = RGB(255, 255, 255)
        End If
      Next C
    Next R
  Else
    For R = 1 To IMGHeight
      For C = 1 To IMGWidth
        If C <= bmpInfoHeader.lngWidth Then
          Get #1, , Palette24
          DataArr(IMGHeight + 1 - R, C) = RGB(Palette24.Red, _
                                              Palette24.Green, _
                                              Palette24.Blue)
        Else
          Get #1, , Padding
          DataArr(IMGHeight + 1 - R, C) = RGB(255, 255, 255)
        End If
      Next C
    Next R
  End If
  LoadBMP = DataArr
CloseFile:
  If Len(Err.Description) > 0 Then MsgBox Err.Description
  Close #1
End Function


