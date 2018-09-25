VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar2lv 
   Caption         =   "Please Wait... Macro Running"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   OleObjectBlob   =   "ProgressBar2lv.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar2lv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Lv1Steps As Double
Dim Lv1Step  As Double
Dim Lv1Inc   As Double
Dim Lv1Prog  As Double
Dim Lv2Steps As Double
Dim Lv2Step  As Double
Dim Lv2Inc   As Double
Dim Lv2Prog  As Double
Dim CurrCaption As String

' This form is not fully functioning yet, all it does is increment on a timer

Private Sub UserForm_Activate()
  Lv1Steps = 1:  Lv1Step = 0:  Lv1Inc = 1 / Lv1Steps:  Lv1Prog = 0
  Lv2Steps = 1:  Lv2Step = 0:  Lv2Inc = 1 / Lv2Steps:  Lv2Prog = 0
  CurrCaption = vbNullString:  DrawBars
End Sub

Public Sub setInitialSteps(ByVal LV1_Steps As Long, ByVal LV2_Steps As Long, Optional ByVal Caption As String = vbNullString)
  Lv1Steps = LV1_Steps:  Lv1Step = 0:  Lv1Inc = 1 / Lv1Steps:  Lv1Prog = 0
  Lv2Steps = LV2_Steps:  Lv2Step = 0:  Lv2Inc = 1 / Lv2Steps:  Lv2Prog = 0
  CurrCaption = Caption:  DrawBars
End Sub

Public Sub IncrementL1(ByVal LV2_Steps As Long, Optional ByVal Caption As String = vbNullString)
  If Lv1Step >= Lv1Steps Then Exit Sub
  Lv1Step = Lv1Step + 1:  Lv1Prog = Lv1Prog + Lv1Inc
  Lv2Steps = LV2_Steps:  Lv2Step = 0:  Lv2Inc = 1 / Lv2Steps:  Lv2Prog = 0
  CurrCaption = Caption:  DrawBars
End Sub

Public Sub IncrementL2()
  If Lv2Step >= Lv2Steps Then Exit Sub
  Lv2Step = Lv2Step + 1:  Lv2Prog = Lv2Prog + Lv2Inc
  DrawBars
End Sub

Private Sub DrawBars()
  With Me
    .MainProgress.Caption = Format(Lv1Prog, "0%") & IIf(CurrCaption = vbNullString, vbNullString, " - " & CurrCaption)
    .MainLabelProgress.Width = Lv1Prog * (.MainProgress.Width - 10)
    .SubProgress.Caption = Format(Lv2Prog, "0%")
    .SubLabelProgress.Width = Lv2Prog * (.SubProgress.Width - 10)
  End With
  Me.Repaint
  DoEvents
End Sub
