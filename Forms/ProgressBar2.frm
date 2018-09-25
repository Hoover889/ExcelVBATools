VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar2 
   Caption         =   "Please Wait... Macro Running"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   OleObjectBlob   =   "ProgressBar2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------- API Declaration ----------
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Dim Canceled As Boolean

' This form is not fully functioning yet, all it does is increment on a timer

Private Sub UserForm_Activate()
    Dim Timeval As Single
    Dim I As Long
    Dim S As Long
    Dim StepText As String
    Canceled = False
    UpdatePct 0, 0
    Timeval = InputBox("Time to Spend in Minutes:", "Time", 5#)
    Timeval = Timeval * 100
    For S = 0 To 5
        For I = 1 To 100
            Sleep Timeval
            UpdatePct S, I
            If Canceled Then Exit Sub
        Next I
    Next S
    Unload Me
End Sub

Private Sub UpdatePct(ByVal S As Long, I As Long)
    Dim PctDone, SubPctDone As Single
    Dim StepText As String
    PctDone = (S * 100 + I) / 600
    SubPctDone = I / 100
    Select Case S
        Case 0: StepText = "Reading Data"
        Case 1: StepText = "Compiling Data"
        Case 2: StepText = "Processing Data"
        Case 3: StepText = "Merging Data"
        Case 4: StepText = "Error Checks"
        Case 5: StepText = "Generating Final Report"
    End Select
        
    With Me
        .MainProgress.Caption = Format(PctDone, "0%")
        .MainLabelProgress.Width = PctDone * (.MainProgress.Width - 10)
        .SubProgress.Caption = Format(SubPctDone, "0%") & " - " & StepText
        .SubLabelProgress.Width = SubPctDone * (.SubProgress.Width - 10)
    End With
    DoEvents
    Me.Repaint
End Sub

Private Sub UserForm_Terminate()
Canceled = True
End Sub

