VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   " Macro"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   OleObjectBlob   =   "ProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------------------------------------------------------------------------
'|      _____                                                                                                            |
'|     |  __ \                                                                                                           |
'|     | |__) |_ __  ___    __ _  _ __  ___  ___  ___                                                                    |
'|     |  ___/| '__|/ _ \  / _` || '__|/ _ \/ __|/ __|                                                                   |
'|     | |    | |  | (_) || (_| || |  |  __/\__ \\__ \                                                                   |
'|     |_|__  |_|   \___/  \__, ||_|   \___||___/|___/                                                                   |
'|     |  _ \               __/ |                                                                                        |
'|     | |_) |  __ _  _ __ |___/                                                                                         |
'|     |  _ <  / _` || '__|                                                                                              |
'|     | |_) || (_| || |                                                                                                 |
'|     |____/  \__,_||_|                                                                                                 |
'|                                                                                                                       |
'| Created By Ryan Hoover                                                                                                |
'|            Ryan.Hoover@Loreal.com                                                                                     |
'-------------------------------------------------------------------------------------------------------------------------
'|                                                                                                                       |
'| The Generic Progress Bar is a Tool to easily add a progress bar to any macro.                                         |
'|                                                                                                                       |
'| The Overhead of the ProgressBar is about 71k Assembly Instructions and refreshes the monitor                          |
'|     on most Computers this overhead is about 0.57 miliseconds                                                         |
'|                                                                                                                       |
'| To add the Generic Progress Bar to a Macro just call the progress bar using the Show Method                           |
'|                                                                                                                       |
'|       ProgressBar.Show                                                                                                |
'|                                                                                                                       |
'| At Any time Call the UpdatePct method to update the progress bar percentage                                           |
'|                                                                                                                       |
'|       ProgressBar.UpdatePct(1,100) '1 percent done                                                                    |
'|       ProgressBar.UpdatePct(2,100) '2 percent done                                                                    |
'|                                                                                                                       |
'| At Any time Call the ChangeText method to update the progress bar text                                                |
'|                                                                                                                       |
'|       ProgressBar.ChangeText("Opening Files")                                                                         |
'|                                                                                                                       |
'| To get rid of the ProgressBar use the Unload command                                                                  |
'|                                                                                                                       |
'|       Unload ProgressBar                                                                                              |
'|                                                                                                                       |
'|                                                                                                                       |
'| Enjoy...                                                                                                              |
'| Created By Ryan Hoover                                                                                                |
'|            Ryan.Hoover@LOreal.com                                                                                     |
'|                                                                                                                       |
'| Feel free to use or modify this code however you like but please provide attribution                                  |
'| If you like my work be sure to say thanks.                                                                            |
'-------------------------------------------------------------------------------------------------------------------------

Option Explicit
Private TotSteps  As Long
Private CurrStep  As Long

Public Property Get TotalSteps() As Long
  TotalSteps = TotSteps
End Property

Public Property Let TotalSteps(ByVal Val As Long)
  TotSteps = Val
End Property

Private Sub UserForm_Activate()
    Me.Repaint
    UpdatePct 0, 1
End Sub

Public Sub UpdatePct(ByVal Step As Long, ByVal Total As Long)
    Dim PctDone As Single
    PctDone = Step / Total
    With Me
        .FrameProgress.Caption = Format(PctDone, "0%")
        .LabelProgress.Width = PctDone * (.FrameProgress.Width - 10)
    End With
    DoEvents
    Me.Repaint
End Sub

Public Sub ChangeText(ByVal NewText As String)
    Me.StatusText.Caption = NewText
End Sub
