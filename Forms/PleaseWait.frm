VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PleaseWait 
   Caption         =   "Please Wait..."
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6915
   OleObjectBlob   =   "PleaseWait.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PleaseWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------------------------------------
'|      _____   _                                                                                                        |
'|     |  __ \ | |                                                                                                       |
'|     | |__) || |  ___   __ _  ___   ___                                                                                |
'|     |  ___/ | | / _ \ / _` |/ __| / _ \                                                                               |
'|     | |     | ||  __/| (_| |\__ \|  __/                                                                               |
'|     |_|     |_| \___| \__,_||___/ \___|                                                                               |
'|      \ \        / /    (_)| |                                                                                         |
'|       \ \  /\  / /__ _  _ | |_                                                                                        |
'|        \ \/  \/ // _` || || __|                                                                                       |
'|         \  /\  /| (_| || || |_                                                                                        |
'|          \/  \/  \__,_||_| \__|                                                                                       |
'|                                                                                                                       |
'| Created By Ryan Hoover                                                                                                |
'-------------------------------------------------------------------------------------------------------------------------
'|                                                                                                                       |
'| The Please Wait Form is a Tool for when your macro does not have any way of making a progress bar                     |
'|                                                                                                                       |
'| To add the Plese Wait Form to a Macro just call the Form using the Show Method                                        |
'|                                                                                                                       |
'|       Pleasewait.Show                                                                                                 |
'|                                                                                                                       |
'| At Any time Call the ChangeText method to update the Form's Text                                                      |
'|                                                                                                                       |
'|       PleaseWait.ChangeText("Opening Files")                                                                          |
'|                                                                                                                       |
'| To get rid of the Form use the Unload command                                                                         |
'|                                                                                                                       |
'|       Unload PleaseWait                                                                                               |
'|                                                                                                                       |
'|                                                                                                                       |
'| Enjoy...                                                                                                              |
'| Created By Ryan Hoover                                                                                                |
'|            Ryan.Hoover@LOreal.com                                                                                     |
'|                                                                                                                       |
'| Feel free to use or modify this code however you like but please provide attribution                                  |
'| If you like my work be sure to say thanks.                                                                            |
'-------------------------------------------------------------------------------------------------------------------------

Public Sub ChangeText(ByVal Txt As String)
    Me.CustomLabel.Caption = Txt: Me.Repaint
End Sub

Private Sub UserForm_Activate():    Me.Repaint: End Sub

