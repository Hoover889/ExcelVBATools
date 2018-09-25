VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExplorer 
   Caption         =   "Excplorer"
   ClientHeight    =   11985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13905
   OleObjectBlob   =   "frmExplorer.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnGo_Click()
    Call wbMain.Navigate(tbURL.Text)
End Sub

Private Sub tbURL_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        If InStr(1, tbURL.Text, "://") < 4 Then tbURL.Text = "http://" & tbURL.Text
        Call wbMain.Navigate(tbURL.Text)
    End If
End Sub

Private Sub UserForm_Initialize()
    Call ResizeAll
    wbMain.Navigate ("http://google.com")
End Sub

Private Sub ResizeAll()
    Me.Width = Application.Width - 100
    Me.Height = Application.Height - 100
    tbURL.Width = Me.Width - 60
    btnGo.Left = tbURL.Width + 10
    wbMain.Width = Me.Width - (wbMain.Left * 3)
    wbMain.Height = Me.Height - (wbMain.Top * 2)
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    MsgBox KeyCode
End Sub

Private Sub wbMain_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    tbURL.Text = URL
End Sub

Private Sub wbMain_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub wbMain_TitleChange(ByVal Text As String)
    frmExplorer.Caption = "Excplorer: " & Text
End Sub
