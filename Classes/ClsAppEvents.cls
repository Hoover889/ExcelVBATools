VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variable Declarations
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Dim Prevsheet As Worksheet

Public Sub Setup(ByRef ActiveWS As Worksheet)
    Set Prevsheet = ActiveWS
End Sub

Public Sub GoBack()
    On Error Resume Next
      Prevsheet.Activate
    On Error GoTo 0
End Sub

Public Property Set Xl(Applic As Excel.Application)
    Set App = Applic
End Property

Public Property Get Xl() As Excel.Application
    Set Xl = App
End Property

Private Sub App_SheetDeactivate(ByVal Sh As Object)
    If TypeName(Sh) = "Worksheet" Then Set Prevsheet = Sh
End Sub

Private Sub App_WorkbookDeactivate(ByVal WB As Workbook)
    If TypeName(WB.ActiveSheet) = "Worksheet" Then Set Prevsheet = WB.ActiveSheet
End Sub

Private Sub App_WorkbookBeforeClose(ByVal WB As Workbook, Cancel As Boolean)
    On Error Resume Next
        If Prevsheet.Parent Is WB Then Set Prevsheet = Nothing
    On Error GoTo 0
End Sub


