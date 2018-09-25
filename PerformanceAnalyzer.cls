VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'-------------------------------------------------------------------------------------------------------------------------
'|      _____              __                                                                                            |
'|     |  __ \            / _|                                                                                           |
'|     | |__) |___  _ __ | |_  ___   _ __  _ __ ___    __ _  _ __    ___  ___                                            |
'|     |  ___// _ \| '__||  _|/ _ \ | '__|| '_ ` _ \  / _` || '_ \  / __|/ _ \                                           |
'|     | |   |  __/| |   | | | (_) || |   | | | | | || (_| || | | || (__|  __/                                           |
'|     |_|    \___||_|   |_|  \___/ |_|   |_| |_| |_| \__,_||_| |_| \___|\___|                                           |
'|                               _                                                                                       |
'|          /\                  | |                                                                                      |
'|         /  \    _ __    __ _ | | _   _  ____ ___  _ __                                                                |
'|        / /\ \  | '_ \  / _` || || | | ||_  // _ \| '__|                                                               |
'|       / ____ \ | | | || (_| || || |_| | / /|  __/| |                                                                  |
'|      /_/    \_\|_| |_| \__,_||_| \__, |/___|\___||_|                                                                  |
'|                                   __/ |                                                                               |
'|                                  |___/                                                                                |
'|                                                                                                                       |
'|   Version 2.0                                                                                                         |
'|                                                                                                                       |
'| Created By Ryan Hoover                                                                                                |
'|            RHoover@us.Loreal.com                                                                                      |
'-------------------------------------------------------------------------------------------------------------------------
'|                                                                                                                       |
'| The Performance Analyzer is a Tool to help you identify what parts of a macro are slowing it down.                    |
'|                                                                                                                       |
'| The overhead of the PerformanceAnalyzer is minimal, about 125 assembly instructions, which can execute in paralel     |
'|     This equates to about 1 millionth of a second on most computers                                                   |
'|                                                                                                                       |
'|                                                                                                                       |
'| To add the Performance Analyzer to a Macro define an object as a 'PerformanceAnalyzer'                                |
'|                                                                                                                       |
'|       Dim Test as PerformanceAnalyzer                                                                                 |
'|                                                                                                                       |
'| If your macro spans multiple Modules you may want to make the variable Global                                         |
'|                                                                                                                       |
'|       Public Test as PerformanceAnalyzer                                                                              |
'|                                                                                                                       |
'| At the Start of the macro initialize the analyzer like this:                                                          |
'|                                                                                                                       |
'|       Set Test = New PerformanceAnalyzer                                                                              |
'|                                                                                                                       |
'| The timer will start as soon as the object is initialized                                                             |
'|                                                                                                                       |
'| After an Operation call the MarkStep method with a string describing what step was just performed                     |
'| Example:                                                                                                              |
'|                                                                                                                       |
'|       ...Do Something...                                                                                              |
'|       Test.MarkStep "First Step"                                                                                      |
'|       ...Do Something...                                                                                              |
'|       Test.MarkStep "Second Step"                                                                                     |
'|       ...Do Something...                                                                                              |
'|       Test.MarkStep "Third Step"                                                                                      |
'|                                                                                                                       |
'| At the end of your macro call the ExportPerformanceReport function to get a report of your macro                      |
'|                                                                                                                       |
'|       Dim MacroReport as Workbook                                                                                     |
'|       Set MacroReport = Test.ExportPerformanceReport                                                                  |
'|                                                                                                                       |
'| Or if you dont want to store the Performance Report as a variable use this:                                           |
'|                                                                                                                       |
'|       Test.ExportPerformanceReport                                                                                    |
'|                                                                                                                       |
'| The Performance report shows how long each step takes as well as the total runtime.                                   |
'| This will allow you to better identify what parts of the macro need optimization                                      |
'|                                                                                                                       |
'| If you don't want to break the macro down by steps, you can use the function RunTime                                  |
'| This returns the time since the Performance Analyzer was initiated in seconds                                         |
'|                                                                                                                       |
'|       Dim TimeTaken as Double                                                                                         |
'|       TimeTaken = Test.Runtime                                                                                        |
'|       MsgBox "The macro took " & TimeTaken & " seconds to complete"                                                   |
'|                                                                                                                       |
'| To Completely re-set the Performance Analyzer call the method Reset                                                   |
'| This is Useful if PerformanceAnalyzer is a public variable but multiple macros need separate performance reports      |
'|                                                                                                                       |
'| The Performance Analyzer can also be used to speed up a macro.                                                        |
'| Call the EnableSpeedup Method to change Excel's settings to make your macro run faster                                |
'| The EnableSpeedup method take in 1 parameter, telling it whether to change the calculation mode                       |
'|                                                                                                                       |
'|       Test.EnableSpeedup True                                                                                         |
'|                                                                                                                       |
'| to turn off the optimizations use the DisableSpeedup method                                                           |
'|                                                                                                                       |
'|       Test.DisableSpeedup                                                                                             |
'|                                                                                                                       |
'| Enjoy...                                                                                                              |
'| Created By Ryan Hoover                                                                                                |
'|            RHoover@US.LOreal.com                                                                                      |
'|                                                                                                                       |
'| Feel free to use or modify this code however you like but please provide attribution                                  |
'| If you like my work be sure to say thanks.                                                                            |
'-------------------------------------------------------------------------------------------------------------------------


Option Explicit
Option Base 1
'---------- API Declarations ----------
#If VBA7 Then
  Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
  Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
  Private Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
  Private Declare Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If

Private Const DefaultLim  As Long = 8192

'---------- Global Variable Declarations ----------
Private StepNum             As Long
Private Lim                 As Long
Private STime               As Currency
Private Freq                As Currency
Private DescArr()           As Variant
Private TimeArr()           As Currency
Private Enabled             As Boolean
Private ScreenState         As Boolean
Private EventState          As Boolean
Private AnimationState      As Boolean
Private AlertState          As Boolean
Private StatusbarState      As Boolean
Private PrintComState       As Boolean
Private CalculationState    As XlCalculation

Private Sub Class_Initialize()
  StepNum = 0: Lim = DefaultLim: getFrequency Freq:  getTickCount STime:     Enabled = False
  ReDim Preserve DescArr(Lim): ReDim Preserve TimeArr(Lim)
End Sub

Private Sub Class_Terminate(): Me.DisableSpeedup: Erase DescArr, TimeArr:  End Sub

Public Sub MarkStep(Optional ByVal Description As String = vbNullString)
  Dim Temp As Currency:  StepNum = StepNum + 1
  If StepNum >= Lim Then: Lim = Lim * 2: ReDim Preserve DescArr(Lim): ReDim Preserve TimeArr(Lim)
  If Description = vbNullString Then Description = "Step# " & StepNum
  DescArr(StepNum) = Description: getTickCount Temp: TimeArr(StepNum) = Temp
End Sub

Public Property Get RunTime() As Double
Attribute RunTime.VB_UserMemId = 0
  Dim Temp As Currency: getTickCount Temp: RunTime = (Temp - STime) / Freq
End Property
Public Function RunTimeMsg(Optional ByVal Message As String = "Macro Complete", Optional ByVal Title As String = "Macro Complete") As Long
  Dim Msg As String: Msg = Message & vbNewLine & "Run time was " & Format(Me.RunTime, "#,##0.00") & " seconds"
  RunTimeMsg = MsgBox(Msg, 64, Title)
End Function

Public Function ExportPerformanceReport() As Workbook
  Dim I As Long, Logsheet As Workbook:    Set Logsheet = Workbooks.Add
  With Logsheet.Sheets(1)
    .Cells(1, 1).Value = "Step#":      .Cells(2, 1).Value = 0:            .Cells(1, 2).Value = "Step Descrip"
    .Cells(2, 2).Value = "Initialize": .Cells(1, 3).Value = "Time Stamp": .Cells(2, 3).Value = 0
    .Cells(1, 4).Value = "Time taken": .Cells(2, 4).Value = 0
    For I = 1 To StepNum
      .Cells(2 + I, 1).Value = I:                           .Cells(2 + I, 2).Value = DescArr(I)
      .Cells(2 + I, 3).Value = (TimeArr(I) - STime) / Freq: .Cells(2 + I, 4).FormulaR1C1 = "=RC[-1]-R[-1]C[-1]"
    Next I:
  End With:   Set ExportPerformanceReport = Logsheet
End Function

Public Sub Reset()
  Erase DescArr, TimeArr: Lim = DefaultLim: StepNum = 0: getTickCount STime
  ReDim DescArr(Lim): ReDim TimeArr(Lim)
End Sub

Public Sub EnableSpeedup(Optional ByVal Calculation As Boolean = False, Optional ByVal SaveCurrent As Boolean = True)
  Dim WB As Workbook
  Dim WS As Worksheet
  If Enabled Then Exit Sub
  With Application
    If SaveCurrent Then
      ScreenState = .ScreenUpdating:      EventState = .EnableEvents:
      AnimationState = .EnableAnimations: AlertState = .DisplayAlerts:
      StatusbarState = .DisplayStatusBar: CalculationState = .Calculation:
      PrintComState = .PrintCommunication
    Else
      ScreenState = True:                 EventState = True
      AnimationState = True:              AlertState = True
      StatusbarState = True:              CalculationState = xlCalculationAutomatic
      PrintComState = True
    End If
    .ScreenUpdating = False:                .EnableEvents = False
    .EnableAnimations = False:              .DisplayAlerts = False
    .DisplayStatusBar = False:              If Not Calculation Then .Calculation = xlCalculationManual
    .PrintCommunication = False
  End With:
  For Each WB In Application.Workbooks:    For Each WS In WB.Worksheets
      WS.DisplayPageBreaks = False
  Next WS:  Next WB
  Enabled = True
End Sub

Public Sub DisableSpeedup()
  If Not Enabled Then Exit Sub
  With Application
    .ScreenUpdating = ScreenState:      .EnableEvents = EventState
    .EnableAnimations = AnimationState: .DisplayAlerts = AlertState
    .DisplayStatusBar = StatusbarState: .Calculation = CalculationState
  End With:   Enabled = False
End Sub
