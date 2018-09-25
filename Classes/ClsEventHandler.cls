VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------
'|          ______                   _                                             |
'|         |  ____|                 | |                                            |
'|         | |__ __   __ ___  _ __  | |_                                           |
'|         |  __|\ \ / // _ \| '_ \ | __|                                          |
'|         | |____\ V /|  __/| | | || |_                                           |
'|         |______|\_/  \___||_| |_| \__|                                          |
'|          _    _                    _  _                                         |
'|         | |  | |                  | || |                                        |
'|         | |__| |  __ _  _ __    __| || |  ___  _ __                             |
'|         |  __  | / _` || '_ \  / _` || | / _ \| '__|                            |
'|         | |  | || (_| || | | || (_| || ||  __/| |                               |
'|         |_|  |_| \__,_||_| |_| \__,_||_| \___||_|                               |
'|                                                                                 |
'|---------------------------------------------------------------------------------|
'|  -A simple framework for capturing Application/Workbook/Worksheet/Chart events  |
'|  -Constructor/Destructor and All setters and getters are already written        |
'|  -What to do when an event triggers will have to be written by you              |
'|  -The most common events already have Subs defined but others can be added      |
'|  -Each sub starts by disabling events and ends with enabling events             |
'|      -This is to prevent an infinite loop with events triggering each-other     |
'|---------------------------------------------------------------------------------|
'| Enjoy...                                                                        |
'| Created By Ryan Hoover                                                          |
'|                                                                                 |
'| Feel free to use or modify this code  but please provide attribution            |
'| If you like my work be sure to say thanks.                                      |
'-----------------------------------------------------------------------------------

    '/---------------------------------\
    '|      Variable Declarations      |
    '\---------------------------------/
Private WithEvents XLApp As Application
Attribute XLApp.VB_VarHelpID = -1
Private WithEvents XLWB As Workbook
Attribute XLWB.VB_VarHelpID = -1
Private WithEvents XLWS As Worksheet
Attribute XLWS.VB_VarHelpID = -1
Private WithEvents XLChar As Chart
Attribute XLChar.VB_VarHelpID = -1

    '/---------------------------------\
    '|  Class Constructor/Destructor   |
    '\---------------------------------/
Public Function Setup(Optional ByRef App As Application, _
                      Optional ByRef WB As Workbook, _
                      Optional ByRef WS As Worksheet, _
                      Optional ByRef Char As Chart) As ClsEventHandler
                      
    If Not App Is Nothing Then Set XLApp = App
    If Not WB Is Nothing Then Set XLWB = WB
    If Not WS Is Nothing Then Set XLWS = WS
    If Not Char Is Nothing Then Set XLChar = Char
End Function
Private Sub Class_Terminate()
    Set XLChar = Nothing:   Set XLWS = Nothing:     Set XLWB = Nothing: Set XLApp = Nothing
End Sub

    '/---------------------------------\
    '|     Property Getter/Setters     |
    '\---------------------------------/
Public Property Set App(ByRef Obj As Application):  Set XLApp = Obj:    End Property
Public Property Get App() As Application:           Set App = XLApp:    End Property
Public Property Set WB(ByRef Obj As Workbook):      Set XLWB = Obj:     End Property
Public Property Get WB() As Workbook:               Set WB = XLWB:      End Property
Public Property Set WS(ByRef Obj As Worksheet):     Set XLWS = Obj:     End Property
Public Property Get WS() As Worksheet:              Set WS = XLWS:      End Property
Public Property Set Char(ByRef Obj As Chart):       Set XLChar = Obj:   End Property
Public Property Get Char() As Chart:                Set Char = XLChar:  End Property


    '/---------------------------------\
    '|       Application Events        |-------------------------------------------------------------------------------------------------------------------
    '\---------------------------------/
        
Private Sub XLApp_NewWorkbook(ByVal WB As Workbook)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when any new blank workbook is created                          |
'|  Wb is the new workbook that was just created                                  |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLApp_WorkbookOpen(ByVal WB As Workbook)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when any workbook as opened                                     |
'|  Wb is the workbook that was just opened                                       |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLApp_WorkbookBeforeSave(ByVal WB As Workbook, _
                                     ByVal SaveAsUI As Boolean, _
                                     Cancel As Boolean)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers before any workbook is saved                                    |
'|  Wb is the workbook to be saved                                                |
'|  SaveAsUI can be used to determine if a Save-as menu will be displayed         |
'|      (useful if you don't want your template to be overwritten)                |
'|  Cancel is used to cancel saving, by default it is False,                      |
'|      have your macro change it to true to prevent saving                       |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLApp_WorkbookAfterSave(ByVal WB As Workbook, _
                                    ByVal Success As Boolean)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers after any workbook is saved                                     |
'|  Wb is the Workbook that was saved (or attempted to be saved)                  |
'|  Success indicates if the save attempt was successful, it can't be changed     |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLApp_WorkbookBeforeClose(ByVal WB As Workbook, _
                                      Cancel As Boolean)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers before any workbook is closed                                   |
'|  Wb is the workbook that is being closed                                       |
'|  Cancel is used to prevent closing, by default it is False                     |
'|      have your macro change the value of Cancel to prevent closing a WB        |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub


    '/---------------------------------\
    '|         Workbook Events         |-------------------------------------------------------------------------------------------------------------------
    '\---------------------------------/
    
Private Sub XLWB_BeforeSave(ByVal SaveAsUI As Boolean, _
                            Cancel As Boolean)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers before the workbook pointed to by XLWB is saved                 |
'|  Unlike the Application level event it only applies to XLWB                    |
'|  SaveAsUI can be used to determine if a Save-as menu will be displayed         |
'|      (useful if you don't want your template to be overwritten)                |
'|  Cancel is used to cancel saving, by default it is False,                      |
'|      have your macro change it to true to prevent saving                       |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWB_AfterSave(ByVal Success As Boolean)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers after the workbook pointed to by XLWB is saved                  |
'|  Success indicates if the save attempt was successful, it can't be changed     |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWB_BeforeClose(Cancel As Boolean)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers before the workbook pointed to by XLWB is closed                |
'|  Cancel is used to prevent closing, by default it is False                     |
'|      have your macro change the value of Cancel to prevent closing a WB        |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWB_Activate()
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when the Workbook pointed to by XLWB becomes active             |
'|  Activation happens when a workbook is selected by the user                    |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWB_Deactivate()
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when the Workbook pointed to by XLWB is no longer active        |
'|  Deactivation happens when another workbook is selected by the user            |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWB_NewSheet(ByVal Sh As Object)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when a new worksheet is added to XLWB                           |
'|  Sh is the newly created worksheet                                             |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWB_SheetDeactivate(ByVal Sh As Object)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when the active worksheet in XLWB is changed                    |
'|  Sh is the Sheet being deactivated                                             |
'|  This happens BEFORE the SheetActivate event                                   |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWB_SheetActivate(ByVal Sh As Object)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when the active worksheet in XLWB is changed                    |
'|  Sh is the Sheet being activated                                               |
'|  This happens AFTER the SheetDeactivate event                                  |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWB_SheetChange(ByVal Sh As Object, _
                             ByVal Target As Range)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'| This Triggers when a cell inside XLWB is modified                              |
'| Sh is the Sheet where the modification took place                              |
'| Target is the exact Range where the modification took place                    |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub


    '/---------------------------------\
    '|        Worksheet Events         |-------------------------------------------------------------------------------------------------------------------
    '\---------------------------------/

Private Sub XLWS_Activate()
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when XLWS is Activated (Selected)                               |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWS_Deactivate()
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when XLWS is Deactivated (Deselected)                           |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWS_Calculate()
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when XLWS is Calculated                                         |
'|  If Calculation is set to Automatic this triggers when any change is made      |
'|  If Calculation is set to Manual this triggers when XLWS.Calculate is called   |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWS_BeforeRightClick(ByVal Target As Range, _
                                  Cancel As Boolean)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'| This Triggers when a cell inside XLWS is Right Clicked                         |
'| Target is the exact Range where the modification took place                    |
'| Cancel allows you to prevent the right click menu from appearing               |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWS_SelectionChange(ByVal Target As Range)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'| This Triggers when a cell inside XLWS is selected                              |
'| Target is the exact Range that was selected                                    |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWS_Change(ByVal Target As Range)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'| This Triggers when a cell inside XLWS is modified                              |
'| Target is the exact Range where the modification took place                    |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLWS_PivotTableUpdate(ByVal Target As PivotTable)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'| This Triggers when a Pivot Table inside XLWS is updated                        |
'| Target is the Pivot Table that was updated                                     |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub


    '/---------------------------------\
    '|          Chart Events           |-------------------------------------------------------------------------------------------------------------------
    '\---------------------------------/

Private Sub XLChar_Activate()
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when XLChar is Activated (Selected)                             |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLChar_Deactivate()
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'|  This triggers when XLChar is Deactivated (Deselected)                         |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub


Private Sub XLChar_BeforeRightClick(Cancel As Boolean)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'| This Triggers when XLChar is Right Clicked                                     |
'| Cancel allows you to prevent the right click menu from appearing               |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLChar_MouseDown(ByVal Button As Long, _
                             ByVal Shift As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'| This Triggers when XLChar is Clicked (Before Mouse button is released)         |
'| Button indicates what Mouse button was pressed:                                |
'|     xlPrimaryButton   = 1 = Left Click                                         |
'|     xlSecondaryButton = 2 = Right Click                                        |
'| Shift indicates what Key was held when the Click happened                      |
'|     0 = No Key Pressed                                                         |
'|     1 = SHIFT Key                                                              |
'|     2 = CTRL Key                                                               |
'|     4 = ALT Key                                                                |
'| x is the horizontal coordinate of where the mouse pointer was                  |
'| y is the Vertical coordinate of where the mouse pointer was                    |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLChar_MouseUp(ByVal Button As Long, _
                           ByVal Shift As Long, _
                           ByVal X As Long, _
                           ByVal Y As Long)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'| This Triggers when XLChar is Clicked (After Mouse button is released)          |
'| Button indicates what Mouse button was pressed:                                |
'|     xlPrimaryButton   = 1 = Left Click                                         |
'|     xlSecondaryButton = 2 = Right Click                                        |
'| Shift indicates what Key was held when the Click happened                      |
'|     0 = No Key Pressed                                                         |
'|     1 = SHIFT Key                                                              |
'|     2 = CTRL Key                                                               |
'|     4 = ALT Key                                                                |
'| x is the horizontal coordinate of where the mouse pointer was                  |
'| y is the Vertical coordinate of where the mouse pointer was                    |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLChar_MouseMove(ByVal Button As Long, _
                             ByVal Shift As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'| This Triggers when the mouse is moved over the chart                           |
'| Button indicates what Mouse button was pressed:                                |
'|     xlPrimaryButton   = 1 = Left Click                                         |
'|     xlSecondaryButton = 2 = Right Click                                        |
'| Shift indicates what Key was held when the Click happened                      |
'|     0 = No Key Pressed                                                         |
'|     1 = SHIFT Key                                                              |
'|     2 = CTRL Key                                                               |
'|     4 = ALT Key                                                                |
'| x is the horizontal coordinate of where the mouse pointer was                  |
'| y is the Vertical coordinate of where the mouse pointer was                    |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

Private Sub XLChar_Select(ByVal ElementID As Long, _
                          ByVal Arg1 As Long, _
                          ByVal Arg2 As Long)
Application.EnableEvents = False
'/--------------------------------------------------------------------------------\
'| This Triggers when an element of the chart is selected                         |
'| ElementID indicates what Chart element was selected                            |
'| Arg1 and Arg2 provide info on the element selected                             |
'|         /--------------------------------------------\                         |
'|         |ElementID        |Arg1       |Arg2          |                         |
'|         |-----------------|-----------|--------------|                         |
'|         |xlAxis           |AxisIndex  |AxisType      |                         |
'|         |xlAxisTitle      |AxisIndex  |AxisType      |                         |
'|         |xlMajorGridlines |AxisIndex  |AxisType      |                         |
'|         |xlMinorGridlines |AxisIndex  |AxisType      |                         |
'|         |xlChartArea      |None       |None          |                         |
'|         |xlChartTitle     |None       |None          |                         |
'|         |xlLegend         |None       |None          |                         |
'|         |xlPlotArea       |None       |None          |                         |
'|         |xlDataLabel      |SeriesIndex|PointIndex    |                         |
'|         |xlLegendEntry    |SeriesIndex|None          |                         |
'|         |xlLegendKey      |SeriesIndex|None          |                         |
'|         |xlSeries         |SeriesIndex|PointIndex    |                         |
'|         |xlTrendline      |SeriesIndex|TrendLineIndex|                         |
'|         |xlErrorBars      |SeriesIndex|None          |                         |
'|         |xlShape          |ShapeIndex |None          |                         |
'|         \--------------------------------------------/                         |
'\--------------------------------------------------------------------------------/

Application.EnableEvents = True:    End Sub

