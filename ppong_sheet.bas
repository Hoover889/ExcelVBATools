Attribute VB_Name = "ppong_sheet"
Option Explicit
Option Base 0

'/---------------------------------------------------\
'|         _         ||                              |
'|        |-|        ||                              |
'|    ____| |____    ||  _____                       |
'|   /   _| |_   \   // |  __ \                      |
'|  |  / ,| |. \  |_//  | |__) |__  _ __   __ _      |
'|  | ( ( '-' ) ) |-'   |  ___/ _ \| '_ \ / _` |     |
'|  |  \ `'"'' /  |     | |  | (_) | | | | (_| |     |
'|  |   `-----'   ;     |_|   \___/|_| |_|\__, |     |
'|  |\___________/|                        __/ |     |
'|  |             ;                       |___/      |
'|   \___________/                                   |
'|                                                   |
'|---------------------------------------------------|
'|                                                   |
'|  Start game by pressing ctrl + m                  |
'|  move paddles using up & down arrow keys          |
'|  Exit game by pressing esc                        |
'|                                                   |
'\---------------------------------------------------/

#If VBA7 Then
  Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
  Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Integer) As Long
#End If

Private Const Key_Esc  As Long = 27    ' 0x1B
Private Const Key_Up   As Long = 38    ' 0x26
Private Const Key_Down As Long = 40    ' 0x28

Dim X_pscale         As Double   ' scale for drawing board
Dim X_PongShapes(5)  As Shape    ' all shapes
Dim XY_ball(4)       As Double   ' ball position, direction and speed
Dim XY_pongs(1)      As Double   ' pads positions Y only
Dim XPP_Xpos         As Double
Dim XPP_Ypos         As Double   'general reference positions
Dim XPP_Stats(6)     As Long     ' 0 status, 1 p-wins, 2 c-wins, settings: 3,4,5: size, speed, difficulty, colorindex

Private Sub PP_SetBoard()
    Dim X_Pcount As Integer
    On Error Resume Next
    With ActiveSheet
      For X_Pcount = 0 To 5
        .Shapes("XPPal_" & X_Pcount + 1).Delete
      Next X_Pcount
    On Error GoTo 0
      X_pscale = 1
      XPP_Xpos = ActiveCell.Left
      XPP_Ypos = ActiveCell.Offset(2, 0).Top
      With .Shapes
        .AddShape(msoShapeRectangle, XPP_Xpos, XPP_Ypos + 81, 9, 60).Name = "XPPal_1"
        .AddShape(msoShapeRectangle, XPP_Xpos + 321, XPP_Ypos + 81, 9, 60).Name = "XPPal_2"
        .AddShape(msoShapeRectangle, XPP_Xpos, XPP_Ypos, 330, 6).Name = "XPPal_3"
        .AddShape(msoShapeRectangle, XPP_Xpos, XPP_Ypos + 216, 330, 6).Name = "XPPal_4"
        .AddShape(msoShapeOval, XPP_Xpos + 162, XPP_Ypos + 102, 12, 12).Name = "XPPal_5"
        .AddLabel(msoTextOrientationHorizontal, XPP_Xpos + 162, XPP_Ypos + 90, 10, 10).Name = "XPPal_6"
      End With
      Set X_PongShapes(5) = .Shapes("XPPal_6")
    End With
    For X_Pcount = 0 To 4
      Set X_PongShapes(X_Pcount) = ActiveSheet.Shapes("XPPal_" & X_Pcount + 1)
      With X_PongShapes(X_Pcount)
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = ThisWorkbook.Colors(XPP_Stats(6))
        .Fill.Solid
      End With
    Next X_Pcount
    With X_PongShapes(5).TextFrame.Characters.Font
      .Name = "Arial"
      .Bold = True
      .Size = 38
      .ColorIndex = 2
    End With
End Sub

Sub XPPong()
Attribute XPPong.VB_ProcData.VB_Invoke_Func = "m\n14"
    Dim xi                As Double
    Dim X_time            As Long
    Dim XiS               As Long
    Dim XBallScan(333, 2) As Double
    
    With Application: .OnKey "{UP}", "": .OnKey "{DOWN}", "": End With
    XPP_Stats(1) = 0: XPP_Stats(2) = 0: XPP_Stats(4) = 2: XPP_Stats(5) = 1: XPP_Stats(6) = 1
    Do
      xi = 0: PP_SetBoard: XPP_Stats(0) = 1
      XY_pongs(0) = XPP_Ypos + 81: XY_pongs(1) = XPP_Ypos + 81: XY_ball(0) = XPP_Ypos + 102: XY_ball(1) = XPP_Xpos + 162
      X_PongShapes(4).Top = XY_ball(0):
      X_PongShapes(4).Left = XY_ball(1)
      XY_ball(3) = (Int(Rnd * 2) * 2 - 1) * (1.94 + 0.16 * Rnd)
      XY_ball(4) = 4.5
      XY_ball(2) = (Int(Rnd * 2) * 2 - 1) * (XY_ball(4) - XY_ball(3) ^ 2) ^ 0.5
      XBallScan(333, 0) = XPP_Ypos + 102: DoEvents: X_time = timeGetTime
      Do
        If GetAsyncKeyState(Key_Esc) <> 0 Then XPP_Stats(0) = -2
      Loop While timeGetTime - X_time < 500 And XPP_Stats(0) > -2
      Do
        xi = xi + 1
        XY_ball(0) = XY_ball(0) + XY_ball(2)
        XY_ball(1) = XY_ball(1) + XY_ball(3)
        Sleep 12 - 2 * XPP_Stats(4)
        With X_PongShapes(4)
          .Top = XY_ball(0)
          .Left = XY_ball(1)
        End With
        If XPP_Stats(0) > 0 Then
          If GetAsyncKeyState(Key_Up) <> 0 And xi Mod 3 = 0 And XY_pongs(0) > XPP_Ypos + 6 Then
            XY_pongs(0) = XY_pongs(0) - 3
            X_PongShapes(0).Top = XY_pongs(0)
          End If
          If GetAsyncKeyState(Key_Down) <> 0 And xi Mod 3 = 0 And XY_pongs(0) < XPP_Ypos + 156 Then
            XY_pongs(0) = XY_pongs(0) + 3
            X_PongShapes(0).Top = XY_pongs(0)
          End If
          If GetAsyncKeyState(Key_Esc) <> 0 Then XPP_Stats(0) = -2
          If xi Mod 3 = 1 Then
            Select Case XPP_Stats(5)
              Case 2
                If XY_ball(3) > 0 Then
                  If XY_pongs(1) + 3 + 3 * Int(Abs(XY_ball(2))) > XBallScan(333, 0) Then XY_pongs(1) = XY_pongs(1) - 3
                  If XY_pongs(1) + 45 - 3 * Int(Abs(XY_ball(2))) < XBallScan(333, 0) Then XY_pongs(1) = XY_pongs(1) + 3
                Else
                  If XY_pongs(1) + 18 > Application.Min(XPP_Ypos + 111, XY_ball(0)) Then XY_pongs(1) = XY_pongs(1) - 3
                  If XY_pongs(1) + 30 < Application.Max(XPP_Ypos + 111, XY_ball(0)) Then XY_pongs(1) = XY_pongs(1) + 3
                End If
              Case 1
                If XY_ball(3) > 0 Then
                  If XY_pongs(1) + 6 > XY_ball(0) Then
                    If XY_ball(2) > 1 And XY_ball(1) < 252 - 24 * XY_ball(2) Then
                      XY_pongs(1) = XY_pongs(1) + 3
                    Else
                      XY_pongs(1) = XY_pongs(1) - 3
                    End If
                  End If
                  If XY_pongs(1) + 42 < XY_ball(0) Then
                    If XY_ball(2) < -1 And XY_ball(1) < 252 + 24 * XY_ball(2) Then
                      XY_pongs(1) = XY_pongs(1) - 3
                    Else
                      XY_pongs(1) = XY_pongs(1) + 3
                    End If
                  End If
                Else
                  XY_pongs(1) = XY_pongs(1) + 3 * Sgn(XPP_Ypos + 81 - 3 _
                    * Int(Abs(XY_ball(2)) * 12) * Sgn(XY_ball(2)) - XY_pongs(1))
                End If
              Case 0
                If XY_pongs(1) + 6 > XY_ball(0) Then XY_pongs(1) = XY_pongs(1) - 3
                If XY_pongs(1) + 42 < XY_ball(0) Then XY_pongs(1) = XY_pongs(1) + 3
            End Select
            If XY_pongs(1) <= XPP_Ypos + 6 Then XY_pongs(1) = XPP_Ypos + 6
            If XY_pongs(1) >= XPP_Ypos + 156 Then XY_pongs(1) = XPP_Ypos + 156
            X_PongShapes(1).Top = XY_pongs(1)
          End If
        End If
        DoEvents
        If XY_ball(1) <= XPP_Xpos + 6 And XY_ball(3) < 0 Then
          If XY_ball(0) < XY_pongs(0) - 12 Or XY_ball(0) > XY_pongs(0) + 60 Or XPP_Stats(0) = 0 Then
            XPP_Stats(0) = XPP_Stats(0) - 1
          Else
            If Abs(XY_pongs(0) + 24 - XY_ball(0)) <= 18 Then
              XY_ball(3) = -XY_ball(3)
            Else
              XY_ball(3) = -XY_ball(3)
              XY_ball(2) = XY_ball(2) - 0.0375 * ((XY_ball(4) / 4.5) ^ 0.5) _
                * Sgn(XY_ball(3)) * Sgn(XY_pongs(0) + 24 - XY_ball(0)) * (Abs(XY_pongs(0) + 24 - XY_ball(0)) - 18)
              XY_ball(4) = XY_ball(4) + 0.25
            End If
          End If
        End If
        If XY_ball(1) >= XPP_Xpos + 312 And XY_ball(3) > 0 Then
          If XY_ball(0) < XY_pongs(1) - 12 Or XY_ball(0) > XY_pongs(1) + 60 Or XPP_Stats(0) = 0 Then
            XPP_Stats(0) = XPP_Stats(0) - 1
          Else
            If Abs(XY_pongs(1) + 24 - XY_ball(0)) <= 18 Then
              XY_ball(3) = -XY_ball(3)
            Else
              XY_ball(3) = -XY_ball(3)
              XY_ball(2) = XY_ball(2) + 0.0375 * ((XY_ball(4) / 4.5) ^ 0.5) _
                * Sgn(XY_ball(3)) * Sgn(XY_pongs(1) + 24 - XY_ball(0)) _
                * (Abs(XY_pongs(1) + 24 - XY_ball(0)) - 18)
              XY_ball(4) = XY_ball(4) + 0.25
            End If
          End If
        End If
        If (XY_ball(0) <= XPP_Ypos + 3 And XY_ball(2) < 0) _
          Or (XY_ball(0) >= XPP_Ypos + 207 And XY_ball(2) > 0) Then XY_ball(2) = -XY_ball(2)
        If Abs(XY_ball(3)) < 1 Then
          XY_ball(3) = Sgn(XY_ball(3)) + (1 - Abs(Sgn(XY_ball(3)))) * Sgn(XY_ball(3))
          XY_ball(2) = Sgn(XY_ball(2)) * (XY_ball(4) - XY_ball(3) ^ 2) ^ 0.5
        End If
        If Abs(XY_ball(2)) > (XY_ball(4) - 0.000002) ^ 0.5 Then XY_ball(2) = (XY_ball(4) - 0.000002) ^ 0.5 * Sgn(XY_ball(2))
        XY_ball(3) = Sgn(XY_ball(3)) * (XY_ball(4) - XY_ball(2) ^ 2) ^ 0.5
        If XY_ball(1) > XPP_Xpos + 12 And XY_ball(3) > 0 And xi Mod 3 = 2 Then
          XBallScan(0, 0) = XY_ball(2): XBallScan(0, 1) = XY_ball(0): XBallScan(0, 2) = XY_ball(1): XiS = 0
          Do
            XiS = XiS + 1
            If (XBallScan(XiS - 1, 1) <= XPP_Ypos + 3 And XY_ball(2) < 0) _
              Or (XBallScan(XiS - 1, 1) >= XPP_Ypos + 207 And XY_ball(2) > 0) Then
              XBallScan(XiS, 0) = -XBallScan(XiS - 1, 0)
            Else
              XBallScan(XiS, 0) = XBallScan(XiS - 1, 0)
            End If
            XBallScan(XiS, 2) = XBallScan(XiS - 1, 2) + XY_ball(3):
            XBallScan(XiS, 1) = XBallScan(XiS - 1, 1) + XBallScan(XiS, 0)
          Loop While XBallScan(XiS, 2) < XPP_Xpos + 312
          XBallScan(333, 0) = XBallScan(XiS, 1)
        End If
      Loop While XPP_Stats(0) >= 0
      If XPP_Stats(0) = -1 Then
        If XY_ball(1) <= XPP_Xpos + 6 Then XPP_Stats(2) = XPP_Stats(2) + 1 Else XPP_Stats(1) = XPP_Stats(1) + 1
        With X_PongShapes(5)
          .TextFrame.Characters.Text = XPP_Stats(1) & ":" & XPP_Stats(2)
          .Left = XPP_Xpos + 165 - .Width * 0.5: .Top = XPP_Ypos + 111 - .Height * 0.5
        End With
        DoEvents: X_time = timeGetTime
        Do
          If GetAsyncKeyState(Key_Esc) <> 0 Then XPP_Stats(0) = -2
        Loop While timeGetTime - X_time < 1000 And XPP_Stats(0) > -2
        'X_PongShapes(5).TextFrame.Characters.Text = ""
      End If
  Loop While XPP_Stats(0) > -2
  For xi = 0 To 5:    X_PongShapes(xi).Delete:  Next xi
  With Application: .OnKey "{UP}": .OnKey "{DOWN}": End With
End Sub

