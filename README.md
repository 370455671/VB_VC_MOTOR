# VB_VC_MOTOR
此编程语言基于中文Visual Basic 6.0
Option Explicit
Const Max_Axes_Num = 1
Dim glTotalAxes As Long                     '当前计算机中轴数
Dim glTotalCards As Long                    '当前计算机中卡数
Dim glDllMajor As Long, glDllMinor1 As Long, glDllMinor2 As Long
Dim glSysMajor As Long, glSysMinor1 As Long, glSysMinor2 As Long
Dim glCardMajor As Long, glCardMinor1 As Long, glCardMinor2 As Long
Dim glCardtype As Long
Dim FlagLoading As Boolean              '初始化过程标记，初始化完成后该值为False

Dim condis As Double                    '相对距离参数
Dim MAxesNum(Max_Axes_Num) As Long      '各轴轴号
Dim MDis(Max_Axes_Num) As Long          '各轴运动位移量
Dim MLowSpeed(Max_Axes_Num) As Double   '梯形速度低速速度
Dim MHighSpeed(Max_Axes_Num) As Double  '梯形速度高速速度
Dim MAccel(Max_Axes_Num) As Double      '加速度
Dim MDec(Max_Axes_Num) As Double        '减速度
Dim maxpeed(Max_Axes_Num) As Double

Dim RaxesNum(Max_Axes_Num) As Long
Dim RconSpeed(Max_Axes_Num) As Double                      '特殊功能运动参数
Dim RDis(Max_Axes_Num) As Long                            '运动距离
Dim RTimes(Max_Axes_Num) As Long                          '往复次数
Dim lastRTimes(Max_Axes_Num) As Long                      '剩余往复次数

Private Sub Command1_Click()
  If SSTab1.Tab = 0 Then
    GetMultAxesParam
    StartMultAxesMove
  End If
  If SSTab1.Tab = 1 Then
  Timer3.Enabled = True
     GetSpecialFunctionParam
     StartSpecialFunctionMove
  End If
End Sub

Private Sub Command2_Click()
Dim i As Long
 Timer3.Enabled = False
 Timer1.Enabled = False
     For i = 1 To glTotalAxes
         sudden_stop i
     Next i
End Sub

Private Sub Command3_Click()
 Timer1.Enabled = True
 Dim i As Integer
    For i = 0 To 1
       set_maxspeed MAxesNum(i), 20000
       set_conspeed MAxesNum(i), 10000
 Next i
    con_vmove2 MAxesNum(0), 1, MAxesNum(1), -1
End Sub

Private Sub Command4_Click()
Dim opx As Long
Dim opy As Long
opx = check_done(2)
opy = check_done(4)
If opx = 0 And opy = 0 Then
reset_pos 2
reset_pos 4
End If
End Sub

Private Sub disspeed_Change()
GetSpecialFunctionParam
speed(1).Text = RconSpeed(1)
dis(1).Text = RDis(1)
End Sub

Private Sub Form_Load()
 InitBoard
 GetMultAxesParam
 Dim i As Integer
    For i = 0 To 1
       set_conspeed MAxesNum(i), 10000
   Next i
   con_vmove2 MAxesNum(0), 1, MAxesNum(1), -1
End Sub
Private Function GetMultAxesParam() As Long  '读取各轴参数
    Dim i As Long
        For i = 0 To Max_Axes_Num
            MAxesNum(i) = Val(AxesNum(i).Text)
            MLowSpeed(i) = Val(InVel(i).Text)
            MHighSpeed(i) = Val(HighSpeed(i).Text)
            MDis(i) = Val(Distance(i).Text)
            MAccel(i) = Val(Acce(i).Text)
            MDec(i) = Val(Dece(i).Text)
            maxpeed(i) = Val(Maxspeed.Text)
        Next i
    End Function
    Private Function StartMultAxesMove()    '运动参数赋值
  Dim i As Integer
      For i = 0 To 1
      set_maxspeed MAxesNum(i), maxpeed(i)
      set_profile MAxesNum(i), MLowSpeed(i), MHighSpeed(i), MAccel(i), MDec(i)
  Next i
      fast_pmove2 MAxesNum(0), MDis(0), MAxesNum(1), MDis(1)
End Function

    Private Function GetSpecialFunctionParam() As Long  '往复运动参数
    Dim i As Long                                       '读取各轴参数
    Dim a As Double
    Dim b As Double
    For i = 0 To Max_Axes_Num
        RaxesNum(i) = Val(num(i).Text)
        RTimes(i) = Val(time.Text)
        maxpeed(i) = Rmaxpeed.Text
        set_maxspeed RaxesNum(i), maxpeed(i)
    Next i
        b = Val(disspeed.Text)
        RDis(0) = Val(dis(0).Text)
        RDis(1) = b + RDis(0)
        RconSpeed(0) = Val(speed(0).Text)
        a = RDis(0) / RconSpeed(0)
        RconSpeed(1) = Int(RDis(1) / a)
    End Function
    Private Function StartSpecialFunctionMove() '根据运动参数设置速度
    Dim i As Double
    
    For i = 0 To Max_Axes_Num
            set_conspeed RaxesNum(i), RconSpeed(i)
            lastRTimes(i) = RTimes(i)
    Next i
          '启动运动
          con_pmove2 RaxesNum(0), RDis(0), RaxesNum(1), RDis(1)
   Timer3.Enabled = True
End Function
 Private Function SetBoard() As Long
    Dim i As Long
    i = 1
    glTotalAxes = auto_set()            '对板卡进行自动设置
    If glTotalAxes <= 0 Then            '若自动设置错误则返回-1
        If glTotalAxes = -1 Then
            SetBoard = -1
        ElseIf glTotalAxes = -10 Then
            SetBoard = -10
        Else
            SetBoard = -2
        End If
        Exit Function
    End If
    glTotalCards = init_board            '初始化板卡
    If glTotalCards <= 0 Then            '若自动设置错误则返回-2
        If glTotalCards = -5 Then
            SetBoard = -5
        ElseIf glTotalCards = -6 Then
            SetBoard = -6
        ElseIf glTotalCards = -10 Then
            SetBoard = -10
        Else
            SetBoard = -2
        End If
        Exit Function
    End If
    SetBoard = 0                '正确返回0
      For i = 1 To glTotalAxes
        set_alm_logic i, 0
        set_el_logic i, 0
        set_org_logic i, 0
    Next i
   For i = 1 To glTotalCards
        set_card_alm_logic i, 0
    Next i
End Function
Public Sub InitBoard()   '初始化板卡
    Dim Temp As Long
    Temp = SetBoard
    Select Case Temp
        Case -1
            MsgBox "自动设置错误,请检查MPC08运动控制卡是否正确插入您的计算机" & glTotalAxes & "轴", vbOKOnly + vbCritical

        Case -2
            MsgBox "初始化错误,请检查MPC08运动控制卡是否正确插入您的计算机,并检查卡的版本号", vbOKOnly + vbCritical
            
        Case -5
            MsgBox "软件版本错误,请检查MPC08函数库与驱动程序版本", vbOKOnly + vbCritical
           
        Case -6
            MsgBox "型号错误,请检查MPC08运动控制卡型号是否正确", vbOKOnly + vbCritical
           
        Case -10
            MsgBox "多卡型号搭配错误,请检查MPC08运动控制卡型号是否正确", vbOKOnly + vbCritical
           
        Case 0
            MsgBox "初始化成功", vbOKOnly
            
    End Select
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim opx As Long
Dim opy As Long
opx = check_done(2)
opy = check_done(4)
If opx = 0 And opy = 0 Then
   For i = 0 To 1
   set_maxspeed MAxesNum(i), 20000
      set_conspeed MAxesNum(i), 10000
   Next i
  con_pmove2 MAxesNum(0), -80000, MAxesNum(1), 80000
  Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Dim mPosa As Double
Dim mposb As Double
    get_abs_pos 2, mPosa
    Label11.Caption = mPosa
    get_abs_pos 4, mposb
    Label12.Caption = mposb
End Sub

Private Sub Timer3_Timer()
Dim i As Integer
 Dim opx As Long
 Dim opy As Long
 Dim rconsp As Long
 Dim rconds As Long
 opx = check_done(2)
 opy = check_done(4)
 If opx = 0 And opy = 0 Then
        lastRTimes(0) = lastRTimes(0) - 1
        If lastRTimes(0) > 0 Then
        rconsp = RconSpeed(0)
        RconSpeed(0) = RconSpeed(1)
        RconSpeed(1) = rconsp
        
        rconds = RDis(0)
        RDis(0) = RDis(1)
        RDis(1) = rconds
             
        For i = 0 To 1
            RDis(i) = -RDis(i)
             set_conspeed RaxesNum(i), RconSpeed(i)
            Next i
        con_pmove2 RaxesNum(0), RDis(0), RaxesNum(1), RDis(1)
         
        Else
            Timer3.Enabled = False
        End If
    End If
End Sub
