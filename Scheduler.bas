' ============================================
' Scheduler.bas - 进程级调度器
' ============================================
' 设计原因：
' 1. 进程级唯一，不依赖具体 Workbook
' 2. 只通过 WorkbookRuntime 接口驱动，不知道 Task/Lua
' 3. 使用 OnTime 实现异步调度
' ============================================

Option Explicit

' ===== 全局状态（仅调度器可见）=====
Private g_Runnables As Collection      ' WorkbookRuntime 集合
Private g_SchedulerRunning As Boolean
Private g_SchedulerIntervalSec As Double
Private g_NextTaskId As Long

' ===== 配置常量 =====
Private Const DEFAULT_INTERVAL_SEC As Double = 1#  ' 1 秒

' ============================================
' 公共接口
' ============================================

' 初始化调度器
Public Sub InitScheduler()
    If g_Runnables Is Nothing Then
        Set g_Runnables = New Collection
        g_SchedulerIntervalSec = DEFAULT_INTERVAL_SEC
        g_NextTaskId = 0
    End If
End Sub

' 注册运行时（由 CoreRegistry 调用）
Public Sub RegisterRunnable(rt As WorkbookRuntime)
    If g_Runnables Is Nothing Then InitScheduler
    On Error Resume Next
    g_Runnables.Add rt
End Sub

' 注销运行时
Public Sub UnregisterRunnable(rt As WorkbookRuntime)
    On Error Resume Next
    Dim i As Long
    For i = g_Runnables.Count To 1 Step -1
        If g_Runnables(i) Is rt Then
            g_Runnables.Remove i
            Exit For
        End If
    Next
End Sub

' 启动调度器
Public Sub StartScheduler()
    If g_SchedulerRunning Then Exit Sub
    If g_Runnables Is Nothing Then InitScheduler
    
    g_SchedulerRunning = True
    SchedulerTick
End Sub

' 停止调度器
Public Sub StopScheduler()
    g_SchedulerRunning = False
    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                       Procedure:="Scheduler.SchedulerTick", _
                       Schedule:=False
End Sub

' 设置调度间隔
Public Sub SetSchedulerInterval(intervalSec As Double)
    If intervalSec < 0.01 Or intervalSec > 3600 Then
        Err.Raise vbObjectError + 1, , "间隔必须在 0.01-3600 秒之间"
    End If
    g_SchedulerIntervalSec = intervalSec
End Sub

' 生成唯一 TaskId（工具函数，供 WorkbookRuntime 使用）
Public Function GenerateTaskId(wbKey As String, cellAddr As String) As String
    GenerateTaskId = "TASK_" & g_NextTaskId & "_" & wbKey & "_" & cellAddr
    g_NextTaskId = g_NextTaskId + 1
End Function

' ============================================
' 内部实现
' ============================================

' 调度心跳（OnTime 回调）
Public Sub SchedulerTick()
    On Error Resume Next
    
    ' 检查运行标志
    If Not g_SchedulerRunning Then Exit Sub
    If g_Runnables Is Nothing Then Exit Sub
    
    ' 关键设计：调度器只知道接口，不知道实现
    ' 20 行核心逻辑，符合架构要求
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim rt As WorkbookRuntime
    Dim needsRefresh As Boolean
    needsRefresh = False
    
    ' 遍历所有运行时
    For Each rt In g_Runnables
        If rt.HasRunnable Then
            rt.Tick
        needsRefresh = True
        End If
    Next rt
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' 只有状态变化时才刷新
    If needsRefresh Then
        On Error Resume Next
        ActiveSheet.Calculate
    End If
    
    ' 继续调度
    If g_SchedulerRunning And g_Runnables.Count > 0 Then
        Application.OnTime EarliestTime:=Now + g_SchedulerIntervalSec / 86400#, Procedure:="Scheduler.SchedulerTick"
    Else
        g_SchedulerRunning = False
    End If
End Sub
