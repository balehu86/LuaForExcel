' ============================================
' Scheduler.bas - 进程级调度器
' ============================================
' 设计原因：
' 1. 进程级唯一，不依赖具体 Workbook
' 2. 只通过 WorkbookRuntime 接口驱动，不知道 Task/Lua
' 3. 使用 OnTime 实现异步调度
' ============================================

Option Explicit

' 全局状态（仅调度器可见）
Private g_Runnables As Collection      ' WorkbookRuntime 集合
Private g_SchedulerRunning As Boolean
Private g_SchedulerIntervalMs As Integer
Private g_NextTaskId As Long
' 配置常量
Private Const DEFAULT_INTERVAL_Ms As Integer = 1000  ' 1000 毫秒

' ====公共接口====
' 初始化调度器
Public Sub InitScheduler()
    If g_Runnables Is Nothing Then Set g_Runnables = New Collection
    If g_SchedulerIntervalSec Is Nothing Then g_SchedulerIntervalSec = DEFAULT_INTERVAL_SEC
    'g_NextTaskId = 0
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
    If g_Runnables Is Nothing Then InitScheduler
    If g_SchedulerRunning Then
        MessageBox "调度器已启动", vbInformation
        Exit Sub
    End If
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
Public Sub SetSchedulerInterval(intervalMs As Double)
    If intervalMs < 10 Or intervalMs > 3600000 Then
        Err.Raise vbObjectError + 1, , "间隔必须在 10-3600000 毫秒之间"
    End If
    g_SchedulerIntervalMs = intervalMs
End Sub
' 生成唯一 TaskId（工具函数，供 WorkbookRuntime 使用）
Public Function GenerateTaskId(wbKey As String, cellAddr As String) As String
    GenerateTaskId = "TASK_" & g_NextTaskId & "_" & wbKey & "_" & cellAddr
    g_NextTaskId = g_NextTaskId + 1
End Function

' ====内部实现====
' 调度心跳（OnTime 回调）
Public Sub SchedulerTick()
    On Error Resume Next
    
    ' 检查运行标志
    If Not g_SchedulerRunning Then Exit Sub
    If g_Runnables Is Nothing Then Exit Sub
    
    ' 关键设计：调度器只知道接口，不知道实现
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
    If g_Runnables.Count > 0 Then
        Application.OnTime EarliestTime:=Now + TimeValue("00:00:01") * g_SchedulerIntervalMs / 1000, Procedure:="Scheduler.SchedulerTick", , False
    Else
        g_SchedulerRunning = False
    End If
End Sub
