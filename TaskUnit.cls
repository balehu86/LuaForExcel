' TaskUnit.cls
Option Explicit

Public taskId As Integer         ' taskId int
Public taskFunc As String        ' taskId -> func name
Public taskWorkbook As String    ' taskId -> workbookInfo instance
Public taskStartArgs As Variant  ' taskId -> startArgs array
Public taskResumeSpec As Variant ' taskId -> resumeSpec array
Public taskCell As String        ' taskId -> taskCell address
Public taskStatus As String      ' taskId -> status string:defined、yield、paused、done、error、terminated
Public taskProgress As Double    ' taskId -> progress number
Public taskMessage As String     ' taskId -> message variant
Public taskValue As Variant      ' taskId -> value variant
Public taskError As String       ' taskId -> error message
Public taskCoThread As LongPtr   ' taskId -> coThread LongPtr

Public taskLastTime As Double    ' 任务上次运行时间(ms)
Public taskTotalTime As Double   ' 任务总运行时间(ms)
Public taskTickCount As Long     ' 任务调度次数

Public taskCoRef As Long         ' Lua 注册表引用号（防止 GC）

' === CFS 调度字段 ===
Public CFS_vruntime As Double      ' 虚拟运行时间（核心调度依据）
Public CFS_weight As Double        ' 任务权重（默认1024，越大优先级越高）
Public CFS_lastScheduled As Double ' 上次被调度的时间戳（GetTickCount）

Private Sub Class_Initialize()
    taskCoRef = 0
    taskCoThread = 0
    CFS_vruntime = 0
    CFS_weight = 1024
    CFS_lastScheduled = 0
End Sub

' ===== 释放协程资源（唯一入口）=====
Public Sub ReleaseCoroutine()
    On Error Resume Next
    
    If taskCoRef = 0 Then Exit Sub
    
    ' 获取全局 Lua 状态机
    Dim L As LongPtr
    L = GetLuaState()
    
    If L <> 0 Then
        UnrefFromRegistry L, taskCoRef
        Debug.Print "TaskUnit.ReleaseCoroutine: Task_" & taskId & " 释放协程 Ref=" & taskCoRef
    End If
    
    taskCoRef = 0
    taskCoThread = 0
End Sub

' ===== 检查协程是否有效 =====
Public Function HasValidCoroutine() As Boolean
    HasValidCoroutine = (taskCoRef <> 0 And taskCoThread <> 0)
End Function

' ===== 类销毁时自动释放 =====
Private Sub Class_Terminate()
    ReleaseCoroutine
End Sub
