' TaskUnit.cls
Option Explicit

Public taskId As Integer         ' taskId int
Public taskFunc As String        ' taskId -> func name
Public taskWorkbook As String    ' taskId -> workbookInfo instance
Public taskStartArgs As Variant  ' taskId -> startArgs array
Public taskResumeSpec As Variant ' taskId -> resumeSpec array
Public taskCell As String        ' taskId -> taskCell address
Public taskStatus As CoStatus ' taskId -> status string:defined、yield、paused、done、error、terminated
Public taskProgress As Double    ' taskId -> progress number
Public taskMessage As String     ' taskId -> message variant
Public taskValue As Variant      ' taskId -> value variant
Public taskError As String       ' taskId -> error message
Public taskCoThread As LongPtr   ' taskId -> coThread LongPtr
Public taskCoRef As Long         ' Lua 注册表引用号（防止 GC）

Public taskLastTime As Double    ' 任务上次运行时间(ms)
Public taskTotalTime As Double   ' 任务总运行时间(ms)
Public taskTickCount As Long     ' 任务调度次数
' === CFS 调度字段 ===
Public CFS_vruntime As Double      ' 虚拟运行时间（核心调度依据）
Public CFS_weight As Double        ' 任务权重（默认1024，越大优先级越高）
Public CFS_lastScheduled As Double ' 上次被调度的时间戳（GetTickCount）

Public taskWatches As Collection ' 任务监视的Watchcell字符串的集合

Private Sub Class_Initialize()
    taskCoRef = 0
    taskCoThread = 0
    CFS_vruntime = 0
    CFS_lastScheduled = 0
    Set taskWatches = New Collection
End Sub

' 清除协程引用（仅清除本地存储，不执行释放）
Public Sub ClearCoroutineRef()
    taskCoRef = 0
    taskCoThread = 0
End Sub

' 添加监控单元格到任务的监控列表
Public Sub AddWatche(watchCell As String)
    ' 检查是否已存在
    Dim exists As Boolean
    exists = False
    On Error Resume Next
    Dim i As Long
    For i = 1 To Me.taskWatches.Count
        If Me.taskWatches(i) = watchCell Then
            exists = True
            Exit For
        End If
    Next
    On Error GoTo 0
    ' 不存在才添加
    If Not exists Then
        Me.taskWatches.Add watchCell
    End If
End Sub

' 从任务的监控列表中移除监控单元格
Public Sub RemoveWatches(watchCell As String)
    On Error Resume Next
    Dim i As Long
    For i = Me.taskWatches.Count To 1 Step -1
        If Me.taskWatches(i) = watchCell Then
            Me.taskWatches.Remove i
            Exit For
        End If
    Next
End Sub
