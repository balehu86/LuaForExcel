Option Explicit

Public taskId As Integer         ' taskId int
Public taskFunc As String        ' taskId -> func name
Public taskWorkbook As String ' taskId -> workbookInfo instance
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
    ' 初始化 CFS 字段
    CFS_vruntime = 0
    CFS_weight = 1024        ' 默认权重，与 Linux CFS 的 nice=0 对应
    CFS_lastScheduled = 0
End Sub

