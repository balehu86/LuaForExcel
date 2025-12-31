Option Explicit

Public TaskId As Integer         ' taskId int
Public TaskFunc As String        ' taskId -> func name
Public TaskWorkbook As String ' taskId -> workbookInfo instance
Public TaskStartArgs As Variant  ' taskId -> startArgs array
Public TaskResumeSpec As Variant ' taskId -> resumeSpec array
Public TaskCell As String        ' taskId -> taskCell address
Public TaskStatus As String      ' taskId -> status string:defined、yield、done、error、terminated、paused
Public TaskProgress As Double    ' taskId -> progress number
Public TaskMessage As String     ' taskId -> message variant
Public TaskValue As Variant      ' taskId -> value variant
Public TaskError As String       ' taskId -> error message
Public TaskCoThread As LongPtr   ' taskId -> coThread LongPtr

Public TaskLastTime As Double    ' 任务上次运行时间(ms)
Public TaskTotalTime As Double   ' 任务总运行时间(ms)
Public TaskTickCount As Long     ' 任务调度次数

Private Sub Class_Initialize()
    ' Debug.Print "[INFO] " & "一个新的TaskInfo实例被创建"
End Sub

