Option Explicit

Public Id As Integer         ' taskId int
Public Func As String        ' taskId -> func name
Public Workbook As WorkbookInfo ' taskId -> workbookInfo instance
Public StartArgs As Variant  ' taskId -> startArgs array
Public ResumeSpec As Variant ' taskId -> resumeSpec array
Public Cell As String        ' taskId -> taskCell address
Public Status As String      ' taskId -> status string:defined、yield、done、error、terminated、paused
Public Progress As Double    ' taskId -> progress number
Public Message As String     ' taskId -> message variant
Public Value As Variant      ' taskId -> value variant
Public Error As String       ' taskId -> error message
Public CoThread As LongPtr   ' taskId -> coThread LongPtr

Public LastTime As Double    ' 任务上次运行时间(ms)
Public TotalTime As Double   ' 任务总运行时间(ms)
Public TickCount As Long     ' 任务调度次数

Private Sub Class_Initialize()
    Debug.Print "[INFO] " & "一个新的TaskInfo实例被创建"
End Sub
