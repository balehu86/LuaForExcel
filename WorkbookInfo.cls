Public Name As String
Public Cursor As Integer
Public AllowedTickCount As Integer ' 仅mode=1时使用，若设置此值将覆盖此工作簿的g_WorkbookTicks
Public Workbook As Object ' workbookName -> Workbook instance
Public LastTime As Double    ' 上次调度时间(ms)
Public TotalTime As Double   ' 总运行时间(ms)
Public TickCount As Integer   ' 调度次数(统计用)
Private m_Tasks As Object
Private Sub Class_Initialize()
    Cursor = 0
    AllowedTickCount = -1 ' 默认不覆盖
    LastTime = 0
    TotalTime = 0
    TickCount = 0
    Set m_Tasks = CreateObject("Scripting.Dictionary")
End Sub

' 属性访问器
Public Property Get Tasks() As Object
    Set Tasks = m_Tasks
End Property

' 可选：设置器
Public Property Set Dictionary(ByVal dict As Object)
    Set pDict = dict
End Property

Public Sub AddTask(task As TaskInfo)
    m_Tasks.Add task
End Sub

Public Sub RemoveTask(taskId As Integer)
    On Error Resume Next
    m_Tasks.Remove CStr(taskId)
End Sub