' WorkbookInfo.cls
Option Explicit

Public WbName As String
Public WbCursor As Integer
Public WbAllowedTickCount As Integer ' 仅mode=1时使用，若设置此值将覆盖此工作簿的g_WorkbookTicks

Public WbLastTime As Double    ' 上次调度时间(ms)
Public WbTotalTime As Double   ' 总运行时间(ms)
Public WbTickCount As Integer   ' 调度次数(统计用)

' Private m_Workbook As Workbook
Private m_Tasks As Object

Private Sub Class_Initialize()
    WbCursor = 0
    WbAllowedTickCount = -1 ' 默认不覆盖
    WbLastTime = 0
    WbTotalTime = 0
    WbTickCount = 0
    Set m_Tasks = CreateObject("Scripting.Dictionary")
End Sub

' 属性访问器
Public Property Get Tasks() As Object
    Set Tasks = m_Tasks
End Property

Public Sub AddTask(taskId As String, task As TaskInfo)
    m_Tasks.Add taskId, task
End Sub

Public Sub RemoveTask(taskId As String)
    On Error Resume Next
    m_Tasks.Remove taskId
End Sub