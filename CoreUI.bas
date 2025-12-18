' ============================================
' CoreUI.bas - 用户界面模块
' ============================================
' 设计原因：
' 1. 统一管理菜单创建/销毁
' 2. OnAction 路由到对应 WorkbookRuntime
' 3. 提供日志/消息框工具
' 4. 不访问 TaskTable，不修改 Task 状态
' ============================================

Option Explicit

' ===== 日志工具 =====
Public Sub LogInfo(msg As String)
    Debug.Print "[INFO] " & msg
End Sub

Public Sub LogError(msg As String)
    Debug.Print "[ERROR] " & msg
    MsgBox msg, vbCritical, "错误"
End Sub

' ============================================
' 菜单管理
' ============================================

' 启用 Lua 任务菜单
Public Sub EnableLuaTaskMenu()
    On Error Resume Next
    
    DisableLuaTaskMenu
    
    Dim cMenu As CommandBar
    Set cMenu = Application.CommandBars("Cell")
    
    ' 任务管理菜单
    Dim taskMenu As CommandBarControl
    Set taskMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    taskMenu.Caption = "Lua 任务管理"
    taskMenu.Tag = "LuaTaskMenu"
    
    AddMenuItem taskMenu, "启动任务", "OnAction_StartTask"
    AddMenuItem taskMenu, "暂停任务", "OnAction_PauseTask"
    AddMenuItem taskMenu, "恢复任务", "OnAction_ResumeTask"
    AddMenuItem taskMenu, "终止任务", "OnAction_TerminateTask"
    AddMenuItem taskMenu, "查看任务详情", "OnAction_ShowTaskDetail"
    
    ' 调度管理菜单
    Dim schedulerMenu As CommandBarControl
    Set schedulerMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    schedulerMenu.Caption = "Lua 调度管理"
    schedulerMenu.Tag = "LuaSchedulerMenu"
    
    AddMenuItem schedulerMenu, "启动调度器", "OnAction_StartScheduler"
    AddMenuItem schedulerMenu, "停止调度器", "OnAction_StopScheduler"
    AddMenuItem schedulerMenu, "设置调度间隔", "OnAction_SetSchedulerInterval"
    
    ' 配置菜单
    Dim configMenu As CommandBarControl
    Set configMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    configMenu.Caption = "Lua 配置管理"
    configMenu.Tag = "LuaConfigMenu"
    
    AddMenuItem configMenu, "重载 functions.lua", "OnAction_ReloadFunctions"
End Sub

' 禁用菜单
Public Sub DisableLuaTaskMenu()
    On Error Resume Next
    Dim cMenu As CommandBar
    Set cMenu = Application.CommandBars("Cell")
    
    Dim ctrl As CommandBarControl
    For Each ctrl In cMenu.Controls
        If ctrl.Tag = "LuaTaskMenu" Or _
           ctrl.Tag = "LuaSchedulerMenu" Or _
           ctrl.Tag = "LuaConfigMenu" Then
            ctrl.Delete
        End If
    Next
End Sub

' 添加菜单项
Private Sub AddMenuItem(parent As CommandBarControl, caption As String, onAction As String)
    Dim ctrl As CommandBarControl
    Set ctrl = parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    ctrl.Caption = caption
    ctrl.OnAction = onAction
End Sub

' ============================================
' OnAction 路由（关键设计）
' ============================================

' 获取当前选中单元格的运行时
Private Function GetRuntimeFromSelection() As IRuntime
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Set GetRuntimeFromSelection = Nothing
        Exit Function
    End If
    
    Set GetRuntimeFromSelection = CoreRegistry.GetRuntimeByWorkbook(wb)
End Function

' 获取当前选中单元格的 TaskId
Private Function GetTaskIdFromSelection() As String
    On Error Resume Next
    Dim cellAddr As String
    cellAddr = Selection.Address(External:=True)
    
    ' 从单元格值读取 TaskId（假设单元格包含 =LuaTask(...) 返回的 taskId）
    Dim taskId As String
    taskId = CStr(Selection.Value)
    
    If Left(taskId, 5) = "TASK_" Then
        GetTaskIdFromSelection = taskId
    Else
        GetTaskIdFromSelection = ""
    End If
End Function

' ============================================
' 菜单回调函数
' ============================================

Public Sub OnAction_StartTask()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then
        MsgBox "未找到工作簿运行时", vbExclamation
        Exit Sub
    End If
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then
        MsgBox "当前单元格没有任务", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    rt.StartTask taskId
    MsgBox "任务已启动: " & taskId, vbInformation
    Exit Sub
    
ErrorHandler:
    LogError "启动任务失败: " & Err.Description
End Sub

Public Sub OnAction_PauseTask()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then Exit Sub
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then Exit Sub
    
    rt.PauseTask taskId
    MsgBox "任务已暂停", vbInformation
End Sub

Public Sub OnAction_ResumeTask()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then Exit Sub
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then Exit Sub
    
    rt.ResumeTaskManual taskId
    MsgBox "任务已恢复", vbInformation
End Sub

Public Sub OnAction_TerminateTask()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then Exit Sub
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then Exit Sub
    
    Dim result As VbMsgBoxResult
    result = MsgBox("确定终止任务？", vbYesNo + vbExclamation)
    If result = vbYes Then
        rt.TerminateTask taskId
        MsgBox "任务已终止", vbInformation
    End If
End Sub

Public Sub OnAction_ShowTaskDetail()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then Exit Sub
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then Exit Sub
    
    Dim msg As String
    msg = "任务详情" & vbCrLf & vbCrLf
    msg = msg & "TaskId: " & taskId & vbCrLf
    msg = msg & "状态: " & rt.GetTaskField(taskId, "status") & vbCrLf
    msg = msg & "进度: " & rt.GetTaskField(taskId, "progress") & "%" & vbCrLf
    msg = msg & "消息: " & rt.GetTaskField(taskId, "message") & vbCrLf
    
    MsgBox msg, vbInformation, "任务详情"
End Sub

Public Sub OnAction_StartScheduler()
    Scheduler.StartScheduler
    MsgBox "调度器已启动", vbInformation
End Sub

Public Sub OnAction_StopScheduler()
    Scheduler.StopScheduler
    MsgBox "调度器已停止", vbInformation
End Sub

Public Sub OnAction_SetSchedulerInterval()
    Dim input As String
    input = InputBox("请输入调度间隔（秒）：", "设置调度间隔", "1")
    
    If input = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    Dim intervalSec As Double
    intervalSec = CDbl(input)
    
    Scheduler.SetSchedulerInterval intervalSec
    MsgBox "调度间隔已设置为 " & intervalSec & " 秒", vbInformation
    Exit Sub
    
ErrorHandler:
    LogError "设置失败: " & Err.Description
End Sub

Public Sub OnAction_ReloadFunctions()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then
        MsgBox "未找到工作簿运行时", vbExclamation
        Exit Sub
    End If
    
    If rt.ReloadFunctions() Then
        MsgBox "functions.lua 已重载", vbInformation
    Else
        MsgBox "重载失败", vbCritical
    End If
End Sub