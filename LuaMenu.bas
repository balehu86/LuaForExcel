' ============================================
' LuaMenu.bas - 用户界面模块
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

' ====菜单管理====
' 创建右键菜单
Public Sub EnableLuaTaskMenu()
    On Error Resume Next
    DisableLuaTaskMenu

    Dim cellMenu As CommandBar
    Set cellMenu = Application.CommandBars("Cell")
    
    ' 任务管理菜单
    Dim luaTaskMenu As CommandBarControl
    Set luaTaskMenu = cellMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaTaskMenu.Caption = "Lua 任务管理"
    luaTaskMenu.Tag = "LuaTaskMenu"
    AddLuaMenuItem luaTaskMenu, "启动任务", "LuaTaskMenu_StartTask"
    AddLuaMenuItem luaTaskMenu, "暂停任务", "LuaTaskMenu_PauseTask"
    AddLuaMenuItem luaTaskMenu, "恢复任务", "LuaTaskMenu_ResumeTask"
    AddLuaMenuItem luaTaskMenu, "终止任务", "LuaTaskMenu_TerminateTask"
    AddLuaMenuItem luaTaskMenu, "查看任务详情", "LuaTaskMenu_ShowTaskDetail"

    ' 调度管理菜单
    Dim luaSchedulerMenu As CommandBarControl
    Set luaSchedulerMenu = cellMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaSchedulerMenu.Caption = "Lua 调度管理"
    luaSchedulerMenu.Tag = "LuaSchedulerMenu"
    AddLuaMenuItem luaSchedulerMenu, "启动所有 defined 任务", "LuaSchedulerMenu_StartAllDefinedTasks"
    AddLuaMenuItem luaSchedulerMenu, "清理所有完成、错误任务", "LuaSchedulerMenu_CleanupFinishedTasks"
    AddLuaMenuItem luaSchedulerMenu, "删除所有任务", "LuaSchedulerMenu_ClearAllTasks"
    AddLuaMenuItem luaSchedulerMenu, "显示所有任务信息", "LuaSchedulerMenu_ShowAllTasks"
    AddLuaMenuItem luaSchedulerMenu, "启动调度器", "LuaSchedulerMenu_StartScheduler"
    AddLuaMenuItem luaSchedulerMenu, "停止调度器", "LuaSchedulerMenu_StopScheduler"

    ' 设置管理菜单
    Dim luaConfigMenu As CommandBarControl
    Set luaConfigMenu = cellMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaConfigMenu.Caption = "Lua 设置管理"
    luaConfigMenu.Tag = "luaConfigMenu"
    AddLuaMenuItem luaConfigMenu, "启用热重载", "LuaConfigMenu_EnableHotReload"
    AddLuaMenuItem luaConfigMenu, "禁用热重载", "LuaConfigMenu_DisableHotReload"
    AddLuaMenuItem luaConfigMenu, "手动重载 functions.lua", "LuaConfigMenu_ReloadFunctions"
    AddLuaMenuItem luaConfigMenu, "设置调度间隔（秒）", "LuaConfigMenu_SetSchedulerInterval"

    MsgBox "Lua 任务右键菜单已启用。", vbInformation
End Sub

Public Sub DisableLuaTaskMenu()
    On Error Resume Next
    
    Dim cellMenu As CommandBar
    Set cellMenu = Application.CommandBars("Cell")
    
    Dim ctrl As CommandBarControl
    For Each ctrl In cellMenu.Controls
        If ctrl.Tag = "LuaTaskMenu" Then ctrl.Delete
        If ctrl.Tag = "LuaSchedulerMenu" Then ctrl.Delete
        If ctrl.Tag = "LuaConfigMenu" Then ctrl.Delete
    Next
End Sub

Private Sub AddLuaMenuItem(parent As CommandBarControl, caption As String, onAction As String)
    Dim ctrl As CommandBarControl
    Set ctrl = parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    ctrl.Caption = caption
    ctrl.OnAction = onAction
End Sub

Private Function GetRuntimeFromSelection() As WorkbookRuntime
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Set GetRuntimeFromSelection = Nothing
        Exit Function
    End If
    
    Set GetRuntimeFromSelection = CoreRegistry.GetRuntimeByWorkbook(wb)
End Function

Private Function GetTaskIdFromSelection() As String
    On Error Resume Next
    Dim taskId As String
    taskId = CStr(Selection.Value)
    
    If Left(taskId, 5) = "TASK_" Then
        GetTaskIdFromSelection = taskId
    Else
        GetTaskIdFromSelection = ""
    End If
End Function

' ============================================
' 任务管理回调函数
' ============================================

Public Sub LuaTaskMenu_StartTask()
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

Public Sub LuaTaskMenu_PauseTask()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then Exit Sub
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then Exit Sub
    
    rt.PauseTask taskId
    MsgBox "任务已暂停", vbInformation
End Sub

Public Sub LuaTaskMenu_ResumeTask()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then Exit Sub
    
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then Exit Sub
    
    rt.ResumeTaskManual taskId
    MsgBox "任务已恢复", vbInformation
End Sub

Public Sub LuaTaskMenu_TerminateTask()
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

Public Sub LuaTaskMenu_ShowTaskDetail()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then
        MsgBox "未找到工作簿运行时", vbExclamation
        Exit Sub
    End If

    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If

    Dim info As Object
    Set info = rt.GetTaskDetail(taskId)
    If info Is Nothing Then
        MsgBox "任务不存在或已被清理。", vbExclamation
        Exit Sub
    End If

    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  Lua 任务详情" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf

    msg = msg & "任务 ID: " & info("taskId") & vbCrLf
    msg = msg & "函数名: " & info("funcName") & vbCrLf
    msg = msg & "单元格: " & info("cellAddr") & vbCrLf
    msg = msg & "状态: " & info("status") & vbCrLf
    msg = msg & "进度: " & Format(info("progress"), "0.00") & "%" & vbCrLf

    If Not IsEmpty(info("message")) Then
        msg = msg & "消息: " & CStr(info("message")) & vbCrLf
    End If

    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf

    ' 返回值
    If IsArray(info("value")) Then
        msg = msg & "返回值: (数组)" & vbCrLf
    ElseIf IsEmpty(info("value")) Then
        msg = msg & "返回值: (空)" & vbCrLf
    Else
        msg = msg & "返回值: " & CStr(info("value")) & vbCrLf
    End If

    ' 错误信息
    If info("status") = "error" Then
        msg = msg & vbCrLf & "错误信息:" & vbCrLf
        msg = msg & CStr(info("error")) & vbCrLf
    End If

    ' 协程
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf
    msg = msg & "协程线程: "
    If info("coThread") = 0 Then
        msg = msg & "未创建" & vbCrLf
    Else
        msg = msg & "0x" & Hex(info("coThread")) & vbCrLf
    End If

    MsgBox msg, vbInformation, "Lua 任务详情"
End Sub

' ============================================
' 调度管理回调函数
' ============================================

Public Sub LuaSchedulerMenu_StartAllDefinedTasks()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then
        MsgBox "未找到工作簿运行时", vbExclamation
        Exit Sub
    End If
    
    ' 注意：这需要在 WorkbookRuntime 中添加 GetAllTasks 方法
    ' 当前简化实现：提示用户手动启动
    MsgBox "请在需要启动的任务单元格上右键选择 '启动任务'" & vbCrLf & _
           "批量启动功能需要额外实现 WorkbookRuntime.GetAllTasks()", _
           vbInformation
End Sub

Public Sub LuaSchedulerMenu_CleanupFinishedTasks()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then
        MsgBox "未找到工作簿运行时", vbExclamation
        Exit Sub
    End If
    rt.CleanUpFinishedTasks
End Sub

Public Sub LuaSchedulerMenu_ClearAllTasks()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then
        MsgBox "未找到工作簿运行时", vbExclamation
        Exit Sub
    End If
    
    Dim result As VbMsgBoxResult
    result = MsgBox("确定删除所有任务？此操作不可撤销！", vbYesNo + vbExclamation)
    If result <> vbYes Then Exit Sub
    
    rt.ClearAllTasks
    MsgBox "所有任务已删除", vbInformation
End Sub

Public Sub LuaSchedulerMenu_ShowAllTasks()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then
        MsgBox "未找到工作簿运行时", vbExclamation
        Exit Sub
    End If
    
    MsgBox rt.GetAllTasksInfo(), vbInformation, "Lua 任务列表"
End Sub

Public Sub LuaSchedulerMenu_StartScheduler()
    Scheduler.StartScheduler
    MsgBox "调度器已启动", vbInformation
End Sub

Public Sub LuaSchedulerMenu_StopScheduler()
    Scheduler.StopScheduler
    MsgBox "调度器已停止", vbInformation
End Sub

' ============================================
' 设置管理回调函数
' ============================================

Public Sub LuaConfigMenu_EnableHotReload()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then
        MsgBox "未找到工作簿运行时", vbExclamation
        Exit Sub
    End If
    
    ' 注意：需要在 WorkbookRuntime 中添加 SetHotReloadEnabled 方法
    ' 或者直接访问 m_HotReloadEnabled（需要改为 Public Property）
    MsgBox "热重载功能需要在 WorkbookRuntime 中添加 SetHotReloadEnabled() 方法" & vbCrLf & _
           "当前热重载默认启用", _
           vbInformation
End Sub

Public Sub LuaConfigMenu_DisableHotReload()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then
        MsgBox "未找到工作簿运行时", vbExclamation
        Exit Sub
    End If
    
    MsgBox "热重载功能需要在 WorkbookRuntime 中添加 SetHotReloadEnabled() 方法" & vbCrLf & _
           "临时解决方案：手动修改 WorkbookRuntime 的 DEFAULT_HOT_RELOAD_ENABLED 常量", _
           vbInformation
End Sub

Public Sub LuaConfigMenu_ReloadFunctions()
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeFromSelection()
    If rt Is Nothing Then
        MsgBox "未找到工作簿运行时", vbExclamation
        Exit Sub
    End If
    
    If rt.ReloadFunctions() Then
        MsgBox "functions.lua 已重载", vbInformation
    Else
        MsgBox "重载失败，请检查 functions.lua 是否存在语法错误", vbCritical
    End If
End Sub

Public Sub LuaConfigMenu_SetSchedulerInterval()
    Dim i As String
    i = InputBox("请输入调度间隔（秒）：", "设置调度间隔", "1")
    
    If i = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    Dim intervalSec As Double
    intervalSec = CDbl(i)
    If intervalSec < 0.01 Or intervalSec > 3600 Then
        MsgBox "间隔必须在 0.01-3600 秒之间", vbExclamation
        Exit Sub
    End If
    
    Scheduler.SetSchedulerInterval intervalSec * 1000
    MsgBox "调度间隔已设置为 " & intervalSec & " 秒", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "输入无效：" & Err.Description, vbCritical
End Sub