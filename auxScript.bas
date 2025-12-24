' ===== 日志工具 =====
Public Sub LogInfo(msg As String)
    Debug.Print "[INFO] " & msg
End Sub
Public Sub LogError(msg As String)
    Debug.Print "[ERROR] " & msg
    MsgBox msg, vbCritical, "错误"
End Sub

' 辅助函数：获取数组维度
Private Function ArrayDimensions(arr As Variant) As String
    On Error Resume Next
    Dim dims As String
    Dim i As Long
    
    For i = 1 To 10  ' 最多检查10维
        Dim ub As Long
        ub = UBound(arr, i)
        If Err.Number <> 0 Then
            Exit For
        End If
        
        Dim lb As Long
        lb = LBound(arr, i)
        
        If dims <> "" Then dims = dims & " x "
        dims = dims & (ub - lb + 1)
    Next i
    
    If dims = "" Then dims = "未知"
    ArrayDimensions = dims
End Function

' 辅助函数：增加子菜单项
Private Sub AddLuaMenuItem(parent As CommandBarControl, caption As String, onAction As String)
    Dim ctrl As CommandBarControl
    Set ctrl = parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    ctrl.Caption = caption
    ctrl.OnAction = onAction
End Sub

' 辅助函数：根据单元格地址获取任务ID
Private Function GetTaskIdFromSelection() As String
    Dim cellAddr As String
    cellAddr = Selection.Address(External:=True)
    GetTaskIdFromSelection = FindTaskByCell(cellAddr)
End Function

' 辅助函数：按工作簿过滤任务
Private Function GetTasksByWorkbook(wbName As String) As Object
    Dim result As Object
    Set result = CreateObject("System.Collections.ArrayList")
    
    If g_TaskWorkbook Is Nothing Then Exit Function
    
    Dim taskId As Variant
    For Each taskId In g_TaskWorkbook.Keys
        If g_TaskWorkbook(CStr(taskId)) = wbName Then
            result.Add taskId
        End If
    Next
    
    Set GetTasksByWorkbook = result
End Function

' 辅助函数：清理特定工作簿的任务
Public Sub CleanupWorkbookTasks(wbName As String)
    On Error Resume Next
    
    If g_TaskWorkbook Is Nothing Then Exit Sub

    Dim tasksToRemove As Object
    Set tasksToRemove = GetTasksByWorkbook(wbName)
    
    Dim i As Long
    For i = 0 To tasksToRemove.Count - 1
        Dim tid As String
        tid = CStr(tasksToRemove(i))
        
        ' 从队列中移除
        If g_TaskQueue.Exists(tid) Then g_TaskQueue.Remove tid
        
        ' 删除所有相关数据
        g_TaskFunc.Remove tid
        g_TaskWorkbook.Remove tid
        g_TaskStartArgs.Remove tid
        g_TaskResumeSpec.Remove tid
        g_TaskCell.Remove tid
        g_TaskStatus.Remove tid
        g_TaskProgress.Remove tid
        g_TaskMessage.Remove tid
        g_TaskValue.Remove tid
        g_TaskError.Remove tid
        g_TaskCoThread.Remove tid
    Next
End Sub  
' ============================================
' 第七部分：可视化操作函数
' ============================================
' 打开并创建右键菜单功能
Public Sub EnableLuaTaskMenu()
    On Error Resume Next

    ' 删除已有菜单，避免重复
    Call DisableLuaTaskMenu

    ' 获取右键菜单（Cell）
    Dim cMenu As CommandBar
    Set cMenu = Application.CommandBars("Cell")

    ' 添加单个任务的主菜单
    Dim luaTaskMenu As CommandBarControl
    Set luaTaskMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaTaskMenu.Caption = "Lua 任务管理"
    luaTaskMenu.Tag = "LuaTaskMenu"
    ' 添加单个任务的子菜单
    AddLuaMenuItem luaTaskMenu, "启动任务", "LuaTaskMenu_StartTask"
    AddLuaMenuItem luaTaskMenu, "启动本簿所有任务", "LuaTaskMenu_StartAllWorkbookTasks"
    AddLuaMenuItem luaTaskMenu, "暂停任务", "LuaTaskMenu_PauseTask"
    AddLuaMenuItem luaTaskMenu, "恢复任务", "LuaTaskMenu_ResumeTask"
    AddLuaMenuItem luaTaskMenu, "终止任务", "LuaTaskMenu_TerminateTask"
    AddLuaMenuItem luaTaskMenu, "查看任务详情", "LuaTaskMenu_ShowDetail"

    ' 添加调度的主菜单
    Dim luaSchedulerMenu As CommandBarControl
    Set luaSchedulerMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaSchedulerMenu.Caption = "Lua 调度管理"
    luaSchedulerMenu.Tag = "LuaSchedulerMenu"
    ' 添加调度的子菜单
    AddLuaMenuItem luaSchedulerMenu, "启动调度器", "LuaSchedulerMenu_StartScheduler"
    AddLuaMenuItem luaSchedulerMenu, "停止调度器", "LuaSchedulerMenu_StopScheduler"
    AddLuaMenuItem luaSchedulerMenu, "启动所有 defined 任务", "LuaSchedulerMenu_StartAllDefinedTasks"
    AddLuaMenuItem luaSchedulerMenu, "清理所有完成、错误任务", "LuaSchedulerMenu_CleanupFinishedTasks"
    AddLuaMenuItem luaSchedulerMenu, "删除此工作簿任务", "LuaSchedulerMenu_CleanupWorkbookTasks"
    AddLuaMenuItem luaSchedulerMenu, "删除所有任务", "LuaSchedulerMenu_ClearAllTasks"
    AddLuaMenuItem luaSchedulerMenu, "显示所有任务信息", "LuaSchedulerMenu_ShowAllTasks"

    ' 添加管理的主菜单
    Dim luaConfigMenu As CommandBarControl
    Set luaConfigMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaConfigMenu.Caption = "Lua 设置管理"
    luaConfigMenu.Tag = "luaConfigMenu"
    ' 添加管理的子菜单
    AddLuaMenuItem luaConfigMenu, "启用热重载", "LuaConfigMenu_EnableHotReload"
    AddLuaMenuItem luaConfigMenu, "禁用热重载", "LuaConfigMenu_DisableHotReload"
    AddLuaMenuItem luaConfigMenu, "手动重载 functions.lua", "LuaConfigMenu_ReloadFunctions"
    AddLuaMenuItem luaConfigMenu, "设置调度间隔（毫秒）", "LuaConfigMenu_SetSchedulerInterval"
    AddLuaMenuItem luaConfigMenu, "设置调度步数", "LuaConfigMenu_SetSchedulerBatchSize"
    AddLuaMenuItem luaConfigMenu, "切换调度模式", "LuaConfigMenu_ToggleScheduleMode"
    AddLuaMenuItem luaConfigMenu, "设置工作簿Tick数", "LuaConfigMenu_SetWorkbookTicks"

    ' 添加调试的主菜单
    Dim luaDebugMenu As CommandBarControl
    Set luaDebugMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaDebugMenu.Caption = "Lua 调试管理"
    luaDebugMenu.Tag = "luaDebugMenu"
    ' 添加调试的子菜单
    AddLuaMenuItem luaDebugMenu, "显示插件状态", "LuaDebugMenu_ShowAddinStatus"

    ' 在右键菜单最后添加性能统计菜单
    Dim luaPerfMenu As CommandBarControl
    Set luaPerfMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaPerfMenu.Caption = "Lua 性能统计"
    luaPerfMenu.Tag = "LuaPerfMenu"

    AddLuaMenuItem luaPerfMenu, "调度器统计", "LuaPerfMenu_ShowSchedulerStats"
    AddLuaMenuItem luaPerfMenu, "任务性能统计", "LuaPerfMenu_ShowTaskStats"
    AddLuaMenuItem luaPerfMenu, "工作簿性能统计", "LuaPerfMenu_ShowWorkbookStats"
    AddLuaMenuItem luaPerfMenu, "重置性能统计", "LuaPerfMenu_ResetStats"
    MsgBox "Lua 任务右键菜单已启用。", vbInformation
End Sub

' 关闭右键菜单
Public Sub DisableLuaTaskMenu()
    On Error Resume Next
    Dim cMenu As CommandBar
    Set cMenu = Application.CommandBars("Cell")

    Dim ctrl As CommandBarControl
    For Each ctrl In cMenu.Controls
        If ctrl.Tag = "LuaTaskMenu" Then ctrl.Delete
        If ctrl.Tag = "LuaSchedulerMenu" Then ctrl.Delete
        If ctrl.Tag = "LuaConfigMenu" Then ctrl.Delete
        If ctrl.Tag = "LuaPerfMenu" Then ctrl.Delete  ' 新增
    Next
End Sub

' 重新加载菜单（修复菜单丢失问题）
Public Sub ReloadMenus()
    DisableLuaTaskMenu
    EnableLuaTaskMenu
    MsgBox "右键菜单已重新加载！", vbInformation, "菜单重载"
End Sub

' 启动任务
Private Sub LuaTaskMenu_StartTask()
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If
    If g_TaskStatus(taskId) = "defined" Then
        StartLuaCoroutine taskId
        MsgBox "任务已启动: " & taskId, vbInformation
    Else
        MsgBox "任务状态为 " & g_TaskStatus(taskId) & "，无法启动。", vbExclamation
    End If
End Sub

' 启动本工作簿的所有defined任务
Private Sub LuaTaskMenu_StartAllWorkbookTasks()
    On Error Resume Next

    ' 获取当前工作簿名称
    Dim wbName As String
    On Error Resume Next
    wbName = ActiveWorkbook.Name
    On Error GoTo ErrorHandler

    If wbName = "" Then
        MsgBox "无法获取当前工作簿。", vbExclamation, "错误"
        Exit Sub
    End If

    If g_TaskFunc Is Nothing Then
        InitCoroutineSystem
    End If

    If g_TaskFunc.Count = 0 Then
        MsgBox "当前没有任何任务。", vbInformation, "提示"
        Exit Sub
    End If

    ' 收集本工作簿的所有defined任务
    Dim taskId As Variant
    Dim count As Long
    count = 0

    For Each taskId In g_TaskFunc.Keys
        If g_TaskWorkbook.Exists(CStr(taskId)) Then
            If g_TaskWorkbook(CStr(taskId)) = wbName Then
                If g_TaskStatus(CStr(taskId)) = "defined" Then
                    StartLuaCoroutine CStr(taskId)
                    count = count + 1
                End If
            End If
        End If
    Next taskId

    If count = 0 Then
        MsgBox "工作簿 [" & wbName & "] 没有 defined 状态的任务。", vbInformation, "提示"
    Else
        MsgBox "已启动工作簿 [" & wbName & "] 的 " & count & " 个任务。", vbInformation, "启动完成"
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "启动任务时出错: " & Err.Description, vbCritical, "错误"
End Sub

' 暂停任务
Private Sub LuaTaskMenu_PauseTask()
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If
    If Not g_TaskFunc.Exists(taskId) Then
        MsgBox "任务已不存在。", vbExclamation
        Exit Sub
    End If

    If g_TaskQueue.Exists(taskId) Then
        g_TaskQueue.Remove taskId
        g_TaskStatus(taskId) = "paused"
        MsgBox "任务 " & taskId & " 已暂停。" & vbCrLf & _
               "使用 ResumeTask 恢复。", vbInformation, "任务已暂停"
    Else
        MsgBox "任务 " & taskId & " 不在活跃队列中。", vbExclamation, "提示"
    End If
End Sub

' 恢复任务
Private Sub LuaTaskMenu_ResumeTask()
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If
    If Not g_TaskFunc.Exists(taskId) Then
        MsgBox "任务 " & taskId & " 不存在。", vbCritical, "错误"
        Exit Sub
    End If

    Dim status As String
    status = g_TaskStatus(taskId)
    If status <> "yielded" And status <> "paused" Then
        MsgBox "任务 " & taskId & " 状态为 " & status & "，无法恢复。", vbExclamation, "无法恢复"
        Exit Sub
    End If

    If Not g_TaskQueue.Exists(taskId) Then
        g_TaskQueue(taskId) = True
        StartSchedulerIfNeeded
        MsgBox "任务 " & taskId & " 已恢复。", vbInformation, "任务已恢复"
    Else
        MsgBox "任务 " & taskId & " 已在活跃队列中。", vbInformation, "提示"
    End If
End Sub

' 终止任务
Private Sub LuaTaskMenu_terminateTask()
    On Error Resume Next
    If g_TaskFunc Is Nothing Then Exit Sub

    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Or Not g_TaskFunc.Exists(taskId) Then
        MsgBox "任务不存在或已删除", vbExclamation
        Exit Sub
    End If

    ' 从队列移除
    If Not g_TaskQueue Is Nothing Then
        If g_TaskQueue.Exists(taskId) Then g_TaskQueue.Remove taskId
    End If

    ' 设置终止状态并标记为脏
    g_TaskStatus(taskId) = "terminated"
    g_StateDirty = True

    ' 删除所有数据
    g_TaskFunc.Remove taskId
    g_TaskWorkbook.Remove taskId
    g_TaskStartArgs.Remove taskId
    g_TaskResumeSpec.Remove taskId
    g_TaskCell.Remove taskId
    g_TaskStatus.Remove taskId
    g_TaskProgress.Remove taskId
    g_TaskMessage.Remove taskId
    g_TaskValue.Remove taskId
    g_TaskError.Remove taskId
    g_TaskCoThread.Remove taskId

    MsgBox "任务已终止并删除: " & taskId, vbInformation
End Sub

' 查看任务详情
Private Sub LuaTaskMenu_ShowDetail()
    On Error GoTo ErrorHandler

    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = "" Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If

    If g_TaskFunc Is Nothing Then
        InitCoroutineSystem
    End If
    
    If Not g_TaskFunc.Exists(taskId) Then
        MsgBox "任务 " & taskId & " 不存在！", vbCritical, "错误"
        Exit Sub
    End If
    
    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  任务详细信息" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf
    
    msg = msg & "任务ID: " & taskId & vbCrLf
    msg = msg & "函数名: " & g_TaskFunc(taskId) & vbCrLf
    msg = msg & "单元格: " & g_TaskCell(taskId) & vbCrLf
    msg = msg & "状态: " & g_TaskStatus(taskId) & vbCrLf
    msg = msg & "进度: " & Format(g_TaskProgress(taskId), "0.00") & "%" & vbCrLf
    msg = msg & "消息: " & CStr(g_TaskMessage(taskId)) & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf & vbCrLf
    
    ' 启动参数
    msg = msg & "启动参数:" & vbCrLf
    Dim startArgs As Variant
    startArgs = g_TaskStartArgs(taskId)
    If IsArray(startArgs) Then
        Dim i As Long
        For i = LBound(startArgs) To UBound(startArgs)
            msg = msg & "   [" & i & "] " & CStr(startArgs(i)) & vbCrLf
        Next i
    Else
        msg = msg & "   (无)" & vbCrLf
    End If

    ' Resume 参数
    msg = msg & vbCrLf & "Resume 参数:" & vbCrLf
    Dim resumeSpec As Variant
    resumeSpec = g_TaskResumeSpec(taskId)
    If IsArray(resumeSpec) Then
        For i = LBound(resumeSpec) To UBound(resumeSpec)
            msg = msg & "   [" & i & "] " & CStr(resumeSpec(i)) & vbCrLf
        Next i
    Else
        msg = msg & "   (无)" & vbCrLf
    End If

    ' 当前值
    msg = msg & vbCrLf & "当前值:" & vbCrLf
    Dim value As Variant
    value = g_TaskValue(taskId)
    If IsArray(value) Then
        msg = msg & "   (数组，维度: " & ArrayDimensions(value) & ")" & vbCrLf
    ElseIf IsEmpty(value) Then
        msg = msg & "   (空)" & vbCrLf
    Else
        Dim valueStr As String
        valueStr = CStr(value)
        If Len(valueStr) > 100 Then valueStr = Left(valueStr, 97) & "..."
        msg = msg & "   " & valueStr & vbCrLf
    End If
    
    ' 错误信息
    If g_TaskStatus(taskId) = "error" Then
        msg = msg & vbCrLf & " 错误信息:" & vbCrLf
        msg = msg & "   " & g_TaskError(taskId) & vbCrLf
    End If
    
    ' 调度信息
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf
    msg = msg & "在活跃队列中: " & IIf(g_TaskQueue.Exists(taskId), "是", "否") & vbCrLf
    msg = msg & "协程线程: " & IIf(g_TaskCoThread(taskId) = 0, "未创建", "0x" & Hex(g_TaskCoThread(taskId))) & vbCrLf
    
    MsgBox msg, vbInformation, "任务详情 - " & taskId
    
    Exit Sub

ErrorHandler:
    MsgBox "显示任务详情时出错: " & Err.Description, vbCritical, "错误"
End Sub
' ====调度管理功能====
' 手动启动调度器
Private Sub LuaSchedulerMenu_StartScheduler()
    If g_TaskQueue Is Nothing Then
        InitCoroutineSystem
    End If
    
    If g_TaskQueue.Count = 0 Then
        MsgBox "队列中没有任务，无需启动调度器。" & vbCrLf & vbCrLf & _
               "请先启动一些任务，或使用【启动所有 defined 任务】。", _
               vbExclamation, "无任务"
        Exit Sub
    End If
    
    If g_SchedulerRunning Then
        MsgBox "调度器已在运行中。" & vbCrLf & vbCrLf & _
               "当前活跃任务数: " & g_TaskQueue.Count, _
               vbInformation, "调度器状态"
        Exit Sub
    End If
    
    g_SchedulerRunning = True
    g_NextScheduleTime = Now + g_SchedulerIntervalMilliSec / 86400000#
    Application.OnTime g_NextScheduleTime, "SchedulerTick"
    
    MsgBox "调度器已启动。" & vbCrLf & vbCrLf & _
           "调度间隔: " & g_SchedulerIntervalMilliSec & " ms" & vbCrLf & _
           "调度模式: " & IIf(g_ScheduleMode = 0, "按任务顺序", "按工作簿") & vbCrLf & _
           "当前队列任务数: " & g_TaskQueue.Count, _
           vbInformation, "调度器已启动"
End Sub

' 手动停止调度器
Private Sub LuaSchedulerMenu_StopScheduler()
    If Not g_SchedulerRunning Then
        MsgBox "调度器已经是停止状态。", vbInformation, "调度器状态"
        Exit Sub
    End If
    
    Dim result As VbMsgBoxResult
    result = MsgBox("确定要停止调度器吗？" & vbCrLf & vbCrLf & _
                    "活跃任务将不会继续执行。" & vbCrLf & _
                    "当前队列任务数: " & g_TaskQueue.Count & vbCrLf & vbCrLf & _
                    "提示：可使用【启动调度器】重新启动。", _
                    vbQuestion + vbYesNo, "确认停止")
    
    If result = vbNo Then Exit Sub
    
    StopScheduler
End Sub

' 批量启动所有 defined 状态的任务
Private Sub LuaSchedulerMenu_StartAllDefinedTasks()
    Dim taskId As Variant
    Dim count As Long
    
    For Each taskId In g_TaskFunc.Keys
        If g_TaskStatus(CStr(taskId)) = "defined" Then
            StartLuaCoroutine CStr(taskId)
            count = count + 1
        End If
    Next
    
    MsgBox "已启动 " & count & " 个任务", vbInformation
End Sub

' 清理所有已完成或错误的任务
Private Sub LuaSchedulerMenu_CleanupFinishedTasks()
    On Error Resume Next
    
    If g_TaskFunc Is Nothing Then
        MsgBox "没有任务需要清理。", vbInformation, "清理任务"
        Exit Sub
    End If
    
    Dim taskId As Variant
    Dim tasksToRemove As Object
    Set tasksToRemove = CreateObject("System.Collections.ArrayList")
    
    ' 收集需要清理的任务
    For Each taskId In g_TaskFunc.Keys
        Dim status As String
        status = g_TaskStatus(CStr(taskId))
        If status = "done" Or status = "error" Then
            tasksToRemove.Add taskId
        End If
    Next taskId
    
    ' 清理任务
    Dim i As Long
    For i = 0 To tasksToRemove.Count - 1
        Dim tid As String
        tid = CStr(tasksToRemove(i))
        
        g_TaskFunc.Remove tid
        g_TaskStartArgs.Remove tid
        g_TaskResumeSpec.Remove tid
        g_TaskCell.Remove tid
        g_TaskStatus.Remove tid
        g_TaskProgress.Remove tid
        g_TaskMessage.Remove tid
        g_TaskValue.Remove tid
        g_TaskError.Remove tid
        g_TaskCoThread.Remove tid
        
        If g_TaskQueue.Exists(tid) Then
            g_TaskQueue.Remove tid
        End If
    Next i
    
    MsgBox "已清理 " & tasksToRemove.Count & " 个已完成或错误的任务。" & vbCrLf & _
           "剩余任务: " & g_TaskFunc.Count, vbInformation, "清理完成"
End Sub

' 清理特定工作簿的任务
Private Sub LuaSchedulerMenu_CleanupWorkbookTasks()
    Dim taskCell As String
    Dim wbName As String
    taskCell = Application.Caller.Address(External:=True)
    wbName = Application.Caller.Worksheet.Parent.Name
    CleanupWorkbookTasks wbName
    MsgBox "已清理工作簿 " & wbName & " 的任务。", vbInformation
End Sub

' 清空所有任务和队列
Private Sub LuaSchedulerMenu_ClearAllTasks()
    Dim result As VbMsgBoxResult
    result = MsgBox("确定要清空所有任务吗？" & vbCrLf & vbCrLf & _
                    "这将删除所有任务数据，无法恢复！", _
                    vbExclamation + vbYesNo, "确认清空")
    
    If result = vbNo Then Exit Sub
    
    ' 停止调度器
    g_SchedulerRunning = False
    
    ' 清空所有 Dictionary
    If Not g_TaskFunc Is Nothing Then
        g_TaskFunc.RemoveAll
        g_TaskStartArgs.RemoveAll
        g_TaskResumeSpec.RemoveAll
        g_TaskCell.RemoveAll
        g_TaskStatus.RemoveAll
        g_TaskProgress.RemoveAll
        g_TaskMessage.RemoveAll
        g_TaskValue.RemoveAll
        g_TaskError.RemoveAll
        g_TaskCoThread.RemoveAll
        g_TaskQueue.RemoveAll
    End If
    
    MsgBox "所有任务已清空。", vbInformation, "清空完成"
End Sub

' 显示所有任务（按工作簿分组）
Private Sub LuaSchedulerMenu_ShowAllTasks()
    On Error GoTo ErrorHandler
    
    If g_TaskFunc Is Nothing Then
        InitCoroutineSystem
    End If
    
    If g_TaskFunc.Count = 0 Then
        MsgBox "当前没有任何任务。", vbInformation, "任务列表"
        Exit Sub
    End If
    
    ' 按工作簿分组统计
    Dim wbStats As Object
    Set wbStats = CreateObject("Scripting.Dictionary")
    
    Dim taskId As Variant
    Dim taskCount As Long
    Dim runningCount As Long, yieldedCount As Long, doneCount As Long, errorCount As Long
    For Each taskId In g_TaskWorkbook.Keys
        Dim wbName As String
        wbName = g_TaskWorkbook(CStr(taskId))
        
        If Not wbStats.Exists(wbName) Then
            wbStats(wbName) = 0
        End If
        wbStats(wbName) = wbStats(wbName) + 1
    Next
    
    ' 构建消息
    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  Lua 协程任务管理器（单实例）" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf
    
    msg = msg & "任务总数: " & g_TaskFunc.Count & vbCrLf
    msg = msg & "活跃队列: " & g_TaskQueue.Count & vbCrLf
    msg = msg & "调度器: " & IIf(g_SchedulerRunning, "运行中", "已停止") & vbCrLf
    msg = msg & vbCrLf & "按工作簿分组:" & vbCrLf
    
    Dim wb As Variant
    For Each wb In wbStats.Keys
        msg = msg & "  [" & CStr(wb) & "]: " & wbStats(CStr(wb)) & " 个任务" & vbCrLf
    Next
    
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf
    
    ' 统计各状态任务数
    For Each taskId In g_TaskFunc.Keys
        Select Case g_TaskStatus(CStr(taskId))
            Case "running": runningCount = runningCount + 1
            Case "yielded": yieldedCount = yieldedCount + 1
            Case "done": doneCount = doneCount + 1
            Case "error": errorCount = errorCount + 1
        End Select
    Next taskId
    
    msg = msg & "状态统计:" & vbCrLf
    msg = msg & "   运行中: " & runningCount & vbCrLf
    msg = msg & "   暂停中: " & yieldedCount & vbCrLf
    msg = msg & "   已完成: " & doneCount & vbCrLf
    msg = msg & "   错误: " & errorCount & vbCrLf
    msg = msg & vbCrLf & "========================================" & vbCrLf & vbCrLf
    
    ' 详细列出每个任务
    taskCount = 0
    For Each taskId In g_TaskFunc.Keys
        taskCount = taskCount + 1
        msg = msg & "【任务 #" & taskCount & "】" & vbCrLf
        msg = msg & "  ID: " & CStr(taskId) & vbCrLf
        msg = msg & "  函数: " & g_TaskFunc(CStr(taskId)) & vbCrLf
        msg = msg & "  单元格: " & g_TaskCell(CStr(taskId)) & vbCrLf
        msg = msg & "  状态: " & g_TaskStatus(CStr(taskId)) & vbCrLf
        msg = msg & "  进度: " & Format(g_TaskProgress(CStr(taskId)), "0.0") & "%" & vbCrLf
        
        ' 显示消息
        Dim msgText As String
        msgText = CStr(g_TaskMessage(CStr(taskId)))
        If Len(msgText) > 50 Then msgText = Left(msgText, 47) & "..."
        msg = msg & "  消息: " & msgText & vbCrLf
        
        ' 如果有错误，显示错误信息
        If g_TaskStatus(CStr(taskId)) = "error" Then
            Dim errText As String
            errText = CStr(g_TaskError(CStr(taskId)))
            If Len(errText) > 60 Then errText = Left(errText, 57) & "..."
            msg = msg & "   错误: " & errText & vbCrLf
        End If
        
        ' 显示是否在活跃队列中
        If g_TaskQueue.Exists(CStr(taskId)) Then
            msg = msg & "  队列: 是" & vbCrLf
        End If
        
        msg = msg & "----------------------------------------" & vbCrLf
    Next taskId
    
    ' 显示消息框
    MsgBox msg, vbInformation, "Lua 协程任务列表 (" & g_TaskFunc.Count & " 个任务)"
    
    Exit Sub
ErrorHandler:
    MsgBox "显示任务信息时出错: " & Err.Description, vbCritical, "错误"
End Sub

' ====插件设置功能====
' 启用热重载
Private Sub LuaConfigMenu_EnableHotReload()
    g_HotReloadEnabled = True
    MsgBox "Lua 自动热重载已启用。" & vbCrLf & _
           "当 functions.lua 修改后，系统将自动重新加载。", _
           vbInformation, "热重载已启用"
End Sub

' 禁用热重载
Private Sub LuaConfigMenu_DisableHotReload()
    g_HotReloadEnabled = False
    MsgBox "Lua 自动热重载已禁用。" & vbCrLf & _
           "如需更新 functions.lua，请手动运行 ""ReloadFunctions""。", _
           vbExclamation, "热重载已禁用"
End Sub

' 手动重载 functions.lua
Private Sub LuaConfigMenu_ReloadFunctions()
    If Not g_Initialized Then
        If Not InitLuaState() Then
            MsgBox "无法初始化 Lua 引擎。", vbCritical, "重载失败"
            Exit Sub
        End If
    End If
    
    If TryLoadFunctionsFile() Then
        MsgBox "functions.lua 已成功重载！", vbInformation, "重载成功"
    Else
        MsgBox "functions.lua 重载失败。" & vbCrLf & _
               "请检查文件语法。", vbCritical, "重载失败"
    End If
End Sub

' 设置调度间隔（毫秒）
Private Sub LuaConfigMenu_SetSchedulerInterval()
    If g_SchedulerRunning Then
        StopScheduler
    End If

    Dim seconds As Long
    seconds = Application.InputBox( _
            "请输入调度的间隔时间（>=10ms且<=60000ms）", _
            "调度参数", _
            g_SchedulerIntervalMilliSec, _
            Type:=1 _
        )

    If seconds = False Then Exit Sub
    If seconds < 10 Or seconds > 60 Then
        MsgBox "间隔不能小于 10 ms。且不能大于 60 秒。", vbExclamation, "无效值"
        Exit Sub
    End If

    g_SchedulerIntervalMilliSec = seconds
    ResumeScheduler
End Sub

' 设置调度步数 
Private Sub LuaConfigMenu_SetSchedulerBatchSize()
    If g_SchedulerRunning Then
        StopScheduler
    End If

    Dim v As Variant
    v = Application.InputBox( _
            "请输入每次调度的最大任务数（>=1）", _
            "调度参数", _
            g_MaxIterationsPerTick, _
            Type:=1 _
        )

    If v = False Then Exit Sub
    If v < 1 Or v <> CLng(v) Then
        MsgBox "请输入 >=1 的整数", vbExclamation
        Exit Sub
    End If

    g_MaxIterationsPerTick = CLng(v)
    ResumeScheduler
End Sub

' 切换调度模式
Private Sub LuaConfigMenu_ToggleScheduleMode()
    If g_ScheduleMode = 0 Then
        g_ScheduleMode = 1
        MsgBox "已切换到【按工作簿调度】模式" & vbCrLf & _
               "每个工作簿可独立设置执行的tick数", vbInformation
    Else
        g_ScheduleMode = 0
        MsgBox "已切换到【按任务顺序调度】模式" & vbCrLf & _
               "使用Round-Robin轮询所有任务", vbInformation
    End If
End Sub

' 设置工作簿Tick数
Private Sub LuaConfigMenu_SetWorkbookTicks()
    If g_ScheduleMode = 0 Then
        MsgBox "当前为【按任务顺序】模式，此设置无效" & vbCrLf & _
               "请先切换到【按工作簿调度】模式", vbExclamation
        Exit Sub
    End If

    Dim choice As String
    choice = InputBox("请选择设置方式：" & vbCrLf & _
                     "1 - 设置默认tick数（所有工作簿）" & vbCrLf & _
                     "2 - 设置当前工作簿tick数", "设置工作簿Tick", "1")

    If choice = "1" Then
        Dim defaultTicks As Variant
        defaultTicks = Application.InputBox("请输入默认tick数（>=1）", "默认设置", g_WorkbookTicks, Type:=1)
        If defaultTicks = False Then Exit Sub
        If defaultTicks < 1 Then
            MsgBox "tick数必须>=1", vbExclamation
            Exit Sub
        End If
        g_WorkbookTicks = CLng(defaultTicks)
        MsgBox "默认工作簿tick数已设置为: " & g_WorkbookTicks, vbInformation

    ElseIf choice = "2" Then
        Dim wbName As String
        wbName = ActiveWorkbook.Name

        Dim wbTicks As Variant
        Dim currentTicks As Long
        If g_WorkbookTickCount.Exists(wbName) Then
            currentTicks = g_WorkbookTickCount(wbName)
        Else
            currentTicks = g_WorkbookTicks
        End If

        wbTicks = Application.InputBox("设置工作簿 [" & wbName & "] 的tick数（>=1）", "工作簿设置", currentTicks, Type:=1)
        If wbTicks = False Then Exit Sub
        If wbTicks < 1 Then
            MsgBox "tick数必须>=1", vbExclamation
            Exit Sub
        End If
        g_WorkbookTickCount(wbName) = CLng(wbTicks)
        MsgBox "工作簿 [" & wbName & "] tick数已设置为: " & wbTicks, vbInformation
    End If
End Sub
' ====调试和诊断功能====
' 显示加载宏状态
Private Sub LuaDebugMenu_ShowAddinStatus()
    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  Excel-Lua 5.4 加载宏状态" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf
    
    msg = msg & "加载宏名称: " & ThisWorkbook.Name & vbCrLf
    msg = msg & "加载宏路径: " & ThisWorkbook.Path & vbCrLf
    msg = msg & "Lua初始化: " & IIf(g_Initialized, "已初始化", "未初始化") & vbCrLf
    msg = msg & "热重载: " & IIf(g_HotReloadEnabled, "已启用", "已禁用") & vbCrLf
    msg = msg & "调度器: " & IIf(g_SchedulerRunning, "运行中", "已停止") & vbCrLf
    msg = msg & "调度间隔: " & g_SchedulerIntervalMilliSec & " 毫秒" & vbCrLf
    msg = msg & "调度步数: " & g_MaxIterationsPerTick & vbCrLf
    msg = msg & "调度模式: " & IIf(g_ScheduleMode = 0, "按任务顺序", "按工作簿") & vbCrLf
    If g_ScheduleMode = 1 Then
        msg = msg & "默认Tick数: " & g_WorkbookTicks & vbCrLf
    End If
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf
    
    If g_TaskFunc Is Nothing Then
        msg = msg & "任务总数: 0" & vbCrLf
    Else
        msg = msg & "任务总数: " & g_TaskFunc.Count & vbCrLf
        msg = msg & "活跃任务: " & g_TaskQueue.Count & vbCrLf
    End If
    
    msg = msg & vbCrLf & "functions.lua: " & vbCrLf
    msg = msg & "  路径: " & g_FunctionsPath & vbCrLf
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(g_FunctionsPath) Then
        msg = msg & "  状态: 存在" & vbCrLf
        msg = msg & "  修改时间: " & FileDateTime(g_FunctionsPath) & vbCrLf
        msg = msg & "  最后加载: " & g_LastModified & vbCrLf
    Else
        msg = msg & "  状态: 不存在" & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "加载宏状态"
End Sub
' ====性能统计功能====
' 显示调度器统计
Private Sub LuaPerfMenu_ShowSchedulerStats()
    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  调度器性能统计" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf

    msg = msg & "启动时间: " & Format(g_SchedulerStats.StartTime, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    msg = msg & "运行时长: " & Format(Now - g_SchedulerStats.StartTime, "hh:nn:ss") & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf

    msg = msg & "总调度次数: " & g_SchedulerStats.TotalCount & vbCrLf
    msg = msg & "总运行时间: " & Format(g_SchedulerStats.TotalTime, "0.00") & " ms" & vbCrLf

    If g_SchedulerStats.TotalCount > 0 Then
        msg = msg & "平均每次: " & Format(g_SchedulerStats.TotalTime / g_SchedulerStats.TotalCount, "0.00") & " ms" & vbCrLf
    End If

    msg = msg & vbCrLf & "上次调度: " & Format(g_SchedulerStats.LastTime, "0.00") & " ms" & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf
    
    msg = msg & "调度模式: " & IIf(g_ScheduleMode = 0, "按任务顺序", "按工作簿") & vbCrLf
    msg = msg & "调度间隔: " & g_SchedulerIntervalMilliSec & " ms" & vbCrLf
    
    If g_ScheduleMode = 0 Then
        msg = msg & "每次执行: " & g_MaxIterationsPerTick & " 个任务" & vbCrLf
    Else
        msg = msg & "默认Tick: " & g_WorkbookTicks & " 次/工作簿" & vbCrLf
    End If
    
    msg = msg & vbCrLf & "当前状态: " & IIf(g_SchedulerRunning, "运行中", "已停止") & vbCrLf
    msg = msg & "活跃任务: " & g_TaskQueue.Count & vbCrLf
    
    MsgBox msg, vbInformation, "调度器性能统计"
End Sub

' 显示任务性能统计
Private Sub LuaPerfMenu_ShowTaskStats()
    If g_TaskFunc Is Nothing Or g_TaskFunc.Count = 0 Then
        MsgBox "当前没有任何任务。", vbInformation, "任务性能统计"
        Exit Sub
    End If
    
    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  任务性能统计" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf
    
    msg = msg & "任务总数: " & g_TaskFunc.Count & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf & vbCrLf
    
    Dim taskId As Variant
    Dim taskNum As Long
    taskNum = 0
    
    For Each taskId In g_TaskFunc.Keys
        taskNum = taskNum + 1
        msg = msg & "【任务 #" & taskNum & "】" & vbCrLf
        msg = msg & "  ID: " & CStr(taskId) & vbCrLf
        msg = msg & "  函数: " & g_TaskFunc(CStr(taskId)) & vbCrLf
        msg = msg & "  状态: " & g_TaskStatus(CStr(taskId)) & vbCrLf
        
        If g_TaskRunCount.Exists(CStr(taskId)) Then
            msg = msg & "  调度次数: " & g_TaskStats(CStr(taskId)).RunCount & vbCrLf
            msg = msg & "  总运行时间: " & Format(g_TaskStats(CStr(taskId)).TotalTime, "0.00") & " ms" & vbCrLf
            msg = msg & "  平均时间: " & Format(g_TaskStats(CStr(taskId)).TotalTime / g_TaskStats(CStr(taskId)).RunCount, "0.00") & " ms" & vbCrLf
            msg = msg & "  上次运行: " & Format(g_TaskStats(CStr(taskId)).LastTime, "0.00") & " ms" & vbCrLf
        Else
            msg = msg & "  (尚未执行)" & vbCrLf
        End If
        
        msg = msg & "----------------------------------------" & vbCrLf
    Next taskId
    
    MsgBox msg, vbInformation, "任务性能统计 (" & g_TaskFunc.Count & " 个任务)"
End Sub

' 显示工作簿性能统计
Private Sub LuaPerfMenu_ShowWorkbookStats()
    If g_TaskWorkbook Is Nothing Or g_TaskWorkbook.Count = 0 Then
        MsgBox "当前没有任何任务。", vbInformation, "工作簿性能统计"
        Exit Sub
    End If
    
    ' 统计每个工作簿的任务数
    Dim wbTaskCount As Object
    Set wbTaskCount = CreateObject("Scripting.Dictionary")
    
    Dim taskId As Variant
    For Each taskId In g_TaskWorkbook.Keys
        Dim wbName As String
        wbName = g_TaskWorkbook(CStr(taskId))
        
        If Not wbTaskCount.Exists(wbName) Then
            wbTaskCount(wbName) = 0
        End If
        wbTaskCount(wbName) = wbTaskCount(wbName) + 1
    Next
    
    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  工作簿性能统计" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf
    
    msg = msg & "工作簿总数: " & wbTaskCount.Count & vbCrLf
    msg = msg & "调度模式: " & IIf(g_ScheduleMode = 0, "按任务顺序", "按工作簿") & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf & vbCrLf
    
    Dim wb As Variant
    Dim wbNum As Long
    wbNum = 0
    
    For Each wb In wbTaskCount.Keys
        wbNum = wbNum + 1
        msg = msg & "【工作簿 #" & wbNum & "】" & vbCrLf
        msg = msg & "  名称: " & CStr(wb) & vbCrLf
        msg = msg & "  任务数: " & wbTaskCount(CStr(wb)) & vbCrLf
        
        If g_WorkbookStats.Exists(CStr(wb)) Then
            msg = msg & "  总调度次数: " & g_WorkbookStats(CStr(wb)).TickCount & vbCrLf
            msg = msg & "  总运行时间: " & Format(g_WorkbookStats(CStr(wb)).TotalTime, "0.00") & " ms" & vbCrLf
            msg = msg & "  平均时间: " & Format(g_WorkbookStats(CStr(wb)).TotalTime / g_WorkbookStats(CStr(wb)).TickCount, "0.00") & " ms" & vbCrLf
            msg = msg & "  上次调度: " & Format(g_WorkbookStats(CStr(wb)).LastTime, "0.00") & " ms" & vbCrLf
        Else
            msg = msg & "  (尚未执行)" & vbCrLf
        End If
        
        ' 显示配置的tick数（仅在按工作簿调度模式下）
        If g_ScheduleMode = 1 Then
            If g_WorkbookTickCount.Exists(CStr(wb)) Then
                msg = msg & "  配置Tick数: " & g_WorkbookTickCount(CStr(wb)) & vbCrLf
            Else
                msg = msg & "  配置Tick数: " & g_WorkbookTicks & " (默认)" & vbCrLf
            End If
        End If
        
        msg = msg & "----------------------------------------" & vbCrLf
    Next wb
    
    MsgBox msg, vbInformation, "工作簿性能统计 (" & wbTaskCount.Count & " 个工作簿)"
End Sub

' 重置性能统计
Private Sub LuaPerfMenu_ResetStats()
    Dim result As VbMsgBoxResult
    result = MsgBox("确定要重置所有性能统计数据吗？" & vbCrLf & vbCrLf & _
                    "这将清除所有计时数据，但不影响任务运行。", _
                    vbQuestion + vbYesNo, "确认重置")

    If result = vbNo Then Exit Sub

    ' 重置调度器统计
    g_SchedulerStats.TotalTime = 0
    g_SchedulerStats.LastTime = 0
    g_SchedulerStats.TotalCount = 0
    g_SchedulerStats.StartTime = Now

    ' 重置任务统计
    If Not g_TaskStats Is Nothing Then g_TaskStats.RemoveAll

    ' 重置工作簿统计
    If Not g_WorkbookStats Is Nothing Then g_WorkbookStats.RemoveAll

    MsgBox "所有性能统计数据已重置。", vbInformation, "重置完成"
End Sub
' ============================================
' 手动初始化/清理函数（供外部调用）
' ============================================
' 手动初始化Lua引擎
Public Sub ManualInitLua()
    If InitLuaState() Then
        MsgBox "Lua引擎初始化成功！", vbInformation, "初始化完成"
    Else
        MsgBox "Lua引擎初始化失败！", vbCritical, "初始化失败"
    End If
End Sub

' 手动清理Lua引擎（慎用！）
Public Sub ManualCleanupLua()
    Dim result As VbMsgBoxResult
    result = MsgBox("警告：这将清理所有Lua资源和任务！" & vbCrLf & vbCrLf & _
                    "所有工作簿的Lua任务都会停止。" & vbCrLf & _
                    "确定要继续吗？", _
                    vbExclamation + vbYesNo, "确认清理")
    
    If result = vbYes Then
        CleanupLua
        MsgBox "Lua引擎已清理。", vbInformation, "清理完成"
    End If
End Sub