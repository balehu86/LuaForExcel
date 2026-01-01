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

        If dims <> vbNullString Then dims = dims & " x "
        dims = dims & (ub - lb + 1)
    Next i

    If dims = vbNullString Then dims = "未知"
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

' 辅助函数：清理特定工作簿的任务
Public Sub CleanupWorkbookTasks(wbName As String)
    On Error Resume Next

    If g_Tasks Is Nothing Then Exit Sub

    ' 收集要删除的任务ID（不能在遍历时删除）
    Dim toRemove As Collection
    Set toRemove = New Collection

    Dim taskId As Variant
    For Each taskId In g_Tasks.Keys
        If g_Tasks(CStr(taskId)).taskWorkbook = wbName Then
            toRemove.Add CStr(taskId)
        End If
    Next

    ' 删除任务
    Dim removeId As Variant
    For Each removeId In toRemove
        CollectionRemove g_TaskQueue, CStr(removeId)
        g_Tasks.Remove CStr(removeId)
    Next

    ' 清理工作簿记录
    If g_Workbooks.Exists(wbName) Then
        g_Workbooks.Remove wbName
    End If
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
    AddLuaMenuItem luaTaskMenu, "暂停任务", "LuaTaskMenu_PauseTask"
    AddLuaMenuItem luaTaskMenu, "恢复任务", "LuaTaskMenu_ResumeTask"
    AddLuaMenuItem luaTaskMenu, "终止任务", "LuaTaskMenu_TerminateTask"
    AddLuaMenuItem luaTaskMenu, "查看任务详情", "LuaTaskMenu_ShowDetail"
    AddLuaMenuItem luaTaskMenu, "设置任务权重", "LuaConfigMenu_SetTaskWeight"

    ' 添加调度的主菜单
    Dim luaSchedulerMenu As CommandBarControl
    Set luaSchedulerMenu = cMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    luaSchedulerMenu.Caption = "Lua 调度管理"
    luaSchedulerMenu.Tag = "LuaSchedulerMenu"
    ' 添加调度的子菜单
    AddLuaMenuItem luaSchedulerMenu, "启动调度器", "LuaSchedulerMenu_StartScheduler"
    AddLuaMenuItem luaSchedulerMenu, "停止调度器", "LuaSchedulerMenu_StopScheduler"
    AddLuaMenuItem luaSchedulerMenu, "启动本簿所有任务", "LuaSchedulerMenu_StartAllWorkbookTasks"
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
        If ctrl.Tag = "LuaDebugMenu" Then ctrl.Delete
        If ctrl.Tag = "LuaPerfMenu" Then ctrl.Delete
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
    If taskId = vbNullString Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If
    If g_Tasks(taskId).taskStatus = "defined" Then
        StartLuaCoroutine taskId
        MsgBox "任务已启动: " & taskId, vbInformation
    Else
        MsgBox "任务状态为 " & g_Tasks(taskId).taskStatus & "，无法启动。", vbExclamation
    End If
End Sub

' 暂停任务
Private Sub LuaTaskMenu_PauseTask()
    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = vbNullString Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If

    If Not g_Tasks.Exists(taskId) Then
        MsgBox "任务已不存在。", vbExclamation
        Exit Sub
    End If

    If CollectionExists(g_TaskQueue, taskId) Then
        CollectionRemove g_TaskQueue, taskId
        g_Tasks(taskId).taskStatus = "paused"
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
    If taskId = vbNullString Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If
    If Not g_Tasks.Exists(taskId) Then
        MsgBox "任务 " & taskId & " 不存在。", vbCritical, "错误"
        Exit Sub
    End If

    Dim status As String
    status = g_Tasks(taskId).taskStatus
    If status <> "yielded" And status <> "paused" Then
        MsgBox "任务 " & taskId & " 状态为 " & status & "，无法恢复。", vbExclamation, "无法恢复"
        Exit Sub
    End If

    If Not CollectionExists(g_TaskQueue, taskId) Then
        CollectionAdd g_TaskQueue, taskId
        StartSchedulerIfNeeded
        MsgBox "任务 " & taskId & " 已恢复。", vbInformation, "任务已恢复"
    Else
        MsgBox "任务 " & taskId & " 已在活跃队列中。", vbInformation, "提示"
    End If
End Sub

' 终止任务
Private Sub LuaTaskMenu_terminateTask()
    On Error Resume Next
    If g_Tasks Is Nothing Then Exit Sub

    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = vbNullString Then
        MsgBox "任务不存在或已删除", vbExclamation
        Exit Sub
    End If

    ' 从队列移除
    If Not g_TaskQueue Is Nothing Then
        CollectionRemove g_TaskQueue, taskId
    End If

    ' 设置终止状态并标记为脏
    g_Tasks(taskId).taskStatus = "terminated"
    g_StateDirty = True

    ' ' 删除所有数据
    ' g_TaskFunc.Remove taskId
    ' g_TaskWorkbook.Remove taskId
    ' g_TaskStartArgs.Remove taskId
    ' g_TaskResumeSpec.Remove taskId
    ' g_TaskCell.Remove taskId
    ' g_TaskStatus.Remove taskId
    ' g_TaskProgress.Remove taskId
    ' g_TaskMessage.Remove taskId
    ' g_TaskValue.Remove taskId
    ' g_TaskError.Remove taskId
    ' g_TaskCoThread.Remove taskId

    MsgBox "任务已终止并删除: " & taskId, vbInformation
End Sub

' 查看任务详情
Private Sub LuaTaskMenu_ShowDetail()
    On Error GoTo ErrorHandler

    Dim taskId As String
    taskId = GetTaskIdFromSelection()
    If taskId = vbNullString Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If

    If g_Tasks Is Nothing Then
        InitCoroutineSystem
    End If
    If Not g_Tasks.Exists(taskId) Then
        MsgBox "任务 " & taskId & " 不存在！", vbCritical, "错误"
        Exit Sub
    End If
    Dim task As TaskUnit
    Set task = g_Tasks(taskId)
    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  任务详细信息" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf

    msg = msg & "任务ID: " & taskId & vbCrLf
    msg = msg & "函数名: " & task.taskFunc & vbCrLf
    msg = msg & "单元格: " & task.taskCell & vbCrLf
    msg = msg & "状态: " & task.taskStatus & vbCrLf
    msg = msg & "进度: " & Format(task.taskProgress, "0.00") & "%" & vbCrLf
    msg = msg & "消息: " & task.taskMessage & vbCrLf
    msg = msg & "  CFS vruntime: " & Format(task.CFS_vruntime, "0.00") & " ms" & vbCrLf
    msg = msg & "  CFS 权重: " & task.CFS_weight & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf & vbCrLf

    ' 启动参数
    msg = msg & "启动参数:" & vbCrLf
    Dim startArgs As Variant
    startArgs = task.taskStartArgs
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
    resumeSpec = task.TaskResumeSpec
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
    value = task.taskValue
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
    If task.taskStatus = "error" Then
        msg = msg & vbCrLf & " 错误信息:" & vbCrLf
        msg = msg & "   " & task.taskError & vbCrLf
    End If

    ' 调度信息
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf
    msg = msg & "在活跃队列中: " & IIf(CollectionExists(g_TaskQueue, taskId), "是", "否") & vbCrLf
    msg = msg & "协程线程: " & IIf(task.taskCoThread = 0, "未创建", "0x" & Hex(task.taskCoThread)) & vbCrLf

    MsgBox msg, vbInformation, "任务详情 - " & taskId

    Exit Sub
ErrorHandler:
    MsgBox "显示任务详情时出错: " & Err.Description, vbCritical, "错误"
End Sub

' 设置任务权重
Private Sub LuaConfigMenu_SetTaskWeight()
    Dim taskId As String
    taskId = GetTaskIdFromSelection()

    If taskId = vbNullString Then
        MsgBox "当前单元格没有 Lua 任务。", vbExclamation
        Exit Sub
    End If

    Dim task As TaskUnit
    Set task = g_Tasks(taskId)

    Dim newWeight As Variant
    newWeight = Application.InputBox( _
        "设置任务权重（默认1024，越大优先级越高）" & vbCrLf & _
        "建议范围: 256 ~ 4096", _
        "CFS 权重设置", _
        task.CFS_weight, _
        Type:=1)

    If newWeight = False Then Exit Sub
    If newWeight < 1 Then newWeight = 1
    If newWeight > 65536 Then newWeight = 65536

    task.CFS_weight = CDbl(newWeight)
    MsgBox "任务 " & taskId & " 权重已设置为: " & task.CFS_weight, vbInformation
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

    StartSchedulerIfNeeded

    MsgBox "调度器已启动。" & vbCrLf & vbCrLf & _
           "调度间隔: " & g_SchedulerIntervalMilliSec & " ms" & vbCrLf & _
           "调度模式: CFS (完全公平调度)" & vbCrLf & _
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

' 启动本工作簿的所有defined任务
Private Sub LuaSchedulerMenu_StartAllWorkbookTasks()
    On Error Resume Next

    ' 获取当前工作簿名称
    Dim wbName As String
    On Error Resume Next
    wbName = ActiveWorkbook.Name
    On Error GoTo ErrorHandler

    If wbName = vbNullString Then
        MsgBox "无法获取当前工作簿。", vbExclamation, "错误"
        Exit Sub
    End If

    If g_Tasks Is Nothing Then
        InitCoroutineSystem
    End If

    If g_Tasks.Count = 0 Then
        MsgBox "当前没有任何任务。", vbInformation, "提示"
        Exit Sub
    End If

    ' 收集本工作簿的所有defined任务
    Dim taskId As Variant
    Dim count As Long
    count = 0
    Dim task As TaskUnit

    For Each taskId In g_Tasks.Keys
        Set task = g_Tasks(taskId)
        If task.taskWorkbook = wbName Then
            If task.taskStatus = "defined" Then
                StartLuaCoroutine "Task_" & CStr(task.taskId)
                count = count + 1
            End If
        End If
    Next

    If count = 0 Then
        MsgBox "工作簿 [" & wbName & "] 没有 defined 状态的任务。", vbInformation, "提示"
    Else
        StartSchedulerIfNeeded
        MsgBox "已启动工作簿 [" & wbName & "] 的 " & count & " 个任务。", vbInformation, "启动完成"
    End If

    Exit Sub
ErrorHandler:
    MsgBox "启动任务时出错: " & Err.Description, vbCritical, "错误"
End Sub

' 批量启动所有 defined 状态的任务
Private Sub LuaSchedulerMenu_StartAllDefinedTasks()
    Dim taskId As Variant
    Dim count As Integer
    count = 0

    If Not InitLuaState() Then
        MsgBox "Lua状态初始化失败", vbCritical
        Exit Sub
    End If
    If g_Tasks.Count = 0 Then
        MsgBox "当前没有任何任务。", vbInformation, "提示"
        Exit Sub
    End If

    For Each taskId In g_Tasks.Keys
        If g_Tasks(taskId).taskStatus = "defined" Then
            StartLuaCoroutine CStr(taskId)
            count = count + 1
        End If
    Next

    MsgBox "已启动 " & count & " 个任务", vbInformation
End Sub

' 清理所有已完成或错误的任务
Private Sub LuaSchedulerMenu_CleanupFinishedTasks()
    On Error Resume Next
    If g_Tasks Is Nothing Then
        InitCoroutineSystem
    End If
    If g_Tasks.Count = 0 Then
        MsgBox "当前没有任务需要清理。", vbInformation, "清理任务"
        Exit Sub
    End If


    ' 收集需要清理的任务
    Dim taskId As Variant
    Dim count As Integer
    count = 0
    Dim status As String
    For Each taskId In g_Tasks.Keys
        status = g_Tasks(taskId).taskStatus
        If status = "done" Or status = "error" Then
            CollectionRemove g_TaskQueue, CStr(taskId)
            count = count +1
        End If
    Next

    MsgBox "已清理 " & count & " 个已完成或错误的任务。" & vbCrLf & _
           "剩余任务: " & g_Tasks.Count, vbInformation, "清理完成"
End Sub

' 清理特定工作簿的任务
Private Sub LuaSchedulerMenu_CleanupWorkbookTasks()
    Dim wb As String
    wb = Selection.Worksheet.Parent.Name
    CleanupWorkbookTasks wb
    MsgBox "已清理工作簿 " & wb & " 的任务。", vbInformation
End Sub

' 清空所有任务队列
Private Sub LuaSchedulerMenu_ClearAllTasks()
    Dim result As VbMsgBoxResult
    result = MsgBox("确定要清空所有任务吗？" & vbCrLf & vbCrLf & _
                    "这将删除所有任务数据，无法恢复！", _
                    vbExclamation + vbYesNo, "确认清空")

    If result = vbNo Then Exit Sub

    ' 停止调度器
    g_SchedulerRunning = False

    ' 清空所有 Dictionary
    If Not g_Tasks Is Nothing Then
        Set g_TaskQueue = New Collection
    End If

    MsgBox "所有任务已清空。", vbInformation, "清空完成"
End Sub

' 显示所有任务（按工作簿分组）
Private Sub LuaSchedulerMenu_ShowAllTasks()
    On Error GoTo ErrorHandler
    If g_Tasks Is Nothing Then InitCoroutineSystem
    If g_Tasks.Count = 0 Then
        MsgBox "当前没有任何任务。", vbInformation, "任务列表"
        Exit Sub
    End If

    ' 按工作簿分组统计
    Dim taskId As Variant
    Dim task As TaskUnit
    Dim definedCount As Integer, runningCount As Integer, yieldedCount As Integer
    Dim doneCount As Integer, errorCount As Integer, pausedCount As Integer
    
    ' 统计各状态任务数
    For Each taskId In g_Tasks.Keys
        Set task = g_Tasks(taskId)
        Select Case task.taskStatus
            Case "defined": definedCount = definedCount + 1
            Case "running": runningCount = runningCount + 1
            Case "yielded": yieldedCount = yieldedCount + 1
            Case "done": doneCount = doneCount + 1
            Case "error": errorCount = errorCount + 1
            Case "paused": pausedCount = pausedCount + 1
        End Select
    Next

    ' 按工作簿统计任务数
    Dim wbTaskCount As Object
    Set wbTaskCount = CreateObject("Scripting.Dictionary")
    
    For Each taskId In g_Tasks.Keys
        Set task = g_Tasks(taskId)
        If wbTaskCount.Exists(task.taskWorkbook) Then
            wbTaskCount(task.taskWorkbook) = wbTaskCount(task.taskWorkbook) + 1
        Else
            wbTaskCount.Add task.taskWorkbook, 1
        End If
    Next

    ' 构建消息
    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  Lua 协程任务管理器" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf
    msg = msg & "任务总数: " & g_Tasks.Count & vbCrLf
    msg = msg & "活跃队列: " & g_TaskQueue.Count & vbCrLf
    msg = msg & "调度器: " & IIf(g_SchedulerRunning, "运行中", "已停止") & vbCrLf
    msg = msg & vbCrLf & "按工作簿分组:" & vbCrLf

    Dim wbName As Variant
    For Each wbName In wbTaskCount.Keys
        msg = msg & "  [" & wbName & "]: " & wbTaskCount(wbName) & " 个任务" & vbCrLf
    Next

    msg = msg & "----------------------------------------" & vbCrLf
    msg = msg & "状态统计:" & vbCrLf
    msg = msg & "   已定义: " & definedCount & vbCrLf
    msg = msg & "   运行中: " & runningCount & vbCrLf
    msg = msg & "   中止中: " & yieldedCount & vbCrLf
    msg = msg & "   已完成: " & doneCount & vbCrLf
    msg = msg & "   错误: " & errorCount & vbCrLf
    msg = msg & "   已暂停: " & pausedCount & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf

    ' 详细列出每个任务
    For Each taskId In g_Tasks.Keys
        Set task = g_Tasks(CStr(taskId))
        msg = msg & "【任务 #" & task.taskId & "】" & vbCrLf
        msg = msg & "  ID: " & CStr(taskId) & vbCrLf
        msg = msg & "  函数: " & task.taskFunc & vbCrLf
        msg = msg & "  工作簿: " & task.taskWorkbook & vbCrLf
        msg = msg & "  单元格: " & task.taskCell & vbCrLf
        msg = msg & "  状态: " & task.taskStatus & vbCrLf
        msg = msg & "  进度: " & Format(task.taskProgress, "0.0") & "%" & vbCrLf

        ' 显示消息
        Dim msgText As String
        msgText = CStr(task.taskMessage)
        If Len(msgText) > 50 Then msgText = Left(msgText, 47) & "..."
        msg = msg & "  消息: " & msgText & vbCrLf

        ' 如果有错误，显示错误信息
        If task.taskStatus = "error" Then
            Dim errText As String
            errText = CStr(task.taskError)
            If Len(errText) > 60 Then errText = Left(errText, 57) & "..."
            msg = msg & "   错误: " & errText & vbCrLf
        End If

        ' 显示是否在活跃队列中
        If CollectionExists(g_TaskQueue, CStr(taskId)) Then
            msg = msg & "  队列: 是" & vbCrLf
        End If

        msg = msg & "----------------------------------------" & vbCrLf
    Next

    MsgBox msg, vbInformation, "Lua 协程任务列表 (" & g_Tasks.Count & " 个任务)"
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
           "如需更新 functions.lua，请手动运行 vbNullStringReloadFunctionsvbNullString。", _
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
    Dim seconds As Long
    seconds = Application.InputBox( _
            "请输入调度的间隔时间（>=10ms且<=60000ms）", _
            "调度参数", _
            g_SchedulerIntervalMilliSec, _
            Type:=1 _
        )

    If seconds = False Then Exit Sub
    If seconds < 10 Or seconds > 60000 Then
        MsgBox "间隔不能小于 10 ms。且不能大于 60 秒。", vbExclamation, "无效值"
        Exit Sub
    End If

    g_SchedulerIntervalMilliSec = seconds
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
    msg = msg & "调度算法: CFS (完全公平调度)" & vbCrLf
    msg = msg & "目标延迟: " & g_CFS_targetLatency & " ms" & vbCrLf
    msg = msg & "最小粒度: " & g_CFS_minGranularity & " ms" & vbCrLf
    msg = msg & "当前 min_vruntime: " & Format(g_CFS_minVruntime, "0.00") & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf
    If g_Tasks Is Nothing Then
        msg = msg & "任务总数: 0" & vbCrLf
    Else
        msg = msg & "任务总数: " & g_Tasks.Count & vbCrLf
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
    msg = msg & "调度算法: CFS" & vbCrLf
    msg = msg & "调度间隔: " & g_SchedulerIntervalMilliSec & " ms" & vbCrLf
    msg = msg & "当前 min_vruntime: " & Format(g_CFS_minVruntime, "0.00") & " ms" & vbCrLf
    msg = msg & vbCrLf & "当前状态: " & IIf(g_SchedulerRunning, "运行中", "已停止") & vbCrLf
    msg = msg & "活跃任务: " & g_TaskQueue.Count & vbCrLf

    MsgBox msg, vbInformation, "调度器性能统计"
End Sub

' 显示任务性能统计
Private Sub LuaPerfMenu_ShowTaskStats()
    If g_Tasks Is Nothing Or g_Tasks.Count = 0 Then
        MsgBox "当前没有任何任务。", vbInformation, "任务性能统计"
        Exit Sub
    End If

    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  任务性能统计" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf
    msg = msg & "任务总数: " & g_Tasks.Count & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf
    Dim taskNum As Long
    taskNum = 0
    Dim taskId As Variant
    Dim task As TaskUnit
    For Each taskId In g_Tasks.Keys
        taskNum = taskNum + 1
        Set task = g_Tasks(taskId)
        msg = msg & "【任务 #" & taskNum & "】" & vbCrLf
        msg = msg & "  ID: " & "Task_" & CStr(task.taskId) & vbCrLf
        msg = msg & "  函数: " & task.taskFunc & vbCrLf
        msg = msg & "  状态: " & task.taskStatus & vbCrLf
        If task.taskTickCount = 0 Then
            msg = msg & "  (尚未执行)" & vbCrLf
        Else
            msg = msg & "  调度次数: " & task.taskTickCount & vbCrLf
            msg = msg & "  总运行时间: " & Format(task.TaskTotalTime, "0.00") & " ms" & vbCrLf
            msg = msg & "  平均时间: " & Format(task.TaskTotalTime / task.taskTickCount, "0.00") & " ms" & vbCrLf
            msg = msg & "  上次运行: " & Format(task.taskLastTime, "0.00") & " ms" & vbCrLf
        End If
        msg = msg & "----------------------------------------" & vbCrLf
    Next
    MsgBox msg, vbInformation, "任务性能统计 (" & g_Tasks.Count & " 个任务)"

End Sub

' 显示工作簿性能统计
Private Sub LuaPerfMenu_ShowWorkbookStats()
    If g_Tasks Is Nothing Or g_Tasks.Count = 0 Then
        MsgBox "当前没有任何任务。", vbInformation, "工作簿性能统计"
        Exit Sub
    End If

    ' 统计每个工作簿的任务数
    Dim msg As String
    msg = "========================================" & vbCrLf
    msg = msg & "  工作簿性能统计" & vbCrLf
    msg = msg & "========================================" & vbCrLf & vbCrLf

    msg = msg & "工作簿总数: " & g_Workbooks.Count & vbCrLf
    msg = msg & "调度模式: CFS (完全公平调度)" & vbCrLf
    msg = msg & vbCrLf & "----------------------------------------" & vbCrLf

    Dim wb As WorkbookInfo
    Dim wbName As Variant
    Dim wbNum As Integer
    wbNum = 0
    For Each wbName In g_Workbooks.Keys
        Set wb = g_Workbooks(wbName)

        wbNum = wbNum + 1
        msg = msg & "【工作簿 #" & wbNum & "】" & vbCrLf
        msg = msg & "  名称: " & wbName & vbCrLf
        msg = msg & "  总调度次数: " & wb.wbTickCount & vbCrLf
        msg = msg & "  总运行时间: " & Format(wb.wbTotalTime, "0.00") & " ms" & vbCrLf
        msg = msg & "  平均时间: " & Format(wb.wbTotalTime / wb.wbTickCount, "0.00") & " ms" & vbCrLf
        msg = msg & "  上次调度: " & Format(wb.wbLastTime, "0.00") & " ms" & vbCrLf
        msg = msg & "----------------------------------------" & vbCrLf
    Next

    MsgBox msg, vbInformation, "工作簿性能统计 (" & g_Workbooks.Count & " 个工作簿)"
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
    Dim taskId As Variant
    Dim task As TaskUnit
    For Each taskId In g_Tasks.Keys
        Set task = g_Tasks(taskId)
        task.taskLastTime = 0
        task.TaskTotalTime = 0
        task.taskTickCount = 0
    Next

    ' 重置工作簿统计
    Dim wbName As Variant
    Dim wb As WorkbookInfo
    For Each wbName In g_Workbooks.Keys
        Set wb = g_Workbooks(wbName)
        wb.wbLastTime = 0
        wb.wbTotalTime = 0
        wb.wbTickCount = 0
    Next

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