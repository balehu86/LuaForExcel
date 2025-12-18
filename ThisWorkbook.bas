' ============================================
' ThisWorkbook - 工作簿事件处理
' ============================================
' 设计原因：
' 1. 管理 WorkbookRuntime 生命周期
' 2. Open 时注册，BeforeClose 时注销
' 3. 确保资源正确清理
' ============================================

Option Explicit

Private m_Runtime As WorkbookRuntime

' 工作簿打开时
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    
    ' 初始化调度器（进程级唯一）
    Scheduler.InitScheduler
    
    ' 创建本工作簿的运行时
    Set m_Runtime = New WorkbookRuntime
    m_Runtime.BindWorkbook Me
    
    ' 注册到核心注册表
    CoreRegistry.RegisterWorkbookRuntime Me, m_Runtime
    
    ' 启用菜单
    CoreUI.EnableLuaTaskMenu
    
    CoreUI.LogInfo "工作簿已初始化: " & Me.Name
    Exit Sub
    
ErrorHandler:
    CoreUI.LogError "工作簿初始化失败: " & Err.Description
End Sub

' 工作簿关闭前
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' 注销运行时（会自动清理资源）
    CoreRegistry.UnregisterWorkbookRuntime Me
    
    ' 清理菜单
    CoreUI.DisableLuaTaskMenu
    
    CoreUI.LogInfo "工作簿已清理: " & Me.Name
End Sub