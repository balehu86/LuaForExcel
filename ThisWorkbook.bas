' ============================================
' ThisWorkbook - 工作簿事件处理
' ============================================
' 设计原因：
' 1. 管理 WorkbookRuntime 生命周期
' 2. Open 时注册，BeforeClose 时注销
' 3. 确保资源正确清理
' ============================================

' ============================================
' ThisWorkbook - 加载项生命周期管理（完整版）
' ============================================

Option Explicit

Private WithEvents App As Application

' 加载项打开时
Private Sub Workbook_Open()
    On Error Resume Next
    
    ' 初始化调度器
    Scheduler.InitScheduler
    
    ' 监听 Application 事件
    Set App = Application
    
    ' 启用右键菜单
    LuaMenu.EnableLuaTaskMenu
    
    Debug.Print "[LuaTask加载项] 已加载"
End Sub

' 加载项关闭前
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' 停止调度器
    Scheduler.StopScheduler
    
    ' 禁用菜单
    LuaMenu.DisableLuaTaskMenu
    
    ' 移除事件监听
    Set App = Nothing
    
    Debug.Print "[LuaTask加载项] 已卸载"
End Sub

' 当任何工作簿打开时
Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    
    ' 跳过加载项自身
    If Wb Is Me Then Exit Sub
    
    ' 为新工作簿创建运行时
    Dim rt As New WorkbookRuntime
    rt.BindWorkbook Wb
    
    ' 注册到全局注册表
    CoreRegistry.RegisterWorkbookRuntime Wb, rt
    
    Debug.Print "[LuaTask] 已为工作簿创建运行时: " & Wb.Name
End Sub

' 当任何工作簿关闭前
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    
    ' 跳过加载项自身
    If Wb Is Me Then Exit Sub
    
    ' 注销运行时
    CoreRegistry.UnregisterWorkbookRuntime Wb
    
    Debug.Print "[LuaTask] 已清理工作簿运行时: " & Wb.Name
End Sub

Private Sub Workbook_AddinInstall()
    On Error Resume Next
    LuaMenu.EnableLuaTaskMenu
End Sub

Private Sub Workbook_AddinUninstall()
    On Error Resume Next
    LuaMenu.DisableLuaTaskMenu
    Scheduler.StopScheduler
End Sub
