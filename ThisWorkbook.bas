' ============================================
' ThisWorkbook.cls - 工作簿级事件处理
' ============================================
' 设计原因：
' 1. 监听本工作簿的 Open/BeforeClose 事件
' 2. 监听 Application 级别的 WorkbookOpen/WorkbookBeforeClose 事件
' 3. 为每个打开的工作簿创建/销毁 WorkbookRuntime
' 4. 统一管理右键菜单
' ============================================

Option Explicit

' ===== Application 事件监听 =====
' WithEvents 使得可以监听 Application 级别的事件
Private WithEvents App As Application

' ====工作簿打开时====
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    ' 初始化调度器
    Scheduler.InitScheduler
    ' 设置 Application 事件监听
    Set App = Application
    ' 为当前工作簿创建 Runtime
    InitializeWorkbookRuntime Me
    ' 为已打开的其他工作簿创建 Runtime
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If Not wb Is Me Then
            InitializeWorkbookRuntime wb
        End If
    Next
    ' 启用右键菜单
    LuaMenu.EnableLuaTaskMenu
    Exit Sub
ErrorHandler:
    MsgBox "初始化失败: " & Err.Description, vbCritical
End Sub

' ====工作簿关闭前====
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ' 停止调度器
    Scheduler.StopScheduler
    ' 销毁当前工作簿的 Runtime
    CleanupWorkbookRuntime Me
    ' 禁用右键菜单
    LuaMenu.DisableLuaTaskMenu
    ' 清除 Application 事件监听
    Set App = Nothing
End Sub
' ============================================
' Application 事件处理
' ============================================

' ====其他工作簿打开时====
Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    
    ' 跳过当前工作簿（已在 Workbook_Open 中处理）
    If Wb Is Me Then Exit Sub
    
    ' 为新打开的工作簿创建 Runtime
    InitializeWorkbookRuntime Wb
End Sub

' ====其他工作簿关闭前====
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    
    ' 跳过当前工作簿（已在 Workbook_BeforeClose 中处理）
    If Wb Is Me Then Exit Sub
    
    ' 销毁关闭的工作簿的 Runtime
    CleanupWorkbookRuntime Wb
End Sub

' ============================================
' 辅助方法
' ============================================

' 初始化工作簿的 Runtime
Private Sub InitializeWorkbookRuntime(wb As Workbook)
    On Error GoTo ErrorHandler
    
    ' 检查是否已存在 Runtime
    If Not CoreRegistry.GetRuntimeByWorkbook(wb) Is Nothing Then
        Exit Sub
    End If
    
    ' 创建新的 Runtime
    Dim rt As New WorkbookRuntime
    rt.BindWorkbook wb
    
    ' 注册到全局注册表
    CoreRegistry.RegisterWorkbookRuntime wb, rt
    
    Debug.Print "[Init] 已为工作簿创建 Runtime: " & wb.Name
    Exit Sub
    
ErrorHandler:
    Debug.Print "[Init Error] " & wb.Name & ": " & Err.Description
End Sub

' 清理工作簿的 Runtime
Private Sub CleanupWorkbookRuntime(wb As Workbook)
    On Error Resume Next
    
    ' 从全局注册表注销
    CoreRegistry.UnregisterWorkbookRuntime wb
    
    Debug.Print "[Cleanup] 已清理工作簿 Runtime: " & wb.Name
End Sub

Private Sub Workbook_AddinInstall()
    Workbook_Open
End Sub

Private Sub Workbook_AddinUninstall()
    Workbook_BeforeClose False
End Sub