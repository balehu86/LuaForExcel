' ============================================
' CoreRegistry.bas - 运行时注册表（完整版）
' ============================================
' 设计原因：
' 1. 封装所有 Dictionary 访问
' 2. 提供 taskId → Runtime 路由
' 3. 管理 Workbook → Runtime 映射
' 4. 新增 cellAddr → taskId 索引
' 5. 提供 WorkbookKey 工具
' ============================================

Option Explicit

' ===== 全局注册表（仅本模块访问）=====
Private g_RuntimeByWorkbook As Object  ' wbKey → WorkbookRuntime
Private g_RuntimeByTaskIndex As Object ' taskId → WorkbookRuntime
Private g_TaskIdByCellAddr As Object   ' cellAddr → taskId

' ====初始化====
Private Sub InitRegistry()
    If g_RuntimeByWorkbook Is Nothing Then
        Set g_RuntimeByWorkbook = CreateObject("Scripting.Dictionary")
        Set g_RuntimeByTaskIndex = CreateObject("Scripting.Dictionary")
        Set g_TaskIdByCellAddr = CreateObject("Scripting.Dictionary")
    End If
End Sub

' ====运行时管理====
' 注册工作簿运行时
Public Sub RegisterWorkbookRuntime(wb As Workbook, rt As WorkbookRuntime)
    InitRegistry
    Dim wbKey As String
    wbKey = GetWorkbookKey(wb)
    g_RuntimeByWorkbook(wbKey) = rt
    
    ' 注册到调度器
    Scheduler.RegisterRunnable rt
End Sub

' 注销工作簿运行时
Public Sub UnregisterWorkbookRuntime(wb As Workbook)
    InitRegistry
    Dim wbKey As String
    wbKey = GetWorkbookKey(wb)

    If g_RuntimeByWorkbook.Exists(wbKey) Then
        Dim rt As WorkbookRuntime
        Set rt = g_RuntimeByWorkbook(wbKey)
        ' 从调度器注销
        Scheduler.UnregisterRunnable rt
        ' 清理资源
        rt.Dispose
        ' 从注册表移除
        g_RuntimeByWorkbook.Remove wbKey
    End If
End Sub

' 根据 Workbook 获取运行时
Public Function GetRuntimeByWorkbook(wb As Workbook) As WorkbookRuntime
    InitRegistry
    Dim wbKey As String
    wbKey = GetWorkbookKey(wb)
    
    If g_RuntimeByWorkbook.Exists(wbKey) Then
        Set GetRuntimeByWorkbook = g_RuntimeByWorkbook(wbKey)
    Else
        Set GetRuntimeByWorkbook = Nothing
    End If
End Function

' ====Task 路由表（GlobalTaskIndex）====
' 注册任务到运行时的映射
Public Sub RegisterTask(taskId As String, rt As WorkbookRuntime, cellAddr As String)
    InitRegistry
    g_RuntimeByTaskIndex(taskId) = rt
    
    ' 新增：注册 cellAddr → taskId 映射
    If cellAddr <> "" Then
        g_TaskIdByCellAddr(cellAddr) = taskId
    End If
End Sub

' 根据 taskId 解析运行时
Public Function ResolveRuntime(taskId As String) As WorkbookRuntime
    InitRegistry
    If g_RuntimeByTaskIndex.Exists(taskId) Then
        Set ResolveRuntime = g_RuntimeByTaskIndex(taskId)
    Else
        Set ResolveRuntime = Nothing
    End If
End Function

' 注销任务
Public Sub UnregisterTask(taskId As String, cellAddr As String)
    InitRegistry
    
    If g_RuntimeByTaskIndex.Exists(taskId) Then
        g_RuntimeByTaskIndex.Remove taskId
    End If
    
    ' 新增：清理 cellAddr 索引
    If cellAddr <> "" And g_TaskIdByCellAddr.Exists(cellAddr) Then
        g_TaskIdByCellAddr.Remove cellAddr
    End If
End Sub

' ============================================
' 单元格地址索引（新增）
' ============================================

' 根据单元格地址查找任务 ID
Public Function FindTaskByCell(cellAddr As String) As String
    InitRegistry
    
    If g_TaskIdByCellAddr.Exists(cellAddr) Then
        FindTaskByCell = g_TaskIdByCellAddr(cellAddr)
    Else
        FindTaskByCell = ""
    End If
End Function

' 检查单元格是否已有任务
Public Function CellHasTask(cellAddr As String) As Boolean
    InitRegistry
    CellHasTask = g_TaskIdByCellAddr.Exists(cellAddr)
End Function

' ============================================
' 工具函数（WorkbookKey 生成）
' ============================================

' 生成唯一的 Workbook 标识
' 设计原因：FullName 可能重复（不同路径同名文件）
' 使用 Name + CreationDate 组合
Public Function GetWorkbookKey(wb As Workbook) As String
    On Error Resume Next
    GetWorkbookKey = wb.Name & "_" & Format(wb.BuiltinDocumentProperties("Creation Date"), "yyyymmddhhnnss")
    If Err.Number <> 0 Then
        ' 回退方案：使用 Name + 当前时间
        GetWorkbookKey = wb.Name & "_" & Format(Now, "yyyymmddhhnnss")
    End If
End Function

' 从外部地址提取 WorkbookKey
' 例如：'[Book1.xlsx]Sheet1'!$A$1 → Book1.xlsx
Public Function ExtractWorkbookFromAddress(addr As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(addr, "[")
    endPos = InStr(addr, "]")
    
    If startPos > 0 And endPos > startPos Then
        ExtractWorkbookFromAddress = Mid(addr, startPos + 1, endPos - startPos - 1)
    Else
        ExtractWorkbookFromAddress = ""
    End If
End Function

' 在 CoreRegistry.bas 中添加自动初始化函数
Public Function EnsureRuntimeForWorkbook(wb As Workbook) As WorkbookRuntime
    InitRegistry
    
    ' 如果是加载宏本身，返回 Nothing
    If wb.IsAddin Then
        Set EnsureRuntimeForWorkbook = Nothing
        Exit Function
    End If
    
    ' 检查是否已存在
    Dim rt As WorkbookRuntime
    Set rt = GetRuntimeByWorkbook(wb)
    
    If Not rt Is Nothing Then
        Set EnsureRuntimeForWorkbook = rt
        Exit Function
    End If
    
    ' 创建新的 Runtime
    Set rt = New WorkbookRuntime
    rt.BindWorkbook wb
    
    ' 注册
    RegisterWorkbookRuntime wb, rt
    
    Set EnsureRuntimeForWorkbook = rt
End Function