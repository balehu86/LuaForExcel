' WatchInfo.cls
Option Explicit
Public watchCell As String        ' LuaWatch 公式所在单元格地址（外部格式）
Public watchTaskId As String      ' 监控的任务ID
Public watchTask As TaskUnit      ' 监控的任务实例
Public watchField As String       ' 监控的字段名
Public watchTargetCell As String  ' 写入目标单元格地址（外部格式）
Public watchDirection As Integer  ' 写入方向 0=右,1=上,2=左,3=下
Public watchLastValue As Variant  ' 上次写入的值
Public watchWorkbook As String    ' 所属工作簿名称
Public watchDirty As Boolean      ' 脏标记，True 表示需要刷新
Private Sub Class_Initialize()
    watchDirection = 0
    watchLastValue = Empty
    watchDirty = True  ' 新建时默认需要首次刷新
End Sub