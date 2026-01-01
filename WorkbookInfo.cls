' WorkbookInfo.cls
Option Explicit

Public wbName As String         ' workbook name

Public wbLastTime As Double    ' 上次调度时间(ms)
Public wbTotalTime As Double   ' 总运行时间(ms)
Public wbTickCount As Integer   ' 调度次数(统计用)

Private Sub Class_Initialize()
    wbLastTime = 0
    wbTotalTime = 0
    wbTickCount = 0
End Sub
