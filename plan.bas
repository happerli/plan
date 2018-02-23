Attribute VB_Name = "模块1"
Option Explicit

Const allCol As Integer = 365 '最大列数
Const allRow As Integer = 300 '最大行数
Const widthDayCol As Integer = 2 '日期单元格宽度

Dim btnStatus As Range '刷新状态按钮位置
Dim btnDate As Range '刷新日期按钮位置
Dim celStatus As Range '填充任务状态区域
Dim colStatus As String '任务状态列
Dim colStatusCtrlStart As String '状态控制起始列
Dim colStatusCtrlEnd As String '状态控制终止列
Dim strStatus As String '任务状态列表
Dim arrStatusClr '任务状态对应填充颜色
Dim colStart As String '任务起始日所在列
Dim colEnd As String '任务结束日所在列
Dim colTotal As String '总天数
Dim colRemain As String '完成/剩余天数

Dim celPeriod As Range '显示区间下拉列表位置
Dim strPeriod As String '显示区间列表
Dim clrToday As Variant '当日颜色
Dim clrWeekend As Variant '周末颜色
Dim clrMonth As Variant '每月开始颜色
Dim clrDay As Variant
Dim lineDay As Integer
Dim colDateStart As String '日期显示第一列
Dim rowTitle As Integer '标题行
Dim bottomLine As Integer '最后一行

Private Sub init_inner()
    Set btnStatus = Range("K1")
    Set btnDate = Range("K2")
    Set celPeriod = Range("J2")
    
    colStatus = "E"
    colStatusCtrlStart = "A"
    colStatusCtrlEnd = "K"
    colStart = "H"
    colEnd = "I"
    colTotal = "J"
    colRemain = "K"
    rowTitle = 3
    
    colDateStart = "L"
    clrToday = RGB(255, 80, 80)
    clrWeekend = RGB(192, 192, 192)
    clrMonth = RGB(127, 221, 212)
    lineDay = xlDash
    clrDay = RGB(83, 141, 213)
    '----------------------------------------------------------------------------------
    Set celStatus = Range(colStatus & (rowTitle + 1) & ":" & colStatus & (rowTitle + allRow))
    strStatus = "未开始,进行中,已完成,推迟,无效,等待中"
    arrStatusClr = Array(RGB(255, 255, 255), RGB(204, 255, 255), RGB(160, 228, 200), RGB(191, 191, 191), RGB(128, 128, 128), RGB(250, 191, 143), RGB(255, 153, 153), RGB(255, 255, 0)) '3: is exceed the time limit. 56:error
    '----------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------
    strPeriod = "所有,前一月,前两周,前一周,本周,本月,后一周,后两周,后一月,截止现在,现在以后"
    
    
    '----------------------------------------------------------------------------------
    bottomLine = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).row
End Sub

Sub init()
    init_inner
    Dim btn As Button
    Application.ScreenUpdating = False
    
    With Range("A1:H2")
        .Merge
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(248, 248, 186)
    End With
    Range("I1").Value = "今日日期:"
    Range("J1").Formula = "=TODAY()"
    Range("J1").Value = Format(Date, "yyyy/mm/dd")
    Range("I2").Value = "日期区间:"
    
    Range("A3").Value = "序号"
    Range("B3").Value = "任务"
    Range("C3").Value = "优先级"
    Range("D3").Value = "详情"
    Range("E3").Value = "状态"
    Range("F3").Value = "完成(%)"
    Range("G3").Value = "负责人"
    Range("H3").Value = "开始日"
    Range("I3").Value = "结束日"
    Range("J3").Value = "总天数"
    Range("K3").Value = "完成/剩余"
    Range("A1:K3").VerticalAlignment = xlCenter
    Range("A1:K3").HorizontalAlignment = xlCenter
    Range("A3:K3").Font.Bold = True
    With Range("A1:K3").Borders
        .LineStyle = xlContinuous
        .ColorIndex = 1
        .Weight = xlThin
    End With
    Range("A3:K3").Interior.Color = RGB(132, 174, 224)
    Range("I1:I2").Interior.Color = RGB(215, 229, 245)
    
    ActiveSheet.Buttons.Delete
    Set btn = ActiveSheet.Buttons.Add(btnStatus.Left + 1, btnStatus.top + 1, btnStatus.width - 1, btnStatus.Height - 1)
    With btn
        .OnAction = "refreshStatus"
        .Caption = "刷新状态"
        .Name = "刷新状态"
    End With
    Set btn = ActiveSheet.Buttons.Add(btnDate.Left, btnDate.top, btnDate.width, btnDate.Height)
    With btn
        .OnAction = "refreshDate"
        .Caption = "刷新日期"
        .Name = "刷新日期"
    End With
    
    Range("A1:K3").Rows.AutoFit
    Range("A1:K3").Columns.AutoFit
    
    fillList celPeriod, strPeriod
    Application.ScreenUpdating = True
End Sub

Private Sub fillList(cel As Range, str As String)
    With cel.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=str
    End With
End Sub


Private Sub refreshStatus()
    On Error Resume Next
    Application.ScreenUpdating = False
    init_inner
    ToGroup
    fillList celStatus, strStatus
    calcDays
    
    Dim status As String
    Dim arrStatus
    Dim today As Date
    Dim endDay As Date
    Dim i As Integer
    Dim m As Variant
    Dim clr As Variant
    
    arrStatus = Split(strStatus, ",")
    today = Date
    
    For i = rowTitle + 1 To bottomLine
        clr = arrStatusClr(7)
        If IsEmpty(Range(colStatus & i)) Or IsEmpty(Range(colEnd & i)) Or Range(colEnd & i).EntireRow.Hidden Then
            GoTo work
        End If
        
        endDay = Range(colEnd & i).Value
        status = Range(colStatus & i).Value
        m = Application.Match(status, arrStatus, 0)
        
        If Not IsError(m) Then
            m = m - 1
            If m >= 0 And m < 6 Then
                clr = arrStatusClr(m)
            End If
        End If
        
        If (m = 0 Or m = 1) And endDay <= today Then
            clr = arrStatusClr(6)
        End If
work:
        With Range(colStatusCtrlStart & i & ":" & colStatusCtrlEnd & i)
            .Interior.Color = clr
            .Borders.LineStyle = 1
            .Borders.Weight = xlHairline
        End With
    Next
    Application.ScreenUpdating = True
End Sub

Private Sub calcDays()
    Dim today As Date
    Dim startDay As Date
    Dim endDay As Date
    Dim i As Integer
    Dim totalDays As Integer
    Dim remainDays As Integer
    Dim passDays As Integer
    
    today = Date
    
    For i = rowTitle + 1 To bottomLine
        If IsEmpty(Range(colStart & i)) Or IsEmpty(Range(colEnd & i)) Or Range(colEnd & i).EntireRow.Hidden Then
            GoTo work
        End If
        
        endDay = Range(colEnd & i).Value
        startDay = Range(colStart & i).Value
        totalDays = DateDiff("d", startDay, endDay) + 1
        Range(colTotal & i).Value = totalDays
        If today < startDay Then
            passDays = 0
            remainDays = totalDays
        ElseIf today > endDay Then
            passDays = totalDays
            remainDays = 0
        Else
            passDays = DateDiff("d", startDay, today)
            remainDays = DateDiff("d", today, endDay) + 1
        End If
        With Range(colRemain & i)
            .NumberFormatLocal = "@"
            .Value = passDays & "/" & remainDays
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With
work:
    Next
    Application.ScreenUpdating = True
End Sub

Private Sub refreshDate()
    Application.ScreenUpdating = False
    init_inner
    
    Dim arrPeriod
    Dim Period As String
    Dim cel As Range
    
    Set cel = Range(colDateStart + ":" + colDateStart)
    Columns(colDateStart + ":" + Split(cel.Offset(0, allCol).Address(1, 0), ":")(0)).Delete
    
    Dim today As Date
    Dim firstDay As Date
    Dim lastDay As Date
    
    today = Date
    Period = celPeriod.Value
    
    Select Case Period
        Case "前一月"
            firstDay = DateAdd("m", -1, today)
            lastDay = today
        Case "前两周"
            firstDay = DateAdd("ww", -2, today)
            lastDay = today
        Case "前一周"
            firstDay = DateAdd("ww", -1, today)
            lastDay = today
        Case "截止现在"
            Call FirstLastLoop(rowTitle + 1, allRow, colStart, colEnd, firstDay, lastDay)
            lastDay = today
        Case "后一月"
            lastDay = DateAdd("m", 1, today)
            firstDay = today
        Case "后两周"
            lastDay = DateAdd("ww", 2, today)
            firstDay = today
        Case "后一周"
            lastDay = DateAdd("ww", 1, today)
            firstDay = today
        Case "现在以后"
            Call FirstLastLoop(rowTitle + 1, allRow, colStart, colEnd, firstDay, lastDay)
            firstDay = today
        Case "本周"
            firstDay = Date - Weekday(Date, vbUseSystem) + 1
            lastDay = Date - Weekday(Date, vbUseSystem) + 7
        Case "本月"
            firstDay = DateSerial(Year(Now), Month(Now), 1)  '本月第一天
            lastDay = DateSerial(Year(Now), Month(Now) + 1, 0) '本月最后一天
        Case "所有"
            Call FirstLastLoop(rowTitle + 1, allRow, colStart, colEnd, firstDay, lastDay)
        Case ""
            GoTo rtn
        Case Else
            MsgBox "Unknown date period!"
            GoTo rtn
    End Select

    Dim diff As Integer
    diff = DateDiff("d", firstDay, lastDay)

    Dim rf As Date
    Dim rl As Date
    Dim uf As Date
    Dim ul As Date
    Dim i, nf, nl As Integer
       
    For i = rowTitle + 1 To bottomLine
        If IsEmpty(Range(colStart & i)) Or IsEmpty(Range(colEnd & i)) Or Range(colStart & i).EntireRow.Hidden Then
            GoTo nxt
        End If
        
        rf = Range(colStart & i).Value
        rl = Range(colEnd & i).Value
        If rf > rl Then
            Range(colStart & i).Select
            MsgBox "Date error!"
            GoTo rtn
        End If
        
        If rf > lastDay Or rl < firstDay Then
            GoTo nxt
        End If
            
        If rf <= firstDay Then
            uf = firstDay
        Else
            uf = rf
        End If
        
        If rl >= lastDay Then
            ul = lastDay
        Else
            ul = rl
        End If
        nf = DateDiff("d", firstDay, uf)
        nl = DateDiff("d", firstDay, ul)
        Set cel = Range(colDateStart & i)
        Range(cel.Offset(0, nf), cel.Offset(0, nl)).Interior.Color = clrDay
        Range(cel.Offset(0, nf), cel.Offset(0, nl)).Borders.LineStyle = lineDay
nxt:
    Next i
       
    Dim wkd As Integer
    Dim beginidx As Integer
    Dim endidx As Integer
    Dim thisdate As Date
    
    beginidx = 0
    endidx = diff
    For i = diff To 0 Step -1
        thisdate = DateAdd("d", i, firstDay)
        Range(colDateStart & rowTitle - 1).Offset(0, i).Value = thisdate
        Range(colDateStart & rowTitle - 1).Offset(0, i).ColumnWidth = widthDayCol
        Range(colDateStart & rowTitle - 1).Offset(0, i).NumberFormatLocal = "d"
        Range(colDateStart & rowTitle).Offset(0, i).Value = thisdate
        Range(colDateStart & rowTitle).Offset(0, i).ColumnWidth = widthDayCol
        Range(colDateStart & rowTitle).Offset(0, i).NumberFormatLocal = "aaa"
        wkd = Weekday(thisdate)
        
        If wkd = 1 Or wkd = 7 Then
            Range(colDateStart & rowTitle - 1).Offset(0, i).ColumnWidth = widthDayCol
            Range(colDateStart & rowTitle - 1 & ":" & colDateStart & bottomLine).Offset(0, i).Interior.Color = clrWeekend
        End If
        
        If Day(thisdate) = 1 Then
            Range(colDateStart & rowTitle - 1 & ":" & colDateStart & rowTitle).Offset(0, i).Interior.Color = clrMonth
        End If
        
        If Day(thisdate) = 1 Then
            With Range(Range(colDateStart & rowTitle - 2).Offset(0, i), Range(colDateStart & rowTitle - 2).Offset(0, endidx))
                .Merge
                .Value = thisdate
                .NumberFormatLocal = "yyyy/mm"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
            
            endidx = i - 1
        End If
            
        If DateDiff("d", thisdate, today) = 0 Then
            Range(colDateStart & rowTitle - 1 & ":" & colDateStart & bottomLine).Offset(0, i).Interior.Color = clrToday
            
        End If
    Next
    With Range(Range(colDateStart & rowTitle - 2).Offset(0, 0), Range(colDateStart & rowTitle - 2).Offset(0, endidx))
        .Merge
        .Value = thisdate
        .NumberFormatLocal = "yyyy/mm"
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
   
rtn:
    Application.ScreenUpdating = True
End Sub

Private Sub refreshAll()
    refreshStatus
    refreshDate
End Sub


Private Function FirstLast(row As Integer, colFirst As String, colLast As String, ByRef firstDay As Date, ByRef lastDay As Date)
    If IsEmpty(Range(colFirst & row)) Or IsEmpty(Range(colLast & row)) Or Range(colFirst & row).EntireRow.Hidden Then
        GoTo rtn
    End If
    
    If Range(colFirst & row).Value < firstDay Then
        firstDay = Range(colFirst & row).Value
    End If
    
    If Range(colLast & row).Value > lastDay Then
        lastDay = Range(colLast & row).Value
    End If
    
rtn:
End Function

Private Function FirstLastLoop(row As Integer, num As Integer, colFirst As String, colLast As String, ByRef firstDay As Date, ByRef lastDay As Date)
    Dim today As Date
    Dim i As Integer
    
    today = Date
    firstDay = DateAdd("m", 100, today)
    lastDay = DateAdd("m", -100, today)
    
    For i = 0 To num Step 1
        Call FirstLast(row + i, colFirst, colLast, firstDay, lastDay)
    Next
End Function

Private Sub ToGroup()
    Dim cell As Range
    Dim ji As Integer
    Dim sv As String
    
    Rows.ClearOutline
    For Each cell In Range("A" & (rowTitle + 1), "A" & bottomLine)
        sv = cell.Value
        ji = Len(sv) - Len(Replace(sv, ".", "")) + 1
        cell.EntireRow.OutlineLevel = ji
        cell.Offset(0, 1).IndentLevel = ji - 1
        If ji < 2 Then
            cell.EntireRow.Font.Bold = True
        Else
            cell.EntireRow.Font.Bold = False
        End If
    Next cell
    
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlAbove
        .SummaryColumn = xlRight
    End With
End Sub

