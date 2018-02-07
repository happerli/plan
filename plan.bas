Attribute VB_Name = "模块1"
Option Explicit

Const allCol As Integer = 365
Const allRow As Integer = 300
Const widthDayCol As Integer = 2

Dim btnStatus As Range
Dim btnDate As Range
Dim btnAll As Range
Dim celStatus As Range
Dim colStatus As String
Dim colStatusCtrlStart As String
Dim colStatusCtrlEnd As String
Dim strStatus As String
Dim arrStatusClr
Dim colStart As String
Dim colEnd As String

Dim celPeriod As Range
Dim strPeriod As String
Dim clrToday As Integer
Dim clrWeekend As Integer
Dim clrMonth As Integer
Dim clrDay As Integer
Dim lineDay As Integer
Dim colDateStart As String
Dim rowTitle As Integer
Dim bottomLine As Integer

Sub init()
    Range("A1").Value = "日期区间"
    Range("A1").Interior.ColorIndex = 23
    Set celPeriod = Range("B1")
    Set btnStatus = Range("C1")
    Set btnDate = Range("D1")
    Set btnAll = Range("E1")
    
    colStatus = "D"
    colStatusCtrlStart = "A"
    colStatusCtrlEnd = "F"
    colStart = "E"
    colEnd = "F"
    rowTitle = 2
    
    colDateStart = "G"
    clrToday = 6
    clrWeekend = 15
    clrMonth = 3
    lineDay = xlDash
    clrDay = 4
    '----------------------------------------------------------------------------------
    Set celStatus = Range(colStatus & (rowTitle + 1) & ":" & colStatus & (rowTitle + allRow))
    strStatus = "未开始,进行中,已完成,推迟,无效,等待中"
    arrStatusClr = Array(0, 34, 50, 48, 16, 18, 3, 56) '3: is exceed the time limit. 56:error
    '----------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------
    strPeriod = "所有,前一月,前两周,前一周,本周,本月,后一周,后两周,后一月,截止现在,现在以后"
    
    
    '----------------------------------------------------------------------------------
    bottomLine = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).row
End Sub

Sub fillList(cel As Range, str As String)
    With cel.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=str
    End With
End Sub

Sub createButtonRefresh()
    Dim btn As Button
    Application.ScreenUpdating = False
    init
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
    Set btn = ActiveSheet.Buttons.Add(btnAll.Left, btnAll.top, btnAll.width, btnAll.Height)
    With btn
        .OnAction = "refreshAll"
        .Caption = "全部刷新"
        .Name = "全部刷新"
    End With
    Application.ScreenUpdating = True
End Sub

Sub refreshStatus()
    On Error Resume Next
    Application.ScreenUpdating = False
    init
    fillList celStatus, strStatus
    Dim status As String
    Dim arrStatus
    Dim today As Date
    Dim endDay As Date
    Dim i As Integer
    Dim m As Variant
    Dim clr As Integer
    
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
        Range(colStatusCtrlStart & i & ":" & colStatusCtrlEnd & i).Interior.ColorIndex = clr
    Next
    Application.ScreenUpdating = True
End Sub

Sub refreshDate()
    Application.ScreenUpdating = False
    init
    fillList celPeriod, strPeriod
    
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
        Range(cel.Offset(0, nf), cel.Offset(0, nl)).Interior.ColorIndex = clrDay
        Range(cel.Offset(0, nf), cel.Offset(0, nl)).Borders.LineStyle = lineDay
nxt:
    Next i
       
    Dim wkd As Integer
    Dim thisdate As Date
    For i = 0 To diff Step 1
        thisdate = DateAdd("d", i, firstDay)
        Range(colDateStart & rowTitle).Offset(0, i).Value = thisdate
        Range(colDateStart & rowTitle).Offset(0, i).ColumnWidth = widthDayCol
        Range(colDateStart & rowTitle).Offset(0, i).NumberFormatLocal = "d"
        wkd = Weekday(thisdate)
        If wkd = 1 Or wkd = 7 Then
            Range(colDateStart & rowTitle).Offset(0, i).ColumnWidth = 1
            Range(colDateStart & rowTitle & ":" & colDateStart & bottomLine).Offset(0, i).Interior.ColorIndex = clrWeekend
        End If
        If Day(thisdate) = 1 Then
            Range(colDateStart & rowTitle & ":" & colDateStart & rowTitle).Offset(0, i).Interior.ColorIndex = clrMonth
        End If
    
        If DateDiff("d", thisdate, today) = 0 Then
            Range(colDateStart & rowTitle & ":" & colDateStart & rowTitle).Offset(0, i).Interior.ColorIndex = clrToday
        End If
    Next
   
rtn:
    Application.ScreenUpdating = True
End Sub

Sub refreshAll()
    refreshStatus
    refreshDate
End Sub


Function FirstLast(row As Integer, colFirst As String, colLast As String, ByRef firstDay As Date, ByRef lastDay As Date)
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

Function FirstLastLoop(row As Integer, num As Integer, colFirst As String, colLast As String, ByRef firstDay As Date, ByRef lastDay As Date)
    Dim today As Date
    Dim i As Integer
    
    today = Date
    firstDay = DateAdd("m", 100, today)
    lastDay = DateAdd("m", -100, today)
    
    For i = 0 To num Step 1
        Call FirstLast(row + i, colFirst, colLast, firstDay, lastDay)
    Next
End Function

