Attribute VB_Name = "ģ��1"
Option Explicit

Const allCol As Integer = 365 '�������
Const allRow As Integer = 300 '�������
Const widthDayCol As Integer = 2 '���ڵ�Ԫ����

Dim btnStatus As Range 'ˢ��״̬��ťλ��
Dim btnDate As Range 'ˢ�����ڰ�ťλ��
Dim celStatus As Range '�������״̬����
Dim colStatus As String '����״̬��
Dim colStatusCtrlStart As String '״̬������ʼ��
Dim colStatusCtrlEnd As String '״̬������ֹ��
Dim strStatus As String '����״̬�б�
Dim arrStatusClr '����״̬��Ӧ�����ɫ
Dim colStart As String '������ʼ��������
Dim colEnd As String '���������������
Dim colTotal As String '������
Dim colRemain As String '���/ʣ������

Dim celPeriod As Range '��ʾ���������б�λ��
Dim strPeriod As String '��ʾ�����б�
Dim clrToday As Variant '������ɫ
Dim clrWeekend As Variant '��ĩ��ɫ
Dim clrMonth As Variant 'ÿ�¿�ʼ��ɫ
Dim clrDay As Variant
Dim lineDay As Integer
Dim colDateStart As String '������ʾ��һ��
Dim rowTitle As Integer '������
Dim bottomLine As Integer '���һ��

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
    strStatus = "δ��ʼ,������,�����,�Ƴ�,��Ч,�ȴ���"
    arrStatusClr = Array(RGB(255, 255, 255), RGB(204, 255, 255), RGB(160, 228, 200), RGB(191, 191, 191), RGB(128, 128, 128), RGB(250, 191, 143), RGB(255, 153, 153), RGB(255, 255, 0)) '3: is exceed the time limit. 56:error
    '----------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------
    strPeriod = "����,ǰһ��,ǰ����,ǰһ��,����,����,��һ��,������,��һ��,��ֹ����,�����Ժ�"
    
    
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
    Range("I1").Value = "��������:"
    Range("J1").Formula = "=TODAY()"
    Range("J1").Value = Format(Date, "yyyy/mm/dd")
    Range("I2").Value = "��������:"
    
    Range("A3").Value = "���"
    Range("B3").Value = "����"
    Range("C3").Value = "���ȼ�"
    Range("D3").Value = "����"
    Range("E3").Value = "״̬"
    Range("F3").Value = "���(%)"
    Range("G3").Value = "������"
    Range("H3").Value = "��ʼ��"
    Range("I3").Value = "������"
    Range("J3").Value = "������"
    Range("K3").Value = "���/ʣ��"
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
        .Caption = "ˢ��״̬"
        .Name = "ˢ��״̬"
    End With
    Set btn = ActiveSheet.Buttons.Add(btnDate.Left, btnDate.top, btnDate.width, btnDate.Height)
    With btn
        .OnAction = "refreshDate"
        .Caption = "ˢ������"
        .Name = "ˢ������"
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
        Case "ǰһ��"
            firstDay = DateAdd("m", -1, today)
            lastDay = today
        Case "ǰ����"
            firstDay = DateAdd("ww", -2, today)
            lastDay = today
        Case "ǰһ��"
            firstDay = DateAdd("ww", -1, today)
            lastDay = today
        Case "��ֹ����"
            Call FirstLastLoop(rowTitle + 1, allRow, colStart, colEnd, firstDay, lastDay)
            lastDay = today
        Case "��һ��"
            lastDay = DateAdd("m", 1, today)
            firstDay = today
        Case "������"
            lastDay = DateAdd("ww", 2, today)
            firstDay = today
        Case "��һ��"
            lastDay = DateAdd("ww", 1, today)
            firstDay = today
        Case "�����Ժ�"
            Call FirstLastLoop(rowTitle + 1, allRow, colStart, colEnd, firstDay, lastDay)
            firstDay = today
        Case "����"
            firstDay = Date - Weekday(Date, vbUseSystem) + 1
            lastDay = Date - Weekday(Date, vbUseSystem) + 7
        Case "����"
            firstDay = DateSerial(Year(Now), Month(Now), 1)  '���µ�һ��
            lastDay = DateSerial(Year(Now), Month(Now) + 1, 0) '�������һ��
        Case "����"
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

