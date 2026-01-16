'===============================================================================
' 외주처 생산계획 자동화 시스템 - VBA 매크로 v4.0
' CAPA 로직 수정: 한 라인의 하루에 한 제품만 배치
'===============================================================================

Option Explicit

Public Const SHEET_RAW As String = "계획"
Public Const SHEET_VENDORS As String = "설정_외주처"
Public Const SHEET_MAPPING As String = "설정_고객매칭"
Public Const SHEET_OUTPUT As String = "auto-planning"

Public Const COL_TRANSFER_DATE As Integer = 1
Public Const COL_STATUS As Integer = 2
Public Const COL_SPECIAL_PROCESS As Integer = 3
Public Const COL_VENDOR As Integer = 4
Public Const COL_MANAGER As Integer = 5
Public Const COL_PRODUCT_CODE As Integer = 6
Public Const COL_PRODUCT_NAME As Integer = 7
Public Const COL_PROCESS_TYPE As Integer = 8
Public Const COL_QUANTITY As Integer = 9
Public Const COL_MFG_DATE As Integer = 10
Public Const COL_MATERIAL_DATE As Integer = 11
Public Const COL_DELIVERY_DATE As Integer = 12
Public Const COL_URGENCY As Integer = 13
Public Const COL_NIGHT_SHIFT As Integer = 14
Public Const COL_REMARKS As Integer = 15

Public Type VendorInfo
    ID As String
    Name As String
    LineCount As Integer
    DailyCapacity As Long
    Capabilities As String
    MonthlyTarget As Long
    Priority As Integer
End Type

Public Type ProductionItem
    RowNum As Long
    TransferDate As String
    Status As String
    SpecialProcess As String
    AssignedVendor As String
    Manager As String
    ProductCode As String
    ProductName As String
    ProcessType As String
    Quantity As Long
    DeliveryDate As String
    ClientCode As String
End Type

Public Type AllocationResult
    Item As ProductionItem
    VendorName As String
    LineNumber As Integer
    StartDate As Date
    DaysNeeded As Integer
End Type

Public Vendors() As VendorInfo
Public VendorCount As Integer
Public ClientMappings As Object

' 라인 스케줄 (전역) - Key: "외주처_라인_yyyy-mm-dd" -> True
Public LineSchedule As Object

'===============================================================================
' 메인 실행
'===============================================================================
Public Sub 자동배정실행()
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' 라인 스케줄 초기화
    Set LineSchedule = CreateObject("Scripting.Dictionary")
    
    If Not LoadSettings() Then
        MsgBox "설정을 로드할 수 없습니다." & vbCrLf & "'초기설정' 매크로를 먼저 실행해주세요.", vbExclamation
        GoTo Cleanup
    End If
    
    Dim items() As ProductionItem
    Dim itemCount As Long
    If Not ReadRawData(items, itemCount) Then
        MsgBox "계획 시트에서 데이터를 읽을 수 없습니다.", vbExclamation
        GoTo Cleanup
    End If
    
    If itemCount = 0 Then
        MsgBox "처리할 데이터가 없습니다.", vbInformation
        GoTo Cleanup
    End If
    
    Dim results() As AllocationResult
    AllocateProduction items, itemCount, results
    
    CreateCalendarView results, itemCount
    
    MsgBox "완료! (" & itemCount & "건, " & Format(Timer - startTime, "0.0") & "초)", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "오류: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub 초기설정()
    Application.ScreenUpdating = False
    CreateVendorSheet
    CreateMappingSheet
    CreateOrClearSheet SHEET_OUTPUT
    Application.ScreenUpdating = True
    MsgBox "초기 설정 완료!", vbInformation
End Sub

'===============================================================================
' 자동 배정 로직 (CAPA 수정)
'===============================================================================
Private Sub AllocateProduction(ByRef items() As ProductionItem, ByVal itemCount As Long, ByRef results() As AllocationResult)
    ReDim results(1 To itemCount)
    
    Dim targetMonth As Date: targetMonth = Date
    Dim monthEnd As Date: monthEnd = DateSerial(Year(targetMonth), Month(targetMonth) + 1, 0)
    
    Dim i As Long
    For i = 1 To itemCount
        results(i).Item = items(i)
        
        ' 이미 외주처가 지정되어 있는 경우
        If Len(items(i).AssignedVendor) > 0 Then
            results(i).VendorName = items(i).AssignedVendor
            results(i).DaysNeeded = CalculateDaysNeeded(items(i).Quantity, GetVendorCapacity(items(i).AssignedVendor))
            
            ' 해당 외주처에서 빈 라인/날짜 찾기
            Dim foundSlot As Boolean
            foundSlot = FindAvailableSlot(results(i), items(i).TransferDate, Year(targetMonth), monthEnd)
            
            If Not foundSlot Then
                results(i).VendorName = "배정불가"
                results(i).LineNumber = 0
                results(i).DaysNeeded = 0
            End If
        Else
            ' 자동 배정
            Dim vendorIdx As Integer
            vendorIdx = SelectVendor(items(i))
            
            If vendorIdx > 0 Then
                results(i).VendorName = Vendors(vendorIdx).Name
                results(i).DaysNeeded = CalculateDaysNeeded(items(i).Quantity, Vendors(vendorIdx).DailyCapacity)
                
                foundSlot = FindAvailableSlot(results(i), items(i).TransferDate, Year(targetMonth), monthEnd)
                
                If Not foundSlot Then
                    results(i).VendorName = "배정불가"
                    results(i).LineNumber = 0
                    results(i).DaysNeeded = 0
                End If
            Else
                results(i).VendorName = "배정불가"
                results(i).LineNumber = 0
                results(i).StartDate = Date
                results(i).DaysNeeded = 0
            End If
        End If
    Next i
End Sub

'===============================================================================
' 빈 라인/날짜 슬롯 찾기 (핵심 로직)
'===============================================================================
Private Function FindAvailableSlot(ByRef result As AllocationResult, ByVal transferDateStr As String, _
                                   ByVal targetYear As Integer, ByVal monthEnd As Date) As Boolean
    FindAvailableSlot = False
    
    Dim vendorName As String: vendorName = result.VendorName
    Dim daysNeeded As Integer: daysNeeded = result.DaysNeeded
    
    ' 외주처 정보 가져오기
    Dim lineCount As Integer: lineCount = GetVendorLineCount(vendorName)
    If lineCount = 0 Then lineCount = 1
    
    ' 생산 시작 가능일 (이동일 + 1 근무일)
    Dim minStartDate As Date
    minStartDate = GetProductionStartDate(transferDateStr, targetYear)
    
    ' 각 라인별로 빈 슬롯 찾기
    Dim line As Integer
    Dim searchDate As Date
    
    ' 시작일부터 월말까지 검색
    searchDate = minStartDate
    
    Do While searchDate <= monthEnd
        ' 주말이면 스킵
        If Weekday(searchDate) = 1 Or Weekday(searchDate) = 7 Then
            searchDate = searchDate + 1
            GoTo ContinueSearch
        End If
        
        ' 각 라인에서 연속으로 빈 날짜 찾기
        For line = 1 To lineCount
            If CanPlaceOnLine(vendorName, line, searchDate, daysNeeded, monthEnd) Then
                ' 슬롯 찾음! 예약하기
                ReserveLine vendorName, line, searchDate, daysNeeded
                
                result.LineNumber = line
                result.StartDate = searchDate
                FindAvailableSlot = True
                Exit Function
            End If
        Next line
        
        searchDate = searchDate + 1
ContinueSearch:
    Loop
    
    ' 빈 슬롯을 찾지 못함
    FindAvailableSlot = False
End Function

'===============================================================================
' 특정 라인에 연속으로 배치 가능한지 확인
'===============================================================================
Private Function CanPlaceOnLine(ByVal vendorName As String, ByVal line As Integer, _
                                ByVal startDate As Date, ByVal daysNeeded As Integer, _
                                ByVal monthEnd As Date) As Boolean
    CanPlaceOnLine = True
    
    Dim checkDate As Date: checkDate = startDate
    Dim placedDays As Integer: placedDays = 0
    
    Do While placedDays < daysNeeded
        ' 월말 초과 체크
        If checkDate > monthEnd Then
            CanPlaceOnLine = False
            Exit Function
        End If
        
        ' 주말 스킵
        If Weekday(checkDate) = 1 Or Weekday(checkDate) = 7 Then
            checkDate = checkDate + 1
            GoTo ContinueCheck
        End If
        
        ' 이미 점유되어 있는지 확인
        Dim scheduleKey As String
        scheduleKey = vendorName & "_" & line & "_" & Format(checkDate, "yyyy-mm-dd")
        
        If LineSchedule.Exists(scheduleKey) Then
            CanPlaceOnLine = False
            Exit Function
        End If
        
        placedDays = placedDays + 1
        checkDate = checkDate + 1
ContinueCheck:
    Loop
    
    CanPlaceOnLine = True
End Function

'===============================================================================
' 라인 예약 (스케줄에 등록)
'===============================================================================
Private Sub ReserveLine(ByVal vendorName As String, ByVal line As Integer, _
                        ByVal startDate As Date, ByVal daysNeeded As Integer)
    Dim checkDate As Date: checkDate = startDate
    Dim placedDays As Integer: placedDays = 0
    
    Do While placedDays < daysNeeded
        ' 주말 스킵
        If Weekday(checkDate) = 1 Or Weekday(checkDate) = 7 Then
            checkDate = checkDate + 1
            GoTo ContinueReserve
        End If
        
        Dim scheduleKey As String
        scheduleKey = vendorName & "_" & line & "_" & Format(checkDate, "yyyy-mm-dd")
        LineSchedule(scheduleKey) = True
        
        placedDays = placedDays + 1
        checkDate = checkDate + 1
ContinueReserve:
    Loop
End Sub

'===============================================================================
' 외주처 선택
'===============================================================================
Private Function SelectVendor(ByRef Item As ProductionItem) As Integer
    SelectVendor = 0
    
    ' 1. 고객사 매칭 확인
    If Len(Item.ClientCode) > 0 Then
        If ClientMappings.Exists(Item.ClientCode) Then
            Dim mappedVendorId As String
            mappedVendorId = ClientMappings(Item.ClientCode)
            
            Dim j As Integer
            For j = 1 To VendorCount
                If Vendors(j).ID = mappedVendorId Then
                    If CanHandleProcess(Vendors(j), Item.SpecialProcess) Then
                        SelectVendor = j
                        Exit Function
                    End If
                End If
            Next j
        End If
    End If
    
    ' 2. 우선순위 기반 선택
    Dim bestVendor As Integer: bestVendor = 0
    Dim bestPriority As Integer: bestPriority = 999
    
    For j = 1 To VendorCount
        If CanHandleProcess(Vendors(j), Item.SpecialProcess) Then
            If Vendors(j).Priority < bestPriority Then
                bestPriority = Vendors(j).Priority
                bestVendor = j
            ElseIf Vendors(j).Priority = bestPriority Then
                If bestVendor > 0 Then
                    If Vendors(j).MonthlyTarget > Vendors(bestVendor).MonthlyTarget Then
                        bestVendor = j
                    End If
                End If
            End If
        End If
    Next j
    
    SelectVendor = bestVendor
End Function

Private Function CanHandleProcess(ByRef vendor As VendorInfo, ByVal ProcessType As String) As Boolean
    If ProcessType = "normal" Or Len(ProcessType) = 0 Then
        CanHandleProcess = True
        Exit Function
    End If
    CanHandleProcess = InStr(1, vendor.Capabilities, ProcessType, vbTextCompare) > 0
End Function

'===============================================================================
' 달력 생성
'===============================================================================
Private Sub CreateCalendarView(ByRef results() As AllocationResult, ByVal itemCount As Long)
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_OUTPUT)
    
    Dim targetMonth As Date: targetMonth = Date
    Dim monthStart As Date: monthStart = DateSerial(Year(targetMonth), Month(targetMonth), 1)
    Dim monthEnd As Date: monthEnd = DateSerial(Year(targetMonth), Month(targetMonth) + 1, 0)
    Dim daysInMonth As Integer: daysInMonth = Day(monthEnd)
    
    ws.Cells.Font.Name = "맑은 고딕"
    ws.Cells.Font.Size = 9
    
    ' 타이틀
    ws.Range("A1").Value = Year(targetMonth) & "년 " & Month(targetMonth) & "월 생산계획표"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(30, 41, 59)
    
    ws.Range("A2").Value = "Updated: " & Format(Now, "yyyy-mm-dd hh:mm")
    ws.Range("A2").Font.Color = RGB(148, 163, 184)
    
    ' 헤더
    Dim headerRow As Integer: headerRow = 4
    ws.Cells(headerRow, 1).Value = "외주처 / 라인"
    ws.Columns(1).ColumnWidth = 14
    
    Dim d As Integer
    For d = 1 To daysInMonth
        Dim currentDate As Date
        currentDate = DateSerial(Year(targetMonth), Month(targetMonth), d)
        Dim dow As Integer: dow = Weekday(currentDate)
        Dim col As Integer: col = d + 1
        
        ws.Cells(headerRow, col).Value = d & vbLf & GetDayName(dow)
        ws.Cells(headerRow, col).WrapText = True
        ws.Cells(headerRow, col).HorizontalAlignment = xlCenter
        ws.Cells(headerRow, col).VerticalAlignment = xlCenter
        
        If dow = 1 Or dow = 7 Then
            ws.Cells(headerRow, col).Interior.Color = RGB(254, 226, 226)
            ws.Cells(headerRow, col).Font.Color = RGB(185, 28, 28)
        ElseIf currentDate = Date Then
            ws.Cells(headerRow, col).Interior.Color = RGB(59, 130, 246)
            ws.Cells(headerRow, col).Font.Color = RGB(255, 255, 255)
            ws.Cells(headerRow, col).Font.Bold = True
        Else
            ws.Cells(headerRow, col).Interior.Color = RGB(241, 245, 249)
            ws.Cells(headerRow, col).Font.Color = RGB(71, 85, 105)
        End If
        
        ws.Columns(col).ColumnWidth = 14
    Next d
    
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, daysInMonth + 1)).Font.Bold = True
    ws.Rows(headerRow).RowHeight = 36
    ws.Cells(headerRow, 1).Interior.Color = RGB(241, 245, 249)
    
    ' 외주처별 행
    Dim currentRow As Integer: currentRow = headerRow + 1
    Dim v As Integer
    
    For v = 1 To VendorCount
        Dim vendorName As String: vendorName = Vendors(v).Name
        
        ' 배정된 품목 있는지 확인
        Dim hasItems As Boolean: hasItems = False
        Dim i As Long
        For i = 1 To itemCount
            If results(i).VendorName = vendorName Then
                hasItems = True
                Exit For
            End If
        Next i
        
        If Not hasItems Then GoTo NextVendor
        
        ' 외주처 헤더
        ws.Cells(currentRow, 1).Value = vendorName
        ws.Cells(currentRow, 1).Font.Bold = True
        ws.Cells(currentRow, 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(currentRow, 1).Interior.Color = GetVendorColor(vendorName)
        
        For d = 1 To daysInMonth
            ws.Cells(currentRow, d + 1).Interior.Color = GetVendorPaleColor(vendorName)
        Next d
        ws.Rows(currentRow).RowHeight = 22
        currentRow = currentRow + 1
        
        ' 라인별 행
        Dim line As Integer
        For line = 1 To Vendors(v).LineCount
            ws.Cells(currentRow, 1).Value = "   Line " & line
            ws.Cells(currentRow, 1).Font.Color = RGB(100, 116, 139)
            ws.Cells(currentRow, 1).Interior.Color = RGB(248, 250, 252)
            
            ' 날짜 셀 초기화
            For d = 1 To daysInMonth
                currentDate = DateSerial(Year(targetMonth), Month(targetMonth), d)
                dow = Weekday(currentDate)
                col = d + 1
                
                If dow = 1 Or dow = 7 Then
                    ws.Cells(currentRow, col).Interior.Color = RGB(254, 242, 242)
                Else
                    ws.Cells(currentRow, col).Interior.Color = RGB(255, 255, 255)
                End If
            Next d
            
            ' 해당 라인에 배정된 품목 표시
            For i = 1 To itemCount
                If results(i).VendorName = vendorName And results(i).LineNumber = line Then
                    PlaceItemOnCalendar ws, results(i), currentRow, targetMonth, daysInMonth
                End If
            Next i
            
            ws.Rows(currentRow).RowHeight = 50
            currentRow = currentRow + 1
        Next line
        
NextVendor:
    Next v
    
    ' 배정불가
    Dim hasUnassigned As Boolean: hasUnassigned = False
    For i = 1 To itemCount
        If results(i).VendorName = "배정불가" Then
            hasUnassigned = True
            Exit For
        End If
    Next i
    
    If hasUnassigned Then
        ws.Cells(currentRow, 1).Value = "배정불가"
        ws.Cells(currentRow, 1).Font.Bold = True
        ws.Cells(currentRow, 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(currentRow, 1).Interior.Color = RGB(148, 163, 184)
        ws.Rows(currentRow).RowHeight = 22
        currentRow = currentRow + 1
        
        ' 배정불가 품목 나열
        For i = 1 To itemCount
            If results(i).VendorName = "배정불가" Then
                ws.Cells(currentRow, 1).Value = "   " & Left(results(i).Item.ProductName, 15)
                ws.Cells(currentRow, 2).Value = Format(results(i).Item.Quantity, "#,##0")
                ws.Cells(currentRow, 1).Font.Color = RGB(185, 28, 28)
                ws.Cells(currentRow, 2).Font.Color = RGB(185, 28, 28)
                ws.Rows(currentRow).RowHeight = 20
                currentRow = currentRow + 1
            End If
        Next i
    End If
    
    ' 테두리
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(currentRow - 1, daysInMonth + 1))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(226, 232, 240)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(226, 232, 240)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(226, 232, 240)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = RGB(241, 245, 249)
    End With
    
    ' 틀 고정
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("B5").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

'===============================================================================
' 품목을 달력에 배치 (한 셀에 한 제품만!)
'===============================================================================
Private Sub PlaceItemOnCalendar(ByRef ws As Worksheet, ByRef result As AllocationResult, _
                                ByVal rowNum As Integer, ByVal targetMonth As Date, ByVal daysInMonth As Integer)
    
    If result.DaysNeeded = 0 Then Exit Sub
    If Month(result.StartDate) <> Month(targetMonth) Then Exit Sub
    
    Dim dailyCapacity As Long: dailyCapacity = GetVendorDailyCapacity(result.VendorName)
    Dim remainingQty As Long: remainingQty = result.Item.Quantity
    Dim currentDate As Date: currentDate = result.StartDate
    Dim dayNum As Integer: dayNum = 1
    
    Do While remainingQty > 0 And Day(currentDate) <= daysInMonth And Month(currentDate) = Month(targetMonth)
        ' 주말 스킵
        If Weekday(currentDate) = 1 Or Weekday(currentDate) = 7 Then
            currentDate = currentDate + 1
            GoTo ContinuePlace
        End If
        
        Dim col As Integer: col = Day(currentDate) + 1
        Dim dailyQty As Long: dailyQty = remainingQty
        If dailyQty > dailyCapacity Then dailyQty = dailyCapacity
        
        ' 셀 내용
        Dim cellText As String
        Dim prodName As String: prodName = result.Item.ProductName
        If Len(prodName) > 12 Then prodName = Left(prodName, 11) & ".."
        
        cellText = prodName & vbLf & Format(dailyQty, "#,##0")
        If result.DaysNeeded > 1 Then
            cellText = cellText & vbLf & "(" & dayNum & "/" & result.DaysNeeded & ")"
        End If
        
        ' 셀에 표시 (이미 값이 있으면 안 됨 - 로직상 없어야 함)
        ws.Cells(rowNum, col).Value = cellText
        
        With ws.Cells(rowNum, col)
            .Interior.Color = GetVendorLightColor(result.VendorName)
            .Font.Size = 9
            .Font.Color = RGB(30, 41, 59)
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .WrapText = True
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeLeft).Color = GetVendorColor(result.VendorName)
        End With
        
        remainingQty = remainingQty - dailyQty
        dayNum = dayNum + 1
        currentDate = currentDate + 1
ContinuePlace:
    Loop
End Sub

'===============================================================================
' 설정 시트
'===============================================================================
Private Sub CreateVendorSheet()
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_VENDORS)
    
    ws.Range("A1:G1").Value = Array("외주처ID", "외주처명", "라인수", "라인당일일생산량", "가능공정", "월간목표", "우선순위")
    
    Dim data As Variant
    data = Array( _
        Array("withmom", "위드맘", 8, 15000, "normal,shrink,highFrequency", 1600000, 1), _
        Array("linear", "리니어", 5, 13000, "normal,shrink,mixing,highFrequency", 1600000, 1), _
        Array("gram", "그램", 3, 10000, "normal", 500000, 1), _
        Array("isis", "이시스", 2, 7000, "normal,mixing", 0, 2), _
        Array("elluo", "엘루오", 1, 7000, "normal", 0, 2), _
        Array("kcostech", "케이코스텍", 2, 7000, "normal,mixing", 0, 2), _
        Array("dami", "다미", 2, 7000, "normal,shrink", 0, 2) _
    )
    
    Dim i As Integer
    For i = 0 To UBound(data)
        ws.Range("A" & (i + 2) & ":G" & (i + 2)).Value = data(i)
    Next i
    
    FormatSettingSheet ws, 7, UBound(data) + 2
End Sub

Private Sub CreateMappingSheet()
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_MAPPING)
    
    ws.Range("A1:C1").Value = Array("고객사코드", "외주처ID", "우선순위")
    
    Dim data As Variant
    data = Array( _
        Array("CLO", "gram", 1), _
        Array("ERK", "gram", 1), _
        Array("DPD", "withmom", 1), _
        Array("GDI", "withmom", 1), _
        Array("MDH", "linear", 1), _
        Array("APS", "linear", 1), _
        Array("PUR", "kcostech", 1) _
    )
    
    Dim i As Integer
    For i = 0 To UBound(data)
        ws.Range("A" & (i + 2) & ":C" & (i + 2)).Value = data(i)
    Next i
    
    FormatSettingSheet ws, 3, UBound(data) + 2
End Sub

Private Function LoadSettings() As Boolean
    On Error GoTo ErrorHandler
    
    Dim wsVendor As Worksheet
    Set wsVendor = ThisWorkbook.Sheets(SHEET_VENDORS)
    
    Dim lastRow As Long
    lastRow = wsVendor.Cells(wsVendor.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        LoadSettings = False
        Exit Function
    End If
    
    VendorCount = lastRow - 1
    ReDim Vendors(1 To VendorCount)
    
    Dim i As Long
    For i = 2 To lastRow
        Vendors(i - 1).ID = CStr(wsVendor.Cells(i, 1).Value)
        Vendors(i - 1).Name = CStr(wsVendor.Cells(i, 2).Value)
        Vendors(i - 1).LineCount = CInt(wsVendor.Cells(i, 3).Value)
        Vendors(i - 1).DailyCapacity = CLng(wsVendor.Cells(i, 4).Value)
        Vendors(i - 1).Capabilities = CStr(wsVendor.Cells(i, 5).Value)
        Vendors(i - 1).MonthlyTarget = CLng(wsVendor.Cells(i, 6).Value)
        Vendors(i - 1).Priority = CInt(wsVendor.Cells(i, 7).Value)
    Next i
    
    Dim wsMapping As Worksheet
    Set wsMapping = ThisWorkbook.Sheets(SHEET_MAPPING)
    
    Set ClientMappings = CreateObject("Scripting.Dictionary")
    
    lastRow = wsMapping.Cells(wsMapping.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        Dim clientCode As String, vendorId As String
        clientCode = CStr(wsMapping.Cells(i, 1).Value)
        vendorId = CStr(wsMapping.Cells(i, 2).Value)
        If Len(clientCode) > 0 And Len(vendorId) > 0 Then
            ClientMappings(clientCode) = vendorId
        End If
    Next i
    
    LoadSettings = True
    Exit Function
    
ErrorHandler:
    LoadSettings = False
End Function

Private Function ReadRawData(ByRef items() As ProductionItem, ByRef itemCount As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_RAW)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_PRODUCT_CODE).End(xlUp).Row
    
    If lastRow < 2 Then
        itemCount = 0
        ReadRawData = True
        Exit Function
    End If
    
    itemCount = 0
    Dim i As Long
    For i = 2 To lastRow
        Dim productCode As String, qty As Long, transferDate As String
        productCode = Trim(CStr(ws.Cells(i, COL_PRODUCT_CODE).Value))
        qty = ParseQuantity(ws.Cells(i, COL_QUANTITY).Value)
        transferDate = Trim(CStr(ws.Cells(i, COL_TRANSFER_DATE).Value))
        
        If Len(productCode) = 0 Or qty = 0 Then GoTo NextRow
        If transferDate = "미정" Or Len(transferDate) = 0 Then GoTo NextRow
        itemCount = itemCount + 1
NextRow:
    Next i
    
    If itemCount = 0 Then
        ReadRawData = True
        Exit Function
    End If
    
    ReDim items(1 To itemCount)
    Dim idx As Long: idx = 0
    
    For i = 2 To lastRow
        productCode = Trim(CStr(ws.Cells(i, COL_PRODUCT_CODE).Value))
        qty = ParseQuantity(ws.Cells(i, COL_QUANTITY).Value)
        transferDate = Trim(CStr(ws.Cells(i, COL_TRANSFER_DATE).Value))
        
        If Len(productCode) = 0 Or qty = 0 Then GoTo NextRow2
        If transferDate = "미정" Or Len(transferDate) = 0 Then GoTo NextRow2
        
        idx = idx + 1
        With items(idx)
            .RowNum = i
            .TransferDate = transferDate
            .Status = Trim(CStr(ws.Cells(i, COL_STATUS).Value))
            .SpecialProcess = ParseSpecialProcess(CStr(ws.Cells(i, COL_SPECIAL_PROCESS).Value))
            .AssignedVendor = Trim(CStr(ws.Cells(i, COL_VENDOR).Value))
            .Manager = Trim(CStr(ws.Cells(i, COL_MANAGER).Value))
            .ProductCode = productCode
            .ProductName = Trim(CStr(ws.Cells(i, COL_PRODUCT_NAME).Value))
            .ProcessType = Trim(CStr(ws.Cells(i, COL_PROCESS_TYPE).Value))
            .Quantity = qty
            .DeliveryDate = Trim(CStr(ws.Cells(i, COL_DELIVERY_DATE).Value))
            .ClientCode = ExtractClientCode(productCode)
        End With
NextRow2:
    Next i
    
    ReadRawData = True
    Exit Function
    
ErrorHandler:
    ReadRawData = False
End Function

'===============================================================================
' 유틸리티
'===============================================================================
Private Function CreateOrClearSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
        ActiveWindow.FreezePanes = False
    End If
    Set CreateOrClearSheet = ws
End Function

Private Sub FormatSettingSheet(ByRef ws As Worksheet, ByVal colCount As Integer, ByVal rowCount As Integer)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(51, 65, 85)
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(226, 232, 240)
    End With
    ws.Columns("A:" & Chr(64 + colCount)).AutoFit
End Sub

Private Function ParseQuantity(ByVal Value As Variant) As Long
    On Error Resume Next
    If IsEmpty(Value) Or Value = "" Then
        ParseQuantity = 0
    ElseIf IsNumeric(Value) Then
        ParseQuantity = CLng(Value)
    Else
        ParseQuantity = CLng(Replace(CStr(Value), ",", ""))
    End If
    On Error GoTo 0
End Function

Private Function ParseSpecialProcess(ByVal Value As String) As String
    Select Case Trim(Value)
        Case "수축": ParseSpecialProcess = "shrink"
        Case "교반": ParseSpecialProcess = "mixing"
        Case "고주파": ParseSpecialProcess = "highFrequency"
        Case Else: ParseSpecialProcess = "normal"
    End Select
End Function

Private Function ExtractClientCode(ByVal productCode As String) As String
    If Len(productCode) < 4 Then
        ExtractClientCode = ""
        Exit Function
    End If
    Dim startPos As Integer: startPos = 1
    If IsNumeric(Left(productCode, 1)) Then startPos = 2
    Dim code As String: code = Mid(productCode, startPos, 3)
    Dim i As Integer
    For i = 1 To Len(code)
        If Not (Mid(code, i, 1) Like "[A-Z]") Then
            ExtractClientCode = ""
            Exit Function
        End If
    Next i
    ExtractClientCode = code
End Function

Private Function GetProductionStartDate(ByVal transferDateStr As String, ByVal targetYear As Integer) As Date
    Dim transferDate As Date
    On Error Resume Next
    If InStr(transferDateStr, "/") > 0 Then
        Dim parts() As String: parts = Split(transferDateStr, "/")
        If UBound(parts) >= 1 Then
            transferDate = DateSerial(targetYear, CInt(parts(0)), CInt(parts(1)))
        End If
    ElseIf InStr(transferDateStr, "월") > 0 Then
        Dim monthPart As String, dayPart As String
        monthPart = Left(transferDateStr, InStr(transferDateStr, "월") - 1)
        dayPart = Mid(transferDateStr, InStr(transferDateStr, "월") + 1)
        dayPart = Replace(dayPart, "일", "")
        transferDate = DateSerial(targetYear, CInt(Trim(monthPart)), CInt(Trim(dayPart)))
    Else
        transferDate = CDate(transferDateStr)
    End If
    On Error GoTo 0
    If transferDate > 0 Then
        GetProductionStartDate = GetNextWorkingDay(transferDate, 1)
    Else
        GetProductionStartDate = GetNextWorkingDay(Date, 1)
    End If
End Function

Private Function GetNextWorkingDay(ByVal StartDate As Date, ByVal daysToAdd As Integer) As Date
    Dim result As Date: result = StartDate
    Dim addedDays As Integer: addedDays = 0
    Do While addedDays < daysToAdd
        result = result + 1
        If Weekday(result) <> 1 And Weekday(result) <> 7 Then addedDays = addedDays + 1
    Loop
    GetNextWorkingDay = result
End Function

Private Function CalculateDaysNeeded(ByVal Quantity As Long, ByVal DailyCapacity As Long) As Integer
    If DailyCapacity <= 0 Then DailyCapacity = 10000
    CalculateDaysNeeded = Application.WorksheetFunction.Ceiling(Quantity / DailyCapacity, 1)
End Function

Private Function GetVendorCapacity(ByVal vendorName As String) As Long
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).Name = vendorName Then
            GetVendorCapacity = Vendors(i).DailyCapacity
            Exit Function
        End If
    Next i
    GetVendorCapacity = 10000
End Function

Private Function GetVendorDailyCapacity(ByVal vendorName As String) As Long
    GetVendorDailyCapacity = GetVendorCapacity(vendorName)
End Function

Private Function GetVendorLineCount(ByVal vendorName As String) As Integer
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).Name = vendorName Then
            GetVendorLineCount = Vendors(i).LineCount
            Exit Function
        End If
    Next i
    GetVendorLineCount = 1
End Function

Private Function GetDayName(ByVal dayOfWeek As Integer) As String
    Select Case dayOfWeek
        Case 1: GetDayName = "일"
        Case 2: GetDayName = "월"
        Case 3: GetDayName = "화"
        Case 4: GetDayName = "수"
        Case 5: GetDayName = "목"
        Case 6: GetDayName = "금"
        Case 7: GetDayName = "토"
    End Select
End Function

Private Function GetVendorColor(ByVal vendorName As String) As Long
    Select Case vendorName
        Case "위드맘": GetVendorColor = RGB(37, 99, 235)
        Case "리니어": GetVendorColor = RGB(22, 163, 74)
        Case "그램": GetVendorColor = RGB(147, 51, 234)
        Case "이시스": GetVendorColor = RGB(234, 88, 12)
        Case "엘루오": GetVendorColor = RGB(219, 39, 119)
        Case "케이코스텍": GetVendorColor = RGB(8, 145, 178)
        Case "다미": GetVendorColor = RGB(202, 138, 4)
        Case "배정불가": GetVendorColor = RGB(100, 116, 139)
        Case Else: GetVendorColor = RGB(148, 163, 184)
    End Select
End Function

Private Function GetVendorLightColor(ByVal vendorName As String) As Long
    Select Case vendorName
        Case "위드맘": GetVendorLightColor = RGB(239, 246, 255)
        Case "리니어": GetVendorLightColor = RGB(240, 253, 244)
        Case "그램": GetVendorLightColor = RGB(250, 245, 255)
        Case "이시스": GetVendorLightColor = RGB(255, 247, 237)
        Case "엘루오": GetVendorLightColor = RGB(253, 242, 248)
        Case "케이코스텍": GetVendorLightColor = RGB(236, 254, 255)
        Case "다미": GetVendorLightColor = RGB(254, 252, 232)
        Case "배정불가": GetVendorLightColor = RGB(248, 250, 252)
        Case Else: GetVendorLightColor = RGB(248, 250, 252)
    End Select
End Function

Private Function GetVendorPaleColor(ByVal vendorName As String) As Long
    Select Case vendorName
        Case "위드맘": GetVendorPaleColor = RGB(248, 250, 255)
        Case "리니어": GetVendorPaleColor = RGB(248, 255, 250)
        Case "그램": GetVendorPaleColor = RGB(253, 250, 255)
        Case "이시스": GetVendorPaleColor = RGB(255, 252, 248)
        Case "엘루오": GetVendorPaleColor = RGB(255, 250, 253)
        Case "케이코스텍": GetVendorPaleColor = RGB(248, 255, 255)
        Case "다미": GetVendorPaleColor = RGB(255, 254, 248)
        Case "배정불가": GetVendorPaleColor = RGB(250, 251, 252)
        Case Else: GetVendorPaleColor = RGB(250, 251, 252)
    End Select
End Function
