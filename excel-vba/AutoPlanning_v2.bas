'===============================================================================
' 외주처 생산계획 자동화 시스템 - VBA 매크로 v2.0
' 달력식 간트차트 뷰 추가
'
' 사용법:
' 1. 이 코드를 엑셀의 VBA 편집기(Alt+F11)에서 새 모듈에 붙여넣기
' 2. "초기설정" 매크로를 한 번 실행하여 설정 시트 생성
' 3. "계획" 시트에 Raw 데이터 입력
' 4. "자동배정실행" 매크로 실행
'===============================================================================

Option Explicit

'===============================================================================
' 상수 정의
'===============================================================================
Public Const SHEET_RAW As String = "계획"
Public Const SHEET_VENDORS As String = "설정_외주처"
Public Const SHEET_MAPPING As String = "설정_고객매칭"
Public Const SHEET_OUTPUT As String = "auto-planning"

' 엑셀 컬럼 인덱스 (1부터 시작)
Public Const COL_TRANSFER_DATE As Integer = 1   ' A열 - 이동일
Public Const COL_STATUS As Integer = 2          ' B열 - 상태
Public Const COL_SPECIAL_PROCESS As Integer = 3 ' C열 - 특수공정
Public Const COL_VENDOR As Integer = 4          ' D열 - 외주처
Public Const COL_MANAGER As Integer = 5         ' E열 - 담당자
Public Const COL_PRODUCT_CODE As Integer = 6    ' F열 - 제품코드
Public Const COL_PRODUCT_NAME As Integer = 7    ' G열 - 제품명
Public Const COL_PROCESS_TYPE As Integer = 8    ' H열 - 공정
Public Const COL_QUANTITY As Integer = 9        ' I열 - 수량
Public Const COL_MFG_DATE As Integer = 10       ' J열 - 제조일
Public Const COL_MATERIAL_DATE As Integer = 11  ' K열 - 자재입고
Public Const COL_DELIVERY_DATE As Integer = 12  ' L열 - 납기일
Public Const COL_URGENCY As Integer = 13        ' M열 - 긴급
Public Const COL_NIGHT_SHIFT As Integer = 14    ' N열 - 야상
Public Const COL_REMARKS As Integer = 15        ' O열 - 비고

'===============================================================================
' 타입 정의
'===============================================================================
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

' 스케줄 정보 (달력에 표시용)
Public Type ScheduleEntry
    ProductCode As String
    ProductName As String
    Quantity As Long
    DailyQty As Long
    DayNum As Integer      ' 몇 번째 날인지
    TotalDays As Integer   ' 총 며칠 걸리는지
    DeliveryDate As String
End Type

' 전역 변수
Public Vendors() As VendorInfo
Public VendorCount As Integer
Public ClientMappings As Object

' 스케줄 데이터 (외주처명_라인_날짜 -> ScheduleEntry)
Public ScheduleData As Object

'===============================================================================
' 메인 진입점
'===============================================================================
Public Sub 자동배정실행()
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' 스케줄 데이터 초기화
    Set ScheduleData = CreateObject("Scripting.Dictionary")
    
    ' 1. 설정 로드
    If Not LoadSettings() Then
        MsgBox "설정을 로드할 수 없습니다. '초기설정' 매크로를 먼저 실행해주세요.", vbExclamation
        GoTo Cleanup
    End If
    
    ' 2. Raw 데이터 읽기
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
    
    ' 3. 자동 배정 실행
    Dim results() As AllocationResult
    AllocateProduction items, itemCount, results
    
    ' 4. 달력식 생산계획표 생성
    CreateCalendarView results, itemCount
    
    MsgBox "자동 배정 완료!" & vbCrLf & _
           "처리 건수: " & itemCount & "건" & vbCrLf & _
           "소요 시간: " & Format(Timer - startTime, "0.00") & "초", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

'===============================================================================
' 초기 설정
'===============================================================================
Public Sub 초기설정()
    Application.ScreenUpdating = False
    
    CreateVendorSheet
    CreateMappingSheet
    CreateOrClearSheet SHEET_OUTPUT
    
    Application.ScreenUpdating = True
    
    MsgBox "초기 설정 완료!" & vbCrLf & _
           "- " & SHEET_VENDORS & " 시트: 외주처 정보 설정" & vbCrLf & _
           "- " & SHEET_MAPPING & " 시트: 고객사-외주처 매칭 설정" & vbCrLf & vbCrLf & _
           "설정을 확인/수정한 후 '자동배정실행' 매크로를 실행하세요.", vbInformation
End Sub

'===============================================================================
' 달력식 생산계획표 생성 (핵심 기능)
'===============================================================================
Private Sub CreateCalendarView(ByRef results() As AllocationResult, ByVal itemCount As Long)
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_OUTPUT)
    
    ' 현재 월 정보
    Dim targetMonth As Date
    targetMonth = Date
    Dim monthStart As Date, monthEnd As Date
    monthStart = DateSerial(Year(targetMonth), Month(targetMonth), 1)
    monthEnd = DateSerial(Year(targetMonth), Month(targetMonth) + 1, 0)
    Dim daysInMonth As Integer
    daysInMonth = Day(monthEnd)
    
    ' =========================================================================
    ' 1. 대시보드 영역 (1~13행)
    ' =========================================================================
    CreateDashboardSection ws, results, itemCount
    
    ' =========================================================================
    ' 2. 달력식 생산계획표 (15행부터)
    ' =========================================================================
    Dim calendarStartRow As Integer
    calendarStartRow = 15
    
    ' 제목
    ws.Cells(calendarStartRow, 1).Value = "생산계획표 - " & Year(targetMonth) & "년 " & Month(targetMonth) & "월"
    ws.Cells(calendarStartRow, 1).Font.Size = 14
    ws.Cells(calendarStartRow, 1).Font.Bold = True
    
    calendarStartRow = calendarStartRow + 2
    
    ' =========================================================================
    ' 날짜 헤더 생성
    ' =========================================================================
    Dim headerRow As Integer
    headerRow = calendarStartRow
    
    ' 첫 번째 열: 외주처/라인
    ws.Cells(headerRow, 1).Value = "외주처/라인"
    ws.Cells(headerRow, 1).ColumnWidth = 15
    
    ' 날짜 헤더 (1일 ~ 말일)
    Dim d As Integer
    For d = 1 To daysInMonth
        Dim currentDate As Date
        currentDate = DateSerial(Year(targetMonth), Month(targetMonth), d)
        
        Dim dayOfWeek As Integer
        dayOfWeek = Weekday(currentDate)
        
        ' 날짜 표시
        ws.Cells(headerRow, d + 1).Value = d
        ws.Cells(headerRow + 1, d + 1).Value = GetDayName(dayOfWeek)
        
        ' 요일별 색상
        If dayOfWeek = 1 Then  ' 일요일
            ws.Cells(headerRow, d + 1).Interior.Color = RGB(254, 226, 226)
            ws.Cells(headerRow + 1, d + 1).Interior.Color = RGB(254, 226, 226)
            ws.Cells(headerRow + 1, d + 1).Font.Color = RGB(220, 38, 38)
        ElseIf dayOfWeek = 7 Then  ' 토요일
            ws.Cells(headerRow, d + 1).Interior.Color = RGB(254, 226, 226)
            ws.Cells(headerRow + 1, d + 1).Interior.Color = RGB(254, 226, 226)
            ws.Cells(headerRow + 1, d + 1).Font.Color = RGB(220, 38, 38)
        Else
            ws.Cells(headerRow, d + 1).Interior.Color = RGB(243, 244, 246)
            ws.Cells(headerRow + 1, d + 1).Interior.Color = RGB(243, 244, 246)
        End If
        
        ' 오늘 날짜 강조
        If currentDate = Date Then
            ws.Cells(headerRow, d + 1).Interior.Color = RGB(219, 234, 254)
            ws.Cells(headerRow + 1, d + 1).Interior.Color = RGB(219, 234, 254)
            ws.Cells(headerRow, d + 1).Font.Bold = True
        End If
        
        ws.Cells(headerRow, d + 1).HorizontalAlignment = xlCenter
        ws.Cells(headerRow + 1, d + 1).HorizontalAlignment = xlCenter
        ws.Columns(d + 1).ColumnWidth = 12
    Next d
    
    ' 헤더 서식
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow + 1, daysInMonth + 1))
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' =========================================================================
    ' 외주처별 행 생성
    ' =========================================================================
    Dim currentRow As Integer
    currentRow = headerRow + 2
    
    Dim v As Integer
    For v = 1 To VendorCount
        Dim vendorName As String
        vendorName = Vendors(v).Name
        
        ' 해당 외주처에 배정된 품목이 있는지 확인
        Dim hasItems As Boolean
        hasItems = False
        Dim i As Long
        For i = 1 To itemCount
            If results(i).VendorName = vendorName Then
                hasItems = True
                Exit For
            End If
        Next i
        
        If Not hasItems Then GoTo NextVendor
        
        ' 외주처 헤더 행
        ws.Cells(currentRow, 1).Value = vendorName
        ws.Cells(currentRow, 1).Font.Bold = True
        ws.Cells(currentRow, 1).Interior.Color = GetVendorColor(vendorName)
        ws.Cells(currentRow, 1).Font.Color = RGB(255, 255, 255)
        
        ' 외주처 헤더 행 날짜 셀 색상
        For d = 1 To daysInMonth
            ws.Cells(currentRow, d + 1).Interior.Color = RGB(249, 250, 251)
        Next d
        
        currentRow = currentRow + 1
        
        ' 라인별 행 생성
        Dim line As Integer
        For line = 1 To Vendors(v).LineCount
            ws.Cells(currentRow, 1).Value = "  라인 " & line
            ws.Cells(currentRow, 1).Interior.Color = RGB(255, 255, 255)
            
            ' 각 날짜 셀 초기화 (주말 표시)
            For d = 1 To daysInMonth
                currentDate = DateSerial(Year(targetMonth), Month(targetMonth), d)
                dayOfWeek = Weekday(currentDate)
                
                If dayOfWeek = 1 Or dayOfWeek = 7 Then
                    ws.Cells(currentRow, d + 1).Interior.Color = RGB(254, 242, 242)
                Else
                    ws.Cells(currentRow, d + 1).Interior.Color = RGB(255, 255, 255)
                End If
            Next d
            
            ' 해당 라인에 배정된 품목 표시
            For i = 1 To itemCount
                If results(i).VendorName = vendorName And results(i).LineNumber = line Then
                    ' 품목을 달력에 표시
                    PlaceItemOnCalendar ws, results(i), currentRow, targetMonth, daysInMonth
                End If
            Next i
            
            currentRow = currentRow + 1
        Next line
        
NextVendor:
    Next v
    
    ' 배정불가 항목 표시
    Dim hasUnassigned As Boolean
    hasUnassigned = False
    For i = 1 To itemCount
        If results(i).VendorName = "배정불가" Then
            hasUnassigned = True
            Exit For
        End If
    Next i
    
    If hasUnassigned Then
        ws.Cells(currentRow, 1).Value = "배정불가"
        ws.Cells(currentRow, 1).Font.Bold = True
        ws.Cells(currentRow, 1).Interior.Color = RGB(156, 163, 175)
        ws.Cells(currentRow, 1).Font.Color = RGB(255, 255, 255)
        currentRow = currentRow + 1
        
        ws.Cells(currentRow, 1).Value = "  미배정"
        For i = 1 To itemCount
            If results(i).VendorName = "배정불가" Then
                PlaceItemOnCalendar ws, results(i), currentRow, targetMonth, daysInMonth
            End If
        Next i
        currentRow = currentRow + 1
    End If
    
    ' =========================================================================
    ' 전체 테두리 적용
    ' =========================================================================
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(currentRow - 1, daysInMonth + 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' 행 높이 조정
    ws.Rows(headerRow & ":" & (currentRow - 1)).RowHeight = 45
    ws.Rows(headerRow).RowHeight = 20
    ws.Rows(headerRow + 1).RowHeight = 20
    
    ' 첫 번째 열 고정
    ws.Activate
    ws.Range("B" & headerRow).Select
    ActiveWindow.FreezePanes = True
    
    ' =========================================================================
    ' 3. 범례 추가
    ' =========================================================================
    currentRow = currentRow + 2
    ws.Cells(currentRow, 1).Value = "범례"
    ws.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    
    Dim legendCol As Integer
    legendCol = 1
    For v = 1 To VendorCount
        ws.Cells(currentRow, legendCol).Value = Vendors(v).Name
        ws.Cells(currentRow, legendCol).Interior.Color = GetVendorLightColor(Vendors(v).Name)
        ws.Cells(currentRow, legendCol).Borders.LineStyle = xlContinuous
        legendCol = legendCol + 1
    Next v
    
    ' 시트 맨 위로 이동
    ws.Range("A1").Select
End Sub

'===============================================================================
' 품목을 달력에 배치
'===============================================================================
Private Sub PlaceItemOnCalendar(ByRef ws As Worksheet, ByRef result As AllocationResult, _
                                ByVal rowNum As Integer, ByVal targetMonth As Date, ByVal daysInMonth As Integer)
    
    If result.DaysNeeded = 0 Then Exit Sub
    
    Dim startDate As Date
    startDate = result.StartDate
    
    ' 시작일이 해당 월에 있는지 확인
    If Month(startDate) <> Month(targetMonth) Or Year(startDate) <> Year(targetMonth) Then
        Exit Sub
    End If
    
    Dim startDay As Integer
    startDay = Day(startDate)
    
    Dim dailyCapacity As Long
    dailyCapacity = GetVendorDailyCapacity(result.VendorName)
    
    Dim remainingQty As Long
    remainingQty = result.Item.Quantity
    
    Dim dayOffset As Integer
    Dim currentDay As Integer
    Dim dayNum As Integer
    dayNum = 1
    
    currentDay = startDay
    
    Do While remainingQty > 0 And currentDay <= daysInMonth
        Dim currentDate As Date
        currentDate = DateSerial(Year(targetMonth), Month(targetMonth), currentDay)
        
        ' 주말 건너뛰기
        If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
            Dim col As Integer
            col = currentDay + 1
            
            Dim dailyQty As Long
            dailyQty = remainingQty
            If dailyQty > dailyCapacity Then dailyQty = dailyCapacity
            
            ' 셀에 품목 정보 표시
            Dim cellText As String
            cellText = Left(result.Item.ProductName, 8)
            If Len(result.Item.ProductName) > 8 Then cellText = cellText & ".."
            cellText = cellText & vbLf & Format(dailyQty, "#,##0")
            
            If result.DaysNeeded > 1 Then
                cellText = cellText & vbLf & "(" & dayNum & "/" & result.DaysNeeded & ")"
            End If
            
            ' 기존 내용이 있으면 추가
            If Len(ws.Cells(rowNum, col).Value) > 0 Then
                ws.Cells(rowNum, col).Value = ws.Cells(rowNum, col).Value & vbLf & "---" & vbLf & cellText
            Else
                ws.Cells(rowNum, col).Value = cellText
            End If
            
            ' 셀 서식
            ws.Cells(rowNum, col).Interior.Color = GetVendorLightColor(result.VendorName)
            ws.Cells(rowNum, col).Font.Size = 8
            ws.Cells(rowNum, col).VerticalAlignment = xlTop
            ws.Cells(rowNum, col).WrapText = True
            
            ' 테두리 (외주처 색상)
            With ws.Cells(rowNum, col).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .Color = GetVendorColor(result.VendorName)
            End With
            
            remainingQty = remainingQty - dailyQty
            dayNum = dayNum + 1
        End If
        
        currentDay = currentDay + 1
    Loop
End Sub

'===============================================================================
' 대시보드 섹션 생성
'===============================================================================
Private Sub CreateDashboardSection(ByRef ws As Worksheet, ByRef results() As AllocationResult, ByVal itemCount As Long)
    ' 통계 계산
    Dim totalQty As Long
    Dim vendorQty As Object
    Set vendorQty = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To itemCount
        totalQty = totalQty + results(i).Item.Quantity
        
        Dim vName As String
        vName = results(i).VendorName
        If vendorQty.Exists(vName) Then
            vendorQty(vName) = vendorQty(vName) + results(i).Item.Quantity
        Else
            vendorQty(vName) = results(i).Item.Quantity
        End If
    Next i
    
    ' 타이틀
    ws.Range("A1").Value = "외주처 생산계획 대시보드"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    
    ws.Range("A2").Value = "생성일시: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Range("A2").Font.Color = RGB(128, 128, 128)
    
    ' 통계 카드
    CreateStatCard ws, "B", 4, "총 품목 수", itemCount & " 건", RGB(59, 130, 246)
    CreateStatCard ws, "D", 4, "총 생산량", Format(totalQty, "#,##0") & " 개", RGB(34, 197, 94)
    CreateStatCard ws, "F", 4, "배정 외주처", vendorQty.Count & " 곳", RGB(168, 85, 247)
    
    Dim avgQty As Long
    If itemCount > 0 Then avgQty = totalQty / itemCount
    CreateStatCard ws, "H", 4, "품목당 평균", Format(avgQty, "#,##0") & " 개", RGB(249, 115, 22)
    
    ' 외주처별 배분 현황
    ws.Range("B9").Value = "외주처별 배분 현황"
    ws.Range("B9").Font.Bold = True
    ws.Range("B9").Font.Size = 12
    
    Dim col As Integer
    col = 2
    Dim vKey As Variant
    For Each vKey In vendorQty.Keys
        If col > 10 Then Exit For
        
        ws.Cells(10, col).Value = vKey
        ws.Cells(11, col).Value = vendorQty(vKey)
        ws.Cells(11, col).NumberFormat = "#,##0"
        
        Dim targetQty As Long
        targetQty = GetVendorMonthlyTarget(CStr(vKey))
        If targetQty > 0 Then
            ws.Cells(12, col).Value = Format(vendorQty(vKey) / targetQty, "0%")
        Else
            ws.Cells(12, col).Value = "-"
        End If
        
        ws.Cells(10, col).Interior.Color = GetVendorColor(CStr(vKey))
        ws.Cells(10, col).Font.Color = RGB(255, 255, 255)
        ws.Cells(10, col).Font.Bold = True
        
        col = col + 1
    Next vKey
    
    If col > 2 Then
        With ws.Range(ws.Cells(10, 2), ws.Cells(12, col - 1))
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
    End If
End Sub

'===============================================================================
' 설정 시트 생성 - 외주처
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
    
    ws.Range("I1").Value = "가능공정 코드:"
    ws.Range("I2").Value = "normal = 일반"
    ws.Range("I3").Value = "shrink = 수축"
    ws.Range("I4").Value = "mixing = 교반"
    ws.Range("I5").Value = "highFrequency = 고주파"
    ws.Range("I1:I5").Font.Color = RGB(100, 100, 100)
    ws.Range("I1").Font.Bold = True
End Sub

'===============================================================================
' 설정 시트 생성 - 고객매칭
'===============================================================================
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
    
    ws.Range("E1").Value = "* 제품코드에서 추출된 3자리 영문코드가 해당 외주처에 우선 배정됩니다."
    ws.Range("E1").Font.Color = RGB(100, 100, 100)
End Sub

'===============================================================================
' 설정 로드
'===============================================================================
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

'===============================================================================
' Raw 데이터 읽기
'===============================================================================
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
    
    Dim idx As Long
    idx = 0
    
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
' 자동 배정 로직
'===============================================================================
Private Sub AllocateProduction(ByRef items() As ProductionItem, ByVal itemCount As Long, ByRef results() As AllocationResult)
    ReDim results(1 To itemCount)
    
    Dim lineSchedule As Object
    Set lineSchedule = CreateObject("Scripting.Dictionary")
    
    Dim targetMonth As Date
    targetMonth = Date
    
    Dim i As Long
    For i = 1 To itemCount
        results(i).Item = items(i)
        
        If Len(items(i).AssignedVendor) > 0 Then
            results(i).VendorName = items(i).AssignedVendor
            results(i).LineNumber = 1
            results(i).StartDate = GetProductionStartDate(items(i).TransferDate, Year(targetMonth))
            results(i).DaysNeeded = CalculateDaysNeeded(items(i).Quantity, GetVendorCapacity(items(i).AssignedVendor))
        Else
            Dim vendorIdx As Integer
            vendorIdx = SelectVendor(items(i), lineSchedule, targetMonth)
            
            If vendorIdx > 0 Then
                results(i).VendorName = Vendors(vendorIdx).Name
                results(i).StartDate = GetProductionStartDate(items(i).TransferDate, Year(targetMonth))
                results(i).DaysNeeded = CalculateDaysNeeded(items(i).Quantity, Vendors(vendorIdx).DailyCapacity)
                results(i).LineNumber = AssignLine(Vendors(vendorIdx), results(i).StartDate, results(i).DaysNeeded, lineSchedule)
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
' 외주처 선택
'===============================================================================
Private Function SelectVendor(ByRef Item As ProductionItem, ByRef lineSchedule As Object, ByVal targetMonth As Date) As Integer
    SelectVendor = 0
    
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
    
    Dim bestVendor As Integer
    bestVendor = 0
    Dim bestPriority As Integer
    bestPriority = 999
    
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

'===============================================================================
' 특수공정 처리 가능 여부
'===============================================================================
Private Function CanHandleProcess(ByRef vendor As VendorInfo, ByVal ProcessType As String) As Boolean
    If ProcessType = "normal" Or Len(ProcessType) = 0 Then
        CanHandleProcess = True
        Exit Function
    End If
    
    CanHandleProcess = InStr(1, vendor.Capabilities, ProcessType, vbTextCompare) > 0
End Function

'===============================================================================
' 라인 배정
'===============================================================================
Private Function AssignLine(ByRef vendor As VendorInfo, ByVal StartDate As Date, ByVal DaysNeeded As Integer, ByRef lineSchedule As Object) As Integer
    Dim line As Integer
    For line = 1 To vendor.LineCount
        Dim canAssign As Boolean
        canAssign = True
        
        Dim d As Integer
        Dim checkDate As Date
        checkDate = StartDate
        
        For d = 1 To DaysNeeded
            Do While Weekday(checkDate) = 1 Or Weekday(checkDate) = 7
                checkDate = checkDate + 1
            Loop
            
            Dim scheduleKey As String
            scheduleKey = vendor.Name & "_" & line & "_" & Format(checkDate, "yyyy-mm-dd")
            
            If lineSchedule.Exists(scheduleKey) Then
                canAssign = False
                Exit For
            End If
            
            checkDate = checkDate + 1
        Next d
        
        If canAssign Then
            checkDate = StartDate
            For d = 1 To DaysNeeded
                Do While Weekday(checkDate) = 1 Or Weekday(checkDate) = 7
                    checkDate = checkDate + 1
                Loop
                scheduleKey = vendor.Name & "_" & line & "_" & Format(checkDate, "yyyy-mm-dd")
                lineSchedule(scheduleKey) = True
                checkDate = checkDate + 1
            Next d
            
            AssignLine = line
            Exit Function
        End If
    Next line
    
    AssignLine = 1
End Function

'===============================================================================
' 유틸리티 함수들
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
        .Interior.Color = RGB(68, 84, 106)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    With ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ws.Columns("A:" & Chr(64 + colCount)).AutoFit
End Sub

Private Sub CreateStatCard(ByRef ws As Worksheet, ByVal startCol As String, ByVal startRow As Integer, _
                          ByVal title As String, ByVal Value As String, ByVal Color As Long)
    Dim rng As Range
    Set rng = ws.Range(startCol & startRow & ":" & Chr(Asc(startCol) + 1) & (startRow + 2))
    
    rng.Merge
    
    With rng
        .Interior.Color = Color
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range(startCol & startRow).Value = title & vbLf & vbLf & Value
    ws.Range(startCol & startRow).Font.Size = 11
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
        Case "수축"
            ParseSpecialProcess = "shrink"
        Case "교반"
            ParseSpecialProcess = "mixing"
        Case "고주파"
            ParseSpecialProcess = "highFrequency"
        Case Else
            ParseSpecialProcess = "normal"
    End Select
End Function

Private Function ExtractClientCode(ByVal productCode As String) As String
    If Len(productCode) < 4 Then
        ExtractClientCode = ""
        Exit Function
    End If
    
    Dim startPos As Integer
    startPos = 1
    
    If IsNumeric(Left(productCode, 1)) Then
        startPos = 2
    End If
    
    Dim code As String
    code = Mid(productCode, startPos, 3)
    
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
        Dim parts() As String
        parts = Split(transferDateStr, "/")
        If UBound(parts) >= 1 Then
            transferDate = DateSerial(targetYear, CInt(parts(0)), CInt(parts(1)))
        End If
    ElseIf InStr(transferDateStr, "월") > 0 Then
        Dim monthPart As String, dayPart As String
        monthPart = Left(transferDateStr, InStr(transferDateStr, "월") - 1)
        dayPart = Mid(transferDateStr, InStr(transferDateStr, "월") + 1)
        dayPart = Replace(dayPart, "일", "")
        dayPart = Trim(dayPart)
        transferDate = DateSerial(targetYear, CInt(Trim(monthPart)), CInt(dayPart))
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
    Dim result As Date
    result = StartDate
    
    Dim addedDays As Integer
    addedDays = 0
    
    Do While addedDays < daysToAdd
        result = result + 1
        If Weekday(result) <> 1 And Weekday(result) <> 7 Then
            addedDays = addedDays + 1
        End If
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

Private Function GetVendorMonthlyTarget(ByVal vendorName As String) As Long
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).Name = vendorName Then
            GetVendorMonthlyTarget = Vendors(i).MonthlyTarget
            Exit Function
        End If
    Next i
    GetVendorMonthlyTarget = 0
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

' 외주처 색상 (진한색 - 헤더용)
Private Function GetVendorColor(ByVal vendorName As String) As Long
    Select Case vendorName
        Case "위드맘"
            GetVendorColor = RGB(59, 130, 246)
        Case "리니어"
            GetVendorColor = RGB(34, 197, 94)
        Case "그램"
            GetVendorColor = RGB(168, 85, 247)
        Case "이시스"
            GetVendorColor = RGB(249, 115, 22)
        Case "엘루오"
            GetVendorColor = RGB(236, 72, 153)
        Case "케이코스텍"
            GetVendorColor = RGB(6, 182, 212)
        Case "다미"
            GetVendorColor = RGB(245, 158, 11)
        Case "배정불가"
            GetVendorColor = RGB(156, 163, 175)
        Case Else
            GetVendorColor = RGB(200, 200, 200)
    End Select
End Function

' 외주처 색상 (연한색 - 셀 배경용)
Private Function GetVendorLightColor(ByVal vendorName As String) As Long
    Select Case vendorName
        Case "위드맘"
            GetVendorLightColor = RGB(219, 234, 254)
        Case "리니어"
            GetVendorLightColor = RGB(220, 252, 231)
        Case "그램"
            GetVendorLightColor = RGB(243, 232, 255)
        Case "이시스"
            GetVendorLightColor = RGB(255, 237, 213)
        Case "엘루오"
            GetVendorLightColor = RGB(252, 231, 243)
        Case "케이코스텍"
            GetVendorLightColor = RGB(207, 250, 254)
        Case "다미"
            GetVendorLightColor = RGB(254, 243, 199)
        Case "배정불가"
            GetVendorLightColor = RGB(243, 244, 246)
        Case Else
            GetVendorLightColor = RGB(249, 250, 251)
    End Select
End Function
