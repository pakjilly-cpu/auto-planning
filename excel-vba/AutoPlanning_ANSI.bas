'===============================================================================
' 외주처 생산계획 자동화 시스템 - VBA 매크로
' Version: 1.0
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
    Capabilities As String  ' 콤마로 구분: "normal,shrink,mixing"
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

' 전역 변수
Public Vendors() As VendorInfo
Public VendorCount As Integer
Public ClientMappings As Object ' Dictionary: ClientCode -> VendorID

'===============================================================================
' 메인 진입점
'===============================================================================
Public Sub 자동배정실행()
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
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
    
    ' 4. 결과 시트 생성
    CreateOutputSheet results, itemCount
    
    ' 5. 대시보드 생성
    CreateDashboard results, itemCount
    
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
' 초기 설정 (시트 생성)
'===============================================================================
Public Sub 초기설정()
    Application.ScreenUpdating = False
    
    ' 설정_외주처 시트 생성
    CreateVendorSheet
    
    ' 설정_고객매칭 시트 생성
    CreateMappingSheet
    
    ' auto-planning 시트 생성 (빈 시트)
    CreateOrClearSheet SHEET_OUTPUT
    
    Application.ScreenUpdating = True
    
    MsgBox "초기 설정 완료!" & vbCrLf & _
           "- " & SHEET_VENDORS & " 시트: 외주처 정보 설정" & vbCrLf & _
           "- " & SHEET_MAPPING & " 시트: 고객사-외주처 매칭 설정" & vbCrLf & vbCrLf & _
           "설정을 확인/수정한 후 '자동배정실행' 매크로를 실행하세요.", vbInformation
End Sub

'===============================================================================
' 설정 시트 생성 - 외주처
'===============================================================================
Private Sub CreateVendorSheet()
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_VENDORS)
    
    ' 헤더
    ws.Range("A1:G1").Value = Array("외주처ID", "외주처명", "라인수", "라인당일일생산량", "가능공정", "월간목표", "우선순위")
    
    ' 기본 데이터 (웹 버전과 동일)
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
    
    ' 서식
    FormatSettingSheet ws, 7, UBound(data) + 2
    
    ' 가능공정 설명 추가
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
    
    ' 헤더
    ws.Range("A1:C1").Value = Array("고객사코드", "외주처ID", "우선순위")
    
    ' 기본 데이터
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
    
    ' 서식
    FormatSettingSheet ws, 3, UBound(data) + 2
    
    ' 설명 추가
    ws.Range("E1").Value = "* 제품코드에서 추출된 3자리 영문코드(예: 9CLO... → CLO)가"
    ws.Range("E2").Value = "  해당 외주처에 우선 배정됩니다."
    ws.Range("E1:E2").Font.Color = RGB(100, 100, 100)
End Sub

'===============================================================================
' 설정 로드
'===============================================================================
Private Function LoadSettings() As Boolean
    On Error GoTo ErrorHandler
    
    ' 외주처 설정 로드
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
    
    ' 고객매칭 설정 로드
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
    
    ' 먼저 유효한 데이터 개수 세기
    itemCount = 0
    Dim i As Long
    For i = 2 To lastRow
        Dim productCode As String, qty As Long, transferDate As String
        productCode = Trim(CStr(ws.Cells(i, COL_PRODUCT_CODE).Value))
        qty = ParseQuantity(ws.Cells(i, COL_QUANTITY).Value)
        transferDate = Trim(CStr(ws.Cells(i, COL_TRANSFER_DATE).Value))
        
        ' 제품코드 없거나 수량 0이면 스킵
        If Len(productCode) = 0 Or qty = 0 Then GoTo NextRow
        ' 이동일이 "미정" 또는 빈값이면 스킵
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
    
    ' 라인별 점유 현황: Dictionary of Dictionary
    ' Key: "외주처명_라인번호_날짜" -> 점유여부
    Dim lineSchedule As Object
    Set lineSchedule = CreateObject("Scripting.Dictionary")
    
    ' 현재 월의 근무일 계산
    Dim targetMonth As Date
    targetMonth = Date ' 현재 날짜 기준
    
    Dim monthStart As Date, monthEnd As Date
    monthStart = DateSerial(Year(targetMonth), Month(targetMonth), 1)
    monthEnd = DateSerial(Year(targetMonth), Month(targetMonth) + 1, 0)
    
    Dim i As Long
    For i = 1 To itemCount
        results(i).Item = items(i)
        
        ' 1. 이미 외주처가 배정되어 있으면 그대로 사용
        If Len(items(i).AssignedVendor) > 0 Then
            results(i).VendorName = items(i).AssignedVendor
            results(i).LineNumber = 1
            results(i).StartDate = GetProductionStartDate(items(i).TransferDate, Year(targetMonth))
            results(i).DaysNeeded = CalculateDaysNeeded(items(i).Quantity, GetVendorCapacity(items(i).AssignedVendor))
        Else
            ' 2. 자동 배정
            Dim vendorIdx As Integer
            vendorIdx = SelectVendor(items(i), lineSchedule, targetMonth)
            
            If vendorIdx > 0 Then
                results(i).VendorName = Vendors(vendorIdx).Name
                results(i).StartDate = GetProductionStartDate(items(i).TransferDate, Year(targetMonth))
                results(i).DaysNeeded = CalculateDaysNeeded(items(i).Quantity, Vendors(vendorIdx).DailyCapacity)
                
                ' 라인 배정 및 스케줄 등록
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
' 외주처 선택 로직
'===============================================================================
Private Function SelectVendor(ByRef item As ProductionItem, ByRef lineSchedule As Object, ByVal targetMonth As Date) As Integer
    SelectVendor = 0
    
    ' 1. 고객사 매칭 확인
    If Len(item.ClientCode) > 0 Then
        If ClientMappings.Exists(item.ClientCode) Then
            Dim mappedVendorId As String
            mappedVendorId = ClientMappings(item.ClientCode)
            
            Dim j As Integer
            For j = 1 To VendorCount
                If Vendors(j).ID = mappedVendorId Then
                    If CanHandleProcess(Vendors(j), item.SpecialProcess) Then
                        SelectVendor = j
                        Exit Function
                    End If
                End If
            Next j
        End If
    End If
    
    ' 2. 우선순위 기반 선택 (특수공정 가능 여부 확인)
    Dim bestVendor As Integer
    bestVendor = 0
    Dim bestPriority As Integer
    bestPriority = 999
    
    For j = 1 To VendorCount
        If CanHandleProcess(Vendors(j), item.SpecialProcess) Then
            If Vendors(j).Priority < bestPriority Then
                bestPriority = Vendors(j).Priority
                bestVendor = j
            ElseIf Vendors(j).Priority = bestPriority Then
                ' 같은 우선순위면 월간 목표가 있는 곳 우선
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
Private Function CanHandleProcess(ByRef vendor As VendorInfo, ByVal processType As String) As Boolean
    If processType = "normal" Or Len(processType) = 0 Then
        CanHandleProcess = True
        Exit Function
    End If
    
    CanHandleProcess = InStr(1, vendor.Capabilities, processType, vbTextCompare) > 0
End Function

'===============================================================================
' 라인 배정
'===============================================================================
Private Function AssignLine(ByRef vendor As VendorInfo, ByVal startDate As Date, ByVal daysNeeded As Integer, ByRef lineSchedule As Object) As Integer
    Dim bestLine As Integer
    bestLine = 1
    
    Dim line As Integer
    For line = 1 To vendor.LineCount
        Dim canAssign As Boolean
        canAssign = True
        
        Dim d As Integer
        Dim checkDate As Date
        checkDate = startDate
        
        For d = 1 To daysNeeded
            ' 주말 건너뛰기
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
            ' 이 라인에 스케줄 등록
            checkDate = startDate
            For d = 1 To daysNeeded
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
    
    ' 모든 라인이 꽉 찼으면 1번 라인 반환
    AssignLine = 1
End Function

'===============================================================================
' 결과 시트 생성
'===============================================================================
Private Sub CreateOutputSheet(ByRef results() As AllocationResult, ByVal itemCount As Long)
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_OUTPUT)
    
    ' 대시보드 영역 확보 (상단 12행)
    Dim dashboardRows As Integer
    dashboardRows = 14
    
    ' 생산계획표 시작 위치
    Dim planStartRow As Integer
    planStartRow = dashboardRows + 2
    
    ' 헤더
    ws.Cells(planStartRow, 1).Value = "생산계획표"
    ws.Cells(planStartRow, 1).Font.Size = 14
    ws.Cells(planStartRow, 1).Font.Bold = True
    
    planStartRow = planStartRow + 2
    
    ' 테이블 헤더
    Dim headers As Variant
    headers = Array("No", "제품코드", "제품명", "수량", "외주처", "라인", "생산시작일", "소요일수", "납기일", "이동일", "특수공정")
    
    Dim col As Integer
    For col = 0 To UBound(headers)
        ws.Cells(planStartRow, col + 1).Value = headers(col)
    Next col
    
    ' 헤더 서식
    With ws.Range(ws.Cells(planStartRow, 1), ws.Cells(planStartRow, UBound(headers) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(68, 84, 106)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 데이터 출력
    Dim i As Long
    For i = 1 To itemCount
        Dim row As Long
        row = planStartRow + i
        
        ws.Cells(row, 1).Value = i
        ws.Cells(row, 2).Value = results(i).Item.ProductCode
        ws.Cells(row, 3).Value = results(i).Item.ProductName
        ws.Cells(row, 4).Value = results(i).Item.Quantity
        ws.Cells(row, 5).Value = results(i).VendorName
        ws.Cells(row, 6).Value = results(i).LineNumber
        ws.Cells(row, 7).Value = results(i).StartDate
        ws.Cells(row, 8).Value = results(i).DaysNeeded
        ws.Cells(row, 9).Value = results(i).Item.DeliveryDate
        ws.Cells(row, 10).Value = results(i).Item.TransferDate
        ws.Cells(row, 11).Value = GetProcessName(results(i).Item.SpecialProcess)
        
        ' 외주처별 색상
        Dim vendorColor As Long
        vendorColor = GetVendorColor(results(i).VendorName)
        ws.Cells(row, 5).Interior.Color = vendorColor
        
        ' 배정불가는 빨간색
        If results(i).VendorName = "배정불가" Then
            ws.Cells(row, 5).Font.Color = RGB(255, 255, 255)
        End If
    Next i
    
    ' 테이블 서식
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(planStartRow, 1), ws.Cells(planStartRow + itemCount, UBound(headers) + 1))
    
    With dataRange
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Columns(4).NumberFormat = "#,##0"
        .Columns(7).NumberFormat = "yyyy-mm-dd"
    End With
    
    ' 열 너비 자동 조정
    ws.Columns("A:K").AutoFit
    
    ' 시트 활성화
    ws.Activate
End Sub

'===============================================================================
' 대시보드 생성
'===============================================================================
Private Sub CreateDashboard(ByRef results() As AllocationResult, ByVal itemCount As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_OUTPUT)
    
    ' 통계 계산
    Dim totalQty As Long
    Dim vendorQty As Object ' Dictionary
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
    
    ' 대시보드 타이틀
    ws.Range("A1").Value = "외주처 생산계획 대시보드"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    
    ws.Range("A2").Value = "생성일시: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Range("A2").Font.Color = RGB(128, 128, 128)
    
    ' 통계 카드
    ' 카드 1: 총 품목 수
    CreateStatCard ws, "B", 4, "총 품목 수", itemCount & " 건", RGB(59, 130, 246)
    
    ' 카드 2: 총 생산량
    CreateStatCard ws, "D", 4, "총 생산량", Format(totalQty, "#,##0") & " 개", RGB(34, 197, 94)
    
    ' 카드 3: 외주처 수
    CreateStatCard ws, "F", 4, "배정 외주처", vendorQty.Count & " 곳", RGB(168, 85, 247)
    
    ' 카드 4: 평균 수량
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
        If col > 10 Then Exit For ' 최대 9개까지
        
        ws.Cells(10, col).Value = vKey
        ws.Cells(11, col).Value = vendorQty(vKey)
        ws.Cells(11, col).NumberFormat = "#,##0"
        
        ' 목표 대비 달성률
        Dim targetQty As Long
        targetQty = GetVendorMonthlyTarget(CStr(vKey))
        If targetQty > 0 Then
            ws.Cells(12, col).Value = Format(vendorQty(vKey) / targetQty, "0%")
        Else
            ws.Cells(12, col).Value = "-"
        End If
        
        ' 색상
        ws.Cells(10, col).Interior.Color = GetVendorColor(CStr(vKey))
        ws.Cells(10, col).Font.Color = RGB(255, 255, 255)
        ws.Cells(10, col).Font.Bold = True
        
        col = col + 1
    Next vKey
    
    ' 테두리
    If col > 2 Then
        With ws.Range(ws.Cells(10, 2), ws.Cells(12, col - 1))
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
    End If
End Sub

'===============================================================================
' 통계 카드 생성
'===============================================================================
Private Sub CreateStatCard(ByRef ws As Worksheet, ByVal startCol As String, ByVal startRow As Integer, _
                          ByVal title As String, ByVal value As String, ByVal color As Long)
    Dim rng As Range
    Set rng = ws.Range(startCol & startRow & ":" & Chr(Asc(startCol) + 1) & (startRow + 2))
    
    rng.Merge
    
    With rng
        .Interior.Color = color
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range(startCol & startRow).Value = title & vbLf & vbLf & value
    ws.Range(startCol & startRow).Font.Size = 11
End Sub

'===============================================================================
' 유틸리티 함수들
'===============================================================================

' 시트 생성 또는 초기화
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
    End If
    
    Set CreateOrClearSheet = ws
End Function

' 설정 시트 서식
Private Sub FormatSettingSheet(ByRef ws As Worksheet, ByVal colCount As Integer, ByVal rowCount As Integer)
    ' 헤더 서식
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
        .Font.Bold = True
        .Interior.Color = RGB(68, 84, 106)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' 테두리
    With ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' 열 너비
    ws.Columns("A:" & Chr(64 + colCount)).AutoFit
End Sub

' 수량 파싱 (콤마 제거)
Private Function ParseQuantity(ByVal value As Variant) As Long
    On Error Resume Next
    If IsEmpty(value) Or value = "" Then
        ParseQuantity = 0
    ElseIf IsNumeric(value) Then
        ParseQuantity = CLng(value)
    Else
        ParseQuantity = CLng(Replace(CStr(value), ",", ""))
    End If
    On Error GoTo 0
End Function

' 특수공정 파싱
Private Function ParseSpecialProcess(ByVal value As String) As String
    Select Case Trim(value)
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

' 특수공정 한글명
Private Function GetProcessName(ByVal processCode As String) As String
    Select Case processCode
        Case "shrink"
            GetProcessName = "수축"
        Case "mixing"
            GetProcessName = "교반"
        Case "highFrequency"
            GetProcessName = "고주파"
        Case Else
            GetProcessName = "일반"
    End Select
End Function

' 고객사 코드 추출 (예: 9CLO1360310 -> CLO)
Private Function ExtractClientCode(ByVal productCode As String) As String
    If Len(productCode) < 4 Then
        ExtractClientCode = ""
        Exit Function
    End If
    
    Dim i As Integer
    Dim startPos As Integer
    startPos = 1
    
    ' 첫 글자가 숫자면 건너뛰기
    If IsNumeric(Left(productCode, 1)) Then
        startPos = 2
    End If
    
    ' 3자리 영문 추출
    Dim code As String
    code = Mid(productCode, startPos, 3)
    
    ' 모두 영문인지 확인
    For i = 1 To Len(code)
        If Not (Mid(code, i, 1) Like "[A-Z]") Then
            ExtractClientCode = ""
            Exit Function
        End If
    Next i
    
    ExtractClientCode = code
End Function

' 이동일 파싱 및 생산시작일 계산 (이동일+1근무일)
Private Function GetProductionStartDate(ByVal transferDateStr As String, ByVal targetYear As Integer) As Date
    Dim transferDate As Date
    
    On Error Resume Next
    
    ' 다양한 형식 파싱 시도
    ' "1/9", "01/09" 형식
    If InStr(transferDateStr, "/") > 0 Then
        Dim parts() As String
        parts = Split(transferDateStr, "/")
        If UBound(parts) >= 1 Then
            transferDate = DateSerial(targetYear, CInt(parts(0)), CInt(parts(1)))
        End If
    ' "1월 9일" 형식
    ElseIf InStr(transferDateStr, "월") > 0 Then
        Dim monthPart As String, dayPart As String
        monthPart = Left(transferDateStr, InStr(transferDateStr, "월") - 1)
        dayPart = Mid(transferDateStr, InStr(transferDateStr, "월") + 1)
        dayPart = Replace(dayPart, "일", "")
        dayPart = Trim(dayPart)
        transferDate = DateSerial(targetYear, CInt(Trim(monthPart)), CInt(dayPart))
    Else
        ' 날짜로 직접 변환 시도
        transferDate = CDate(transferDateStr)
    End If
    
    On Error GoTo 0
    
    ' 이동일+1근무일 계산
    If transferDate > 0 Then
        GetProductionStartDate = GetNextWorkingDay(transferDate, 1)
    Else
        GetProductionStartDate = GetNextWorkingDay(Date, 1)
    End If
End Function

' 다음 근무일 계산
Private Function GetNextWorkingDay(ByVal startDate As Date, ByVal daysToAdd As Integer) As Date
    Dim result As Date
    result = startDate
    
    Dim addedDays As Integer
    addedDays = 0
    
    Do While addedDays < daysToAdd
        result = result + 1
        ' 주말이 아니면 카운트
        If Weekday(result) <> 1 And Weekday(result) <> 7 Then
            addedDays = addedDays + 1
        End If
    Loop
    
    GetNextWorkingDay = result
End Function

' 필요 일수 계산
Private Function CalculateDaysNeeded(ByVal quantity As Long, ByVal dailyCapacity As Long) As Integer
    If dailyCapacity <= 0 Then dailyCapacity = 10000
    CalculateDaysNeeded = Application.WorksheetFunction.Ceiling(quantity / dailyCapacity, 1)
End Function

' 외주처명으로 CAPA 조회
Private Function GetVendorCapacity(ByVal vendorName As String) As Long
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).Name = vendorName Then
            GetVendorCapacity = Vendors(i).DailyCapacity
            Exit Function
        End If
    Next i
    GetVendorCapacity = 10000 ' 기본값
End Function

' 외주처 월간 목표 조회
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

' 외주처별 색상
Private Function GetVendorColor(ByVal vendorName As String) As Long
    Select Case vendorName
        Case "위드맘"
            GetVendorColor = RGB(59, 130, 246)   ' Blue
        Case "리니어"
            GetVendorColor = RGB(34, 197, 94)    ' Green
        Case "그램"
            GetVendorColor = RGB(168, 85, 247)   ' Purple
        Case "이시스"
            GetVendorColor = RGB(249, 115, 22)   ' Orange
        Case "엘루오"
            GetVendorColor = RGB(236, 72, 153)   ' Pink
        Case "케이코스텍"
            GetVendorColor = RGB(6, 182, 212)    ' Cyan
        Case "다미"
            GetVendorColor = RGB(245, 158, 11)   ' Amber
        Case "배정불가"
            GetVendorColor = RGB(156, 163, 175)  ' Gray
        Case Else
            GetVendorColor = RGB(200, 200, 200)
    End Select
End Function
