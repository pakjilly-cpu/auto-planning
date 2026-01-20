'===============================================================================
' 외주처 생산계획 자동화 시스템 - VBA 매크로 v9.0
'===============================================================================
' [v9.0 변경사항]
' - Raw 데이터 컬럼 매핑 변경 (실제 사용 양식 반영)
'   A:이동일, B:발주, D:외주처, F:품번, G:품목, H:공정, I:수량, L:납기, O:비고
' - 배정 제외 조건: A,B,L열이 "미정" 또는 빈칸
' - 배정 제외: A열="미정" + D열에 업체 있음 (아직 확정 안됨)
' - D열 직접 지정시 유지 + 문제시 파스텔 빨간색 경고
' - 납기 우선순위: 날짜(1/26) > ASAP
' - ASAP은 배정하되 날짜있는 품목 밀어내지 않음
' - 고객매칭 변경시 파스텔 주황색 + 메모
' - 생산계획표 상단에 조회창 추가
'===============================================================================

Option Explicit

'===============================================================================
' 상수 정의
'===============================================================================
Public Const SHEET_RAW As String = "Raw데이터"
Public Const SHEET_VENDORS As String = "설정_외주처"
Public Const SHEET_MAPPING As String = "설정_고객매칭"
Public Const SHEET_OUTPUT As String = "생산계획표"

' Raw 데이터 컬럼 (v9.0 - 실제 양식 반영)
' A:이동일, B:발주, C:담당자, D:외주처, E:구분, F:품번, G:품목, H:공정, I:수량,
' J:제조, K:자재, L:납기, M:외주사유, N:특이사항, O:비고, P:주간계획
Public Const RAW_COL_TRANSFER As Integer = 1      ' A: 이동일
Public Const RAW_COL_ORDER As Integer = 2         ' B: 발주 (예정/완료)
Public Const RAW_COL_MANAGER As Integer = 3       ' C: 담당자
Public Const RAW_COL_VENDOR As Integer = 4        ' D: 외주처 (수동 지정)
Public Const RAW_COL_TYPE As Integer = 5          ' E: 구분 (초도품/기존품 등)
Public Const RAW_COL_CODE As Integer = 6          ' F: 품번
Public Const RAW_COL_NAME As Integer = 7          ' G: 품목
Public Const RAW_COL_PROCESS As Integer = 8       ' H: 공정
Public Const RAW_COL_QTY As Integer = 9           ' I: 수량
Public Const RAW_COL_MANUFACTURE As Integer = 10  ' J: 제조
Public Const RAW_COL_MATERIAL As Integer = 11     ' K: 자재
Public Const RAW_COL_DUEDATE As Integer = 12      ' L: 납기
Public Const RAW_COL_REASON As Integer = 13       ' M: 외주사유
Public Const RAW_COL_SPECIAL As Integer = 14      ' N: 특이사항
Public Const RAW_COL_REMARKS As Integer = 15      ' O: 비고
Public Const RAW_COL_WEEKLY As Integer = 16       ' P: 주간계획

' 자동 생성 컬럼 (Q~W)
Public Const RAW_COL_CLIENT As Integer = 17       ' Q: 고객사코드 (자동)
Public Const RAW_COL_ASSIGNED As Integer = 18     ' R: 배정업체 (자동)
Public Const RAW_COL_PRODSTART As Integer = 19    ' S: 생산시작일 (자동)
Public Const RAW_COL_PRODEND As Integer = 20      ' T: 생산완료일 (자동)
Public Const RAW_COL_DUECHECK As Integer = 21     ' U: 납기준수 (자동)
Public Const RAW_COL_ORIGINAL As Integer = 22     ' V: 최초배정업체 (자동)
Public Const RAW_COL_CHANGED As Integer = 23      ' W: 업체변경 (자동)
Public Const RAW_COL_WARNING As Integer = 24      ' X: 경고 (자동)

' Capa 관련
Public Const HOURS_BASIC As Integer = 8
Public Const HOURS_OVERTIME As Integer = 3
Public Const HOURS_SATURDAY As Integer = 8
Public Const SET_CAPA_FACTOR As Double = 0.8
Public Const FIXED_DAYS As Integer = 2

'===============================================================================
' 파스텔톤 색상 (v9.0)
'===============================================================================
' 경고 - 파스텔 빨간색
Private Const COLOR_PASTEL_RED As Long = 16764108      ' RGB(252, 220, 220)
Private Const COLOR_PASTEL_RED_TEXT As Long = 6579300  ' RGB(100, 60, 100)

' 고객매칭 변경 - 파스텔 주황색
Private Const COLOR_PASTEL_ORANGE As Long = 13434879   ' RGB(255, 235, 205)
Private Const COLOR_PASTEL_ORANGE_TEXT As Long = 4235520 ' RGB(64, 128, 140)

' 납기 위험 - 파스텔 노란색
Private Const COLOR_PASTEL_YELLOW As Long = 13434879   ' RGB(255, 250, 205)

' 정상 - 파스텔 녹색
Private Const COLOR_PASTEL_GREEN As Long = 13434828    ' RGB(220, 252, 220)

'===============================================================================
' 타입 정의
'===============================================================================
Public Type VendorInfo
    ID As String
    Name As String
    LineCount As Integer
    CapaPerLine As Long
    Capabilities As String
    MonthlyTarget As Long
    Priority As Integer
End Type

Public Type ProductionItem
    RowNum As Long
    ProductCode As String
    ProductName As String
    Quantity As Long
    TransferDate As Date          ' 이동일
    TransferStr As String         ' 이동일 원본 문자열
    OrderStatus As String         ' 발주상태 (예정/완료)
    DueDate As Date               ' 납기
    DueDateStr As String          ' 납기 원본 문자열 (ASAP 등)
    IsASAP As Boolean             ' ASAP 여부
    ProcessType As String
    SpecialProcess As String      ' 특이사항에서 추출
    DesignatedVendor As String    ' D열 수동 지정 업체
    ItemType As String            ' 구분 (초도품/기존품)
    Remarks As String
    ClientCode As String
    
    ' 자동 계산 필드
    ProductionDays As Integer
    ProductionStartDate As Date
    ProductionEndDate As Date
    DueCheckStatus As String
    
    ' 배정 결과
    AssignedVendor As String
    AssignedLine As Integer
    OriginalVendor As String
    VendorChanged As Boolean
    OriginalMatchedVendor As String  ' 고객매칭 원래 업체
    
    ' 상태
    ShouldSkip As Boolean         ' 배정 제외 여부
    SkipReason As String          ' 제외 사유
    
    ' 세트품 관련
    IsSetItem As Boolean
    SetGroupID As String
    SetTotalQty As Long
    SetLeaderRow As Long
End Type

Public Type AllocationResult
    Item As ProductionItem
    Success As Boolean
    FailReason As String
    AlternativeVendors As String
    
    ' 경고
    HasWarning As Boolean
    WarningType As String
    WarningMessage As String
End Type

'===============================================================================
' 전역 변수
'===============================================================================
Public Vendors() As VendorInfo
Public VendorCount As Integer
Public ClientMappings As Object
Public LineSchedule As Object

'===============================================================================
' 색상 팔레트
'===============================================================================
Private Function GetVendorColor(vendorName As String) As Long
    Select Case vendorName
        Case "위드맘": GetVendorColor = RGB(59, 130, 246)
        Case "리니어": GetVendorColor = RGB(16, 185, 129)
        Case "그램": GetVendorColor = RGB(249, 115, 22)
        Case "이시스": GetVendorColor = RGB(139, 92, 246)
        Case "엘루오": GetVendorColor = RGB(236, 72, 153)
        Case "케이코스텍": GetVendorColor = RGB(20, 184, 166)
        Case "다미": GetVendorColor = RGB(245, 158, 11)
        Case Else: GetVendorColor = RGB(107, 114, 128)
    End Select
End Function

Private Function GetVendorLightColor(vendorName As String) As Long
    Select Case vendorName
        Case "위드맘": GetVendorLightColor = RGB(219, 234, 254)
        Case "리니어": GetVendorLightColor = RGB(209, 250, 229)
        Case "그램": GetVendorLightColor = RGB(255, 237, 213)
        Case "이시스": GetVendorLightColor = RGB(237, 233, 254)
        Case "엘루오": GetVendorLightColor = RGB(252, 231, 243)
        Case "케이코스텍": GetVendorLightColor = RGB(204, 251, 241)
        Case "다미": GetVendorLightColor = RGB(254, 243, 199)
        Case Else: GetVendorLightColor = RGB(243, 244, 246)
    End Select
End Function

' 파스텔톤 경고 색상 (v9.0)
Private Function GetPastelRed() As Long
    GetPastelRed = RGB(255, 228, 225)  ' 연한 살몬핑크
End Function

Private Function GetPastelOrange() As Long
    GetPastelOrange = RGB(255, 239, 213)  ' 연한 피치
End Function

Private Function GetPastelYellow() As Long
    GetPastelYellow = RGB(255, 250, 205)  ' 연한 레몬
End Function

Private Function GetPastelGreen() As Long
    GetPastelGreen = RGB(220, 252, 220)  ' 연한 민트
End Function

'===============================================================================
' 메인 실행 매크로
'===============================================================================
Public Sub 자동배정실행()
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    ' 초기화
    Set LineSchedule = CreateObject("Scripting.Dictionary")
    
    ' 설정 로드
    If Not LoadSettings() Then
        MsgBox "설정을 로드할 수 없습니다." & vbCrLf & _
               "'초기설정' 매크로를 먼저 실행해주세요.", vbExclamation, "오류"
        GoTo Cleanup
    End If
    
    ' Raw 데이터 읽기
    Dim items() As ProductionItem
    Dim itemCount As Long
    If Not ReadRawData(items, itemCount) Then
        MsgBox "Raw데이터 시트에서 데이터를 읽을 수 없습니다.", vbExclamation, "오류"
        GoTo Cleanup
    End If
    
    If itemCount = 0 Then
        MsgBox "처리할 데이터가 없습니다.", vbInformation, "알림"
        GoTo Cleanup
    End If
    
    ' 배정 제외 품목 체크 (미정/빈칸)
    CheckSkipItems items, itemCount
    
    ' 세트품 식별
    IdentifySetItems items, itemCount
    
    ' 순산 로직: 이동일 -> 생산완료 -> 납기 체크
    CalculateProductionDates items, itemCount
    
    ' 날짜 범위 파악
    Dim minDate As Date, maxDate As Date
    GetDateRange items, itemCount, minDate, maxDate
    
    ' 자동 배정 실행 (납기>ASAP 우선순위 적용)
    Dim results() As AllocationResult
    AllocateProduction items, itemCount, results, minDate, maxDate
    
    ' Raw 데이터 시트에 결과 기록
    WriteResultsToRaw items, itemCount, results
    
    ' 생산계획표 생성 (조회창 포함)
    On Error Resume Next
    CreateProductionPlanView results, itemCount, minDate, maxDate
    If Err.Number <> 0 Then
        MsgBox "생산계획표 생성 중 오류: " & Err.Description, vbExclamation, "경고"
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' 완료 메시지
    Dim elapsed As Double
    elapsed = Timer - startTime
    
    Dim skipCount As Long, assignedCount As Long, failCount As Long
    Dim i As Long
    For i = 1 To itemCount
        If items(i).ShouldSkip Then
            skipCount = skipCount + 1
        ElseIf results(i).Success Then
            assignedCount = assignedCount + 1
        Else
            failCount = failCount + 1
        End If
    Next i
    
    MsgBox "자동배정 완료!" & vbCrLf & vbCrLf & _
           "전체: " & itemCount & "건" & vbCrLf & _
           "배정 완료: " & assignedCount & "건" & vbCrLf & _
           "배정 제외 (미정): " & skipCount & "건" & vbCrLf & _
           "배정 불가: " & failCount & "건" & vbCrLf & vbCrLf & _
           "소요 시간: " & Format(elapsed, "0.0") & "초", vbInformation, "완료"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    MsgBox "오류가 발생했습니다:" & vbCrLf & Err.Description, vbCritical, "오류"
    Resume Cleanup
End Sub

'===============================================================================
' 초기 설정 매크로
'===============================================================================
Public Sub 초기설정()
    Application.ScreenUpdating = False
    
    ' 외주처 설정 시트 생성
    CreateVendorSheet
    
    ' 고객매칭 설정 시트 생성
    CreateMappingSheet
    
    ' 생산계획표 시트 생성
    CreateOrClearSheet SHEET_OUTPUT
    
    Application.ScreenUpdating = True
    
    MsgBox "초기 설정이 완료되었습니다!" & vbCrLf & vbCrLf & _
           "1. 'Raw데이터' 시트에 데이터를 붙여넣으세요." & vbCrLf & _
           "2. '설정_외주처' 시트에서 외주처 정보를 확인하세요." & vbCrLf & _
           "3. '설정_고객매칭' 시트에서 고객매칭을 설정하세요." & vbCrLf & _
           "4. '자동배정실행' 매크로를 실행하세요.", vbInformation, "초기 설정 완료"
End Sub

'===============================================================================
' 외주처 설정 시트 생성
'===============================================================================
Private Sub CreateVendorSheet()
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_VENDORS)
    
    ' 헤더
    ws.Range("A1:G1").Value = Array("외주처ID", "외주처명", "라인수", "라인당일일Capa", _
                                      "가능공정", "월생산목표", "우선순위")
    
    With ws.Range("A1:G1")
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(59, 130, 246)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With
    
    ' 샘플 데이터
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
    
    ws.Columns("A:G").AutoFit
End Sub

'===============================================================================
' 고객매칭 설정 시트 생성
'===============================================================================
Private Sub CreateMappingSheet()
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_MAPPING)
    
    ' 헤더
    ws.Range("A1:C1").Value = Array("고객사코드", "외주처명", "우선순위")
    
    With ws.Range("A1:C1")
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(59, 130, 246)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With
    
    ' 샘플 데이터 (외주처명으로 변경)
    Dim data As Variant
    data = Array( _
        Array("CLO", "그램", 1), _
        Array("ERK", "그램", 1), _
        Array("DPD", "위드맘", 1), _
        Array("GDI", "위드맘", 1), _
        Array("MDH", "리니어", 1), _
        Array("APS", "리니어", 1), _
        Array("PUR", "케이코스텍", 1), _
        Array("OLV", "위드맘", 1), _
        Array("AMR", "리니어", 1), _
        Array("CJO", "그램", 1), _
        Array("LGH", "리니어", 1), _
        Array("EMT", "위드맘", 1), _
        Array("BNU", "리니어", 1), _
        Array("DTL", "리니어", 1), _
        Array("WAT", "위드맘", 1), _
        Array("HOT", "다미", 1) _
    )
    
    Dim i As Integer
    For i = 0 To UBound(data)
        ws.Range("A" & (i + 2) & ":C" & (i + 2)).Value = data(i)
    Next i
    
    ws.Columns("A:C").AutoFit
End Sub

'===============================================================================
' 설정 로드
'===============================================================================
Private Function LoadSettings() As Boolean
    On Error GoTo ErrorHandler
    
    ' 외주처 로드
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
        Vendors(i - 1).CapaPerLine = CLng(wsVendor.Cells(i, 4).Value)
        Vendors(i - 1).Capabilities = CStr(wsVendor.Cells(i, 5).Value)
        Vendors(i - 1).MonthlyTarget = CLng(wsVendor.Cells(i, 6).Value)
        Vendors(i - 1).Priority = CInt(wsVendor.Cells(i, 7).Value)
    Next i
    
    ' 고객매칭 로드 (외주처명으로 매핑)
    Dim wsMapping As Worksheet
    Set wsMapping = ThisWorkbook.Sheets(SHEET_MAPPING)
    
    Set ClientMappings = CreateObject("Scripting.Dictionary")
    
    lastRow = wsMapping.Cells(wsMapping.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        Dim clientCode As String, vendorName As String
        clientCode = Trim(CStr(wsMapping.Cells(i, 1).Value))
        vendorName = Trim(CStr(wsMapping.Cells(i, 2).Value))
        If Len(clientCode) > 0 And Len(vendorName) > 0 Then
            ClientMappings(clientCode) = vendorName
        End If
    Next i
    
    LoadSettings = True
    Exit Function
    
ErrorHandler:
    LoadSettings = False
End Function

'===============================================================================
' Raw 데이터 읽기 (v9.0 - 실제 양식 반영)
'===============================================================================
Private Function ReadRawData(ByRef items() As ProductionItem, ByRef itemCount As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' Raw데이터 시트 찾기 (없으면 첫번째 시트 사용)
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_RAW)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets(1)
    End If
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, RAW_COL_CODE).End(xlUp).Row
    
    If lastRow < 2 Then
        itemCount = 0
        ReadRawData = True
        Exit Function
    End If
    
    itemCount = lastRow - 1
    ReDim items(1 To itemCount)
    
    Dim i As Long
    For i = 2 To lastRow
        With items(i - 1)
            .RowNum = i
            
            ' 기본 정보 읽기
            .TransferStr = Trim(CStr(ws.Cells(i, RAW_COL_TRANSFER).Value))
            .OrderStatus = Trim(CStr(ws.Cells(i, RAW_COL_ORDER).Value))
            .DesignatedVendor = Trim(CStr(ws.Cells(i, RAW_COL_VENDOR).Value))
            .ItemType = Trim(CStr(ws.Cells(i, RAW_COL_TYPE).Value))
            .ProductCode = Trim(CStr(ws.Cells(i, RAW_COL_CODE).Value))
            .ProductName = Trim(CStr(ws.Cells(i, RAW_COL_NAME).Value))
            .ProcessType = Trim(CStr(ws.Cells(i, RAW_COL_PROCESS).Value))
            .SpecialProcess = Trim(CStr(ws.Cells(i, RAW_COL_SPECIAL).Value))
            .Remarks = Trim(CStr(ws.Cells(i, RAW_COL_REMARKS).Value))
            
            ' 수량 파싱 (콤마 제거)
            Dim qtyStr As String
            qtyStr = Replace(CStr(ws.Cells(i, RAW_COL_QTY).Value), ",", "")
            qtyStr = Replace(qtyStr, " ", "")
            .Quantity = CLng(Val(qtyStr))
            
            ' 이동일 파싱
            .TransferDate = ParseTransferDate(.TransferStr)
            
            ' 납기 파싱 (ASAP 체크)
            .DueDateStr = Trim(CStr(ws.Cells(i, RAW_COL_DUEDATE).Value))
            ParseDueDate .DueDateStr, .DueDate, .IsASAP
            
            ' 고객사코드 추출 (품번 2~4번째 글자)
            If Len(.ProductCode) >= 4 Then
                .ClientCode = Mid(.ProductCode, 2, 3)
            End If
            
            ' 기존 배정 정보 (이미 있으면 유지)
            .AssignedVendor = Trim(CStr(ws.Cells(i, RAW_COL_ASSIGNED).Value))
            .OriginalVendor = Trim(CStr(ws.Cells(i, RAW_COL_ORIGINAL).Value))
        End With
    Next i
    
    ReadRawData = True
    Exit Function
    
ErrorHandler:
    ReadRawData = False
End Function

'===============================================================================
' 이동일 파싱
'===============================================================================
Private Function ParseTransferDate(dateStr As String) As Date
    On Error GoTo ErrorHandler
    
    ParseTransferDate = 0
    dateStr = Trim(dateStr)
    
    If Len(dateStr) = 0 Then Exit Function
    If LCase(dateStr) = "미정" Then Exit Function
    
    ' 날짜 형식 파싱 시도
    If IsDate(dateStr) Then
        ParseTransferDate = CDate(dateStr)
        Exit Function
    End If
    
    ' M/D 형식
    If InStr(dateStr, "/") > 0 Then
        Dim parts() As String
        parts = Split(dateStr, "/")
        If UBound(parts) >= 1 Then
            Dim m As Integer, d As Integer
            m = CInt(parts(0))
            d = CInt(parts(1))
            Dim y As Integer: y = Year(Date)
            If m < Month(Date) Then y = y + 1
            ParseTransferDate = DateSerial(y, m, d)
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ParseTransferDate = 0
End Function

'===============================================================================
' 납기 파싱 (ASAP 체크)
'===============================================================================
Private Sub ParseDueDate(dueDateStr As String, ByRef dueDate As Date, ByRef isASAP As Boolean)
    dueDate = 0
    isASAP = False
    
    dueDateStr = Trim(dueDateStr)
    If Len(dueDateStr) = 0 Then Exit Sub
    If LCase(dueDateStr) = "미정" Then Exit Sub
    
    ' ASAP 체크 (여러 패턴)
    If InStr(1, LCase(dueDateStr), "asap", vbTextCompare) > 0 Or _
       InStr(1, dueDateStr, "연속생산", vbTextCompare) > 0 Then
        isASAP = True
        
        ' ASAP이지만 날짜도 있는 경우 (예: "12/29 5만개만 출고요청, 나머지는 연속생산납품")
        Dim dateMatch As Date
        dateMatch = ExtractDateFromString(dueDateStr)
        If dateMatch > 0 Then
            dueDate = dateMatch
        End If
        Exit Sub
    End If
    
    ' 날짜 추출 시도
    dueDate = ExtractDateFromString(dueDateStr)
End Sub

'===============================================================================
' 문자열에서 날짜 추출
'===============================================================================
Private Function ExtractDateFromString(str As String) As Date
    On Error GoTo ErrorHandler
    
    ExtractDateFromString = 0
    str = Trim(str)
    If Len(str) = 0 Then Exit Function
    
    ' 직접 날짜인 경우
    If IsDate(str) Then
        ExtractDateFromString = CDate(str)
        Exit Function
    End If
    
    ' M/D 패턴 찾기 (예: 1/26, 12/29)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "(\d{1,2})/(\d{1,2})"
    
    If regex.Test(str) Then
        Dim matches As Object
        Set matches = regex.Execute(str)
        Dim m As Integer, d As Integer
        m = CInt(matches(0).SubMatches(0))
        d = CInt(matches(0).SubMatches(1))
        
        Dim y As Integer: y = Year(Date)
        ' 현재 월보다 이전이면 내년으로
        If m < Month(Date) Or (m = Month(Date) And d < Day(Date)) Then
            y = y + 1
        End If
        
        ExtractDateFromString = DateSerial(y, m, d)
        Exit Function
    End If
    
    ' 괄호 안의 날짜 (예: "1/26(1만개정도만맞춰주면됨)")
    regex.Pattern = "(\d{1,2})/(\d{1,2})\("
    If regex.Test(str) Then
        Set matches = regex.Execute(str)
        m = CInt(matches(0).SubMatches(0))
        d = CInt(matches(0).SubMatches(1))
        y = Year(Date)
        If m < Month(Date) Then y = y + 1
        ExtractDateFromString = DateSerial(y, m, d)
    End If
    
    Exit Function
    
ErrorHandler:
    ExtractDateFromString = 0
End Function

'===============================================================================
' 배정 제외 품목 체크
'===============================================================================
Private Sub CheckSkipItems(ByRef items() As ProductionItem, ByVal itemCount As Long)
    Dim i As Long
    
    For i = 1 To itemCount
        With items(i)
            .ShouldSkip = False
            .SkipReason = ""
            
            ' 품번이 없으면 제외
            If Len(.ProductCode) = 0 Then
                .ShouldSkip = True
                .SkipReason = "품번없음"
                GoTo NextItem
            End If
            
            ' 수량이 0이면 제외
            If .Quantity <= 0 Then
                .ShouldSkip = True
                .SkipReason = "수량없음"
                GoTo NextItem
            End If
            
            ' A열(이동일)이 "미정" 또는 빈칸
            If Len(.TransferStr) = 0 Or LCase(.TransferStr) = "미정" Then
                ' D열에 업체가 있어도 아직 배정 안함
                .ShouldSkip = True
                .SkipReason = "이동일미정"
                GoTo NextItem
            End If
            
            ' B열(발주)이 "미정" 또는 빈칸 -> 이건 체크 안함 (예정도 허용)
            
            ' L열(납기)이 "미정" 또는 빈칸
            If Len(.DueDateStr) = 0 Or LCase(.DueDateStr) = "미정" Then
                ' ASAP이 아니고 날짜도 없으면 제외
                If Not .IsASAP And .DueDate = 0 Then
                    .ShouldSkip = True
                    .SkipReason = "납기미정"
                    GoTo NextItem
                End If
            End If
            
            ' 이동일 파싱 실패
            If .TransferDate = 0 Then
                .ShouldSkip = True
                .SkipReason = "이동일오류"
                GoTo NextItem
            End If
        End With
NextItem:
    Next i
End Sub

'===============================================================================
' 세트품 식별
'===============================================================================
Private Sub IdentifySetItems(ByRef items() As ProductionItem, ByVal itemCount As Long)
    If itemCount = 0 Then Exit Sub
    
    Dim i As Long, j As Long
    Dim setGroupCounter As Integer: setGroupCounter = 0
    
    For i = 1 To itemCount
        If items(i).IsSetItem Then GoTo NextSetItem
        If items(i).ShouldSkip Then GoTo NextSetItem
        
        ' 세트품 공정 확인 (포장/충포장)
        If items(i).ProcessType = "포장/충포장" Then
            setGroupCounter = setGroupCounter + 1
            Dim groupID As String
            groupID = "SET" & Format(setGroupCounter, "000")
            
            items(i).IsSetItem = True
            items(i).SetGroupID = groupID
            items(i).SetLeaderRow = i
            
            Dim setTotalQty As Long
            setTotalQty = items(i).Quantity
            
            Dim leaderDueDate As Date
            leaderDueDate = items(i).DueDate
            
            Dim leaderTransferDate As Date
            leaderTransferDate = items(i).TransferDate
            
            ' 연속된 행에서 같은 고객사코드 + 충전/충포장 찾기
            For j = i + 1 To itemCount
                If items(j).ProcessType = "충전/충포장" Then
                    If items(j).ClientCode = items(i).ClientCode Then
                        items(j).IsSetItem = True
                        items(j).SetGroupID = groupID
                        items(j).SetLeaderRow = i
                        setTotalQty = setTotalQty + items(j).Quantity
                        
                        ' 1코드 납기/이동일 빈칸이면 9코드 값 사용
                        If items(j).DueDate = 0 Then
                            items(j).DueDate = leaderDueDate
                            items(j).IsASAP = items(i).IsASAP
                        End If
                        If items(j).TransferDate = 0 Then
                            items(j).TransferDate = leaderTransferDate
                        End If
                    Else
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next j
            
            ' 세트품 합산 수량 저장
            For j = i To itemCount
                If items(j).SetGroupID = groupID Then
                    items(j).SetTotalQty = setTotalQty
                End If
            Next j
        End If
NextSetItem:
    Next i
End Sub

'===============================================================================
' 순산 로직
'===============================================================================
Private Sub CalculateProductionDates(ByRef items() As ProductionItem, ByVal itemCount As Long)
    Dim i As Long
    
    For i = 1 To itemCount
        With items(i)
            If .ShouldSkip Then GoTo NextCalc
            If .TransferDate = 0 Then
                .DueCheckStatus = "날짜없음"
                GoTo NextCalc
            End If
            
            ' 생산 소요일 계산
            Dim dailyCapa As Long
            dailyCapa = GetVendorCapacity(.DesignatedVendor)
            If dailyCapa = 0 Then dailyCapa = 10000
            
            Dim qty As Long
            If .IsSetItem And .SetTotalQty > 0 Then
                qty = .SetTotalQty
            Else
                qty = .Quantity
            End If
            
            .ProductionDays = Application.WorksheetFunction.RoundUp(qty / dailyCapa, 0)
            If .ProductionDays < 1 Then .ProductionDays = 1
            
            ' 생산시작일 = 이동일 + 1
            .ProductionStartDate = AddWorkdays(.TransferDate, 1)
            
            ' 생산완료일 = 생산시작일 + 생산소요일 - 1
            .ProductionEndDate = AddWorkdays(.ProductionStartDate, .ProductionDays - 1)
            
            ' 납기 체크
            If .DueDate = 0 And Not .IsASAP Then
                .DueCheckStatus = "납기없음"
            ElseIf .IsASAP And .DueDate = 0 Then
                .DueCheckStatus = "ASAP"
            ElseIf .ProductionEndDate <= .DueDate - 1 Then
                .DueCheckStatus = "OK"
            ElseIf .ProductionEndDate = .DueDate Then
                .DueCheckStatus = "위험"
            Else
                .DueCheckStatus = "불가"
            End If
        End With
NextCalc:
    Next i
End Sub

'===============================================================================
' 외주처 Capa 조회
'===============================================================================
Private Function GetVendorCapacity(vendorName As String) As Long
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).Name = vendorName Then
            GetVendorCapacity = Vendors(i).CapaPerLine
            Exit Function
        End If
    Next i
    GetVendorCapacity = 0
End Function

'===============================================================================
' 영업일 추가
'===============================================================================
Private Function AddWorkdays(startDate As Date, days As Integer) As Date
    Dim result As Date: result = startDate
    Dim added As Integer: added = 0
    
    If days = 0 Then
        AddWorkdays = startDate
        Exit Function
    End If
    
    Do While added < days
        result = result + 1
        If Weekday(result) <> vbSunday And Weekday(result) <> vbSaturday Then
            added = added + 1
        End If
    Loop
    
    AddWorkdays = result
End Function

'===============================================================================
' 날짜 범위 파악
'===============================================================================
Private Sub GetDateRange(ByRef items() As ProductionItem, ByVal itemCount As Long, _
                         ByRef minDate As Date, ByRef maxDate As Date)
    minDate = DateSerial(2099, 12, 31)
    maxDate = DateSerial(1900, 1, 1)
    
    Dim i As Long
    For i = 1 To itemCount
        If Not items(i).ShouldSkip Then
            If items(i).TransferDate > 0 Then
                If items(i).TransferDate < minDate Then minDate = items(i).TransferDate
            End If
            If items(i).DueDate > 0 Then
                If items(i).DueDate > maxDate Then maxDate = items(i).DueDate
            End If
            If items(i).ProductionEndDate > 0 Then
                If items(i).ProductionEndDate > maxDate Then maxDate = items(i).ProductionEndDate
            End If
        End If
    Next i
    
    If minDate > maxDate Then
        minDate = Date
        maxDate = Date + 30
    End If
    
    minDate = minDate - 3
    maxDate = maxDate + 7
End Sub

'===============================================================================
' 자동 배정 실행 (납기>ASAP 우선순위 적용)
'===============================================================================
Private Sub AllocateProduction(ByRef items() As ProductionItem, ByVal itemCount As Long, _
                               ByRef results() As AllocationResult, _
                               ByVal minDate As Date, ByVal maxDate As Date)
    ReDim results(1 To itemCount)
    
    ' 배정 순서 결정: 납기 날짜 있는 것 먼저, ASAP은 나중에
    Dim sortedIdx() As Long
    ReDim sortedIdx(1 To itemCount)
    SortItemsByPriority items, itemCount, sortedIdx
    
    Dim idx As Long, i As Long
    For idx = 1 To itemCount
        i = sortedIdx(idx)
        
        results(i).Item = items(i)
        results(i).Success = False
        results(i).FailReason = ""
        
        ' 제외 품목
        If items(i).ShouldSkip Then
            results(i).FailReason = items(i).SkipReason
            GoTo NextAlloc
        End If
        
        ' D열에 지정 업체가 있는 경우
        If Len(items(i).DesignatedVendor) > 0 Then
            results(i).Item.AssignedVendor = items(i).DesignatedVendor
            
            ' 고객매칭 확인 (원래 어디로 가야 했는지)
            Dim matchedVendor As String
            matchedVendor = GetMatchedVendor(items(i))
            If Len(matchedVendor) > 0 And matchedVendor <> items(i).DesignatedVendor Then
                results(i).Item.OriginalMatchedVendor = matchedVendor
            End If
            
            ' 경고 체크
            CheckWarnings results(i), items(i)
            
            ' 슬롯 검색
            If FindSlotForVendor(results(i), items(i), maxDate) Then
                results(i).Success = True
            Else
                results(i).FailReason = "NO_SLOT_MANUAL|" & items(i).DesignatedVendor
                results(i).AlternativeVendors = GetAlternativeVendors(items(i), maxDate)
                
                ' 지정 업체 슬롯 없으면 경고 추가
                results(i).HasWarning = True
                results(i).WarningType = "NO_SLOT"
                results(i).WarningMessage = items(i).DesignatedVendor & " 슬롯 없음"
            End If
        Else
            ' 자동 배정
            ' 1. 고객매칭 업체 시도
            matchedVendor = GetMatchedVendor(items(i))
            If Len(matchedVendor) > 0 Then
                results(i).Item.OriginalMatchedVendor = matchedVendor
                results(i).Item.AssignedVendor = matchedVendor
                
                If FindSlotForVendor(results(i), items(i), maxDate) Then
                    results(i).Success = True
                    GoTo NextAlloc
                End If
            End If
            
            ' 2. 우선순위 순 검색
            Dim vendorIdx As Integer
            For vendorIdx = 1 To VendorCount
                If CanHandleItem(Vendors(vendorIdx), items(i)) Then
                    results(i).Item.AssignedVendor = Vendors(vendorIdx).Name
                    If FindSlotForVendor(results(i), items(i), maxDate) Then
                        results(i).Success = True
                        
                        ' 고객매칭과 다른 업체로 배정된 경우
                        If Len(matchedVendor) > 0 And Vendors(vendorIdx).Name <> matchedVendor Then
                            results(i).Item.VendorChanged = True
                            results(i).HasWarning = True
                            results(i).WarningType = "VENDOR_CHANGED"
                            results(i).WarningMessage = matchedVendor & " → " & Vendors(vendorIdx).Name & " (capa 부족)"
                        End If
                        GoTo NextAlloc
                    End If
                End If
            Next vendorIdx
            
            ' 3. 배정 실패
            results(i).FailReason = "NO_AVAILABLE_SLOT"
            results(i).AlternativeVendors = GetAlternativeVendors(items(i), maxDate)
        End If
        
NextAlloc:
        ' 최초배정업체 기록
        If Len(results(i).Item.OriginalVendor) = 0 Then
            results(i).Item.OriginalVendor = results(i).Item.AssignedVendor
        End If
    Next idx
End Sub

'===============================================================================
' 배정 우선순위 정렬 (납기 날짜 > ASAP)
'===============================================================================
Private Sub SortItemsByPriority(ByRef items() As ProductionItem, ByVal itemCount As Long, ByRef sortedIdx() As Long)
    Dim i As Long, j As Long, tempIdx As Long
    
    ' 초기화
    For i = 1 To itemCount
        sortedIdx(i) = i
    Next i
    
    ' 버블 정렬 (납기 날짜 있는 것이 먼저, 같으면 이동일 빠른 것)
    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            Dim swap As Boolean: swap = False
            
            Dim a As Long, b As Long
            a = sortedIdx(i)
            b = sortedIdx(j)
            
            ' Skip 제외 품목은 맨 뒤로
            If items(a).ShouldSkip And Not items(b).ShouldSkip Then
                swap = True
            ElseIf Not items(a).ShouldSkip And Not items(b).ShouldSkip Then
                ' ASAP은 뒤로
                If items(a).IsASAP And Not items(b).IsASAP Then
                    If items(b).DueDate > 0 Then swap = True
                ElseIf Not items(a).IsASAP And items(b).IsASAP Then
                    ' a가 납기 있으면 그대로
                ElseIf Not items(a).IsASAP And Not items(b).IsASAP Then
                    ' 둘 다 납기 있으면 이동일 빠른 순
                    If items(a).TransferDate > items(b).TransferDate Then
                        swap = True
                    End If
                End If
            End If
            
            If swap Then
                tempIdx = sortedIdx(i)
                sortedIdx(i) = sortedIdx(j)
                sortedIdx(j) = tempIdx
            End If
        Next j
    Next i
End Sub

'===============================================================================
' 매칭 업체 조회
'===============================================================================
Private Function GetMatchedVendor(ByRef Item As ProductionItem) As String
    GetMatchedVendor = ""
    
    If Len(Item.ClientCode) = 0 Then Exit Function
    If Not ClientMappings.Exists(Item.ClientCode) Then Exit Function
    
    Dim vendorName As String
    vendorName = ClientMappings(Item.ClientCode)
    
    ' 해당 업체가 처리 가능한지 확인
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).Name = vendorName Then
            If CanHandleItem(Vendors(i), Item) Then
                GetMatchedVendor = vendorName
            End If
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' 품목 처리 가능 여부
'===============================================================================
Private Function CanHandleItem(ByRef vendor As VendorInfo, ByRef Item As ProductionItem) As Boolean
    ' 특수공정 체크 (특이사항에서 확인)
    If InStr(1, Item.SpecialProcess, "수축", vbTextCompare) > 0 Then
        If InStr(1, vendor.Capabilities, "shrink", vbTextCompare) = 0 Then
            CanHandleItem = False
            Exit Function
        End If
    End If
    If InStr(1, Item.SpecialProcess, "교반", vbTextCompare) > 0 Then
        If InStr(1, vendor.Capabilities, "mixing", vbTextCompare) = 0 Then
            CanHandleItem = False
            Exit Function
        End If
    End If
    If InStr(1, Item.SpecialProcess, "고주파", vbTextCompare) > 0 Then
        If InStr(1, vendor.Capabilities, "highFrequency", vbTextCompare) = 0 Then
            CanHandleItem = False
            Exit Function
        End If
    End If
    
    CanHandleItem = True
End Function

'===============================================================================
' 경고 체크
'===============================================================================
Private Sub CheckWarnings(ByRef result As AllocationResult, ByRef Item As ProductionItem)
    result.HasWarning = False
    
    ' 납기 체크
    If Item.DueCheckStatus = "불가" Then
        result.HasWarning = True
        result.WarningType = "DUE_OVER"
        result.WarningMessage = "납기초과 (완료예정: " & Format(Item.ProductionEndDate, "m/d") & _
                                ", 납기: " & Format(Item.DueDate, "m/d") & ")"
        Exit Sub
    End If
    
    If Item.DueCheckStatus = "위험" Then
        result.HasWarning = True
        result.WarningType = "DUE_RISK"
        result.WarningMessage = "당일직납 (완료예정: " & Format(Item.ProductionEndDate, "m/d") & ")"
        Exit Sub
    End If
    
    ' capa 체크 (간단히)
    Dim vendorIdx As Integer
    vendorIdx = GetVendorIndex(Item.DesignatedVendor)
    If vendorIdx > 0 Then
        Dim dailyCapa As Long
        dailyCapa = Vendors(vendorIdx).CapaPerLine * Vendors(vendorIdx).LineCount
        If Item.Quantity > dailyCapa * 3 Then  ' 3일치 이상이면 경고
            result.HasWarning = True
            result.WarningType = "CAPA_WARNING"
            result.WarningMessage = "대량 주문 (일 capa: " & Format(dailyCapa, "#,##0") & ")"
        End If
    End If
End Sub

'===============================================================================
' 업체 인덱스 조회
'===============================================================================
Private Function GetVendorIndex(vendorName As String) As Integer
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).Name = vendorName Then
            GetVendorIndex = i
            Exit Function
        End If
    Next i
    GetVendorIndex = 0
End Function

'===============================================================================
' 슬롯 검색
'===============================================================================
Private Function FindSlotForVendor(ByRef result As AllocationResult, _
                                    ByRef Item As ProductionItem, _
                                    maxDate As Date) As Boolean
    FindSlotForVendor = False
    
    Dim vendorIdx As Integer
    vendorIdx = GetVendorIndex(result.Item.AssignedVendor)
    If vendorIdx = 0 Then Exit Function
    
    Dim startDate As Date
    startDate = Item.TransferDate + 1
    If startDate < Date Then startDate = Date
    
    Dim daysNeeded As Integer
    daysNeeded = Item.ProductionDays
    If daysNeeded < 1 Then daysNeeded = 1
    
    Dim line As Integer
    Dim searchDate As Date: searchDate = startDate
    
    Do While searchDate <= maxDate
        If Weekday(searchDate) = vbSunday Or Weekday(searchDate) = vbSaturday Then
            searchDate = searchDate + 1
            GoTo ContinueSearch
        End If
        
        For line = 1 To Vendors(vendorIdx).LineCount
            If CanPlaceOnLine(Vendors(vendorIdx).Name, line, searchDate, daysNeeded, maxDate) Then
                ReserveLine Vendors(vendorIdx).Name, line, searchDate, daysNeeded
                result.Item.AssignedLine = line
                result.Item.ProductionStartDate = searchDate
                result.Item.ProductionEndDate = AddWorkdays(searchDate, daysNeeded - 1)
                
                ' 납기 재체크
                If Item.DueDate > 0 Then
                    If result.Item.ProductionEndDate > Item.DueDate Then
                        result.Item.DueCheckStatus = "불가"
                    ElseIf result.Item.ProductionEndDate = Item.DueDate Then
                        result.Item.DueCheckStatus = "위험"
                    Else
                        result.Item.DueCheckStatus = "OK"
                    End If
                End If
                
                FindSlotForVendor = True
                Exit Function
            End If
        Next line
        
        searchDate = searchDate + 1
ContinueSearch:
    Loop
End Function

'===============================================================================
' 슬롯 배치 가능 여부
'===============================================================================
Private Function CanPlaceOnLine(vendorName As String, line As Integer, _
                                startDate As Date, daysNeeded As Integer, _
                                maxDate As Date) As Boolean
    CanPlaceOnLine = True
    
    Dim checkDate As Date: checkDate = startDate
    Dim placedDays As Integer: placedDays = 0
    
    Do While placedDays < daysNeeded
        If checkDate > maxDate Then
            CanPlaceOnLine = False
            Exit Function
        End If
        
        If Weekday(checkDate) = vbSunday Or Weekday(checkDate) = vbSaturday Then
            checkDate = checkDate + 1
            GoTo ContinueCheck
        End If
        
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
End Function

'===============================================================================
' 슬롯 예약
'===============================================================================
Private Sub ReserveLine(vendorName As String, line As Integer, _
                        startDate As Date, daysNeeded As Integer)
    Dim checkDate As Date: checkDate = startDate
    Dim placedDays As Integer: placedDays = 0
    
    Do While placedDays < daysNeeded
        If Weekday(checkDate) = vbSunday Or Weekday(checkDate) = vbSaturday Then
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
' 대안 업체 리스트 생성
'===============================================================================
Private Function GetAlternativeVendors(ByRef Item As ProductionItem, maxDate As Date) As String
    Dim alternatives As String: alternatives = ""
    Dim i As Integer
    
    For i = 1 To VendorCount
        If CanHandleItem(Vendors(i), Item) Then
            Dim earliestDate As Date
            earliestDate = FindEarliestSlot(Vendors(i).Name, Item.ProductionDays, maxDate)
            If earliestDate > 0 Then
                If Len(alternatives) > 0 Then alternatives = alternatives & "; "
                alternatives = alternatives & Vendors(i).Name & "(" & Format(earliestDate, "m/d") & ")"
            End If
        End If
    Next i
    
    GetAlternativeVendors = alternatives
End Function

'===============================================================================
' 가장 빠른 슬롯 찾기
'===============================================================================
Private Function FindEarliestSlot(vendorName As String, daysNeeded As Integer, maxDate As Date) As Date
    FindEarliestSlot = 0
    
    Dim vendorIdx As Integer
    vendorIdx = GetVendorIndex(vendorName)
    If vendorIdx = 0 Then Exit Function
    
    Dim searchDate As Date: searchDate = Date
    
    Do While searchDate <= maxDate
        If Weekday(searchDate) = vbSunday Or Weekday(searchDate) = vbSaturday Then
            searchDate = searchDate + 1
            GoTo ContinueFind
        End If
        
        Dim line As Integer
        For line = 1 To Vendors(vendorIdx).LineCount
            If CanPlaceOnLine(vendorName, line, searchDate, daysNeeded, maxDate) Then
                FindEarliestSlot = searchDate
                Exit Function
            End If
        Next line
        
        searchDate = searchDate + 1
ContinueFind:
    Loop
End Function

'===============================================================================
' Raw 데이터 시트에 결과 기록
'===============================================================================
Private Sub WriteResultsToRaw(ByRef items() As ProductionItem, ByVal itemCount As Long, _
                              ByRef results() As AllocationResult)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_RAW)
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets(1)
    On Error GoTo 0
    
    ' 자동 생성 컬럼 헤더 추가
    ws.Cells(1, RAW_COL_CLIENT).Value = "고객사코드"
    ws.Cells(1, RAW_COL_ASSIGNED).Value = "배정업체"
    ws.Cells(1, RAW_COL_PRODSTART).Value = "생산시작일"
    ws.Cells(1, RAW_COL_PRODEND).Value = "생산완료일"
    ws.Cells(1, RAW_COL_DUECHECK).Value = "납기준수"
    ws.Cells(1, RAW_COL_ORIGINAL).Value = "최초배정업체"
    ws.Cells(1, RAW_COL_CHANGED).Value = "업체변경"
    ws.Cells(1, RAW_COL_WARNING).Value = "경고"
    
    Dim i As Long
    For i = 1 To itemCount
        With results(i).Item
            If .RowNum > 0 Then
                ws.Cells(.RowNum, RAW_COL_CLIENT).Value = .ClientCode
                ws.Cells(.RowNum, RAW_COL_ASSIGNED).Value = .AssignedVendor
                
                If .ProductionStartDate > 0 Then
                    ws.Cells(.RowNum, RAW_COL_PRODSTART).Value = .ProductionStartDate
                    ws.Cells(.RowNum, RAW_COL_PRODSTART).NumberFormat = "m/d"
                End If
                
                If .ProductionEndDate > 0 Then
                    ws.Cells(.RowNum, RAW_COL_PRODEND).Value = .ProductionEndDate
                    ws.Cells(.RowNum, RAW_COL_PRODEND).NumberFormat = "m/d"
                End If
                
                ws.Cells(.RowNum, RAW_COL_DUECHECK).Value = .DueCheckStatus
                ws.Cells(.RowNum, RAW_COL_ORIGINAL).Value = .OriginalVendor
                ws.Cells(.RowNum, RAW_COL_CHANGED).Value = IIf(.VendorChanged, "변경됨", "")
                
                ' 경고 메시지
                If results(i).HasWarning Then
                    ws.Cells(.RowNum, RAW_COL_WARNING).Value = results(i).WarningMessage
                ElseIf .ShouldSkip Then
                    ws.Cells(.RowNum, RAW_COL_WARNING).Value = .SkipReason
                End If
            End If
        End With
    Next i
    
    ws.Columns(RAW_COL_CLIENT & ":" & RAW_COL_WARNING).AutoFit
End Sub

'===============================================================================
' 생산계획표 생성 (조회창 포함)
'===============================================================================
Private Sub CreateProductionPlanView(ByRef results() As AllocationResult, _
                                     ByVal itemCount As Long, _
                                     ByVal minDate As Date, ByVal maxDate As Date)
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_OUTPUT)
    
    ws.Cells.Font.Name = "맑은 고딕"
    ws.Cells.Font.Size = 9
    
    On Error Resume Next
    ws.Cells.ClearComments
    On Error GoTo 0
    
    Dim currentRow As Long: currentRow = 1
    
    ' ========== 조회창 영역 (상단) ==========
    ws.Cells(currentRow, 1).Value = "품목 조회"
    With ws.Cells(currentRow, 1)
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(59, 130, 246)
    End With
    currentRow = currentRow + 1
    
    ws.Cells(currentRow, 1).Value = "품번/품목명:"
    ws.Cells(currentRow, 1).Font.Bold = True
    ws.Cells(currentRow, 2).Value = ""
    ws.Cells(currentRow, 2).Interior.Color = RGB(255, 255, 224)  ' 연한 노랑 (입력칸)
    With ws.Cells(currentRow, 2).Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(59, 130, 246)
    End With
    ws.Columns(2).ColumnWidth = 20
    
    ws.Cells(currentRow, 3).Value = "← 검색어 입력 후 Ctrl+F로 찾기"
    ws.Cells(currentRow, 3).Font.Color = RGB(107, 114, 128)
    ws.Cells(currentRow, 3).Font.Italic = True
    
    currentRow = currentRow + 2
    
    ' ========== 타이틀 ==========
    ws.Cells(currentRow, 1).Value = "외주 생산계획표"
    With ws.Cells(currentRow, 1)
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = RGB(30, 41, 59)
    End With
    currentRow = currentRow + 1
    
    ws.Cells(currentRow, 1).Value = Format(minDate, "yyyy년 m월 d일") & " ~ " & _
                                     Format(maxDate, "yyyy년 m월 d일") & "  |  " & _
                                     "Updated: " & Format(Now, "yyyy-mm-dd hh:mm")
    With ws.Cells(currentRow, 1)
        .Font.Size = 11
        .Font.Color = RGB(100, 116, 139)
    End With
    currentRow = currentRow + 1
    
    ' 범례
    ws.Cells(currentRow, 1).Value = "범례:"
    ws.Cells(currentRow, 1).Font.Bold = True
    
    ws.Cells(currentRow, 2).Value = "경고"
    ws.Cells(currentRow, 2).Interior.Color = GetPastelRed()
    
    ws.Cells(currentRow, 3).Value = "업체변경"
    ws.Cells(currentRow, 3).Interior.Color = GetPastelOrange()
    
    ws.Cells(currentRow, 4).Value = "위험"
    ws.Cells(currentRow, 4).Interior.Color = GetPastelYellow()
    
    ws.Cells(currentRow, 5).Value = "정상"
    ws.Cells(currentRow, 5).Interior.Color = GetPastelGreen()
    
    currentRow = currentRow + 2
    
    ' ========== 캘린더 헤더 ==========
    Dim totalDays As Long
    totalDays = DateDiff("d", minDate, maxDate) + 1
    
    Dim headerRow As Long: headerRow = currentRow
    
    ws.Cells(headerRow, 1).Value = "외주처 / 라인"
    With ws.Cells(headerRow, 1)
        .Interior.Color = RGB(30, 41, 59)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Columns(1).ColumnWidth = 14
    
    ' 날짜 헤더
    Dim d As Long
    Dim currentDate As Date
    For d = 1 To totalDays
        currentDate = minDate + d - 1
        Dim dow As Integer: dow = Weekday(currentDate)
        Dim col As Long: col = d + 1
        
        Dim headerText As String
        If Day(currentDate) = 1 Or d = 1 Then
            headerText = Format(currentDate, "m월") & vbLf & Day(currentDate) & vbLf & GetDayName(dow)
        Else
            headerText = Day(currentDate) & vbLf & GetDayName(dow)
        End If
        
        ws.Cells(headerRow, col).Value = headerText
        With ws.Cells(headerRow, col)
            .WrapText = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 9
            
            If dow = vbSunday Or dow = vbSaturday Then
                .Interior.Color = RGB(254, 226, 226)
                .Font.Color = RGB(185, 28, 28)
            ElseIf currentDate = Date Then
                .Interior.Color = RGB(187, 247, 208)
                .Font.Color = RGB(22, 101, 52)
            Else
                .Interior.Color = RGB(241, 245, 249)
                .Font.Color = RGB(51, 65, 85)
            End If
        End With
        
        ws.Columns(col).ColumnWidth = 13
    Next d
    
    ws.Rows(headerRow).RowHeight = 42
    
    ' ========== 업체별 행 ==========
    currentRow = headerRow + 1
    Dim v As Integer
    
    For v = 1 To VendorCount
        Dim vendorName As String: vendorName = Vendors(v).Name
        
        ' 해당 업체에 배정된 품목 있는지 확인
        Dim hasItems As Boolean: hasItems = False
        Dim i As Long
        For i = 1 To itemCount
            If results(i).Item.AssignedVendor = vendorName And results(i).Success Then
                hasItems = True
                Exit For
            End If
        Next i
        
        If Not hasItems Then GoTo NextVendor
        
        ' 업체 헤더 행
        ws.Cells(currentRow, 1).Value = vendorName
        With ws.Cells(currentRow, 1)
            .Font.Bold = True
            .Font.Size = 12
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = GetVendorColor(vendorName)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        For d = 1 To totalDays
            ws.Cells(currentRow, d + 1).Interior.Color = GetVendorLightColor(vendorName)
        Next d
        ws.Rows(currentRow).RowHeight = 26
        currentRow = currentRow + 1
        
        ' 라인별 행
        Dim line As Integer
        For line = 1 To Vendors(v).LineCount
            ws.Cells(currentRow, 1).Value = "Line " & line
            With ws.Cells(currentRow, 1)
                .Font.Color = RGB(51, 65, 85)
                .Font.Size = 10
                .Font.Bold = True
                .Interior.Color = RGB(248, 250, 252)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            ' 날짜별 배경
            For d = 1 To totalDays
                currentDate = minDate + d - 1
                dow = Weekday(currentDate)
                col = d + 1
                
                If dow = vbSunday Or dow = vbSaturday Then
                    ws.Cells(currentRow, col).Interior.Color = RGB(254, 242, 242)
                Else
                    ws.Cells(currentRow, col).Interior.Color = RGB(255, 255, 255)
                End If
            Next d
            
            ' 품목 표시
            For i = 1 To itemCount
                If results(i).Item.AssignedVendor = vendorName And _
                   results(i).Item.AssignedLine = line And _
                   results(i).Success Then
                    PlaceItemOnCalendar ws, results(i), currentRow, minDate, totalDays
                End If
            Next i
            
            ws.Rows(currentRow).RowHeight = 58
            currentRow = currentRow + 1
        Next line
        
NextVendor:
    Next v
    
    ' 테두리
    If currentRow > headerRow + 1 Then
        ApplyTableBorders ws, headerRow, currentRow - 1, totalDays + 1
    End If
    
    ' ========== 배정불가 섹션 ==========
    currentRow = currentRow + 2
    currentRow = CreateUnassignedSection(ws, results, itemCount, currentRow)
    
    ' ========== 배정 제외(미정) 섹션 ==========
    currentRow = currentRow + 2
    currentRow = CreateSkippedSection(ws, results, itemCount, currentRow)
    
    ' 틀 고정 (오류 방지)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(headerRow + 1, 2).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

'===============================================================================
' 요일명 반환
'===============================================================================
Private Function GetDayName(dow As Integer) As String
    Select Case dow
        Case vbSunday: GetDayName = "일"
        Case vbMonday: GetDayName = "월"
        Case vbTuesday: GetDayName = "화"
        Case vbWednesday: GetDayName = "수"
        Case vbThursday: GetDayName = "목"
        Case vbFriday: GetDayName = "금"
        Case vbSaturday: GetDayName = "토"
    End Select
End Function

'===============================================================================
' 품목을 캘린더에 배치 (파스텔톤 적용)
'===============================================================================
Private Sub PlaceItemOnCalendar(ByRef ws As Worksheet, ByRef result As AllocationResult, _
                                ByVal rowNum As Long, ByVal minDate As Date, _
                                ByVal totalDays As Long)
    
    If result.Item.ProductionDays = 0 Then Exit Sub
    If result.Item.ProductionStartDate = 0 Then Exit Sub
    
    Dim dailyCapacity As Long
    dailyCapacity = GetVendorCapacity(result.Item.AssignedVendor)
    If dailyCapacity = 0 Then dailyCapacity = 10000
    
    Dim remainingQty As Long: remainingQty = result.Item.Quantity
    Dim currentDate As Date: currentDate = result.Item.ProductionStartDate
    Dim dayNum As Integer: dayNum = 1
    
    Do While remainingQty > 0
        Dim dayOffset As Long
        dayOffset = DateDiff("d", minDate, currentDate)
        If dayOffset < 0 Or dayOffset >= totalDays Then
            currentDate = currentDate + 1
            GoTo ContinuePlace
        End If
        
        If Weekday(currentDate) = vbSunday Or Weekday(currentDate) = vbSaturday Then
            currentDate = currentDate + 1
            GoTo ContinuePlace
        End If
        
        Dim col As Long: col = dayOffset + 2
        Dim dailyQty As Long: dailyQty = remainingQty
        If dailyQty > dailyCapacity Then dailyQty = dailyCapacity
        
        ' 셀 내용
        Dim prodCode As String: prodCode = result.Item.ProductCode
        If Len(prodCode) > 10 Then prodCode = Left(prodCode, 10)
        
        Dim prodName As String: prodName = result.Item.ProductName
        If Len(prodName) > 8 Then prodName = Left(prodName, 7) & ".."
        
        Dim cellText As String
        cellText = prodCode & vbLf & prodName & vbLf & Format(dailyQty, "#,##0")
        
        If result.Item.ProductionDays > 1 Then
            cellText = cellText & " [" & dayNum & "/" & result.Item.ProductionDays & "]"
        End If
        
        ws.Cells(rowNum, col).Value = cellText
        
        ' 셀 스타일 (파스텔톤)
        Dim cellBgColor As Long
        Dim fontColor As Long
        
        If result.HasWarning Then
            Select Case result.WarningType
                Case "DUE_OVER", "NO_SLOT", "PROCESS_UNABLE"
                    ' 경고 - 파스텔 빨간색
                    cellBgColor = GetPastelRed()
                    fontColor = RGB(127, 29, 29)
                Case "DUE_RISK"
                    ' 위험 - 파스텔 노란색
                    cellBgColor = GetPastelYellow()
                    fontColor = RGB(133, 77, 14)
                Case "VENDOR_CHANGED", "CAPA_WARNING"
                    ' 업체변경/capa경고 - 파스텔 주황색
                    cellBgColor = GetPastelOrange()
                    fontColor = RGB(146, 64, 14)
                Case Else
                    cellBgColor = GetVendorLightColor(result.Item.AssignedVendor)
                    fontColor = RGB(30, 41, 59)
            End Select
        ElseIf result.Item.VendorChanged Then
            ' 업체 변경 - 파스텔 주황색
            cellBgColor = GetPastelOrange()
            fontColor = RGB(146, 64, 14)
        Else
            ' 정상 - 업체 색상
            cellBgColor = GetVendorLightColor(result.Item.AssignedVendor)
            fontColor = RGB(30, 41, 59)
        End If
        
        With ws.Cells(rowNum, col)
            .Interior.Color = cellBgColor
            .Font.Size = 8
            .Font.Color = fontColor
            .Font.Bold = False
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .WrapText = True
            
            ' 테두리
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeLeft).Color = GetVendorColor(result.Item.AssignedVendor)
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeRight).Color = GetVendorColor(result.Item.AssignedVendor)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeTop).Color = GetVendorColor(result.Item.AssignedVendor)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeBottom).Color = GetVendorColor(result.Item.AssignedVendor)
        End With
        
        ' 메모 추가
        Dim commentText As String
        commentText = ""
        
        If result.HasWarning Then
            commentText = "⚠️ " & result.WarningMessage & vbLf & "────────────────" & vbLf
        End If
        
        If result.Item.VendorChanged And Len(result.Item.OriginalMatchedVendor) > 0 Then
            commentText = commentText & "📌 원래 매칭: " & result.Item.OriginalMatchedVendor & _
                          " → " & result.Item.AssignedVendor & vbLf & "────────────────" & vbLf
        End If
        
        commentText = commentText & "품목코드: " & result.Item.ProductCode & vbLf
        commentText = commentText & "품목명: " & result.Item.ProductName & vbLf
        commentText = commentText & "────────────────" & vbLf
        commentText = commentText & "총 수량: " & Format(result.Item.Quantity, "#,##0") & vbLf
        commentText = commentText & "금일 수량: " & Format(dailyQty, "#,##0") & vbLf
        commentText = commentText & "────────────────" & vbLf
        commentText = commentText & "이동일: " & Format(result.Item.TransferDate, "m/d") & vbLf
        If result.Item.DueDate > 0 Then
            commentText = commentText & "납기: " & Format(result.Item.DueDate, "m/d") & vbLf
        ElseIf result.Item.IsASAP Then
            commentText = commentText & "납기: ASAP" & vbLf
        End If
        commentText = commentText & "납기준수: " & result.Item.DueCheckStatus & vbLf
        commentText = commentText & "외주처: " & result.Item.AssignedVendor & " Line " & result.Item.AssignedLine
        
        ' 메모 추가 (Excel 버전 호환)
        On Error Resume Next
        ws.Cells(rowNum, col).Comment.Delete
        ws.Cells(rowNum, col).AddComment commentText
        If Not ws.Cells(rowNum, col).Comment Is Nothing Then
            ws.Cells(rowNum, col).Comment.Shape.TextFrame.AutoSize = True
        End If
        On Error GoTo 0
        
        remainingQty = remainingQty - dailyQty
        dayNum = dayNum + 1
        currentDate = currentDate + 1
ContinuePlace:
    Loop
End Sub

'===============================================================================
' 테이블 테두리 적용
'===============================================================================
Private Sub ApplyTableBorders(ws As Worksheet, startRow As Long, endRow As Long, endCol As Long)
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, endCol))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeLeft).Color = RGB(30, 41, 59)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeRight).Color = RGB(30, 41, 59)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeTop).Color = RGB(30, 41, 59)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeBottom).Color = RGB(30, 41, 59)
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideVertical).Color = RGB(226, 232, 240)
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideHorizontal).Color = RGB(226, 232, 240)
    End With
End Sub

'===============================================================================
' 배정불가 섹션
'===============================================================================
Private Function CreateUnassignedSection(ws As Worksheet, ByRef results() As AllocationResult, _
                                         ByVal itemCount As Long, ByVal startRow As Long) As Long
    Dim hasUnassigned As Boolean: hasUnassigned = False
    Dim i As Long
    For i = 1 To itemCount
        If Not results(i).Success And Not results(i).Item.ShouldSkip And Len(results(i).FailReason) > 0 Then
            hasUnassigned = True
            Exit For
        End If
    Next i
    
    If Not hasUnassigned Then
        CreateUnassignedSection = startRow
        Exit Function
    End If
    
    ' 타이틀
    ws.Cells(startRow, 1).Value = "⚠️ 배정불가 항목"
    With ws.Cells(startRow, 1)
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(185, 28, 28)
    End With
    startRow = startRow + 1
    
    ' 헤더
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 8)).Value = _
        Array("품번", "품목", "수량", "이동일", "납기", "지정업체", "사유", "대안업체")
    
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 8))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(185, 28, 28)
        .HorizontalAlignment = xlCenter
    End With
    startRow = startRow + 1
    
    ' 데이터
    For i = 1 To itemCount
        If Not results(i).Success And Not results(i).Item.ShouldSkip And Len(results(i).FailReason) > 0 Then
            ws.Cells(startRow, 1).Value = results(i).Item.ProductCode
            ws.Cells(startRow, 2).Value = results(i).Item.ProductName
            ws.Cells(startRow, 3).Value = results(i).Item.Quantity
            ws.Cells(startRow, 3).NumberFormat = "#,##0"
            ws.Cells(startRow, 4).Value = results(i).Item.TransferDate
            ws.Cells(startRow, 4).NumberFormat = "m/d"
            If results(i).Item.DueDate > 0 Then
                ws.Cells(startRow, 5).Value = results(i).Item.DueDate
                ws.Cells(startRow, 5).NumberFormat = "m/d"
            ElseIf results(i).Item.IsASAP Then
                ws.Cells(startRow, 5).Value = "ASAP"
            End If
            ws.Cells(startRow, 6).Value = results(i).Item.DesignatedVendor
            
            Dim reasonText As String
            Select Case Split(results(i).FailReason, "|")(0)
                Case "NO_SLOT_MANUAL"
                    reasonText = "지정업체 슬롯없음"
                Case "NO_AVAILABLE_SLOT"
                    reasonText = "전체 슬롯없음"
                Case Else
                    reasonText = results(i).FailReason
            End Select
            ws.Cells(startRow, 7).Value = reasonText
            ws.Cells(startRow, 8).Value = results(i).AlternativeVendors
            
            ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 8)).Interior.Color = GetPastelRed()
            startRow = startRow + 1
        End If
    Next i
    
    ws.Columns(7).ColumnWidth = 18
    ws.Columns(8).ColumnWidth = 40
    
    CreateUnassignedSection = startRow
End Function

'===============================================================================
' 배정 제외(미정) 섹션
'===============================================================================
Private Function CreateSkippedSection(ws As Worksheet, ByRef results() As AllocationResult, _
                                      ByVal itemCount As Long, ByVal startRow As Long) As Long
    Dim hasSkipped As Boolean: hasSkipped = False
    Dim i As Long
    For i = 1 To itemCount
        If results(i).Item.ShouldSkip Then
            hasSkipped = True
            Exit For
        End If
    Next i
    
    If Not hasSkipped Then
        CreateSkippedSection = startRow
        Exit Function
    End If
    
    ' 타이틀
    ws.Cells(startRow, 1).Value = "⏸️ 배정 제외 항목 (미정/빈칸)"
    With ws.Cells(startRow, 1)
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(107, 114, 128)
    End With
    startRow = startRow + 1
    
    ' 헤더
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 7)).Value = _
        Array("품번", "품목", "수량", "이동일", "납기", "지정업체", "제외사유")
    
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 7))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(107, 114, 128)
        .HorizontalAlignment = xlCenter
    End With
    startRow = startRow + 1
    
    ' 데이터
    For i = 1 To itemCount
        If results(i).Item.ShouldSkip Then
            ws.Cells(startRow, 1).Value = results(i).Item.ProductCode
            ws.Cells(startRow, 2).Value = results(i).Item.ProductName
            ws.Cells(startRow, 3).Value = results(i).Item.Quantity
            ws.Cells(startRow, 3).NumberFormat = "#,##0"
            ws.Cells(startRow, 4).Value = results(i).Item.TransferStr
            ws.Cells(startRow, 5).Value = results(i).Item.DueDateStr
            ws.Cells(startRow, 6).Value = results(i).Item.DesignatedVendor
            ws.Cells(startRow, 7).Value = results(i).Item.SkipReason
            
            ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 7)).Interior.Color = RGB(243, 244, 246)
            startRow = startRow + 1
        End If
    Next i
    
    CreateSkippedSection = startRow
End Function

'===============================================================================
' 시트 생성/초기화
'===============================================================================
Private Function CreateOrClearSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
        ws.Cells.ClearComments
    End If
    
    Set CreateOrClearSheet = ws
End Function

'===============================================================================
' 전체 리포트 생성
'===============================================================================
Public Sub 전체리포트생성()
    Call 자동배정실행
    
    MsgBox "전체 리포트 생성 완료!" & vbCrLf & vbCrLf & _
           "생성된 시트:" & vbCrLf & _
           "- 생산계획표 (전체 캘린더 뷰)", vbInformation, "완료"
End Sub
