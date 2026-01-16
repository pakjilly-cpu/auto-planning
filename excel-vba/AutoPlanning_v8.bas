'===============================================================================
' 외주처 생산계획 자동화 시스템 - VBA 매크로 v8.0
'===============================================================================
' [v8.0 주요 기능]
' - FRAMEWORK.md 기반 완전 재설계
' - 역산 로직: 납기 → 생산일 → 이동일 자동 계산
' - 상태값: 신규/보류/배정완료/발주완료
' - 긴급도: 일반/급함/긴급 (Capa 범위 결정)
' - D+2 확정 구간 / D+3+ 조정 가능 구간
' - 세트품 묶음 처리 (80% capa, 합산)
' - 이동일 변경 시 기존 업체 우선 유지
' - 배정불가 시 대안 리스트 제공
' - 세련된 UI/가독성 향상
'===============================================================================

Option Explicit

'===============================================================================
' 상수 정의
'===============================================================================
Public Const SHEET_RAW As String = "Raw데이터"
Public Const SHEET_VENDORS As String = "설정_외주처"
Public Const SHEET_MAPPING As String = "설정_고객매칭"
Public Const SHEET_OUTPUT As String = "생산계획표"
Public Const SHEET_VENDOR_PLAN As String = "업체별계획"

' Raw 데이터 컬럼 (FRAMEWORK.md 기준)
Public Const RAW_COL_CODE As Integer = 1          ' A: 품번
Public Const RAW_COL_NAME As Integer = 2          ' B: 품명
Public Const RAW_COL_QTY As Integer = 3           ' C: 수량
Public Const RAW_COL_DUEDATE As Integer = 4       ' D: 납기
Public Const RAW_COL_PROCESS As Integer = 5       ' E: 공정명
Public Const RAW_COL_SPECIAL As Integer = 6       ' F: 특수공정
Public Const RAW_COL_VENDOR As Integer = 7        ' G: 지정업체
Public Const RAW_COL_URGENCY As Integer = 8       ' H: 긴급도
Public Const RAW_COL_STATUS As Integer = 9        ' I: 상태
Public Const RAW_COL_REMARKS As Integer = 10      ' J: 비고
Public Const RAW_COL_CLIENT As Integer = 11       ' K: 고객사코드 (자동)
Public Const RAW_COL_ASSIGNED As Integer = 12     ' L: 배정업체 (자동)
Public Const RAW_COL_TRANSFER As Integer = 13     ' M: 이동일 (자동)
Public Const RAW_COL_PRODDATE As Integer = 14     ' N: 생산일 (자동)
Public Const RAW_COL_ORIGINAL As Integer = 15     ' O: 최초배정업체 (자동)
Public Const RAW_COL_CHANGED As Integer = 16      ' P: 업체변경여부 (자동)

' Capa 설정 (시간 단위)
Public Const HOURS_BASIC As Integer = 8           ' 기본 근무시간
Public Const HOURS_OVERTIME As Integer = 3        ' 잔업 시간
Public Const HOURS_SATURDAY As Integer = 8        ' 토요일 특근

' 세트품 capa 계수
Public Const SET_CAPA_FACTOR As Double = 0.8      ' 세트품은 80% capa

' 확정 구간
Public Const FIXED_DAYS As Integer = 2            ' D+2 이내는 확정

'===============================================================================
' 타입 정의
'===============================================================================
Public Type VendorInfo
    ID As String
    Name As String
    LineCount As Integer
    CapaPerLine As Long           ' 라인당 일일 생산량 (8시간 기준)
    Capabilities As String        ' 가능 공정 (콤마 구분)
    MonthlyTarget As Long
    Priority As Integer
End Type

Public Type ProductionItem
    RowNum As Long
    ProductCode As String
    ProductName As String
    Quantity As Long
    DueDate As Date               ' 납기
    ProcessType As String         ' 공정명 (포장/충전/충포장/충전충포장/포장충포장)
    SpecialProcess As String      ' 특수공정 (수축/교반/고주파)
    DesignatedVendor As String    ' 지정업체
    Urgency As String             ' 긴급도 (일반/급함/긴급)
    Status As String              ' 상태 (신규/보류/배정완료/발주완료)
    Remarks As String
    ClientCode As String          ' 고객사코드 (품번에서 추출)
    
    ' 자동 계산 필드
    ProductionDays As Integer     ' 생산 소요일
    TransferDate As Date          ' 이동일
    ProductionStartDate As Date   ' 생산시작일
    ProductionEndDate As Date     ' 생산완료일
    
    ' 배정 결과
    AssignedVendor As String      ' 배정업체
    AssignedLine As Integer       ' 배정라인
    OriginalVendor As String      ' 최초배정업체
    VendorChanged As Boolean      ' 업체변경여부
    
    ' 세트품 관련
    IsSetItem As Boolean          ' 세트품 여부
    SetGroupID As String          ' 세트품 그룹 ID
    SetTotalQty As Long           ' 세트품 합산 수량
End Type

Public Type AllocationResult
    Item As ProductionItem
    Success As Boolean
    FailReason As String
    AlternativeVendors As String  ' 대안 업체 리스트
    
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
' 색상 팔레트 (세련된 디자인)
'===============================================================================
' 메인 색상
Private Const COLOR_PRIMARY As Long = 3955058       ' RGB(50, 60, 75) - 진한 네이비
Private Const COLOR_PRIMARY_LIGHT As Long = 15724527 ' RGB(239, 246, 255) - 연한 파랑
Private Const COLOR_ACCENT As Long = 15124704       ' RGB(96, 165, 230) - 스카이블루
Private Const COLOR_SUCCESS As Long = 7978352       ' RGB(16, 185, 129) - 민트그린
Private Const COLOR_WARNING As Long = 7053567       ' RGB(251, 191, 107) - 골드
Private Const COLOR_DANGER As Long = 5765887        ' RGB(239, 68, 87) - 코랄레드

' 배경 색상
Private Const COLOR_BG_WHITE As Long = 16777215     ' RGB(255, 255, 255)
Private Const COLOR_BG_GRAY As Long = 15921906      ' RGB(242, 244, 247)
Private Const COLOR_BG_WEEKEND As Long = 15263484   ' RGB(252, 231, 232)

' 텍스트 색상
Private Const COLOR_TEXT_DARK As Long = 2500134     ' RGB(38, 38, 38)
Private Const COLOR_TEXT_GRAY As Long = 8421504     ' RGB(128, 128, 128)
Private Const COLOR_TEXT_WHITE As Long = 16777215   ' RGB(255, 255, 255)

' 업체별 색상
Private Function GetVendorColor(vendorName As String) As Long
    Select Case vendorName
        Case "위드맘": GetVendorColor = RGB(59, 130, 246)   ' 블루
        Case "리니어": GetVendorColor = RGB(16, 185, 129)   ' 그린
        Case "그램": GetVendorColor = RGB(249, 115, 22)     ' 오렌지
        Case "이시스": GetVendorColor = RGB(139, 92, 246)   ' 퍼플
        Case "엘루오": GetVendorColor = RGB(236, 72, 153)   ' 핑크
        Case "케이코스텍": GetVendorColor = RGB(20, 184, 166) ' 틸
        Case "다미": GetVendorColor = RGB(245, 158, 11)     ' 앰버
        Case Else: GetVendorColor = RGB(107, 114, 128)      ' 그레이
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
    
    ' 세트품 식별 및 그룹핑
    IdentifySetItems items, itemCount
    
    ' 역산 로직 실행 (납기 → 생산일 → 이동일)
    CalculateProductionDates items, itemCount
    
    ' 날짜 범위 파악
    Dim minDate As Date, maxDate As Date
    GetDateRange items, itemCount, minDate, maxDate
    
    ' 자동 배정 실행
    Dim results() As AllocationResult
    AllocateProduction items, itemCount, results, minDate, maxDate
    
    ' Raw 데이터 시트에 결과 기록
    WriteResultsToRaw items, itemCount
    
    ' 생산계획표 생성
    CreateProductionPlanView results, itemCount, minDate, maxDate
    
    ' 완료 메시지
    Dim elapsed As Double
    elapsed = Timer - startTime
    MsgBox "자동배정 완료!" & vbCrLf & vbCrLf & _
           "처리 건수: " & itemCount & "건" & vbCrLf & _
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
    
    ' Raw 데이터 시트 생성
    CreateRawDataSheet
    
    ' 외주처 설정 시트 생성
    CreateVendorSheet
    
    ' 고객매칭 설정 시트 생성
    CreateMappingSheet
    
    ' 생산계획표 시트 생성
    CreateOrClearSheet SHEET_OUTPUT
    
    Application.ScreenUpdating = True
    
    MsgBox "초기 설정이 완료되었습니다!" & vbCrLf & vbCrLf & _
           "1. 'Raw데이터' 시트에 데이터를 입력하세요." & vbCrLf & _
           "2. '설정_외주처' 시트에서 외주처 정보를 확인하세요." & vbCrLf & _
           "3. '설정_고객매칭' 시트에서 고객매칭을 설정하세요." & vbCrLf & _
           "4. '자동배정실행' 매크로를 실행하세요.", vbInformation, "초기 설정 완료"
End Sub

'===============================================================================
' Raw 데이터 시트 생성
'===============================================================================
Private Sub CreateRawDataSheet()
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_RAW)
    
    ' 헤더 설정
    Dim headers As Variant
    headers = Array("품번", "품명", "수량", "납기", "공정명", "특수공정", _
                    "지정업체", "긴급도", "상태", "비고", _
                    "고객사코드", "배정업체", "이동일", "생산일", "최초배정업체", "업체변경")
    
    Dim i As Integer
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    
    ' 헤더 스타일
    With ws.Range("A1:P1")
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = COLOR_TEXT_WHITE
        .Interior.Color = COLOR_PRIMARY
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' 입력 영역 (A~J) vs 자동 영역 (K~P) 구분
    With ws.Range("A1:J1")
        .Interior.Color = RGB(59, 130, 246)  ' 파란색 (입력 영역)
    End With
    With ws.Range("K1:P1")
        .Interior.Color = RGB(107, 114, 128) ' 회색 (자동 영역)
    End With
    
    ' 열 너비 설정
    ws.Columns("A").ColumnWidth = 15  ' 품번
    ws.Columns("B").ColumnWidth = 20  ' 품명
    ws.Columns("C").ColumnWidth = 10  ' 수량
    ws.Columns("D").ColumnWidth = 12  ' 납기
    ws.Columns("E").ColumnWidth = 12  ' 공정명
    ws.Columns("F").ColumnWidth = 10  ' 특수공정
    ws.Columns("G").ColumnWidth = 12  ' 지정업체
    ws.Columns("H").ColumnWidth = 8   ' 긴급도
    ws.Columns("I").ColumnWidth = 10  ' 상태
    ws.Columns("J").ColumnWidth = 15  ' 비고
    ws.Columns("K").ColumnWidth = 10  ' 고객사코드
    ws.Columns("L").ColumnWidth = 12  ' 배정업체
    ws.Columns("M").ColumnWidth = 12  ' 이동일
    ws.Columns("N").ColumnWidth = 12  ' 생산일
    ws.Columns("O").ColumnWidth = 12  ' 최초배정업체
    ws.Columns("P").ColumnWidth = 10  ' 업체변경
    
    ' 드롭다운 설정 (공정명)
    With ws.Range("E2:E1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="포장,충전,충포장,충전/충포장,포장/충포장"
    End With
    
    ' 드롭다운 설정 (특수공정)
    With ws.Range("F2:F1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="수축,교반,고주파"
    End With
    
    ' 드롭다운 설정 (긴급도)
    With ws.Range("H2:H1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="급함,긴급"
    End With
    
    ' 드롭다운 설정 (상태)
    With ws.Range("I2:I1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="보류,배정완료,발주완료"
    End With
    
    ' 고객사코드 자동 추출 수식 (K열)
    ws.Range("K2").Formula = "=IF(A2="""","""",MID(A2,2,3))"
    ws.Range("K2").AutoFill Destination:=ws.Range("K2:K100")
    
    ' 조건부 서식 - 상태별 색상
    ApplyConditionalFormatting ws
    
    ' 샘플 데이터 추가
    AddSampleData ws
    
    ' 틀 고정
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

'===============================================================================
' 조건부 서식 적용
'===============================================================================
Private Sub ApplyConditionalFormatting(ws As Worksheet)
    Dim rng As Range
    Set rng = ws.Range("I2:I1000")
    
    ' 기존 조건부 서식 삭제
    rng.FormatConditions.Delete
    
    ' 상태별 색상
    ' 보류 - 회색
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""보류""")
        .Interior.Color = RGB(229, 231, 235)
        .Font.Color = RGB(107, 114, 128)
    End With
    
    ' 배정완료 - 연한 초록
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""배정완료""")
        .Interior.Color = RGB(209, 250, 229)
        .Font.Color = RGB(6, 95, 70)
    End With
    
    ' 발주완료 - 연한 파랑
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""발주완료""")
        .Interior.Color = RGB(219, 234, 254)
        .Font.Color = RGB(30, 64, 175)
    End With
    
    ' 긴급도별 색상 (H열)
    Set rng = ws.Range("H2:H1000")
    rng.FormatConditions.Delete
    
    ' 급함 - 주황
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""급함""")
        .Interior.Color = RGB(255, 237, 213)
        .Font.Color = RGB(194, 65, 12)
        .Font.Bold = True
    End With
    
    ' 긴급 - 빨강
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""긴급""")
        .Interior.Color = RGB(254, 226, 226)
        .Font.Color = RGB(185, 28, 28)
        .Font.Bold = True
    End With
End Sub

'===============================================================================
' 샘플 데이터 추가
'===============================================================================
Private Sub AddSampleData(ws As Worksheet)
    Dim data As Variant
    data = Array( _
        Array("1OLV01", "샴푸 500ml", 15000, "2026-01-25", "충포장", "", "", "", "", ""), _
        Array("2AMR05", "로션 200ml", 8000, "2026-01-24", "충전", "", "", "급함", "", "수출건"), _
        Array("9CJO12", "크림 50g", 5000, "2026-01-22", "포장", "교반", "", "긴급", "", "홈쇼핑"), _
        Array("9AMR10", "세트A-포장", 5000, "2026-01-26", "포장/충포장", "", "", "", "", ""), _
        Array("1AMR10", "세트A-충전1", 5000, "2026-01-26", "충전/충포장", "", "", "", "", ""), _
        Array("1AMR11", "세트A-충전2", 3000, "2026-01-26", "충전/충포장", "", "", "", "", ""), _
        Array("1LGH03", "에센스 30ml", 3000, "2026-01-28", "충포장", "", "리니어", "", "", ""), _
        Array("9EMT07", "샴푸 1L", 10000, "2026-01-30", "충포장", "", "", "", "보류", "자재 1/22") _
    )
    
    Dim i As Integer, j As Integer
    For i = 0 To UBound(data)
        For j = 0 To UBound(data(i))
            ws.Cells(i + 2, j + 1).Value = data(i)(j)
        Next j
    Next i
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
    
    ' 헤더 스타일
    With ws.Range("A1:G1")
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = COLOR_TEXT_WHITE
        .Interior.Color = COLOR_PRIMARY
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With
    
    ' 데이터
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
    
    ' 열 너비
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 35
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 10
    
    ' 테두리
    Dim lastRow As Integer: lastRow = UBound(data) + 2
    With ws.Range("A1:G" & lastRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(203, 213, 225)
    End With
End Sub

'===============================================================================
' 고객매칭 설정 시트 생성
'===============================================================================
Private Sub CreateMappingSheet()
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_MAPPING)
    
    ' 헤더
    ws.Range("A1:C1").Value = Array("고객사코드", "외주처ID", "우선순위")
    
    ' 헤더 스타일
    With ws.Range("A1:C1")
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = COLOR_TEXT_WHITE
        .Interior.Color = COLOR_PRIMARY
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With
    
    ' 데이터
    Dim data As Variant
    data = Array( _
        Array("CLO", "gram", 1), _
        Array("ERK", "gram", 1), _
        Array("DPD", "withmom", 1), _
        Array("GDI", "withmom", 1), _
        Array("MDH", "linear", 1), _
        Array("APS", "linear", 1), _
        Array("PUR", "kcostech", 1), _
        Array("OLV", "withmom", 1), _
        Array("AMR", "linear", 1), _
        Array("CJO", "gram", 1), _
        Array("LGH", "linear", 1), _
        Array("EMT", "withmom", 1) _
    )
    
    Dim i As Integer
    For i = 0 To UBound(data)
        ws.Range("A" & (i + 2) & ":C" & (i + 2)).Value = data(i)
    Next i
    
    ' 열 너비
    ws.Columns("A").ColumnWidth = 15
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 10
    
    ' 테두리
    Dim lastRow As Integer: lastRow = UBound(data) + 2
    With ws.Range("A1:C" & lastRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(203, 213, 225)
    End With
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
    
    ' 고객매칭 로드
    Dim wsMapping As Worksheet
    Set wsMapping = ThisWorkbook.Sheets(SHEET_MAPPING)
    
    Set ClientMappings = CreateObject("Scripting.Dictionary")
    
    lastRow = wsMapping.Cells(wsMapping.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        Dim clientCode As String, vendorId As String
        clientCode = Trim(CStr(wsMapping.Cells(i, 1).Value))
        vendorId = Trim(CStr(wsMapping.Cells(i, 2).Value))
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
            .ProductCode = Trim(CStr(ws.Cells(i, RAW_COL_CODE).Value))
            .ProductName = Trim(CStr(ws.Cells(i, RAW_COL_NAME).Value))
            .Quantity = CLng(Val(ws.Cells(i, RAW_COL_QTY).Value))
            
            ' 납기 파싱
            If IsDate(ws.Cells(i, RAW_COL_DUEDATE).Value) Then
                .DueDate = CDate(ws.Cells(i, RAW_COL_DUEDATE).Value)
            Else
                .DueDate = ParseDateString(CStr(ws.Cells(i, RAW_COL_DUEDATE).Value))
            End If
            
            .ProcessType = Trim(CStr(ws.Cells(i, RAW_COL_PROCESS).Value))
            .SpecialProcess = Trim(CStr(ws.Cells(i, RAW_COL_SPECIAL).Value))
            .DesignatedVendor = Trim(CStr(ws.Cells(i, RAW_COL_VENDOR).Value))
            .Urgency = Trim(CStr(ws.Cells(i, RAW_COL_URGENCY).Value))
            .Status = Trim(CStr(ws.Cells(i, RAW_COL_STATUS).Value))
            .Remarks = Trim(CStr(ws.Cells(i, RAW_COL_REMARKS).Value))
            
            ' 고객사코드 추출 (품번 2~4번째 자리)
            If Len(.ProductCode) >= 4 Then
                .ClientCode = Mid(.ProductCode, 2, 3)
            End If
            
            ' 기존 배정 정보 (재배정 시 참고)
            .AssignedVendor = Trim(CStr(ws.Cells(i, RAW_COL_ASSIGNED).Value))
            .OriginalVendor = Trim(CStr(ws.Cells(i, RAW_COL_ORIGINAL).Value))
            
            ' 기본값 설정 (빈칸 = 기본값)
            If Len(.Urgency) = 0 Then .Urgency = "일반"
            If Len(.Status) = 0 Then .Status = "신규"
        End With
    Next i
    
    ReadRawData = True
    Exit Function
    
ErrorHandler:
    ReadRawData = False
End Function

'===============================================================================
' 날짜 문자열 파싱
'===============================================================================
Private Function ParseDateString(dateStr As String) As Date
    On Error GoTo ErrorHandler
    
    dateStr = Trim(dateStr)
    If Len(dateStr) = 0 Then
        ParseDateString = 0
        Exit Function
    End If
    
    ' "1/20", "1월20일", "2026-01-20" 등 다양한 형식 처리
    If InStr(dateStr, "-") > 0 Then
        ' YYYY-MM-DD 형식
        ParseDateString = CDate(dateStr)
    ElseIf InStr(dateStr, "/") > 0 Then
        ' M/D 형식
        Dim parts() As String
        parts = Split(dateStr, "/")
        If UBound(parts) >= 1 Then
            ParseDateString = DateSerial(Year(Date), CInt(parts(0)), CInt(parts(1)))
        End If
    ElseIf InStr(dateStr, "월") > 0 Then
        ' M월D일 형식
        Dim m As Integer, d As Integer
        m = CInt(Left(dateStr, InStr(dateStr, "월") - 1))
        Dim dayPart As String
        dayPart = Mid(dateStr, InStr(dateStr, "월") + 1)
        dayPart = Replace(dayPart, "일", "")
        d = CInt(dayPart)
        ParseDateString = DateSerial(Year(Date), m, d)
    Else
        ParseDateString = CDate(dateStr)
    End If
    
    Exit Function
    
ErrorHandler:
    ParseDateString = 0
End Function

'===============================================================================
' 세트품 식별
'===============================================================================
Private Sub IdentifySetItems(ByRef items() As ProductionItem, ByVal itemCount As Long)
    If itemCount = 0 Then Exit Sub
    
    Dim i As Long, j As Long
    Dim setGroupCounter As Integer: setGroupCounter = 0
    
    For i = 1 To itemCount
        ' 이미 세트품으로 식별된 경우 스킵
        If items(i).IsSetItem Then GoTo NextItem
        
        ' 세트품 공정인지 확인
        If items(i).ProcessType = "포장/충포장" Or items(i).ProcessType = "충전/충포장" Then
            ' 세트품 그룹 시작
            setGroupCounter = setGroupCounter + 1
            Dim groupID As String
            groupID = "SET" & Format(setGroupCounter, "000")
            
            items(i).IsSetItem = True
            items(i).SetGroupID = groupID
            
            Dim setTotalQty As Long
            setTotalQty = items(i).Quantity
            
            ' 연속된 행에서 같은 고객사코드 + 세트품 공정인 것들 묶기
            For j = i + 1 To itemCount
                If items(j).ProcessType = "포장/충포장" Or items(j).ProcessType = "충전/충포장" Then
                    If items(j).ClientCode = items(i).ClientCode Then
                        items(j).IsSetItem = True
                        items(j).SetGroupID = groupID
                        setTotalQty = setTotalQty + items(j).Quantity
                    Else
                        Exit For  ' 다른 고객사코드면 세트 끝
                    End If
                Else
                    Exit For  ' 세트품 공정이 아니면 세트 끝
                End If
            Next j
            
            ' 세트품 합산 수량 기록
            For j = i To itemCount
                If items(j).SetGroupID = groupID Then
                    items(j).SetTotalQty = setTotalQty
                End If
            Next j
        End If
NextItem:
    Next i
End Sub

'===============================================================================
' 역산 로직: 납기 → 생산일 → 이동일
'===============================================================================
Private Sub CalculateProductionDates(ByRef items() As ProductionItem, ByVal itemCount As Long)
    Dim i As Long
    
    For i = 1 To itemCount
        With items(i)
            ' 발주완료 건은 이미 확정된 날짜 유지
            If .Status = "발주완료" Then GoTo NextCalc
            
            ' 납기가 없으면 스킵
            If .DueDate = 0 Then GoTo NextCalc
            
            ' 생산 소요일 계산
            Dim dailyCapa As Long
            dailyCapa = GetEffectiveCapacity(.AssignedVendor, .Urgency, .IsSetItem)
            If dailyCapa = 0 Then dailyCapa = 10000 ' 기본값
            
            ' 세트품이면 합산 수량 기준
            Dim qty As Long
            If .IsSetItem And .SetTotalQty > 0 Then
                qty = .SetTotalQty
            Else
                qty = .Quantity
            End If
            
            .ProductionDays = WorksheetFunction.RoundUp(qty / dailyCapa, 0)
            If .ProductionDays < 1 Then .ProductionDays = 1
            
            ' 생산완료일 = 납기 - 1 (일반) / 납기 당일 (긴급 직납)
            If .Urgency = "긴급" Then
                .ProductionEndDate = .DueDate  ' 당일 직납 가능
            Else
                .ProductionEndDate = .DueDate - 1
            End If
            
            ' 생산시작일 = 생산완료일 - 생산소요일 + 1 (주말 제외)
            .ProductionStartDate = SubtractWorkdays(.ProductionEndDate, .ProductionDays - 1)
            
            ' 이동일 = 생산시작일 - 1 (영업일 기준)
            .TransferDate = SubtractWorkdays(.ProductionStartDate, 1)
        End With
NextCalc:
    Next i
End Sub

'===============================================================================
' 영업일 빼기 (주말 제외)
'===============================================================================
Private Function SubtractWorkdays(startDate As Date, days As Integer) As Date
    Dim result As Date: result = startDate
    Dim subtracted As Integer: subtracted = 0
    
    Do While subtracted < days
        result = result - 1
        If Weekday(result) <> vbSunday And Weekday(result) <> vbSaturday Then
            subtracted = subtracted + 1
        End If
    Loop
    
    ' 결과가 주말이면 금요일로
    Do While Weekday(result) = vbSunday Or Weekday(result) = vbSaturday
        result = result - 1
    Loop
    
    SubtractWorkdays = result
End Function

'===============================================================================
' 유효 Capa 계산 (긴급도, 세트품 반영)
'===============================================================================
Private Function GetEffectiveCapacity(vendorName As String, urgency As String, isSetItem As Boolean) As Long
    Dim baseCapa As Long
    baseCapa = GetVendorCapacity(vendorName)
    If baseCapa = 0 Then baseCapa = 10000
    
    ' 긴급도에 따른 Capa 확장
    Dim capaHours As Double
    Select Case urgency
        Case "긴급"
            capaHours = HOURS_BASIC + HOURS_OVERTIME + (HOURS_SATURDAY / 5)  ' 토특까지 분배
        Case "급함"
            capaHours = HOURS_BASIC + HOURS_OVERTIME  ' 잔업까지
        Case Else
            capaHours = HOURS_BASIC  ' 기본만
    End Select
    
    Dim effectiveCapa As Double
    effectiveCapa = baseCapa * (capaHours / HOURS_BASIC)
    
    ' 세트품이면 80% 적용
    If isSetItem Then
        effectiveCapa = effectiveCapa * SET_CAPA_FACTOR
    End If
    
    GetEffectiveCapacity = CLng(effectiveCapa)
End Function

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
' 날짜 범위 파악
'===============================================================================
Private Sub GetDateRange(ByRef items() As ProductionItem, ByVal itemCount As Long, _
                         ByRef minDate As Date, ByRef maxDate As Date)
    minDate = DateSerial(2099, 12, 31)
    maxDate = DateSerial(1900, 1, 1)
    
    Dim i As Long
    For i = 1 To itemCount
        If items(i).TransferDate > 0 Then
            If items(i).TransferDate < minDate Then minDate = items(i).TransferDate
        End If
        If items(i).DueDate > 0 Then
            If items(i).DueDate > maxDate Then maxDate = items(i).DueDate
        End If
    Next i
    
    ' 기본값
    If minDate > maxDate Then
        minDate = Date
        maxDate = Date + 30
    End If
    
    ' 월 시작/끝으로 맞추기 (여유 있게)
    minDate = minDate - 3
    maxDate = maxDate + 7
End Sub

'===============================================================================
' 자동 배정 로직
'===============================================================================
Private Sub AllocateProduction(ByRef items() As ProductionItem, ByVal itemCount As Long, _
                               ByRef results() As AllocationResult, _
                               ByVal minDate As Date, ByVal maxDate As Date)
    ReDim results(1 To itemCount)
    
    Dim i As Long
    For i = 1 To itemCount
        results(i).Item = items(i)
        results(i).Success = False
        results(i).FailReason = ""
        
        ' 상태 확인
        Select Case items(i).Status
            Case "발주완료"
                ' 절대 고정 - 기존 배정 유지
                results(i).Success = True
                results(i).Item.AssignedVendor = items(i).AssignedVendor
                GoTo NextAlloc
                
            Case "보류"
                ' 스킵
                results(i).FailReason = "HOLD"
                GoTo NextAlloc
                
            Case "배정완료"
                ' D+2 이내면 고정, D+3 이후면 재배정 가능
                If IsWithinFixedPeriod(items(i).TransferDate) Then
                    results(i).Success = True
                    results(i).Item.AssignedVendor = items(i).AssignedVendor
                    GoTo NextAlloc
                End If
                ' D+3 이후면 재배정 진행
        End Select
        
        ' 이동일 유효성 체크
        If items(i).TransferDate = 0 Then
            results(i).FailReason = "INVALID_DATE"
            GoTo NextAlloc
        End If
        
        ' D+2 이내 신규건은 수동 처리 알림
        If items(i).Status = "신규" And IsWithinFixedPeriod(items(i).TransferDate) Then
            results(i).FailReason = "WITHIN_FIXED_PERIOD"
            GoTo NextAlloc
        End If
        
        ' 지정업체 있는 경우
        If Len(items(i).DesignatedVendor) > 0 Then
            results(i).Item.AssignedVendor = items(i).DesignatedVendor
            
            ' 경고 체크
            CheckWarnings results(i), items(i)
            
            ' 슬롯 검색
            If FindSlotForVendor(results(i), items(i), maxDate) Then
                results(i).Success = True
            Else
                results(i).FailReason = "NO_SLOT_MANUAL|" & items(i).DesignatedVendor
                ' 대안 리스트 생성
                results(i).AlternativeVendors = GetAlternativeVendors(items(i), maxDate)
            End If
        Else
            ' 자동 배정
            ' 1. 기존 업체 우선 (이동일 변경 시)
            If Len(items(i).AssignedVendor) > 0 Then
                results(i).Item.AssignedVendor = items(i).AssignedVendor
                If FindSlotForVendor(results(i), items(i), maxDate) Then
                    results(i).Success = True
                    results(i).Item.VendorChanged = False
                    GoTo NextAlloc
                End If
            End If
            
            ' 2. 고객매칭 업체
            Dim matchedVendor As String
            matchedVendor = GetMatchedVendor(items(i))
            If Len(matchedVendor) > 0 Then
                results(i).Item.AssignedVendor = matchedVendor
                If FindSlotForVendor(results(i), items(i), maxDate) Then
                    results(i).Success = True
                    If items(i).AssignedVendor <> matchedVendor Then
                        results(i).Item.VendorChanged = True
                    End If
                    GoTo NextAlloc
                End If
            End If
            
            ' 3. 우선순위 순 검색
            Dim vendorIdx As Integer
            For vendorIdx = 1 To VendorCount
                If CanHandleItem(Vendors(vendorIdx), items(i)) Then
                    results(i).Item.AssignedVendor = Vendors(vendorIdx).Name
                    If FindSlotForVendor(results(i), items(i), maxDate) Then
                        results(i).Success = True
                        If items(i).AssignedVendor <> Vendors(vendorIdx).Name Then
                            results(i).Item.VendorChanged = True
                        End If
                        GoTo NextAlloc
                    End If
                End If
            Next vendorIdx
            
            ' 4. 배정 실패
            results(i).FailReason = "NO_AVAILABLE_SLOT"
            results(i).AlternativeVendors = GetAlternativeVendors(items(i), maxDate)
        End If
        
NextAlloc:
        ' 최초배정업체 기록
        If Len(results(i).Item.OriginalVendor) = 0 Then
            results(i).Item.OriginalVendor = results(i).Item.AssignedVendor
        End If
    Next i
End Sub

'===============================================================================
' D+2 이내인지 확인
'===============================================================================
Private Function IsWithinFixedPeriod(transferDate As Date) As Boolean
    If transferDate = 0 Then
        IsWithinFixedPeriod = False
        Exit Function
    End If
    IsWithinFixedPeriod = (transferDate <= Date + FIXED_DAYS)
End Function

'===============================================================================
' 경고 체크
'===============================================================================
Private Sub CheckWarnings(ByRef result As AllocationResult, ByRef Item As ProductionItem)
    result.HasWarning = False
    
    Dim vendorIdx As Integer
    vendorIdx = GetVendorIndex(Item.DesignatedVendor)
    If vendorIdx = 0 Then Exit Sub
    
    ' 특수공정 불가 체크
    If Len(Item.SpecialProcess) > 0 Then
        If Not CanHandleProcess(Vendors(vendorIdx), Item.SpecialProcess) Then
            result.HasWarning = True
            result.WarningType = "PROCESS_UNABLE"
            result.WarningMessage = Item.DesignatedVendor & "은(는) " & _
                                    Item.SpecialProcess & " 공정 불가"
            Exit Sub
        End If
    End If
    
    ' 고객매칭 불일치 체크
    If Len(Item.ClientCode) > 0 Then
        If ClientMappings.Exists(Item.ClientCode) Then
            Dim mappedVendorId As String
            mappedVendorId = ClientMappings(Item.ClientCode)
            If Vendors(vendorIdx).ID <> mappedVendorId Then
                result.HasWarning = True
                result.WarningType = "CLIENT_MISMATCH"
                result.WarningMessage = "고객매칭: " & GetVendorNameById(mappedVendorId) & _
                                        " / 지정: " & Item.DesignatedVendor
            End If
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
' 업체ID로 이름 조회
'===============================================================================
Private Function GetVendorNameById(vendorId As String) As String
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).ID = vendorId Then
            GetVendorNameById = Vendors(i).Name
            Exit Function
        End If
    Next i
    GetVendorNameById = vendorId
End Function

'===============================================================================
' 매칭 업체 조회
'===============================================================================
Private Function GetMatchedVendor(ByRef Item As ProductionItem) As String
    GetMatchedVendor = ""
    
    If Len(Item.ClientCode) = 0 Then Exit Function
    If Not ClientMappings.Exists(Item.ClientCode) Then Exit Function
    
    Dim vendorId As String
    vendorId = ClientMappings(Item.ClientCode)
    
    ' 특수공정 가능 여부 확인
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).ID = vendorId Then
            If CanHandleItem(Vendors(i), Item) Then
                GetMatchedVendor = Vendors(i).Name
            End If
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' 품목 처리 가능 여부
'===============================================================================
Private Function CanHandleItem(ByRef vendor As VendorInfo, ByRef Item As ProductionItem) As Boolean
    CanHandleItem = CanHandleProcess(vendor, Item.SpecialProcess)
End Function

'===============================================================================
' 공정 처리 가능 여부
'===============================================================================
Private Function CanHandleProcess(ByRef vendor As VendorInfo, specialProcess As String) As Boolean
    If Len(specialProcess) = 0 Then
        CanHandleProcess = True
        Exit Function
    End If
    
    Dim processCode As String
    Select Case specialProcess
        Case "수축": processCode = "shrink"
        Case "교반": processCode = "mixing"
        Case "고주파": processCode = "highFrequency"
        Case Else: processCode = specialProcess
    End Select
    
    CanHandleProcess = (InStr(1, vendor.Capabilities, processCode, vbTextCompare) > 0)
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
    startDate = Item.TransferDate + 1  ' 이동일 다음날부터 생산
    If startDate < Date Then startDate = Date
    
    Dim daysNeeded As Integer
    daysNeeded = Item.ProductionDays
    If daysNeeded < 1 Then daysNeeded = 1
    
    Dim line As Integer
    Dim searchDate As Date: searchDate = startDate
    
    Do While searchDate <= maxDate
        ' 주말 스킵
        If Weekday(searchDate) = vbSunday Or Weekday(searchDate) = vbSaturday Then
            searchDate = searchDate + 1
            GoTo ContinueSearch
        End If
        
        For line = 1 To Vendors(vendorIdx).LineCount
            If CanPlaceOnLine(Vendors(vendorIdx).Name, line, searchDate, daysNeeded, maxDate) Then
                ReserveLine Vendors(vendorIdx).Name, line, searchDate, daysNeeded
                result.Item.AssignedLine = line
                result.Item.ProductionStartDate = searchDate
                FindSlotForVendor = True
                Exit Function
            End If
        Next line
        
        searchDate = searchDate + 1
ContinueSearch:
    Loop
End Function

'===============================================================================
' 라인 배치 가능 여부
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
' 라인 예약
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
Private Sub WriteResultsToRaw(ByRef items() As ProductionItem, ByVal itemCount As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_RAW)
    
    Dim i As Long
    For i = 1 To itemCount
        With items(i)
            If .RowNum > 0 Then
                ws.Cells(.RowNum, RAW_COL_CLIENT).Value = .ClientCode
                ws.Cells(.RowNum, RAW_COL_ASSIGNED).Value = .AssignedVendor
                If .TransferDate > 0 Then
                    ws.Cells(.RowNum, RAW_COL_TRANSFER).Value = .TransferDate
                    ws.Cells(.RowNum, RAW_COL_TRANSFER).NumberFormat = "yyyy-mm-dd"
                End If
                If .ProductionStartDate > 0 Then
                    ws.Cells(.RowNum, RAW_COL_PRODDATE).Value = .ProductionStartDate
                    ws.Cells(.RowNum, RAW_COL_PRODDATE).NumberFormat = "yyyy-mm-dd"
                End If
                ws.Cells(.RowNum, RAW_COL_ORIGINAL).Value = .OriginalVendor
                ws.Cells(.RowNum, RAW_COL_CHANGED).Value = IIf(.VendorChanged, "변경됨", "유지")
            End If
        End With
    Next i
End Sub

'===============================================================================
' 생산계획표 생성 (세련된 디자인)
'===============================================================================
Private Sub CreateProductionPlanView(ByRef results() As AllocationResult, _
                                     ByVal itemCount As Long, _
                                     ByVal minDate As Date, ByVal maxDate As Date)
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_OUTPUT)
    
    ' 기본 폰트 설정
    ws.Cells.Font.Name = "맑은 고딕"
    ws.Cells.Font.Size = 9
    
    ' 기존 메모 삭제
    On Error Resume Next
    ws.Cells.ClearComments
    On Error GoTo 0
    
    Dim currentRow As Long: currentRow = 1
    
    ' =====================================================================
    ' 타이틀 섹션 (세련된 디자인)
    ' =====================================================================
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
    
    ' 구간 안내
    ws.Cells(currentRow, 1).Value = "D+2 확정구간: ~" & Format(Date + FIXED_DAYS, "m/d") & _
                                    "  |  D+3 조정가능구간: " & Format(Date + FIXED_DAYS + 1, "m/d") & "~"
    With ws.Cells(currentRow, 1)
        .Font.Size = 10
        .Font.Color = RGB(59, 130, 246)
        .Font.Italic = True
    End With
    currentRow = currentRow + 2
    
    ' =====================================================================
    ' 달력 헤더
    ' =====================================================================
    Dim totalDays As Long
    totalDays = DateDiff("d", minDate, maxDate) + 1
    
    Dim headerRow As Long: headerRow = currentRow
    
    ' A열 헤더
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
        
        ' 월 변경 시 월 표시
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
            
            ' 색상 설정
            If dow = vbSunday Or dow = vbSaturday Then
                ' 주말
                .Interior.Color = RGB(254, 226, 226)
                .Font.Color = RGB(185, 28, 28)
            ElseIf currentDate = Date Then
                ' 오늘
                .Interior.Color = RGB(187, 247, 208)
                .Font.Color = RGB(22, 101, 52)
            ElseIf currentDate <= Date + FIXED_DAYS Then
                ' D+2 확정구간
                .Interior.Color = RGB(219, 234, 254)
                .Font.Color = RGB(30, 64, 175)
            Else
                ' 일반 평일
                .Interior.Color = RGB(241, 245, 249)
                .Font.Color = RGB(51, 65, 85)
            End If
        End With
        
        ws.Columns(col).ColumnWidth = 13
    Next d
    
    ws.Rows(headerRow).RowHeight = 42
    
    ' =====================================================================
    ' 외주처별 행
    ' =====================================================================
    currentRow = headerRow + 1
    Dim v As Integer
    
    For v = 1 To VendorCount
        Dim vendorName As String: vendorName = Vendors(v).Name
        
        ' 해당 업체에 배정된 품목 있는지 확인
        Dim hasItems As Boolean: hasItems = False
        Dim i As Long
        For i = 1 To itemCount
            If results(i).Item.AssignedVendor = vendorName Then
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
        
        ' 업체 행 배경
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
                ElseIf currentDate <= Date + FIXED_DAYS Then
                    ws.Cells(currentRow, col).Interior.Color = RGB(239, 246, 255)
                Else
                    ws.Cells(currentRow, col).Interior.Color = RGB(255, 255, 255)
                End If
            Next d
            
            ' 품목 표시
            For i = 1 To itemCount
                If results(i).Item.AssignedVendor = vendorName And _
                   results(i).Item.AssignedLine = line Then
                    PlaceItemOnCalendar ws, results(i), currentRow, minDate, totalDays
                End If
            Next i
            
            ws.Rows(currentRow).RowHeight = 58
            currentRow = currentRow + 1
        Next line
        
NextVendor:
    Next v
    
    ' =====================================================================
    ' 테두리 적용
    ' =====================================================================
    If currentRow > headerRow + 1 Then
        ApplyTableBorders ws, headerRow, currentRow - 1, totalDays + 1
    End If
    
    ' =====================================================================
    ' 배정불가 섹션
    ' =====================================================================
    currentRow = currentRow + 2
    currentRow = CreateUnassignedSection(ws, results, itemCount, currentRow)
    
    ' =====================================================================
    ' 틀 고정
    ' =====================================================================
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(headerRow + 1, 2).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
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
' 품목을 달력에 배치
'===============================================================================
Private Sub PlaceItemOnCalendar(ByRef ws As Worksheet, ByRef result As AllocationResult, _
                                ByVal rowNum As Integer, ByVal minDate As Date, _
                                ByVal totalDays As Long)
    
    If result.Item.ProductionDays = 0 Then Exit Sub
    If result.Item.ProductionStartDate = 0 Then Exit Sub
    
    Dim dailyCapacity As Long
    dailyCapacity = GetEffectiveCapacity(result.Item.AssignedVendor, result.Item.Urgency, result.Item.IsSetItem)
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
        
        ' 셀 내용 (가독성 향상)
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
        
        ' 셀 스타일
        Dim cellBgColor As Long
        Dim fontColor As Long
        
        If result.HasWarning Then
            ' 경고 색상
            If result.WarningType = "PROCESS_UNABLE" Then
                cellBgColor = RGB(254, 202, 202)
                fontColor = RGB(127, 29, 29)
            Else
                cellBgColor = RGB(254, 226, 226)
                fontColor = RGB(153, 27, 27)
            End If
        ElseIf result.Item.VendorChanged Then
            ' 업체 변경
            cellBgColor = RGB(254, 243, 199)
            fontColor = RGB(146, 64, 14)
        Else
            ' 일반
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
        
        ' 메모 (상세정보)
        Dim commentText As String
        commentText = ""
        
        If result.HasWarning Then
            commentText = "[ 경고 ] " & result.WarningMessage & vbLf & "───────────────" & vbLf
        End If
        If result.Item.VendorChanged Then
            commentText = commentText & "[ 업체변경 ] " & result.Item.OriginalVendor & " → " & _
                          result.Item.AssignedVendor & vbLf & "───────────────" & vbLf
        End If
        
        commentText = commentText & "품목코드: " & result.Item.ProductCode & vbLf
        commentText = commentText & "품목명: " & result.Item.ProductName & vbLf
        commentText = commentText & "───────────────" & vbLf
        commentText = commentText & "총 수량: " & Format(result.Item.Quantity, "#,##0") & vbLf
        commentText = commentText & "금일 수량: " & Format(dailyQty, "#,##0") & vbLf
        commentText = commentText & "───────────────" & vbLf
        commentText = commentText & "납기: " & Format(result.Item.DueDate, "yyyy-mm-dd") & vbLf
        commentText = commentText & "이동일: " & Format(result.Item.TransferDate, "yyyy-mm-dd") & vbLf
        commentText = commentText & "외주처: " & result.Item.AssignedVendor & vbLf
        commentText = commentText & "라인: Line " & result.Item.AssignedLine
        
        If result.Item.ProductionDays > 1 Then
            commentText = commentText & vbLf & "───────────────" & vbLf
            commentText = commentText & "진행: " & dayNum & "일차 / " & result.Item.ProductionDays & "일"
        End If
        
        On Error Resume Next
        ws.Cells(rowNum, col).ClearComments
        ws.Cells(rowNum, col).AddComment commentText
        ws.Cells(rowNum, col).Comment.Shape.TextFrame.AutoSize = True
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
        ' 외곽 테두리
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
        ' 내부 테두리
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
                                         ByVal itemCount As Long, ByVal startRow As Integer) As Integer
    Dim hasUnassigned As Boolean: hasUnassigned = False
    Dim i As Long
    For i = 1 To itemCount
        If Not results(i).Success And Len(results(i).FailReason) > 0 And results(i).FailReason <> "HOLD" Then
            hasUnassigned = True
            Exit For
        End If
    Next i
    
    If Not hasUnassigned Then
        CreateUnassignedSection = startRow
        Exit Function
    End If
    
    ' 타이틀
    ws.Cells(startRow, 1).Value = "배정불가 / 수동처리 필요 항목"
    With ws.Cells(startRow, 1)
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(185, 28, 28)
    End With
    startRow = startRow + 1
    
    ' 헤더
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 8)).Value = _
        Array("품번", "품명", "수량", "납기", "이동일", "원인", "대안 업체", "조치")
    
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 8))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(185, 28, 28)
        .HorizontalAlignment = xlCenter
    End With
    startRow = startRow + 1
    
    ' 데이터
    For i = 1 To itemCount
        If Not results(i).Success And Len(results(i).FailReason) > 0 And results(i).FailReason <> "HOLD" Then
            ws.Cells(startRow, 1).Value = results(i).Item.ProductCode
            ws.Cells(startRow, 2).Value = results(i).Item.ProductName
            ws.Cells(startRow, 3).Value = results(i).Item.Quantity
            ws.Cells(startRow, 3).NumberFormat = "#,##0"
            ws.Cells(startRow, 4).Value = results(i).Item.DueDate
            ws.Cells(startRow, 4).NumberFormat = "yyyy-mm-dd"
            ws.Cells(startRow, 5).Value = results(i).Item.TransferDate
            ws.Cells(startRow, 5).NumberFormat = "yyyy-mm-dd"
            
            ' 원인 파싱
            Dim reasonText As String
            Select Case Split(results(i).FailReason, "|")(0)
                Case "WITHIN_FIXED_PERIOD"
                    reasonText = "D+2 이내 (수동처리)"
                Case "NO_SLOT_MANUAL"
                    reasonText = "지정업체 슬롯 부족"
                Case "NO_AVAILABLE_SLOT"
                    reasonText = "가용 슬롯 없음"
                Case "INVALID_DATE"
                    reasonText = "날짜 오류"
                Case Else
                    reasonText = results(i).FailReason
            End Select
            ws.Cells(startRow, 6).Value = reasonText
            
            ' 대안 업체
            ws.Cells(startRow, 7).Value = results(i).AlternativeVendors
            
            ' 조치
            ws.Cells(startRow, 8).Value = "[ ] 억지배정 / [ ] 내부생산 / [ ] 납기협의"
            
            ' 스타일
            ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 8)).Interior.Color = RGB(254, 242, 242)
            ws.Cells(startRow, 6).Font.Color = RGB(185, 28, 28)
            ws.Cells(startRow, 6).Font.Bold = True
            ws.Cells(startRow, 7).Font.Color = RGB(22, 101, 52)
            
            startRow = startRow + 1
        End If
    Next i
    
    ' 열 너비
    ws.Columns(6).ColumnWidth = 20
    ws.Columns(7).ColumnWidth = 40
    ws.Columns(8).ColumnWidth = 35
    
    CreateUnassignedSection = startRow
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
' 업체별 계획표 생성 매크로 (공개)
'===============================================================================
Public Sub 업체별계획표생성()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
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
    
    ' 업체별 시트 생성
    Dim vendorSheetCount As Integer: vendorSheetCount = 0
    Dim v As Integer
    
    For v = 1 To VendorCount
        If CreateVendorPlanSheet(Vendors(v), items, itemCount) Then
            vendorSheetCount = vendorSheetCount + 1
        End If
    Next v
    
    MsgBox "업체별 계획표 생성 완료!" & vbCrLf & vbCrLf & _
           "생성된 시트: " & vendorSheetCount & "개", vbInformation, "완료"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "오류가 발생했습니다:" & vbCrLf & Err.Description, vbCritical, "오류"
    Resume Cleanup
End Sub

'===============================================================================
' 업체별 계획표 시트 생성
'===============================================================================
Private Function CreateVendorPlanSheet(ByRef vendor As VendorInfo, _
                                       ByRef items() As ProductionItem, _
                                       ByVal itemCount As Long) As Boolean
    CreateVendorPlanSheet = False
    
    ' 해당 업체에 배정된 품목 수집
    Dim vendorItems() As ProductionItem
    Dim vendorItemCount As Long: vendorItemCount = 0
    ReDim vendorItems(1 To itemCount)
    
    Dim i As Long
    For i = 1 To itemCount
        If items(i).AssignedVendor = vendor.Name Then
            vendorItemCount = vendorItemCount + 1
            vendorItems(vendorItemCount) = items(i)
        End If
    Next i
    
    ' 배정된 품목 없으면 시트 생성 안함
    If vendorItemCount = 0 Then Exit Function
    
    ' 이동일 기준 정렬
    SortItemsByTransferDate vendorItems, vendorItemCount
    
    ' 시트 생성
    Dim sheetName As String
    sheetName = vendor.Name & "_계획표"
    
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(sheetName)
    
    ' 폰트 기본 설정
    ws.Cells.Font.Name = "맑은 고딕"
    ws.Cells.Font.Size = 10
    
    Dim currentRow As Long: currentRow = 1
    
    ' =====================================================================
    ' 헤더 섹션
    ' =====================================================================
    ' 업체명 타이틀
    ws.Cells(currentRow, 1).Value = "[" & vendor.Name & "] 생산계획표"
    With ws.Cells(currentRow, 1)
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = GetVendorColor(vendor.Name)
    End With
    currentRow = currentRow + 1
    
    ' 기준일시
    ws.Cells(currentRow, 1).Value = "기준일시: " & Format(Now, "yyyy-mm-dd hh:mm") & _
                                    "  |  총 " & vendorItemCount & "건"
    With ws.Cells(currentRow, 1)
        .Font.Size = 11
        .Font.Color = RGB(100, 116, 139)
    End With
    currentRow = currentRow + 1
    
    ' 안내 문구
    ws.Cells(currentRow, 1).Value = "* D+3 이후 계획은 조정될 수 있습니다. 매일 9시 이후 최신 계획표를 공유드립니다."
    With ws.Cells(currentRow, 1)
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(107, 114, 128)
    End With
    currentRow = currentRow + 2
    
    ' =====================================================================
    ' 변경사항 요약
    ' =====================================================================
    Dim hasChanges As Boolean: hasChanges = False
    Dim addedCount As Integer: addedCount = 0
    Dim changedCount As Integer: changedCount = 0
    
    For i = 1 To vendorItemCount
        If vendorItems(i).Status = "신규" Then
            addedCount = addedCount + 1
            hasChanges = True
        End If
        If vendorItems(i).VendorChanged Then
            changedCount = changedCount + 1
            hasChanges = True
        End If
    Next i
    
    If hasChanges Then
        ws.Cells(currentRow, 1).Value = "변경사항 있음"
        With ws.Cells(currentRow, 1)
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(220, 38, 38)
        End With
        
        Dim changeText As String: changeText = ""
        If addedCount > 0 Then changeText = changeText & "신규 추가: " & addedCount & "건  "
        If changedCount > 0 Then changeText = changeText & "변경: " & changedCount & "건"
        
        ws.Cells(currentRow, 2).Value = changeText
        With ws.Cells(currentRow, 2)
            .Font.Size = 11
            .Font.Color = RGB(220, 38, 38)
        End With
        currentRow = currentRow + 2
    Else
        ws.Cells(currentRow, 1).Value = "변경사항 없음"
        With ws.Cells(currentRow, 1)
            .Font.Size = 12
            .Font.Color = RGB(34, 197, 94)
        End With
        currentRow = currentRow + 2
    End If
    
    ' =====================================================================
    ' 테이블 헤더
    ' =====================================================================
    Dim headerRow As Long: headerRow = currentRow
    
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, 8)).Value = _
        Array("이동일", "품번", "품명", "수량", "납기", "공정", "긴급도", "변경")
    
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, 8))
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = GetVendorColor(vendor.Name)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With
    currentRow = currentRow + 1
    
    ' =====================================================================
    ' 데이터 행
    ' =====================================================================
    Dim prevTransferDate As Date: prevTransferDate = 0
    
    For i = 1 To vendorItemCount
        With vendorItems(i)
            ' 이동일
            ws.Cells(currentRow, 1).Value = .TransferDate
            ws.Cells(currentRow, 1).NumberFormat = "m/d (aaa)"
            
            ' 품번
            ws.Cells(currentRow, 2).Value = .ProductCode
            
            ' 품명
            ws.Cells(currentRow, 3).Value = .ProductName
            
            ' 수량
            ws.Cells(currentRow, 4).Value = .Quantity
            ws.Cells(currentRow, 4).NumberFormat = "#,##0"
            
            ' 납기
            ws.Cells(currentRow, 5).Value = .DueDate
            ws.Cells(currentRow, 5).NumberFormat = "m/d"
            
            ' 공정
            ws.Cells(currentRow, 6).Value = .ProcessType
            If Len(.SpecialProcess) > 0 Then
                ws.Cells(currentRow, 6).Value = .ProcessType & " (" & .SpecialProcess & ")"
            End If
            
            ' 긴급도
            ws.Cells(currentRow, 7).Value = .Urgency
            If .Urgency = "긴급" Then
                ws.Cells(currentRow, 7).Font.Color = RGB(185, 28, 28)
                ws.Cells(currentRow, 7).Font.Bold = True
            ElseIf .Urgency = "급함" Then
                ws.Cells(currentRow, 7).Font.Color = RGB(194, 65, 12)
                ws.Cells(currentRow, 7).Font.Bold = True
            End If
            
            ' 변경 상태
            Dim changeStatus As String: changeStatus = ""
            If .Status = "신규" Then
                changeStatus = "추가"
                ws.Cells(currentRow, 8).Font.Color = RGB(220, 38, 38)
                ws.Cells(currentRow, 8).Font.Bold = True
            ElseIf .VendorChanged Then
                changeStatus = "변경"
                ws.Cells(currentRow, 8).Font.Color = RGB(234, 88, 12)
                ws.Cells(currentRow, 8).Font.Bold = True
            End If
            ws.Cells(currentRow, 8).Value = changeStatus
            
            ' 행 스타일
            Dim rowBgColor As Long
            
            ' D+2 확정 구간 표시
            If .TransferDate <= Date + FIXED_DAYS Then
                rowBgColor = RGB(239, 246, 255)  ' 연한 파란색 (확정)
            Else
                ' 짝수/홀수 행 구분
                If (currentRow - headerRow) Mod 2 = 0 Then
                    rowBgColor = RGB(249, 250, 251)
                Else
                    rowBgColor = RGB(255, 255, 255)
                End If
            End If
            
            ' 신규/변경 건 하이라이트
            If Len(changeStatus) > 0 Then
                rowBgColor = RGB(254, 243, 199)  ' 연한 노란색
            End If
            
            ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 8)).Interior.Color = rowBgColor
            
            ' 이동일 변경 시 구분선
            If prevTransferDate > 0 And .TransferDate <> prevTransferDate Then
                ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
                ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 8)).Borders(xlEdgeTop).Weight = xlMedium
                ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 8)).Borders(xlEdgeTop).Color = RGB(203, 213, 225)
            End If
            
            prevTransferDate = .TransferDate
        End With
        
        ws.Rows(currentRow).RowHeight = 22
        currentRow = currentRow + 1
    Next i
    
    ' =====================================================================
    ' 테이블 테두리
    ' =====================================================================
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(currentRow - 1, 8))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = GetVendorColor(vendor.Name)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeRight).Color = GetVendorColor(vendor.Name)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeTop).Color = GetVendorColor(vendor.Name)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = GetVendorColor(vendor.Name)
        
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideVertical).Color = RGB(226, 232, 240)
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideHorizontal).Color = RGB(226, 232, 240)
    End With
    
    ' =====================================================================
    ' 열 너비 설정
    ' =====================================================================
    ws.Columns(1).ColumnWidth = 12  ' 이동일
    ws.Columns(2).ColumnWidth = 14  ' 품번
    ws.Columns(3).ColumnWidth = 20  ' 품명
    ws.Columns(4).ColumnWidth = 10  ' 수량
    ws.Columns(5).ColumnWidth = 8   ' 납기
    ws.Columns(6).ColumnWidth = 16  ' 공정
    ws.Columns(7).ColumnWidth = 8   ' 긴급도
    ws.Columns(8).ColumnWidth = 8   ' 변경
    
    ' =====================================================================
    ' 하단 요약
    ' =====================================================================
    currentRow = currentRow + 1
    
    ' 총 수량
    Dim totalQty As Long: totalQty = 0
    For i = 1 To vendorItemCount
        totalQty = totalQty + vendorItems(i).Quantity
    Next i
    
    ws.Cells(currentRow, 3).Value = "총 수량:"
    ws.Cells(currentRow, 3).Font.Bold = True
    ws.Cells(currentRow, 3).HorizontalAlignment = xlRight
    ws.Cells(currentRow, 4).Value = totalQty
    ws.Cells(currentRow, 4).NumberFormat = "#,##0"
    ws.Cells(currentRow, 4).Font.Bold = True
    
    ' =====================================================================
    ' 틀 고정
    ' =====================================================================
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(headerRow + 1, 1).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    
    CreateVendorPlanSheet = True
End Function

'===============================================================================
' 이동일 기준 정렬
'===============================================================================
Private Sub SortItemsByTransferDate(ByRef items() As ProductionItem, ByVal itemCount As Long)
    Dim i As Long, j As Long
    Dim tempItem As ProductionItem
    
    ' 버블 정렬 (간단한 구현)
    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            If items(j).TransferDate < items(i).TransferDate Then
                tempItem = items(i)
                items(i) = items(j)
                items(j) = tempItem
            End If
        Next j
    Next i
End Sub

'===============================================================================
' 전체 리포트 생성 (메인 계획표 + 업체별 계획표)
'===============================================================================
Public Sub 전체리포트생성()
    ' 자동배정 실행
    Call 자동배정실행
    
    ' 업체별 계획표 생성
    Call 업체별계획표생성
    
    MsgBox "전체 리포트 생성 완료!" & vbCrLf & vbCrLf & _
           "생성된 시트:" & vbCrLf & _
           "- 생산계획표 (전체 캘린더 뷰)" & vbCrLf & _
           "- [업체명]_계획표 (업체별 리스트 뷰)", vbInformation, "완료"
End Sub
