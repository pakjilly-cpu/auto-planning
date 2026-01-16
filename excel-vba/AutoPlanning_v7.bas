'===============================================================================
' 외주처 생산계획 자동화 시스템 - VBA 매크로 v7.0
' - 월별 표 분리 없이 연속 달력
' - 굵은 테두리 + 통일된 열 폭(15)
' - A열 큰 글씨 (업체명/라인)
' - 파스텔 톤 날짜 헤더
' - [NEW] 수동 지정 경고 시스템:
'   * 고객매칭과 다른 업체 지정 시 → 노란색 테두리 + 메모 경고
'   * 특수공정 불가 업체 지정 시 → 빨간색 테두리 + 메모 경고
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
    FailReason As String
    ' Warning 필드 (수동 지정 시 경고)
    HasWarning As Boolean
    WarningType As String      ' "CLIENT_MISMATCH" | "PROCESS_UNABLE"
    WarningMessage As String
End Type

Public Vendors() As VendorInfo
Public VendorCount As Integer
Public ClientMappings As Object
Public LineSchedule As Object

'===============================================================================
' 메인 실행
'===============================================================================
Public Sub 자동배정실행()
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
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
    
    ' 데이터 날짜 범위 파악
    Dim minDate As Date, maxDate As Date
    GetDateRange items, itemCount, minDate, maxDate
    
    Dim results() As AllocationResult
    AllocateProduction items, itemCount, results, minDate, maxDate
    
    CreateContinuousCalendarView results, itemCount, minDate, maxDate
    
    MsgBox "완료! (" & itemCount & "건, " & Format(Timer - startTime, "0.0") & "초)", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
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
' 품목 이동 기능 (두 셀 선택 후 실행)
'===============================================================================
' 사용법:
' 1. 이동할 품목 셀 클릭
' 2. Ctrl 누른 채로 이동할 위치(빈 셀) 클릭
' 3. Alt+F8 → "품목이동" 실행 (또는 버튼에 연결)
'===============================================================================
Public Sub 품목이동()
    On Error GoTo ErrorHandler
    
    ' auto-planning 시트에서만 동작
    If ActiveSheet.Name <> SHEET_OUTPUT Then
        MsgBox "auto-planning 시트에서만 사용 가능합니다.", vbExclamation
        Exit Sub
    End If
    
    ' 선택된 셀이 2개인지 확인
    If Selection.Areas.Count <> 2 Then
        If Selection.Cells.Count <> 2 Then
            MsgBox "사용법:" & vbCrLf & vbCrLf & _
                   "1. 이동할 품목 셀 클릭" & vbCrLf & _
                   "2. Ctrl 누른 채로 이동할 위치 클릭" & vbCrLf & _
                   "3. 이 매크로 실행", vbInformation, "품목 이동"
            Exit Sub
        End If
    End If
    
    Dim sourceCell As Range, targetCell As Range
    
    ' 두 셀 가져오기
    If Selection.Areas.Count = 2 Then
        Set sourceCell = Selection.Areas(1).Cells(1, 1)
        Set targetCell = Selection.Areas(2).Cells(1, 1)
    Else
        Set sourceCell = Selection.Cells(1, 1)
        Set targetCell = Selection.Cells(1, 2)
    End If
    
    ' 원본 셀에 내용이 있는지 확인
    If Len(Trim(sourceCell.Value)) = 0 Then
        MsgBox "첫 번째 선택한 셀에 품목이 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 대상 셀이 비어있는지 확인
    If Len(Trim(targetCell.Value)) > 0 Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("이동할 위치에 이미 품목이 있습니다." & vbCrLf & _
                       "기존 품목: " & Left(targetCell.Value, 30) & "..." & vbCrLf & vbCrLf & _
                       "서로 교환하시겠습니까?", vbYesNo + vbQuestion, "품목 교환")
        If answer = vbNo Then Exit Sub
        
        ' 교환 모드
        SwapCells sourceCell, targetCell
    Else
        ' 이동 모드
        MoveCellContent sourceCell, targetCell
    End If
    
    MsgBox "이동 완료!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "오류: " & Err.Description, vbCritical
End Sub

' 셀 내용 이동 (원본 → 대상)
Private Sub MoveCellContent(ByRef sourceCell As Range, ByRef targetCell As Range)
    ' 내용 복사
    targetCell.Value = sourceCell.Value
    
    ' 스타일 복사
    targetCell.Interior.Color = sourceCell.Interior.Color
    targetCell.Font.Size = sourceCell.Font.Size
    targetCell.Font.Color = sourceCell.Font.Color
    targetCell.Font.Bold = sourceCell.Font.Bold
    targetCell.HorizontalAlignment = sourceCell.HorizontalAlignment
    targetCell.VerticalAlignment = sourceCell.VerticalAlignment
    targetCell.WrapText = sourceCell.WrapText
    
    ' 테두리 복사
    CopyBorders sourceCell, targetCell
    
    ' 메모(Comment) 복사
    On Error Resume Next
    If Not sourceCell.Comment Is Nothing Then
        targetCell.ClearComments
        targetCell.AddComment sourceCell.Comment.Text
        targetCell.Comment.Shape.TextFrame.AutoSize = True
    End If
    On Error GoTo 0
    
    ' 원본 셀 초기화
    ClearCell sourceCell
End Sub

' 두 셀 내용 교환
Private Sub SwapCells(ByRef cell1 As Range, ByRef cell2 As Range)
    ' 임시 저장
    Dim tempValue As String: tempValue = cell1.Value
    Dim tempColor As Long: tempColor = cell1.Interior.Color
    Dim tempFontColor As Long: tempFontColor = cell1.Font.Color
    Dim tempComment As String: tempComment = ""
    
    On Error Resume Next
    If Not cell1.Comment Is Nothing Then tempComment = cell1.Comment.Text
    On Error GoTo 0
    
    ' cell2 → cell1
    cell1.Value = cell2.Value
    cell1.Interior.Color = cell2.Interior.Color
    cell1.Font.Color = cell2.Font.Color
    
    On Error Resume Next
    cell1.ClearComments
    If Not cell2.Comment Is Nothing Then
        cell1.AddComment cell2.Comment.Text
        cell1.Comment.Shape.TextFrame.AutoSize = True
    End If
    On Error GoTo 0
    
    ' temp → cell2
    cell2.Value = tempValue
    cell2.Interior.Color = tempColor
    cell2.Font.Color = tempFontColor
    
    On Error Resume Next
    cell2.ClearComments
    If Len(tempComment) > 0 Then
        cell2.AddComment tempComment
        cell2.Comment.Shape.TextFrame.AutoSize = True
    End If
    On Error GoTo 0
End Sub

' 셀 초기화 (빈 셀로)
Private Sub ClearCell(ByRef cell As Range)
    cell.Value = ""
    cell.Interior.Color = RGB(255, 255, 255)
    cell.Font.Color = RGB(0, 0, 0)
    cell.Font.Bold = False
    
    ' 테두리 초기화
    cell.Borders(xlEdgeLeft).LineStyle = xlNone
    cell.Borders(xlEdgeRight).LineStyle = xlNone
    cell.Borders(xlEdgeTop).LineStyle = xlNone
    cell.Borders(xlEdgeBottom).LineStyle = xlNone
    
    ' 메모 삭제
    On Error Resume Next
    cell.ClearComments
    On Error GoTo 0
End Sub

' 테두리 복사
Private Sub CopyBorders(ByRef source As Range, ByRef target As Range)
    On Error Resume Next
    
    With target.Borders(xlEdgeLeft)
        .LineStyle = source.Borders(xlEdgeLeft).LineStyle
        .Weight = source.Borders(xlEdgeLeft).Weight
        .Color = source.Borders(xlEdgeLeft).Color
    End With
    
    With target.Borders(xlEdgeRight)
        .LineStyle = source.Borders(xlEdgeRight).LineStyle
        .Weight = source.Borders(xlEdgeRight).Weight
        .Color = source.Borders(xlEdgeRight).Color
    End With
    
    With target.Borders(xlEdgeTop)
        .LineStyle = source.Borders(xlEdgeTop).LineStyle
        .Weight = source.Borders(xlEdgeTop).Weight
        .Color = source.Borders(xlEdgeTop).Color
    End With
    
    With target.Borders(xlEdgeBottom)
        .LineStyle = source.Borders(xlEdgeBottom).LineStyle
        .Weight = source.Borders(xlEdgeBottom).Weight
        .Color = source.Borders(xlEdgeBottom).Color
    End With
    
    On Error GoTo 0
End Sub

' 선택 취소 (ESC 대용)
Public Sub 선택취소()
    If ActiveSheet.Name = SHEET_OUTPUT Then
        ActiveSheet.Range("A1").Select
        MsgBox "선택이 취소되었습니다.", vbInformation
    End If
End Sub

'===============================================================================
' 데이터 날짜 범위 파악
'===============================================================================
Private Sub GetDateRange(ByRef items() As ProductionItem, ByVal itemCount As Long, _
                         ByRef minDate As Date, ByRef maxDate As Date)
    minDate = DateSerial(2099, 12, 31)
    maxDate = DateSerial(1900, 1, 1)
    
    Dim i As Long
    For i = 1 To itemCount
        Dim itemDate As Date
        itemDate = ParseTransferDate(items(i).TransferDate, Year(Date))
        
        If itemDate > 0 Then
            If itemDate < minDate Then minDate = itemDate
            If itemDate > maxDate Then maxDate = itemDate
        End If
    Next i
    
    ' 기본값 설정
    If minDate > maxDate Then
        minDate = Date
        maxDate = Date
    End If
    
    ' 월 시작/끝으로 맞추기
    minDate = DateSerial(Year(minDate), Month(minDate), 1)
    maxDate = DateSerial(Year(maxDate), Month(maxDate) + 1, 0)
End Sub

'===============================================================================
' 자동 배정
'===============================================================================
Private Sub AllocateProduction(ByRef items() As ProductionItem, ByVal itemCount As Long, _
                               ByRef results() As AllocationResult, ByVal minDate As Date, ByVal maxDate As Date)
    ReDim results(1 To itemCount)
    
    Dim i As Long
    For i = 1 To itemCount
        results(i).Item = items(i)
        results(i).FailReason = ""
        
        ' 이동일 유효성 체크
        Dim transferDate As Date
        transferDate = ParseTransferDate(items(i).TransferDate, Year(minDate))
        If transferDate = 0 Then
            results(i).VendorName = "배정불가"
            results(i).FailReason = "INVALID_DATE"
            results(i).LineNumber = 0
            results(i).DaysNeeded = 0
            GoTo NextItem
        End If
        
        If Len(items(i).AssignedVendor) > 0 Then
            ' 수동 지정된 외주처
            results(i).VendorName = items(i).AssignedVendor
            results(i).DaysNeeded = CalculateDaysNeeded(items(i).Quantity, GetVendorCapacity(items(i).AssignedVendor))
            
            ' ★ 경고 체크: 수동 지정이 자동배정과 다른지 확인
            CheckManualAssignmentWarning results(i), items(i)
            
            If Not FindAvailableSlot(results(i), items(i).TransferDate, Year(minDate), maxDate) Then
                results(i).VendorName = "배정불가"
                results(i).FailReason = "NO_SLOT_MANUAL|" & items(i).AssignedVendor & "|" & results(i).DaysNeeded
                results(i).LineNumber = 0
                results(i).DaysNeeded = 0
            End If
        Else
            Dim vendorIdx As Integer
            Dim selectReason As String
            vendorIdx = SelectVendorWithReason(items(i), selectReason)
            
            If vendorIdx > 0 Then
                results(i).VendorName = Vendors(vendorIdx).Name
                results(i).DaysNeeded = CalculateDaysNeeded(items(i).Quantity, Vendors(vendorIdx).DailyCapacity)
                
                If Not FindAvailableSlot(results(i), items(i).TransferDate, Year(minDate), maxDate) Then
                    results(i).VendorName = "배정불가"
                    results(i).FailReason = "NO_SLOT|" & Vendors(vendorIdx).Name & "|" & results(i).DaysNeeded
                    results(i).LineNumber = 0
                    results(i).DaysNeeded = 0
                End If
            Else
                results(i).VendorName = "배정불가"
                results(i).FailReason = selectReason
                results(i).LineNumber = 0
                results(i).StartDate = Date
                results(i).DaysNeeded = 0
            End If
        End If
NextItem:
    Next i
End Sub

Private Function FindAvailableSlot(ByRef result As AllocationResult, ByVal transferDateStr As String, _
                                   ByVal targetYear As Integer, ByVal maxDate As Date) As Boolean
    FindAvailableSlot = False
    
    Dim vendorName As String: vendorName = result.VendorName
    Dim daysNeeded As Integer: daysNeeded = result.DaysNeeded
    Dim lineCount As Integer: lineCount = GetVendorLineCount(vendorName)
    If lineCount = 0 Then lineCount = 1
    
    Dim minStartDate As Date
    minStartDate = GetProductionStartDate(transferDateStr, targetYear)
    
    Dim searchDate As Date: searchDate = minStartDate
    
    Do While searchDate <= maxDate
        If Weekday(searchDate) = 1 Or Weekday(searchDate) = 7 Then
            searchDate = searchDate + 1
            GoTo ContinueSearch
        End If
        
        Dim line As Integer
        For line = 1 To lineCount
            If CanPlaceOnLine(vendorName, line, searchDate, daysNeeded, maxDate) Then
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
End Function

Private Function CanPlaceOnLine(ByVal vendorName As String, ByVal line As Integer, _
                                ByVal startDate As Date, ByVal daysNeeded As Integer, _
                                ByVal maxDate As Date) As Boolean
    CanPlaceOnLine = True
    
    Dim checkDate As Date: checkDate = startDate
    Dim placedDays As Integer: placedDays = 0
    
    Do While placedDays < daysNeeded
        If checkDate > maxDate Then
            CanPlaceOnLine = False
            Exit Function
        End If
        
        If Weekday(checkDate) = 1 Or Weekday(checkDate) = 7 Then
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

Private Sub ReserveLine(ByVal vendorName As String, ByVal line As Integer, _
                        ByVal startDate As Date, ByVal daysNeeded As Integer)
    Dim checkDate As Date: checkDate = startDate
    Dim placedDays As Integer: placedDays = 0
    
    Do While placedDays < daysNeeded
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

Private Function SelectVendor(ByRef Item As ProductionItem) As Integer
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

Private Function SelectVendorWithReason(ByRef Item As ProductionItem, ByRef reason As String) As Integer
    SelectVendorWithReason = 0
    reason = ""
    
    ' 고객사 매칭 우선 시도
    If Len(Item.ClientCode) > 0 Then
        If ClientMappings.Exists(Item.ClientCode) Then
            Dim mappedVendorId As String
            mappedVendorId = ClientMappings(Item.ClientCode)
            
            Dim j As Integer
            For j = 1 To VendorCount
                If Vendors(j).ID = mappedVendorId Then
                    If CanHandleProcess(Vendors(j), Item.SpecialProcess) Then
                        SelectVendorWithReason = j
                        Exit Function
                    Else
                        ' 매칭 업체가 특수공정 불가
                        reason = "SPECIAL_PROCESS|" & Vendors(j).Name & "|" & Item.SpecialProcess
                        Exit Function
                    End If
                End If
            Next j
        End If
    End If
    
    ' 일반 배정 시도
    Dim bestVendor As Integer: bestVendor = 0
    Dim bestPriority As Integer: bestPriority = 999
    Dim hasSpecialCapable As Boolean: hasSpecialCapable = False
    
    For j = 1 To VendorCount
        If CanHandleProcess(Vendors(j), Item.SpecialProcess) Then
            hasSpecialCapable = True
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
    
    If bestVendor = 0 Then
        If Len(Item.SpecialProcess) > 0 And Item.SpecialProcess <> "normal" Then
            reason = "NO_SPECIAL_VENDOR|" & Item.SpecialProcess
        Else
            reason = "NO_VENDOR"
        End If
    End If
    
    SelectVendorWithReason = bestVendor
End Function

Private Function CanHandleProcess(ByRef vendor As VendorInfo, ByVal ProcessType As String) As Boolean
    If ProcessType = "normal" Or Len(ProcessType) = 0 Then
        CanHandleProcess = True
        Exit Function
    End If
    CanHandleProcess = InStr(1, vendor.Capabilities, ProcessType, vbTextCompare) > 0
End Function

'===============================================================================
' 수동 지정 경고 체크
' - 고객매칭과 다른 업체 지정 시 → 노란색 경고
' - 특수공정 불가 업체 지정 시 → 빨간색 경고
'===============================================================================
Private Sub CheckManualAssignmentWarning(ByRef result As AllocationResult, ByRef Item As ProductionItem)
    result.HasWarning = False
    result.WarningType = ""
    result.WarningMessage = ""
    
    Dim manualVendorName As String: manualVendorName = Item.AssignedVendor
    Dim manualVendorIdx As Integer: manualVendorIdx = GetVendorIndexByName(manualVendorName)
    
    If manualVendorIdx = 0 Then Exit Sub
    
    ' 1) 특수공정 불가 체크 (빨간색 - 더 심각)
    If Len(Item.SpecialProcess) > 0 And Item.SpecialProcess <> "normal" Then
        If Not CanHandleProcess(Vendors(manualVendorIdx), Item.SpecialProcess) Then
            result.HasWarning = True
            result.WarningType = "PROCESS_UNABLE"
            result.WarningMessage = manualVendorName & "은(는) " & GetProcessName(Item.SpecialProcess) & " 공정 불가"
            Exit Sub
        End If
    End If
    
    ' 2) 고객매칭과 다른지 체크 (노란색)
    If Len(Item.ClientCode) > 0 Then
        If ClientMappings.Exists(Item.ClientCode) Then
            Dim mappedVendorId As String
            mappedVendorId = ClientMappings(Item.ClientCode)
            
            ' 수동 지정 업체의 ID 가져오기
            Dim manualVendorId As String
            manualVendorId = Vendors(manualVendorIdx).ID
            
            ' 매칭 업체와 다르면 경고
            If mappedVendorId <> manualVendorId Then
                Dim mappedVendorName As String
                mappedVendorName = GetVendorNameById(mappedVendorId)
                
                result.HasWarning = True
                result.WarningType = "CLIENT_MISMATCH"
                result.WarningMessage = "고객매칭: " & mappedVendorName & " → 수동지정: " & manualVendorName
            End If
        End If
    End If
End Sub

' 외주처명으로 인덱스 찾기
Private Function GetVendorIndexByName(ByVal vendorName As String) As Integer
    GetVendorIndexByName = 0
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).Name = vendorName Then
            GetVendorIndexByName = i
            Exit Function
        End If
    Next i
End Function

' 외주처ID로 이름 찾기
Private Function GetVendorNameById(ByVal vendorId As String) As String
    GetVendorNameById = vendorId
    Dim i As Integer
    For i = 1 To VendorCount
        If Vendors(i).ID = vendorId Then
            GetVendorNameById = Vendors(i).Name
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' 연속 달력 뷰 (월 구분 없이 쭉 이어지게)
'===============================================================================
Private Sub CreateContinuousCalendarView(ByRef results() As AllocationResult, ByVal itemCount As Long, _
                               ByVal minDate As Date, ByVal maxDate As Date)
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_OUTPUT)
    
    ws.Cells.Font.Name = "맑은 고딕"
    ws.Cells.Font.Size = 9
    
    ' 기존 메모 모두 삭제
    On Error Resume Next
    ws.Cells.ClearComments
    On Error GoTo 0
    
    Dim currentRow As Integer: currentRow = 1
    
    ' 전체 기간 표시
    ws.Cells(currentRow, 1).Value = "생산계획표"
    ws.Cells(currentRow, 1).Font.Size = 20
    ws.Cells(currentRow, 1).Font.Bold = True
    ws.Cells(currentRow, 1).Font.Color = RGB(15, 23, 42)
    currentRow = currentRow + 1
    
    ws.Cells(currentRow, 1).Value = Format(minDate, "yyyy년 m월") & " ~ " & Format(maxDate, "yyyy년 m월") & "  |  Updated: " & Format(Now, "yyyy-mm-dd hh:mm")
    ws.Cells(currentRow, 1).Font.Color = RGB(100, 116, 139)
    ws.Cells(currentRow, 1).Font.Size = 10
    currentRow = currentRow + 2
    
    ' 전체 날짜 수 계산
    Dim totalDays As Long
    totalDays = DateDiff("d", minDate, maxDate) + 1
    
    Dim headerRow As Integer: headerRow = currentRow
    
    ' A열 헤더
    ws.Cells(headerRow, 1).Value = "외주처 / 라인"
    ws.Cells(headerRow, 1).Interior.Color = RGB(51, 65, 85)
    ws.Cells(headerRow, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(headerRow, 1).Font.Bold = True
    ws.Cells(headerRow, 1).Font.Size = 11
    ws.Columns(1).ColumnWidth = 15
    
    ' 날짜 헤더 (파스텔 톤)
    Dim d As Long
    Dim currentDate As Date
    For d = 1 To totalDays
        currentDate = minDate + d - 1
        Dim dow As Integer: dow = Weekday(currentDate)
        Dim col As Long: col = d + 1
        
        ' 월 변경 시 월 표시 추가
        Dim headerText As String
        If Day(currentDate) = 1 Or d = 1 Then
            headerText = Format(currentDate, "m월") & vbLf & Day(currentDate) & vbLf & GetDayName(dow)
        Else
            headerText = Day(currentDate) & vbLf & GetDayName(dow)
        End If
        
        ws.Cells(headerRow, col).Value = headerText
        ws.Cells(headerRow, col).WrapText = True
        ws.Cells(headerRow, col).HorizontalAlignment = xlCenter
        ws.Cells(headerRow, col).VerticalAlignment = xlCenter
        ws.Cells(headerRow, col).Font.Bold = True
        ws.Cells(headerRow, col).Font.Size = 9
        
        ' 파스텔 톤 색상
        If dow = 1 Or dow = 7 Then
            ' 주말: 파스텔 핑크/살몬
            ws.Cells(headerRow, col).Interior.Color = RGB(254, 205, 211)
            ws.Cells(headerRow, col).Font.Color = RGB(159, 18, 57)
        ElseIf currentDate = Date Then
            ' 오늘: 파스텔 민트
            ws.Cells(headerRow, col).Interior.Color = RGB(167, 243, 208)
            ws.Cells(headerRow, col).Font.Color = RGB(6, 95, 70)
        Else
            ' 평일: 파스텔 블루/그레이
            ws.Cells(headerRow, col).Interior.Color = RGB(226, 232, 240)
            ws.Cells(headerRow, col).Font.Color = RGB(51, 65, 85)
        End If
        
        ' B열 이후 모두 15로 통일
        ws.Columns(col).ColumnWidth = 15
    Next d
    
    ws.Rows(headerRow).RowHeight = 38
    
    ' 외주처별 행
    currentRow = headerRow + 1
    Dim v As Integer
    
    For v = 1 To VendorCount
        Dim vendorName As String: vendorName = Vendors(v).Name
        
        ' 해당 업체에 배정된 품목 있는지 확인
        Dim hasItems As Boolean: hasItems = False
        Dim i As Long
        For i = 1 To itemCount
            If results(i).VendorName = vendorName Then
                hasItems = True
                Exit For
            End If
        Next i
        
        If Not hasItems Then GoTo NextVendor
        
        ' 외주처 헤더 (큰 글씨)
        ws.Cells(currentRow, 1).Value = vendorName
        ws.Cells(currentRow, 1).Font.Bold = True
        ws.Cells(currentRow, 1).Font.Size = 12
        ws.Cells(currentRow, 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(currentRow, 1).Interior.Color = GetVendorColor(vendorName)
        ws.Cells(currentRow, 1).HorizontalAlignment = xlCenter
        ws.Cells(currentRow, 1).VerticalAlignment = xlCenter
        
        For d = 1 To totalDays
            ws.Cells(currentRow, d + 1).Interior.Color = GetVendorPaleColor(vendorName)
        Next d
        ws.Rows(currentRow).RowHeight = 24
        currentRow = currentRow + 1
        
        ' 라인별
        Dim line As Integer
        For line = 1 To Vendors(v).LineCount
            ws.Cells(currentRow, 1).Value = "Line " & line
            ws.Cells(currentRow, 1).Font.Color = RGB(51, 65, 85)
            ws.Cells(currentRow, 1).Font.Size = 11
            ws.Cells(currentRow, 1).Font.Bold = True
            ws.Cells(currentRow, 1).Interior.Color = RGB(241, 245, 249)
            ws.Cells(currentRow, 1).HorizontalAlignment = xlCenter
            ws.Cells(currentRow, 1).VerticalAlignment = xlCenter
            
            For d = 1 To totalDays
                currentDate = minDate + d - 1
                dow = Weekday(currentDate)
                col = d + 1
                
                If dow = 1 Or dow = 7 Then
                    ws.Cells(currentRow, col).Interior.Color = RGB(254, 242, 242)
                Else
                    ws.Cells(currentRow, col).Interior.Color = RGB(255, 255, 255)
                End If
            Next d
            
            ' 품목 표시
            For i = 1 To itemCount
                If results(i).VendorName = vendorName And results(i).LineNumber = line Then
                    PlaceItemOnContinuousCalendar ws, results(i), currentRow, minDate, totalDays
                End If
            Next i
            
            ws.Rows(currentRow).RowHeight = 55
            currentRow = currentRow + 1
        Next line
        
NextVendor:
    Next v
    
    ' 굵은 테두리 적용
    If currentRow > headerRow + 1 Then
        With ws.Range(ws.Cells(headerRow, 1), ws.Cells(currentRow - 1, totalDays + 1))
            ' 외곽 테두리 - 굵게
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeLeft).Color = RGB(51, 65, 85)
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThick
            .Borders(xlEdgeRight).Color = RGB(51, 65, 85)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeTop).Color = RGB(51, 65, 85)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeBottom).Color = RGB(51, 65, 85)
            ' 내부 테두리 - 중간 굵기
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Weight = xlMedium
            .Borders(xlInsideVertical).Color = RGB(203, 213, 225)
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlMedium
            .Borders(xlInsideHorizontal).Color = RGB(203, 213, 225)
        End With
    End If
    
    ' 품목 셀에 굵은 테두리 적용 (전체 테두리 후에 적용해야 덮이지 않음)
    ApplyItemCellBorders ws, headerRow, currentRow - 1, totalDays + 1
    
    ' 배정불가 섹션
    currentRow = currentRow + 2
    currentRow = CreateUnassignedSection(ws, results, itemCount, currentRow)
    
    ' 틀 고정
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("B5").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

'===============================================================================
' 품목을 연속 달력에 배치
'===============================================================================
Private Sub PlaceItemOnContinuousCalendar(ByRef ws As Worksheet, ByRef result As AllocationResult, _
                                ByVal rowNum As Integer, ByVal minDate As Date, ByVal totalDays As Long)
    
    If result.DaysNeeded = 0 Then Exit Sub
    
    Dim dailyCapacity As Long: dailyCapacity = GetVendorDailyCapacity(result.VendorName)
    Dim remainingQty As Long: remainingQty = result.Item.Quantity
    Dim currentDate As Date: currentDate = result.StartDate
    Dim dayNum As Integer: dayNum = 1
    
    Do While remainingQty > 0
        ' 날짜가 범위 내인지 확인
        Dim dayOffset As Long
        dayOffset = DateDiff("d", minDate, currentDate)
        If dayOffset < 0 Or dayOffset >= totalDays Then
            currentDate = currentDate + 1
            GoTo ContinuePlace
        End If
        
        If Weekday(currentDate) = 1 Or Weekday(currentDate) = 7 Then
            currentDate = currentDate + 1
            GoTo ContinuePlace
        End If
        
        Dim col As Long: col = dayOffset + 2
        Dim dailyQty As Long: dailyQty = remainingQty
        If dailyQty > dailyCapacity Then dailyQty = dailyCapacity
        
        ' =====================================================================
        ' 셀 내용 구성 (품목코드 + 품목명(축약) + 수량)
        ' =====================================================================
        Dim prodCode As String: prodCode = result.Item.ProductCode
        If Len(prodCode) > 12 Then prodCode = Left(prodCode, 12)
        
        Dim prodName As String: prodName = result.Item.ProductName
        If Len(prodName) > 8 Then prodName = Left(prodName, 7) & ".."
        
        Dim cellText As String
        cellText = prodCode & vbLf & prodName & vbLf & Format(dailyQty, "#,##0")
        
        If result.DaysNeeded > 1 Then
            cellText = cellText & " (" & dayNum & "/" & result.DaysNeeded & ")"
        End If
        
        ws.Cells(rowNum, col).Value = cellText
        
        ' =====================================================================
        ' 셀 스타일 + 굵은 테두리
        ' =====================================================================
        Dim vendorColor As Long: vendorColor = GetVendorColor(result.VendorName)
        Dim lightColor As Long: lightColor = GetVendorLightColor(result.VendorName)
        Dim cellBgColor As Long: cellBgColor = lightColor
        Dim fontColor As Long: fontColor = RGB(30, 41, 59)
        
        ' ★ 경고가 있으면 배경색 변경 (둘 다 빨간색 계열)
        If result.HasWarning Then
            If result.WarningType = "PROCESS_UNABLE" Then
                ' 진한 빨간색 배경 (특수공정 불가 - 더 심각)
                cellBgColor = RGB(254, 202, 202)
                fontColor = RGB(127, 29, 29)
            ElseIf result.WarningType = "CLIENT_MISMATCH" Then
                ' 빨간색 배경 (고객매칭 불일치)
                cellBgColor = RGB(254, 226, 226)
                fontColor = RGB(153, 27, 27)
            End If
        End If
        
        With ws.Cells(rowNum, col)
            ' 배경
            .Interior.Color = cellBgColor
            
            ' 폰트
            .Font.Size = 8
            .Font.Color = fontColor
            .Font.Bold = False
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .WrapText = True
        End With
        
        ' =====================================================================
        ' 마우스 호버 시 상세정보 (메모/Comment)
        ' =====================================================================
        Dim commentText As String
        
        ' ★ 경고가 있으면 상단에 경고 메시지 표시
        If result.HasWarning Then
            If result.WarningType = "PROCESS_UNABLE" Then
                commentText = "⚠️ [ 경 고 ] ⚠️" & vbLf
                commentText = commentText & result.WarningMessage & vbLf
                commentText = commentText & "═════════════" & vbLf & vbLf
            ElseIf result.WarningType = "CLIENT_MISMATCH" Then
                commentText = "⚡ [ 주 의 ] ⚡" & vbLf
                commentText = commentText & result.WarningMessage & vbLf
                commentText = commentText & "═════════════" & vbLf & vbLf
            End If
        End If
        
        commentText = commentText & "[ 상 세 정 보 ]" & vbLf & vbLf
        commentText = commentText & "품목코드: " & result.Item.ProductCode & vbLf
        commentText = commentText & "품목명: " & result.Item.ProductName & vbLf
        commentText = commentText & "─────────────" & vbLf
        commentText = commentText & "총 수량: " & Format(result.Item.Quantity, "#,##0") & vbLf
        commentText = commentText & "금일 수량: " & Format(dailyQty, "#,##0") & vbLf
        commentText = commentText & "─────────────" & vbLf
        commentText = commentText & "납기: " & IIf(Len(result.Item.DeliveryDate) > 0, result.Item.DeliveryDate, "-") & vbLf
        commentText = commentText & "이동일: " & result.Item.TransferDate & vbLf
        commentText = commentText & "외주처: " & result.VendorName & vbLf
        commentText = commentText & "라인: Line " & result.LineNumber & vbLf
        
        If result.DaysNeeded > 1 Then
            commentText = commentText & "─────────────" & vbLf
            commentText = commentText & "진행: " & dayNum & "일차 / " & result.DaysNeeded & "일"
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
' 품목 셀에 굵은 테두리 적용
' - 경고 셀은 배경색에 따라 테두리 색상 변경
'===============================================================================
Private Sub ApplyItemCellBorders(ByRef ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal endCol As Long)
    Dim r As Long, c As Long
    Dim borderColor As Long
    Dim bgColor As Long
    
    ' 경고 배경색 상수 (RGB를 Long으로: R + G*256 + B*65536)
    Const WARNING_RED_DARK_BG As Long = 13289214  ' RGB(254, 202, 202) - 특수공정 불가
    Const WARNING_RED_LIGHT_BG As Long = 14869246 ' RGB(254, 226, 226) - 고객매칭 불일치
    
    For r = startRow + 1 To endRow
        For c = 2 To endCol
            ' 셀에 내용이 있으면 (품목 셀)
            If Len(Trim(ws.Cells(r, c).Value)) > 0 Then
                bgColor = ws.Cells(r, c).Interior.Color
                
                ' ★ 배경색에 따라 테두리 색상 결정 (둘 다 빨간색)
                If bgColor = WARNING_RED_DARK_BG Then
                    ' 진한 빨간 테두리 (특수공정 불가)
                    borderColor = RGB(185, 28, 28)
                ElseIf bgColor = WARNING_RED_LIGHT_BG Then
                    ' 빨간 테두리 (고객매칭 불일치)
                    borderColor = RGB(220, 38, 38)
                Else
                    ' 기본 테두리
                    borderColor = RGB(30, 30, 30)
                End If
                
                With ws.Cells(r, c)
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeLeft).Color = borderColor
                    
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlEdgeRight).Color = borderColor
                    
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeTop).Color = borderColor
                    
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).Color = borderColor
                End With
            End If
        Next c
    Next r
End Sub

'===============================================================================
' 배정불가 섹션
'===============================================================================
Private Function CreateUnassignedSection(ByRef ws As Worksheet, ByRef results() As AllocationResult, _
                                         ByVal itemCount As Long, ByVal startRow As Integer) As Integer
    Dim hasUnassigned As Boolean: hasUnassigned = False
    Dim i As Long
    For i = 1 To itemCount
        If results(i).VendorName = "배정불가" Then
            hasUnassigned = True
            Exit For
        End If
    Next i
    
    If Not hasUnassigned Then
        CreateUnassignedSection = startRow
        Exit Function
    End If
    
    ws.Cells(startRow, 1).Value = "배정불가 품목"
    ws.Cells(startRow, 1).Font.Size = 14
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Color = RGB(185, 28, 28)
    startRow = startRow + 1
    
    ' 헤더 (원인, 해결방안 추가)
    ws.Cells(startRow, 1).Value = "품목코드"
    ws.Cells(startRow, 2).Value = "품목명"
    ws.Cells(startRow, 3).Value = "수량"
    ws.Cells(startRow, 4).Value = "납기"
    ws.Cells(startRow, 5).Value = "이동일"
    ws.Cells(startRow, 6).Value = "원인"
    ws.Cells(startRow, 7).Value = "해결방안"
    
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 7))
        .Font.Bold = True
        .Interior.Color = RGB(185, 28, 28)
        .Font.Color = RGB(255, 255, 255)
    End With
    startRow = startRow + 1
    
    For i = 1 To itemCount
        If results(i).VendorName = "배정불가" Then
            ws.Cells(startRow, 1).Value = results(i).Item.ProductCode
            ws.Cells(startRow, 2).Value = results(i).Item.ProductName
            ws.Cells(startRow, 3).Value = results(i).Item.Quantity
            ws.Cells(startRow, 3).NumberFormat = "#,##0"
            ws.Cells(startRow, 4).Value = results(i).Item.DeliveryDate
            ws.Cells(startRow, 5).Value = results(i).Item.TransferDate
            
            ' 원인과 해결방안 파싱
            Dim reasonText As String, solutionText As String
            ParseFailReason results(i).FailReason, results(i).Item, reasonText, solutionText
            
            ws.Cells(startRow, 6).Value = reasonText
            ws.Cells(startRow, 7).Value = solutionText
            
            ' 스타일
            ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 7)).Interior.Color = RGB(254, 242, 242)
            ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 5)).Font.Color = RGB(153, 27, 27)
            ws.Cells(startRow, 6).Font.Color = RGB(185, 28, 28)
            ws.Cells(startRow, 6).Font.Bold = True
            ws.Cells(startRow, 7).Font.Color = RGB(22, 101, 52)
            ws.Cells(startRow, 7).Font.Bold = True
            
            startRow = startRow + 1
        End If
    Next i
    
    ' 열 너비 조정
    ws.Columns(6).ColumnWidth = 35
    ws.Columns(7).ColumnWidth = 45
    
    CreateUnassignedSection = startRow
End Function

'===============================================================================
' 실패 원인 파싱 및 해결방안 생성
'===============================================================================
Private Sub ParseFailReason(ByVal failReason As String, ByRef Item As ProductionItem, _
                            ByRef reasonText As String, ByRef solutionText As String)
    reasonText = ""
    solutionText = ""
    
    If Len(failReason) = 0 Then
        reasonText = "알 수 없음"
        solutionText = "관리자 확인 필요"
        Exit Sub
    End If
    
    Dim parts() As String
    parts = Split(failReason, "|")
    Dim reasonCode As String: reasonCode = parts(0)
    
    Select Case reasonCode
        Case "INVALID_DATE"
            reasonText = "이동일 형식 오류 (" & Item.TransferDate & ")"
            solutionText = "이동일을 올바른 날짜 형식으로 수정 (예: 1/20, 1월20일)"
            
        Case "NO_SLOT"
            ' NO_SLOT|업체명|필요일수
            Dim vendorName As String: vendorName = ""
            Dim daysNeeded As String: daysNeeded = ""
            If UBound(parts) >= 1 Then vendorName = parts(1)
            If UBound(parts) >= 2 Then daysNeeded = parts(2)
            reasonText = vendorName & " 슬롯 부족 (" & daysNeeded & "일 필요)"
            solutionText = "1) 이동일 앞당기기 2) D열에 다른 외주처 지정 3) 수량 분할"
            
        Case "NO_SLOT_MANUAL"
            ' NO_SLOT_MANUAL|업체명|필요일수
            If UBound(parts) >= 1 Then vendorName = parts(1)
            If UBound(parts) >= 2 Then daysNeeded = parts(2)
            reasonText = "지정 외주처(" & vendorName & ") 슬롯 부족 (" & daysNeeded & "일 필요)"
            solutionText = "1) D열 외주처 지정 해제 (자동배정) 2) 이동일 앞당기기"
            
        Case "SPECIAL_PROCESS"
            ' SPECIAL_PROCESS|업체명|공정타입
            If UBound(parts) >= 1 Then vendorName = parts(1)
            Dim processType As String: processType = ""
            If UBound(parts) >= 2 Then processType = parts(2)
            Dim processName As String: processName = GetProcessName(processType)
            reasonText = "매칭업체(" & vendorName & ")가 " & processName & " 공정 불가"
            solutionText = "1) D열에 " & processName & " 가능 업체 직접 지정 2) C열 특수공정 해제"
            
        Case "NO_SPECIAL_VENDOR"
            ' NO_SPECIAL_VENDOR|공정타입
            If UBound(parts) >= 1 Then processType = parts(1)
            processName = GetProcessName(processType)
            reasonText = processName & " 공정 가능 업체 없음"
            solutionText = "1) 설정_외주처 시트에서 해당 공정 가능 업체 추가 2) C열 특수공정 해제"
            
        Case "NO_VENDOR"
            reasonText = "배정 가능한 외주처 없음"
            solutionText = "설정_외주처 시트 확인 및 업체 추가"
            
        Case Else
            reasonText = failReason
            solutionText = "관리자 확인 필요"
    End Select
End Sub

Private Function GetProcessName(ByVal processType As String) As String
    Select Case processType
        Case "shrink": GetProcessName = "수축"
        Case "mixing": GetProcessName = "교반"
        Case "highFrequency": GetProcessName = "고주파"
        Case Else: GetProcessName = processType
    End Select
End Function

'===============================================================================
' 설정 시트
'===============================================================================
Private Sub CreateVendorSheet()
    Dim ws As Worksheet
    Set ws = CreateOrClearSheet(SHEET_VENDORS)
    
    ws.Range("A1:G1").Value = Array("외주처ID", "외주처명", "라인수", "라인당일일생산량", "가능공정", "월생산목표", "우선순위")
    
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
        ' 1코드(충전반제품) 제외 - 9코드(완제품)만 배정
        If Left(productCode, 1) = "1" Then GoTo NextRow
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
        ' 1코드(충전반제품) 제외 - 9코드(완제품)만 배정
        If Left(productCode, 1) = "1" Then GoTo NextRow2
        
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
        ws.Cells.ClearComments
        ActiveWindow.FreezePanes = False
    End If
    Set CreateOrClearSheet = ws
End Function

Private Sub FormatSettingSheet(ByRef ws As Worksheet, ByVal colCount As Integer, ByVal rowCount As Integer)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 58, 138)
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(203, 213, 225)
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

Private Function ParseTransferDate(ByVal transferDateStr As String, ByVal targetYear As Integer) As Date
    On Error Resume Next
    ParseTransferDate = 0
    
    If InStr(transferDateStr, "/") > 0 Then
        Dim parts() As String: parts = Split(transferDateStr, "/")
        If UBound(parts) >= 1 Then
            ParseTransferDate = DateSerial(targetYear, CInt(parts(0)), CInt(parts(1)))
        End If
    ElseIf InStr(transferDateStr, "월") > 0 Then
        Dim monthPart As String, dayPart As String
        monthPart = Left(transferDateStr, InStr(transferDateStr, "월") - 1)
        dayPart = Mid(transferDateStr, InStr(transferDateStr, "월") + 1)
        dayPart = Replace(dayPart, "일", "")
        ParseTransferDate = DateSerial(targetYear, CInt(Trim(monthPart)), CInt(Trim(dayPart)))
    Else
        ParseTransferDate = CDate(transferDateStr)
    End If
    On Error GoTo 0
End Function

Private Function GetProductionStartDate(ByVal transferDateStr As String, ByVal targetYear As Integer) As Date
    Dim transferDate As Date
    transferDate = ParseTransferDate(transferDateStr, targetYear)
    
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
        Case "위드맘": GetVendorLightColor = RGB(219, 234, 254)
        Case "리니어": GetVendorLightColor = RGB(220, 252, 231)
        Case "그램": GetVendorLightColor = RGB(243, 232, 255)
        Case "이시스": GetVendorLightColor = RGB(255, 237, 213)
        Case "엘루오": GetVendorLightColor = RGB(252, 231, 243)
        Case "케이코스텍": GetVendorLightColor = RGB(207, 250, 254)
        Case "다미": GetVendorLightColor = RGB(254, 249, 195)
        Case "배정불가": GetVendorLightColor = RGB(241, 245, 249)
        Case Else: GetVendorLightColor = RGB(241, 245, 249)
    End Select
End Function

Private Function GetVendorPaleColor(ByVal vendorName As String) As Long
    Select Case vendorName
        Case "위드맘": GetVendorPaleColor = RGB(239, 246, 255)
        Case "리니어": GetVendorPaleColor = RGB(240, 253, 244)
        Case "그램": GetVendorPaleColor = RGB(250, 245, 255)
        Case "이시스": GetVendorPaleColor = RGB(255, 247, 237)
        Case "엘루오": GetVendorPaleColor = RGB(253, 242, 248)
        Case "케이코스텍": GetVendorPaleColor = RGB(236, 254, 255)
        Case "다미": GetVendorPaleColor = RGB(254, 252, 232)
        Case "배정불가": GetVendorPaleColor = RGB(248, 250, 252)
        Case Else: GetVendorPaleColor = RGB(248, 250, 252)
    End Select
End Function

' 색상 어둡게 (그림자 효과용)
Private Function DarkenColor(ByVal Color As Long) As Long
    Dim r As Integer, g As Integer, b As Integer
    r = Color Mod 256
    g = (Color \ 256) Mod 256
    b = (Color \ 65536) Mod 256
    
    r = r - 30: If r < 0 Then r = 0
    g = g - 30: If g < 0 Then g = 0
    b = b - 30: If b < 0 Then b = 0
    
    DarkenColor = RGB(r, g, b)
End Function
