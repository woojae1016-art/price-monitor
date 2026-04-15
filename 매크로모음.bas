Attribute VB_Name = "매크로모음"
' =============================================
' 1. 하이퍼링크 일괄 열기
' =============================================
Sub OpenHyperlinksInRange()
    Dim cell As Range
    Dim selectedRange As Range
    Dim visibleRange As Range
    Dim ws As Worksheet

    Set ws = ActiveSheet

    On Error Resume Next
    Set selectedRange = Application.InputBox("하이퍼링크를 열기 위한 셀 범위를 선택하세요:", Type:=8)
    On Error GoTo 0

    If selectedRange Is Nothing Then
        MsgBox "유효한 셀 범위가 선택되지 않았습니다."
        Exit Sub
    End If

    On Error Resume Next
    Set visibleRange = selectedRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If visibleRange Is Nothing Then
        MsgBox "선택한 범위 내에 표시된 셀이 없습니다."
        Exit Sub
    End If

    For Each cell In visibleRange
        If cell.Hyperlinks.Count > 0 Then
            cell.Hyperlinks(1).Follow
        End If
    Next cell
End Sub

' =============================================
' 2. 168/169행 + DE/DF열 계산 후
'    C2점 권장가 위반 정리 시트 생성
'
' ※ 실행 전: 오픈마켓확인 시트의 쿠팡 G열(판매자)을 직접 입력 후 실행
' =============================================
Sub RunIntegratedProcess()
    Dim wsResult     As Worksheet
    Dim wsTarget     As Worksheet
    Dim lastCol      As Long
    Dim lastDataRow  As Long
    Dim targetRow    As Long
    Dim dataStartRow As Long
    Dim dataEndRow   As Long
    Dim extractRow   As Long
    Dim i            As Long
    Dim r            As Long
    Dim c            As Long
    Dim row          As Long
    Dim cellValue    As Variant
    Dim currentDate  As String
    Dim filterRow    As Range
    Dim dfValue      As String
    Dim dealerList   As Variant
    Dim dealer       As String
    Dim dataRange    As Range

    Application.ScreenUpdating = False
    currentDate = Format(Now, "yyyy.mm.dd")

    ' ── 통합결과 시트 확인 ──────────────────────
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("통합결과")
    On Error GoTo 0
    If wsResult Is Nothing Then
        MsgBox "통합결과 시트를 찾을 수 없습니다.", vbExclamation
        GoTo CleanUp
    End If

    ' DE/DF열 추가 전 실제 데이터 마지막 열 (108열 이하로 제한)
    lastCol = wsResult.Cells(1, wsResult.Columns.Count).End(xlToLeft).Column
    If lastCol > 108 Then lastCol = 108
    lastDataRow = wsResult.Cells(wsResult.Rows.Count, 1).End(xlUp).row

    ' ══════════════════════════════════════════
    ' STEP 1) DE(109)/DF(110)열 계산
    '         각 모델(3행묶음)의 전체 판매처 중 최저가+판매자 기록
    ' ══════════════════════════════════════════
    ' 헤더
    wsResult.Cells(1, 109).Value = "최저가"
    wsResult.Cells(2, 109).Value = "최저가격"
    wsResult.Cells(2, 110).Value = "c2 or WILO"

    Dim rowIdx   As Long
    Dim colIdx   As Long
    Dim minPrice As Variant
    Dim minSeller As String
    Dim minMsrp  As Variant
    Dim minDc    As Variant
    Dim lp       As Variant

    rowIdx = 3
    Do While rowIdx <= lastDataRow
        If wsResult.Cells(rowIdx, 1).Value = "" Then
            rowIdx = rowIdx + 3
        Else
            minPrice  = Null
            minSeller = ""
            minMsrp   = 0
            minDc     = 0

            For colIdx = 3 To lastCol
                lp = wsResult.Cells(rowIdx, colIdx).Value
                If IsNumeric(lp) And lp > 0 Then
                    If IsNull(minPrice) Or lp < minPrice Then
                        minPrice  = lp
                        minSeller = wsResult.Cells(1, colIdx).Value
                        minMsrp   = wsResult.Cells(rowIdx + 1, colIdx).Value
                        minDc     = wsResult.Cells(rowIdx + 2, colIdx).Value
                    End If
                End If
            Next colIdx

            If Not IsNull(minPrice) Then
                wsResult.Cells(rowIdx,     109).Value         = minPrice
                wsResult.Cells(rowIdx + 1, 109).Value         = minMsrp
                wsResult.Cells(rowIdx + 2, 109).Value         = minDc
                wsResult.Cells(rowIdx,     110).Value         = minSeller
                wsResult.Cells(rowIdx,     109).NumberFormat  = "#,##0"
                wsResult.Cells(rowIdx + 1, 109).NumberFormat  = "#,##0"
                wsResult.Cells(rowIdx + 2, 109).NumberFormat  = "0.0%"
            Else
                wsResult.Cells(rowIdx,     109).Value = 0
                wsResult.Cells(rowIdx + 1, 109).Value = 0
                wsResult.Cells(rowIdx + 2, 109).Value = 0
                wsResult.Cells(rowIdx,     110).Value = "wilo"
            End If

            ' 테두리
            Dim rr As Long, cc As Long
            For rr = rowIdx To rowIdx + 2
                For cc = 109 To 110
                    With wsResult.Cells(rr, cc).Borders
                        .LineStyle = xlContinuous
                        .Weight    = xlThin
                    End With
                Next cc
            Next rr

            rowIdx = rowIdx + 3
        End If
    Loop

    ' 헤더 서식
    For cc = 109 To 110
        For rr = 1 To 2
            With wsResult.Cells(rr, cc)
                .Font.Bold = True
                .Interior.Color = RGB(192, 192, 192)
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Weight    = xlThin
            End With
        Next rr
    Next cc

    ' ══════════════════════════════════════════
    ' STEP 2) 168행: 판매처별 평균 DC율 계산
    '         169행: 판매처별 권장가 미만 개수
    ' ══════════════════════════════════════════
    wsResult.Cells(168, 2).Value = "평균 dc 율"
    wsResult.Cells(169, 2).Value = "권장가 미만 개수"

    Dim dcSum   As Double
    Dim dcCount As Long
    Dim dcVal   As Variant
    Dim lpVal   As Variant
    Dim cnt     As Long

    For colIdx = 3 To lastCol
        ' 168행: DC율 평균 (DC율 행: 5, 8, 11 ...)
        dcSum   = 0
        dcCount = 0
        r = 5
        Do While r <= lastDataRow
            dcVal = wsResult.Cells(r, colIdx).Value
            If IsNumeric(dcVal) And dcVal <> 0 Then
                If dcVal < 1 Then
                    dcSum = dcSum + dcVal
                Else
                    dcSum = dcSum + dcVal / 100
                End If
                dcCount = dcCount + 1
            End If
            r = r + 3
        Loop
        If dcCount > 0 Then
            wsResult.Cells(168, colIdx).Value = dcSum / dcCount
        Else
            wsResult.Cells(168, colIdx).Value = 0
        End If
        wsResult.Cells(168, colIdx).NumberFormat = "0.0%"

        ' 169행: 최저가 있는 개수 (최저가 행: 3, 6, 9 ...)
        cnt = 0
        r = 3
        Do While r <= lastDataRow
            lpVal = wsResult.Cells(r, colIdx).Value
            If IsNumeric(lpVal) And lpVal > 0 Then
                cnt = cnt + 1
            End If
            r = r + 3
        Loop
        wsResult.Cells(169, colIdx).Value = cnt
    Next colIdx

    ' 168/169행 헤더 서식
    For rr = 168 To 169
        With wsResult.Cells(rr, 2)
            .Font.Bold = True
            .Interior.Color = RGB(192, 192, 192)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight    = xlThin
        End With
    Next rr
    ' 168/169행 데이터 셀 테두리
    For colIdx = 3 To lastCol
        For rr = 168 To 169
            With wsResult.Cells(rr, colIdx).Borders
                .LineStyle = xlContinuous
                .Weight    = xlThin
            End With
        Next rr
    Next colIdx

    ' ══════════════════════════════════════════
    ' STEP 3) C2점 권장가 위반 정리 시트 생성
    ' ══════════════════════════════════════════
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("C2점 권장가 위반 정리")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Sheets.Add
        wsTarget.Name = "C2점 권장가 위반 정리"
    Else
        wsTarget.Cells.Clear
    End If
    targetRow = 1

    ' 날짜 헤더
    wsTarget.Cells(targetRow, 1).Value = "진행 날짜: " & currentDate
    wsTarget.Cells(targetRow, 1).Font.Bold = True
    targetRow = targetRow + 1

    ' 요약 헤더
    wsTarget.Cells(targetRow, 1).Value = "판매처"
    wsTarget.Cells(targetRow, 2).Value = "대리점"
    wsTarget.Cells(targetRow, 3).Value = "평균 DC율"
    wsTarget.Cells(targetRow, 4).Value = "권장가 미만 개수"
    targetRow = targetRow + 1
    dataStartRow = targetRow

    ' 169행 기반 위반 판매처 집계
    For i = 1 To lastCol
        If i = 109 Or i = 110 Then GoTo NextCol
        cellValue = wsResult.Cells(169, i).Value
        If IsNumeric(cellValue) And cellValue <> 0 Then
            wsTarget.Cells(targetRow, 1).Value = wsResult.Cells(1, i).Value
            wsTarget.Cells(targetRow, 2).Value = wsResult.Cells(2, i).Value
            wsTarget.Cells(targetRow, 3).Value = wsResult.Cells(168, i).Value
            wsTarget.Cells(targetRow, 4).Value = wsResult.Cells(169, i).Value
            targetRow = targetRow + 1
        End If
NextCol:
    Next i
    dataEndRow = targetRow - 1

    ' 요약 서식
    wsTarget.Range("C" & dataStartRow & ":C" & dataEndRow).NumberFormat = "0.0%"
    With wsTarget.Range("A" & dataStartRow - 1 & ":D" & dataEndRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Color     = RGB(0, 0, 0)
        .Borders.Weight    = xlThin
    End With

    ' 요약 대리점 노란색
    dealerList = Array( _
        "서우기업", "LG윌로펌프", "경동기전", "고강C&P", "광진종합상사", "굿펌프", "나인티에스", "대림상사", _
        "대영상사", "대풍상사", "미라클YT펌프", "삼흥E&P", "서울종합펌프", "서울펌프랜드", "세광사", _
        "수중모터주식회사", "시대상사", "에스에이치테크", "엘지산업", "윌로종합상사 영천", "이조", _
        "이피컴퍼니", "전진", "주식회사 리텍솔루션", "주식회사 세종종합상사", "카토건설중기", "투빈", _
        "퍼맥스", "펌스", "하경상사", "국제티에스", "광명상사", "희성산업", "펌프랜드", "대산종합상사" _
    )
    For row = dataStartRow To dataEndRow
        dealer = wsTarget.Cells(row, 2).Value
        If Not IsError(Application.Match(dealer, dealerList, 0)) Then
            wsTarget.Range("A" & row & ":D" & row).Interior.Color = RGB(255, 255, 0)
        End If
    Next row

    ' 상세 테이블 헤더
    extractRow = dataEndRow + 3
    wsTarget.Cells(extractRow, 1).Value = "모델명"
    wsTarget.Cells(extractRow, 2).Value = "최저가"
    wsTarget.Cells(extractRow, 3).Value = "DC율"
    wsTarget.Cells(extractRow, 4).Value = "대리점명"
    extractRow = extractRow + 1

    ' DE/DF열 기반 상세 데이터 추출
    Set dataRange = wsResult.Range("A3:A" & lastDataRow)
    For Each filterRow In dataRange.Rows
        dfValue = Trim(wsResult.Cells(filterRow.row, 110).Value)
        If LCase(dfValue) <> "wilo" And LCase(dfValue) <> "c2" And dfValue <> "" Then
            Dim dePrice As Variant
            dePrice = wsResult.Cells(filterRow.row, 109).Value
            If IsNumeric(dePrice) And dePrice > 0 Then
                wsTarget.Cells(extractRow, 1).Value = wsResult.Cells(filterRow.row, 1).Value
                wsTarget.Cells(extractRow, 2).Value = dePrice
                wsTarget.Cells(extractRow, 3).Value = wsResult.Cells(filterRow.row + 2, 109).Value
                wsTarget.Cells(extractRow, 4).Value = dfValue
                extractRow = extractRow + 1
            End If
        End If
    Next filterRow

    ' 상세 서식
    If extractRow > dataEndRow + 5 Then
        wsTarget.Range("B" & dataEndRow + 4 & ":B" & extractRow - 1).NumberFormat = "#,##0"
        wsTarget.Range("C" & dataEndRow + 4 & ":C" & extractRow - 1).NumberFormat = "0.0%"
        With wsTarget.Range("A" & dataEndRow + 3 & ":D" & extractRow - 1)
            .Borders.LineStyle = xlContinuous
            .Borders.Color     = RGB(0, 0, 0)
            .Borders.Weight    = xlThin
        End With
    End If

    ' 상세 대리점 노란색
    dealerList = Array( _
        "서우기업", "윌로펌프백화점", "오아시스 펌프", "서울피엠", "펌프365", "윌로펌프총판", "펌프굿", _
        "나인티에스", "펌프파트너", "이엔지마켓", "따뜻함", "펌프산업", "워터테크", "펌프닷컴", _
        "윌로프로", "샌프란시스코2", "pump-damoa", "서울펌프몰", "윌로공식 SKS윌로펌프", _
        "수중모터주식회사", "시대몰", "펌프프렌드", "윌로펌프마켓", "윌로종합", "윌로펌프모터", _
        "이피컴퍼니", "EP COMPANY", "펌프몰", "윌로펌프온라인쇼핑몰", "주식회사 리텍솔루션", _
        "주식회사 세종종합상사", "여담고", "주식회사 투빈", "펌프의 모든 것", "펌프뱅크", _
        "펌스pums", "펌프탑", "윈디샵", "신세계몰", "광명상사", "신한일전기공식인증몰", _
        "펌프마스터", "대산공구" _
    )
    For row = dataEndRow + 4 To extractRow - 1
        dealer = wsTarget.Cells(row, 4).Value
        If Not IsError(Application.Match(dealer, dealerList, 0)) Then
            wsTarget.Range("A" & row & ":D" & row).Interior.Color = RGB(255, 255, 0)
        End If
    Next row

    wsTarget.Columns("A").AutoFit
    wsTarget.Columns("B:D").AutoFit

CleanUp:
    Application.ScreenUpdating = True
    MsgBox "완료! C2점 권장가 위반 정리 시트가 생성되었습니다.", vbInformation
End Sub

' =============================================
' 3. 테두리만 적용 (별도 실행용)
' =============================================
Sub ApplyBordersOnly()
    Dim wsResult As Worksheet
    Dim lastCol  As Long
    Dim lastRow  As Long
    Dim i As Long, c As Long

    Application.ScreenUpdating = False

    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("통합결과")
    On Error GoTo 0
    If wsResult Is Nothing Then
        MsgBox "통합결과 시트를 찾을 수 없습니다.", vbExclamation
        GoTo Done
    End If

    lastCol = wsResult.Cells(1, wsResult.Columns.Count).End(xlToLeft).Column
    If lastCol > 108 Then lastCol = 108
    lastRow = wsResult.Cells(wsResult.Rows.Count, 1).End(xlUp).Row

    With wsResult.Range(wsResult.Cells(1, 109), wsResult.Cells(lastRow, 110))
        .Borders.LineStyle = xlContinuous
        .Borders.Color     = RGB(0, 0, 0)
        .Borders.Weight    = xlThin
    End With

    For i = 1 To 2
        For c = 2 To lastCol
            If c <> 109 And c <> 110 Then
                With wsResult.Cells(167 + i, c).Borders
                    .LineStyle = xlContinuous
                    .Weight    = xlThin
                End With
            End If
        Next c
    Next i

    Dim wsTarget As Worksheet
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("C2점 권장가 위반 정리")
    On Error GoTo 0
    If Not wsTarget Is Nothing Then
        Dim tLast As Long
        tLast = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
        If tLast >= 2 Then
            With wsTarget.Range("A1:D" & tLast)
                .Borders.LineStyle = xlContinuous
                .Borders.Color     = RGB(0, 0, 0)
                .Borders.Weight    = xlThin
            End With
        End If
    End If

Done:
    Application.ScreenUpdating = True
    MsgBox "테두리 적용 완료!", vbInformation
End Sub
