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

' =============================================
' 4. 쿠팡 판매자 빠른 입력 - 공통 처리 함수
'    Ctrl+Shift+1 → 오아시스 펌프
'    Ctrl+Shift+2 → 서울종합펌프
'    Ctrl+Shift+3 → 펌프샵
' =============================================
Sub ApplySeller(sellerName As String)
    Dim wsResult  As Worksheet
    Dim wsMk      As Worksheet
    Dim curRow    As Long
    Dim modelName As String
    Dim lprice    As Variant
    Dim msrp      As Variant
    Dim dc        As Variant
    Dim link      As String

    If ActiveSheet.Name <> "오픈마켓확인" Then
        MsgBox "오픈마켓확인 시트에서 실행해주세요.", vbExclamation
        Exit Sub
    End If

    Set wsMk = ActiveSheet
    curRow   = ActiveCell.Row
    If curRow <= 1 Then Exit Sub

    modelName = Trim(wsMk.Cells(curRow, 2).Value)
    lprice    = wsMk.Cells(curRow, 4).Value
    msrp      = wsMk.Cells(curRow, 5).Value
    dc        = wsMk.Cells(curRow, 6).Value
    link      = wsMk.Cells(curRow, 8).Value

    If modelName = "" Then Exit Sub

    ' 오픈마켓확인 G열에 판매자 입력 + 연초록 배경
    wsMk.Cells(curRow, 7).Value = sellerName
    wsMk.Cells(curRow, 7).Interior.Color = RGB(204, 255, 204)

    ' 통합결과 시트에서 모델행 + 판매처열 찾기
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("통합결과")
    On Error GoTo 0
    If wsResult Is Nothing Then Exit Sub

    Dim modelRow  As Long
    Dim sellerCol As Long
    Dim sr As Long, sc As Long

    modelRow  = 0
    sellerCol = 0

    For sr = 3 To wsResult.Cells(wsResult.Rows.Count, 1).End(xlUp).Row
        If Trim(wsResult.Cells(sr, 1).Value) = modelName Then
            modelRow = sr
            Exit For
        End If
    Next sr

    For sc = 3 To wsResult.Cells(1, wsResult.Columns.Count).End(xlToLeft).Column
        If Trim(wsResult.Cells(1, sc).Value) = sellerName Then
            sellerCol = sc
            Exit For
        End If
    Next sc

    If modelRow = 0 Or sellerCol = 0 Then
        ' 통합결과에 없어도 오픈마켓확인엔 이미 입력됨 - 다음 행으로 이동
        GoTo MoveNext
    End If

    ' 기존 값보다 낮을 때만 덮어쓰기
    Dim existVal As Variant
    existVal = wsResult.Cells(modelRow, sellerCol).Value
    If IsNumeric(existVal) And existVal > 0 Then
        If IsNumeric(lprice) And lprice >= existVal Then GoTo MoveNext
    End If

    ' 최저가 + 하이퍼링크
    With wsResult.Cells(modelRow, sellerCol)
        .Value = lprice
        .NumberFormat = "#,##0"
        .HorizontalAlignment = xlRight
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        If link <> "" Then
            wsResult.Hyperlinks.Add Anchor:=wsResult.Cells(modelRow, sellerCol), _
                Address:=link, TextToDisplay:=Format(lprice, "#,##0")
            .Font.Color = RGB(5, 99, 193)
            .Font.Underline = xlUnderlineStyleSingle
        End If
    End With

    ' 인터넷 권장가
    With wsResult.Cells(modelRow + 1, sellerCol)
        .Value = msrp
        .NumberFormat = "#,##0"
        .HorizontalAlignment = xlRight
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' DC율 + 색상
    With wsResult.Cells(modelRow + 2, sellerCol)
        If IsNumeric(dc) Then
            Dim dcPct As Double
            dcPct = IIf(dc > 1, dc / 100, dc)
            .Value = dcPct
            .NumberFormat = "0.0%"
            Select Case dcPct * 100
                Case Is >= 25: .Interior.Color = RGB(255, 1, 1)
                Case Is >= 22: .Interior.Color = RGB(255, 150, 150)
                Case Is >= 20: .Interior.Color = RGB(255, 150, 1)
                Case Is >= 17: .Interior.Color = RGB(255, 255, 1)
            End Select
        End If
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

MoveNext:
    ' 다음 행으로 자동 이동
    Dim lastRow As Long
    lastRow = wsMk.Cells(wsMk.Rows.Count, 1).End(xlUp).Row
    If curRow + 1 <= lastRow Then
        wsMk.Cells(curRow + 1, 1).Select
    End If
End Sub

' Ctrl+Shift+1 → 오아시스 펌프
Sub Seller_1()
    Call ApplySeller("오아시스 펌프")
End Sub

' Ctrl+Shift+2 → 서울종합펌프
Sub Seller_2()
    Call ApplySeller("서울종합펌프")
End Sub

' Ctrl+Shift+3 → 펌프샵
Sub Seller_3()
    Call ApplySeller("펌프샵")
End Sub

' Ctrl+Shift+4 → 윈디샵
Sub Seller_4()
    Call ApplySeller("윈디샵")
End Sub

' Ctrl+Shift+5 → 직접 입력
Sub Seller_Custom()
    Dim name As String
    name = InputBox("판매자명 입력:", "직접 입력")
    If name <> "" Then Call ApplySeller(name)
End Sub

' ── 단축키 등록 (파일 열릴 때 자동 실행) ──────
Sub Auto_Open()
    ' 오픈마켓확인 시트 단축키 (Ctrl+Shift+1~5)
    Application.OnKey "^+1", "Seller_1"
    Application.OnKey "^+2", "Seller_2"
    Application.OnKey "^+3", "Seller_3"
    Application.OnKey "^+4", "Seller_4"
    Application.OnKey "^+5", "Seller_Custom"
    ' 통합결과 시트 단축키 (Ctrl+Shift+Q~T)
    Application.OnKey "^+q", "Result_1"
    Application.OnKey "^+w", "Result_2"
    Application.OnKey "^+e", "Result_3"
    Application.OnKey "^+r", "Result_4"
    Application.OnKey "^+t", "Result_Custom"
End Sub

' =============================================
' 5. 통합결과 시트에서 선택 범위의 판매자 열 변경
'    선택한 셀들의 판매처(1행)를 지정 판매자로 일괄 변경
'    Ctrl+Shift+Q → 오아시스 펌프
'    Ctrl+Shift+W → 서울종합펌프
'    Ctrl+Shift+E → 펌프샵
'    Ctrl+Shift+R → 윈디샵
'    Ctrl+Shift+T → 직접 입력
' =============================================
Sub ApplySellerToResult(sellerName As String)
    Dim wsResult As Worksheet
    Dim sel      As Range
    Dim cell     As Range
    Dim sellerCol As Long
    Dim sc       As Long
    Dim lastCol  As Long
    Dim newCol   As Long

    If ActiveSheet.Name <> "통합결과" Then
        MsgBox "통합결과 시트에서 실행해주세요.", vbExclamation
        Exit Sub
    End If

    Set wsResult = ActiveSheet
    Set sel = Selection

    If sel Is Nothing Then Exit Sub

    ' 변경할 판매자 열 찾기 (1행에서 sellerName 검색)
    lastCol  = wsResult.Cells(1, wsResult.Columns.Count).End(xlToLeft).Column
    sellerCol = 0
    For sc = 3 To lastCol
        If Trim(wsResult.Cells(1, sc).Value) = sellerName Then
            sellerCol = sc
            Exit For
        End If
    Next sc

    If sellerCol = 0 Then
        MsgBox "'" & sellerName & "' 판매처를 통합결과 1행에서 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 선택 범위의 각 셀을 처리
    Dim movedCount As Long
    movedCount = 0

    For Each cell In sel
        Dim r As Long
        r = cell.Row

        ' 유효한 데이터 행인지 확인 (값 있고, 3행 이상)
        If r >= 3 And IsNumeric(cell.Value) And cell.Value > 0 Then
            Dim srcCol As Long
            srcCol = cell.Column

            ' 같은 열이면 건너뜀
            If srcCol = sellerCol Then GoTo NextCell

            ' 기존 값 읽기
            Dim lprice As Variant
            Dim msrp   As Variant
            Dim dc     As Variant
            Dim link   As String

            ' 최저가 행인지 판단 (3, 6, 9... → (r-3) mod 3 = 0)
            Dim baseRow As Long
            baseRow = r - ((r - 3) Mod 3)  ' 해당 모델의 최저가 행

            lprice = wsResult.Cells(baseRow,     srcCol).Value
            msrp   = wsResult.Cells(baseRow + 1, srcCol).Value
            dc     = wsResult.Cells(baseRow + 2, srcCol).Value

            ' 하이퍼링크 추출
            link = ""
            If wsResult.Cells(baseRow, srcCol).Hyperlinks.Count > 0 Then
                link = wsResult.Cells(baseRow, srcCol).Hyperlinks(1).Address
            End If

            If Not IsNumeric(lprice) Or lprice <= 0 Then GoTo NextCell

            ' 대상 열 기존값 확인 (낮을 때만 덮어쓰기)
            Dim existVal As Variant
            existVal = wsResult.Cells(baseRow, sellerCol).Value
            If IsNumeric(existVal) And existVal > 0 Then
                If lprice >= existVal Then GoTo NextCell
            End If

            ' 원본 셀 지우기
            wsResult.Cells(baseRow,     srcCol).ClearContents
            wsResult.Cells(baseRow + 1, srcCol).ClearContents
            wsResult.Cells(baseRow + 2, srcCol).ClearContents
            wsResult.Cells(baseRow + 2, srcCol).Interior.ColorIndex = xlNone

            ' 대상 열에 입력
            With wsResult.Cells(baseRow, sellerCol)
                .Value = lprice
                .NumberFormat = "#,##0"
                .HorizontalAlignment = xlRight
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                If link <> "" Then
                    wsResult.Hyperlinks.Add Anchor:=wsResult.Cells(baseRow, sellerCol), _
                        Address:=link, TextToDisplay:=Format(lprice, "#,##0")
                    .Font.Color = RGB(5, 99, 193)
                    .Font.Underline = xlUnderlineStyleSingle
                End If
            End With

            With wsResult.Cells(baseRow + 1, sellerCol)
                .Value = msrp
                .NumberFormat = "#,##0"
                .HorizontalAlignment = xlRight
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
            End With

            With wsResult.Cells(baseRow + 2, sellerCol)
                If IsNumeric(dc) Then
                    Dim dcPct As Double
                    dcPct = IIf(dc > 1, dc / 100, dc)
                    .Value = dcPct
                    .NumberFormat = "0.0%"
                    Select Case dcPct * 100
                        Case Is >= 25: .Interior.Color = RGB(255, 1, 1)
                        Case Is >= 22: .Interior.Color = RGB(255, 150, 150)
                        Case Is >= 20: .Interior.Color = RGB(255, 150, 1)
                        Case Is >= 17: .Interior.Color = RGB(255, 255, 1)
                    End Select
                End If
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
            End With

            movedCount = movedCount + 1
        End If
NextCell:
    Next cell

    If movedCount > 0 Then
        MsgBox movedCount & "개 항목을 '" & sellerName & "' 열로 이동했습니다.", vbInformation
    Else
        MsgBox "이동할 항목이 없습니다." & vbCrLf & "(이미 같은 열이거나 기존값이 더 낮음)", vbInformation
    End If
End Sub

Sub Result_1(): Call ApplySellerToResult("오아시스 펌프"): End Sub
Sub Result_2(): Call ApplySellerToResult("서울종합펌프"): End Sub
Sub Result_3(): Call ApplySellerToResult("펌프샵"): End Sub
Sub Result_4(): Call ApplySellerToResult("윈디샵"): End Sub
Sub Result_Custom()
    Dim name As String
    name = InputBox("이동할 판매처명 입력:", "직접 입력")
    If name <> "" Then Call ApplySellerToResult(name)
End Sub
